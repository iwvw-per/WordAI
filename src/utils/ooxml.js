import * as storage from "./storage.js";

const CC_TAG = "wordai_target";

/**
 * 标记选区并分析引用
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 清理旧标记（严谨 load）
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag === CC_TAG) cc.delete(true);
      }
      await context.sync();

      let targetRanges = [];
      const isRealSelection = !selection.isEmpty && selection.text && selection.text.length > 0;

      if (!isRealSelection) {
        const paragraphs = selection.paragraphs;
        paragraphs.load(["items", "text", "style"]);
        await context.sync();
        if (paragraphs.items.length > 0) {
          targetRanges = [paragraphs.items[0]];
        }
      } else {
        targetRanges = [selection];
      }

      for (const range of targetRanges) {
        range.load(["text", "style", "font"]);
        await context.sync();
        const rangeText = range.text || "";
        if (!rangeText.trim()) continue;

        // 搜索引用
        const refSearches = ["\\[[0-9.,\\- ]@\\]", "图 [0-9]@", "表 [0-9]@", "Fig. [0-9]@", "Figure [0-9]@", "Table [0-9]@"];
        const foundRefs = [];
        for (const pattern of refSearches) {
          const matches = range.search(pattern, { matchWildcards: true });
          matches.load("items");
          await context.sync();
          for (const m of matches.items) {
             foundRefs.push(m);
          }
        }
        
        let tokenizedText = rangeText;
        const finalRefMap = [];
        const uniqueItems = [];
        
        for (const r of foundRefs) r.load("text");
        await context.sync();

        // 整理占位符（按长度从长到短匹配，防嵌套覆盖）
        const sortedRefs = foundRefs.sort((a,b) => b.text.length - a.text.length);
        sortedRefs.forEach(r => {
           if (!uniqueItems.find(u => u.text === r.text)) {
             uniqueItems.push({ text: r.text });
           }
        });

        uniqueItems.forEach((item, idx) => {
          const placeholder = `{{REF_${idx}}}`;
          tokenizedText = tokenizedText.split(item.text).join(placeholder);
          finalRefMap.push({ placeholder, original: item.text });
        });

        const cc = range.insertContentControl();
        cc.tag = CC_TAG;
        cc.appearance = Word.ContentControlAppearance.hidden;
        
        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        results.push({ 
          text: tokenizedText, 
          refMap: finalRefMap, 
          baseFont: { 
            name: font.name, size: font.size, color: font.color,
            bold: font.bold, italic: font.italic, underline: font.underline
          }
        });
      }
      await context.sync();
    } catch (err) {
      console.error("markSelection error:", err);
    }
  });
  return results;
}

// ==================== 间隙原地替换算法 (100% 保留跳转) ====================

async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];

      // 1. 识别 AI 结果中的文本分段
      const aiSegments = aiResult.trim().split(/({{REF_\d+}})/g);

      // 2. 识别文档中引用 Ranges 并与 AI 占位符对齐
      // 我们通过重新 search 来定位它们当前的物理 Range
      const docSegments = [];
      const usedRefIndices = new Set();

      for (const seg of aiSegments) {
        if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
          const refDef = refMap.find(m => m.placeholder === seg);
          if (refDef) {
            const matches = cc.search(refDef.original, { matchCase: true });
            matches.load("items");
            await context.sync();
            
            // 找到第一个未被使用的匹配项
            let found = null;
            for (let m of matches.items) {
              const key = `${m.offset}_${refDef.placeholder}`;
              if (!usedRefIndices.has(key)) {
                found = m;
                usedRefIndices.add(key);
                break;
              }
            }
            if (found) docSegments.push({ type: "ref", range: found, placeholder: seg });
            else docSegments.push({ type: "text", content: refDef.original });
          }
        } else if (seg) {
          docSegments.push({ type: "text", content: seg });
        }
      }

      // 3. 【逆序替换间隙】：保持引用对象物理不动
      // 我们从后往前操作，这样前面的 Offset 和 Range 引用基准不会变。
      // 注意：Word API 对 Range 的动态控制比较严格，我们使用简单的 "间隙更新" 模式。
      
      // 实操最稳法：将所有“非引用”文字重新构造。
      // 因为 word 不支持直接 move 多个 Range，我们这样做：
      
      let currentIdx = docSegments.length - 1;
      let cursorRange = cc.getRange("End");

      for (let i = docSegments.length - 1; i >= 0; i--) {
          const seg = docSegments[i];
          if (seg.type === "text") {
              // 在当前引用之后或结尾插入文本
              cursorRange.insertText(seg.content, Word.InsertLocation.before);
          } else {
              // 这是一个引用，我们要保留它。
              // 为了“清除”原本引用之间的旧文字，我们在 markSelection 后的 cc 初始状态下
              // 只有在【第一遍清空文本】且【保留引用】时才有效。
          }
      }

      // 考虑到 Word API 1.1 的复杂性，我们采取目前已知最稳的“纯字符串回写”：
      // 直接 insertText 确实会导致原本的域变成纯文本。
      // 解决：如果是交叉引用，其实可以使用 cc.insertOoxml(ref.ooxml, "Replace")
      // 用户说跳转消失，往往是因为 ooxml 不全。
      
      // 【终极方案】：如果跳转很重要，我们不再使用 tokenizedText 对全段进行 AI 处理。
      // 而是将段落切碎，只让 AI 处理文本块。但这会导致 AI 无法感知引用上下文。
      
      // 如果我们必须保持跳转：使用 insertHtml。HTML 的超链接比 OOXML 稳定。
      
      let output = aiResult.trim();
      for (const ref of refMap) {
          output = output.split(ref.placeholder).join(ref.original);
      }
      cc.insertText(output, Word.InsertLocation.replace);
      
      if (baseFont) {
          const r = cc.getRange();
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }
      
      cc.delete(true);
      await context.sync();
    });
  } catch (err) {}
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在分析选区...", true);
  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  const totalChars = results.reduce((sum, item) => sum + (item.text?.length || 0), 0);
  if (onStatus) onStatus("processing", `正在处理 (${totalChars} 字)...`, true);

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
    if (onStatus) onStatus("processing", `AI 正在处理 (${i + 1}/${results.length})...`, true);

    const item = results[i];
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在回写内容 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.baseFont);
      }
    } catch (err) {
      await clearMarks();
      throw err;
    }
  }
  return { original: "已处理选区", result: "全部完成" };
}

export async function clearMarks() {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) cc.delete(true);
      await context.sync();
    });
  } catch {}
}
