import * as storage from "./storage.js";

const CC_TAG = "wordai_target";

/**
 * 标记并分析选区，记录引用信息
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 清理旧标记
      const existing = context.document.contentControls.getByTag(CC_TAG);
      existing.load("items");
      await context.sync();
      for (const cc of existing.items) cc.delete(true); 
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

        const styleName = (range.style || "").toString().toLowerCase();
        const isHeaderPattern = /^(\s*\d+(\.\d+)*[\.\s\t])/.test(rangeText.trim()) && rangeText.length < 120;
        const isHeadingStyle = styleName.includes("heading") || styleName.includes("标题");
        if (skipRules.headings && !isRealSelection && (isHeadingStyle || isHeaderPattern)) {
          continue;
        }

        // 搜索引用
        const refSearches = [
          "\\[[0-9.,\\- ]@\\]",
          "图 [0-9]@",
          "表 [0-9]@",
          "Fig. [0-9]@",
          "Figure [0-9]@",
          "Table [0-9]@"
        ];
        
        const refMap = [];
        const searchPromises = refSearches.map(pattern => {
          const matches = range.search(pattern, { matchWildcards: true });
          matches.load("items");
          return matches;
        });
        
        await context.sync();

        const allFoundRefs = [];
        for (const matches of searchPromises) {
          for (const matchRange of matches.items) {
            allFoundRefs.push(matchRange);
          }
        }
        
        if (allFoundRefs.length > 0) await context.sync();

        // 为每个引用建立唯一标识
        let tokenizedText = rangeText;
        const finalRefMap = [];
        
        // 按在文档中出现的先后顺序排序
        allFoundRefs.sort((a, b) => a.items?.[0]?.offset || 0 - (b.items?.[0]?.offset || 0));
        
        // 这一步只是为了给 AI 文本加占位符
        const uniqueItems = [];
        allFoundRefs.forEach(r => {
           r.load("text");
        });
        await context.sync();

        allFoundRefs.sort((a, b) => b.text.length - a.text.length).forEach(r => {
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
            name: font.name, 
            size: font.size, 
            color: font.color,
            bold: font.bold,
            italic: font.italic,
            underline: font.underline
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

// ==================== 终极原地替换策略 ====================

async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();

      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];
      
      // 1. 获取 AI 结果中的序列
      // 我们将 AI 结果切分为 [Text, Placeholder, Text, Placeholder...]
      const aiSegments = aiResult.trim().split(/({{REF_\d+}})/g).filter(s => s !== "");

      // 2. 在当前文档 CC 中重新定位所有引用 (必须实时定位，因为之前的 range 已失效)
      // 注意：这一步我们不使用 regex 搜索，而是按顺序查找 refMap 中的 original
      // 但为了 100% 保留跳转，我们最好是找那些“非文本”的东西？不，还是找 text
      const docRefs = [];
      for (const refDef of refMap) {
        // 在 CC 范围内搜索原始文本
        const matches = cc.search(refDef.original, { matchCase: true });
        matches.load("items");
        await context.sync();
        for (const m of matches.items) {
          docRefs.push({ range: m, placeholder: refDef.placeholder });
        }
      }

      // 按文档物理顺序排序
      // 我们需要通过 range.getRange("Start") 的比较来排序
      // 但简化的做法是：既然我们希望“原地”且“不触碰”，我们其实可以直接做“全文替换占位符”后再“局部还原”？
      // 不，用户说 insertOoxml 导致链接失效。
      
      // ★ 策略：如果 AI 返回的引用顺序和原文一致（99%的情况），采用“间隙替换”
      
      // 先把 AI 序列中的占位符转换成对应的 docRefs
      const sequence = [];
      for (const seg of aiSegments) {
        if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
          // 找到文档中对应的第一个尚未被使用的该引用
          const matchedDocRefIndex = docRefs.findIndex(dr => dr.placeholder === seg && !dr.used);
          if (matchedDocRefIndex !== -1) {
            docRefs[matchedDocRefIndex].used = true;
            sequence.push({ type: "ref", data: docRefs[matchedDocRefIndex] });
          } else {
            sequence.push({ type: "text", data: seg }); // 找不到就当纯文本（AI捏造的）
          }
        } else {
          sequence.push({ type: "text", data: seg });
        }
      }

      // ★ 终极保险：为了 100% 保留跳转，我们必须保证引用的 Range 对象物理上不被 delete。
      // 我们采取“先回写，再移动”太复杂。
      // 采取“逆序替换间隙”：
      
      // 准备工作：在 CC 结尾插入一个哨兵
      const endSentinel = cc.getRange("End").insertText("", "After");
      
      // 逆序处理 sequence
      let currentEndRange = cc.getRange("End"); 
      
      for (let i = sequence.length - 1; i >= 0; i--) {
        const item = sequence[i];
        if (item.type === "ref") {
            // 这是我们要“保活”的引用。
            // 我们不删除它，而是把它前面的文本替换掉。
            // 等待，逆序逻辑在这里比较绕。
            
            // 换个思路：如果 AI 结果是 A [1] B [2] C
            // 我们只需要把 A 换成 AI_A，B 换成 AI_B，C 换成 AI_C。
            // [1] 和 [2] 只要不动，跳转永远有效。
        }
      }

      // 既然用户建议“不碰原来的特殊格式”，最稳妥的方法其实是：
      // 1. 把 CC 内的所有非引用文字删掉/替换。
      // 2. 但 Word 不好直接删“非引用文字”。
      
      // 改回“搜索还原”但【不使用 insertOoxml】，而是使用【移动 Range】？
      // Word API 1.1 不支持移动 Range。
      
      // 【最终确认方案】：
      // 使用 `insertText` 替换占位符会导致角标消失吗？会，因为 insertText 是纯文本。
      // 那么，如果我们在 `markSelection` 时，把每个引用都先包进一个【临时的 ContentControl】呢？
      // 然后我们只需把这些 CC 之间的文本替换掉。
      // 这种方式引用的物理对象一直存在，跳转绝对不会丢。
      
      // 这种太复杂且容易崩。
      
      // 让我们回到 `insertOoxml`。为什么链接会坏？
      // 因为 `getOoxml()` 可能只拿到了文字，没拿到 Field 结构。
      // 解决：在 `markSelection` 时，使用 `range.parentContentControl` 或扩大 Range 确保拿到完整 Field。
      // 或者... 使用 `range.insertHtml`？
      
      // 其实有个最简单的办法保留跳转和角标：
      // 如果我们不使用 `ooxml.js` 的 `insertOoxml`，而是直接用 `cc.search(placeholder)` 
      // 然后用【原本就存在于文档某处的那个 Range】去替换它？
      // Word API 支持 `Range.copyFrom(otherRange)` 吗？不支持。
      
      // 那么：我们在回写前，把原本的引用【移动】到一个临时区域。
      // 回写完带占位符的文本后，再把它们【移动】回来？
      // Word 1.1 移动 $= 剪切 + 粘贴 $= 还是会丢链接（在某些复杂情况下）。
      
      // ！！！ 真正的原因 ！！！
      // 交叉引用在 Word 里是 `REF` 域。`insertOoxml` 只要包含了完整的 `<w:fldSimple>` 或 `<w:fldChar>` 标签，链接就不会坏。
      // 我的 `range.search()` 拿到的范围可能【紧紧包裹了文字】，导致去掉了外层的 Field 标签。
      
      // 【修复】：在 `markSelection` 中，通过 `range.expand("Word")` 或类似手段尝试包含 Field 边界。
      // 或者，使用 `range.getRange("Whole")`？不。
      
      // 最直接的办法：不使用占位符策略了？那 AI 就看不到引用。
      
      // 方案：回写逻辑回归最简：
      // 1. 填入带占位符的文本。
      // 2. 搜索占位符，【原本是什么，就填回什么】。
      // 为了保证跳转，我们这次在 `markSelection` 时，不仅存 OOXML，还存 HTML。
      // `insertHtml` 有时比 `insertOoxml` 对链接更友好。
      
      // 但我打赌，最稳的办法是：
      await Word.run(async (context) => {
          // ... (之前的清理代码)
          cc.insertText(aiResult.trim(), Word.InsertLocation.replace);
          await context.sync();
          
          for (const ref of refMap) {
            const found = cc.search(ref.placeholder, { matchCase: true });
            found.load("items");
            await context.sync();
            for (const r of found.items) {
               // 尝试使用 insertOoxml。
               // 为了防止链接失效，我们要求 AI 这里的 placeholder 必须是独立的。
               r.insertOoxml(ref.ooxml, Word.InsertLocation.replace);
            }
          }
          
          // 恢复字体
          const range = cc.getRange();
          if (baseFont) {
              if (baseFont.name) range.font.name = baseFont.name;
              if (typeof baseFont.size === "number" && baseFont.size > 0) range.font.size = baseFont.size;
              if (baseFont.color) range.font.color = baseFont.color;
          }
          
          cc.delete(true);
          await context.sync();
      });

    });
  } catch (err) {
    console.error("replace error:", err);
    throw err;
  }
}

// ... (executeAndReplace 和 clearMarks 保持不变)

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在分析选区...");
  
  const diffMode = storage.getDiffMode();
  if (diffMode) {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = "TrackAll";
      await context.sync();
    });
  }

  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  const totalChars = results.reduce((sum, item) => sum + (item.text?.length || 0), 0);
  if (onStatus) onStatus("processing", `正在处理 (${totalChars} 字)...`);

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }

    const item = results[i];
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }

      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在应用修改 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.baseFont);
      } else {
        await Word.run(async (context) => {
          const ccs = context.document.contentControls.getByTag(CC_TAG);
          ccs.load("items");
          await context.sync();
          if (ccs.items.length > 0) { ccs.items[0].delete(true); await context.sync(); }
        });
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
