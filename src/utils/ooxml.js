import * as storage from "./storage.js";

const CC_TAG = "wordai_target";

// ==================== 标记选区 ====================

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

        // 过滤标题
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

        for (const matches of searchPromises) {
          for (const matchRange of matches.items) {
            matchRange.load("text");
            const ooxml = matchRange.getOoxml();
            refMap.push({ range: matchRange, ooxmlRequest: ooxml });
          }
        }
        
        if (refMap.length > 0) await context.sync();

        let tokenizedText = rangeText;
        const finalRefMap = [];
        const uniqueRefs = [];
        refMap.sort((a, b) => b.range.text.length - a.range.text.length).forEach(item => {
           if (!uniqueRefs.find(u => u.text === item.range.text)) {
             uniqueRefs.push({ text: item.range.text, ooxml: item.ooxmlRequest.value });
           }
        });

        uniqueRefs.forEach((item, idx) => {
          const placeholder = `{{REF_${idx}}}`;
          tokenizedText = tokenizedText.split(item.text).join(placeholder);
          finalRefMap.push({ placeholder, ooxml: item.ooxml, original: item.text });
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

// ==================== 回写：保留引用角标版 ====================

async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(CC_TAG);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) return;
    const cc = ccs.items[0];
    
    // 1. 清空内容并准备按需分段
    cc.insertText("", Word.InsertLocation.replace);
    await context.sync();

    // 2. 切分新文本行
    const lines = aiResult.split(/\r?\n/).filter(line => line.trim());
    
    for (let i = 0; i < lines.length; i++) {
        const lineText = lines[i].trim();
        if (!lineText) continue;

        // 如果有多行且不是第一行，创建一个新段落承载内容
        const insertionContainer = (i === 0) ? cc : cc.insertParagraph("", Word.InsertLocation.end);
        
        // 解析引用占位符并注入 (使用正则表达式切分避免丢失文本)
        const parts = lineText.split(/({{REF_\d+}})/g);
        for (const part of parts) {
            if (!part) continue;
            
            if (part.startsWith("{{REF_") && part.endsWith("}}")) {
                const ref = refMap.find(m => m.placeholder === part);
                if (ref && ref.ooxml) {
                    // ★ 核心：使用 insertOoxml 还原包含角标和链接的引用
                    insertionContainer.insertOoxml(ref.ooxml, Word.InsertLocation.end);
                } else {
                    insertionContainer.insertText(part, Word.InsertLocation.end);
                }
            } else {
                // 普通文本：应用基础字体属性
                const run = insertionContainer.insertText(part, Word.InsertLocation.end);
                if (baseFont) {
                    run.font.name = baseFont.name;
                    if (typeof baseFont.size === "number" && baseFont.size > 0) run.font.size = baseFont.size;
                    if (baseFont.color) run.font.color = baseFont.color;
                    run.font.bold = baseFont.bold;
                    run.font.italic = baseFont.italic;
                    run.font.underline = baseFont.underline;
                }
            }
        }
    }

    cc.delete(true); 
    await context.sync();
  });
}

// ==================== 批量串行执行 ====================

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
  if (!results || results.length === 0) throw new Error("请先选择要处理的文字内容");

  const totalChars = results.reduce((sum, item) => sum + (item.text?.length || 0), 0);
  if (onStatus) onStatus("processing", `正在处理 (${totalChars} 字)...`);

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }

    if (onStatus) onStatus("processing", `AI 正在处理 (${i + 1}/${results.length})...`, true);

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
