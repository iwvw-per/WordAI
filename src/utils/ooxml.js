import * as storage from "./storage.js";
import * as format from "./format.js";

const CC_TAG = "wordai_target";

// ==================== 标记选区（选区感知版） ====================

export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 清理旧标记
      const existing = context.document.contentControls.getByTag(CC_TAG);
      existing.load("items");
      await context.sync();
      for (const cc of existing.items) cc.delete(true);
      await context.sync();

      let targetRanges = [];
      const text = selection.text.trim();

      if (!text) {
        // 选区为空，处理当前光标所在段落
        const paragraphs = selection.paragraphs;
        paragraphs.load(["items", "text", "style"]);
        await context.sync();
        if (paragraphs.items.length > 0) {
          targetRanges = [paragraphs.items[0]];
        }
      } else {
        // 选区不为空，直接处理选区
        targetRanges = [selection];
      }

      for (const range of targetRanges) {
        range.load(["text", "style"]);
        await context.sync();
        const rangeText = range.text.trim();
        if (!rangeText) continue;

        // 过滤标题（仅在处理全段落时且不是唯一段落时应用）
        const styleName = (range.style || "").toString().toLowerCase();
        const isHeaderPattern = /^(\s*\d+(\.\d+)*[\.\s\t])/.test(rangeText) && rangeText.length < 120;
        const isHeadingStyle = styleName.includes("heading") || styleName.includes("标题");
        if (skipRules.headings && !text && (isHeadingStyle || isHeaderPattern)) {
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
        
        // 捕获原始字体信息以备回写恢复
        const font = range.font;
        font.load(["name", "size", "color"]);
        await context.sync();

        results.push({ 
          text: tokenizedText, 
          refMap: finalRefMap, 
          baseFont: { name: font.name, size: font.size, color: font.color }
        });
      }
      await context.sync();
    } catch (err) {
      console.error("markSelection error:", err);
    }
  });
  return results;
}

// ==================== 逐段回写（支持格式恢复） ====================

async function replaceSingleMarkedContent(newText, refMap, baseFont) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(CC_TAG);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) return;
    const cc = ccs.items[0];
    
    // 强制先清空内容
    cc.insertText("", Word.InsertLocation.replace);
    await context.sync();

    const lines = newText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line && i !== lines.length - 1) {
        cc.insertParagraph("", Word.InsertLocation.end);
        continue;
      }
      if (line) {
        // processMarkdownLine 内部会处理加粗和引用
        await format.processMarkdownLine(cc, line, refMap);
        
        // 恢复基本字体属性（防止因清空导致的手动格式丢失）
        if (baseFont) {
          const range = cc.getRange();
          range.font.name = baseFont.name;
          if (baseFont.size) range.font.size = baseFont.size;
          if (baseFont.color) range.font.color = baseFont.color;
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
