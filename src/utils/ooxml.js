import * as storage from "./storage.js";
import * as format from "./format.js";

const CC_TAG = "wordai_target";

// ==================== 标记选区（高性能版） ====================

export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load(["items", "text", "style"]);
      await context.sync();

      let workParagraphs = [...paragraphs.items];
      if (workParagraphs.length === 0) {
        const startPars = selection.getRange("Start").paragraphs;
        startPars.load(["items", "text", "style"]);
        await context.sync();
        workParagraphs = [...startPars.items];
      }

      if (workParagraphs.length === 0) return;

      // 清理旧标记
      const existing = context.document.contentControls.getByTag(CC_TAG);
      existing.load("items");
      await context.sync();
      for (const cc of existing.items) cc.delete(true);
      await context.sync();

      const isSingle = workParagraphs.length === 1;
      const refSearches = [
        "\\[[0-9.,\\- ]@\\]",
        "图 [0-9]@",
        "表 [0-9]@",
        "Fig. [0-9]@",
        "Figure [0-9]@",
        "Table [0-9]@"
      ];

      for (const p of workParagraphs) {
        const text = p.text.trim();
        if (!text) continue;

        // 过滤逻辑
        const styleName = (p.style || "").toString().toLowerCase();
        const isHeaderPattern = /^(\s*\d+(\.\d+)*[\.\s\t])/.test(text) && text.length < 120;
        const isHeadingStyle = styleName.includes("heading") || styleName.includes("标题");

        if (skipRules.headings && !isSingle && (isHeadingStyle || isHeaderPattern)) {
          continue;
        }

        // 搜索引用：并发收集 search 对象
        const refMap = [];
        const searchPromises = refSearches.map(pattern => {
          const matches = p.search(pattern, { matchWildcards: true });
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

        let tokenizedText = text;
        const finalRefMap = [];
        
        // 按长度倒序，防止子串误替换
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

        const cc = p.insertContentControl();
        cc.tag = CC_TAG;
        cc.appearance = Word.ContentControlAppearance.hidden;
        results.push({ text: tokenizedText, refMap: finalRefMap });
      }
      await context.sync();
    } catch (err) {
      console.error("markSelection error:", err);
    }
  });
  return results;
}

// ==================== 逐段回写（调用统一格式化逻辑） ====================

async function replaceSingleMarkedContent(newText, refMap) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(CC_TAG);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) return;
    const cc = ccs.items[0];
    
    // 清空并按行注入
    cc.insertText("", Word.InsertLocation.replace);
    const lines = newText.split(/\r?\n/);
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line && i !== lines.length - 1) {
        cc.insertParagraph("", Word.InsertLocation.end);
        continue;
      }
      if (line) {
        await format.processMarkdownLine(cc, line, refMap);
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
  if (!results || results.length === 0) throw new Error("无有效待处理段落");

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }

    if (onStatus) onStatus("processing", `AI 正在处理 (${i + 1}/${results.length})...`, true);

    const item = results[i];
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }

      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在应用修改 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap);
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
