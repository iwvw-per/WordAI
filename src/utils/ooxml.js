import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";
const REF_BOOKMARK_PREFIX = "wordai_ref_";

/**
 * 标记选区并建立参考文献索引
 * 深度修复：隔离 results 数组，避免逻辑干扰
 */
export async function markSelection() {
  let finalItems = [];
  await Word.run(async (context) => {
    try {
      const doc = context.document;
      const selection = doc.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 1. 系统级清理
      const allCCs = doc.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          const tag = cc.tag || "";
          if (tag.includes(BOU_START) || tag.includes(BOU_END) || tag.includes(SHIELD_PREFIX)) {
              cc.delete(true);
          }
      }
      await context.sync();

      // 2. 书签化参考文献 (自愈基础)
      // 仅在首次分析时进行全量扫描 (基于性能考虑，我们可以只搜本段或常见模式)
      const refParas = doc.body.paragraphs.search("^\\[[0-9]+\\]", { matchWildcards: true });
      refParas.load("items");
      await context.sync();
      
      for (const p of refParas.items) {
          p.load("text");
      }
      await context.sync();

      for (const p of refParas.items) {
          const match = p.text.match(/^\[(\d+)\]/);
          if (match) {
              doc.bookmarks.add(`${REF_BOOKMARK_PREFIX}${match[1]}`, p);
          }
      }
      await context.sync();

      // 3. 确定目标处理范围
      let targetRanges = [];
      if (selection.isEmpty) {
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");
        await context.sync();
        if (paragraphs.items.length > 0) targetRanges = [paragraphs.items[0]];
      } else {
        targetRanges = [selection];
      }

      let globalIdx = 0;
      for (const range of targetRanges) {
        range.load(["text", "font"]);
        await context.sync();
        const text = range.text || "";
        if (!text.trim()) continue;

        const sessionId = Date.now() + "_" + Math.floor(Math.random() * 100);
        
        // 边界锚点
        const sM = range.getRange("Start").insertContentControl();
        sM.tag = `${BOU_START}_${sessionId}`;
        sM.appearance = Word.ContentControlAppearance.hidden;

        const eM = range.getRange("End").insertContentControl();
        eM.tag = `${BOU_END}_${sessionId}`;
        eM.appearance = Word.ContentControlAppearance.hidden;

        // 识别正文引用
        const matches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        matches.load("items");
        await context.sync();

        const refMetadata = [];
        let aiInput = text;

        for (let i = 0; i < matches.items.length; i++) {
           const m = matches.items[i];
           m.load("text");
           m.track();
           await context.sync();
           
           const id = globalIdx++;
           const placeholder = `{{REF_${id}}}`;
           
           const shield = m.getRange("Start").insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${id}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           const pos = aiInput.indexOf(m.text);
           if (pos >= 0) {
              aiInput = aiInput.substring(0, pos) + placeholder + aiInput.substring(pos + m.text.length);
           }
           refMetadata.push({ placeholder, id });
        }

        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        finalItems.push({ 
          text: aiInput, 
          refMap: refMetadata, 
          boundaryTags: { start: sM.tag, end: eM.tag },
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
  return finalItems;
}

/**
 * 后置自愈重链器 (VBA 核心移植)
 */
async function autoRelinkRange(range) {
    try {
        const matches = range.search("\\[[0-9]+\\]", { matchWildcards: true });
        matches.load("items");
        await range.context.sync();
        
        for (const m of matches.items) {
            m.load(["text", "hyperlink"]);
            await range.context.sync();
            if (m.hyperlink) continue;
            
            const num = m.text.match(/\[(\d+)\]/)?.[1];
            if (num) {
                const bmName = `${REF_BOOKMARK_PREFIX}${num}`;
                const bm = range.context.document.bookmarks.getAtOrNullObject(bmName);
                await range.context.sync();
                if (!bm.isNullObject) {
                    m.insertHyperlink(`#${bmName}`, m.text, Word.InsertLocation.replace);
                    m.font.underline = Word.UnderlineStyle.none;
                    m.font.color = "black";
                }
            }
        }
    } catch (e) {}
}

/**
 * 复合回写系统 6.1
 */
async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      const start = ccs.items.find(c => c.tag === boundaryTags.start);
      const end = ccs.items.find(c => c.tag === boundaryTags.end);
      if (!start || !end) return;

      const segments = aiResult.trim().split(/({{REF_\d+}})/g);
      let insertionCursor = start.getRange("After");

      // 分步同步填缝 (保活)
      for (const seg of segments) {
          if (!seg) continue;
          if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
              const id = seg.match(/\d+/)[0];
              const shield = ccs.items.find(c => c.tag === `${SHIELD_PREFIX}${id}`);
              if (shield) insertionCursor = shield.getRange("After");
          } else {
              let nextAnchor = null;
              const nextIdx = segments.indexOf(seg) + 1;
              const nextSeg = segments[nextIdx];
              if (nextSeg && nextSeg.startsWith("{{REF_")) {
                  const nId = nextSeg.match(/\d+/)[0];
                  nextAnchor = ccs.items.find(c => c.tag === `${SHIELD_PREFIX}${nId}`);
              } else {
                  nextAnchor = end;
              }

              if (nextAnchor) {
                  const gap = insertionCursor.expandTo(nextAnchor.getRange("Before"));
                  gap.insertText(seg, Word.InsertLocation.replace);
                  await context.sync();
              }
          }
      }

      const r = start.getRange("After").expandTo(end.getRange("Before"));
      await autoRelinkRange(r);

      if (baseFont) {
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }

      for (const c of ccs.items) {
          if (c.tag && (c.tag.includes(SHIELD_PREFIX) || c.tag.includes(BOU_START) || c.tag.includes(BOU_END))) {
              c.delete(true);
          }
      }
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement Error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在扫描参考文献...", true);
  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
    const item = results[i];
    if (onStatus) onStatus("processing", `AI 润色中 (${i + 1}/${results.length})...`, true);
    
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `同步内容并执行自愈功能 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.boundaryTags, item.baseFont);
      }
    } catch (err) {
      await clearMarks();
      throw err;
    }
  }
  return { result: "全部完成" };
}

export async function clearMarks() {
  try {
    await Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag && (cc.tag.includes(SHIELD_PREFIX) || cc.tag.includes(BOU_START) || cc.tag.includes(BOU_END))) {
              cc.delete(true);
          }
      }
      await context.sync();
    });
  } catch {}
}
