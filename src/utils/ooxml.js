import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";
const REF_BOOKMARK_PREFIX = "wordai_ref_";

/**
 * 标记选区：优先锁定选区，确保稳定性 6.2
 */
export async function markSelection() {
  let finalItems = [];
  try {
    await Word.run(async (context) => {
      const doc = context.document;
      
      // 1. 【优先级第1】：第一时间捕获选区状态
      const selection = doc.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 确定目标 Range
      let mainRanges = [];
      if (selection.isEmpty) {
        const ps = selection.paragraphs;
        ps.load("items");
        await context.sync();
        if (ps.items.length > 0) mainRanges = [ps.items[0]];
      } else {
        mainRanges = [selection];
      }

      if (mainRanges.length === 0) return;

      // 2. 【清理旧标记】
      const ccs = doc.contentControls;
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) {
          const t = cc.tag || "";
          if (t.includes(BOU_START) || t.includes(BOU_END) || t.includes(SHIELD_PREFIX)) {
              cc.delete(true);
          }
      }
      await context.sync();

      // 3. 【参考文献索引化】(异步防御性处理)
      try {
          // 只扫描可能包含 [1] 的行首段落
          const refMatches = doc.body.paragraphs.search("^\\[[0-9]+\\]", { matchWildcards: true });
          refMatches.load("items");
          await context.sync();
          for (let i = 0; i < refMatches.items.length; i++) {
              const p = refMatches.items[i];
              p.load("text");
              await context.sync();
              const m = p.text.match(/^\[(\d+)\]/);
              if (m) doc.bookmarks.add(`${REF_BOOKMARK_PREFIX}${m[1]}`, p);
          }
      } catch (e) {
          console.warn("Bookmark indexing skip or limited:", e);
      }

      // 4. 【分段处理选区内容】
      let globalCounter = 0;
      for (const range of mainRanges) {
        range.load(["text", "font"]);
        await context.sync();
        const originalText = range.text || "";
        if (!originalText.trim()) continue;

        const session = Date.now() + "_" + Math.floor(Math.random() * 100);
        
        // 分别标记 Start 和 End
        const startCC = range.getRange("Start").insertContentControl();
        startCC.tag = `${BOU_START}_${session}`;
        startCC.appearance = Word.ContentControlAppearance.hidden;

        const endCC = range.getRange("End").insertContentControl();
        endCC.tag = `${BOU_END}_${session}`;
        endCC.appearance = Word.ContentControlAppearance.hidden;

        // 识别正文引用
        const refMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        refMatches.load("items");
        await context.sync();

        const localMap = [];
        let aiInputText = originalText;

        for (let i = 0; i < refMatches.items.length; i++) {
           const m = refMatches.items[i];
           m.load("text");
           m.track();
           await context.sync();
           
           const uid = globalCounter++;
           const token = `{{REF_${uid}}}`;
           
           // 安装盾牌锚点 (零触碰)
           const shield = m.getRange("Start").insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${uid}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           const offset = aiInputText.indexOf(m.text);
           if (offset >= 0) {
              aiInputText = aiInputText.substring(0, offset) + token + aiInputText.substring(offset + m.text.length);
           }
           localMap.push({ placeholder: token, id: uid });
        }

        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        finalItems.push({ 
          text: aiInputText, 
          refMap: localMap, 
          boundaryTags: { start: startCC.tag, end: endCC.tag },
          baseFont: { 
            name: font.name, size: font.size, color: font.color,
            bold: font.bold, italic: font.italic, underline: font.underline
          }
        });
      }
      await context.sync();
    });
  } catch (err) {
    console.error("Critical markSelection error:", err);
    throw err; // 向上抛出以便 executeAndReplace 捕获
  }
  return finalItems;
}

/**
 * 后置自愈重链逻辑 (VBA 移植)
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
            
            const numMatch = m.text.match(/\[(\d+)\]/);
            if (numMatch) {
                const num = numMatch[1];
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
    } catch (e) {
        console.warn("Relink skip:", e);
    }
}

/**
 * 复合回写核心 6.2
 */
async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      const startAnchor = ccs.items.find(c => c.tag === boundaryTags.start);
      const endAnchor = ccs.items.find(c => c.tag === boundaryTags.end);
      if (!startAnchor || !endAnchor) return;

      const segments = aiResult.trim().split(/({{REF_\d+}})/g);
      let insertionPoint = startAnchor.getRange("After");

      // 分步同步填缝 (保活核心)
      for (const seg of segments) {
          if (!seg) continue;
          if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
              const id = seg.match(/\d+/)[0];
              const shield = ccs.items.find(c => c.tag === `${SHIELD_PREFIX}${id}`);
              if (shield) insertionPoint = shield.getRange("After");
          } else {
              let nextAnchor = null;
              const nextIdx = segments.indexOf(seg) + 1;
              if (nextIdx < segments.length) {
                  const ns = segments[nextIdx];
                  if (ns.startsWith("{{REF_")) {
                      const nid = ns.match(/\d+/)[0];
                      nextAnchor = ccs.items.find(c => c.tag === `${SHIELD_PREFIX}${nid}`);
                  }
              }
              if (!nextAnchor) nextAnchor = endAnchor;

              const gap = insertionPoint.expandTo(nextAnchor.getRange("Before"));
              gap.insertText(seg, Word.InsertLocation.replace);
              await context.sync();
          }
      }

      const finalR = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
      await autoRelinkRange(finalR);

      if (baseFont) {
          if (baseFont.name) finalR.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalR.font.size = baseFont.size;
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
  if (onStatus) onStatus("processing", "正在锁定选区与处理指纹...", true);
  try {
    const results = await markSelection();
    if (!results || results.length === 0) throw new Error("请先选择文字内容");

    for (let i = 0; i < results.length; i++) {
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      const item = results[i];
      if (onStatus) onStatus("processing", `AI 深度润色中 (${i + 1}/${results.length})...`, true);
      
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `分步同步与自愈重链 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.boundaryTags, item.baseFont);
      }
    }
    return { result: "全部完成" };
  } catch (err) {
    await clearMarks();
    throw err;
  }
}

export async function clearMarks() {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) {
          if (cc.tag && (cc.tag.includes(SHIELD_PREFIX) || cc.tag.includes(BOU_START) || cc.tag.includes(BOU_END))) {
              cc.delete(true);
          }
      }
      await context.sync();
    });
  } catch {}
}
