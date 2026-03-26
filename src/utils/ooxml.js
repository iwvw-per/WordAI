import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";
const REF_BOOKMARK_PREFIX = "wordai_ref_";

/**
 * 标记选区并建立全局参考文献书签索引 (VBA 思想集成)
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const doc = context.document;
      const selection = doc.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 1. 全局准备：由于要建立书签，我们先清理旧的 wordai 书签（可选，保证干净）
      const bookmarks = doc.bookmarks;
      bookmarks.load("items");
      await context.sync();
      for (const bm of bookmarks.items) {
          if (bm.name.startsWith(REF_BOOKMARK_PREFIX)) {
              // 我们保留之前的书签以防本次未处理到，或者直接全量更新
          }
      }

      // 2. 扫描全文参考文献条目 (如行首的 [1]) 并建立书签
      const body = doc.body;
      const refParas = body.paragraphs.search("^\\[[0-9]+\\]", { matchWildcards: true });
      refParas.load("items");
      await context.sync();
      
      for (const p of refParas.items) {
          p.load("text");
      }
      await context.sync();

      for (const p of refParas.items) {
          const match = p.text.match(/^\[(\d+)\]/);
          if (match) {
              const num = match[1];
              doc.bookmarks.add(`${REF_BOOKMARK_PREFIX}${num}`, p);
          }
      }
      await context.sync();

      // 3. 处理选区内的锚点
      const allCCs = doc.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag && (cc.tag.includes(BOU_START) || cc.tag.includes(BOU_END) || cc.tag.includes(SHIELD_PREFIX))) {
              cc.delete(true);
          }
      }
      await context.sync();

      let targetRanges = [];
      if (selection.isEmpty) {
        const paragraphs = selection.paragraphs;
        paragraphs.load("items");
        await context.sync();
        if (paragraphs.items.length > 0) targetRanges = [paragraphs.items[0]];
      } else {
        targetRanges = [selection];
      }

      let globalCounter = 0;
      for (const range of targetRanges) {
        range.load(["text", "font"]);
        await context.sync();
        const rangeText = range.text || "";
        if (!rangeText.trim()) continue;

        const sessionId = Date.now() + "_" + Math.floor(Math.random() * 100);
        const startMarker = range.getRange("Start").insertContentControl();
        startMarker.tag = `${BOU_START}_${sessionId}`;
        startMarker.appearance = Word.ContentControlAppearance.hidden;

        // 查找引用实例进行 Range 锁定 (5.0 物理隔离)
        const orderedMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        orderedMatches.load("items");
        await context.sync();

        const localRefMap = [];
        let aiInput = rangeText;

        for (let i = 0; i < orderedMatches.items.length; i++) {
           const m = orderedMatches.items[i];
           m.load("text");
           m.track(); 
           await context.sync();
           
           const currentId = globalCounter++;
           const placeholder = `{{REF_${currentId}}}`;
           
           const shield = m.getRange("Start").insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${currentId}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           const idx = aiInput.indexOf(m.text);
           if (idx >= 0) {
              aiInput = aiInput.substring(0, idx) + placeholder + aiInput.substring(idx + m.text.length);
           }
           localRefMap.push({ placeholder, id: currentId });
        }

        const endMarker = range.getRange("End").insertContentControl();
        endMarker.tag = `${BOU_END}_${sessionId}`;
        endMarker.appearance = Word.ContentControlAppearance.hidden;

        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        results.push({ 
          text: aiInput, 
          refMap: localRefMap, 
          boundaryTags: { start: startMarker.tag, end: endMarker.tag },
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

/**
 * 后置自愈重链：确保 AI 修改过的引用也能跳转
 */
async function autoRelinkRange(range) {
    try {
        const matches = range.search("\\[[0-9]+\\]", { matchWildcards: true });
        matches.load("items");
        await range.context.sync();
        
        for (const m of matches.items) {
            m.load(["text", "hyperlink"]);
            await range.context.sync();
            
            // 如果已经有链接或 Field (Word JS 无法直接检测 Field，但 Hyperlink 可检)
            if (m.hyperlink) continue;
            
            const num = m.text.match(/\[(\d+)\]/)?.[1];
            if (num) {
                const bmName = `${REF_BOOKMARK_PREFIX}${num}`;
                const bookmarks = range.context.document.bookmarks;
                const bm = bookmarks.getAtOrNullObject(bmName);
                await range.context.sync();
                
                if (!bm.isNullObject) {
                    // 模拟 VBA：添加链接并移除格式
                    // Office.js 中内部书签链接使用 #name
                    m.insertHyperlink(`#${bmName}`, m.text, Word.InsertLocation.replace);
                    m.font.underline = Word.UnderlineStyle.none;
                    m.font.color = "black";
                }
            }
        }
        await range.context.sync();
    } catch (e) {
        console.warn("AutoRelink skip:", e);
    }
}

/**
 * 复合回写引擎 6.0：隔离填缝 + 后置自愈
 */
async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags, baseFont) {
  try {
    await Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();

      const startAnchor = allCCs.items.find(c => c.tag === boundaryTags.start);
      const endAnchor = allCCs.items.find(c => c.tag === boundaryTags.end);
      if (!startAnchor || !endAnchor) return;

      const segments = aiResult.trim().split(/({{REF_\d+}})/g);
      let currentDrawPoint = startAnchor.getRange("After");

      // 第一重保险：5.0 物理隔离更新
      for (const seg of segments) {
          if (!seg) continue;
          if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
              const shieldId = seg.match(/\d+/)[0];
              const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${shieldId}`);
              if (shield) currentDrawPoint = shield.getRange("After");
          } else {
              let nextAnchor = null;
              const nextRefIdx = segments.indexOf(seg) + 1;
              const nextRefMatch = segments[nextRefIdx];
              if (nextRefMatch && nextRefMatch.startsWith("{{REF_")) {
                  const nId = nextRefMatch.match(/\d+/)[0];
                  nextAnchor = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${nId}`);
              } else {
                  nextAnchor = endAnchor;
              }

              if (nextAnchor) {
                  const gapRange = currentDrawPoint.expandTo(nextAnchor.getRange("Before"));
                  gapRange.insertText(seg, Word.InsertLocation.replace);
                  await context.sync();
              }
          }
      }

      const finalRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
      
      // 第二重保险：后置自愈 (处理可能被 AI 修改出的新引文)
      await autoRelinkRange(finalRange);

      if (baseFont) {
          if (baseFont.name) finalRange.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalRange.font.size = baseFont.size;
      }

      for (const c of allCCs.items) {
          if (c.tag && (c.tag.includes(SHIELD_PREFIX) || c.tag === boundaryTags.start || c.tag === boundaryTags.end)) {
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
  if (onStatus) onStatus("processing", "正在构建全局文献指纹...", true);
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
        if (onStatus) onStatus("processing", `正在执行物理随动与自愈重链 (${i + 1}/${results.length})...`, true);
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
