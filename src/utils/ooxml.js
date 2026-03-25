import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区：确保全局唯一 ID 与 AI 占位符 1:1 对应
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      const allCCs = context.document.contentControls;
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

      // 全局计数器：关键！
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

        const endMarker = range.getRange("End").insertContentControl();
        endMarker.tag = `${BOU_END}_${sessionId}`;
        endMarker.appearance = Word.ContentControlAppearance.hidden;

        const matches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        matches.load("items");
        await context.sync();

        const localRefMap = [];
        let aiInput = rangeText;

        // 【修正核心】：占位符必须使用全局 ID
        for (let i = 0; i < matches.items.length; i++) {
           const m = matches.items[i];
           m.load("text");
           await context.sync();
           
           const currentId = globalCounter++;
           const placeholder = `{{REF_${currentId}}}`;
           
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${currentId}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           // 物理替换 aiInput
           const idx = aiInput.indexOf(m.text);
           if (idx >= 0) {
              aiInput = aiInput.substring(0, idx) + placeholder + aiInput.substring(idx + m.text.length);
           }
           
           localRefMap.push({ placeholder, original: m.text, id: currentId });
        }

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
 * 灵魂回填 3.2：全局 ID 映射缝合
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

      const targetRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
      
      // 1. 采集当前段落内的盾牌 OOXML
      const backups = [];
      for (const ref of refMap) {
          const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${ref.id}`);
          if (shield) {
              backups.push({ tag: shield.tag, ooxml: shield.getOoxml(), obj: shield });
          }
      }
      await context.sync();

      // 2. 更新文本 (暂时保留全球唯一占位符)
      targetRange.insertText(aiResult.trim(), Word.InsertLocation.replace);
      await context.sync();

      // 3. 全局唯一占位符回缝
      for (const b of backups) {
          const placeholder = `{{REF_${b.tag.replace(SHIELD_PREFIX, "")}}}`;
          const matches = targetRange.search(placeholder, { matchCase: true });
          matches.load("items");
          await context.sync();
          
          if (matches.items.length > 0) {
              matches.items[0].insertOoxml(b.ooxml.value, Word.InsertLocation.replace);
          }
      }

      // 4. 清理
      if (baseFont) {
          const r = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }

      for (const b of backups) b.obj.delete(true);
      startAnchor.delete(true);
      endAnchor.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replace Error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在同步全局物理锚点...", true);
  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
    const item = results[i];
    if (onStatus) onStatus("processing", `AI 正在深度处理 (${i + 1}/${results.length})...`, true);
    
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在执行物理级归位 (${i + 1}/${results.length})...`, true);
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
