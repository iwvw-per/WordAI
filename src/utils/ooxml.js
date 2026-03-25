import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区：使用全局唯一 ID 锁定每一个物理引用
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 1. 全局清理
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag === BOU_START || cc.tag === BOU_END || (cc.tag && cc.tag.startsWith(SHIELD_PREFIX))) {
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

      // 【核心】：全局计数器，确保全文 ID 唯一
      let globalShieldCount = 0;

      for (const range of targetRanges) {
        range.load(["text", "font"]);
        await context.sync();
        const rangeText = range.text || "";
        if (!rangeText.trim()) continue;

        // 段落级锚点
        const startMarker = range.getRange("Start").insertContentControl();
        const paragraphSessionId = Date.now() + "_" + Math.floor(Math.random() * 1000);
        startMarker.tag = `${BOU_START}_${paragraphSessionId}`;
        startMarker.appearance = Word.ContentControlAppearance.hidden;

        const endMarker = range.getRange("End").insertContentControl();
        endMarker.tag = `${BOU_END}_${paragraphSessionId}`;
        endMarker.appearance = Word.ContentControlAppearance.hidden;

        // 搜索当前 Range 内的所有引用
        const orderedMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        orderedMatches.load("items");
        await context.sync();

        const localRefMap = [];
        let aiInput = rangeText;

        // 为段落内每个引用安装全局唯一盾牌
        for (let i = 0; i < orderedMatches.items.length; i++) {
           const m = orderedMatches.items[i];
           m.load("text");
           await context.sync();
           
           const absoluteId = globalShieldCount++;
           const placeholder = `{{REF_${absoluteId}}}`;
           
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${absoluteId}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           // 物理替换 aiInput 中的文字为占位符 (按顺序寻找避免冲突)
           const idx = aiInput.indexOf(m.text);
           if (idx >= 0) {
              aiInput = aiInput.substring(0, idx) + placeholder + aiInput.substring(idx + m.text.length);
           }
           
           localRefMap.push({ placeholder, original: m.text, id: absoluteId });
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
 * 坐标映射回填 3.1 版：绝对物理隔离回缝
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
      
      // 1. 采集当前段落内的盾牌数据
      const backups = [];
      for (const refDef of refMap) {
          const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${refDef.id}`);
          if (shield) {
              backups.push({ tag: shield.tag, ooxml: shield.getOoxml(), shieldObj: shield });
          }
      }
      await context.sync();

      // 2. 润色结果入场
      targetRange.insertText(aiResult.trim(), Word.InsertLocation.replace);
      await context.sync();

      // 3. 灵魂回填：搜索 {{REF_N}} 并炸开缝合
      for (const b of backups) {
          const id = b.tag.replace(SHIELD_PREFIX, "");
          const placeholder = `{{REF_${id}}}`;
          const matches = targetRange.search(placeholder, { matchCase: true });
          matches.load("items");
          await context.sync();
          
          if (matches.items.length > 0) {
              matches.items[0].insertOoxml(b.ooxml.value, Word.InsertLocation.replace);
          }
      }

      // 4. 重设格式
      if (baseFont) {
          const finalR = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) finalR.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalR.font.size = baseFont.size;
      }

      // 彻底清理
      for (const b of backups) b.shieldObj.delete(true);
      startAnchor.delete(true);
      endAnchor.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement Error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在锁定全局物理坐标...", true);
  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
    const item = results[i];
    if (onStatus) onStatus("processing", `AI 正在润色 (${i + 1}/${results.length})...`, true);
    
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在执行物理回缝 (${i + 1}/${results.length})...`, true);
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
          if (cc.tag.includes(SHIELD_PREFIX) || cc.tag.includes(BOU_START) || cc.tag.includes(BOU_END)) {
              cc.delete(true);
          }
      }
      await context.sync();
    });
  } catch {}
}
