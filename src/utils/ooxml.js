import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区并安装物理锚点
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 清理旧标记
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

        for (let i = 0; i < matches.items.length; i++) {
           const m = matches.items[i];
           m.load("text");
           await context.sync();
           
           const currentId = globalCounter++;
           const placeholder = `{{REF_${currentId}}}`;
           
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${currentId}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
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
 * 填缝算法 4.0：串行物理分段构建
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

      // 1. 采集灵魂备份 (OOXML)
      const backups = new Map();
      for (const ref of refMap) {
          const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${ref.id}`);
          if (shield) {
              backups.set(ref.placeholder, { ooxml: shield.getOoxml(), obj: shield });
          }
      }
      await context.sync();

      // 2. 将 AI 结果按占位符拆分
      // 例: "Text0 {{REF_0}} Text1" -> ["Text0", "{{REF_0}}", "Text1"]
      const segments = aiResult.trim().split(/({{REF_\d+}})/g);

      // 3. 【串行物理缝合】：清空选区，重新构建
      const targetRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
      targetRange.clear(); // 物理清空
      await context.sync();

      // 我们使用这个 Range 的“开头”不断往里填
      let insertionPoint = startAnchor.getRange("After");

      for (const seg of segments) {
         if (!seg) continue;
         
         if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
             const backup = backups.get(seg);
             if (backup) {
                 // 缝入原始灵魂
                 insertionPoint.insertOoxml(backup.ooxml.value, Word.InsertLocation.before);
             }
         } else {
             // 填入润色文本
             insertionPoint.insertText(seg, Word.InsertLocation.before);
         }
         // 关键：插入后 insertionPoint 会保持在原位（Before 插入），
         // 但由于我们是顺序构建，我们希望能持续在“已插入内容的后面”插入。
         // 所以我们实际上应该在 endAnchor 的 "Before" 位置不断插入？
         // 不，最稳的做法是：每次 insert 之后，使用返回的 Range 的 "After" 作为下一个点。
      }

      // 修正后的串行构建逻辑：
      targetRange.clear();
      let cursor = startAnchor.getRange("After");
      for (const seg of segments) {
          if (!seg) continue;
          let nextRange;
          if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
              const backup = backups.get(seg);
              if (backup) {
                  nextRange = cursor.insertOoxml(backup.ooxml.value, Word.InsertLocation.replace);
              } else {
                  nextRange = cursor.insertText(seg, Word.InsertLocation.replace);
              }
          } else {
              nextRange = cursor.insertText(seg, Word.InsertLocation.replace);
          }
          await context.sync(); // 必须同步以获取 nextRange 的物理位置
          cursor = nextRange.getRange("After");
      }

      // 恢复格式
      if (baseFont) {
          const r = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }

      // 清理
      backups.forEach(b => b.obj.delete(true));
      startAnchor.delete(true);
      endAnchor.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement Error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在采集物理指纹...", true);
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
        if (onStatus) onStatus("processing", `正在串行分段缝合 (${i + 1}/${results.length})...`, true);
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
