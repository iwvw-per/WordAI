import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区：获取物理引用的 Range 本身，不加干扰
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      // 清理旧标记（仅清理辅助锚点）
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
        
        // 关键：在选区最开头安装“定海锚点”，这是唯一需要的 CC，且不包含其它内容
        const startMarker = range.getRange("Start").insertContentControl();
        startMarker.tag = `${BOU_START}_${sessionId}`;
        startMarker.appearance = Word.ContentControlAppearance.hidden;

        // 查找引用 Range，并使用 track() 锁定它们
        const orderedMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        orderedMatches.load("items");
        await context.sync();

        const localRefMap = [];
        let aiInput = rangeText;

        for (let i = 0; i < orderedMatches.items.length; i++) {
           const m = orderedMatches.items[i];
           m.load("text");
           m.track(); // 锁定这个 Range，让它在文字更新时自动随动
           await context.sync();
           
           const currentId = globalCounter++;
           const placeholder = `{{REF_${currentId}}}`;
           
           // 为引用加一个极小的 Shield 仅仅为了标记编号，绝不包裹引用文本
           // 我们在这里采用“之前”插入法来安置盾牌锚点
           const shield = m.getRange("Start").insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${currentId}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           const idx = aiInput.indexOf(m.text);
           if (idx >= 0) {
              aiInput = aiInput.substring(0, idx) + placeholder + aiInput.substring(idx + m.text.length);
           }
           localRefMap.push({ placeholder, id: currentId });
        }

        // 选区末尾锚点
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
 * 零触碰手术填缝 5.0：分步同步实现物理随动
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
      
      // 我们通过盾牌锚点来寻找每一个间隙的物理地址
      let currentDrawPoint = startAnchor.getRange("After");

      for (const seg of segments) {
          if (!seg) continue;
          
          if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
              const shieldId = seg.match(/\d+/)[0];
              const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${shieldId}`);
              if (shield) {
                  // 这个 Shield 的 Before 就是当前文字间隙的终点
                  // 我们在这里不触碰 Shield 以后的内容
                  currentDrawPoint = shield.getRange("After");
              }
          } else {
              // 文本段落：我们要更新到下一个 Shield 之前，或者选区末尾之前
              // 但最稳的方法是：直接在 currentDrawPoint "Before" 插入并删除旧内容？
              // 不，我们这样做：
              // 下一个物理点是谁？
              let nextAnchor = null;
              const nextRefMatch = segments[segments.indexOf(seg) + 1];
              if (nextRefMatch && nextRefMatch.startsWith("{{REF_")) {
                  const nId = nextRefMatch.match(/\d+/)[0];
                  nextAnchor = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${nId}`);
              } else {
                  nextAnchor = endAnchor;
              }

              if (nextAnchor) {
                  const gapRange = currentDrawPoint.expandTo(nextAnchor.getRange("Before"));
                  gapRange.insertText(seg, Word.InsertLocation.replace);
                  // 【分步同步】：这是 5.0 的灵魂
                  // 同步后，Word 物理引擎会重新计算文档各处 Range 的位置
                  // 我们的 shield 和 引用 此时只是在内存中平移了位置，没有被干掉
                  await context.sync();
              }
          }
      }

      // 最后设置字体
      if (baseFont) {
          const finalRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) finalRange.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalRange.font.size = baseFont.size;
      }

      // 清理锚点
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
  if (onStatus) onStatus("processing", "正在进行物理指纹锁定...", true);
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
        if (onStatus) onStatus("processing", `分步同步手术填坑中 (${i + 1}/${results.length})...`, true);
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
