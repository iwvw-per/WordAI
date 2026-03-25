import * as storage from "./storage.js";

const CC_TAG = "wordai_target";
const SHIELD_PREFIX = "wordai_shield_";

/**
 * 标记选区并安装“物理盾牌”
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 清理旧标记
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag === CC_TAG || (cc.tag && cc.tag.startsWith(SHIELD_PREFIX))) cc.delete(true);
      }
      await context.sync();

      let targetRanges = [];
      const isRealSelection = !selection.isEmpty && selection.text && selection.text.length > 0;

      if (!isRealSelection) {
        const paragraphs = selection.paragraphs;
        paragraphs.load(["items", "text", "style"]);
        await context.sync();
        if (paragraphs.items.length > 0) targetRanges = [paragraphs.items[0]];
      } else {
        targetRanges = [selection];
      }

      for (const range of targetRanges) {
        range.load(["text", "style", "font"]);
        await context.sync();
        const rangeText = range.text || "";
        if (!rangeText.trim()) continue;

        // 查找引用并记录 (使用通配符覆盖多种格式)
        const refPatterns = ["\\[[0-9.,\\- ]@\\]", "图 [0-9]@", "表 [0-9]@", "Fig. [0-9]@", "Figure [0-9]@", "Table [0-9]@"];
        const foundInstances = [];
        
        // 我们需要按在文档中的【物理出现的先后顺序】来安装盾牌
        const allMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        allMatches.load("items");
        await context.sync();

        for (let m of allMatches.items) {
           m.load("text");
        }
        await context.sync();

        const finalRefMap = [];
        const uniqueKeys = [];
        let tokenizedText = rangeText;

        for (let i = 0; i < allMatches.items.length; i++) {
           const m = allMatches.items[i];
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${i}`;
           shield.title = m.text; // 将原始文本存入 title，方便回写时切分
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           if (!uniqueKeys.includes(m.text)) uniqueKeys.push(m.text);
        }

        // 构建发给 AI 的占位符文本
        uniqueKeys.sort((a,b) => b.length - a.length).forEach((txt, idx) => {
           const placeholder = `{{REF_${idx}}}`;
           tokenizedText = tokenizedText.split(txt).join(placeholder);
           finalRefMap.push({ placeholder, original: txt });
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
 * 物理间隙原位填缝 (Surgical Gap-Filling)
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];

      const shields = cc.contentControls;
      shields.load("items");
      await context.sync();
      const shieldItems = shields.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      
      // 1. 还原 AI 结果到带原文引用的状态
      let reconstructedText = aiResult.trim();
      for (const refDef of refMap) {
          reconstructedText = reconstructedText.split(refDef.placeholder).join(refDef.original);
      }
      
      // 2. 将重建文本按物理采集时的 Shield 文本内容切分
      const textGaps = [];
      let tempText = reconstructedText;
      for (const s of shieldItems) {
          const pattern = s.title; // 获取我们之前存入的原始内容
          const idx = tempText.indexOf(pattern);
          if (idx >= 0) {
              textGaps.push(tempText.substring(0, idx));
              tempText = tempText.substring(idx + pattern.length);
          } else {
              textGaps.push(""); 
          }
      }
      textGaps.push(tempText);

      // 3. 【精准间隙更新】
      // 如果没有引用，直接替换
      if (shieldItems.length === 0) {
          cc.insertText(aiResult.trim(), Word.InsertLocation.replace);
      } else {
          // 最后一个间隙
          const lastGapRange = shieldItems[shieldItems.length-1].getRange("After").expandTo(cc.getRange("End"));
          lastGapRange.insertText(textGaps[textGaps.length - 1] || "", Word.InsertLocation.replace);

          // 中间及头部间隙 (逆序)
          for (let i = shieldItems.length - 1; i >= 0; i--) {
              const cur = shieldItems[i];
              const prev = shieldItems[i-1];
              if (prev) {
                  const gap = prev.getRange("After").expandTo(cur.getRange("Before"));
                  gap.insertText(textGaps[i] || "", Word.InsertLocation.replace);
              } else {
                  const firstGap = cc.getRange("Start").expandTo(cur.getRange("Before"));
                  firstGap.insertText(textGaps[0] || "", Word.InsertLocation.replace);
              }
          }
      }

      // 处理字体
      if (baseFont) {
          const r = cc.getRange();
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }

      // 清理盾牌并保留内部 Fields
      for (const s of shieldItems) s.delete(true);
      cc.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replace error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在分析选区...", true);
  const results = await markSelection();
  if (!results || results.length === 0) throw new Error("请先选择文字内容");

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
    const item = results[i];
    if (onStatus) onStatus("processing", `AI 正在处理 (${i + 1}/${results.length})...`, true);
    
    try {
      const aiResult = await processText(item.text);
      if (signal?.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在原位更新文本间隙 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.baseFont);
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
          if (cc.tag === CC_TAG || (cc.tag && cc.tag.startsWith(SHIELD_PREFIX))) {
              cc.delete(true);
          }
      }
      await context.sync();
    });
  } catch {}
}
