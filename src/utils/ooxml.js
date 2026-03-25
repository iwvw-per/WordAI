import * as storage from "./storage.js";

const CC_TAG = "wordai_target";
const SHIELD_PREFIX = "wordai_shield_";

/**
 * 标记选区并为每一个物理引用安装“唯一物理盾牌”
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 1. 清理旧标记
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();
      for (const cc of allCCs.items) {
          if (cc.tag === CC_TAG || (cc.tag && cc.tag.startsWith(SHIELD_PREFIX))) {
              cc.delete(true);
          }
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

        // 搜索所有可能的引用模式
        const refPatterns = ["\\[[0-9.,\\- ]@\\]", "图 [0-9]@", "表 [0-9]@", "Fig. [0-9]@", "Figure [0-9]@", "Table [0-9]@"];
        const foundInstances = [];
        
        for (const pattern of refPatterns) {
          const matches = range.search(pattern, { matchWildcards: true });
          matches.load("items");
          await context.sync();
          for (const m of matches.items) {
             foundInstances.push(m);
          }
        }

        // 重要：我们必须按在文档中的物理出现顺序来标记 Shields
        // 为了确保顺序，我们暂存这些 Range 及其初始文本
        for (const m of foundInstances) m.load("text");
        await context.sync();
        
        // 我们通过 insertContentControl 的返回顺序来天然维持顺序（Word API 的集合通常是有序的）
        // 但为了 100% 保险，我们先给它们打上物理 ID。
        // AI 占位符只需要和原始文本对齐即可。
        
        const finalRefMap = [];
        const uniqueTexts = [];
        foundInstances.forEach(m => {
           if (!uniqueTexts.includes(m.text)) uniqueTexts.push(m.text);
        });
        uniqueTexts.sort((a,b) => b.length - a.length);
        
        let tokenizedText = rangeText;
        uniqueTexts.forEach((txt, idx) => {
           const placeholder = `{{REF_${idx}}}`;
           // 注意：这里是对发给 AI 的全文进行占位替换
           tokenizedText = tokenizedText.split(txt).join(placeholder);
           finalRefMap.push({ placeholder, original: txt });
        });

        // 【核心：物理锁定】
        // 我们按发现的顺序安装盾牌，不管文字是不是一样的
        // 这样在回写时，我们可以根据盾牌的物理顺序来对齐 AI 分段
        let shieldCount = 0;
        const orderedMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        orderedMatches.load("items");
        await context.sync();
        
        for (let m of orderedMatches.items) {
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${shieldCount++}`; // 每个实例唯一的物理 ID
           shield.appearance = Word.ContentControlAppearance.hidden;
        }

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
 * 物理间隙填缝算法：1:1 对等物理恢复
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];

      // 1. 获取所有物理盾牌（有序）
      const shields = cc.contentControls;
      shields.load("items");
      await context.sync();
      const shieldItems = shields.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      
      // 2. 将 AI 结果按引用的逻辑顺序还原
      // 注意：AI 结果中可能有 {{REF_0}}。我们需要将其替换回原始文本
      // 这里的策略是：把 AI 的润色结果变回“带原文原引用的全文”，但我们仅分段填入间隙。
      let reconstructedText = aiResult.trim();
      for (const refDef of refMap) {
          reconstructedText = reconstructedText.split(refDef.placeholder).join(refDef.original);
      }
      
      // 3. 【分段填缝】：按物理盾牌将 reconstructedText 切碎
      // 假设：AI 没有增加、删除或重新排列引用。
      // 我们通过 shieldItems 的原始 text 来进行分段。
      const splitPatterns = shieldItems.map(s => s.placeholderText); // 原文引用的列表 [ "[1]", "[2]" ]
      
      const textGaps = [];
      let tempText = reconstructedText;
      for (const pattern of splitPatterns) {
          const idx = tempText.indexOf(pattern);
          if (idx >= 0) {
              textGaps.push(tempText.substring(0, idx));
              tempText = tempText.substring(idx + pattern.length);
          } else {
              // 防御：如果 AI 把 [1] 删了或改了，尝试模糊匹配？
              // 暂先直接推入，保活优先
              textGaps.push(""); 
          }
      }
      textGaps.push(tempText); // 最后一个末尾缝隙

      // 4. 【倒序填空】：物理级不动引用
      // 后到前：Gap N
      const lastRange = (shieldItems.length > 0) 
          ? shieldItems[shieldItems.length - 1].getRange("After").expandTo(cc.getRange("End"))
          : cc.getRange();
      lastRange.insertText(textGaps[textGaps.length - 1] || "", Word.InsertLocation.replace);

      // 前到中
      for (let i = shieldItems.length - 1; i >= 0; i--) {
          const cur = shieldItems[i];
          const prev = shieldItems[i-1];
          if (prev) {
              const gap = prev.getRange("After").expandTo(cur.getRange("Before"));
              gap.insertText(textGaps[i] || "", Word.InsertLocation.replace);
          } else if (shieldItems.length > 0) {
              const firstGap = cc.getRange("Start").expandTo(cur.getRange("Before"));
              firstGap.insertText(textGaps[0] || "", Word.InsertLocation.replace);
          }
      }

      const r = cc.getRange();
      if (baseFont) {
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }

      for (const s of shieldItems) s.delete(true);
      cc.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement Failure:", err);
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
        if (onStatus) onStatus("processing", `正在执行原位填缝 (${i + 1}/${results.length})...`, true);
        await replaceSingleMarkedContent(aiResult, item.refMap, item.baseFont);
      }
    } catch (err) {
      console.error(err);
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
