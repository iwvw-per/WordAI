import * as storage from "./storage.js";

const CC_TAG = "wordai_target";
const SHIELD_PREFIX = "wordai_shield_";

/**
 * 标记选区并加装引用“物理盾牌”
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 清理旧标记（严谨 load）
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
        if (paragraphs.items.length > 0) {
          targetRanges = [paragraphs.items[0]];
        }
      } else {
        targetRanges = [selection];
      }

      for (const range of targetRanges) {
        range.load(["text", "style", "font"]);
        await context.sync();
        const rangeText = range.text || "";
        if (!rangeText.trim()) continue;

        // 搜索引用并记录
        const refSearches = ["\\[[0-9.,\\- ]@\\]", "图 [0-9]@", "表 [0-9]@", "Fig. [0-9]@", "Figure [0-9]@", "Table [0-9]@"];
        const foundRefs = [];
        for (const pattern of refSearches) {
          const matches = range.search(pattern, { matchWildcards: true });
          matches.load("items");
          await context.sync();
          for (const m of matches.items) {
             foundRefs.push(m);
          }
        }
        
        // 分配物理盾牌（用于回写时定位）
        // 我们必须按偏移量排序，以保证 REF_0, REF_1 的顺序在文档中是递增的
        // Word 1.1 中 item.offset 不可用，我们使用 insertContentControl 顺序来推断
        const finalRefMap = [];
        let tokenizedText = rangeText;

        // 整理占位符
        const sortedUnique = [];
        foundRefs.forEach(r => {
           if (!sortedUnique.find(u => u.text === r.text)) sortedUnique.push({ text: r.text });
        });
        sortedUnique.sort((a,b) => b.text.length - a.text.length);

        sortedUnique.forEach((item, idx) => {
          const placeholder = `{{REF_${idx}}}`;
          tokenizedText = tokenizedText.split(item.text).join(placeholder);
          finalRefMap.push({ placeholder, original: item.text, id: idx });
        });

        // 加装盾牌：这步是“物理固定”
        // 我们需要重新遍历，因为我们需要真实的实例
        const actualShields = [];
        for (const def of finalRefMap) {
            const matches = range.search(def.original, { matchCase: true });
            matches.load("items");
            await context.sync();
            for (const m of matches.items) {
               const s = m.insertContentControl();
               s.tag = `${SHIELD_PREFIX}${def.id}`;
               s.appearance = Word.ContentControlAppearance.hidden;
            }
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
 * 物理间隙填缝算法：彻底不触碰引用 Range
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];

      // 1. 解析 AI 结果及占位符
      const aiSegments = aiResult.trim().split(/({{REF_\d+}})/);
      const textSegments = aiSegments.filter((s, i) => i % 2 === 0);
      const refPlaceholders = aiSegments.filter((s, i) => i % 2 === 1);

      // 2. 获取文档中的盾牌 CC (按物理顺序)
      const shields = cc.contentControls;
      shields.load("items");
      await context.sync();
      
      const shieldItems = shields.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      
      // 3. 【逆序替换间隙】
      // 逆序是为了保证前面的 Offset 指向不失效
      
      // Gap Last: 最后一个盾牌之后到选区结束
      let lastShield = shieldItems[shieldItems.length - 1];
      if (lastShield) {
          const afterRange = lastShield.getRange("After").expandTo(cc.getRange("End"));
          afterRange.insertText(textSegments[textSegments.length - 1] || "", Word.InsertLocation.replace);
      } else {
          // 没有引用，直接全量替换
          cc.insertText(aiResult.trim(), Word.InsertLocation.replace);
      }

      // Gap Middle & Initial
      for (let i = shieldItems.length - 1; i >= 0; i--) {
          const curShield = shieldItems[i];
          const prevShield = shieldItems[i - 1];
          
          if (prevShield) {
              // 两个盾牌之间的间隙
              const gapRange = prevShield.getRange("After").expandTo(curShield.getRange("Before"));
              gapRange.insertText(textSegments[i] || "", Word.InsertLocation.replace);
          } else if (shieldItems.length > 0) {
              // 第一个盾牌之前的间隙
              const firstGap = cc.getRange("Start").expandTo(curShield.getRange("Before"));
              firstGap.insertText(textSegments[0] || "", Word.InsertLocation.replace);
          }
      }

      const r = cc.getRange();
      if (baseFont) {
          if (baseFont.name) r.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) r.font.size = baseFont.size;
      }
      
      // 删除所有标记
      for (const s of shieldItems) s.delete(true);
      cc.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replace Error:", err);
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
      if (signal.aborted) { await clearMarks(); throw new Error("已取消"); }
      if (aiResult && aiResult.trim()) {
        if (onStatus) onStatus("processing", `正在手术式填缝 (${i + 1}/${results.length})...`, true);
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
