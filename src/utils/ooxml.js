import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区：安装无级联风险的并列锚点链
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 1. 清理所有旧标记
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

        // 【最稳模式】：在选区最开头和最末尾分别钉上“定海神针”
        const startMarker = range.getRange("Start").insertContentControl();
        startMarker.tag = BOU_START;
        startMarker.appearance = Word.ContentControlAppearance.hidden;

        const endMarker = range.getRange("End").insertContentControl();
        endMarker.tag = BOU_END;
        endMarker.appearance = Word.ContentControlAppearance.hidden;

        // 搜索并录入引用物理盾牌
        const allMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        allMatches.load("items");
        await context.sync();

        for (let m of allMatches.items) m.load("text");
        await context.sync();

        const uniqueKeys = [];
        let tokenizedText = rangeText;

        for (let i = 0; i < allMatches.items.length; i++) {
           const m = allMatches.items[i];
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${i}`;
           shield.title = m.text; // 用于分段
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           if (!uniqueKeys.includes(m.text)) uniqueKeys.push(m.text);
        }

        // 占位符转换（长匹配优先）
        uniqueKeys.sort((a,b) => b.length - a.length).forEach((txt, idx) => {
           const placeholder = `{{REF_${idx}}}`;
           tokenizedText = tokenizedText.split(txt).join(placeholder);
           results.push({ placeholder, original: txt }); // 暂存用于回写匹配
        });

        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        // 将本段结果收集
        results.push({ 
          text: tokenizedText, 
          refMap: results.filter(r => r.placeholder), // 仅仅是 mapping
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
  // 过滤出干净的结果对象
  return results.filter(r => r.text);
}

/**
 * 无容器填缝方案：绝对物理隔离更新
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      // 1. 获取链上的所有锚点内容
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();

      const startAnchor = allCCs.items.find(c => c.tag === BOU_START);
      const endAnchor = allCCs.items.find(c => c.tag === BOU_END);
      const shieldItems = allCCs.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      
      if (!startAnchor || !endAnchor) return;

      // 2. 将 AI 结果解析为待填入的片段
      let reconstructedText = aiResult.trim();
      for (const refDef of refMap) {
          reconstructedText = reconstructedText.split(refDef.placeholder).join(refDef.original);
      }
      
      const textGaps = [];
      let tempText = reconstructedText;
      for (const s of shieldItems) {
          const pattern = s.title; 
          const idx = tempText.indexOf(pattern);
          if (idx >= 0) {
              textGaps.push(tempText.substring(0, idx));
              tempText = tempText.substring(idx + pattern.length);
          } else {
              textGaps.push(""); 
          }
      }
      textGaps.push(tempText);

      // 3. 【无容器填缝手术】：按反向顺序更新间隙
      // 如果没有引用盾牌：起始锚点 -> 结束锚点 之间的全部内容
      if (shieldItems.length === 0) {
          const totalGap = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          totalGap.insertText(aiResult.trim(), Word.InsertLocation.replace);
      } else {
          // 最后一个间隙：最后一个盾牌 -> 结束锚点
          const lastGap = shieldItems[shieldItems.length - 1].getRange("After").expandTo(endAnchor.getRange("Before"));
          lastGap.insertText(textGaps[textGaps.length - 1] || "", Word.InsertLocation.replace);

          // 中间及头部
          for (let i = shieldItems.length - 1; i >= 0; i--) {
              const cur = shieldItems[i];
              const prev = shieldItems[i-1];
              if (prev) {
                  const midGap = prev.getRange("After").expandTo(cur.getRange("Before"));
                  midGap.insertText(textGaps[i] || "", Word.InsertLocation.replace);
              } else {
                  const firstGap = startAnchor.getRange("After").expandTo(cur.getRange("Before"));
                  firstGap.insertText(textGaps[0] || "", Word.InsertLocation.replace);
              }
          }
      }

      // 最后清理锚点并恢复字体
      if (baseFont) {
          const finalRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) finalRange.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalRange.font.size = baseFont.size;
      }

      for (const s of shieldItems) s.delete(true);
      startAnchor.delete(true);
      endAnchor.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement error:", err);
  }
}

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在锁定物理锚点...", true);
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
        if (onStatus) onStatus("processing", `正在物理同步内容 (${i + 1}/${results.length})...`, true);
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
          if (cc.tag === BOU_START || cc.tag === BOU_END || (cc.tag && cc.tag.startsWith(SHIELD_PREFIX))) {
              cc.delete(true);
          }
      }
      await context.sync();
    });
  } catch {}
}
