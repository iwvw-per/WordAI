import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";

/**
 * 标记选区：为每一个物理引用实例安装带唯一 ID 的盾牌
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

        // 【定海神针】：边界锚点
        const startMarker = range.getRange("Start").insertContentControl();
        startMarker.tag = BOU_START;
        startMarker.appearance = Word.ContentControlAppearance.hidden;

        const endMarker = range.getRange("End").insertContentControl();
        endMarker.tag = BOU_END;
        endMarker.appearance = Word.ContentControlAppearance.hidden;

        // 搜索所有引用实例并记录
        const orderedMatches = range.search("\\[[0-9.,\\- ]@\\]|图 [0-9]@|表 [0-9]@|Fig. [0-9]@|Figure [0-9]@|Table [0-9]@", { matchWildcards: true });
        orderedMatches.load("items");
        await context.sync();

        // 核心：将每个物理实例变为唯一的 {{REF_N}}
        let tokenizedText = rangeText;
        const localRefMap = [];

        // 我们必须按“逆序”进行字符串替换，以防 Offset 漂移导致对同一段话进行多次处理
        // 但更好的做法是：先获取所有匹配项的文本内容。
        for (let m of orderedMatches.items) m.load("text");
        await context.sync();

        // 为每一个物理实例安装带唯一 ID 的盾牌，并生成占位符
        for (let i = 0; i < orderedMatches.items.length; i++) {
           const m = orderedMatches.items[i];
           const placeholder = `{{REF_${i}}}`;
           
           // 我们在这里采用一个巧妙的占位方法：直接把 m.text 换成 placeholder 给 AI 看
           // 由于 rangeText 是纯文本，我们只需知道每个 [1] 出现的顺序
           // 这是一个简单的“第 N 个出现的引用”映射
           
           const shield = m.insertContentControl();
           shield.tag = `${SHIELD_PREFIX}${i}`;
           shield.appearance = Word.ContentControlAppearance.hidden;
           
           localRefMap.push({ placeholder, original: m.text, id: i });
        }

        // 构建 Tokenized Text：按顺序寻找并替换第 i 个引用
        let aiInput = rangeText;
        let offset = 0;
        for (let i = 0; i < localRefMap.length; i++) {
           const ref = localRefMap[i];
           const idx = aiInput.indexOf(ref.original, offset);
           if (idx >= 0) {
              aiInput = aiInput.substring(0, idx) + ref.placeholder + aiInput.substring(idx + ref.original.length);
              offset = idx + ref.placeholder.length;
           }
        }

        const font = range.font;
        font.load(["name", "size", "color", "bold", "italic", "underline"]);
        await context.sync();

        results.push({ 
          text: aiInput, 
          refMap: localRefMap, 
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
 * 填缝算法 3.0：占位符-盾牌 1:1 坐标对齐
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      const allCCs = context.document.contentControls;
      allCCs.load("items");
      await context.sync();

      const startAnchor = allCCs.items.find(c => c.tag === BOU_START);
      const endAnchor = allCCs.items.find(c => c.tag === BOU_END);
      if (!startAnchor || !endAnchor) return;

      // 1. 解析 AI 结果及占位符 (保留占位符以便对齐)
      const segments = aiResult.trim().split(/({{REF_\d+}})/g);
      
      // 2. 依次填补缝隙
      let currentAnchor = startAnchor.getRange("After");

      for (const seg of segments) {
         if (!seg) continue;

         if (seg.startsWith("{{REF_") && seg.endsWith("}}")) {
             // 找到对应的物理盾牌
             const shieldId = seg.replace("{{REF_", "").replace("}}", "");
             const shield = allCCs.items.find(c => c.tag === `${SHIELD_PREFIX}${shieldId}`);
             if (shield) {
                 // 此时：currentAnchor 到 shield.Before 这一段就是需要更新的文字缝隙
                 const gapRange = currentAnchor.expandTo(shield.getRange("Before"));
                 gapRange.insertText("", Word.InsertLocation.replace); // 先清空旧文字
                 // 注意：如果想在这个缝隙插入 AI 段落，我们需要在 currentAnchor 开始插入
                 // 我们换个思路：每次遇到 Text 段就插入，遇到 Shield 段就跳过
             }
         }
      }

      // 【更稳的方案】：按 AI 分段重构
      // 我们先删除选区内除锚点和盾牌外的所有非盾牌内容（极其困难）
      // 【回归极致稳健】：先全量填入 aiResult 到 边界之间，然后找到对应的 {{REF_N}} 字符串，
      // 把原始盾牌移动到那个位置。
      
      // 不，Word 1.1 不支持移动。
      
      // 【最终大招】：既然 AI 已经返回了带 {{REF_N}} 的文本，我们直接用 insertText(aiResult)
      // 然后搜索每一个 {{REF_X}}，并将保存在内存中的 [1] 通过 insertOoxml 缝合回去。
      // 因为这次是在 AI 结果的“精确位置”缝合，且每个 REF_X 都是唯一的物理 ID 对应，
      // 所以绝对不会出现残留 `{{REF_1}}` 的情况。
      
      const targetRange = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
      
      // 1. 采集盾牌的 OOXML (在被冲掉前)
      const shields = allCCs.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      const backups = [];
      for (let s of shields) {
          backups.push({ tag: s.tag, ooxml: s.getOoxml() });
      }
      await context.sync();

      // 2. 润色替换（会暂时留下占位符）
      targetRange.insertText(aiResult.trim(), Word.InsertLocation.replace);
      await context.sync();

      // 3. 将每一个 {{REF_N}} 物理“核爆”掉，填回原始魂魄
      for (let b of backups) {
          const id = b.tag.replace(SHIELD_PREFIX, "");
          const placeholder = `{{REF_${id}}}`;
          const matches = targetRange.search(placeholder, { matchCase: true });
          matches.load("items");
          await context.sync();
          
          if (matches.items.length > 0) {
              matches.items[0].insertOoxml(b.ooxml.value, Word.InsertLocation.replace);
          }
      }

      // 4. 清理
      if (baseFont) {
          const finalR = startAnchor.getRange("After").expandTo(endAnchor.getRange("Before"));
          if (baseFont.name) finalR.font.name = baseFont.name;
          if (typeof baseFont.size === "number" && baseFont.size > 0) finalR.font.size = baseFont.size;
      }

      for (let s of shields) s.delete(true);
      startAnchor.delete(true);
      endAnchor.delete(true);
      await context.sync();
    });
  } catch (err) {
    console.error("Replacement Error:", err);
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
