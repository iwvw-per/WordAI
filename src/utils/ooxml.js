import * as storage from "./storage.js";

const CC_TAG = "wordai_target";
const SHIELD_PREFIX = "wordai_shield_";

/**
 * 标记选区并加装引用“盾牌”
 */
export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    try {
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      selection.load(["text", "isEmpty", "style"]);
      await context.sync();

      // 1. 清理旧标记（严谨顺序）
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

        // 搜索引用并加装物理盾牌 (ContentControl)
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
        
        // 按文档顺序排序，方便后续计算间隙
        // 注意：FoundRefs 本身不带偏移量，我们需要用 r.getRange() 来准确定位
        const finalRefMap = [];
        let tokenizedText = rangeText;

        // 加装盾牌并构建占位符
        // 为了防止重复匹配，我们使用一个唯一且非中文的替换逻辑
        const sortedRefs = foundRefs.sort((a, b) => b.text.length - a.text.length);
        const uniqueItems = [];
        sortedRefs.forEach(r => {
           if (!uniqueItems.find(u => u.text === r.text)) uniqueItems.push({ text: r.text });
        });

        uniqueItems.forEach((item, idx) => {
          const placeholder = `{{REF_${idx}}}`;
          tokenizedText = tokenizedText.split(item.text).join(placeholder);
          finalRefMap.push({ placeholder, original: item.text, id: idx });
        });

        // 在正式回写前，我们需要在物理上把这些引用包裹起来作为“锚点”
        // 这里只是为了后续回写时识别 Boundary
        for (const m of foundRefs) {
            const def = finalRefMap.find(d => d.original === m.text);
            if (def) {
               const shield = m.insertContentControl();
               shield.tag = `${SHIELD_PREFIX}${def.id}`;
               shield.appearance = Word.ContentControlAppearance.hidden;
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
 * 外科手术式回写：仅替换引用间的间隙
 */
async function replaceSingleMarkedContent(aiResult, refMap, baseFont) {
  try {
    await Word.run(async (context) => {
      // 1. 获取主容器和所有盾牌
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      if (ccs.items.length === 0) return;
      const cc = ccs.items[0];

      const allShields = cc.contentControls;
      allShields.load("items");
      await context.sync();
      
      const shields = allShields.items.filter(s => s.tag && s.tag.startsWith(SHIELD_PREFIX));
      // 按物理顺序排列
      // 因为 shields 带 Range 信息，我们可以直接按偏移排序
      // 但其实 Word API 的 collection 默认顺序通常就是文档物理顺序
      
      // 2. 解析 AI 结果分段
      const segments = aiResult.trim().split(/({{REF_\d+}})/g).filter(s => s !== "");

      // 3. 【极致稳健策略】：
      // 如果 AI 没有重新排列引用顺序（绝大多数学术润色情况），
      // 我们直接使用“非破坏性顺序填入”。
      
      // 当前最稳的方法：逐段构造 HTML 片段然后再插。
      // 但用户说不需要重排，只要 [1] 呆在原地。
      
      // 我们换个思路：如果我们就直接把 [1] 占位符填回去，会有什么后果？
      // 如果我们用 insertText，[1] 确实会变成纯文本。
      
      // 【终极方案】：不重插文字，而是直接修改 Range。
      // 为每一个盾牌 CC 设定对应的文本缓冲区。
      
      // 这里的逻辑需要极高精度：
      // A [1] B [2] C
      // Segments: [AI_A, {{REF_0}}, AI_B, {{REF_1}}, AI_C]
      
      // 给每一个文字段落（间隙）分配 Range 并 update。
      
      let cursor = cc.getRange("Start");
      for (let seg of segments) {
         if (seg.startsWith("{{REF_")) {
             // 这是一个引用。找到对应的盾牌，跳过它。
             const id = seg.replace("{{REF_", "").replace("}}", "");
             const scc = shields.find(s => s.tag === `${SHIELD_PREFIX}${id}`);
             if (scc) {
                 cursor = scc.getRange("After");
                 // 此时偏移已跳过 [1]
             }
         } else {
             // 这是纯文本。我们在当前光标位置插入 AI 文本。
             // 如果原本这里有旧文本，我们首先要“清除”到下一个盾牌为止的区域。
             // 这是一个极大的挑战。
             
             // 简化法：我们先用传统的 insertText 填回 [占位符]，然后搜索占位符，
             // 此时占位符还在原位，我们将 Shields 的 Range 对应到占位符上。
             
             // 不，这还是会造成删除。
         }
      }

      // 【回归用户最易理解的解法】：
      // 既然用户说 [1] 不要动，那我们就只替换 [1] 前后的文字段。
      
      let finalOutput = aiResult.trim();
      for (const ref of refMap) {
          finalOutput = finalOutput.split(ref.placeholder).join(ref.original);
      }
      
      // 关键：我们不能调用 cc.insertText(..., "Replace")。
      // 这会删除内部所有 CC。
      // 我们调用 cc.insertText(..., "Before") 后直接删除旧的内容？也不行。
      
      // 目前唯一的 100% 保活跳转的办法是：
      // 不管三七二十一，先用 insertText(finalOutput, "Replace")
      // 然后对 [1] 进行“二次手术”：找到这个新插入的文本 [1]，
      // 此时它已经失去了跳转功能。我们用原始 Shields 里的内容（OOXML）把它【替换回来】。
      // 为什么这种比之前的“全文 OOXML 还原”更稳？
      // 因为我们这次只还原【非常小的一个点】，Word 对这种局部 OOXML 替换的 fieldId 匹配极其宽容。
      
      // 1. 先把全选区变成润色后的样子（此时引用变成纯文本了）
      cc.insertText(finalOutput, Word.InsertLocation.replace);
      await context.sync();
      
      // 2. 局部原封不动“神还原”
      for (const ref of refMap) {
          const sites = cc.search(ref.original, { matchCase: true });
          sites.load("items");
          await context.sync();
          for (const site of sites.items) {
             // 只要之前 markSelection 时备份了 OOXML
             // 注意：我们在 markSelection 里没存 OOXML，我们要现在补上
          }
      }
      
      // 既然局部同步极其难，我们换个最稳的：
      // 重新启用 search -> insertOoxml(原始内容) 策略，
      // 但这次我们在 markSelection 时对原始引用范围调用 getOoxml!
      
      cc.delete(true);
      await context.sync();
    });
  } catch (err) {}
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
        if (onStatus) onStatus("processing", `正在手术式回写 (${i + 1}/${results.length})...`, true);
        
        await Word.run(async (context) => {
           const ccList = context.document.contentControls.getByTag(CC_TAG);
           ccList.load("items");
           await context.sync();
           if (ccList.items.length === 0) return;
           const cc = ccList.items[0];

           // 1. 获取所有盾牌的 OOXML 备份 (重要：必须在 delete 之前获取)
           const shields = cc.contentControls;
           shields.load("items");
           await context.sync();
           
           const refBackups = [];
           for (let s of shields.items) {
               if (s.tag && s.tag.startsWith(SHIELD_PREFIX)) {
                  refBackups.push({ tag: s.tag, ooxml: s.getOoxml(), text: s.placeholderText }); // 这里用 getOoxml!
               }
           }
           await context.sync();

           // 2. 润色替换（会暂时冲掉跳转）
           let output = aiResult.trim();
           for (const ref of item.refMap) {
               output = output.split(ref.placeholder).join(ref.original);
           }
           cc.insertText(output, Word.InsertLocation.replace);
           await context.sync();

           // 3. 【局部还原手术】：找回每个引用的“灵魂” (Field 结构)
           for (let backup of refBackups) {
               const placeholderText = item.refMap.find(m => `${SHIELD_PREFIX}${m.id}` === backup.tag)?.original;
               if (placeholderText) {
                   const site = cc.search(placeholderText, { matchCase: true });
                   site.load("items");
                   await context.sync();
                   if (site.items.length > 0) {
                      // 将存下的完整 Field 结构插回这一小块地方
                      site.items[0].insertOoxml(backup.ooxml.value, Word.InsertLocation.replace);
                   }
               }
           }

           if (item.baseFont) {
               const r = cc.getRange();
               if (item.baseFont.name) r.font.name = item.baseFont.name;
               if (typeof item.baseFont.size === "number" && item.baseFont.size > 0) r.font.size = item.baseFont.size;
           }
           
           cc.delete(true);
           await context.sync();
        });
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
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) cc.delete(true);
      await context.sync();
    });
  } catch {}
}
