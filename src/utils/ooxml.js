/**
 * ooxml.js - 格式保留替换引擎
 * 使用 ContentControl 标记重点选区 + 过滤杂音标题 + 修复上标
 */
import * as storage from "./storage.js";

const CC_TAG = "wordai_target";

// ==================== 标记选区 ====================

// ==================== 标记选区（增强：捕获引用范围的 OOXML） ====================

export async function markSelection() {
  let results = [];
  await Word.run(async (context) => {
    // 清理之前的标记
    const existing = context.document.contentControls.getByTag(CC_TAG);
    existing.load("items");
    await context.sync();
    for (const cc of existing.items) {
      cc.delete(true);
    }
    await context.sync();

    // 获取选区段落与规则
    const skipRules = storage.getSkipRules();
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    
    // 采用更保守的分步加载策略，确保在所有 Word 版本中稳定
    paragraphs.load("items");
    await context.sync();

    if (paragraphs.items.length === 0) {
      // 备选方案：通过获取选区起点附近的段落集合并查数，避免直接 .getFirst() 导致 ItemNotFound
      const fallbackPars = selection.getRange("Start").paragraphs;
      fallbackPars.load("items");
      await context.sync();
      
      if (fallbackPars.items.length > 0) {
         const p = fallbackPars.items[0];
         p.load(["text", "style"]);
         await context.sync();
         paragraphs.items = [p];
      }
    } else {
      // 显式加载每个 items 的属性
      for (let p of paragraphs.items) {
        p.load(["text", "style"]);
      }
      await context.sync();
    }

    const isTableCheckSupported = Office.context.requirements.isSetSupported("WordApi", "1.3");
    const paragraphRanges = paragraphs.items.map(p => p.getRange());
    if (skipRules.tables && isTableCheckSupported) {
       paragraphRanges.forEach(r => r.load("parentTableCell"));
    }
    
    await context.sync();

    const validParagraphs = [];
    const validRanges = [];

    const isSingleParagraphSelected = paragraphs.items.length === 1;
    let validParagraphData = [];

    for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const range = paragraphRanges[i];
        let text = "";
        try { text = p.text.trim(); } catch(e) {}
        if (!text) continue;

        const styleName = (p.style || "").toString().toLowerCase();
        // 改进：仅当文本较短且符合编号特征时，才认为可能是标题（防止误杀长列表项）
        const isHeaderPattern = /^(\s*\d+(\.\d+)*[\.\s\t])/.test(text) && text.length < 120;
        const isHeadingStyle = styleName.includes("heading") || 
                             styleName.includes("标题") || 
                             (styleName.includes("title") && !styleName.includes("subtitle") && !styleName.includes("副标题"));

        let shouldSkip = false;
        if (skipRules.headings && !isSingleParagraphSelected) {
            if (isHeadingStyle || isHeaderPattern) shouldSkip = true;
        }

        if (skipRules.tables && isTableCheckSupported) {
            try { 
                if (range.parentTableCell && !range.parentTableCell.isNullObject) {
                    shouldSkip = true;
                }
            } catch(e) {}
        }

        validParagraphData.push({ p, range, shouldSkip });
    }

    // --- 核心优化：兜底机制 ---
    // 如果所有选中的非空段落都被规则过滤掉了，则忽略过滤规则（强制包含它们），确保用户有反馈
    const allSkipped = validParagraphData.length > 0 && validParagraphData.every(d => d.shouldSkip);
    
    for (const data of validParagraphData) {
        if (!data.shouldSkip || allSkipped) {
            validParagraphs.push(data.p);
            validRanges.push(data.range);
        }
    }

    // 搜索引用用的通配符模式（Word 语法）
    const refSearches = [
      "\\[[0-9.,\\- ]@\\]", // [1], [1-3], [1, 2]
      "图 [0-9]@",           // 图 1
      "表 [0-9]@",           // 表 1
      "式([0-9]@)",         // 式(1)
      "Fig. [0-9]@",        // Fig. 1
      "Figure [0-9]@",      // Figure 1
      "Table [0-9]@",       // Table 1
      "Section [0-9]@"      // Section 1
    ];

    for (let i = 0; i < validParagraphs.length; i++) {
      const p = validParagraphs[i];
      let targetRange;

      try {
        if (validParagraphs.length === 1) {
          targetRange = selection;
        } else {
          // 为了避免 getRange("Content") 抛出 ItemNotFound，直接使用完整的 paragraph range
          targetRange = p.getRange();
        }

        targetRange.load("text");
        await context.sync();

        const text = targetRange.text.trim();
        if (!text) continue;

        // --- 核心：提取范围内的引用片段 OOXML ---
        const refMap = [];
        for (const pattern of refSearches) {
          try {
            const matches = targetRange.search(pattern, { matchWildcards: true });
            matches.load("items");
            await context.sync();
            for (const matchRange of matches.items) {
              matchRange.load(["text", "address"]); 
              const ooxml = matchRange.getOoxml();
              await context.sync();
              refMap.push({
                text: matchRange.text,
                ooxml: ooxml.value,
                // 用于排序，确保占位符顺序与文中一致
                address: matchRange.address 
              });
            }
          } catch (e) {}
        }

        // 按文中出现位置排序（根据 Word 的 range address 启发式排序，或者简单依靠 search 结果）
        // 这里简化处理：search 结果通常已经是有序的。如果不放心，可以根据 index 排序。
        // 由于我们将 refMap 用于 tokenize，我们需要确保它唯一
        
        let tokenizedText = text;
        const finalRefMap = [];
        
        // 倒序替换以防索引偏移（或者使用特殊 Token 避免二次匹配）
        // 这里我们直接用 detokenizeReferences 的逻辑，提前准备好 finalRefMap
        // 为确保精准，我们按长度从长到短匹配（防止 [1-3] 被匹配为 [1]）
        const uniqueRefs = [];
        refMap.sort((a, b) => b.text.length - a.text.length).forEach(item => {
           if (!uniqueRefs.find(u => u.text === item.text)) {
              uniqueRefs.push(item);
           }
        });

        uniqueRefs.forEach((item, idx) => {
          const placeholder = `{{REF_${idx}}}`;
          tokenizedText = tokenizedText.split(item.text).join(placeholder);
          finalRefMap.push({ placeholder, ooxml: item.ooxml, original: item.text });
        });

        const cc = targetRange.insertContentControl();
        cc.tag = CC_TAG;
        cc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();

        results.push({
          text: tokenizedText,
          refMap: finalRefMap
        });
      } catch (err) {
        console.warn("Paragraph marking failed:", err);
      }
    }
  });
  return results;
}

// ==================== 逐段回写（物理拼合文本与 OOXML） ====================

/**
 * 替换首个标记内容
 * @param {string} newText - LLM 返回的（含占位符）文字
 * @param {Array} refMap - 对应的占位符映射
 */
async function replaceSingleMarkedContent(newText, refMap) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(CC_TAG);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) return;

    const cc = ccs.items[0];
    
    // 1. 预处理：处理 literal \n 和 Markdown 标记
    let processedText = newText.replace(/\\n/g, "\n");
    const lines = processedText.split(/\r?\n/);
    
    // 清空内容，准备注入
    cc.insertText("", Word.InsertLocation.replace);
    await context.sync();

    let isFirstLine = true;

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      
      // 处理空行
      if (!line) {
        if (!isFirstLine && i !== lines.length - 1) {
          cc.insertParagraph("", Word.InsertLocation.end);
          await context.sync();
        }
        continue;
      }

      // 清理 Markdown 标题和列表标记
      line = line.replace(/^#{1,6}\s+/, "");
      // 如果不是普通的带数字列表（由于 Word 机制，列表最好转纯文本），清理掉破折号
      line = line.replace(/^-\s+/, "");

      // 插入换行
      if (!isFirstLine) {
        cc.insertParagraph("", Word.InsertLocation.end);
        await context.sync();
      }
      isFirstLine = false;

      // 采用正则切分，保留占位符以便识别
      const parts = line.split(/({{REF_\d+}})/g);

      for (const part of parts) {
        if (!part) continue;
        
        const match = part.match(/{{REF_(\d+)}}/);
        if (match) {
          const refItem = refMap.find(m => m.placeholder === part);
          if (refItem) {
            cc.insertOoxml(refItem.ooxml, Word.InsertLocation.end);
          } else {
            cc.insertText(part, Word.InsertLocation.end);
          }
        } else {
          // 进一步处理行内加粗 **text**
          const subParts = part.split(/(\*\*.*?\*\*)/g);
          for (const subPart of subParts) {
            if (subPart.startsWith("**") && subPart.endsWith("**")) {
              const boldText = subPart.substring(2, subPart.length - 2);
              const run = cc.insertText(boldText, Word.InsertLocation.end);
              run.font.bold = true;
            } else {
              cc.insertText(subPart, Word.InsertLocation.end);
            }
          }
        }
        await context.sync();
      }
    }

    cc.delete(true);
    await context.sync();
  });
}

// ==================== 批量串行执行 ====================

export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在分析选区...");
  
  const diffMode = storage.getDiffMode();
  if (diffMode) {
    await Word.run(async (context) => {
      context.document.changeTrackingMode = "TrackAll";
      await context.sync();
    });
  }

  const results = await markSelection();

  if (!results || results.length === 0) {
    throw new Error("无有效待处理段落");
  }

  for (let i = 0; i < results.length; i++) {
    if (signal?.aborted) {
      await clearMarks();
      throw new Error("已取消");
    }

    if (onStatus) {
      onStatus("processing", `AI 润色中 (${i + 1}/${results.length})...`, true);
    }

    const item = results[i];
    let aiResult = null;
    try {
      aiResult = await processText(item.text);
    } catch (err) {
      await clearMarks();
      throw err;
    }

    if (signal?.aborted) {
      await clearMarks();
      throw new Error("已取消");
    }

    if (aiResult && aiResult.trim()) {
      if (onStatus) onStatus("processing", `回写中 (${i + 1}/${results.length})...`, true);
      await replaceSingleMarkedContent(aiResult.trim(), item.refMap);
    } else {
      await Word.run(async (context) => {
        const ccs = context.document.contentControls.getByTag(CC_TAG);
        ccs.load("items");
        await context.sync();
        if (ccs.items.length > 0) {
           ccs.items[0].delete(true);
           await context.sync();
        }
      });
    }
  }

  return { original: "已处理选区", result: "全部完成" };
}

// ==================== 统一清理标记 ====================

export async function clearMarks() {
  try {
    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag(CC_TAG);
      ccs.load("items");
      await context.sync();
      for (const cc of ccs.items) {
        cc.delete(true);
      }
      await context.sync();
    });
  } catch {}
}
