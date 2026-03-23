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

    // 获取选区段落
    const skipRules = storage.getSkipRules();
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    
    paragraphs.load(["items", "style", "text"]);
    await context.sync();

    const isTableCheckSupported = Office.context.requirements.isSetSupported("WordApi", "1.3");
    const paragraphRanges = paragraphs.items.map(p => p.getRange());
    if (skipRules.tables && isTableCheckSupported) {
       paragraphRanges.forEach(r => r.load("parentTableCell"));
    }
    
    try { await context.sync(); } catch (e) {}

    const validParagraphs = [];
    const validRanges = [];

    for (let i = 0; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      const range = paragraphRanges[i];
      let text = "";
      try { text = p.text.trim(); } catch(e) {}
      if (!text) continue;

      if (skipRules.headings) {
        const styleName = (p.style || "").toString().toLowerCase();
        const isNumberedHeader = /^(\s*(\d+\.)+\d*\s)/.test(text);
        if (styleName.includes("heading") || styleName.includes("title") || styleName.includes("标题") || isNumberedHeader) {
          continue;
        }
      }

      if (skipRules.tables && isTableCheckSupported) {
        try { if (range.parentTableCell && !range.parentTableCell.isNullObject) continue; } catch(e) {}
      }

      validParagraphs.push(p);
      validRanges.push(range);
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
      let targetRange = validRanges[i];

      try {
        if (validParagraphs.length === 1) {
          targetRange = selection;
        } else if (i === 0) {
          targetRange = selection.getRange("Start").expandTo(p.getRange("Content").getRange("End"));
        } else if (i === validParagraphs.length - 1) {
          targetRange = p.getRange("Content").getRange("Start").expandTo(selection.getRange("End"));
        } else {
          targetRange = p.getRange("Content");
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
    const cleanText = newText.split(/\r?\n/).filter(l => l.trim()).join("\n");
    
    // 清空内容，准备注入
    cc.insertText("", Word.InsertLocation.replace);
    await context.sync();

    // 采用正则切分，保留占位符以便识别
    // 匹配 {{REF_(\d+)}}
    const parts = cleanText.split(/({{REF_\d+}})/g);

    for (const part of parts) {
      if (!part) continue;
      
      const match = part.match(/{{REF_(\d+)}}/);
      if (match) {
        const refItem = refMap.find(m => m.placeholder === part);
        if (refItem) {
          // 物理注入原始 OOXML，保留所有链接、域、格式
          cc.insertOoxml(refItem.ooxml, Word.InsertLocation.end);
        } else {
          // 如果 AI 伪造了不存在的占位符，当作普通文本
          cc.insertText(part, Word.InsertLocation.end);
        }
      } else {
        // 普通文本片段
        cc.insertText(part, Word.InsertLocation.end);
      }
      // 每一小块 insert 后可以 sync 一次确保顺序，或者积累一定量再 sync
      // 为了稳定，我们每一块都 sync
      await context.sync();
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
