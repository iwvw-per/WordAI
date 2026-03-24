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
    try {
      // ===== 第一步：获取选区段落（在清理旧CC之前，防止选区失效）=====
      const skipRules = storage.getSkipRules();
      const selection = context.document.getSelection();
      // 注意：getSelection() 已经返回 Range，不要再调用 .getRange()
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let workParagraphs = [...paragraphs.items]; // 复制到本地数组
      if (workParagraphs.length === 0) {
        try {
          const startPars = selection.getRange("Start").paragraphs;
          startPars.load("items");
          await context.sync();
          if (startPars.items.length > 0) {
            workParagraphs = [startPars.items[0]];
          }
        } catch (e) {}
      }

      if (workParagraphs.length === 0) return;

      // 加载段落属性
      for (let p of workParagraphs) {
        p.load(["text", "style"]);
      }
      await context.sync();

      // ===== 第二步：清理旧标记（选区信息已安全缓存）=====
      try {
        const existing = context.document.contentControls.getByTag(CC_TAG);
        existing.load("items");
        await context.sync();
        for (const cc of existing.items) cc.delete(true);
        await context.sync();
      } catch (e) {}

      // ===== 第三步：过滤段落 =====
      const isSingle = workParagraphs.length === 1;
      const resultParagraphs = [];

      for (let i = 0; i < workParagraphs.length; i++) {
        const p = workParagraphs[i];
        let text = "";
        try { text = p.text.trim(); } catch(e) {}
        if (!text) continue;

        const styleName = (p.style || "").toString().toLowerCase();
        const isHeaderPattern = /^(\s*\d+(\.\d+)*[\.\s\t])/.test(text) && text.length < 120;
        const isHeadingStyle = styleName.includes("heading") || styleName.includes("标题");

        let shouldSkip = false;
        if (skipRules.headings && !isSingle) {
          if (isHeadingStyle || isHeaderPattern) shouldSkip = true;
        }
        if (!shouldSkip) resultParagraphs.push(p);
      }

      // 兜底：全部被过滤则强制包含
      if (resultParagraphs.length === 0 && workParagraphs.length > 0) {
        for (let p of workParagraphs) {
          let text = "";
          try { text = p.text.trim(); } catch(e) {}
          if (text) resultParagraphs.push(p);
        }
      }

      // ===== 第四步：搜索引用并标记 =====
      const refSearches = [
        "\\[[0-9.,\\- ]@\\]",
        "图 [0-9]@",
        "表 [0-9]@",
        "Fig. [0-9]@",
        "Figure [0-9]@",
        "Table [0-9]@"
      ];

      for (let i = 0; i < resultParagraphs.length; i++) {
        const p = resultParagraphs[i];
        try {
          // 直接使用段落自身的 range，不调用 getRange()
          // 对于单段落场景，也使用段落 range 而非 selection
          p.load("text");
          await context.sync();

          const text = p.text.trim();
          if (!text) continue;

          const refMap = [];
          for (const pattern of refSearches) {
            try {
              const matches = p.search(pattern, { matchWildcards: true });
              matches.load("items");
              await context.sync();
              for (const matchRange of matches.items) {
                try {
                  matchRange.load(["text"]);
                  const ooxml = matchRange.getOoxml();
                  await context.sync();
                  refMap.push({ text: matchRange.text, ooxml: ooxml.value });
                } catch(e) {}
              }
            } catch (e) {}
          }

          let tokenizedText = text;
          const finalRefMap = [];
          const uniqueRefs = [];
          refMap.sort((a, b) => b.text.length - a.text.length).forEach(item => {
             if (!uniqueRefs.find(u => u.text === item.text)) uniqueRefs.push(item);
          });
          uniqueRefs.forEach((item, idx) => {
            const placeholder = `{{REF_${idx}}}`;
            tokenizedText = tokenizedText.split(item.text).join(placeholder);
            finalRefMap.push({ placeholder, ooxml: item.ooxml, original: item.text });
          });

          // 使用段落自身插入 ContentControl，而非 targetRange
          const cc = p.insertContentControl();
          cc.tag = CC_TAG;
          cc.appearance = Word.ContentControlAppearance.hidden;
          await context.sync();

          results.push({ text: tokenizedText, refMap: finalRefMap });
        } catch (err) {
          console.warn("Paragraph marking failed:", err);
        }
      }
    } catch (outerErr) {
      console.error("markSelection top-level error:", outerErr);
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
