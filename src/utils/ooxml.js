/**
 * ooxml.js - 格式保留替换引擎
 * 使用 ContentControl 标记重点选区 + 过滤杂音标题 + 修复上标
 */

const CC_TAG = "wordai_target";

// ==================== 标记选区 ====================

export async function markSelection() {
  let texts = [];
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
    const selection = context.document.getSelection();
    const paragraphs = selection.paragraphs;
    paragraphs.load(["items", "styleBuiltIn", "style", "text"]);
    await context.sync();

    // 过滤出非标题段落
    const validParagraphs = [];
    for (const p of paragraphs.items) {
      const styleBuiltIn = p.styleBuiltIn.toString().toLowerCase();
      const styleName = p.style.toString().toLowerCase();
      if (styleBuiltIn.includes("heading") || styleBuiltIn.includes("title") || styleName.includes("标题")) {
        continue;
      }
      validParagraphs.push(p);
    }

    // 逐段处理，使用边界裁剪确保只处理用户实际选中的文本
    for (let i = 0; i < validParagraphs.length; i++) {
      const p = validParagraphs[i];

      try {
        let targetRange;

        if (validParagraphs.length === 1) {
          // 只有一个段落：直接使用用户的选区范围（最精准，绝不碰未选中的字）
          targetRange = selection;
        } else if (i === 0) {
          // 首段：从选区起点 → 段落内容末尾（裁剪掉段落中选区之前的文字）
          try {
            targetRange = selection.getRange("Start").expandTo(p.getRange("Content").getRange("End"));
          } catch {
            targetRange = p.getRange("Content");
          }
        } else if (i === validParagraphs.length - 1) {
          // 末段：从段落内容起点 → 选区终点（裁剪掉段落中选区之后的文字，如交叉引用上标）
          try {
            targetRange = p.getRange("Content").getRange("Start").expandTo(selection.getRange("End"));
          } catch {
            targetRange = p.getRange("Content");
          }
        } else {
          // 中间段落：完全被选区包含，直接使用全段
          targetRange = p.getRange("Content");
        }

        targetRange.load("text");
        await context.sync();

        const text = targetRange.text.trim();
        if (!text) continue;

        let validText = text;

        // 如果选区内仍有尾部引用，尝试搜寻并收缩选中区域跳过它
        const tailRegex = /((?:\[[^\]]+\]|【[^】]+】)[\s。.,，;；、]*)+$/;
        const match = text.match(tailRegex);

        if (match && match[0].length < text.length && match[0].length < 200) {
          const tailText = match[0];
          try {
            const searchResults = targetRange.search(tailText, { matchWholeWord: false, matchCase: false });
            searchResults.load("items");
            await context.sync();

            if (searchResults.items.length > 0) {
              const lastMatch = searchResults.items[searchResults.items.length - 1];
              targetRange = targetRange.getRange("Start").expandTo(lastMatch.getRange("Before"));
              validText = text.substring(0, text.length - tailText.length);
            }
          } catch (e) {
            // search 报错则保持 targetRange 不变
          }
        }

        // 插入标记框
        const cc = targetRange.insertContentControl();
        cc.tag = CC_TAG;
        cc.title = "";
        cc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
        texts.push(validText);
      } catch (err) {
        // 极端兜底：如果裁剪也失败，尝试框住整个段落
        try {
          const fallbackRange = p.getRange("Content");
          fallbackRange.load("text");
          await context.sync();
          const fbText = fallbackRange.text.trim();
          if (!fbText) continue;
          const cc = fallbackRange.insertContentControl();
          cc.tag = CC_TAG;
          cc.title = "";
          cc.appearance = Word.ContentControlAppearance.hidden;
          await context.sync();
          texts.push(fbText);
        } catch (fatalErr) {
          console.warn("Paragraph marking failed:", fatalErr);
        }
      }
    }
  });
  return texts;
}

// ==================== 逐段回写（保留段落格式，恢复引用） ====================

/**
 * 替换首个标记内容，并恢复其中的上标
 * @param {string} newText - LLM 返回的新文字
 */
async function replaceSingleMarkedContent(newText) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag(CC_TAG);
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0) return;

    const cc = ccs.items[0];
    const newLines = newText.split(/\r?\n/).filter((line) => line.trim().length > 0);

    // 行内 ContentControl 不允许通过 API 直接 insertParagraph 插入块级硬回车段落，否则将抛出 InvalidArgument 错误
    // 因此使用换行符 \n 拼合在一起进行同选区替换（视觉表现为段内的垂直软换行）
    const fullText = newLines.join("\n");
    cc.insertText(fullText, Word.InsertLocation.replace);
    await context.sync();

    // 在当前最新插入完成的范围内，全局搜索匹配形如 [1], [13] 这种样式的引用标记，并施加上标
    // 注意：Word 原生通配符不支持标准的 '+' 取而代之的是 '@'
    try {
      const searchResults = cc.search("\\[[0-9]@\\]", { matchWildcards: true });
      searchResults.load("items");
      await context.sync();
      for (const res of searchResults.items) {
        res.font.superscript = true;
      }
    } catch {
      // 忽略部分不支持通配符环境的报错
    }

    // 仅删除处理完毕的这个 CC
    cc.delete(true);
    await context.sync();
  });
}

// ==================== 批量串行执行（含状态展示） ====================

/**
 * 逐段执行：跳过标题 → 每个正文段落轮流给 LLM 处理并依次插入
 * @param {function} processText - (text) => Promise<string>
 * @param {function} onStatus - 状态回调 (type, message)
 * @param {AbortSignal} signal - 取消信号
 * @returns {Promise<{original: string, result: string}>}
 */
export async function executeAndReplace(processText, onStatus, signal) {
  if (onStatus) onStatus("processing", "正在分析并切分选区(过滤标题)...");
  
  const texts = await markSelection();

  if (!texts || texts.length === 0) {
    throw new Error("请先选中正文段落（已自动排除标题）");
  }

  // 逐段处理以保持独立上下文，防止长文报错，亦能实时更新进度
  for (let i = 0; i < texts.length; i++) {
    if (signal?.aborted) {
      await clearMarks();
      throw new Error("已取消");
    }

    if (onStatus) {
      onStatus("processing", `AI 润色中 (${i + 1}/${texts.length})...`, true);
    }

    let result = null;
    try {
      result = await processText(texts[i]);
    } catch (err) {
      await clearMarks();
      throw err;
    }

    if (signal?.aborted) {
      await clearMarks();
      throw new Error("已取消");
    }

    if (result && result.trim()) {
      if (onStatus) onStatus("processing", `回写段落 (${i + 1}/${texts.length})...`, true);
      await replaceSingleMarkedContent(result.trim());
    } else {
      // 移除由于空返回而没有被 replaceSingleMarkedContent 清理掉的头部段落标签
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
  } catch {
    // 忽略清理错误
  }
}
