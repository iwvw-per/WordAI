/**
 * 将简单的 Markdown 文本转换为 Word 富文本并插入指定位置
 * 支持：### 标题, **加粗**, 1. 列表, {{REF_N}} 占位符以及换行
 * 
 * 修复：不再在内部创建新的 Word.run，而是接收 context 参数
 * 避免跨 Word.run 传递代理对象的问题
 */

/**
 * 核心处理逻辑：将单行 Markdown 注入到 Range 或 ContentControl
 * @param {Word.Range | Word.ContentControl} container 目标容器
 * @param {string} line 文本行
 * @param {Array} refMap 占位符映射 (可选)
 */
export async function processMarkdownLine(container, line, refMap = []) {
  // 1. 处理标题 (### Title)
  if (line.startsWith("###")) {
    const titleText = line.replace(/^###\s*/, "");
    const insertedRange = container.insertParagraph(titleText, "End");
    insertedRange.font.bold = true;
    insertedRange.font.size = 14;
    insertedRange.spacingBefore = 12;
    return;
  }

  if (line.startsWith("##")) {
    const titleText = line.replace(/^##\s*/, "");
    const insertedRange = container.insertParagraph(titleText, "End");
    insertedRange.font.bold = true;
    insertedRange.font.size = 16;
    insertedRange.spacingBefore = 14;
    return;
  }

  // 2. 基础段落处理：支持加粗和占位符
  const insertedRange = container.insertParagraph("", "End");

  // 采用复合正则切分：同时匹配 **加粗** 和 {{REF_N}}
  const parts = line.split(/(\*\*.*?\*\*|{{REF_\d+}})/g);

  for (const part of parts) {
    if (!part) continue;

    if (part.startsWith("**") && part.endsWith("**")) {
      const boldText = part.substring(2, part.length - 2);
      const run = insertedRange.insertText(boldText, "End");
      run.font.bold = true;
    } else if (part.startsWith("{{REF_") && part.endsWith("}}")) {
      const refItem = refMap.find(m => m.placeholder === part);
      if (refItem) {
        insertedRange.insertOoxml(refItem.ooxml, "End");
      } else {
        insertedRange.insertText(part, "End");
      }
    } else {
      insertedRange.insertText(part, "End");
    }
  }
}

/**
 * 批量插入 Markdown 文本
 * 修复：接收 context 参数而非在内部创建新的 Word.run
 * 如果未传入 context，则创建新的 Word.run（向后兼容）
 */
export async function insertMarkdownAsRichText(target, markdown, location = "End", refMap = []) {
  const doWork = async (context) => {
    const cleanMarkdown = markdown.replace(/\\n/g, "\n");
    const lines = cleanMarkdown.split(/\r?\n/);

    if (location === "Replace") {
      target.insertText("", "Replace");
    }

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) {
        if (i !== lines.length - 1) {
          target.insertParagraph("", "End");
        }
        continue;
      }
      await processMarkdownLine(target, line, refMap);
    }
    await context.sync();
  };

  // 如果 target 有关联的 context，直接使用它（同一个 Word.run 内）
  if (target.context) {
    await doWork(target.context);
  } else {
    // 降级：创建新的 Word.run（不推荐，但向后兼容）
    await Word.run(async (context) => {
      await doWork(context);
    });
  }
}
