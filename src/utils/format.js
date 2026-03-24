/**
 * 将简单的 Markdown 文本转换为 Word 富文本并插入指定位置
 * 支持：### 标题, **加粗**, 1. 列表, 以及换行
 * @param {Word.Range | Word.Body} target - 插入目标
 * @param {string} markdown - 待转换的 Markdown 文本
 * @param {string} location - 插入位置 ("Start", "End", "Replace")
 */
export async function insertMarkdownAsRichText(target, markdown, location = "End") {
  await Word.run(async (context) => {
    // 处理 literal \n 字符串
    const cleanMarkdown = markdown.replace(/\\n/g, "\n");
    const lines = cleanMarkdown.split("\n");
    
    // 如果是替换模式，先清空内容
    if (location === "Replace") {
      target.insertText("", "Replace");
    }

    let currentContainer = target;

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (!line) {
        // 插入空行
        currentContainer.insertParagraph("", "End");
        continue;
      }

      let insertedRange;
      
      // 1. 处理标题 (### Title)
      if (line.startsWith("###")) {
        const titleText = line.replace(/^###\s*/, "");
        insertedRange = currentContainer.insertParagraph(titleText, "End");
        insertedRange.font.bold = true;
        insertedRange.font.size = 14;
        insertedRange.spacingBefore = 12;
      } 
      else if (line.startsWith("##")) {
        const titleText = line.replace(/^##\s*/, "");
        insertedRange = currentContainer.insertParagraph(titleText, "End");
        insertedRange.font.bold = true;
        insertedRange.font.size = 16;
        insertedRange.spacingBefore = 14;
      }
      // 2. 处理加粗 (**Text**) - 简化版：仅处理全行加粗或通过正则二次处理
      else {
        // 基础段落插入
        insertedRange = currentContainer.insertParagraph("", "End");
        
        // 分解行内格式
        // 匹配 **...**
        const parts = line.split(/(\*\*.*?\*\*)/g);
        for (const part of parts) {
          if (part.startsWith("**") && part.endsWith("**")) {
            const boldText = part.substring(2, part.length - 2);
            const run = insertedRange.insertText(boldText, "End");
            run.font.bold = true;
          } else {
            insertedRange.insertText(part, "End");
          }
        }
      }
      
      await context.sync();
    }
  });
}
