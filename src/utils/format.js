/**
 * format.js - Word 格式操作工具类
 */

/**
 * 在选区中搜索并高亮关键字
 * @param {Array<string>} keywords - 关键字列表
 * @param {string} color - 高亮颜色（Hex）
 */
export async function highlightKeywords(keywords, color = "#2563eb") {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    
    for (const keyword of keywords) {
      if (!keyword || keyword.length < 2) continue; // 忽略过短或空关键字
      
      const searchResults = selection.search(keyword, { 
        matchCase: false,
        matchWildcards: false
      });
      
      searchResults.load("items");
      await context.sync();
      
      searchResults.items.forEach((range) => {
        range.font.bold = true;
        range.font.color = color;
      });
    }
    
    await context.sync();
  });
}
