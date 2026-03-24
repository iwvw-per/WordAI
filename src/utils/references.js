/**
 * WordAI 参考文献处理工具类
 * 提供占位符扫描、参考文献解析、编号更新与跳转逻辑
 */

/**
 * 全文扫描占位符 (例如：【作者, 年份】)
 * @returns {Promise<Word.RangeCollection>}
 */
export async function scanPlaceholders() {
    return await Word.run(async (context) => {
        // 使用通配符搜索 【*】 格式的内容
        const results = context.document.body.search("【*】", { matchWildcards: true });
        results.load("items");
        await context.sync();
        return results;
    });
}

/**
 * 尝试解析文末的参考文献列表
 */
export async function parseBibliography() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        // 查找常见的参考文献标题
        const searchResults = body.search("参考文献|References", { matchWildcards: true, matchCase: false });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length === 0) return [];

        // 获取标题后的范围
        const lastTitle = searchResults.items[searchResults.items.length - 1];
        const bibRange = lastTitle.getNextTextRange("End");
        bibRange.load("text");
        await context.sync();

        // 解析条目：按换行拆分，过滤掉过短的行
        const lines = bibRange.text.split(/[\r\n]+/).map(l => l.trim()).filter(l => l.length > 10);
        return lines.map((line, index) => {
            // 简单提取年份用于匹配：识别 19xx 或 20xx
            const yearMatch = line.match(/(19|20)\d{2}/);
            return {
                id: index + 1,
                text: line,
                year: yearMatch ? yearMatch[0] : null
            };
        });
    });
}

/**
 * 将占位符文本与参考文献列表进行匹配
 * @param {string} placeholderText 如 "【张三, 2023】"
 * @param {Array} bibliography 解析出的列表
 */
export function matchPlaceholderToBibliography(placeholderText, bibliography) {
    const clean = placeholderText.replace(/[【】\[\]]/g, "");
    const parts = clean.split(/[,，\s]+/).map(p => p.trim());
    
    const author = parts[0];
    const year = parts.find(p => /^\d{4}$/.test(p));

    // 评分匹配
    const scored = bibliography.map(entry => {
        let score = 0;
        if (author && entry.text.includes(author)) score += 50;
        if (year && entry.year === year) score += 50;
        // 模糊匹配全文
        if (author && !entry.text.includes(author)) {
             // 检查拼音或部分匹配（简单处理）
        }
        return { ...entry, score };
    }).filter(e => e.score > 0);

    return scored.sort((a, b) => b.score - a.score);
}

/**
 * 为指定范围创建书签并链接
 */
export async function createReferenceLink(range, bookmarkName) {
    await Word.run(async (context) => {
        // 这是一个示意：Word JS API 暂时不支持直接创建 Hyperlink 到 Bookmark 的便捷方法
        // 通常需要通过 insertOoxml 或操作 Field Code。
        // 此处简化为直接替换文本，后续可增强。
        range.font.color = "#2563eb";
        range.font.underline = "Single";
        await context.sync();
    });
}
