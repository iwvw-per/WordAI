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
 * @returns {Promise<Array<{id: number, text: string}>>}
 */
export async function parseBibliography() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        // 查找常见的参考文献标题
        const searchResults = body.search("参考文献", { matchCase: false });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length === 0) return [];

        // 获取标题后的所有内容
        const lastTitle = searchResults.items[searchResults.items.length - 1];
        const bibRange = lastTitle.getNextTextRange("End");
        bibRange.load("text");
        await context.sync();

        // 简单的按行拆分逻辑 (后续可增强为语义拆分)
        const lines = bibRange.text.split('\n').filter(line => line.trim().length > 5);
        return lines.map((line, index) => ({
            id: index + 1,
            text: line.trim()
        }));
    });
}

/**
 * 为指定范围创建到书签的超链接
 * @param {Word.Range} range 
 * @param {string} bookmarkName 
 */
export async function createReferenceLink(range, bookmarkName) {
    await Word.run(async (context) => {
        range.hyperlink = "#" + bookmarkName;
        await context.sync();
    });
}
