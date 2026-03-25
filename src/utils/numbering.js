/**
 * WordAI 图表编号同步工具类
 */

/**
 * 重新编排全文图表编号并返回统计信息
 * 这是一个基础实现，基于简单的文本搜索。
 * 高级实现需操作 SEQ 域，但考虑到 API 限制，此处提供启发式文本替换示例。
 */
/**
 * 重新编排全文图表编号
 * 算法：扫描全文匹配“图 X”或“表 X”，优先处理作为独立行的标题。
 */
export async function renumberFiguresAndTables() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load(["items", "text"]);
        await context.sync();
        
        let figCount = 0;
        let tabCount = 0;

        for (const p of paragraphs.items) {
            const text = p.text.trim();
            // 启发式逻辑：标题通常较短且以“图”或“表”开头
            if (text.length > 0 && text.length < 150) {
                // 处理图片标题
                if (/^图\s*[0-9]+/.test(text)) {
                    figCount++;
                    const search = p.search("图\s*[0-9]+", { matchWildcards: true });
                    search.load("items");
                    await context.sync();
                    if (search.items.length > 0) {
                        search.items[0].insertText(`图 ${figCount}`, "Replace");
                    }
                }
                // 处理表格标题
                else if (/^表\s*[0-9]+/.test(text)) {
                    tabCount++;
                    const search = p.search("表\s*[0-9]+", { matchWildcards: true });
                    search.load("items");
                    await context.sync();
                    if (search.items.length > 0) {
                        search.items[0].insertText(`表 ${tabCount}`, "Replace");
                    }
                }
            }
        }

        await context.sync();
        return {
            figures: figCount,
            tables: tabCount,
            message: `重排完成：共发现 ${figCount} 个图片标题和 ${tabCount} 个表格标题。`
        };
    });
}
