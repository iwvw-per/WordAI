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
 * 算法：扫描全文匹配“图 X”或“表 X”，按出现物理顺序重新分配连续编号。
 */
export async function renumberFiguresAndTables() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        
        // 1. 扫描图、表标题 (支持 "图 1" 和 "图1" 等变体)
        // 使用通配符：图字 + 可选空格 + 一个或多个数字
        const figResults = body.search("图[ ]@[0-9]@", { matchWildcards: true });
        const tabResults = body.search("表[ ]@[0-9]@", { matchWildcards: true });
        // 如果通配符匹配不全，再尝试直接搜索（Word Wildcards 限制较多）
        const figResultsBasic = body.search("图[0-9]@", { matchWildcards: true });
        const tabResultsBasic = body.search("表[0-9]@", { matchWildcards: true });
        
        figResults.load("items");
        tabResults.load("items");
        await context.sync();
        
        // 处理图片编号
        for (let i = 0; i < figResults.items.length; i++) {
            const range = figResults.items[i];
            const newNum = i + 1;
            range.insertText(`图 ${newNum}`, "Replace");
        }

        // 处理表格编号
        for (let i = 0; i < tabResults.items.length; i++) {
            const range = tabResults.items[i];
            const newNum = i + 1;
            range.insertText(`表 ${newNum}`, "Replace");
        }

        await context.sync();
        
        return {
            figures: figResults.items.length,
            tables: tabResults.items.length,
            message: `已自动重排 ${figResults.items.length} 个图片标题和 ${tabResults.items.length} 个表格标题。`
        };
    });
}
