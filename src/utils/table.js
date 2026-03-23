/**
 * WordAI 表格处理工具类
 * 提供学术三线表识别、美化与格式控制逻辑
 */

/**
 * 为指定表格应用学术三线表样式
 * @param {Word.Table} table Word 表格对象
 * @param {Object} config 配置参数
 * @param {number} config.topWidth 顶线宽度 (pt)
 * @param {number} config.bottomWidth 底线宽度 (pt)
 * @param {number} config.headerWidth 栏目线宽度 (pt)
 */
export async function applyAcademicStyle(table, config = { topWidth: 1.5, bottomWidth: 1.5, headerWidth: 0.75 }) {
    await Word.run(async (context) => {
        table.load(["rows/items", "rows/count"]);
        await context.sync();

        // 1. 清除所有原有边框
        table.borders.outsideLineWidth = 0;
        table.borders.insideHorizontalLineWidth = 0;
        table.borders.insideVerticalLineWidth = 0;

        // 2. 设置顶线
        table.borders.top.style = "Single";
        table.borders.top.width = config.topWidth;
        table.borders.top.color = "black";

        // 3. 设置底线
        table.borders.bottom.style = "Single";
        table.borders.bottom.width = config.bottomWidth;
        table.borders.bottom.color = "black";

        // 4. 智能识别表头并设置栏目线
        // 目前先处理第一行，后续任务将增强为多级表头识别
        if (table.rows.count > 0) {
            const headerRow = table.rows.getFirst();
            headerRow.borders.bottom.style = "Single";
            headerRow.borders.bottom.width = config.headerWidth;
            headerRow.borders.bottom.color = "black";
            
            // 默认表头居中加粗
            headerRow.font.bold = true;
            headerRow.horizontalAlignment = "Center";
        }

        await context.sync();
    });
}

/**
 * 全文扫描所有表格并返回基本信息
 */
export async function getAllTablesInfo() {
    return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();

        return tables.items.map((table, index) => {
            // 这里可以添加更多识别信息，如表格前后的文字作为标题
            return {
                id: index,
                rowCount: table.rowCount,
                columnCount: table.columnCount,
                handle: table // 返回引用供后续操作
            };
        });
    });
}
