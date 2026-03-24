/**
 * WordAI 表格处理工具类
 * 提供学术三线表识别、美化与格式控制逻辑
 */

/**
 * 探测表格的表头行数 (最多检测前 3 行)
 * 逻辑：如果某行存在合并单元格（单元格数少于总列数），视为表头。
 * @param {Word.Table} table
 * @param {Word.RequestContext} context
 */
async function detectHeaderRowsInternal(table, context) {
    table.load(["rows/items/cells/count", "columnCount"]);
    await context.sync();
    
    const totalCols = table.columnCount;
    let headerRows = 1;

    for (let i = 0; i < Math.min(table.rows.items.length, 3); i++) {
        const row = table.rows.items[i];
        if (row.cells.count < totalCols) {
            headerRows = i + 1;
        }
    }
    return headerRows;
}

/**
 * 为指定表格应用学术三线表样式
 */
export async function applyAcademicStyle(table, config = { topWidth: 1.5, bottomWidth: 1.5, headerWidth: 0.75 }) {
    await Word.run(async (context) => {
        const headerRowCount = await detectHeaderRowsInternal(table, context);

        // 1. 清除并设置顶底线
        table.borders.outsideLineWidth = 0;
        table.borders.insideHorizontalLineWidth = 0;
        table.borders.insideVerticalLineWidth = 0;

        table.borders.top.style = "Single";
        table.borders.top.width = config.topWidth;
        table.borders.bottom.style = "Single";
        table.borders.bottom.width = config.bottomWidth;

        // 2. 应用栏目线
        if (table.rows.count >= headerRowCount) {
            for (let i = 0; i < headerRowCount; i++) {
                const row = table.rows.items[i];
                if (i === headerRowCount - 1) {
                    row.borders.bottom.style = "Single";
                    row.borders.bottom.width = config.headerWidth;
                }
                row.font.bold = true;
                row.horizontalAlignment = "Center";
            }
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
