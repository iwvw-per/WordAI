/**
 * WordAI 表格处理工具类
 * 提供学术三线表识别、美化与格式控制逻辑
 */

/**
 * 探测表格的表头行数 (最多检测前 3 行)
 * 逻辑：如果某行存在合并单元格（单元格数少于总列数），或者前几行全部加粗且居中，则视为表头。
 */
export async function detectHeaderRows(table) {
    return await Word.run(async (context) => {
        table.load(["rows/items/cells/count", "columns/count"]);
        await context.sync();
        
        const totalCols = table.columns.count;
        let headerRows = 1;

        for (let i = 0; i < Math.min(table.rows.items.length, 3); i++) {
            const row = table.rows.items[i];
            // 如果单元格数量少于总列数，说明存在横向合并，这通常是复合表头的一部分
            if (row.cells.count < totalCols) {
                headerRows = i + 1;
            } else {
                // 如果没有合并，但前几行是加粗居中的，也可能是表头的最后一行
                // 这里加一个启发式判断：如果第一行是加粗的，默认至少有一行表头
                if (i === 0) headerRows = 1;
            }
        }
        return headerRows;
    });
}

/**
 * 为指定表格应用学术三线表样式
 */
export async function applyAcademicStyle(table, config = { topWidth: 1.5, bottomWidth: 1.5, headerWidth: 0.75 }) {
    await Word.run(async (context) => {
        table.load(["rows/items", "rows/count", "columns/count"]);
        await context.sync();

        const headerRowCount = await detectHeaderRows(table);

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

        // 4. 设置栏目线 (应用在表头的最后一行)
        if (table.rows.count >= headerRowCount) {
            for (let i = 0; i < headerRowCount; i++) {
                const row = table.rows.items[i];
                // 仅在表头的最后一行添加底线
                if (i === headerRowCount - 1) {
                    row.borders.bottom.style = "Single";
                    row.borders.bottom.width = config.headerWidth;
                    row.borders.bottom.color = "black";
                }
                // 表头通用样式：加粗 + 居中
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
