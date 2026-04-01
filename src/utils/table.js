/**
 * WordAI 表格处理工具类
 */

/**
 * 探测表格的表头行数
 */
async function detectHeaderRowsInternal(table, context) {
    table.load(["rowCount", "columnCount"]);
    await context.sync();

    const totalCols = table.columnCount;
    let headerRows = 1;

    const rowsToCheck = Math.min(table.rowCount, 3);
    for (let i = 0; i < rowsToCheck; i++) {
        const row = table.rows.items[i];
        row.load("cellCount");
    }
    await context.sync();

    for (let i = 0; i < rowsToCheck; i++) {
        const row = table.rows.items[i];
        if (row.cellCount < totalCols) {
            headerRows = i + 1;
        }
    }
    return headerRows;
}

/**
 * 为指定表格应用学术三线表样式
 * 修复：支持上下文复用
 */
export async function applyAcademicStyle(table, config = { topWidth: 1.5, bottomWidth: 1.5, headerWidth: 0.75 }) {
    const doWork = async (context) => {
        table.load(["rowCount", "columnCount"]);
        const rows = table.rows;
        rows.load("items");
        await context.sync();

        const headerRowCount = await detectHeaderRowsInternal(table, context);

        const borderLocations = [
            Word.BorderLocation.top, Word.BorderLocation.bottom, Word.BorderLocation.left,
            Word.BorderLocation.right, Word.BorderLocation.insideHorizontal, Word.BorderLocation.insideVertical,
        ];
        for (const loc of borderLocations) {
            table.getBorder(loc).type = Word.BorderType.none;
        }
        await context.sync();

        table.getBorder(Word.BorderLocation.top).set({ type: Word.BorderType.single, width: config.topWidth, color: "#000000" });
        table.getBorder(Word.BorderLocation.bottom).set({ type: Word.BorderType.single, width: config.bottomWidth, color: "#000000" });
        await context.sync();

        if (table.rowCount >= headerRowCount) {
            for (let i = 0; i < headerRowCount; i++) {
                const row = rows.items[i];
                if (i === headerRowCount - 1) {
                    row.getBorder(Word.BorderLocation.bottom).set({ type: Word.BorderType.single, width: config.headerWidth, color: "#000000" });
                }
                row.font.bold = true;
                row.horizontalAlignment = Word.Alignment.centered;
            }
        }
        await context.sync();
    };

    if (table.context) {
        await doWork(table.context);
    } else {
        await Word.run(async (context) => await doWork(context));
    }
}

/**
 * 全文扫描所有表格
 */
export async function getAllTablesInfo() {
    return await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();
        for (const table of tables.items) table.load(["rowCount", "columnCount"]);
        await context.sync();
        return tables.items.map((table, index) => ({
            id: index,
            rowCount: table.rowCount,
            columnCount: table.columnCount
        }));
    });
}
