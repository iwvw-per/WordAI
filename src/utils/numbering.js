/**
 * WordAI 图表编号同步工具类
 */

/**
 * 重新编排全文图表编号
 * 修复：支持上下文复用
 */
export async function renumberFiguresAndTables(passedContext = null) {
    const doWork = async (context) => {
        const body = context.document.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        for (const p of paragraphs.items) p.load("text");
        await context.sync();

        let figCount = 0;
        let tabCount = 0;

        for (const p of paragraphs.items) {
            const text = p.text.trim();
            if (text.length > 0 && text.length < 150) {
                if (/^图\s*[0-9]+/.test(text)) {
                    figCount++;
                    const search = p.search("图[ ][0-9]@", { matchWildcards: true });
                    search.load("items");
                    await context.sync();
                    if (search.items.length > 0) search.items[0].insertText(`图 ${figCount}`, "Replace");
                } else if (/^表\s*[0-9]+/.test(text)) {
                    tabCount++;
                    const search = p.search("表[ ][0-9]@", { matchWildcards: true });
                    search.load("items");
                    await context.sync();
                    if (search.items.length > 0) search.items[0].insertText(`表 ${tabCount}`, "Replace");
                }
            }
        }
        await context.sync();
        return { figures: figCount, tables: tabCount, message: `重排完成：共发现 ${figCount} 图和 ${tabCount} 表。` };
    };

    if (passedContext) return await doWork(passedContext);
    return await Word.run(async (context) => await doWork(context));
}
