/**
 * WordAI 图表编号同步工具类
 */

/**
 * 重新编排全文图表编号
 * 支持：图 N、表 N、Figure N、Table N、图 N-N、表 N.N 等格式
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

        // 匹配中文和英文的图表编号（支持章节编号如 图 2-1、Figure 3.2）
        const figRegex = /^(图|Figure|Fig\.?)\s*[0-9]+([.\-][0-9]+)*/i;
        const tabRegex = /^(表|Table)\s*[0-9]+([.\-][0-9]+)*/i;

        const tasks = [];
        for (const p of paragraphs.items) {
            const text = p.text.trim();
            if (text.length === 0 || text.length > 200) continue;

            const figMatch = text.match(figRegex);
            if (figMatch) {
                figCount++;
                const search = p.search(figMatch[0], { matchWildcards: false, matchCase: false });
                search.load("items");
                tasks.push({ search, prefix: figMatch[1], count: figCount });
            } else {
                const tabMatch = text.match(tabRegex);
                if (tabMatch) {
                    tabCount++;
                    const search = p.search(tabMatch[0], { matchWildcards: false, matchCase: false });
                    search.load("items");
                    tasks.push({ search, prefix: tabMatch[1], count: tabCount });
                }
            }
        }
        await context.sync(); // ⚡ 一键批量获取所有图表搜索代理对象，将 IPC 次数由 O(N) 降至常数 1 次

        for (const task of tasks) {
            if (task.search.items.length > 0) {
                task.search.items[0].insertText(`${task.prefix} ${task.count}`, "Replace");
            }
        }
        await context.sync();
        return { figures: figCount, tables: tabCount, message: `重排完成：共发现 ${figCount} 图和 ${tabCount} 表。` };
    };

    if (passedContext) return await doWork(passedContext);
    return await Word.run(async (context) => await doWork(context));
}
