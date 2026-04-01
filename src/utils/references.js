/**
 * WordAI 参考文献处理工具类
 */

/**
 * 全文扫描占位符
 */
export async function scanPlaceholders() {
    return await Word.run(async (context) => {
        const results = context.document.body.search("【*】", { matchWildcards: true });
        results.load("items");
        await context.sync();
        for (const item of results.items) item.load("text");
        await context.sync();
        return {
            count: results.items.length,
            texts: results.items.map((item, index) => ({ text: item.text, index: index }))
        };
    });
}

/**
 * 解析文末参考文献列表
 */
export async function parseBibliography() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        const searchChinese = body.search("参考文献", { matchWildcards: false, matchCase: false });
        const searchEnglish = body.search("References", { matchWildcards: false, matchCase: false });
        searchChinese.load("items");
        searchEnglish.load("items");
        await context.sync();

        const allTitleResults = [...searchChinese.items, ...searchEnglish.items];
        const styledParagraphs = body.paragraphs;
        styledParagraphs.load(["items"]);
        await context.sync();
        for (const p of styledParagraphs.items) p.load(["style", "text"]);
        await context.sync();

        let bibRange = null;
        if (allTitleResults.length > 0) {
            bibRange = allTitleResults[allTitleResults.length - 1].expandTo(body.getRange("End"));
        } else {
            const bibStartPar = styledParagraphs.items.find(p =>
                (p.style && (p.style.toLowerCase().includes("bib") || p.style.includes("参考文献"))) ||
                /^\[1\]|1\./.test(p.text.trim())
            );
            if (bibStartPar) bibRange = bibStartPar.expandTo(body.getRange("End"));
        }

        if (!bibRange) return [];
        bibRange.load("text");
        await context.sync();

        const lines = bibRange.text.split(/[\r\n]+/).map(l => l.trim()).filter(l => l.length > 8);
        return lines.map((line, index) => {
            const yearMatch = line.match(/(19|20)\d{2}/);
            return { id: index + 1, text: line, year: yearMatch ? yearMatch[0] : null };
        });
    });
}

/**
 * 匹配占位符与参考文献
 */
export function matchPlaceholderToBibliography(placeholderText, bibliography) {
    const clean = placeholderText.replace(/[【】\[\]]/g, "");
    const parts = clean.split(/[,，\s]+/).map(p => p.trim());
    const author = parts[0];
    const year = parts.find(p => /^\d{4}$/.test(p));
    const scored = bibliography.map(entry => {
        let score = 0;
        if (author && entry.text.includes(author)) score += 50;
        if (year && entry.year === year) score += 50;
        return { ...entry, score };
    }).filter(e => e.score > 0);
    return scored.sort((a, b) => b.score - a.score);
}

/**
 * 创建并链接参考文献
 * 修复：支持上下文复用
 */
export async function createReferenceLink(range, bookmarkName, passedContext = null) {
    const doWork = async (context) => {
        range.hyperlink = `#${bookmarkName}`;
        range.font.color = "#2563eb";
        range.font.underline = Word.UnderlineType.none;
        await context.sync();
    };

    if (range.context) await doWork(range.context);
    else if (passedContext) await doWork(passedContext);
    else await Word.run(async (context) => await doWork(context));
}
