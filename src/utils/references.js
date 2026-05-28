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
        let bibRange = null;
        if (allTitleResults.length > 0) {
            bibRange = allTitleResults[allTitleResults.length - 1].expandTo(body.getRange("End"));
        } else {
            // ⚡ 降级策略优先：使用 search 快速锚定文献库首项 [1] 或 1.
            const searchBracket1 = body.search("\\[1\\]", { matchWildcards: true });
            const searchDot1 = body.search("1. ", { matchWildcards: false });
            searchBracket1.load("items");
            searchDot1.load("items");
            await context.sync();

            let bibStartPar = null;
            if (searchBracket1.items.length > 0) {
                bibStartPar = searchBracket1.items[0];
            } else if (searchDot1.items.length > 0) {
                bibStartPar = searchDot1.items[0];
            }

            // 万不得已作为最后后备：只扫描文档最后 200 个段落的样式，完全避免载入全身大 payload 崩溃
            if (!bibStartPar) {
                const styledParagraphs = body.paragraphs;
                styledParagraphs.load("items");
                await context.sync();

                const count = styledParagraphs.items.length;
                const checkStartIndex = Math.max(0, count - 200);
                const subItems = styledParagraphs.items.slice(checkStartIndex);

                for (const p of subItems) p.load(["style", "text"]);
                await context.sync();

                const found = subItems.find(p =>
                    (p.style && (p.style.toLowerCase().includes("bib") || p.style.includes("参考文献"))) ||
                    /^\[1\]|1\./.test(p.text.trim())
                );
                if (found) bibStartPar = found;
            }

            if (bibStartPar) {
                bibRange = bibStartPar.expandTo(body.getRange("End"));
            }
        }

        if (!bibRange) {
            console.warn("WordAI parseBibliography: bibRange is null, did not find references section start!");
            return [];
        }
        bibRange.load("text");
        await context.sync();

        console.log("WordAI parseBibliography Range Text length:", bibRange.text.length);
        console.log("WordAI parseBibliography Range Text:", bibRange.text);
        const lines = bibRange.text.split(/[\r\n]+/).map(l => l.trim()).filter(l => l.length > 8);
        console.log("WordAI parseBibliography Lines:", lines);
        const results = lines.map((line, index) => {
            const yearMatch = line.match(/(19|20)\d{2}/);
            const cleanText = line.replace(/^(\s*(?:[\[【]\s*\d+\s*[\]】]|\d+\s*[\.．、])\s*)/, "");
            const authorMatch = cleanText.match(/[\u4e00-\u9fff]{2,4}|[a-zA-Z\-]{2,}/);
            return {
                id: index + 1,
                text: line,
                year: yearMatch ? yearMatch[0] : null,
                coreAuthor: authorMatch ? authorMatch[0].toLowerCase() : null
            };
        });
        console.log("WordAI parseBibliography Parsed Results:", results);
        return results;
    });
}

/**
 * 匹配占位符与参考文献
 */
export function matchPlaceholderToBibliography(placeholderText, bibliography) {
    const clean = placeholderText.replace(/[【】\[\]]/g, "");
    
    // ⚡ 精准模糊比对：提取占位符内的核心作者（2~4个中文或2个字符以上的英文）与年份
    const authorMatch = clean.match(/[\u4e00-\u9fff]{2,4}|[a-zA-Z\-]{2,}/);
    const yearMatch = clean.match(/\b(19|20)\d{2}\b/);
    
    const pAuthor = authorMatch ? authorMatch[0].toLowerCase() : null;
    const pYear = yearMatch ? yearMatch[0] : null;

    console.log(`WordAI matchPlaceholderToBibliography: placeholder="${placeholderText}" clean="${clean}" pAuthor="${pAuthor}" pYear="${pYear}"`);
    console.log("WordAI matchPlaceholderToBibliography bibliography list:", bibliography);

    const scored = bibliography.map(entry => {
        let score = 0;
        
        // 1. 如果核心作者（前2-4字）精确匹配，给极高分
        if (pAuthor && entry.coreAuthor === pAuthor) {
            score += 60;
        } else if (pAuthor && entry.text.toLowerCase().includes(pAuthor)) {
            // 2. 降级：如果整行文献包含核心作者
            score += 30;
        }
        
        // 3. 年份相同，给高分
        if (pYear && entry.year === pYear) {
            score += 40;
        }
        
        console.log(`  --> Entry ID=${entry.id} CoreAuthor="${entry.coreAuthor}" Year="${entry.year}" -> Score=${score}`);
        return { ...entry, score };
    }).filter(e => e.score > 0);
    
    const sorted = scored.sort((a, b) => b.score - a.score);
    console.log("WordAI matchPlaceholderToBibliography sorted matched results:", sorted);
    return sorted;
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
