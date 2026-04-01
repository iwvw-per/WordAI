import * as llm from "./llm.js";

/**
 * 扫描文本并提取术语及冲突
 */
export async function extractTerminology(text) {
    const prompt = `请从以下学术文本中提取核心术语冲突。返回 JSON 数组格式: [{"standard": "标准术语", "aliases": ["别名1"]}]。无冲突返回 []。`;
    const result = await llm.callLLM(prompt, text);
    try {
        let jsonStr = result.trim().replace(/^```json\n?|\n?```$/gi, "");
        const arrayMatch = jsonStr.match(/\[[\s\S]*\]/);
        if (arrayMatch) jsonStr = arrayMatch[0];
        return JSON.parse(jsonStr);
    } catch (e) {
        return [];
    }
}

/**
 * 将别名替换为标准术语
 * 修复：支持上下文复用
 */
export async function replaceTerminology(aliases, standardTerm, passedContext = null) {
    const doWork = async (context) => {
        const body = context.document.body;
        for (const alias of aliases) {
            const searchResults = body.search(alias, { matchCase: false });
            searchResults.load("items");
            await context.sync();
            for (let i = 0; i < searchResults.items.length; i++) {
                searchResults.items[i].insertText(standardTerm, "Replace");
            }
        }
        await context.sync();
    };

    if (passedContext) await doWork(passedContext);
    else await Word.run(async (context) => await doWork(context));
}
