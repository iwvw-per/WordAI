import * as llm from "./llm.js";

/**
 * 扫描文本并提取术语及冲突
 * @param {string} text 文档文本 
 */
export async function extractTerminology(text) {
    const prompt = `请从以下学术文本中提取核心术语。如果发现同一个概念使用了不同的表达方式（如“卷积神经网络”与“卷积感知机”、“CNN”等），请指出潜在的冲突。
请严格返回 JSON 格式数组，格式如下：
[
  { "standard": "推荐的标准术语", "aliases": ["发现的别名1", "发现的别名2"] }
]
如果没有冲突，返回空数组 []。仅返回 JSON 字符串，不要有任何其他解释。`;

    const result = await llm.callLLM(prompt, text);
    try {
        const jsonStr = result.replace(/```json\n?|\n?```/g, "").trim();
        return JSON.parse(jsonStr);
    } catch (e) {
        throw new Error("术语提取失败，LLM 返回格式错误");
    }
}

/**
 * 将别名替换为标准术语
 */
export async function replaceTerminology(aliases, standardTerm) {
    await Word.run(async (context) => {
        const body = context.document.body;
        for (const alias of aliases) {
            const searchResults = body.search(alias, { matchCase: false });
            searchResults.load("items");
            await context.sync();
            for (let i = 0; i < searchResults.items.length; i++) {
                // 仅替换纯文本，尽量不破坏格式
                searchResults.items[i].insertText(standardTerm, "Replace");
            }
        }
        await context.sync();
    });
}
