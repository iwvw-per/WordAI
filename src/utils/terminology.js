import * as llm from "./llm.js";

/**
 * 扫描文本并提取术语及冲突
 */
export async function extractTerminology(text) {
    const prompt = `请从以下学术文本中提取核心术语冲突。返回 JSON 数组格式: [{"standard": "标准术语", "aliases": ["别名1"]}]。无冲突返回 []。`;
    const result = await llm.callLLM(prompt, text);
    try {
        let jsonStr = result.trim().replace(/^```json\n?|\n?```$/gi, "");
        const arrayMatch = jsonStr.match(/[\s\S]*\]/);
        if (arrayMatch) jsonStr = arrayMatch[0];
        return JSON.parse(jsonStr);
    } catch (e) {
        return [];
    }
}

/**
 * 将别名替换为标准术语
 * 保护公式内容不被替换（通过检查匹配项的 OOXML 是否包含 oMath 标签）
 */
export async function replaceTerminology(aliases, standardTerm, passedContext = null) {
    const doWork = async (context) => {
        const body = context.document.body;
        for (const alias of aliases) {
            const searchResults = body.search(alias, { matchCase: false });
            searchResults.load("items");
            await context.sync();

            // 逐一检查，跳过公式内的匹配项
            for (let i = 0; i < searchResults.items.length; i++) {
                const item = searchResults.items[i];
                // 检查所在段落是否包含公式
                const parentParagraph = item.paragraphs.getFirst();
                const paraOoxml = parentParagraph.getOoxml();
                await context.sync();

                // 如果段落包含 oMath 公式，则需要精确判断匹配项自身是否在公式内
                if (paraOoxml.value && (paraOoxml.value.includes("<m:oMath") || paraOoxml.value.includes("<m:oMathPara"))) {
                    // 获取匹配项自身的 OOXML
                    const itemOoxml = item.getOoxml();
                    await context.sync();
                    // 如果匹配项自身的 OOXML 包含 oMath，说明它在公式内，跳过
                    if (itemOoxml.value && (itemOoxml.value.includes("<m:oMath") || itemOoxml.value.includes("<m:oMathPara") || itemOoxml.value.includes("<m:r"))) {
                        continue;
                    }
                }

                item.insertText(standardTerm, "Replace");
            }
        }
        await context.sync();
    };

    if (passedContext) await doWork(passedContext);
    else await Word.run(async (context) => await doWork(context));
}
