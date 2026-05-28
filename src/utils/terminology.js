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

            const tasks = [];
            for (let i = 0; i < searchResults.items.length; i++) {
                const item = searchResults.items[i];
                const parentParagraph = item.paragraphs.getFirst();
                const paraOoxml = parentParagraph.getOoxml();
                tasks.push({ item, parentParagraph, paraOoxml });
            }
            await context.sync(); // ⚡ 集中批量 sync，仅需 1 次往返即可获取所有段落的 OOXML

            const needItemOoxmlTasks = [];
            for (const task of tasks) {
                if (task.paraOoxml.value && (task.paraOoxml.value.includes("<m:oMath") || task.paraOoxml.value.includes("<m:oMathPara"))) {
                    const itemOoxml = task.item.getOoxml();
                    needItemOoxmlTasks.push({ task, itemOoxml });
                }
            }
            if (needItemOoxmlTasks.length > 0) {
                await context.sync(); // ⚡ 集中批量 sync，仅当公式检测段落内有匹配项时才进行二次提取
            }

            for (const task of tasks) {
                const matchedNeed = needItemOoxmlTasks.find(n => n.task === task);
                if (matchedNeed) {
                    const val = matchedNeed.itemOoxml.value;
                    if (val && (val.includes("<m:oMath") || val.includes("<m:oMathPara") || val.includes("<m:r"))) {
                        continue;
                    }
                }
                task.item.insertText(standardTerm, "Replace");
            }
        }
        await context.sync();
    };

    if (passedContext) await doWork(passedContext);
    else await Word.run(async (context) => await doWork(context));
}
