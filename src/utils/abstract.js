import * as llm from "./llm.js";

/**
 * 生成摘要与关键词
 * 修复：支持上下文复用
 */
export async function generateAbstract(passedContext = null) {
    const doWork = async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        let contextText = body.text;
        if (!contextText || contextText.trim().length === 0) throw new Error("文档内容为空");
        // 改写截断逻辑
        if (contextText.length > 12000) {
            contextText = contextText.substring(0, 12000);
            const lastEnd = Math.max(contextText.lastIndexOf("。"), contextText.lastIndexOf("\n"));
            if (lastEnd > 8000) contextText = contextText.substring(0, lastEnd + 1);
            contextText += "\n... (省略)";
        }

        const prompt = `提炼论文全文草稿，生成结构化摘要。包含：背景、方法、结果、意义，及 3-5 个关键词。`;
        return await llm.callLLM(prompt, contextText);
    };

    if (passedContext) return await doWork(passedContext);
    return await Word.run(async (context) => await doWork(context));
}
