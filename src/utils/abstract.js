import * as llm from "./llm.js";

/**
 * 生成多合一摘要与关键词
 */
export async function generateAbstract() {
    return await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        const text = body.text;
        if (!text || text.trim().length === 0) {
            throw new Error("文档内容为空，无法生成摘要");
        }

        // 截取前 12000 个字符以避免超出常见模型的 Token 限制
        const contextText = text.length > 12000 ? text.substring(0, 12000) + "..." : text;

        const prompt = `请作为资深学术编辑，根据以下论文全文草稿，生成一段结构化的标准学术摘要。
要求：
1. 包含四个维度：研究背景、核心方法、实验/研究结果、结论与意义。
2. 语言要专业、精炼、符合学术规范。
3. 摘要之后，提供 3-5 个专业关键词。

请直接输出摘要和关键词，无需前缀废话。

待处理论文文本：
`;
        
        const abstract = await llm.callLLM(prompt, contextText);
        return abstract;
    });
}
