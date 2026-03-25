/**
 * llm.js - LLM API 客户端
 * 支持 OpenAI 兼容格式的 API 调用和模型列表获取
 */

import { getEndpoint, getApiKey, getModel, getTemperature } from "./storage.js";

/**
 * 获取模型列表
 * @returns {Promise<Array<{id: string, name: string}>>}
 */
export async function fetchModels() {
  const endpoint = getEndpoint();
  const apiKey = getApiKey();

  if (!endpoint || !apiKey) {
    throw new Error("请先配置 API 端点和 API Key");
  }

  const url = `${endpoint}/v1/models`;
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(new Error("请求超时，请检查网络和 API 地址")), 15000); // 15s 超时

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      signal: controller.signal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`获取模型列表失败 (${response.status}): ${errorText}`);
    }

    const data = await response.json();
    let models = [];

    if (data.data && Array.isArray(data.data)) {
      models = data.data.map((m) => ({
        id: m.id,
        name: m.id,
      }));
    } else if (Array.isArray(data)) {
      models = data.map((m) => ({
        id: typeof m === "string" ? m : m.id,
        name: typeof m === "string" ? m : m.id,
      }));
    }

    // 按名称排序
    models.sort((a, b) => a.name.localeCompare(b.name));
    return models;
  } catch (error) {
    if (error.name === "AbortError" || error.message.includes("超时")) {
      throw new Error("连接超时，请检查网络或代理设置");
    }
    if (error.message.includes("获取模型列表失败")) {
      throw error;
    }
    throw new Error(`无法连接到 API 端点: ${error.message}`);
  } finally {
    clearTimeout(timeoutId);
  }
}

/**
 * 调用 LLM API（非流式）
 * @param {string} systemPrompt - 系统提示词
 * @param {string} userContent - 用户内容（选中的文字）
 * @param {AbortSignal} signal - 取消信号
 * @returns {Promise<string>} 处理后的文字
 */
export async function callLLM(systemPrompt, userContent, signal) {
  const endpoint = getEndpoint();
  const apiKey = getApiKey();
  const model = getModel();
  const temperature = getTemperature();

  if (!endpoint || !apiKey || !model) {
    throw new Error("请先完成 API 配置");
  }

  const url = `${endpoint}/v1/chat/completions`;

  const body = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent },
    ],
    temperature: temperature,
    stream: false,
  };

  let fetchSignal = signal;
  let timeoutId;

  if (!signal) {
    const controller = new AbortController();
    timeoutId = setTimeout(() => controller.abort(new Error("请求超时")), 60000); // 60s 超时
    fetchSignal = controller.signal;
  }

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal: fetchSignal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API 调用失败 (${response.status}): ${errorText}`);
    }

    const data = await response.json();

    if (data.choices && data.choices.length > 0) {
      return data.choices[0].message.content.trim();
    }

    throw new Error("API 返回了空的结果");
  } catch (error) {
    if (error.name === "AbortError") {
      throw error; 
    }
    throw error;
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

/**
 * 调用 LLM API（流式）
 * @param {string} systemPrompt - 系统提示词
 * @param {string} userContent - 用户内容
 * @param {function} onChunk - 每收到一段文字时的回调
 * @param {AbortSignal} signal - 取消信号
 * @returns {Promise<string>} 完整的处理后文字
 */
export async function callLLMStream(systemPrompt, userContent, onChunk, signal) {
  const endpoint = getEndpoint();
  const apiKey = getApiKey();
  const model = getModel();
  const temperature = getTemperature();

  if (!endpoint || !apiKey || !model) {
    throw new Error("请先完成 API 配置");
  }

  const url = `${endpoint}/v1/chat/completions`;

  const body = {
    model: model,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent },
    ],
    temperature: temperature,
    stream: true,
  };

  // 支持 90s 无响应超时
  let fetchSignal = signal;
  let timeoutId;
  const controller = new AbortController();

  if (signal) {
    // 监听外部取消
    signal.addEventListener("abort", () => controller.abort(signal.reason));
  }
  
  timeoutId = setTimeout(() => controller.abort(new Error("请求超时，请检查网络或 API 服务状态")), 90000);
  fetchSignal = controller.signal;

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
      signal: fetchSignal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API 调用失败 (${response.status}): ${errorText}`);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullText = "";
    let buffer = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split("\n");
      buffer = lines.pop() || "";

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || !trimmed.startsWith("data: ")) continue;
        const data = trimmed.slice(6);
        if (data === "[DONE]") continue;

        try {
          const parsed = JSON.parse(data);
          const delta = parsed.choices?.[0]?.delta?.content;
          if (delta) {
            fullText += delta;
            // 有数据返回就重置流式获取超时（30s不返回新数据则中断）
            clearTimeout(timeoutId);
            timeoutId = setTimeout(() => controller.abort(new Error("流式输出中断，已超时")), 30000);
            
            if (onChunk) onChunk(delta, fullText);
          }
        } catch {
          // 忽略流数据的单行解析错误
        }
      }
    }

    return fullText.trim();
  } catch (error) {
    if (error.name === "AbortError") {
      throw error;
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

/**
 * 清理 AI 回复中的冗余信息（如引导语和代码块包裹）
 * @param {string} text 
 */
export function cleanAiResponse(text) {
  if (!text) return "";
  
  let result = text.trim();
  // 1. 移除常见的 AI 引导语（支持多行）
  const prefixes = [
    /^([\s\n]*(\*\*|__)?(这里是|这是)?(修改后|润色后|翻译后|重写后|扩展后|缩写后|处理后)的?(内容|文本|结果|段落)?[\s\n]*(\*\*|__)?[\s\n]*[：:\n]+)/i,
    /^(好的[，, ]?|当然[，, ]?|没问题[，, ]?)/i
  ];
  
  for (const regex of prefixes) {
    result = result.replace(regex, "");
  }

  // 2. 移除可能包裹的 Markdown 代码块
  result = result.replace(/^```[a-zA-Z]*\n/i, "").replace(/\n```$/i, "");
  
  return result.trim();
}

/**
 * 测试 API 连接
 * @returns {Promise<{success: boolean, message: string}>}
 */
export async function testConnection() {
  try {
    const models = await fetchModels();
    return {
      success: true,
      message: `连接成功！发现 ${models.length} 个模型`,
      models: models,
    };
  } catch (error) {
    return {
      success: false,
      message: error.message,
      models: [],
    };
  }
}
