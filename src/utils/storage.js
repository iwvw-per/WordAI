/**
 * storage.js - 设置存储管理
 * 使用 localStorage 持久化所有配置
 */

const STORAGE_KEYS = {
  API_ENDPOINT: "wordai_api_endpoint",
  API_KEY: "wordai_api_key",
  MODEL: "wordai_model",
  PROMPTS: "wordai_prompts",
  SKIP_RULES: "wordai_skip_rules",
  TEMPERATURE: "wordai_temperature",
};

// 默认预设提示词
const DEFAULT_PROMPTS = [
  {
    id: "polish",
    name: "润色",
    icon: "✨",
    prompt: "请润色以下文字，保持原意不变，提升表达质量和流畅度。保持原文的语言（中文或英文），不要翻译。只输出润色后的文字，不要输出任何解释。",
    color: "#6366f1",
  },
  {
    id: "deai",
    name: "降AI",
    icon: "🎭",
    prompt: "请改写以下文字，使其更像人类自然书写的风格。避免AI常见的套话和模式化表达，增加口语化、个性化的表达方式。保持原文含义不变，保持原文的语言。只输出改写后的文字，不要输出任何解释。",
    color: "#ec4899",
  },
  {
    id: "rewrite",
    name: "改写",
    icon: "🔄",
    prompt: "请用不同的表达方式改写以下文字，保持原意不变。保持原文的语言。只输出改写后的文字，不要输出任何解释。",
    color: "#f59e0b",
  },
  {
    id: "translate_en",
    name: "至英",
    icon: "🌐",
    prompt: "请将以下文字翻译成英文。只输出翻译结果，不要输出任何解释。",
    color: "#10b981",
  },
  {
    id: "translate_zh",
    name: "至中",
    icon: "🇨🇳",
    prompt: "请将以下文字翻译成中文。只输出翻译结果，不要输出任何解释。",
    color: "#3b82f6",
  },
  {
    id: "fix",
    name: "纠错",
    icon: "🔍",
    prompt: "请检查并修正以下文字中的错别字、语法错误和标点符号问题。保持原文的语言和风格。只输出修正后的文字，不要输出任何解释。",
    color: "#ef4444",
  },
  {
    id: "shorten",
    name: "缩写",
    icon: "📐",
    prompt: "请精简以下文字，去除冗余表达，保留核心含义。保持原文的语言。只输出缩写后的文字，不要输出任何解释。",
    color: "#8b5cf6",
  },
  {
    id: "expand",
    name: "扩写",
    icon: "📝",
    prompt: "请在保持原意的基础上扩写以下文字，增加细节和描述，使内容更加丰富充实。保持原文的语言。只输出扩写后的文字，不要输出任何解释。",
    color: "#0ea5e9",
  },
];

const DEFAULT_SKIP_RULES = {
  headings: true,
  tables: true,
  formulas: true,
  crossReferences: true,
  images: true,
  toc: true,
};

/**
 * 获取 API 端点
 */
export function getEndpoint() {
  return localStorage.getItem(STORAGE_KEYS.API_ENDPOINT) || "";
}

/**
 * 设置 API 端点
 */
export function setEndpoint(endpoint) {
  localStorage.setItem(STORAGE_KEYS.API_ENDPOINT, endpoint.replace(/\/+$/, ""));
}

/**
 * 获取 API Key
 */
export function getApiKey() {
  return localStorage.getItem(STORAGE_KEYS.API_KEY) || "";
}

/**
 * 设置 API Key
 */
export function setApiKey(key) {
  localStorage.setItem(STORAGE_KEYS.API_KEY, key);
}

/**
 * 获取当前选择的模型
 */
export function getModel() {
  return localStorage.getItem(STORAGE_KEYS.MODEL) || "";
}

/**
 * 设置当前模型
 */
export function setModel(model) {
  localStorage.setItem(STORAGE_KEYS.MODEL, model);
}

/**
 * 获取温度参数
 */
export function getTemperature() {
  const temp = localStorage.getItem(STORAGE_KEYS.TEMPERATURE);
  return temp !== null ? parseFloat(temp) : 0.7;
}

/**
 * 设置温度参数
 */
export function setTemperature(temp) {
  localStorage.setItem(STORAGE_KEYS.TEMPERATURE, temp.toString());
}

/**
 * 获取提示词列表
 */
export function getPrompts() {
  const stored = localStorage.getItem(STORAGE_KEYS.PROMPTS);
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      // 自动迁移旧版冗长的默认快捷名称，避免用户手动重置缓存
      let migrated = false;
      parsed.forEach(p => {
        if (p.id === "translate_en" && (p.name === "翻译为英文" || p.name === "英译")) {
          p.name = "至英";
          migrated = true;
        }
        if (p.id === "translate_zh" && (p.name === "翻译为中文" || p.name === "中译")) {
          p.name = "至中";
          migrated = true;
        }
      });
      if (migrated) {
        setPrompts(parsed);
      }
      return parsed;
    } catch {
      return [...DEFAULT_PROMPTS];
    }
  }
  return [...DEFAULT_PROMPTS];
}

/**
 * 设置提示词列表
 */
export function setPrompts(prompts) {
  localStorage.setItem(STORAGE_KEYS.PROMPTS, JSON.stringify(prompts));
}

/**
 * 重置提示词为默认
 */
export function resetPrompts() {
  localStorage.setItem(STORAGE_KEYS.PROMPTS, JSON.stringify(DEFAULT_PROMPTS));
  return [...DEFAULT_PROMPTS];
}

/**
 * 获取跳过规则
 */
export function getSkipRules() {
  const stored = localStorage.getItem(STORAGE_KEYS.SKIP_RULES);
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch {
      return { ...DEFAULT_SKIP_RULES };
    }
  }
  return { ...DEFAULT_SKIP_RULES };
}

/**
 * 设置跳过规则
 */
export function setSkipRules(rules) {
  localStorage.setItem(STORAGE_KEYS.SKIP_RULES, JSON.stringify(rules));
}

/**
 * 检查配置是否完整
 */
export function isConfigured() {
  return !!(getEndpoint() && getApiKey() && getModel());
}

export { DEFAULT_PROMPTS, DEFAULT_SKIP_RULES };
