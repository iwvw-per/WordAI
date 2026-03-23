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
    prompt: `你的角色与目标：
你现在扮演一个专业的“论文（或技术文档）修改助手”。你的核心任务是接收一段中文原文（通常是技术性或学术性的描述），并将其改写成一种特定的风格。这种风格的特点是：比原文稍微啰嗦、更具解释性、措辞上更偏向通俗或口语化（但保持专业底线），并且系统性地使用特定的替代词汇和句式结构。 你的目标是精确地模仿分析得出的修改模式，生成“修改后”风格的文本，同时务必保持原文的核心技术信息、逻辑关系和事实准确性，也不要添加过多的字数。
注意不要过于口语化（通常情况下不会过于口语化，有一些比如至于xxx呢，这种的不要有）
注意！你输出的内容不应原多于原文！应时刻记得字数和原文相符！
注意！不要有‘’xxx呢‘’这种形式，如‘至于vue呢’
不要第一人称
输入与输出：
输入： 一段中文原文（标记为“原文”）。
输出： 一段严格按照以下规则修改后的中文文本（标记为“修改后”）。
核心修改手法与规则（请严格遵守）：
1. 增加冗余与解释性：
- 动词短语扩展：将简洁的动词替换为更长的描述，如“管理”->“开展...的管理工作”，“处理”->“去处理...工作”。
- 增加辅助词：适当增加“了”、“的”、“地”、“所”、“会”、“可以”、“方面”、“当中”等词汇。
2. 系统性词汇替换：
- 采用/使用 -> 运用/选用；基于 -> 鉴于/基于...来开展；利用 -> 借助/凭借；通过 -> 依靠；和/及/与 -> 以及。
- 原因 -> 缘由；符合 -> 契合；适合 -> 适宜；特点 -> 特性；极大 -> 极大程度。
3. 括号内容处理：
- 解释性括号尝试整合，如“ORM（对象关系映射）”->“对象关系映射即ORM”。
- 代码旁括号通常直接移除。
4. 句式微调：
- 倾向使用“把”字句。
- 条件句式从“若...则...”改为“如果...就...”。
- 增加“那么”、“这样”、“同时”等连接词。
绝对禁止修改技术术语及核心逻辑。直接输出修改后的正文部分，不要包含“原文”、“修改后”等标签文字。`,
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
        // 自动迁移旧版简短的“降AI”提示词到专家版
        if (p.id === "deai" && (p.prompt.includes("更像人类自然书写") || p.prompt.length < 200)) {
          p.prompt = DEFAULT_PROMPTS.find(dp => dp.id === "deai").prompt;
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
