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
  DIFF_MODE: "wordai_diff_mode",
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
    
    增加冗余与解释性（Verbose Elaboration）：
    
    动词短语扩展： 将简洁的动词或动词短语替换为更长的、带有动作过程描述的短语。
    示例：“管理” -> “开展...的管理工作” 或 “进行管理”
    示例：“交互” -> “进行交互” 或 “开展交互”
    示例：“配置” -> “进行配置”
    示例：“处理” -> “去处理...工作”
    示例：“恢复” -> “进行恢复”
    示例：“实现” -> “得以实现” 或 “来实现”
    增加辅助词/结构： 在句子中添加语法上允许但非必需的词语，使句子更饱满。
    示例：适当增加 “了”、“的”、“地”、“所”、“会”、“可以”、“这个”、“方面”、“当中” 等。
    示例：“提供功能” -> “有...功能” 或 “拥有...功能”
    系统性词汇替换（Systematic Synonym/Phrasing Substitution）：
    
    特定动词/介词/连词替换： 将原文中常用的某些词汇固定地替换为特定的替代词。这是模仿目标风格的关键。
    采用 / 使用 -> 运用 / 选用 / 把...当作...来使用
    基于 -> 鉴于 / 基于...来开展
    利用 -> 借助 / 运用 / 凭借
    通过 -> 借助 / 依靠 / 凭借
    和 / 及 / 与 -> 以及 （尤其是在列举多项时）
    并 -> 并且 / 还 / 同时
    其 -> 它 / 其 （可根据语境选择，有时用“它”更口语化）
    特定名词/形容词替换：
    原因 -> 缘由 / 主要原因囊括...
    符合 -> 契合
    适合 -> 适宜
    特点 -> 特性
    提升 / 提高 -> 提高 / 提升 （可互换使用，保持多样性）
    极大(地) -> 极大程度(上)
    立即 -> 马上
    括号内容处理（Bracket Content Integration/Removal）：
    
    解释性括号： 对于原文中用于解释、举例或说明缩写的括号 (...) 或 （...）：
    优先整合： 尝试将括号内的信息自然地融入句子，使用 “也就是”、“即”、“比如”、“像” 等引导词或考虑直接删除括号。
    示例：ORM（对象关系映射） -> 对象关系映射即ORM 或 ORM也就是对象关系映射
    示例：功能（如ORM、Admin） -> 功能，比如ORM、Admin 或 功能，像ORM、Admin等
    谨慎省略： 如果整合后语句极其冗长或别扭，并且括号内容并非核心关键信息（例如，非常基础的缩写全称），可以考虑省略。但要极其小心，避免丢失重要上下文或示例。 在提供的范例中，有时示例信息被省略了，你可以模仿这一点，但要判断是否会损失过多信息。
    代码/标识符旁括号： 对于紧跟在代码、文件名、类名旁的括号，通常直接移除括号。
    示例：视图 (views.py) 中 -> 视图即views.py中
    示例：权限类 (admin_panel.permissions) -> 权限类 admin_panel.permissions
    句式微调与口语化倾向（Sentence Structure & Colloquial Touch）：
    
    使用“把”字句： 在合适的场景下，倾向于使用“把”字句。
    示例：“会将对象移动” -> “会把对象移动”
    条件句式转换： 将较书面的条件句式改为稍口语化的形式。
    示例：“若...，则...” -> “要是...，那就...” 或 “如果...，就...”
    名词化与动词化转换： 根据需要进行调整，有时将名词性结构展开为动词性结构，反之亦然，以符合更自然的口语表达。
    示例：“为了将...解耦” -> “为了实现...的解耦”
    增加语气词/连接词： 如在句首或句中添加“那么”、“这样”、“同时”等。
    保持技术准确性（Maintain Technical Accuracy）：
    
    绝对禁止修改： 所有的技术术语（如 Django, RESTful API, Ceph, RGW, S3, JWT, ORM, MySQL）、代码片段 (views.py, settings.py, accounts.CustomUser, .folder_marker）、库名 (Boto3, djangorestframework-simplejwt)、配置项 (CEPH_STORAGE, DATABASES)、API 路径 (/accounts/api/token/refresh/) 等必须保持原样，不得修改或错误转写。
    核心逻辑不变： 修改后的句子必须表达与原文完全相同的技术逻辑、因果关系和功能描述。
    执行指令：
    
    请根据以上所有规则，对接下来提供的“原文”进行修改，生成符合上述特定风格的“修改后”文本。务必仔细揣摩每个规则的细节和示例，力求在风格上高度一致。注意不要过于口语化（通常情况下不会过于口语化，有一些比如至于xxx呢，这种的不要有）注意！你输出的内容不应原多于原文！应时刻记得字数和原文相符！注意！不要有‘’xxx呢‘’这种形式，如‘至于vue呢’
    不要第一人称`,
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
    icon: "😌",
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
 * 获取是否开启修订模式
 */
export function getDiffMode() {
  return localStorage.getItem(STORAGE_KEYS.DIFF_MODE) === "true";
}

/**
 * 设置是否开启修订模式
 */
export function setDiffMode(enabled) {
  localStorage.setItem(STORAGE_KEYS.DIFF_MODE, enabled ? "true" : "false");
}

/**
 * 检查配置是否完整
 */
export function isConfigured() {
  return !!(getEndpoint() && getApiKey() && getModel());
}

export { DEFAULT_PROMPTS, DEFAULT_SKIP_RULES };
