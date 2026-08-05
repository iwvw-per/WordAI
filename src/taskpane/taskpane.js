/**
 * taskpane.js - WordAI 主逻辑
 * 自动执行模式：选中 → 点击 → 自动替换
 */

import "./taskpane.css";
import * as storage from "../utils/storage.js";
import * as llm from "../utils/llm.js";
import * as ooxml from "../utils/ooxml.js";
import * as tableUtils from "../utils/table.js";
import { escapeHtml } from "../utils/html.js";
import { parseSegmentedResponse } from "../utils/llmOutput.js";
import { redactSensitiveText, restoreSensitiveText } from "../utils/privacy.js";

// ==================== 状态 ====================
let appInitialized = false;
let isOfficeReady = false;
let isProcessing = false;
let editingPromptId = null;
let currentAbortController = null;

// 全局多任务异步排队调度队列状态
let globalTaskQueue = [];
let activeTaskCount = 0;
let taskCounter = 0;

// 防重复触发：同一动作 1 秒内忽略重复点击
const ACTION_DEBOUNCE_MS = 1000;
let lastActionTime = 0;

// ==================== 初始化 ====================
if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    isOfficeReady = true;
    initApp();
  });
} else {
  console.warn("Office.js 未加载，插件以独立模式运行");
}

// Fallback：3 秒后兜底初始化 UI
document.addEventListener("DOMContentLoaded", () => {
  setTimeout(() => {
    if (!appInitialized) {
      console.warn("Office JS 未就绪，以独立模式初始化");
      initApp();
    }
  }, 3000);
});

function initApp() {
  if (appInitialized) return;
  appInitialized = true;

  renderActionButtons();
  renderPromptList();
  loadSettings();
  checkConfig();
  initTheme();

  bindTabEvents();
  bindActionEvents();
  bindSettingsEvents();
  bindRoutingModelSelects();
  bindModalEvents();
  bindAcademicEvents();
  bindDeckEvents();

  handleUrlAction();
}

// 绑定快捷快捷学术操作台卡片的事件
function bindDeckEvents() {
  const btnClose = document.getElementById("btn-close-deck");
  const deck = document.getElementById("quick-action-deck");
  
  if (btnClose && deck) {
    btnClose.addEventListener("click", () => {
      deck.classList.add("hidden");
    });
  }

  const polishBtn = document.getElementById("btn-deck-polish");
  const deaiBtn = document.getElementById("btn-deck-deai");
  const coolBtn = document.getElementById("btn-deck-cool");

  const runAction = async (actionType) => {
    const prompts = storage.getPrompts();
    let targetPrompt = null;
    if (actionType === "polish") {
      targetPrompt = prompts.find(p => p.id === "polish" || p.name.includes("润色"));
    } else if (actionType === "deai") {
      targetPrompt = prompts.find(p => p.id === "deai" || p.name.includes("降"));
    } else if (actionType === "cool") {
      targetPrompt = prompts.find(p => p.id === "cool" || p.name.includes("降温"));
    }

    if (targetPrompt) {
      // 自动切回“操作”选项卡
      const tabBtn = document.getElementById("tab-actions");
      if (tabBtn) tabBtn.click();
      await executeAction(targetPrompt.prompt, targetPrompt.name, null);
    }
  };

  if (polishBtn) {
    polishBtn.addEventListener("click", () => runAction("polish"));
  }
  if (deaiBtn) {
    deaiBtn.addEventListener("click", () => runAction("deai"));
  }
  if (coolBtn) {
    coolBtn.addEventListener("click", () => runAction("cool"));
  }
}

// 处理来自 Word 右键菜单跳转传递的 action 参数并自动触发
function handleUrlAction() {
  try {
    const params = new URLSearchParams(window.location.search);
    const action = params.get("action");
    if (!action) return;

    // 延迟 800ms 确保 Office/Word JS 以及 DOM 真正初始化就绪
    setTimeout(async () => {
      // 抹除 URL 中的 action 参数，防止后续页面重载或切换主题时发生二次意外触发（防老旧 IE 内核/ Office 沙盒崩溃保护）
      if (window.history && typeof window.history.replaceState === "function") {
        window.history.replaceState({}, document.title, window.location.pathname);
      }

      if (action === "general") {
        const deck = document.getElementById("quick-action-deck");
        if (deck) {
          deck.classList.remove("hidden");
        }
        // 自动切回“操作”选项卡
        const tabBtn = document.getElementById("tab-actions");
        if (tabBtn) tabBtn.click();
        return;
      }

      const prompts = storage.getPrompts();
      let targetPrompt = null;

      if (action === "polish") {
        targetPrompt = prompts.find(p => p.id === "polish" || p.name.includes("润色"));
      } else if (action === "deai") {
        targetPrompt = prompts.find(p => p.id === "deai" || p.name.includes("降"));
      } else if (action === "cool") {
        targetPrompt = prompts.find(p => p.id === "cool" || p.name.includes("降温"));
      }

      if (targetPrompt) {
        // 自动切回“操作”选项卡
        const tabBtn = document.getElementById("tab-actions");
        if (tabBtn) tabBtn.click();

        // 立即静默触发 AI 任务
        await executeAction(targetPrompt.prompt, targetPrompt.name, null);
      }
    }, 800);
  } catch (err) {
    console.error("解析并触发右键快捷动作出错:", err);
  }
}

// ==================== 主题 ====================
function initTheme() {
  const saved = localStorage.getItem("wordai_theme");
  const btn = document.getElementById("theme-toggle-btn");
  if (saved === "light") {
    setTheme("light");
    if (btn) btn.textContent = "☀️";
  } else if (saved === "dark") {
    setTheme("dark");
    if (btn) btn.textContent = "🌙";
  } else {
    applyAutoTheme();
    if (btn) btn.textContent = "🌓";
  }

  // 监听系统主题变化（仅在用户未手动选择主题时跟随系统，兼容旧内核）
  const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
  const handleThemeChange = () => {
    const userTheme = localStorage.getItem("wordai_theme");
    if (!userTheme) {
      applyAutoTheme();
    }
  };
  try {
    mediaQuery.addEventListener("change", handleThemeChange);
  } catch (e) {
    try {
      mediaQuery.addListener(handleThemeChange);
    } catch (err) {
      console.warn("当前浏览器内核不支持系统主题变化监听", err);
    }
  }

  // 注意：Office.EventType.OfficeThemeChanged 仅支持 Outlook，Word 不支持，已移除
}

function applyAutoTheme() {
  try {
    // 优先尝试 Office 主题颜色
    if (Office.context?.officeTheme) {
      const bg = Office.context.officeTheme.bodyBackgroundColor;
      if (bg) {
        const hex = bg.replace("#", "");
        if (hex.length === 6) {
          const brightness =
            (parseInt(hex.substring(0, 2), 16) * 299 +
              parseInt(hex.substring(2, 4), 16) * 587 +
              parseInt(hex.substring(4, 6), 16) * 114) / 1000;
          setTheme(brightness < 128 ? "dark" : "light");
          return;
        }
      }
    }
  } catch { }

  // 降级为系统设置
  const isSysDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  setTheme(isSysDark ? "dark" : "light");
}

function setTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
}

function toggleTheme() {
  const btn = document.getElementById("theme-toggle-btn");
  let current = localStorage.getItem("wordai_theme");

  if (!current) {
    // 当前是自动模式，切换到确定的浅色或深色
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const next = isDark ? "light" : "dark";
    setTheme(next);
    localStorage.setItem("wordai_theme", next);
    if (btn) btn.textContent = next === "light" ? "☀️" : "🌙";
  } else if (current === "light") {
    // 浅色 -> 深色
    setTheme("dark");
    localStorage.setItem("wordai_theme", "dark");
    if (btn) btn.textContent = "🌙";
  } else {
    // 深色 -> 自动
    localStorage.removeItem("wordai_theme");
    applyAutoTheme();
    if (btn) btn.textContent = "🌓";
  }
}

// ==================== Tab ====================
function bindTabEvents() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab;
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      document.querySelectorAll(".page").forEach((p) => p.classList.remove("active"));
      document.getElementById(`page-${tab}`).classList.add("active");
    });
  });

  document.getElementById("go-settings-btn")?.addEventListener("click", () => {
    document.querySelector('[data-tab="settings"]').click();
  });

  document.getElementById("theme-toggle-btn")?.addEventListener("click", toggleTheme);
}

// ==================== 操作按钮 ====================
function renderActionButtons() {
  const grid = document.getElementById("action-grid");
  const prompts = storage.getPrompts();

  grid.innerHTML = prompts
    .map(
      (p) => `
    <button class="action-btn" data-prompt-id="${escapeHtml(p.id)}" style="--btn-color: ${escapeHtml(p.color)}" title="${escapeHtml(p.name)}">
      <span class="action-icon">${escapeHtml(p.icon)}</span>
      <span class="action-name">${escapeHtml(p.name)}</span>
    </button>
  `
    )
    .join("");

  grid.querySelectorAll(".action-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const prompt = prompts.find((p) => p.id === btn.dataset.promptId);
      if (prompt) executeAction(prompt.prompt, prompt.name, btn);
    });
  });
}

function bindActionEvents() {
  document.getElementById("custom-run-btn").addEventListener("click", () => {
    const text = document.getElementById("custom-prompt").value.trim();
    if (!text) {
      showToast("请输入自定义指令", "warning");
      return;
    }
    executeAction(text, "自定义", document.getElementById("custom-run-btn"));
  });

  document.getElementById("custom-prompt").addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
      e.preventDefault();
      document.getElementById("custom-run-btn").click();
    }
  });

  document.getElementById("btn-clear-console")?.addEventListener("click", () => {
    const body = document.getElementById("pipeline-console-body");
    const title = document.getElementById("pipeline-console-title");
    if (body) {
      body.innerHTML = '<div class="console-line system">已清空</div>';
    }
    if (title) {
      title.textContent = "就绪";
    }
  });
}

// ==================== 核心：一键执行 ====================
async function executeAction(systemPrompt, actionName, triggerBtn) {
  // 防重复触发：避免用户因无反馈连点导致同一选区重复入队
  const now = Date.now();
  if (now - lastActionTime < ACTION_DEBOUNCE_MS) {
    showToast("正在处理中，请稍候", "warning");
    return;
  }
  lastActionTime = now;

  if (!storage.isConfigured()) {
    showConfigBanner();
    showToast("请先完成 API 配置", "warning");
    return;
  }

  if (!isOfficeReady) {
    showToast("请在 Word 中使用", "error");
    return;
  }

  const consoleSection = document.getElementById("pipeline-console-section");
  const consoleBody = document.getElementById("pipeline-console-body");

  // 展示控制台面板
  if (consoleSection) consoleSection.classList.remove("hidden");

  // 触发按钮触感反馈 + 立即给出处理中的状态提示（避免误以为没反应）
  setAllActionsLoading(true, triggerBtn);
  showInlineStatus("processing", `正在处理：${actionName}...`);

  let segments = [];
  try {
    // ⚡ 瞬时锁定：瞬间在 Word 中锁定并锚定该任务对应选区（约耗时 50ms）
    segments = await ooxml.markSelection();
    if (!segments || segments.length === 0) {
      throw new Error("请先在 Word 中选中需要处理的文字内容");
    }

    // 提示：含公式被跳过的段落
    const skippedInfo = segments.skippedInfo || {};
    if (skippedInfo.formula > 0 && consoleBody) {
      const tip = document.createElement("div");
      tip.className = "console-line system";
      tip.textContent = `${skippedInfo.formula} 个段落含公式，已跳过`;
      consoleBody.appendChild(tip);
      consoleBody.scrollTop = consoleBody.scrollHeight;
    }
  } catch (err) {
    hideInlineStatus();
    showToast(err.message, "error");
    return;
  }

  // 打包任务压入全局队列
  taskCounter++;
  const taskId = `task_${Date.now()}_${taskCounter}`;
  const pipelineTask = {
    id: taskId,
    num: taskCounter,
    actionName: actionName,
    systemPrompt: systemPrompt,
    segments: segments,
    status: "pending",
    createdAt: Date.now(),
    segmentCount: segments.length,
    abortController: new AbortController(),
  };

  globalTaskQueue.push(pipelineTask);

  // 在滚动控制台动态注入精简的单行任务项
  if (consoleBody) {
    const placeholder = consoleBody.querySelector(".console-line.system");
    if (placeholder && placeholder.textContent.includes("就绪")) {
      placeholder.remove();
    }

    const taskLine = document.createElement("div");
    taskLine.id = `console-task-${taskId}`;
    taskLine.className = "console-line pending";
    taskLine.innerHTML = `
      <span class="task-badge">#${taskCounter}</span>
      <span class="task-status-icon" id="console-status-${taskId}">⏳</span>
      <span class="task-name" title="${escapeHtml(actionName)}">${escapeHtml(actionName)}</span>
      <span class="task-progress" id="console-progress-${taskId}">等待中</span>
      <button class="task-cancel-btn" id="console-cancel-${taskId}" title="取消任务">✕</button>
    `;

    consoleBody.appendChild(taskLine);
    consoleBody.scrollTop = consoleBody.scrollHeight;

    // 取消按钮：排队中直接移出队列；执行中中止请求并回滚
    taskLine.querySelector(`#console-cancel-${taskId}`).addEventListener("click", () => {
      const idx = globalTaskQueue.findIndex((t) => t.id === taskId);
      if (idx >= 0) {
        globalTaskQueue.splice(idx, 1);
        taskLine.className = "console-line error";
        const icon = document.getElementById(`console-status-${taskId}`);
        const prog = document.getElementById(`console-progress-${taskId}`);
        if (icon) icon.textContent = "🛑";
        if (prog) prog.textContent = "已取消";
        updateConsoleTitleState();
      } else {
        pipelineTask.abortController.abort(new Error("已取消"));
      }
    });
  }

  showToast(`任务 #${taskCounter} [${actionName}] 已加入队列`, "success");
  hideInlineStatus();

  // 触发全局调度器
  processTaskQueue();
}

/**
 * 全局多任务异步排队调度引擎（支持自定义多线程并行）
 */
async function processTaskQueue() {
  const concurrencyLimit = storage.getConcurrencyLimit();

  while (activeTaskCount < concurrencyLimit && globalTaskQueue.length > 0) {
    const currentTask = globalTaskQueue.shift();
    activeTaskCount++;
    
    // 非阻塞地启动该任务的异步执行链
    runSingleTaskAsync(currentTask).catch(err => {
      console.error("执行任务异步链出错:", err);
    });
  }

  updateConsoleTitleState();
}

/**
 * 非阻塞执行单个任务的完整异步处理链
 */
async function runSingleTaskAsync(currentTask) {
  const consoleBody = document.getElementById("pipeline-console-body");

  currentTask.status = "processing";
  const taskId = currentTask.id;
  const taskLine = document.getElementById(`console-task-${taskId}`);
  const progressSpan = document.getElementById(`console-progress-${taskId}`);
  const statusIcon = taskLine ? taskLine.querySelector(".task-status-icon") : null;

  if (taskLine) {
    taskLine.className = "console-line processing";
  }
  if (statusIcon) statusIcon.textContent = "⚡";
  updateConsoleTitleState();

  const signal = currentTask.abortController.signal;
  currentAbortController = currentTask.abortController; // 接线全局“中断”按钮
  const segments = currentTask.segments; // 获取所有选中的自然段
  let routedModel = storage.getRoutedModel(currentTask.actionName);

  let retryCount = 0;
  const retryLimit = 3;
  let success = false;

  try {
    while (retryCount < retryLimit) {
      if (signal.aborted) break;

      try {
        if (progressSpan) {
          if (retryCount > 0) {
            progressSpan.textContent = `重试 ${retryCount}/3`;
          } else {
            progressSpan.textContent = "处理中";
          }
          if (consoleBody) consoleBody.scrollTop = consoleBody.scrollHeight;
        }

        // 1. 拼接带段落隔离标签的发送大文本
        let normalizedInputText = "";
        let hasShields = false;
        for (let i = 0; i < segments.length; i++) {
          const seg = segments[i];
          if (seg.refMap && seg.refMap.length > 0) {
            hasShields = true;
          }
          const textClean = seg.text.replace(/\r/g, "\n");
          normalizedInputText += `<p id="${i}">${textClean}</p>\n`;
        }

        // 2. 拼装红线限制提示词
        let redLine = "";
        if (hasShields) {
          redLine += "\n\n【绝对禁令】：文中的 [REF_N], [EQN_N], [FNOTE_N] 是物理引用或公式锚点，你必须原封不动地保留所有此类标记（包括内部的类型、编号以及外层的英文中括号 []），必须将其放置在改写后对应的语义位置。严禁删除、修改括号类型（不能改为 【】 或 『』等）！";
        }
        const privacyEnabled = storage.getPrivacyMode();
        let privacyContext = { text: normalizedInputText, replacements: [] };
        if (privacyEnabled) {
          privacyContext = redactSensitiveText(normalizedInputText);
          if (privacyContext.replacements.length > 0) {
            redLine += "\n\n【隐私占位符红线】：文中的 [[WAI_SECRET_N]] 是用户隐私占位符，必须原封不动保留，不要解释、翻译、拆分或改写。";
          }
        }
        // 强力注入段落隔离协议，约束大模型输出
        redLine += "\n\n【段落标签绝对保留红线】：";
        redLine += "\n1. 输入的文本由多个由 <p id=\"N\">...</p> 包裹的自然段组成，各个段落的物理顺序非常关键。";
        redLine += "\n2. 你必须对每个段落 <p id=\"N\"> 内部的文本进行独立的润色或修改。";
        redLine += "\n3. 你必须原封不动地返回所有的段落外层 HTML 标签（即 <p id=\"N\"> 和 </p>），原封不动地保留其原有的 id 编号和原有的段落物理顺序。";
        redLine += "\n4. 严禁将多个标签内的文本合并到同一个段落里，严禁增减、拆分或删除任何标签！";
        redLine += "\n5. 请严格输出如下格式的内容：<p id=\"0\">第一段修改后文本</p>\\n<p id=\"1\">第二段修改后文本</p>";

        const finalPrompt = redLine ? (currentTask.systemPrompt + redLine + "\n") : currentTask.systemPrompt;

        // 触发流式输出并在控制终端行渲染 delta
        let raw = await llm.callLLMStream(finalPrompt, privacyContext.text, (delta, currentText) => {
          if (progressSpan && !signal.aborted) {
            let displaySnippet = currentText.replace(/<\/?p[^>]*>|\[PARAGRAPH_\d+\]|\n/gi, "");
            if (displaySnippet.length > 15) displaySnippet = "..." + displaySnippet.slice(-15);
            progressSpan.textContent = `⚡ 改写中: "${displaySnippet}█"`;
          }
        }, signal, { model: routedModel });
        raw = restoreSensitiveText(raw, privacyContext.replacements);

        if (signal.aborted) throw new Error("已取消");

        if (progressSpan) {
          progressSpan.textContent = `🧩 恢复排版中...`;
        }

        // 3. 解析大模型返回的标签隔离子串
        const cleanRaw = llm.cleanAiResponse(raw);
        const parsedTexts = parseSegmentedResponse(cleanRaw, segments);
        // 4. 精准逐个段落回填！各段落各回各家，100% 保持 Word 原生物理段落样式！
        for (let i = 0; i < segments.length; i++) {
          const seg = segments[i];
          const aiText = parsedTexts[i] || seg.text; // 保底，无此段改写就写回原文
          if (aiText) {
            await ooxml.replaceSingleMarkedContent(aiText, seg.refMap, seg.boundaryTags, seg.text);
          }
        }

        success = true;
        break; // 成功后跳出重试循环

      } catch (err) {
        if (err.name === "AbortError" || err.message === "已取消") {
          break; // 手动取消退场
        }

        retryCount++;
        if (retryCount < retryLimit) {
          console.warn(`任务 #${currentTask.num} 失败，正在进行第 ${retryCount} 次重试. 错误: ${err.message}`);
          await new Promise(resolve => setTimeout(resolve, 500)); // 重试前避让延时
        } else {
          console.error(`任务 #${currentTask.num} 在重试 ${retryLimit} 次后依然失败. 错误: ${err.message}`);
          // 彻底失败：还原引用遮罩并清理边界标记，避免文档残留 [REF_N] 占位符
          try {
            await ooxml.rollbackSegments(segments);
          } catch (rollbackErr) {
            console.error("任务失败回滚出错:", rollbackErr);
          }
          break;
        }
      }
    }
  } finally {
    // 释放全局取消引用（仅当仍指向本任务时）
    if (currentAbortController === currentTask.abortController) {
      currentAbortController = null;
    }

    // 更新 Task 行的完成或中止状态（独立 try，防止 UI 异常卡死调度队列）
    try {
      if (taskLine) {
        if (signal.aborted) {
          taskLine.className = "console-line error";
          if (statusIcon) statusIcon.textContent = "🛑";
          if (progressSpan) progressSpan.textContent = `已中止`;
        } else if (success) {
          taskLine.className = "console-line done";
          if (statusIcon) statusIcon.textContent = "✅";
          if (progressSpan) progressSpan.textContent = `成功`;
        } else {
          taskLine.className = "console-line error";
          if (statusIcon) statusIcon.textContent = "❌";
          if (progressSpan) progressSpan.textContent = `失败`;
        }
        if (consoleBody) consoleBody.scrollTop = consoleBody.scrollHeight;
      }

      storage.addTaskHistory({
        id: currentTask.id,
        actionName: currentTask.actionName,
        status: signal.aborted ? "aborted" : success ? "success" : "failed",
        model: routedModel,
        segmentCount: currentTask.segmentCount || segments.length,
        durationMs: Date.now() - currentTask.createdAt,
        privacy: storage.getPrivacyMode(),
        createdAt: currentTask.createdAt,
      });
      renderTaskHistory();
    } catch (err) {
      console.error("更新任务状态失败:", err);
    }

    // 释放并发任务计数，并重新触发队列调度（无论成功失败都必须执行）
    activeTaskCount--;
    updateConsoleTitleState();
    processTaskQueue();
  }
}

/**
 * 动态刷新控制台总标题的线程执行状态
 */
function updateConsoleTitleState() {
  const consoleTitle = document.getElementById("pipeline-console-title");
  if (!consoleTitle) return;
  if (activeTaskCount > 0) {
    consoleTitle.textContent = `执行中 (${activeTaskCount})`;
  } else {
    consoleTitle.textContent = globalTaskQueue.length > 0 ? "等待中" : "就绪";
  }
}

function setAllActionsLoading(loading, activeBtn) {
  // ⚡ 彻底废除将按钮变灰禁用（disabled = true）的反模式，允许不间断多任务追加！
  if (loading && activeBtn) {
    activeBtn.classList.add("active-loading");
    // 瞬时反馈 600ms 后自动移除加载闪烁状态
    setTimeout(() => {
      activeBtn.classList.remove("active-loading");
    }, 600);
  }
}

// ==================== 内联状态条 ====================
function showInlineStatus(type, message, canCancel = false) {
  let bar = document.getElementById("inline-status");
  if (!bar) {
    bar = document.createElement("div");
    bar.id = "inline-status";
    const grid = document.getElementById("action-grid");
    if (grid) {
      grid.parentNode.insertBefore(bar, grid.nextSibling);
    }
  }

  bar.className = `inline-status ${type}`;
  bar.style.opacity = "1";

  let iconHtml = "";
  if (type === "processing") {
    iconHtml = `<div class="inline-spinner"></div>`;
  } else if (type === "done") {
    iconHtml = `<span class="inline-icon done">✓</span>`;
  } else if (type === "error") {
    iconHtml = `<span class="inline-icon error">✕</span>`;
  }

  let actionHtml = "";
  if (type === "processing" && canCancel) {
    actionHtml = `<button class="btn btn-sm btn-ghost" id="inline-cancel-btn" style="padding: 2px 6px; font-size: 10px; flex-shrink: 0; margin-left: 4px;">中断</button>`;
  } else if (type === "done") {
    actionHtml = `<button class="btn btn-sm btn-ghost" id="inline-dismiss-btn" style="padding: 2px 6px; font-size: 10px; opacity: 0.7; flex-shrink: 0; margin-left: 4px;" title="关闭">✕</button>`;
  }

  bar.innerHTML = `
    ${iconHtml}
    <div class="status-content">
      <span class="status-text">${escapeHtml(message)}</span>
    </div>
    ${actionHtml}
  `;

  // 检查是否需要滚动
  const content = bar.querySelector(".status-content");
  const text = bar.querySelector(".status-text");
  if (content && text) {
    // 延迟一帧确保 DOM 渲染完成
    requestAnimationFrame(() => {
      const overflow = text.scrollWidth - content.clientWidth;
      if (overflow > 0) {
        text.style.setProperty("--scroll-dist", `-${overflow + 10}px`);
        text.classList.add("marquee-active");
      }
    });
  }

  // 绑定取消事件
  const cancelBtn = document.getElementById("inline-cancel-btn");
  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => {
      if (typeof currentAbortController !== "undefined" && currentAbortController) {
        currentAbortController.abort(new Error("已取消"));
      }
    });
  }

  // 绑定完成状态的关闭按钮
  const dismissBtn = document.getElementById("inline-dismiss-btn");
  if (dismissBtn) {
    dismissBtn.addEventListener("click", () => hideInlineStatus());
  }
}

function hideInlineStatus() {
  const bar = document.getElementById("inline-status");
  if (bar) {
    bar.style.opacity = "0";
    setTimeout(() => bar.remove(), 200);
  }
}

// ==================== 设置 ====================
function loadSettings() {
  document.getElementById("api-endpoint").value = storage.getEndpoint();
  document.getElementById("api-key").value = storage.getApiKey();

  const temp = storage.getTemperature();
  document.getElementById("temperature-slider").value = temp;
  document.getElementById("temperature-value").textContent = temp;

  const model = storage.getModel();
  if (model) {
    const select = document.getElementById("model-select");
    const opt = document.createElement("option");
    opt.value = model;
    opt.textContent = model;
    opt.selected = true;
    select.appendChild(opt);
  }

  const rules = storage.getSkipRules();
  document.getElementById("skip-headings").checked = rules.headings;
  document.getElementById("skip-tables").checked = rules.tables;
  document.getElementById("skip-formulas").checked = rules.formulas;
  document.getElementById("skip-crossrefs").checked = rules.crossReferences;
  document.getElementById("skip-images").checked = rules.images;
  document.getElementById("skip-toc").checked = rules.toc;

  document.getElementById("diff-mode-toggle").checked = storage.getDiffMode();
  document.getElementById("privacy-mode-toggle").checked = storage.getPrivacyMode();
  document.getElementById("model-routing-toggle").checked = storage.getModelRouting();
  document.getElementById("input-tracked-author").value = storage.getTrackedAuthor();
  populateModelSelects([]);
  document.getElementById("terminology-bank-input").value = storage.getTerminologyBankRaw();

  const concurrency = storage.getConcurrencyLimit();
  const slider = document.getElementById("concurrency-slider");
  const valueBadge = document.getElementById("concurrency-value");
  if (slider && valueBadge) {
    slider.value = concurrency;
    valueBadge.textContent = concurrency;
  }
  renderTaskHistory();
}

function bindSettingsEvents() {
  // 自动保存
  document.getElementById("api-endpoint").addEventListener("change", (e) => {
    storage.setEndpoint(e.target.value);
    checkConfig();
  });
  document.getElementById("api-key").addEventListener("change", (e) => {
    storage.setApiKey(e.target.value);
    checkConfig();
  });
  document.getElementById("toggle-key-btn").addEventListener("click", () => {
    const inp = document.getElementById("api-key");
    inp.type = inp.type === "password" ? "text" : "password";
  });
  document.getElementById("temperature-slider").addEventListener("input", (e) => {
    const v = parseFloat(e.target.value);
    document.getElementById("temperature-value").textContent = v.toFixed(1);
    storage.setTemperature(v);
  });
  document.getElementById("fetch-models-btn").addEventListener("click", fetchModelList);
  // 点击下拉框时自动刷新模型列表（30 秒防抖）
  let lastModelFetchTime = 0;
  document.getElementById("model-select").addEventListener("focus", () => {
    const now = Date.now();
    if (now - lastModelFetchTime > 30000 && storage.getEndpoint() && storage.getApiKey()) {
      lastModelFetchTime = now;
      fetchModelList();
    }
  });
  document.getElementById("model-select").addEventListener("change", (e) => {
    storage.setModel(e.target.value);
    checkConfig();
  });
  document.getElementById("test-connection-btn").addEventListener("click", testApiConnection);

  // 跳过规则
  ["skip-headings", "skip-tables", "skip-formulas", "skip-crossrefs", "skip-images", "skip-toc"].forEach((id) => {
    document.getElementById(id).addEventListener("change", () => {
      storage.setSkipRules({
        headings: document.getElementById("skip-headings").checked,
        tables: document.getElementById("skip-tables").checked,
        formulas: document.getElementById("skip-formulas").checked,
        crossReferences: document.getElementById("skip-crossrefs").checked,
        images: document.getElementById("skip-images").checked,
        toc: document.getElementById("skip-toc").checked,
      });
    });
  });

  // 显示对比
  document.getElementById("diff-mode-toggle").addEventListener("change", (e) => {
    storage.setDiffMode(e.target.checked);
  });

  // 修订作者名
  document.getElementById("input-tracked-author")?.addEventListener("change", (e) => {
    storage.setTrackedAuthor(e.target.value);
  });

  // 配置备份/还原
  const configModal = document.getElementById("config-modal");
  const configText = document.getElementById("config-modal-text");
  const configTitle = document.getElementById("config-modal-title");
  const configOk = document.getElementById("config-modal-ok");

  document.getElementById("btn-backup-config")?.addEventListener("click", () => {
    if (!configModal || !configText) return;
    configText.readOnly = true;
    configText.value = storage.exportConfig();
    configTitle.textContent = "配置备份（复制并保存）";
    configOk.classList.add("hidden");
    configModal.classList.remove("hidden");
  });

  document.getElementById("btn-restore-config")?.addEventListener("click", () => {
    if (!configModal || !configText) return;
    configText.readOnly = false;
    configText.value = "";
    configTitle.textContent = "还原配置（粘贴 JSON）";
    configOk.classList.remove("hidden");
    configModal.classList.remove("hidden");
  });

  document.getElementById("config-modal-cancel")?.addEventListener("click", () => {
    configModal.classList.add("hidden");
  });
  document.getElementById("config-modal-ok")?.addEventListener("click", () => {
    try {
      const count = storage.importConfig(configText.value);
      showToast(`已还原 ${count} 项配置`, "success");
      configModal.classList.add("hidden");
      loadSettings();
      checkConfig();
      renderActionButtons();
      renderPromptList();
    } catch (err) {
      showToast("还原失败：配置格式无效", "error");
    }
  });
  configModal?.querySelector(".modal-overlay")?.addEventListener("click", () => {
    configModal.classList.add("hidden");
  });

  // 接受全部修订标记
  document.getElementById("btn-accept-diff")?.addEventListener("click", async () => {
    try {
      const btn = document.getElementById("btn-accept-diff");
      if (btn) btn.disabled = true;
      const count = await ooxml.acceptAllDiff();
      showToast(count > 0 ? `已接受 ${count} 处修订` : "没有发现待接受的修订", count > 0 ? "success" : "info");
    } catch (err) {
      showToast("接受修订失败: " + err.message, "error");
    } finally {
      const btn = document.getElementById("btn-accept-diff");
      if (btn) btn.disabled = false;
    }
  });

  document.getElementById("privacy-mode-toggle").addEventListener("change", (e) => {
    storage.setPrivacyMode(e.target.checked);
  });
  document.getElementById("model-routing-toggle").addEventListener("change", (e) => {
    storage.setModelRouting(e.target.checked);
  });
  document.getElementById("fast-model-input").addEventListener("change", (e) => {
    storage.setFastModel(e.target.value);
  });
  document.getElementById("quality-model-input").addEventListener("change", (e) => {
    storage.setQualityModel(e.target.value);
  });
  document.getElementById("save-terminology-bank-btn").addEventListener("click", () => {
    storage.setTerminologyBankRaw(document.getElementById("terminology-bank-input").value);
    showToast("术语库已保存", "success");
  });
  document.getElementById("clear-task-history-btn").addEventListener("click", () => {
    storage.clearTaskHistory();
    renderTaskHistory();
    showToast("任务历史已清空", "success");
  });

  // 并发控制
  const concurrencySlider = document.getElementById("concurrency-slider");
  const concurrencyValue = document.getElementById("concurrency-value");
  if (concurrencySlider) {
    concurrencySlider.addEventListener("input", (e) => {
      const val = parseInt(e.target.value, 10);
      if (concurrencyValue) concurrencyValue.textContent = val;
      storage.setConcurrencyLimit(val);
    });
  }

  // 提示词管理
  document.getElementById("reset-prompts-btn").addEventListener("click", () => {
    storage.resetPrompts();
    renderPromptList();
    renderActionButtons();
    showToast("已恢复默认", "success");
  });
  document.getElementById("add-prompt-btn").addEventListener("click", () => openPromptModal(null));
}

async function fetchModelList() {
  const btn = document.getElementById("fetch-models-btn");
  const status = document.getElementById("model-status");
  const select = document.getElementById("model-select");

  const ep = document.getElementById("api-endpoint").value;
  const key = document.getElementById("api-key").value;
  if (ep) storage.setEndpoint(ep);
  if (key) storage.setApiKey(key);

  btn.disabled = true;
  btn.textContent = "⏳";
  status.textContent = "获取中...";

  try {
    const models = await llm.fetchModels();
    select.innerHTML = "";

    if (models.length === 0) {
      select.innerHTML = '<option value="">-- 无模型 --</option>';
      status.textContent = "未发现模型";
    } else {
      const cur = storage.getModel();
      models.forEach((m) => {
        const o = document.createElement("option");
        o.value = m.id;
        o.textContent = m.name;
        if (m.id === cur) o.selected = true;
        select.appendChild(o);
      });
      if (!cur || !models.find((m) => m.id === cur)) {
        storage.setModel(models[0].id);
        select.value = models[0].id;
      }
      status.textContent = `${models.length} 个模型`;
      status.style.color = "var(--success)";
      populateModelSelects(models);
      checkConfig();
    }
  } catch (err) {
    status.textContent = err.message;
    status.style.color = "var(--error)";
  } finally {
    btn.disabled = false;
    btn.textContent = "🔄";
    setTimeout(() => (status.style.color = ""), 3000);
  }
}

/**
 * 填充“快速模型 / 质量模型”两个路由下拉框
 * @param {Array} models - [{ id, name }] 模型列表；为空则仅保留默认项与当前已存值
 */
function populateModelSelects(models) {
  const fastSelect = document.getElementById("fast-model-input");
  const qualitySelect = document.getElementById("quality-model-input");
  if (!fastSelect || !qualitySelect) return;

  const savedPairs = [
    [fastSelect, storage.getFastModel()],
    [qualitySelect, storage.getQualityModel()],
  ];

  for (const [select, saved] of savedPairs) {
    const html = ['<option value="">-- 默认模型 --</option>'];
    const added = new Set();
    if (saved) {
      html.push(`<option value="${escapeHtml(saved)}" selected>${escapeHtml(saved)}</option>`);
      added.add(saved);
    }
    for (const m of models || []) {
      if (!m || !m.id || added.has(m.id)) continue;
      html.push(`<option value="${escapeHtml(m.id)}">${escapeHtml(m.name || m.id)}</option>`);
      added.add(m.id);
    }
    select.innerHTML = html.join("");
    select.value = saved || "";
  }
}

// 快速/质量模型下拉框：聚焦时自动刷新模型列表（30 秒防抖）
function bindRoutingModelSelects() {
  let lastFetchTime = 0;
  const refresh = (selectId) => {
    const now = Date.now();
    if (now - lastFetchTime > 30000 && storage.getEndpoint() && storage.getApiKey()) {
      lastFetchTime = now;
      llm.fetchModels()
        .then((models) => {
          populateModelSelects(models);
          checkConfig();
        })
        .catch((err) => {
          console.warn("刷新路由模型列表失败:", err.message);
        });
    }
  };
  ["fast-model-input", "quality-model-input"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("focus", () => refresh(id));
  });
}

async function testApiConnection() {
  const btn = document.getElementById("test-connection-btn");
  const el = document.getElementById("connection-result");

  const ep = document.getElementById("api-endpoint").value;
  const key = document.getElementById("api-key").value;
  if (ep) storage.setEndpoint(ep);
  if (key) storage.setApiKey(key);

  btn.disabled = true;
  btn.innerHTML = '<span class="btn-icon">⏳</span> 测试中...';

  const result = await llm.testConnection();
  el.classList.remove("hidden", "success", "error");
  el.classList.add(result.success ? "success" : "error");
  el.textContent = result.message;

  if (result.success && result.models.length > 0) {
    const select = document.getElementById("model-select");
    select.innerHTML = "";
    const cur = storage.getModel();
    result.models.forEach((m) => {
      const o = document.createElement("option");
      o.value = m.id;
      o.textContent = m.name;
      if (m.id === cur) o.selected = true;
      select.appendChild(o);
    });
    if (!cur && result.models.length > 0) storage.setModel(result.models[0].id);
    populateModelSelects(result.models);
    checkConfig();
  }

  btn.disabled = false;
  btn.innerHTML = '<span class="btn-icon">🔗</span> 测试连接';
}

// ==================== 配置检查 ====================
function checkConfig() {
  const banner = document.getElementById("config-banner");
  if (storage.isConfigured()) banner.classList.add("hidden");
  else banner.classList.remove("hidden");
}

function showConfigBanner() {
  document.getElementById("config-banner").classList.remove("hidden");
}

function renderTaskHistory() {
  const list = document.getElementById("task-history-list");
  if (!list) return;
  const history = storage.getTaskHistory();
  if (history.length === 0) {
    list.innerHTML = '<div class="compact-list-item">暂无任务历史</div>';
    return;
  }

  const statusText = {
    success: "成功",
    failed: "失败",
    aborted: "中止",
  };

  list.innerHTML = history.map((item) => {
    const date = new Date(item.createdAt || Date.now()).toLocaleString();
    const seconds = ((item.durationMs || 0) / 1000).toFixed(1);
    const privacyLabel = item.privacy ? "隐私" : "普通";
    return `
      <div class="compact-list-item">
        <div style="flex:1; min-width:0;">
          <div style="font-weight:600; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(item.actionName || "任务")} · ${escapeHtml(statusText[item.status] || item.status)}</div>
          <div style="font-size:10px; color:var(--text-secondary); overflow:hidden; text-overflow:ellipsis;">${escapeHtml(date)} · ${escapeHtml(item.model || "默认模型")} · ${item.segmentCount || 0}段 · ${seconds}s · ${privacyLabel}</div>
        </div>
      </div>
    `;
  }).join("");
}

// ==================== 提示词管理 ====================
function renderPromptList() {
  const list = document.getElementById("prompt-list");
  const prompts = storage.getPrompts();

  list.innerHTML = prompts
    .map(
      (p) => `
    <div class="prompt-item" data-id="${escapeHtml(p.id)}">
      <div class="prompt-item-color" style="background: ${escapeHtml(p.color)}"></div>
      <span class="prompt-item-icon">${escapeHtml(p.icon)}</span>
      <span class="prompt-item-name">${escapeHtml(p.name)}</span>
      <div class="prompt-item-actions">
        <button class="btn btn-sm btn-ghost edit-prompt-btn" data-id="${escapeHtml(p.id)}">✎</button>
        <button class="btn btn-sm btn-ghost delete-prompt-btn" data-id="${escapeHtml(p.id)}">✕</button>
      </div>
    </div>`
    )
    .join("");

  list.querySelectorAll(".edit-prompt-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const p = prompts.find((x) => x.id === btn.dataset.id);
      if (p) openPromptModal(p);
    });
  });
  list.querySelectorAll(".delete-prompt-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      storage.setPrompts(prompts.filter((x) => x.id !== btn.dataset.id));
      renderPromptList();
      renderActionButtons();
    });
  });
}

function openPromptModal(prompt) {
  const modal = document.getElementById("prompt-modal");
  document.getElementById("modal-title").textContent = prompt ? "编辑提示词" : "添加提示词";
  document.getElementById("prompt-name").value = prompt?.name || "";
  document.getElementById("prompt-icon").value = prompt?.icon || "💡";
  document.getElementById("prompt-color").value = prompt?.color || "#6366f1";
  document.getElementById("prompt-text").value = prompt?.prompt || "";
  editingPromptId = prompt?.id || null;
  modal.classList.remove("hidden");
}

function closePromptModal() {
  document.getElementById("prompt-modal").classList.add("hidden");
  editingPromptId = null;
}

function bindModalEvents() {
  document.getElementById("save-prompt-btn").addEventListener("click", () => {
    const name = document.getElementById("prompt-name").value.trim();
    const icon = document.getElementById("prompt-icon").value.trim() || "💡";
    const color = document.getElementById("prompt-color").value;
    const promptText = document.getElementById("prompt-text").value.trim();

    if (!name || !promptText) {
      showToast("名称和内容不能为空", "warning");
      return;
    }

    const prompts = storage.getPrompts();
    if (editingPromptId) {
      const idx = prompts.findIndex((p) => p.id === editingPromptId);
      if (idx >= 0) prompts[idx] = { ...prompts[idx], name, icon, color, prompt: promptText };
    } else {
      prompts.push({ id: "c_" + Date.now(), name, icon, color, prompt: promptText });
    }

    storage.setPrompts(prompts);
    renderPromptList();
    renderActionButtons();
    closePromptModal();
    showToast("已保存", "success");
  });

  document.getElementById("cancel-prompt-btn").addEventListener("click", closePromptModal);
  document.querySelector("#prompt-modal .modal-overlay").addEventListener("click", closePromptModal);
}

// ==================== Toast ====================
function showToast(message, type = "info") {
  document.querySelectorAll(".toast").forEach((t) => t.remove());
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  const colors = { success: "#10b981", error: "#ef4444", warning: "#f59e0b", info: "#6366f1" };
  Object.assign(toast.style, {
    position: "fixed",
    bottom: "12px",
    left: "50%",
    transform: "translateX(-50%) translateY(20px)",
    padding: "6px 14px",
    borderRadius: "6px",
    fontSize: "11.5px",
    fontWeight: "500",
    color: "white",
    background: colors[type] || colors.info,
    zIndex: "200",
    boxShadow: "var(--toast-shadow)",
    whiteSpace: "nowrap",
    opacity: "0",
    transition: "all 0.25s ease",
  });
  document.body.appendChild(toast);
  // 触发动画（不再依赖不存在的 slideDown keyframe）
  requestAnimationFrame(() => {
    toast.style.opacity = "1";
    toast.style.transform = "translateX(-50%) translateY(0)";
  });
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(-50%) translateY(20px)";
    setTimeout(() => toast.remove(), 300);
  }, 2500);
}

// ==================== 学术工具事件 ====================
function bindAcademicEvents() {
  // --- 表格样式 ---
  const styleSelect = document.getElementById("select-table-style");
  const styleInputWrap = document.getElementById("custom-style-wrap");
  const styleInput = document.getElementById("input-custom-style");

  // 下拉加载失败/不可用时，切换到手动输入样式名（兼容自定义样式）
  const fallbackToInput = (msg) => {
    if (msg) console.warn("表格样式下拉加载失败，已切换为手动输入:", msg);
    if (styleSelect) styleSelect.classList.add("hidden");
    if (styleInputWrap) styleInputWrap.classList.remove("hidden");
  };

  // 恢复上次保存的样式名到输入框
  if (styleInput) {
    styleInput.value = storage.getCustomTableStyle();
  }

  // 自动保存用户输入/选择的样式名
  const saveStyleName = (name) => {
    storage.setCustomTableStyle(name);
  };
  styleInput?.addEventListener("change", (e) => saveStyleName(e.target.value));
  styleSelect?.addEventListener("change", (e) => saveStyleName(e.target.value));

  // 异步加载文档中所有表格样式（含用户自定义样式）填充下拉
  if (styleSelect) {
    styleSelect.innerHTML = '<option value="">-- 加载中... --</option>';
    tableUtils
      .getTableStyles()
      .then((names) => {
        if (names.length === 0) {
          styleSelect.innerHTML = '<option value="">-- 未发现表格样式 --</option>';
          fallbackToInput("文档中未发现表格样式");
          return;
        }
        // 优先把含"三线"的自定义样式排前面，其余按名称排序
        const sorted = [...names].sort((a, b) => {
          const aThree = a.includes("三线");
          const bThree = b.includes("三线");
          if (aThree !== bThree) return aThree ? -1 : 1;
          return a.localeCompare(b, "zh");
        });
        styleSelect.innerHTML = sorted
          .map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`)
          .join("");
      })
      .catch((err) => {
        styleSelect.innerHTML = '<option value="">-- 加载失败 --</option>';
        fallbackToInput(err && err.message);
      });
  }

  // 读取当前生效的样式名（下拉优先，其次手动输入）
  const readStyleName = () => {
    if (styleSelect && !styleSelect.classList.contains("hidden")) {
      return styleSelect.value || "";
    }
    return styleInput ? styleInput.value.trim() : "";
  };

  // 读取当前宽度模式
  const readWidthMode = () => document.getElementById("select-width-mode")?.value || "";

  document.getElementById("btn-apply-table-style")?.addEventListener("click", async () => {
    const styleName = readStyleName();
    if (!styleName) {
      showToast("请先选择或输入表格样式名", "warning");
      return;
    }
    saveStyleName(styleName);
    try {
      showInlineStatus("processing", "正在应用表格样式...");
      const count = await tableUtils.optimizeSelection({ styleName, widthMode: readWidthMode() });
      showInlineStatus("done", `已应用「${styleName}」到 ${count} 个表格 ✓`);
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 5000);
    }
  });

  document.getElementById("btn-apply-all-tables")?.addEventListener("click", async () => {
    const styleName = readStyleName();
    if (!styleName) {
      showToast("请先选择或输入表格样式名", "warning");
      return;
    }
    saveStyleName(styleName);
    try {
      showInlineStatus("processing", `正在为全部表格应用「${styleName}」...`);
      const count = await tableUtils.applyStyleToAllTables({ styleName, widthMode: readWidthMode() });
      showInlineStatus("done", `已为全文 ${count} 个表格应用「${styleName}」✓`);
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 5000);
    }
  });

}
