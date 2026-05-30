/**
 * taskpane.js - WordAI 主逻辑
 * 自动执行模式：选中 → 点击 → 自动替换
 */

import "./taskpane.css";
import * as storage from "../utils/storage.js";
import * as llm from "../utils/llm.js";
import * as ooxml from "../utils/ooxml.js";
import * as format from "../utils/format.js";
import * as tableUtils from "../utils/table.js";
import * as refUtils from "../utils/references.js";
import * as termUtils from "../utils/terminology.js";
import * as numberUtils from "../utils/numbering.js";
import * as abstractUtils from "../utils/abstract.js";

// ==================== 工具函数 ====================
function escapeHtml(str) {
  if (!str) return "";
  const s = String(str);
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

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

let refNavState = {
  items: [],
  currentIndex: 0
};

// ==================== 初始化 ====================
Office.onReady((info) => {
  isOfficeReady = true;
  initApp();
});

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

  // 监听系统主题变化（仅在用户未手动选择主题时跟随系统）
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
    // 仅在用户未手动设置主题时才跟随系统变化
    const userTheme = localStorage.getItem("wordai_theme");
    if (!userTheme) {
      applyAutoTheme();
    }
  });

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
      body.innerHTML = '<div class="console-line system">⚡ 流水线控制台已清空。</div>';
    }
    if (title) {
      title.textContent = "空闲";
    }
  });
}

// ==================== 核心：一键执行 ====================
async function executeAction(systemPrompt, actionName, triggerBtn) {
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

  // 触发按钮触感反馈（非灰变禁用）
  setAllActionsLoading(true, triggerBtn);

  let segments = [];
  try {
    // ⚡ 瞬时锁定：瞬间在 Word 中锁定并锚定该任务对应选区（约耗时 50ms）
    segments = await ooxml.markSelection();
    if (!segments || segments.length === 0) {
      throw new Error("请先在 Word 中选中需要处理的文字内容");
    }
  } catch (err) {
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
    status: "pending"
  };

  globalTaskQueue.push(pipelineTask);

  // 在滚动控制台动态注入精简的单行任务项
  if (consoleBody) {
    const placeholder = consoleBody.querySelector(".console-line.system");
    if (placeholder && placeholder.textContent.includes("流水线已准备就绪")) {
      placeholder.remove();
    }

    const taskLine = document.createElement("div");
    taskLine.id = `console-task-${taskId}`;
    taskLine.className = "console-line pending";
    taskLine.innerHTML = `
      <span class="task-badge">#${taskCounter}</span>
      <span class="task-name" title="${escapeHtml(actionName)}">${escapeHtml(actionName)}</span>
      <span class="task-progress" id="console-progress-${taskId}">⏳ 排队中</span>
    `;

    consoleBody.appendChild(taskLine);
    consoleBody.scrollTop = consoleBody.scrollHeight;
  }

  showToast(`任务 #${taskCounter} [${actionName}] 已锁屏排队 ✓`, "success");

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

  if (taskLine) {
    taskLine.className = "console-line processing";
  }
  updateConsoleTitleState();

  currentTask.abortController = new AbortController();
  const signal = currentTask.abortController.signal;
  const seg = currentTask.segments[0]; // 整块化改写，必然只有 1 段

  let retryCount = 0;
  const retryLimit = 3;
  let success = false;

  while (retryCount < retryLimit) {
    if (signal.aborted) break;

    try {
      if (progressSpan) {
        if (retryCount > 0) {
          progressSpan.textContent = `⏳ 重试 ${retryCount}/3...`;
        } else {
          progressSpan.textContent = `⚡ 请求云端中...`;
        }
        if (consoleBody) consoleBody.scrollTop = consoleBody.scrollHeight;
      }

      // 拼装红线限制提示词
      let hasShields = seg.refMap && seg.refMap.length > 0;
      let redLine = "";
      if (hasShields) {
        redLine += "\n\n【绝对禁令】：文中的 [REF_N], [EQN_N], [FNOTE_N] 是物理引用或公式锚点，你必须原封不动地保留所有此类标记（包括内部的类型、编号以及外层的英文中括号 []），必须将其放置在改写后对应的语义位置。严禁删除、修改括号类型（不能改为 【】 或 『』等）！";
      }
      // 强力注入段落保留约束，保障段落不被大模型合并
      redLine += "\n\n【段落结构绝对保留】：输入的文本可能包含多个自然段（段落之间通过换行符 \\n 分隔）。你必须原封不动地保持这些段落结构。改写后的文本必须拥有与输入文本完全一致的段落数量和换行分隔，严禁将多段合并为一段，也严禁自行切分或增加额外的段落！请严格输出带换行符 \\n 的多段内容。";

      const finalPrompt = redLine ? (currentTask.systemPrompt + redLine + "\n") : currentTask.systemPrompt;
      const normalizedInputText = seg.text.replace(/\r/g, "\n");

      // 触发流式输出并在控制终端行渲染 delta
      const raw = await llm.callLLMStream(finalPrompt, normalizedInputText, (delta, currentText) => {
        if (progressSpan && !signal.aborted) {
          let displaySnippet = currentText.replace(/<\/?p[^>]*>|\[PARAGRAPH_\d+\]|\n/gi, "");
          if (displaySnippet.length > 15) displaySnippet = "..." + displaySnippet.slice(-15);
          progressSpan.textContent = `⚡ 改写中: "${displaySnippet}█"`;
        }
      }, signal);

      if (signal.aborted) throw new Error("已取消");

      if (progressSpan) {
        progressSpan.textContent = `🧩 恢复排版中...`;
      }

      // 调用单点 AST 还原回填
      const cleanRaw = llm.cleanAiResponse(raw);
      await ooxml.replaceSingleMarkedContent(cleanRaw, seg.refMap, seg.boundaryTags);

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
        // 彻底失败清理边界标记
        try {
          await Word.run(async (context) => {
            const startCCs = context.document.contentControls.getByTag(seg.boundaryTags.start);
            const endCCs = context.document.contentControls.getByTag(seg.boundaryTags.end);
            startCCs.load("items");
            endCCs.load("items");
            await context.sync();
            if (startCCs.items.length > 0) startCCs.items[0].delete(true);
            if (endCCs.items.length > 0) endCCs.items[0].delete(true);
            await context.sync();
          });
        } catch {}
        break;
      }
    }
  }

  // 更新 Task 行的完成或中止状态
  if (taskLine) {
    if (signal.aborted) {
      taskLine.className = "console-line error";
      if (progressSpan) progressSpan.textContent = `🛑 已中止`;
    } else if (success) {
      taskLine.className = "console-line done";
      if (progressSpan) progressSpan.textContent = `✅ 成功`;
    } else {
      taskLine.className = "console-line error";
      if (progressSpan) progressSpan.textContent = `❌ 失败`;
    }
    if (consoleBody) consoleBody.scrollTop = consoleBody.scrollHeight;
  }

  // 释放并发任务计数，并重新触发队列调度
  activeTaskCount--;
  updateConsoleTitleState();
  processTaskQueue();
}

/**
 * 动态刷新控制台总标题的线程执行状态
 */
function updateConsoleTitleState() {
  const consoleTitle = document.getElementById("pipeline-console-title");
  if (!consoleTitle) return;
  if (activeTaskCount > 0) {
    consoleTitle.textContent = `正在执行 (${activeTaskCount} 个任务)...`;
  } else {
    consoleTitle.textContent = globalTaskQueue.length > 0 ? "等待调度中" : "队列已清空";
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
      <span class="status-text">${message}</span>
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

  const concurrency = storage.getConcurrencyLimit();
  const slider = document.getElementById("concurrency-slider");
  const valueBadge = document.getElementById("concurrency-value");
  if (slider && valueBadge) {
    slider.value = concurrency;
    valueBadge.textContent = concurrency;
  }
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
    boxShadow: "0 4px 12px rgba(0,0,0,0.3)",
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
  // --- 表格工具 ---
  document.getElementById("btn-apply-3line")?.addEventListener("click", async () => {
    try {
      showInlineStatus("processing", "正在美化表格...");
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const tables = selection.tables;
        tables.load("items");
        await context.sync();

        if (tables.items.length === 0) throw new Error("请先选中包含表格的区域");

        const config = {
          topWidth: parseFloat(document.getElementById("input-line-bold").value) || 1.5,
          bottomWidth: parseFloat(document.getElementById("input-line-bold").value) || 1.5,
          headerWidth: 0.75
        };

        for (let table of tables.items) {
          await tableUtils.applyAcademicStyle(table, config);
        }
        showInlineStatus("done", "表格已美化 ✓");
      });
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 5000);
    }
  });

  document.getElementById("btn-scan-tables")?.addEventListener("click", async () => {
    try {
      const resultsDiv = document.getElementById("table-scan-results");
      resultsDiv.classList.toggle("hidden");
      if (resultsDiv.classList.contains("hidden")) return;

      resultsDiv.innerHTML = '<div class="compact-list-item">扫描中...</div>';

      const tables = await tableUtils.getAllTablesInfo();
      if (tables.length === 0) {
        resultsDiv.innerHTML = '<div class="compact-list-item">未发现表格</div>';
      } else {
        resultsDiv.innerHTML = tables.map(t => `
          <div class="compact-list-item" data-id="${t.id}">
            <span>表格 ${t.id + 1} (${t.rowCount}x${t.columnCount})</span>
          </div>
        `).join("");
      }
    } catch (err) {
      showToast(err.message, "error");
    }
  });

  // --- 参考文献 ---
  document.getElementById("btn-match-refs")?.addEventListener("click", async () => {
    try {
      showInlineStatus("processing", "正在扫描占位符与文献列表...");

      // scanPlaceholders 现在返回纯数据
      const [placeholderData, bibliography] = await Promise.all([
        refUtils.scanPlaceholders(),
        refUtils.parseBibliography()
      ]);

      if (placeholderData.count === 0) {
        showInlineStatus("done", "未发现占位符");
        setTimeout(() => hideInlineStatus(), 2000);
        return;
      }

      // 重新在 Word.run 内创建 ContentControl 标记
      await Word.run(async (context) => {
        // 清理旧标记
        const oldCCs = context.document.contentControls.getByTag("wordai_ref_placeholder");
        oldCCs.load("items");
        await context.sync();
        for (let cc of oldCCs.items) cc.delete(true);
        await context.sync();

        // 重新搜索并包裹 CC
        const results = context.document.body.search("【*】", { matchWildcards: true });
        results.load("items");
        await context.sync();

        for (let i = 0; i < results.items.length; i++) {
          const cc = results.items[i].insertContentControl();
          cc.tag = "wordai_ref_placeholder";
          cc.appearance = "Hidden";
        }
        await context.sync();
      });

      refNavState.itemsCount = placeholderData.count;
      refNavState.bibliography = bibliography;
      refNavState.currentIndex = 0;

      document.getElementById("ref-navigator").classList.remove("hidden");
      await updateRefNavigator();
      showInlineStatus("done", `发现 ${placeholderData.count} 处引用，文献库 ${bibliography.length} 条 ✓`);
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 5000);
    }
  });

  document.getElementById("btn-ref-prev")?.addEventListener("click", () => {
    if (refNavState.currentIndex > 0) {
      refNavState.currentIndex--;
      updateRefNavigator();
    }
  });

  document.getElementById("btn-ref-next")?.addEventListener("click", () => {
    if (refNavState.currentIndex < refNavState.items.length - 1) {
      refNavState.currentIndex++;
      updateRefNavigator();
    }
  });

  document.getElementById("btn-ref-confirm")?.addEventListener("click", handleRefConfirm);

  // 一键刷新编号与链接
  document.getElementById("btn-refresh-refs")?.addEventListener("click", async () => {
    try {
      showInlineStatus("processing", "正在刷新引用编号与链接...");
      await Word.run(async (context) => {
        const body = context.document.body;
        // 搜索所有 [N] 格式的引用
        const matches = body.search("\\[[0-9\\- ,]@\\]", { matchWildcards: true });
        matches.load("items");
        await context.sync();
        for (const m of matches.items) m.load("text");
        await context.sync();

        let linkedCount = 0;
        for (const m of matches.items) {
          const numMatch = m.text.match(/\d+/);
          if (numMatch) {
            m.hyperlink = `#wordai_ref_${numMatch[0]}`;
            m.font.color = "black";
            m.font.underline = "None";
            linkedCount++;
          }
        }
        await context.sync();
        showInlineStatus("done", `已刷新 ${linkedCount} 处引用链接 ✓`);
      });
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 5000);
    }
  });

  // --- 写作辅助 ---
  document.getElementById("btn-term-check")?.addEventListener("click", async () => {
    try {
      const resultsDiv = document.getElementById("term-check-results");
      resultsDiv.classList.remove("hidden");
      resultsDiv.innerHTML = '<div class="compact-list-item">扫描并识别术语中...</div>';

      let extractedText = "";

      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        extractedText = body.text.substring(0, 15000); // 截取核心内容
      });

      // 【关键修复】：将耗时巨大的大模型网络请求彻底剥离出随时可能 Timeout 的 Word.run 隔离区
      const conflicts = await termUtils.extractTerminology(extractedText);

      if (!conflicts || conflicts.length === 0) {
        resultsDiv.innerHTML = '<div class="compact-list-item">未发现明显术_语冲突 ✨</div>';
        setTimeout(() => resultsDiv.classList.add("hidden"), 3000);
        return;
      }

      resultsDiv.innerHTML = conflicts.map((c, i) => `
          <div class="compact-list-item term-conflict-item">
            <div style="flex:1">
              <span class="badge" style="background:var(--primary-light)">${c.standard}</span>
              <span style="font-size:10px; color:var(--text-secondary)"> ← ${c.aliases.join(", ")}</span>
            </div>
            <button class="btn btn-xs btn-ghost unify-term-btn" data-standard="${c.standard}" data-aliases='${JSON.stringify(c.aliases)}'>统一</button>
          </div>
        `).join("");

      // 绑定统一事件
      resultsDiv.querySelectorAll(".unify-term-btn").forEach(btn => {
        btn.addEventListener("click", async () => {
          const standard = btn.dataset.standard;
          const aliases = JSON.parse(btn.dataset.aliases);
          btn.disabled = true;
          btn.textContent = "⏳";
          try {
            await termUtils.replaceTerminology(aliases, standard);
            showToast(`全文 ${standard} 已统一 ✓`, "success");
            btn.closest(".term-conflict-item").style.opacity = "0.5";
            btn.textContent = "已统一";
          } catch (err) {
            showToast(err.message, "error");
            btn.disabled = false;
            btn.textContent = "统一";
          }
        });
      });
    } catch (err) {
      showToast(err.message, "error");
      document.getElementById("term-check-results").classList.add("hidden");
    }
  });

  document.getElementById("btn-renumber")?.addEventListener("click", async () => {
    try {
      showInlineStatus("processing", "正在重排图表编号...");
      const result = await numberUtils.renumberFiguresAndTables();
      showToast(result.message, "success");
      hideInlineStatus();
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 3000);
    }
  });

  document.getElementById("btn-gen-abstract")?.addEventListener("click", async () => {
    try {
      showInlineStatus("processing", "正在提炼全文生成摘要...");
      const abstract = await abstractUtils.generateAbstract();

      // 将摘要插入到文档开头，使用富文本渲染
      await Word.run(async (context) => {
        const body = context.document.body;
        const headerRange = body.insertParagraph("【AI 生成摘要与关键词】", "Start");
        headerRange.font.bold = true;
        headerRange.font.size = 14;
        await context.sync();

        await format.insertMarkdownAsRichText(headerRange, abstract, "After");
      });

      showToast("摘要已生成并插入文首", "success");
      hideInlineStatus();
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 3000);
    }
  });
}

async function updateRefNavigator() {
  const status = document.getElementById("ref-nav-status");
  const target = document.getElementById("ref-nav-target");

  await Word.run(async (context) => {
    const ccs = context.document.contentControls.getByTag("wordai_ref_placeholder");
    ccs.load("items");
    await context.sync();

    if (ccs.items.length === 0 || refNavState.currentIndex >= ccs.items.length) {
      document.getElementById("ref-navigator").classList.add("hidden");
      return;
    }

    const currentCC = ccs.items[refNavState.currentIndex];
    currentCC.load("text");
    currentCC.select();
    await context.sync();

    const text = currentCC.text;
    const suggestions = refUtils.matchPlaceholderToBibliography(text, refNavState.bibliography || []);
    refNavState.suggestions = suggestions;

    status.textContent = `第 ${refNavState.currentIndex + 1}/${ccs.items.length} 处: ${text}`;

    if (suggestions.length > 0) {
      target.value = `建议匹配: [${suggestions[0].id}] ${suggestions[0].text.substring(0, 50)}...`;
      target.style.color = "var(--primary)";
    } else {
      target.value = "未找到匹配项";
      target.style.color = "var(--error)";
    }

    // ⚡ 动态诊断与打分渲染
    const detailsContent = document.getElementById("ref-match-details-content");
    if (detailsContent) {
      const clean = text.replace(/[【】\[\]]/g, "");
      const authorMatch = clean.match(/[\u4e00-\u9fff]{2,4}|[a-zA-Z\-]{2,}/);
      const yearMatch = clean.match(/\b(19|20)\d{2}\b/);
      const pAuthor = authorMatch ? authorMatch[0].toLowerCase() : null;
      const pYear = yearMatch ? yearMatch[0] : null;

      let html = `<div style="margin-bottom:4px; font-weight:bold; color:var(--primary);">正文提取：作者="${pAuthor || '无'}" 年份="${pYear || '无'}"</div>`;
      
      if (!refNavState.bibliography || refNavState.bibliography.length === 0) {
        html += `<div style="color:var(--error); font-weight:bold;">⚠️ 侧边栏未检索到文末参考文献！请先确认文档末尾有以“参考文献”或“References”命名的标题，且下方包含完整的文献列表。</div>`;
      } else {
        html += refNavState.bibliography.map(entry => {
          let score = 0;
          let matchLog = [];
          if (pAuthor && entry.coreAuthor === pAuthor) {
            score += 60;
            matchLog.push(`核心作者对齐(+60)`);
          } else if (pAuthor && entry.text.toLowerCase().includes(pAuthor)) {
            score += 30;
            matchLog.push(`包含核心作者(+30)`);
          }
          if (pYear && entry.year === pYear) {
            score += 40;
            matchLog.push(`年份相同(+40)`);
          }
          const logStr = score > 0 ? ` [${matchLog.join(',')}]` : ' [无匹配点]';
          return `<div style="margin-bottom:4px; border-bottom:1px dashed rgba(0,0,0,0.05); padding-bottom:2px; ${score > 0 ? 'color:#10b981; font-weight:500;' : ''}">
            [${entry.id}] 得分: ${score}${logStr}<br/>
            <span style="font-size:9px; opacity:0.8; color:var(--text-secondary);">文献: ${entry.text.substring(0, 45)}...</span>
          </div>`;
        }).join("");
      }
      detailsContent.innerHTML = html;
    }
  });
}

// 确认按钮逻辑（需要在外面绑定，或在此处根据需要修改按钮监听器）
// 之前是在 init 中绑定的，这里我补充一下针对 CC 的逻辑修改
async function handleRefConfirm() {
  try {
    const suggestions = refNavState.suggestions || [];
    const bestMatch = suggestions[0];

    if (!bestMatch) {
      showToast("未找到匹配的参考文献，请手动处理", "warning");
      return;
    }

    await Word.run(async (context) => {
      const ccs = context.document.contentControls.getByTag("wordai_ref_placeholder");
      ccs.load("items");
      await context.sync();

      if (ccs.items.length > refNavState.currentIndex) {
        const currentCC = ccs.items[refNavState.currentIndex];
        const replacement = `[${bestMatch.id}]`;
        const run = currentCC.insertText(replacement, "Replace");
        run.font.color = "#2563eb";
        currentCC.delete(false); // 仅删除容器，保留文字
        await context.sync();

        showToast(`已匹配到: ${bestMatch.text.substring(0, 20)}...`, "success");

        if (refNavState.currentIndex < ccs.items.length - 1) {
          // 下一个（currentIndex 不变，因为删了一个 CC 后后面的索引会顶上来）
          await updateRefNavigator();
        } else {
          document.getElementById("ref-navigator").classList.add("hidden");
          showToast("全部匹配完成", "success");
        }
      }
    });
  } catch (err) {
    showToast(err.message, "error");
  }
}
