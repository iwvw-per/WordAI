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

// ==================== 状态 ====================
let appInitialized = false;
let isOfficeReady = false;
let isProcessing = false;
let editingPromptId = null;
let currentAbortController = null;

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

  // 监听系统主题变化（系统一旦变化，强制解除用户的手动锁定恢复自动）
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
    localStorage.removeItem("wordai_theme");
    const tbtn = document.getElementById("theme-toggle-btn");
    if (tbtn) tbtn.textContent = "🌓";
    applyAutoTheme();
  });

  // 监听 Office 主题变化
  if (isOfficeReady && Office.context?.document?.addHandlerAsync) {
    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.OfficeThemeChanged,
        () => {
          localStorage.removeItem("wordai_theme");
          const tbtn = document.getElementById("theme-toggle-btn");
          if (tbtn) tbtn.textContent = "🌓";
          applyAutoTheme();
        }
      );
    } catch (e) {}
  }
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
  } catch {}

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
    <button class="action-btn" data-prompt-id="${p.id}" style="--btn-color: ${p.color}" title="${p.name}">
      <span class="action-icon">${p.icon}</span>
      <span class="action-name">${p.name}</span>
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

}


// ==================== 核心：一键执行 ====================
async function executeAction(systemPrompt, actionName, triggerBtn) {
  if (isProcessing) return;

  if (!storage.isConfigured()) {
    showConfigBanner();
    showToast("请先完成 API 配置", "warning");
    return;
  }

  if (!isOfficeReady) {
    showToast("请在 Word 中使用", "error");
    return;
  }

  isProcessing = true;
  currentAbortController = new AbortController();
  setAllActionsLoading(true, triggerBtn);
  showInlineStatus("processing", `${actionName}中...`, true);

  try {
    const result = await ooxml.executeAndReplace(
      async (text) => {
        // 动态注入保护指令，确保 AI 不会篡改占位符
        const protectionNotice = "\n\n注意：文中的 {{REF_N}} 是受保护的引用占位符，请务必原封不动保留并放置在正确的语义位置。";
        let result = await llm.callLLM(systemPrompt + protectionNotice, text, currentAbortController.signal);
        
        // 清理常见的 AI 回复套话前缀
        result = result.replace(/^([\s\n]*(\*\*|__)?(这里是|这是)?(修改后|润色后|翻译后|重写后|扩展后|缩写后|处理后)的?(内容|文本|结果|段落)?[\s\n]*(\*\*|__)?[\s\n]*[：:\n]+)/i, "").trim();
        // 清理可能包裹的 markdown 代码块
        result = result.replace(/^```[a-zA-Z]*\n/i, "").replace(/\n```$/i, "").trim();
        return result;
      },
      (type, msg, canCancel) => showInlineStatus(type, msg, canCancel),
      currentAbortController.signal
    );

    showInlineStatus("done", `${actionName}完成 ✓`);
    setTimeout(() => hideInlineStatus(), 2000);
  } catch (err) {
    if (err.name === "AbortError" || err.message === "已取消") {
      showInlineStatus("error", "用户已中止");
    } else {
      showInlineStatus("error", err.message);
    }
    setTimeout(() => hideInlineStatus(), 3000);
    // 确保清理
    await ooxml.clearMarks();
  } finally {
    isProcessing = false;
    currentAbortController = null;
    setAllActionsLoading(false);
  }
}

function setAllActionsLoading(loading, activeBtn) {
  document.querySelectorAll(".action-btn").forEach((btn) => {
    btn.disabled = loading;
    btn.classList.toggle("loading", loading && btn !== activeBtn);
    if (!loading) {
      btn.classList.remove("active-loading");
    }
  });

  if (loading && activeBtn) {
    activeBtn.classList.add("active-loading");
  }

  const customBtn = document.getElementById("custom-run-btn");
  if (customBtn) customBtn.disabled = loading;
}

// ==================== 内联状态条 ====================
function showInlineStatus(type, message, canCancel = false) {
  let bar = document.getElementById("inline-status");
  if (!bar) {
    bar = document.createElement("div");
    bar.id = "inline-status";
    const grid = document.getElementById("action-grid");
    grid.parentNode.insertBefore(bar, grid.nextSibling);
  }

  bar.className = `inline-status ${type}`;

  let html = "";
  if (type === "processing") {
    html = `<span class="inline-spinner"></span><span style="flex:1">${message}</span>`;
    if (canCancel) {
      html += `<button class="btn btn-sm btn-ghost" id="inline-cancel-btn" style="padding: 2px 6px; font-size: 10px;">中断</button>`;
    }
  } else if (type === "done") {
    html = `<span class="inline-icon done">✓</span><span>${message}</span>`;
  } else if (type === "error") {
    html = `<span class="inline-icon error">✕</span><span>${message}</span>`;
  }
  
  bar.innerHTML = html;

  // 绑定取消事件
  const cancelBtn = document.getElementById("inline-cancel-btn");
  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => {
      if (currentAbortController) {
        currentAbortController.abort(new Error("已取消"));
      }
    });
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
    <div class="prompt-item" data-id="${p.id}">
      <div class="prompt-item-color" style="background: ${p.color}"></div>
      <span class="prompt-item-icon">${p.icon}</span>
      <span class="prompt-item-name">${p.name}</span>
      <div class="prompt-item-actions">
        <button class="btn btn-sm btn-ghost edit-prompt-btn" data-id="${p.id}">✎</button>
        <button class="btn btn-sm btn-ghost delete-prompt-btn" data-id="${p.id}">✕</button>
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
    transform: "translateX(-50%)",
    padding: "6px 14px",
    borderRadius: "6px",
    fontSize: "11.5px",
    fontWeight: "500",
    color: "white",
    background: colors[type] || colors.info,
    zIndex: "200",
    boxShadow: "0 4px 12px rgba(0,0,0,0.3)",
    whiteSpace: "nowrap",
    animation: "slideDown 0.2s ease",
  });
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transition = "opacity 0.3s";
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
    }
    setTimeout(() => hideInlineStatus(), 2000);
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
      
      await Word.run(async (context) => {
        // 1. 同时扫描占位符和解析列表
        const [results, bibliography] = await Promise.all([
          refUtils.scanPlaceholders(),
          refUtils.parseBibliography()
        ]);

        if (results.items.length === 0) {
          showInlineStatus("done", "未发现占位符");
          setTimeout(() => hideInlineStatus(), 2000);
          return;
        }

        // 2. 清理旧标记并为每个占位符包裹 CC 以便追踪
        const oldCCs = context.document.contentControls.getByTag("wordai_ref_placeholder");
        oldCCs.load("items");
        await context.sync();
        for (let cc of oldCCs.items) cc.delete(true);

        for (let i = 0; i < results.items.length; i++) {
          const cc = results.items[i].insertContentControl();
          cc.tag = "wordai_ref_placeholder";
          cc.appearance = "Hidden";
        }
        await context.sync();
        
        refNavState.itemsCount = results.items.length;
        refNavState.bibliography = bibliography;
        refNavState.currentIndex = 0;
        
        document.getElementById("ref-navigator").classList.remove("hidden");
        await updateRefNavigator();
        showInlineStatus("done", `发现 ${results.items.length} 处引用，文献库 ${bibliography.length} 条 ✓`);
        setTimeout(() => hideInlineStatus(), 2000);
      });
    } catch (err) {
      showInlineStatus("error", err.message);
      setTimeout(() => hideInlineStatus(), 2000);
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

  // --- 写作辅助 ---
  document.getElementById("btn-term-check")?.addEventListener("click", async () => {
    try {
      const resultsDiv = document.getElementById("term-check-results");
      resultsDiv.classList.remove("hidden");
      resultsDiv.innerHTML = '<div class="compact-list-item">扫描并识别术语中...</div>';

      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();
        const text = body.text.substring(0, 15000); // 截取核心内容
        const conflicts = await termUtils.extractTerminology(text);
        
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
