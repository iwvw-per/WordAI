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
      async (segments, signal) => {
        let hasRefs = segments.some(s => s.refMap && s.refMap.length > 0);
        let redLine = "";
        if (hasRefs) {
            redLine += "\n\n【绝对禁令】：文中的 [REF_N] 是物理引用锚点，你必须原封不动地保留所有 [REF_N]（包括里面的 REF、编号以及外层的英文中括号 []），必须将其放置在改写后对应的语义位置。严禁删除、修改括号类型（不能改为 【】 或 『』等）！";
        }

        let fullText = "";
        if (segments.length > 1) {
            redLine += "\n【多段落严格指令】：本次待处理的文本被切分成了多个片段，并分别使用 <p id=\"N\"> 标签包裹。你必须逐个对每个片段进行处理，并在处理后的内容外层原封不动地保留对应的 <p id=\"N\"> 和 </p> 标签包围！严禁合并段落！严禁遗漏标签！";
            fullText = segments.map((s, i) => `<p id="${i}">\n${s.text}\n</p>`).join("\n\n");
        } else {
            fullText = segments[0].text;
        }

        showInlineStatus("processing", "🔗 等待云端响应...", true);

        const finalPrompt = redLine ? (systemPrompt + redLine + "\n") : systemPrompt;
        const raw = await llm.callLLMStream(finalPrompt, fullText, (delta, currentText) => {
            let displaySnippet = currentText.replace(/<\/?p[^>]*>|\[PARAGRAPH_\d+\]|\n/gi, "");
            if (displaySnippet.length > 15) displaySnippet = "..." + displaySnippet.slice(-15);
            showInlineStatus("processing", `⚡ ${displaySnippet}█`, true);
        }, signal);
        const cleanRaw = llm.cleanAiResponse(raw);

        if (segments.length === 1) {
            return [cleanRaw.replace(/<\/?p[^>]*>/g, '').trim()];
        } else {
            const aiTexts = [];
            for (let i = 0; i < segments.length; i++) {
                // 兼容带不带双引号或者单引号的 XML 属性
                const regex = new RegExp(`<p\\s+id=["']?${i}["']?\\s*>([\\s\\S]*?)</p>`, 'i');
                const match = cleanRaw.match(regex);
                if (match) {
                    aiTexts.push(match[1].trim());
                } else {
                    aiTexts.push("");
                }
            }
            if (aiTexts.some(t => !t)) {
                throw new Error("大模型未能遵循多段落 XML 标签分割指令，请求作废。请调整文本长度后重试。");
            }
            return aiTexts;
        }
      },
      (type, msg, canCancel) => showInlineStatus(type, msg, canCancel),
      currentAbortController.signal
    );

    showInlineStatus("done", `${actionName}完成 ✓`);
    setTimeout(() => hideInlineStatus(), 10000); // 延长至 10 秒供用户查看结果
  } catch (err) {
    if (err.name === "AbortError" || err.message === "已取消") {
      showInlineStatus("error", "用户已中止");
    } else {
      showInlineStatus("error", err.message);
    }
    setTimeout(() => hideInlineStatus(), 10000); // 错误日志同样延长停留
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
      setTimeout(() => hideInlineStatus(), 2000);
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
      setTimeout(() => hideInlineStatus(), 10000);
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
      setTimeout(() => hideInlineStatus(), 10000);
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
