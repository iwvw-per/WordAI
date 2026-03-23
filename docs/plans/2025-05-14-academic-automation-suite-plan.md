# WordAI 学术自动化套件实现计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 实现表格一键三线表美化与智能参考文献（占位符匹配与重格式化）功能。

**Architecture:** 
- `src/utils/table.js`: 提供表格结构识别与样式应用工具函数。
- `src/utils/references.js`: 提供参考文献扫描、正则匹配、编号更新与跳转逻辑。
- `src/taskpane/academic-tools.html/js`: 侧边栏独立模块 UI 与业务绑定。

**Tech Stack:** JavaScript (ES6), Office.js API, CSS.

---

### 任务 1: 创建表格处理工具类 (Table Utils)

**Files:**
- Create: `src/utils/table.js`

**步骤 1: 实现应用三线表逻辑**
包含顶线、底线和基于 `detectHeaderRows` 的表头线下划线。

```javascript
/**
 * @param {Word.Table} table 
 * @param {Object} config { topWidth, bottomWidth, headerWidth, padding }
 */
export async function applyAcademicStyle(table, config) {
  await Word.run(async (context) => {
    table.load(["rows/items/cells", "rows/items/verticalAlignment"]);
    // 清除边框并设置三线
    table.borders.outsideLineWidth = 0;
    table.borders.insideHorizontalLineWidth = 0;
    table.borders.insideVerticalLineWidth = 0;

    table.borders.top.style = "Single";
    table.borders.top.width = config.topWidth || 1.5;
    
    table.borders.bottom.style = "Single";
    table.borders.bottom.width = config.bottomWidth || 1.5;

    // TODO: 识别表头行并加线
    await context.sync();
  });
}
```

**步骤 2: 提交**
```bash
git add src/utils/table.js
git commit -m "feat: 增加学术表格美化底层逻辑"
```

---

### 任务 2: 创建参考文献处理工具类 (Ref Utils)

**Files:**
- Create: `src/utils/references.js`

**步骤 1: 实现占位符扫描正则**
匹配 `【作者, 年份】`。

```javascript
/**
 * 全文扫描占位符
 */
export async function scanPlaceholders() {
  return await Word.run(async (context) => {
    const results = context.document.body.search("【*】", { matchWildcards: true });
    results.load("items");
    await context.sync();
    return results.items;
  });
}
```

**步骤 2: 实现参考文献列表解析**
解析文档末尾的 References 章节。

**步骤 3: 提交**
```bash
git add src/utils/references.js
git commit -m "feat: 增加参考文献扫描与解析逻辑"
```

---

### 任务 3: UI 界面开发 (Academic Panel)

**Files:**
- Modify: `src/taskpane/taskpane.html`
- Modify: `src/taskpane/taskpane.js`

**步骤 1: 增加“学术工具”侧边栏面板**
实现 Tab 切换和功能入口。

**步骤 2: 提交**
```bash
git commit -m "feat: 集成学术工具侧边栏 UI"
```

---

### 任务 4: 实现分步导航确认功能

**Files:**
- Modify: `src/taskpane/taskpane.js`

**步骤 1: 编写“查找下一个占位符”逻辑**
利用 `taskpane.js` 管理当前扫描索引，点击按钮时 `range.select()` 跳转。

**步骤 2: 提交**
```bash
git commit -m "feat: 实现参考文献引用分步导航功能"
```

---

### 任务 5: 术语一致性检查逻辑 (Terminology Utils)

**Files:**
- Create: `src/utils/terminology.js`

**步骤 1: 实现术语提取与聚类**
利用 LLM 接口提取关键字，并在本地进行简单的语义聚类建议。

```javascript
/**
 * 扫描全文提取术语并寻找潜在冲突
 */
export async function scanTerminologyConflicts() {
  // 1. 获取全文文本
  // 2. 分段发送给 LLM 提取核心术语
  // 3. AI 返回冲突报告：{ standard: "卷积神经网络", aliases: ["卷积感知机", "CNN"] }
}
```

**步骤 2: 实现一键替换逻辑**
调用 Word 的全局替换 API。

**步骤 2: 提交**
```bash
git commit -m "feat: 增加术语一致性检查模块"
```

---

### 任务 6: 图表编号同步逻辑 (Numbering Utils)

**Files:**
- Create: `src/utils/numbering.js`

**步骤 1: 实现图表标题搜索与重编**
利用 `search` API 寻找 "图 * " 和 "表 * "。

```javascript
/**
 * 重新编排全文图表编号
 */
export async function renumberFiguresAndTables() {
  // 1. 扫描全文图表标题
  // 2. 按顺序重新编号
  // 3. 寻找正文中对应的引用并同步更新
}
```

**步骤 2: 提交**
```bash
git add src/utils/numbering.js
git commit -m "feat: 增加图表编号同步逻辑"
```

---

### 任务 7: 摘要与关键词生成逻辑 (Abstract Utils)

**Files:**
- Create: `src/utils/abstract.js`

**步骤 1: 实现长文档分段摘要**
调用 LLM 接口处理全文核心段落。

**步骤 2: 提交**
```bash
git add src/utils/abstract.js
git commit -m "feat: 增加摘要与关键词生成功能"
```

