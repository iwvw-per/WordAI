# 表格美化（三线表）实现计划

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 在 WordAI 中增加一键美化表格为学术三线表的功能，支持智能识别复合表头、批量扫描和参数微调。

**Architecture:** 
- 在 `src/utils/table.js` 中实现表格结构识别与 OOXML 操作逻辑。
- 在 `src/taskpane/taskpane.html` 和 `taskpane.js` 中增加表格美化面板。
- 利用 Office.js API 直接操作 `Word.Table` 对象的边框和对齐属性。

**Tech Stack:** JavaScript (ES6), Office.js API, CSS.

---

### 任务 1: 创建表格处理工具类

**Files:**
- Create: `src/utils/table.js`

**步骤 1: 编写基础函数原型**
实现 `applyThreeLineStyle` 函数，接收表格对象和配置参数。

```javascript
/**
 * 为表格应用三线表样式
 * @param {Word.Table} table 
 * @param {Object} config { topWidth, bottomWidth, headerWidth }
 */
export async function applyThreeLineStyle(table, config = { topWidth: 1.5, bottomWidth: 1.5, headerWidth: 0.75 }) {
  await Word.run(async (context) => {
    table.load("rows/items/cells");
    // 清除原有边框
    table.borders.outsideLineWidth = 0;
    table.borders.insideHorizontalLineWidth = 0;
    table.borders.insideVerticalLineWidth = 0;

    // 设置顶线
    table.borders.top.style = "Single";
    table.borders.top.width = config.topWidth;
    
    // 设置底线
    table.borders.bottom.style = "Single";
    table.borders.bottom.width = config.bottomWidth;

    // 设置表头底线 (第一行)
    const headerRow = table.rows.getFirst();
    headerRow.borders.bottom.style = "Single";
    headerRow.borders.bottom.width = config.headerWidth;

    await context.sync();
  });
}
```

**步骤 2: 提交**
```bash
git add src/utils/table.js
git commit -m "feat: 增加基础三线表应用函数"
```

---

### 任务 2: 实现智能表头识别

**Files:**
- Modify: `src/utils/table.js`

**步骤 1: 实现结构识别逻辑**
自动判断是否有复合表头（跨列合并）。

```javascript
export async function detectHeaderRows(table) {
  return await Word.run(async (context) => {
    table.load("rows/items/cells");
    await context.sync();
    
    let headerRowCount = 1;
    // 检查前几行是否存在单元格合并（跨列）
    for (let i = 0; i < Math.min(table.rows.items.length, 3); i++) {
        const row = table.rows.items[i];
        // 逻辑：如果某行单元格数量少于列数，或者有特定的格式，判定为表头
        // 简化版：检测单元格的 horizontalAlignment 或加粗状态
    }
    return headerRowCount;
  });
}
```

**步骤 2: 提交**
```bash
git commit -m "feat: 实现初步的表头识别逻辑"
```

---

### 任务 3: UI 界面集成 (侧边栏)

**Files:**
- Modify: `src/taskpane/taskpane.html`
- Modify: `src/taskpane/taskpane.js`

**步骤 1: 在 HTML 中增加面板**
添加“表格工具”选项卡。

```html
<div id="table-tools-panel" class="panel">
  <h3>表格美化</h3>
  <button id="btn-beautify-table">一键三线表</button>
  <button id="btn-scan-tables">扫描全文表格</button>
  <div id="table-list-container"></div>
</div>
```

**步骤 2: 编写交互代码**
绑定按钮点击事件。

**步骤 3: 提交**
```bash
git commit -m "feat: 集成表格美化 UI 面板"
```

---

### 任务 4: 批量扫描与勾选功能

**Files:**
- Modify: `src/taskpane/taskpane.js`

**步骤 1: 实现全文表格遍历**
```javascript
async function scanAllTables() {
  await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    // 渲染到 table-list-container
  });
}
```

**步骤 2: 提交**
```bash
git commit -m "feat: 实现全文表格扫描列表"
```
