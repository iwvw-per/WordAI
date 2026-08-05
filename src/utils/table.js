/**
 * WordAI 表格样式工具类
 *
 * 支持应用自定义表格样式（用户自建的"三线表"等）与内置样式：
 *  - 自定义/本地化样式名 → table.style = "样式名"
 *  - 内置跨语言样式       → table.styleBuiltIn = Word.BuiltInStyleName.xxx
 *
 * 边框/颜色经验教训（此环境 Word 桌面版的 Office.js bug）：
 *  - TableBorder 设置 type/width 抛 InvalidArgument（#5219）
 *  - table.shadingColor = null 会把文字改成白色
 *  - OOXML 整表回填会破坏部分表格样式与颜色
 * 因此本模块只做"应用样式"这一种稳定操作。
 */

/**
 * 获取选区或光标所在的所有表格
 * 兼容两种场景：真正选中了表格；或光标仅停留在表格内（此时 selection.tables 为空）
 * @param {Word.RequestContext} context - Word.run 内上下文
 * @returns {Promise<Word.Table[]>}
 */
async function getSelectedTables(context) {
  const selection = context.document.getSelection();
  const tables = selection.tables;
  tables.load("items");
  await context.sync();

  if (tables.items.length > 0) return tables.items;

  // 光标停留在表格内：取光标所在段落所属的表格
  const paras = selection.paragraphs;
  paras.load("items");
  await context.sync();
  const probes = [];
  for (const p of paras.items) {
    const pt = p.parentTableOrNullObject;
    pt.load("isNullObject");
    probes.push(pt);
  }
  await context.sync();

  const result = [];
  for (const pt of probes) {
    if (!pt.isNullObject && !result.includes(pt)) result.push(pt);
  }
  return result;
}

/**
 * 读取文档中所有可用的表格样式名（含用户自定义样式，如自建的"三线表"）
 * @returns {Promise<string[]>} 去重后的表格样式名数组
 */
export async function getTableStyles() {
  return await Word.run(async (context) => {
    const styles = context.document.styles;
    styles.load("items/name,items/type");
    await context.sync();

    const names = styles.items
      .filter((s) => s && (s.type === Word.StyleType.table || s.type === "Table"))
      .map((s) => s && s.name)
      .filter(Boolean);
    return [...new Set(names)];
  });
}

/**
 * 为单个表格应用样式与宽度模式
 * @param {Word.Table} table - 带 context 的表格代理
 * @param {Object} options - {
 *   styleName: 自定义/本地化样式名,
 *   styleBuiltIn: 内置枚举键,
 *   widthMode: "window" | "content" | "fixed" | ""（保持原样）
 * }
 */
export async function optimizeTable(table, options = {}) {
  const { styleName = "", styleBuiltIn = "", widthMode = "" } = options;
  try {
    const context = table.context;

    // 1. 应用样式
    if (styleBuiltIn && Word.BuiltInStyleName && Word.BuiltInStyleName[styleBuiltIn]) {
      table.styleBuiltIn = Word.BuiltInStyleName[styleBuiltIn];
      await context.sync();
      console.log(`[table-optimize] 已应用内置样式: ${styleBuiltIn}`);
    } else if (styleName) {
      table.style = styleName;
      await context.sync();
      console.log(`[table-optimize] 已应用样式: ${styleName}`);
    }

    // 2. 宽度模式（autoFitBehavior 需桌面 1.4，window 失败时回退 autoFitWindow）
    if (widthMode) {
      try {
        if (widthMode === "window") {
          table.autoFitBehavior("Window");
        } else if (widthMode === "content") {
          table.autoFitBehavior("Content");
        } else if (widthMode === "fixed") {
          table.autoFitBehavior("FixedSize");
        }
        await context.sync();
        console.log(`[table-optimize] 已应用宽度模式: ${widthMode}`);
      } catch (e) {
        if (widthMode === "window") {
          try {
            table.autoFitWindow();
            await context.sync();
            console.log("[table-optimize] autoFitBehavior 失败，已回退 autoFitWindow");
          } catch (e2) {
            console.error("[table-optimize] 窗口自适应失败:", e2 && e2.message, e2 && e2.name);
          }
        } else {
          console.error(`[table-optimize] 宽度模式 ${widthMode} 失败:`, e && e.message, e && e.name);
        }
      }
    }
  } catch (e) {
    console.error("[table-optimize] 应用失败:", e && e.message, e && e.name);
  }
}

/**
 * 为当前选区/光标所在的表格应用样式
 * @param {Object} options - { styleName | styleBuiltIn }
 * @returns {Promise<number>} 处理的表格数量
 */
export async function optimizeSelection(options = {}) {
  return await Word.run(async (context) => {
    const tables = await getSelectedTables(context);
    if (tables.length === 0) {
      throw new Error("请先把光标点进要处理的表格，或选中表格区域");
    }
    for (const table of tables) {
      await optimizeTable(table, options);
    }
    return tables.length;
  });
}

/**
 * 为文档中所有表格应用样式
 * @param {Object} options - { styleName | styleBuiltIn }
 * @returns {Promise<number>} 处理的表格数量
 */
export async function applyStyleToAllTables(options = {}) {
  return await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    if (tables.items.length === 0) {
      throw new Error("文档中没有表格");
    }
    for (const table of tables.items) {
      await optimizeTable(table, options);
    }
    return tables.items.length;
  });
}

export { getSelectedTables };