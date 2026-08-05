import * as storage from "./storage.js";

const SHIELD_PREFIX = "wordai_shield_";
const BOU_START = "wordai_boundary_start";
const BOU_END = "wordai_boundary_end";
const REF_BOOKMARK_PREFIX = "wordai_ref_";

/**
 * 标记选区并保护引用
 */
export async function markSelection() {
  let finalItems = [];
  let formulaSkippedCount = 0;
  try {
    await Word.run(async (context) => {
      const doc = context.document;
      const selection = doc.getSelection();
      selection.load(["text", "isEmpty"]);
      await context.sync();

      let mainRanges = [];
      if (selection.isEmpty) {
        const ps = selection.paragraphs;
        ps.load("items");
        await context.sync();
        if (ps.items.length > 0) mainRanges = [ps.items[0].getRange()];
      } else {
        mainRanges = [selection];
      }
      if (mainRanges.length === 0) return;

      // 1. 全文参考文献书签索引 (WordApi 1.4)
      if (Office.context.requirements.isSetSupported("WordApi", "1.4")) {
        const pars = doc.body.paragraphs;
        pars.load("items");
        await context.sync();
        for (const p of pars.items) p.load("text");
        await context.sync();
        for (const p of pars.items) {
          const m = p.text.match(/^\[(\d+)\]/);
          if (m) {
            p.getRange("Start").insertBookmark(`${REF_BOOKMARK_PREFIX}${m[1]}`);
          }
        }
        await context.sync();
      }

      // 2. 读取用户配置的跳过规则
      const skipRules = storage.getSkipRules();

      // 3. 提取合法段落
      let globalCounter = 0;
      const validParagraphs = [];
      for (const range of mainRanges) {
        const ps = range.paragraphs;
        ps.load("items");
        await context.sync();

        const ooxmlPromises = [];
        for (const p of ps.items) {
          p.load(["text", "style"]);
          if (skipRules.images) {
            p.inlinePictures.load("items");
          }
          // 始终加载 OOXML 以进行公式自动探测和跳过保护
          const ooxmlObj = p.getOoxml();
          ooxmlPromises.push({ paragraph: p, ooxmlObj });
        }
        await context.sync();

        if (skipRules.tables) {
          for (const p of ps.items) {
            p.parentTableOrNullObject.load("isNullObject");
          }
          await context.sync();
        }

        let paraOoxmlMap = new Map();
        for (const { paragraph, ooxmlObj } of ooxmlPromises) {
          paraOoxmlMap.set(paragraph, ooxmlObj.value);
        }

        for (const p of ps.items) {
          const t = p.text.trim();
          if (!t) continue;

          if (
            skipRules.headings &&
            (p.style.includes("Heading") || p.style.includes("标题"))
          )
            continue;

          if (skipRules.tables && !p.parentTableOrNullObject.isNullObject)
            continue;

          // 保护：含公式（oMath）的段落跳过，避免 AI 破坏公式
          const xml = paraOoxmlMap.get(p);
          if (
            xml &&
            (xml.includes("<m:oMath") || xml.includes("<m:oMathPara"))
          ) {
            formulaSkippedCount++;
            continue;
          }

          if (
            skipRules.crossReferences &&
            /^(图|表|Figure|Table)\s*\d+/.test(t)
          )
            continue;

          if (skipRules.images && p.inlinePictures.items.length > 0) continue;

          if (
            skipRules.toc &&
            (p.style.includes("TOC") || p.style.includes("目录"))
          )
            continue;

          if (
            t.startsWith("摘要") ||
            t.startsWith("Abstract") ||
            t.includes("致谢")
          )
            continue;
          if (t.replace(/[^\w\u4e00-\u9fa5]/g, "").length === 0) continue;

          validParagraphs.push(p);
        }
      }

      if (validParagraphs.length === 0) return;

      // 4. 将提取的自然段并行标记 CC 并准备扫描引文、公式与脚注
      const searchTasks = [];
      const session = Date.now() + "_" + Math.floor(Math.random() * 100);

      for (const p of validParagraphs) {
        const startCC = p.getRange("Start").insertContentControl();
        startCC.tag = `${BOU_START}_${session}_${globalCounter}`;
        startCC.appearance = "Hidden";
        const endCC = p.getRange("End").insertContentControl();
        endCC.tag = `${BOU_END}_${session}_${globalCounter++}`;
        endCC.appearance = "Hidden";

        const refMatches = p.search("\\[[0-9\\- ,]@\\]", {
          matchWildcards: true,
        });
        refMatches.load("items");

        let footnotes = null;
        if (
          Office.context.requirements.isSetSupported("WordApi", "1.5") &&
          p.footnotes
        ) {
          footnotes = p.footnotes;
          footnotes.load("items/reference");
        }

        searchTasks.push({
          paragraph: p,
          refMatches,
          footnotes,
          boundaryTags: { start: startCC.tag, end: endCC.tag },
          startCC,
          endCC,
        });
      }
      await context.sync();

      // 5. 批量加载 OOXML 格式数据
      for (const task of searchTasks) {
        task.itemsToShield = [];
        if (task.refMatches.items) {
          for (const m of task.refMatches.items) {
            task.itemsToShield.push({
              range: m,
              type: "REF",
              xml: m.getOoxml(),
            });
          }
        }
        if (task.footnotes && task.footnotes.items) {
          for (const fn of task.footnotes.items) {
            if (fn.reference) {
              task.itemsToShield.push({
                range: fn.reference,
                type: "FNOTE",
                xml: fn.reference.getOoxml(),
              });
            }
          }
        }
      }
      await context.sync();

      // 6. 逆序对引文进行遮罩替换，并记载遮罩映射表
      const rangeListToLoad = [];
      for (const task of searchTasks) {
        const shieldMap = [];
        const allItems = [...task.itemsToShield];

        for (let i = allItems.length - 1; i >= 0; i--) {
          const item = allItems[i];
          const uid = globalCounter++;
          const token = `[${item.type}_${uid}]`;

          const cc = item.range.insertContentControl();
          cc.tag = `${SHIELD_PREFIX}${uid}`;
          cc.appearance = "Hidden";
          cc.insertText(token, "Replace");

          shieldMap.push({
            placeholder: token,
            id: uid,
            originalXml: item.xml.value,
          });
        }

        const newRange = task.startCC
          .getRange("After")
          .expandTo(task.endCC.getRange("Before"));
        newRange.load("text");
        
        finalItems.push({
          textObj: newRange,
          refMap: shieldMap,
          boundaryTags: task.boundaryTags,
        });
      }
      await context.sync();

      // 7. 正式取得带遮罩后的各物理段落文本内容
      for (const item of finalItems) {
        item.text = item.textObj.text;
        delete item.textObj;
      }
    });
  } catch (err) {
    console.error("markSelection Error:", err);
    throw err;
  }
  finalItems.skippedInfo = { formula: formulaSkippedCount };
  return finalItems;
}

/**
 * 后置自愈：为所有 [N] 格式重新建立超链接
 */
export async function autoRelinkRange(range) {
  try {
    const matches = range.search("\\[[0-9\\- ,]@\\]", { matchWildcards: true });
    matches.load("items");
    await range.context.sync();
    for (const m of matches.items) m.load("text");
    await range.context.sync();

    for (const m of matches.items) {
      const numMatch = m.text.match(/\d+/);
      if (numMatch) {
        m.hyperlink = `#${REF_BOOKMARK_PREFIX}${numMatch[0]}`;
        m.font.color = "black";
        m.font.underline = "None";
      }
    }
    await range.context.sync();
  } catch (e) {
    console.warn("自动重连文献/脚注超链接失败:", e);
  }
}

export function parseAiResult(text, refMap) {
  // 升级正则，同时支持 REF, EQN 和 FNOTE 占位符
  const regex = /[\[【「『](REF|EQN|FNOTE)_(\d+)[\]】」』]/g;
  const parts = [];
  let lastIndex = 0;
  let match;
  const placedIds = new Set();

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      parts.push({ type: "text", val: text.substring(lastIndex, match.index) });
    }
    const type = match[1];
    const id = parseInt(match[2]);
    const placeholder = `[${type}_${id}]`;

    if (refMap.some((m) => m.id === id)) {
      parts.push({ type: "ref", id: id });
      placedIds.add(id);
    }
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) {
    parts.push({ type: "text", val: text.substring(lastIndex) });
  }
  return { parts, placedIds };
}


/**
 * 替换标记内容并恢复引用
 * @param {string} aiResult - 大模型处理后的文本
 * @param {Array} refMap - 引用/公式遮罩映射
 * @param {Object} boundaryTags - 段落边界 CC 标签
 * @param {string} originalText - 原始段落文本（用于字符级对比，可选）
 */
export async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags, originalText) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    const startCC = ccs.items.find((c) => c.tag === boundaryTags.start);
    const endCC = ccs.items.find((c) => c.tag === boundaryTags.end);
    if (!startCC || !endCC) {
      console.warn("Boundary CCs missing");
      return;
    }

    // 1. 在内存中将大模型结果拆解为 AST 节点
    const { parts, placedIds } = parseAiResult(aiResult, refMap);

    // 2. 回填：开启修订模式 → Word 原生修订（TrackAll + delete + 重插，可移植接受）；
    //    否则整段直接替换。稳定优先，不使用 OOXML 字符级注入。
    const diffMode = storage.getDiffMode();
    let originalTrackingMode = null;
    if (diffMode) {
      context.document.load("changeTrackingMode");
      await context.sync();
      originalTrackingMode = context.document.changeTrackingMode;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    // 3. 删除旧文本（修订模式下生成"删除"修订；须用 delete，clear 不生成修订）
    const targetRange = startCC
      .getRange("After")
      .expandTo(endCC.getRange("Before"));
    targetRange.delete();

    // 4. 链式组装新文本与引用的原始 OOXML（修订模式下生成"插入"修订）
    let currentLoc = startCC.getRange("After");
    for (const part of parts) {
      if (part.type === "text") {
        if (part.val) {
          currentLoc = currentLoc.insertText(part.val, "After");
        }
      } else if (part.type === "ref") {
        const mapItem = refMap.find((m) => m.id === part.id);
        if (mapItem) {
          currentLoc = currentLoc.insertOoxml(mapItem.originalXml, "After");
        }
      }
    }

    // 5. 强行恢复被 AI 删掉的孤儿引用
    const orphans = refMap.filter((m) => !placedIds.has(m.id));
    for (const o of orphans) {
      currentLoc = currentLoc.insertOoxml(o.originalXml, "After");
    }

    await context.sync();

    // 6. 回写完毕，恢复 Word 原本的修订跟踪状态
    if (diffMode && originalTrackingMode != null) {
      context.document.changeTrackingMode = originalTrackingMode;
      await context.sync();
    }

    // 5. 强力擦除 AI 幻觉产生的假占位符（所有真的已经被还原为 XML，剩下的全是捏造出的纯文本垃圾）
    try {
      const fakeSearch = startCC
        .getRange("After")
        .expandTo(endCC.getRange("Before"));
      // 搜索 REF, EQN 和 FNOTE 的各种可能括号形式
      const patterns = [
        "\\[REF_[0-9]@\\]",
        "【REF_[0-9]@】",
        "「REF_[0-9]@」",
        "『REF_[0-9]@』",
        "\\[EQN_[0-9]@\\]",
        "【EQN_[0-9]@】",
        "「EQN_[0-9]@」",
        "『EQN_[0-9]@』",
        "\\[FNOTE_[0-9]@\\]",
        "【FNOTE_[0-9]@】",
        "「FNOTE_[0-9]@」",
        "『FNOTE_[0-9]@』",
      ];

      for (const p of patterns) {
        const fakes = fakeSearch.search(p, { matchWildcards: true });
        fakes.load("items");
        await context.sync();
        if (fakes.items) {
          for (const ft of fakes.items) ft.insertText("", "Replace");
        }
      }
      await context.sync();
    } catch (e) {
      console.warn("擦除幻觉生成的假占位符标记失败:", e);
    }

    // 最终自愈：在具有孤儿和新插入内容扩展后的完整段落范围内扫描
    // 【关键修复】必须在删除 startCC 和 endCC 之前获取并操作它们！
    const finalSeg = startCC
      .getRange("After")
      .expandTo(endCC.getRange("Before"));
    await autoRelinkRange(finalSeg);

    // 清理 CC：严格限制只清理本段的起始保护圈，防止把后面段落排队中的作用域给删了
    const startCCs = context.document.contentControls.getByTag(boundaryTags.start);
    const endCCs = context.document.contentControls.getByTag(boundaryTags.end);
    startCCs.load("items");
    endCCs.load("items");
    await context.sync();
    if (startCCs.items.length > 0) {
      startCCs.items[0].delete(true);
    }
    if (endCCs.items.length > 0) {
      endCCs.items[0].delete(true);
    }
    await context.sync();
  });
}

export async function executeAndReplace(processText, onStatus, signal) {
  let segments = null;
  if (onStatus) onStatus("processing", "⚓ 锚定文献中...", true);
  try {
    segments = await markSelection();
    if (!segments || segments.length === 0) throw new Error("未选中内容");

    if (onStatus)
      onStatus("processing", `📦 提取 ${segments.length} 段...`, true);

    // 现在，我们将包含整个选区信息的 segments 一把抛给业务层组合发送，防止 503 限流
    const aiTexts = await processText(segments, signal);

    if (!aiTexts || aiTexts.length !== segments.length) {
      throw new Error("大模型返回格式错乱：未能按结构处理全部段落。");
    }

    if (onStatus) onStatus("processing", `🧩 重组排版中...`, true);

    // 取得结果后，依然走单点的 AST 回填，以确保 Word 排版里夹带的图片/表格被完美留存
    for (let i = 0; i < segments.length; i++) {
      if (signal?.aborted) throw new Error("已取消");
      const aiText = aiTexts[i];
      const seg = segments[i];
      if (aiText && aiText.trim()) {
        await replaceSingleMarkedContent(aiText, seg.refMap, seg.boundaryTags, seg.text);
      }
    }
    return { result: "完成" };
  } catch (err) {
    if (segments) {
      // 智能回滚：如果大模型请求崩溃或被拒绝，必须把文档中由于第一步锁定而生成的 [REF_N] 给还原成原来的角标！
      try {
        await rollbackSegments(segments);
      } catch (rollbackErr) {
        console.error("Rollback failed:", rollbackErr);
      }
    }
    await clearMarks();
    throw err;
  }
}

/**
 * 回滚失败任务：将遮罩占位符还原为原始 OOXML，并清理边界/遮罩 ContentControl
 * @param {Array} segments - markSelection() 返回的段落任务
 */
export async function rollbackSegments(segments) {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
    await context.sync();

    for (const seg of segments || []) {
      const startCC = ccs.items.find((c) => c.tag === seg.boundaryTags.start);
      const endCC = ccs.items.find((c) => c.tag === seg.boundaryTags.end);

      if (startCC && endCC) {
        // 1. 还原 [REF_N] / [EQN_N] / [FNOTE_N] 遮罩为原始 OOXML
        const currentSeg = startCC
          .getRange("After")
          .expandTo(endCC.getRange("Before"));
        for (const mapItem of seg.refMap || []) {
          const placeholder = mapItem.placeholder.replace(/[\[\]]/g, "\\$&");
          const s = currentSeg.search(placeholder, { matchWildcards: false });
          s.load("items");
          await context.sync();
          if (s.items && s.items.length > 0) {
            for (const t of s.items) t.insertOoxml(mapItem.originalXml, "Replace");
          }
        }

        // 2. 清理边界 CC（保留文字内容）
        startCC.delete(true);
        endCC.delete(true);
      }

      // 3. 清理本任务创建的遮罩 CC（内容已还原，delete 时保留）
      for (const mapItem of seg.refMap || []) {
        const shieldCCs = context.document.contentControls.getByTag(
          `${SHIELD_PREFIX}${mapItem.id}`
        );
        shieldCCs.load("items");
        await context.sync();
        for (const c of shieldCCs.items) c.delete(true);
      }
    }
    await context.sync();
  });
}

export async function clearMarks() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag"); // ⚡ 仅 load 每一个 CC 的 tag 属性，极大地降低网络 payload 与耗时
    await context.sync();
    for (const c of ccs.items) {
      if (
        c.tag &&
        (c.tag.startsWith(SHIELD_PREFIX) ||
          c.tag.startsWith(BOU_START) ||
          c.tag.startsWith(BOU_END))
      ) {
        c.delete(true);
      }
    }
    await context.sync();
  });
}

// ==================== 字符级 diff 对比 ====================

/**
 * 字符级 LCS diff：返回 { type: "keep"|"del"|"ins", text } 序列。
 * 文本过长时返回 null（由调用方回退为整段替换）。
 */
export function diffText(original, modified) {
  const a = String(original || "");
  const b = String(modified || "");
  const n = a.length;
  const m = b.length;
  if (n * m > 4000000) return null;

  const dp = new Array(n + 1);
  for (let i = 0; i <= n; i++) dp[i] = new Array(m + 1).fill(0);
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }

  const ops = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      ops.push({ type: "keep", text: a[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      ops.push({ type: "del", text: a[i] });
      i++;
    } else {
      ops.push({ type: "ins", text: b[j] });
      j++;
    }
  }
  while (i < n) {
    ops.push({ type: "del", text: a[i] });
    i++;
  }
  while (j < m) {
    ops.push({ type: "ins", text: b[j] });
    j++;
  }
  return ops;
}

/**
 * 合并连续相同类型的操作，减少回填时的 API 调用次数。
 */
export function mergeDiffOps(ops) {
  const merged = [];
  for (const op of ops || []) {
    const last = merged[merged.length - 1];
    if (last && last.type === op.type) last.text += op.text;
    else merged.push({ type: op.type, text: op.text });
  }
  return merged;
}

/**
 * 接受全部修订（Word 原生修订）
 * @returns {Promise<number>} 接受的修订数
 */
export async function acceptAllDiff() {
  return await Word.run(async (context) => {
    const trackedChanges = context.document.body.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();
    const count = trackedChanges.items.length;
    trackedChanges.acceptAll();
    await context.sync();
    return count;
  });
}

