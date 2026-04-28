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
            try {
              const searchResults = p.search(`\\[${m[1]}\\]`, {
                matchWildcards: true,
              });
              searchResults.load("items");
              await context.sync();
              if (searchResults.items.length > 0) {
                p.getRange("Start")
                  .expandTo(searchResults.items[0])
                  .insertBookmark(`${REF_BOOKMARK_PREFIX}${m[1]}`);
              }
            } catch (e) {}
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

        for (const p of ps.items) {
          p.load(["text", "style"]);
          // 加载 inlinePictures 用于图片检测
          if (skipRules.images) {
            p.inlinePictures.load("items");
          }
        }
        await context.sync();

        // 对需要检测表格的段落，批量加载 parentTableOrNullObject
        if (skipRules.tables) {
          for (const p of ps.items) {
            p.parentTableOrNullObject.load("isNullObject");
          }
          await context.sync();
        }

        // 对需要检测公式的段落，批量获取 OOXML 做 oMath 检测
        let paraOoxmlMap = new Map();
        if (skipRules.formulas) {
          const ooxmlPromises = [];
          for (const p of ps.items) {
            const t = p.text.trim();
            if (!t) continue;
            const ooxmlObj = p.getOoxml();
            ooxmlPromises.push({ paragraph: p, ooxmlObj });
          }
          await context.sync();
          for (const { paragraph, ooxmlObj } of ooxmlPromises) {
            paraOoxmlMap.set(paragraph, ooxmlObj.value);
          }
        }

        for (const p of ps.items) {
          const t = p.text.trim();
          if (!t) continue;

          // 跳过标题
          if (
            skipRules.headings &&
            (p.style.includes("Heading") || p.style.includes("标题"))
          )
            continue;

          // 跳过表格内的段落
          if (skipRules.tables && !p.parentTableOrNullObject.isNullObject)
            continue;

          // 跳过包含公式（行内/单行）的段落
          if (skipRules.formulas) {
            const xml = paraOoxmlMap.get(p);
            if (
              xml &&
              (xml.includes("<m:oMath") || xml.includes("<m:oMathPara"))
            )
              continue;
          }

          // 跳过交叉引用（图表标题段落）
          if (
            skipRules.crossReferences &&
            /^(图|表|Figure|Table)\s*\d+/.test(t)
          )
            continue;

          // 跳过图片段落
          if (skipRules.images && p.inlinePictures.items.length > 0) continue;

          // 跳过目录段落
          if (
            skipRules.toc &&
            (p.style.includes("TOC") || p.style.includes("目录"))
          )
            continue;

          // 跳过摘要/致谢标题
          if (
            t.startsWith("摘要") ||
            t.startsWith("Abstract") ||
            t.includes("致谢")
          )
            continue;
          // 跳过纯符号段落
          if (t.replace(/[^\w\u4e00-\u9fa5]/g, "").length === 0) continue;

          validParagraphs.push(p);
        }
      }

      if (validParagraphs.length === 0) return;

      // 第一步：批量搜索引文与公式
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

        let eqns = null;
        if (p.equations) {
          eqns = p.equations;
          eqns.load("items");
        }

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
          eqns,
          footnotes,
          boundaryTags: { start: startCC.tag, end: endCC.tag },
          startCC,
          endCC,
        });
      }
      await context.sync();

      // 第二步：批量加载 OOXML
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
        if (task.eqns && task.eqns.items) {
          for (const eq of task.eqns.items) {
            task.itemsToShield.push({
              range: eq,
              type: "EQN",
              xml: eq.getOoxml(),
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

      // 第三步：逆序替换
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

        task.shieldMap = shieldMap;
        const newRange = task.startCC
          .getRange("After")
          .expandTo(task.endCC.getRange("Before"));
        newRange.load("text");
        task.newRange = newRange;
      }
      await context.sync();

      // 第四步：提取结果
      for (const task of searchTasks) {
        finalItems.push({
          text: task.newRange.text,
          refMap: task.shieldMap,
          boundaryTags: task.boundaryTags,
        });
      }
    });
  } catch (err) {
    console.error("markSelection Error:", err);
    throw err;
  }
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
  } catch (e) {}
}

function parseAiResult(text, refMap) {
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
 */
async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags) {
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

    // 2. 清空原本的旧文本
    const targetRange = startCC
      .getRange("After")
      .expandTo(endCC.getRange("Before"));
    targetRange.clear();

    // 3. 链式组装新文本与引用的原始 OOXML
    let currentLoc = startCC.getRange("After");
    for (const part of parts) {
      if (part.type === "text") {
        currentLoc = currentLoc.insertText(part.val, "After");
      } else if (part.type === "ref") {
        const mapItem = refMap.find((m) => m.id === part.id);
        if (mapItem) {
          currentLoc = currentLoc.insertOoxml(mapItem.originalXml, "After");
        }
      }
    }

    // 4. 强行恢复被 AI 删掉的孤儿引用
    const orphans = refMap.filter((m) => !placedIds.has(m.id));
    for (const o of orphans) {
      currentLoc = currentLoc.insertOoxml(o.originalXml, "After");
    }

    await context.sync();

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
    } catch (e) {}

    // 最终自愈：在具有孤儿和新插入内容扩展后的完整段落范围内扫描
    // 【关键修复】必须在删除 startCC 和 endCC 之前获取并操作它们！
    const finalSeg = startCC
      .getRange("After")
      .expandTo(endCC.getRange("Before"));
    await autoRelinkRange(finalSeg);

    // 清理 CC：严格限制只清理本段的起始保护圈，防止把后面段落排队中的作用域给删了
    const allCCs = context.document.contentControls;
    allCCs.load("items");
    await context.sync();
    for (const c of allCCs.items) {
      if (
        c.tag &&
        (c.tag === boundaryTags.start || c.tag === boundaryTags.end)
      ) {
        c.delete(true);
      }
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
        await replaceSingleMarkedContent(aiText, seg.refMap, seg.boundaryTags);
      }
    }
    return { result: "完成" };
  } catch (err) {
    if (segments) {
      // 智能回滚：如果大模型请求崩溃或被拒绝，必须把文档中由于第一步锁定而生成的 [REF_N] 给还原成原来的角标！
      try {
        await Word.run(async (context) => {
          const ccs = context.document.contentControls;
          ccs.load("items");
          await context.sync();

          for (const seg of segments) {
            const startCC = ccs.items.find(
              (c) => c.tag === seg.boundaryTags.start,
            );
            const endCC = ccs.items.find((c) => c.tag === seg.boundaryTags.end);
            if (startCC && endCC) {
              const currentSeg = startCC
                .getRange("After")
                .expandTo(endCC.getRange("Before"));
              for (const mapItem of seg.refMap) {
                // 修正回滚正则，支持 REF, EQN 和 FNOTE
                const placeholder = mapItem.placeholder
                  .replace("[", "\\[")
                  .replace("]", "\\]");
                const s = currentSeg.search(placeholder, {
                  matchWildcards: false,
                });
                s.load("items");
                await context.sync();
                if (s.items && s.items.length > 0) {
                  for (const t of s.items)
                    t.insertOoxml(mapItem.originalXml, "Replace");
                }
              }
            }
          }
        });
      } catch (rollbackErr) {
        console.error("Rollback failed:", rollbackErr);
      }
    }
    await clearMarks();
    throw err;
  }
}

export async function clearMarks() {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items");
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
