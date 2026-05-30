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

        for (const p of ps.items) {
          p.load(["text", "style"]);
          if (skipRules.images) {
            p.inlinePictures.load("items");
          }
        }
        await context.sync();

        if (skipRules.tables) {
          for (const p of ps.items) {
            p.parentTableOrNullObject.load("isNullObject");
          }
          await context.sync();
        }

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

          if (
            skipRules.headings &&
            (p.style.includes("Heading") || p.style.includes("标题"))
          )
            continue;

          if (skipRules.tables && !p.parentTableOrNullObject.isNullObject)
            continue;

          if (skipRules.formulas) {
            const xml = paraOoxmlMap.get(p);
            if (
              xml &&
              (xml.includes("<m:oMath") || xml.includes("<m:oMathPara"))
            )
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

      // 4. 将整个合法段落集合包装为唯一一个整体 Task
      const session = Date.now() + "_" + Math.floor(Math.random() * 100);
      const startCC = validParagraphs[0].getRange("Start").insertContentControl();
      startCC.tag = `${BOU_START}_${session}_0`;
      startCC.appearance = "Hidden";
      
      const endCC = validParagraphs[validParagraphs.length - 1].getRange("End").insertContentControl();
      endCC.tag = `${BOU_END}_${session}_0`;
      endCC.appearance = "Hidden";
      await context.sync();

      const totalRange = startCC.getRange("After").expandTo(endCC.getRange("Before"));

      // 批量搜索引文、公式与脚注
      const refMatches = totalRange.search("\\[[0-9\\- ,]@\\]", {
        matchWildcards: true,
      });
      refMatches.load("items");

      let eqns = null;
      if (totalRange.equations) {
        eqns = totalRange.equations;
        eqns.load("items");
      }

      let footnotes = null;
      if (
        Office.context.requirements.isSetSupported("WordApi", "1.5") &&
        totalRange.footnotes
      ) {
        footnotes = totalRange.footnotes;
        footnotes.load("items/reference");
      }
      await context.sync();

      // 批量搜集并加载 OOXML
      const itemsToShield = [];
      if (refMatches.items) {
        for (const m of refMatches.items) {
          itemsToShield.push({
            range: m,
            type: "REF",
            xml: m.getOoxml(),
          });
        }
      }
      if (eqns && eqns.items) {
        for (const eq of eqns.items) {
          itemsToShield.push({
            range: eq,
            type: "EQN",
            xml: eq.getOoxml(),
          });
        }
      }
      if (footnotes && footnotes.items) {
        for (const fn of footnotes.items) {
          if (fn.reference) {
            itemsToShield.push({
              range: fn.reference,
              type: "FNOTE",
              xml: fn.reference.getOoxml(),
            });
          }
        }
      }
      await context.sync();

      // 逆序进行 ContentControl 标记保护
      const shieldMap = [];
      for (let i = itemsToShield.length - 1; i >= 0; i--) {
        const item = itemsToShield[i];
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
      await context.sync();

      // 重新读取带有遮罩后的整体 Range 文本
      totalRange.load("text");
      await context.sync();

      finalItems.push({
        text: totalRange.text,
        refMap: shieldMap,
        boundaryTags: { start: startCC.tag, end: endCC.tag },
      });
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

function splitPartsIntoParagraphs(parts) {
  const paras = [];
  let currentPara = [];

  for (const part of parts) {
    if (part.type === "text") {
      const normalizedText = part.val.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      const lines = normalizedText.split("\n");
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (i > 0) {
          paras.push(currentPara);
          currentPara = [];
        }
        currentPara.push({ type: "text", val: line });
      }
    } else if (part.type === "ref") {
      currentPara.push(part);
    }
  }
  paras.push(currentPara);
  return paras;
}

/**
 * 替换标记内容并恢复引用
 */
export async function replaceSingleMarkedContent(aiResult, refMap, boundaryTags) {
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

    // 1. 在内存中将大模型结果拆解为 AST 节点，并按段落进行划分子树
    const { parts, placedIds } = parseAiResult(aiResult, refMap);
    const parasAST = splitPartsIntoParagraphs(parts);

    // 2. 加载选区内的所有原有物理段落，以便后续复用与精确排版
    const targetRange = startCC
      .getRange("After")
      .expandTo(endCC.getRange("Before"));
    const existingParas = targetRange.paragraphs;
    existingParas.load("items");
    await context.sync();

    // 3. 清空原本的旧文本（内容清空，残留的物理空壳会被自动复用或清除）
    targetRange.clear();
    await context.sync();

    // 4. 复用原有物理段落或在后面创建新段落来承载改写后的段落 AST
    const totalExisting = existingParas.items.length;
    const totalNew = parasAST.length;
    const paragraphRefs = [];

    for (let k = 0; k < totalNew; k++) {
      const paraAST = parasAST[k];
      let pRange = null;

      if (k < totalExisting) {
        // 复用原有段落空壳，绝不把空行往后推
        const currentPara = existingParas.items[k];
        paragraphRefs.push(currentPara);
        pRange = currentPara.getRange();
      } else {
        // 大模型改写后的段落多于原本段落，在最后一个已写完的物理段落后面新建物理段落
        const prevPara = paragraphRefs[paragraphRefs.length - 1];
        const newPara = prevPara.insertParagraph("", "After");
        paragraphRefs.push(newPara);
        pRange = newPara.getRange();
      }

      // 在该物理段落内部，链式回写属于本段落的所有文本和引文
      let currentLoc = pRange.getRange("Start");
      for (const part of paraAST) {
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
    }

    // 5. 如果大模型改写后的段落少于原本的段落，物理删除后部多余残留的空壳段落以彻底抹除空行！
    if (totalNew < totalExisting) {
      for (let k = totalNew; k < totalExisting; k++) {
        existingParas.items[k].delete();
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
