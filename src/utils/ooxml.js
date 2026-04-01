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
        for (const p of pars.items) {
          p.load("text");
        }
        await context.sync();
        for (const p of pars.items) {
          const m = p.text.match(/^\[(\d+)\]/);
          if (m) {
            try { 
              const searchResults = p.search(`\\[${m[1]}\\]`, { matchWildcards: true });
              searchResults.load("items");
              await context.sync();
              if (searchResults.items.length > 0) {
                 p.getRange("Start").expandTo(searchResults.items[0]).insertBookmark(`${REF_BOOKMARK_PREFIX}${m[1]}`); 
              }
            } catch(e){}
          }
        }
        await context.sync();
      }

      // 2. 读取用户配置的跳过规则
      // 默认开启跳过特性，如果 localStorage 中没有则用默认值
      const skipRulesStr = localStorage.getItem("wordai_skip_rules");
      const skipRules = skipRulesStr ? JSON.parse(skipRulesStr) : {
        skipHeadings: true,
        skipFigures: true,
        skipQuotes: true,
      };

      // 3. 细化处理选区，支持跳过标题等特殊格式
      let globalCounter = 0;
      
      // 提取选区内所有的段落，进行细粒度控制
      const validParagraphs = [];
      for (const range of mainRanges) {
        const ps = range.paragraphs;
        ps.load("items");
        await context.sync();
        
        for (const p of ps.items) {
          p.load(["text", "style"]);
        }
        await context.sync();
        
        for (const p of ps.items) {
          const t = p.text.trim();
          if (!t) continue;
          
          // ==== 核心跳过逻辑 ====
          
          // 1. 跳过各级标题 (按样式名称判断，包括中英文环境)
          if (skipRules.skipHeadings && (p.style.includes("Heading") || p.style.includes("标题"))) continue;
          
          // 2. 跳过图表标题
          // 以“图 X”或“表 X”开头
          if (skipRules.skipFigures && /^图\s*\d+|^表\s*\d+/.test(t)) continue;
          
          // 3. 跳过摘要、致谢等特定固定格式段落
          if (t.startsWith("摘要") || t.startsWith("Abstract") || t.includes("致谢")) continue;

          // 4. 不再根据字数随意跳过，只跳过纯标点片段
          if (t.replace(/[^\w\u4e00-\u9fa5]/g, "").length === 0) continue;

          validParagraphs.push(p);
        }
      }

      if (validParagraphs.length === 0) return;

      // 第一步：全局批量发起搜索所有段落里的角标
      const searchTasks = [];
      const session = Date.now() + "_" + Math.floor(Math.random() * 100);

      for (const p of validParagraphs) {
        // 先圈定边界 CC，因为包含引用的文段后续拉宽缩小都不会影响包裹框
        const startCC = p.getRange("Start").insertContentControl();
        startCC.tag = `${BOU_START}_${session}_${globalCounter}`;
        startCC.appearance = "Hidden";
        const endCC = p.getRange("End").insertContentControl();
        endCC.tag = `${BOU_END}_${session}_${globalCounter++}`;
        endCC.appearance = "Hidden";

        const refMatches = p.search("\\[[0-9\\- ,]@\\]", { matchWildcards: true });
        refMatches.load("items");
        searchTasks.push({ paragraph: p, refMatches: refMatches, boundaryTags: { start: startCC.tag, end: endCC.tag }, startCC, endCC });
      }
      await context.sync(); // 全局仅 1 次 Sync

      // 第二步：全局批量读取所有找到的角标的底层 OOXML
      for (const task of searchTasks) {
        task.xmlPromises = [];
        if (task.refMatches.items) {
          for (const m of task.refMatches.items) {
             task.xmlPromises.push(m.getOoxml());
          }
        }
      }
      await context.sync(); // 全局仅 1 次 Sync

      // 第三步：全局逆序切入占位符与保护套
      for (const task of searchTasks) {
        const localMap = [];
        const items = task.refMatches.items;
        
        if (items && items.length > 0) {
            for (let i = items.length - 1; i >= 0; i--) {
                const match = items[i];
                const uid = globalCounter++;
                const token = `[REF_${uid}]`;
                const cc = match.insertContentControl();
                cc.tag = `${SHIELD_PREFIX}${uid}`;
                cc.appearance = "Hidden";
                cc.insertText(token, "Replace");
                localMap.push({ placeholder: token, id: uid, originalXml: task.xmlPromises[i].value });
            }
        }

        task.refMap = localMap;
        
        // 获取插入占位符后，当前段实质承载了 [REF_N] 的新内容（准备一次性 load）
        const newRange = task.startCC.getRange("After").expandTo(task.endCC.getRange("Before"));
        newRange.load("text");
        task.newRange = newRange;
      }
      await context.sync(); // 全局仅 1 次 Sync 写入所有段落变化并提取 Range

      // 第四步：从载入完毕的 Range 中剥离所需的纯文本赋给 LLM 处理流
      for (const task of searchTasks) {
        finalItems.push({
          text: task.newRange.text,
          refMap: task.refMap,
          boundaryTags: task.boundaryTags
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
    const regex = /[\[【「『]REF_(\d+)[\]】」』]/g;
    const parts = [];
    let lastIndex = 0;
    let match;
    const placedIds = new Set();
    
    while ((match = regex.exec(text)) !== null) {
        if (match.index > lastIndex) {
            parts.push({ type: "text", val: text.substring(lastIndex, match.index) });
        }
        const id = parseInt(match[1]);
        if (refMap.some(m => m.id === id)) {
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

    const startCC = ccs.items.find(c => c.tag === boundaryTags.start);
    const endCC = ccs.items.find(c => c.tag === boundaryTags.end);
    if (!startCC || !endCC) {
      console.warn("Boundary CCs missing");
      return;
    }

    // 1. 在内存中将大模型结果拆解为 AST 节点
    const { parts, placedIds } = parseAiResult(aiResult, refMap);

    // 2. 清空原本的旧文本
    const targetRange = startCC.getRange("After").expandTo(endCC.getRange("Before"));
    targetRange.clear();

    // 3. 链式组装新文本与引用的原始 OOXML
    let currentLoc = startCC.getRange("After");
    for (const part of parts) {
        if (part.type === "text") {
            currentLoc = currentLoc.insertText(part.val, "After");
        } else if (part.type === "ref") {
            const mapItem = refMap.find(m => m.id === part.id);
            if (mapItem) {
                currentLoc = currentLoc.insertOoxml(mapItem.originalXml, "After");
            }
        }
    }

    // 4. 强行恢复被 AI 删掉的孤儿引用
    const orphans = refMap.filter(m => !placedIds.has(m.id));
    for (const o of orphans) {
        currentLoc = currentLoc.insertOoxml(o.originalXml, "After");
    }

    await context.sync();
    
    // 5. 强力擦除 AI 幻觉产生的假占位符（所有真的已经被还原为 XML，剩下的全是捏造出的纯文本垃圾）
    try {
        const fakeSearch = startCC.getRange("After").expandTo(endCC.getRange("Before"));
        const fakes1 = fakeSearch.search("\\[REF_[0-9]@\\]", { matchWildcards: true });
        const fakes2 = fakeSearch.search("【REF_[0-9]@】", { matchWildcards: true });
        const fakes3 = fakeSearch.search("「REF_[0-9]@」", { matchWildcards: true });
        const fakes4 = fakeSearch.search("『REF_[0-9]@』", { matchWildcards: true });
        fakes1.load("items"); fakes2.load("items"); fakes3.load("items"); fakes4.load("items");
        await context.sync();

        const allFakes = [];
        if (fakes1.items) allFakes.push(...fakes1.items);
        if (fakes2.items) allFakes.push(...fakes2.items);
        if (fakes3.items) allFakes.push(...fakes3.items);
        if (fakes4.items) allFakes.push(...fakes4.items);

        for (const ft of allFakes) {
            ft.insertText("", "Replace");
        }
        await context.sync();
    } catch (e) {}

    // 最终自愈：在具有孤儿和新插入内容扩展后的完整段落范围内扫描
    // 【关键修复】必须在删除 startCC 和 endCC 之前获取并操作它们！
    const finalSeg = startCC.getRange("After").expandTo(endCC.getRange("Before"));
    await autoRelinkRange(finalSeg);

    // 清理 CC：严格限制只清理本段的起始保护圈，防止把后面段落排队中的作用域给删了
    const allCCs = context.document.contentControls;
    allCCs.load("items");
    await context.sync();
    for (const c of allCCs.items) {
        if (c.tag && (c.tag === boundaryTags.start || c.tag === boundaryTags.end)) {
            c.delete(true);
        }
    }
    await context.sync();
  });
}

export async function executeAndReplace(processText, onStatus, signal) {
  let segments = null;
  if (onStatus) onStatus("processing", "正在锁定选区与处理引用...", true);
  try {
    segments = await markSelection();
    if (!segments || segments.length === 0) throw new Error("未选中内容");

    // 现在，我们将包含整个选区信息的 segments 一把抛给业务层组合发送，防止 503 限流
    const aiTexts = await processText(segments, signal);
    
    if (!aiTexts || aiTexts.length !== segments.length) {
        throw new Error("大模型返回格式错乱：未能按结构处理全部段落。");
    }

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
                    const startCC = ccs.items.find(c => c.tag === seg.boundaryTags.start);
                    const endCC = ccs.items.find(c => c.tag === seg.boundaryTags.end);
                    if (startCC && endCC) {
                        for (const mapItem of seg.refMap) {
                            const currentSeg = startCC.getRange("After").expandTo(endCC.getRange("Before"));
                            const s = currentSeg.search(`[REF_${mapItem.id}]`, { matchWildcards: false });
                            s.load("items");
                            await context.sync();
                            if (s.items && s.items.length > 0) {
                                for (const t of s.items) t.insertOoxml(mapItem.originalXml, "Replace");
                                await context.sync();
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
            if (c.tag && (c.tag.startsWith(SHIELD_PREFIX) || c.tag.startsWith(BOU_START) || c.tag.startsWith(BOU_END))) {
                c.delete(true);
            }
        }
        await context.sync();
    });
}
