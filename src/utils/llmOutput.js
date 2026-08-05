export function parseSegmentedResponse(text, segments) {
  const parsedTexts = [];
  const source = text || "";

  // 优先匹配带 id 的段落标签，兼容 id="0" / id='0' / id=0 / id = 0 等写法
  const regex = /<p\s+id\s*=\s*["']?(\d+)["']?\s*>([\s\S]*?)<\/p>/gi;
  let match;
  while ((match = regex.exec(source)) !== null) {
    const id = parseInt(match[1], 10);
    parsedTexts[id] = match[2].trim();
  }

  if (parsedTexts.some((item) => item !== undefined)) {
    // 带 id 解析成功：缺失的 id 保留 undefined，由调用方按段落回退原文，避免错位
    return parsedTexts;
  }

  // 无 id 标签：剥除所有 p 标签后按行还原（LLM 输出不规范时的兜底）
  const backupLines = source
    .replace(/<\/?p[^>]*>/gi, "\n")
    .split("\n")
    .filter((line) => line.trim() !== "");
  for (let i = 0; i < segments.length; i++) {
    parsedTexts[i] = backupLines[i] !== undefined ? backupLines[i] : segments[i].text;
  }

  return parsedTexts;
}

