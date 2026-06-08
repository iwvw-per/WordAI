export function parseSegmentedResponse(text, segments) {
  const parsedTexts = [];
  const source = text || "";
  const regex = /<p\s+id="(\d+)">([\s\S]*?)<\/p>/gi;
  let match;

  while ((match = regex.exec(source)) !== null) {
    const id = parseInt(match[1], 10);
    parsedTexts[id] = match[2].trim();
  }

  if (parsedTexts.some((item) => item !== undefined)) {
    return parsedTexts;
  }

  const backupLines = source.split("\n").filter((line) => line.trim() !== "");
  for (let i = 0; i < segments.length; i++) {
    parsedTexts[i] = backupLines[i] !== undefined ? backupLines[i] : segments[i].text;
  }

  return parsedTexts;
}

