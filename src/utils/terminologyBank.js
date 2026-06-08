export function parseTerminologyBank(raw) {
  return String(raw || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [standardPart, aliasesPart = ""] = line.split("=");
      const standard = (standardPart || "").trim();
      const aliases = aliasesPart
        .split(/[,，]/)
        .map((item) => item.trim())
        .filter(Boolean)
        .filter((item) => item !== standard);
      return { standard, aliases };
    })
    .filter((entry) => entry.standard && entry.aliases.length > 0);
}

export function findTerminologyBankConflicts(text, entries) {
  const source = String(text || "");
  return entries
    .map((entry) => ({
      standard: entry.standard,
      aliases: entry.aliases.filter((alias) => source.includes(alias)),
    }))
    .filter((entry) => entry.aliases.length > 0);
}

export function mergeTerminologyConflicts(...groups) {
  const map = new Map();
  for (const group of groups) {
    for (const item of group || []) {
      if (!item?.standard) continue;
      const current = map.get(item.standard) || new Set();
      for (const alias of item.aliases || []) {
        if (alias && alias !== item.standard) current.add(alias);
      }
      map.set(item.standard, current);
    }
  }
  return [...map.entries()]
    .map(([standard, aliases]) => ({ standard, aliases: [...aliases] }))
    .filter((item) => item.aliases.length > 0);
}

