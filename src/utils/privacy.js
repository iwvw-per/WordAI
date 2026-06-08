const DEFAULT_PATTERNS = [
  /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi,
  /\b(?:sk|pk|rk)-[A-Za-z0-9_-]{12,}\b/g,
  /\b1[3-9]\d{9}\b/g,
  /\b(?:\+?\d[\d\s().-]{7,}\d)\b/g,
];

export function redactSensitiveText(text, patterns = DEFAULT_PATTERNS) {
  let redacted = String(text || "");
  const replacements = [];

  for (const pattern of patterns) {
    redacted = redacted.replace(pattern, (match) => {
      const token = `[[WAI_SECRET_${replacements.length}]]`;
      replacements.push({ token, value: match });
      return token;
    });
  }

  return { text: redacted, replacements };
}

export function restoreSensitiveText(text, replacements = []) {
  let restored = String(text || "");
  for (const item of replacements) {
    restored = restored.split(item.token).join(item.value);
  }
  return restored;
}

