export function describeModelRoute({ enabled, baseModel, fastModel, qualityModel, actionName }) {
  if (!enabled) return { model: baseModel, route: "base" };
  const name = String(actionName || "").toLowerCase();
  const qualityPattern = /(academic|abstract|term|reference|deai|cool|降|摘要|术语|文献|引用)/i;
  if (qualityPattern.test(name)) {
    return { model: qualityModel || baseModel, route: qualityModel ? "quality" : "base" };
  }
  return { model: fastModel || baseModel, route: fastModel ? "fast" : "base" };
}

