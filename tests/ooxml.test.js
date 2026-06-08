import { describe, expect, it } from "vitest";
import { parseAiResult } from "../src/utils/ooxml.js";

describe("parseAiResult", () => {
  it("splits text around protected placeholders", () => {
    const result = parseAiResult("Before [REF_1] after", [{ id: 1, originalXml: "<xml />" }]);

    expect(result.parts).toEqual([
      { type: "text", val: "Before " },
      { type: "ref", id: 1 },
      { type: "text", val: " after" },
    ]);
    expect(result.placedIds.has(1)).toBe(true);
  });
});

