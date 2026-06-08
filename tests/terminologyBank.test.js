import { describe, expect, it } from "vitest";
import {
  findTerminologyBankConflicts,
  mergeTerminologyConflicts,
  parseTerminologyBank,
} from "../src/utils/terminologyBank.js";

describe("terminology bank", () => {
  it("parses one entry per line", () => {
    expect(parseTerminologyBank("Large Language Model=LLM,large model\n")).toEqual([
      { standard: "Large Language Model", aliases: ["LLM", "large model"] },
    ]);
  });

  it("finds aliases present in text", () => {
    const entries = parseTerminologyBank("Large Language Model=LLM,large model");
    expect(findTerminologyBankConflicts("This LLM is useful.", entries)).toEqual([
      { standard: "Large Language Model", aliases: ["LLM"] },
    ]);
  });

  it("merges conflicts by standard term", () => {
    expect(
      mergeTerminologyConflicts(
        [{ standard: "A", aliases: ["x"] }],
        [{ standard: "A", aliases: ["x", "y"] }],
      ),
    ).toEqual([{ standard: "A", aliases: ["x", "y"] }]);
  });
});

