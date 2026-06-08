import { describe, expect, it } from "vitest";
import { parseSegmentedResponse } from "../src/utils/llmOutput.js";

const segments = [{ text: "first original" }, { text: "second original" }];

describe("parseSegmentedResponse", () => {
  it("parses model output wrapped in paragraph ids", () => {
    expect(parseSegmentedResponse('<p id="0">First</p>\n<p id="1">Second</p>', segments)).toEqual([
      "First",
      "Second",
    ]);
  });

  it("falls back to line splitting when paragraph ids are missing", () => {
    expect(parseSegmentedResponse("First\nSecond", segments)).toEqual(["First", "Second"]);
  });

  it("preserves original text when fallback output has fewer lines than segments", () => {
    expect(parseSegmentedResponse("Only first", segments)).toEqual(["Only first", "second original"]);
  });
});

