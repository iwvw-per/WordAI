import { beforeEach, describe, expect, it, vi } from "vitest";
import * as llm from "../src/utils/llm.js";
import * as storage from "../src/utils/storage.js";
import { extractTerminology } from "../src/utils/terminology.js";

vi.mock("../src/utils/llm.js", () => ({
  callLLM: vi.fn(),
}));

vi.mock("../src/utils/storage.js", () => ({
  getPrivacyMode: vi.fn(),
  getRoutedModel: vi.fn(),
}));

describe("extractTerminology privacy mode", () => {
  beforeEach(() => {
    vi.mocked(llm.callLLM).mockReset();
    vi.mocked(storage.getPrivacyMode).mockReset();
    vi.mocked(storage.getRoutedModel).mockReset();
    vi.mocked(storage.getRoutedModel).mockReturnValue("quality-model");
  });

  it("redacts sensitive values before calling the LLM", async () => {
    vi.mocked(storage.getPrivacyMode).mockReturnValue(true);
    vi.mocked(llm.callLLM).mockResolvedValue("[]");

    await extractTerminology("Contact user@example.com or 13800138000.");

    const [prompt, userContent, signal, options] = vi.mocked(llm.callLLM).mock.calls[0];
    expect(prompt).toContain("[[WAI_SECRET_N]]");
    expect(userContent).toContain("[[WAI_SECRET_0]]");
    expect(userContent).toContain("[[WAI_SECRET_1]]");
    expect(userContent).not.toContain("user@example.com");
    expect(userContent).not.toContain("13800138000");
    expect(signal).toBeUndefined();
    expect(options).toEqual({ model: "quality-model" });
  });

  it("restores placeholders in the model result before parsing", async () => {
    vi.mocked(storage.getPrivacyMode).mockReturnValue(true);
    vi.mocked(llm.callLLM).mockResolvedValue('[{"standard":"Contact","aliases":["[[WAI_SECRET_0]]"]}]');

    await expect(extractTerminology("Contact user@example.com.")).resolves.toEqual([
      { standard: "Contact", aliases: ["user@example.com"] },
    ]);
  });
});
