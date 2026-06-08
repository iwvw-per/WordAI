import { beforeEach, describe, expect, it, vi } from "vitest";
import * as llm from "../src/utils/llm.js";
import * as storage from "../src/utils/storage.js";
import { generateAbstract } from "../src/utils/abstract.js";

vi.mock("../src/utils/llm.js", () => ({
  callLLM: vi.fn(),
}));

vi.mock("../src/utils/storage.js", () => ({
  getPrivacyMode: vi.fn(),
  getRoutedModel: vi.fn(),
}));

function mockWordBodyText(text) {
  globalThis.Word = {
    run: vi.fn(async (callback) => {
      const body = {
        text,
        load: vi.fn(),
      };
      const context = {
        document: { body },
        sync: vi.fn(async () => {}),
      };
      await callback(context);
    }),
  };
}

describe("generateAbstract privacy mode", () => {
  beforeEach(() => {
    vi.mocked(llm.callLLM).mockReset();
    vi.mocked(storage.getPrivacyMode).mockReset();
    vi.mocked(storage.getRoutedModel).mockReset();
    vi.mocked(storage.getRoutedModel).mockReturnValue("quality-model");
  });

  it("redacts document text before sending it to the LLM", async () => {
    mockWordBodyText("Author email: user@example.com. Phone: 13800138000.");
    vi.mocked(storage.getPrivacyMode).mockReturnValue(true);
    vi.mocked(llm.callLLM).mockResolvedValue("摘要");

    await generateAbstract();

    const [prompt, userContent, signal, options] = vi.mocked(llm.callLLM).mock.calls[0];
    expect(prompt).toContain("[[WAI_SECRET_N]]");
    expect(userContent).toContain("[[WAI_SECRET_0]]");
    expect(userContent).toContain("[[WAI_SECRET_1]]");
    expect(userContent).not.toContain("user@example.com");
    expect(userContent).not.toContain("13800138000");
    expect(signal).toBeUndefined();
    expect(options).toEqual({ model: "quality-model" });
  });

  it("restores placeholders in the generated abstract", async () => {
    mockWordBodyText("Author email: user@example.com.");
    vi.mocked(storage.getPrivacyMode).mockReturnValue(true);
    vi.mocked(llm.callLLM).mockResolvedValue("摘要包含 [[WAI_SECRET_0]]");

    await expect(generateAbstract()).resolves.toBe("摘要包含 user@example.com");
  });
});
