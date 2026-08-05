import { describe, expect, it, beforeEach } from "vitest";
import { getPrompts } from "../src/utils/storage.js";

function createLocalStorageMock() {
  const store = new Map();
  return {
    getItem: (k) => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => store.set(k, String(v)),
    removeItem: (k) => store.delete(k),
    clear: () => store.clear(),
  };
}

describe("getPrompts 迁移逻辑", () => {
  beforeEach(() => {
    globalThis.localStorage = createLocalStorageMock();
  });

  it("旧版默认降AI提示词（含特征串）迁移为专家版", () => {
    const oldPrompts = [
      {
        id: "deai",
        name: "降AI",
        icon: "🎭",
        prompt:
          "请改写以下文字，使其更像人类自然书写的风格。避免AI常见的套话和模式化表达，增加口语化、个性化的表达方式。保持原文含义不变，保持原文的语言。只输出改写后的文字，不要输出任何解释。",
        color: "#ec4899",
      },
    ];
    localStorage.setItem("wordai_prompts", JSON.stringify(oldPrompts));
    const prompts = getPrompts();
    expect(prompts[0].prompt).toContain("你的角色与目标");
    expect(prompts[0].prompt).toContain("【学术降温优化指令】");
  });

  it("用户自定义短提示词不被迁移覆盖", () => {
    const customPrompts = [
      { id: "deai", name: "降AI", icon: "🎭", prompt: "帮我改得自然点", color: "#ec4899" },
    ];
    localStorage.setItem("wordai_prompts", JSON.stringify(customPrompts));
    const prompts = getPrompts();
    expect(prompts[0].prompt).toBe("帮我改得自然点");
  });

  it("用户自定义长提示词即使不含学术降温指令也不被覆盖", () => {
    const customPrompts = [
      {
        id: "deai",
        name: "降AI",
        icon: "🎭",
        prompt:
          "这是用户精心编写的自定义降AI提示词，包含各种个性化规则、措辞偏好与示例段落，整体长度已经明显超过两百个字符，但绝不应被内置默认值静默覆盖。",
        color: "#ec4899",
      },
    ];
    localStorage.setItem("wordai_prompts", JSON.stringify(customPrompts));
    const prompts = getPrompts();
    expect(prompts[0].prompt).toContain("这是用户精心编写的自定义降AI提示词");
  });
});
