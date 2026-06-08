import { describe, expect, it } from "vitest";
import { redactSensitiveText, restoreSensitiveText } from "../src/utils/privacy.js";

describe("privacy redaction", () => {
  it("redacts and restores common sensitive values", () => {
    const original = "Contact me at user@example.com or 13800138000.";
    const redacted = redactSensitiveText(original);

    expect(redacted.text).toContain("[[WAI_SECRET_0]]");
    expect(redacted.text).toContain("[[WAI_SECRET_1]]");
    expect(restoreSensitiveText(redacted.text, redacted.replacements)).toBe(original);
  });
});

