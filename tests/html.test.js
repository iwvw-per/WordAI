import { describe, expect, it } from "vitest";
import { escapeHtml } from "../src/utils/html.js";

describe("escapeHtml", () => {
  it("escapes characters that can break out of HTML text or attributes", () => {
    expect(escapeHtml(`<img src=x onerror="alert('x')">&`)).toBe(
      "&lt;img src=x onerror=&quot;alert(&#39;x&#39;)&quot;&gt;&amp;",
    );
  });

  it("treats nullish values as empty text", () => {
    expect(escapeHtml(null)).toBe("");
    expect(escapeHtml(undefined)).toBe("");
  });
});

