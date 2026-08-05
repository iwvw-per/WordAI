import { describe, expect, it } from "vitest";
import { diffText, mergeDiffOps } from "../src/utils/ooxml.js";

describe("diffText 字符级 diff", () => {
  it("无变化时全部为 keep", () => {
    const ops = diffText("hello", "hello");
    expect(ops.every((o) => o.type === "keep")).toBe(true);
    expect(ops.map((o) => o.text).join("")).toBe("hello");
  });

  it("识别删除与插入", () => {
    const ops = mergeDiffOps(diffText("abcd", "abxd"));
    const del = ops.find((o) => o.type === "del");
    const ins = ops.find((o) => o.type === "ins");
    expect(del).toBeTruthy();
    expect(del.text).toContain("c");
    expect(ins).toBeTruthy();
    expect(ins.text).toContain("x");
  });

  it("合并后 keep 文本为公共子序列", () => {
    const ops = mergeDiffOps(diffText("你好世界", "你好新世界"));
    const kept = ops.filter((o) => o.type === "keep").map((o) => o.text).join("");
    expect(kept).toContain("你好");
    expect(kept).toContain("世界");
  });

  it("超长文本返回 null", () => {
    const a = "x".repeat(3000);
    const b = "y".repeat(3000);
    expect(diffText(a, b)).toBeNull();
  });
});