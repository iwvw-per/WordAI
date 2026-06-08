import { describe, expect, it } from "vitest";
import { describeModelRoute } from "../src/utils/modelRouting.js";

describe("describeModelRoute", () => {
  it("uses base model when routing is disabled", () => {
    expect(describeModelRoute({ enabled: false, baseModel: "base", actionName: "polish" })).toEqual({
      model: "base",
      route: "base",
    });
  });

  it("routes academic-style work to the quality model", () => {
    expect(
      describeModelRoute({
        enabled: true,
        baseModel: "base",
        fastModel: "fast",
        qualityModel: "quality",
        actionName: "abstract",
      }),
    ).toEqual({ model: "quality", route: "quality" });
  });

  it("routes ordinary work to the fast model", () => {
    expect(
      describeModelRoute({
        enabled: true,
        baseModel: "base",
        fastModel: "fast",
        qualityModel: "quality",
        actionName: "polish",
      }),
    ).toEqual({ model: "fast", route: "fast" });
  });
});

