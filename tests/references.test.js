import { describe, expect, it } from "vitest";
import { matchPlaceholderToBibliography } from "../src/utils/references.js";

describe("matchPlaceholderToBibliography", () => {
  it("scores bibliography entries by author and year", () => {
    const matches = matchPlaceholderToBibliography("[Smith, 2024]", [
      { id: 1, text: "Smith J. A strong paper. 2024.", year: "2024", coreAuthor: "smith" },
      { id: 2, text: "Jones P. Another paper. 2020.", year: "2020", coreAuthor: "jones" },
    ]);

    expect(matches[0]).toMatchObject({ id: 1, score: 100 });
  });
});

