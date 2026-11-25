import { describe, expect, it } from "vitest";
import { presentColumnFreeze, restoreColumnFreeze } from "./columnFreeze";

describe("column freeze", () => {
  it("drops a frozen column for the view and puts it back on save", () => {
    const freeze = { xSplit: 1, ySplit: 7, startRow: 7, startColumn: 1 };
    const book = {
      sheets: {
        a: { name: "Cont DI", freeze },
        b: { name: "Slidesheet", freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } },
      },
    };
    const { view, kept } = presentColumnFreeze(book);
    expect(freeze.xSplit).toBe(1);
    expect(view.sheets.a.freeze).toEqual({ xSplit: 0, ySplit: 7, startRow: 7, startColumn: -1 });
    expect(view.sheets.b).toBe(book.sheets.b);
    const saved = {
      sheets: {
        renamed: { name: "Cont DI", freeze: { xSplit: 0, ySplit: 7, startRow: 7, startColumn: 0 } },
        b: book.sheets.b,
      },
    };
    expect(restoreColumnFreeze(saved, kept).sheets.renamed.freeze).toEqual(freeze);
    expect(restoreColumnFreeze(saved, kept).sheets.b.freeze.ySplit).toBe(1);
  });

  it("leaves a book with no column freeze alone", () => {
    const book = { sheets: { a: { name: "Data", freeze: { xSplit: 0, ySplit: 2, startRow: 2, startColumn: 0 } } } };
    const { view, kept } = presentColumnFreeze(book);
    expect(view).toBe(book);
    expect(kept).toEqual([]);
    expect(restoreColumnFreeze(book, kept)).toBe(book);
  });
});
