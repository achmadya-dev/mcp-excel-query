import { describe, expect, it } from "@jest/globals";
import {
  formatDateValue,
  normalizeExcelNumFmt,
  parseDateValue,
} from "../date.js";

describe("date helpers", () => {
  it("parses ISO and human-readable date strings", () => {
    expect(parseDateValue("2026-06-08T00:00:00.000Z")?.toISOString()).toBe(
      "2026-06-08T00:00:00.000Z"
    );
    expect(parseDateValue("09 Jun 2026")?.getDate()).toBe(9);
    expect(parseDateValue("08 Jun 26")?.getFullYear()).toBe(2026);
    expect(parseDateValue("08/06/2026")?.getMonth()).toBe(5);
  });

  it("formats dates with Excel-style tokens", () => {
    const date = new Date(2026, 5, 9);
    expect(formatDateValue(date, "dd mmm yyyy")).toBe("09 Jun 2026");
    expect(formatDateValue(date, "dd mmm yy")).toBe("09 Jun 26");
  });

  it("normalizes Excel numFmt escaped quotes", () => {
    expect(normalizeExcelNumFmt('dd" "mmm" "yy')).toBe("dd mmm yy");
  });
});
