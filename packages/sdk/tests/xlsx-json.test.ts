import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { spreadsheetToJsonLines } from "../src/vfs/xlsx-json";

function workbookBytes() {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([["指标", "H1", "H2"], ["公平", 28.6, 29]]);
  sheet.C2.f = "B2+0.4";
  sheet.C2.z = "0.0";
  XLSX.utils.book_append_sheet(workbook, sheet, "数据");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as Uint8Array;
}

describe("styled Excel JSONL export", () => {
  it("preserves source addresses, values and formulas with explicit limitations", () => {
    const rows = spreadsheetToJsonLines(workbookBytes(), "数据", "A1:C2")
      .split("\n").map((line) => JSON.parse(line));
    expect(rows[0].warnings.join(" ")).toContain("Conditional formatting");
    expect(rows.find((row) => row.address === "B2")).toMatchObject({ value: 28.6, sheet: "数据" });
    expect(rows.find((row) => row.address === "C2")).toMatchObject({ formula: "B2+0.4", numberFormat: "0.0" });
  });

  it("supports a sheet index and bounds processing", () => {
    expect(spreadsheetToJsonLines(workbookBytes(), "0", "B2")).toContain('"address":"B2"');
    expect(() => spreadsheetToJsonLines(workbookBytes(), "missing")).toThrow("Sheet not found");
    expect(() => spreadsheetToJsonLines(workbookBytes(), "0", "A1:XFD1048576")).toThrow("10000");
    expect(() => spreadsheetToJsonLines(workbookBytes(), "0", "C3:A1")).toThrow();
    expect(() => spreadsheetToJsonLines(workbookBytes(), "0", "A:A")).toThrow("Range");
  });
});
