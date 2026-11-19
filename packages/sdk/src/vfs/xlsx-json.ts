import * as XLSX from "xlsx";

interface FillColor {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
}

function colorInfo(color: FillColor | undefined) {
  if (!color) return null;
  const rgb = color.rgb?.replace(/^#/, "");
  return {
    ...color,
    resolvedRgb:
      rgb && /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(rgb)
        ? rgb.slice(-6).toUpperCase()
        : null,
  };
}

export function spreadsheetToJsonLines(
  data: Uint8Array,
  sheetArg?: string,
  rangeArg?: string,
): string {
  const workbook = XLSX.read(data, {
    type: "array",
    cellStyles: true,
    cellNF: true,
    cellFormula: true,
    sheetStubs: true,
  });
  const sheetName =
    sheetArg === undefined
      ? workbook.SheetNames[0]
      : workbook.SheetNames.includes(sheetArg)
        ? sheetArg
        : /^\d+$/.test(sheetArg)
          ? workbook.SheetNames[Number(sheetArg)]
          : undefined;
  if (!sheetName) {
    throw new Error(
      `Sheet not found. Available: ${workbook.SheetNames.join(", ")}`,
    );
  }
  const sheet = workbook.Sheets[sheetName];
  const rangeText = rangeArg ?? sheet["!ref"] ?? "A1";
  if (!/^[A-Z]+[1-9]\d*(?::[A-Z]+[1-9]\d*)?$/i.test(rangeText)) {
    throw new Error("Range must be an A1 cell or range, for example A1:Q70.");
  }
  const range = XLSX.utils.decode_range(rangeText.toUpperCase());
  const count = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
  if (
    range.e.r < range.s.r ||
    range.e.c < range.s.c ||
    range.e.r >= 1_048_576 ||
    range.e.c >= 16_384 ||
    count > 10_000
  ) {
    throw new Error("Read at most 10000 cells per call using a smaller range.");
  }
  const lines: string[] = [
    JSON.stringify({
      type: "sheet",
      sheet: sheetName,
      sheets: workbook.SheetNames,
      range: rangeText,
      merges: sheet["!merges"] ?? [],
      warnings: [
        "Conditional formatting is not evaluated; fills are stored cell styles.",
        "Formula values are cached, not recalculated.",
        "If resolvedRgb is null, do not infer a color from the theme/index alone.",
      ],
    }),
  ];
  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address];
      if (!cell) continue;
      const style = cell.s as
        | {
            patternType?: string;
            fgColor?: FillColor;
            bgColor?: FillColor;
          }
        | undefined;
      lines.push(
        JSON.stringify({
          type: "cell",
          sheet: sheetName,
          address,
          value: cell.v ?? null,
          display: cell.w ?? XLSX.utils.format_cell(cell),
          formula: cell.f ?? null,
          numberFormat: cell.z ?? null,
          fill: style
            ? {
                patternType: style.patternType ?? null,
                foreground: colorInfo(style.fgColor),
                background: colorInfo(style.bgColor),
              }
            : null,
        }),
      );
    }
  }
  return lines.join("\n");
}
