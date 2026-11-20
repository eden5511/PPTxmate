// @vitest-environment happy-dom
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { centeredAxisBounds, sharedAxisBounds } from "../src/lib/pptx/chart-axis";
import { chartSnapshot, validateChartEdits } from "../src/lib/pptx/chart-validation";
import { compactPowerPointContext, routePowerPointRequest } from "../src/lib/request-router";

const ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const series = (index: number) =>
  '<c:ser><c:idx val="' + index + '"/><c:order val="' + index +
  '"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>-3.5</c:v></c:pt></c:numLit></c:val></c:ser>';

function chartZip(body = series(0)) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml",
    '<Types><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
  zip.file("ppt/charts/chart1.xml", '<c:chartSpace xmlns:c="' + ns +
    '"><c:chart><c:plotArea><c:barChart>' + body +
    '</c:barChart></c:plotArea></c:chart></c:chartSpace>');
  return zip;
}

describe("Excel to editable PPT workflow", () => {
  it("routes data requests to chart-capable tools", () => {
    expect(routePowerPointRequest("读取Excel并填入PPT")).toBe("layout");
    expect(routePowerPointRequest("根据xlsx检查柱形图数据缺失")).toBe("layout");
  });

  it("uses a shared finite axis including zero and centers line data", () => {
    expect(sharedAxisBounds([-21, 4, -3, 12])).toEqual({ min: -25, max: 15 });
    expect(sharedAxisBounds([12, 18])).toEqual({ min: -5, max: 20 });
    expect(centeredAxisBounds([28, 29, 30, 31])).toEqual({ min: 19.5, max: 39.5 });
    expect(() => sharedAxisBounds([Number.NaN])).toThrow();
    expect(() => centeredAxisBounds([])).toThrow();
  });

  it("marks compressed reads and preserves the newest returned data", () => {
    const messages = [
      { role: "user", content: "fill charts" },
      { role: "assistant", content: [
        { type: "toolCall", id: "old", name: "bash", arguments: { command: "X".repeat(10000) } },
      ] },
      { role: "toolResult", toolCallId: "old", content: [{ type: "text", text: "converted sheet0.csv" }] },
      { role: "assistant", content: [
        { type: "toolCall", id: "new", name: "read", arguments: { path: "sheet0.csv" } },
      ] },
      { role: "toolResult", toolCallId: "new", content: [{ type: "text", text: "indicator,H1,H2\nfairness,28.6,29.0" }] },
    ];
    const original = JSON.stringify(messages);
    const compacted = compactPowerPointContext(messages, { maxChars: 1800 });
    const old = compacted.find((message) => message.toolCallId === "old");
    expect(JSON.parse((old?.content as Array<{ text: string }>)[0].text)).toMatchObject({
      compacted: true, payloadOmitted: true, outputPreview: "converted sheet0.csv",
    });
    expect(compacted.at(-1)?.content).toEqual(messages.at(-1)?.content);
    expect(JSON.stringify(messages)).toBe(original);
    expect(JSON.stringify(compacted).length).toBeLessThanOrEqual(1800);
  });

  it("preserves error diagnostics when compressing tool history", () => {
    const messages = [
      { role: "user", content: "fill charts" },
      { role: "assistant", content: [{ type: "toolCall", id: "a", name: "read", arguments: { path: "x".repeat(10000) } }] },
      { role: "toolResult", toolCallId: "a", isError: true, content: [{ type: "text", text: '{"success":false,"error":"file missing"}' }] },
      { role: "assistant", content: [{ type: "toolCall", id: "b", name: "bash", arguments: { command: "echo ok" } }] },
      { role: "toolResult", toolCallId: "b", content: [{ type: "text", text: "ok" }] },
    ];
    const result = compactPowerPointContext(messages, { maxChars: 1800 });
    const error = result.find((message) => message.toolCallId === "a");
    expect(JSON.parse((error?.content as Array<{ text: string }>)[0].text)).toMatchObject({
      success: false, isError: true, error: "file missing", compacted: true,
    });
  });

  it("accepts a valid chart and rejects duplicate indices before import", async () => {
    await expect(validateChartEdits(chartZip(), new Map())).resolves.toBeUndefined();
    await expect(validateChartEdits(chartZip(series(0) + series(0)), new Map())).rejects.toThrow("duplicate series");
  });

  it("ignores chart style/color parts while retaining custom chart filenames", async () => {
    const zip = chartZip();
    zip.file("ppt/charts/style1.xml",
      '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="102"/>');
    zip.file("ppt/charts/colors1.xml",
      '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="10"/>');
    const original = await zip.file("ppt/charts/chart1.xml")!.async("string");
    zip.remove("ppt/charts/chart1.xml");
    zip.file("ppt/charts/sales.xml", original);
    zip.file("[Content_Types].xml",
      '<Types><Override PartName="/ppt/charts/sales.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>');
    expect([...((await chartSnapshot(zip)).keys())]).toEqual(["ppt/charts/sales.xml"]);
    await expect(validateChartEdits(zip, new Map())).resolves.toBeUndefined();
    zip.file("ppt/charts/sales.xml", original.replace("</c:barChart>", series(0) + "</c:barChart>"));
    await expect(validateChartEdits(zip, new Map())).rejects.toThrow("duplicate series");
  });

  it("rejects late dPt and out-of-range cached points", async () => {
    await expect(validateChartEdits(chartZip(series(0).replace("</c:ser>", '<c:dPt><c:idx val="0"/></c:dPt></c:ser>')), new Map())).rejects.toThrow("dPt");
    await expect(validateChartEdits(chartZip(series(0).replace('ptCount val="1"', 'ptCount val="0"')), new Map())).rejects.toThrow("cache point");
  });

  it("does not mistake sparse blank caches for corruption", async () => {
    await expect(validateChartEdits(chartZip(series(0).replace('ptCount val="1"', 'ptCount val="3"')), new Map())).resolves.toBeUndefined();
  });

  it("requires an editable workbook when requested and detects its removal", async () => {
    const zip = chartZip();
    await expect(validateChartEdits(zip, new Map(), true)).rejects.toThrow("embedded workbook");
    const xml = await zip.file("ppt/charts/chart1.xml")!.async("string");
    zip.file("ppt/charts/chart1.xml", xml.replace("</c:chartSpace>",
      '<c:externalData xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></c:chartSpace>'));
    const before = await chartSnapshot(zip);
    zip.file("ppt/charts/chart1.xml", xml);
    await expect(validateChartEdits(zip, before)).rejects.toThrow("embedded workbook");
  });
});
