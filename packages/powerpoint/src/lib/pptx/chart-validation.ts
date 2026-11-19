import JSZip from "jszip";

const CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(xml: string, part: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error(`CHART_VALIDATION_FAILED: invalid XML in ${part}`);
  }
  return doc;
}

function fail(part: string, message: string): never {
  throw new Error(`CHART_VALIDATION_FAILED: ${part}: ${message}`);
}

function resolvePart(source: string, target: string): string {
  const parts = target.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const segment of target.split("/")) {
    if (segment === "..") {
      if (!parts.length) fail(source, "relationship escapes package");
      parts.pop();
    } else if (segment && segment !== ".") {
      parts.push(segment);
    }
  }
  return parts.join("/");
}

export async function chartSnapshot(zip: JSZip): Promise<Map<string, string>> {
  const snapshot = new Map<string, string>();
  for (const name of Object.keys(zip.files)) {
    if (/^ppt\/charts\/[^/]+\.xml$/.test(name)) {
      snapshot.set(name, await zip.files[name].async("string"));
    }
  }
  return snapshot;
}

export async function validateChartEdits(
  zip: JSZip,
  before: Map<string, string>,
  requireEditableData = false,
): Promise<void> {
  const after = await chartSnapshot(zip);
  const contentTypes = zip.file("[Content_Types].xml");
  const types = contentTypes
    ? parseXml(await contentTypes.async("string"), "[Content_Types].xml")
    : null;
  for (const [name, xml] of after) {
    const doc = parseXml(xml, name);
    const changed = before.get(name) !== xml;
    const externalData = doc.getElementsByTagNameNS(CHART_NS, "externalData")[0];
    const previouslyEmbedded =
      before.has(name) &&
      parseXml(before.get(name) as string, name).getElementsByTagNameNS(
        CHART_NS,
        "externalData",
      ).length > 0;
    if ((previouslyEmbedded || requireEditableData) && !externalData) {
      fail(name, "editable chart lost its embedded workbook reference");
    }
    if (changed) {
      const registered = types
        ? Array.from(types.getElementsByTagName("Override")).some(
            (entry) =>
              entry.getAttribute("PartName") === `/${name}` &&
              entry.getAttribute("ContentType") ===
                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
          )
        : false;
      if (!registered) fail(name, "missing chart content type");
      const indices = new Set<string>();
      const orders = new Map<Element, Set<string>>();
      for (const series of Array.from(doc.getElementsByTagNameNS(CHART_NS, "ser"))) {
        const children = Array.from(series.children);
        for (const key of ["idx", "order"]) {
          const value = children
            .find((child) => child.localName === key)
            ?.getAttribute("val");
          if (!value || !/^\d+$/.test(value)) fail(name, `missing/invalid ${key}`);
          const normalized = String(Number(value));
          const parent = series.parentElement as Element;
          const seen = key === "idx" ? indices : (orders.get(parent) ?? new Set<string>());
          if (seen.has(normalized)) fail(name, `duplicate series ${key}=${value}`);
          seen.add(normalized);
          if (key === "order") orders.set(parent, seen);
        }
        const firstData = children.findIndex((child) =>
          ["dLbls", "cat", "val", "xVal", "yVal", "bubbleSize"].includes(child.localName),
        );
        if (
          firstData >= 0 &&
          children.some((child, index) => child.localName === "dPt" && index > firstData)
        ) {
          fail(name, "dPt must precede labels and category/value data");
        }
      }
      for (const cacheName of ["numCache", "strCache", "numLit", "strLit"]) {
        for (const cache of Array.from(doc.getElementsByTagNameNS(CHART_NS, cacheName))) {
          const children = Array.from(cache.children);
          const countText = children.find((child) => child.localName === "ptCount")?.getAttribute("val");
          if (countText === null || countText === undefined || !/^\d+$/.test(countText)) {
            fail(name, "invalid cache ptCount");
          }
          const count = Number(countText);
          const seen = new Set<number>();
          for (const point of children.filter((child) => child.localName === "pt")) {
            const indexText = point.getAttribute("idx");
            const index = Number(indexText);
            if (!indexText || !/^\d+$/.test(indexText) || index >= count || seen.has(index)) {
              fail(name, "duplicate/out-of-range cache point");
            }
            seen.add(index);
            const value = point.getElementsByTagNameNS(CHART_NS, "v")[0]?.textContent;
            if (
              cacheName.startsWith("num") &&
              (value === undefined || value === null || value.trim() === "" || !Number.isFinite(Number(value)))
            ) {
              fail(name, "non-numeric chart cache value");
            }
          }
        }
      }
    }
    // Check even unchanged chart XML: a callback can delete its workbook/rels.
    if (externalData) {
      const id = externalData.getAttributeNS(OFFICE_REL_NS, "id");
      const relPath = name.replace(/([^/]+)$/, "_rels/$1.rels");
      const relFile = zip.file(relPath);
      if (!id || !relFile) fail(name, "missing embedded workbook relationship");
      const relDoc = parseXml(await relFile.async("string"), relPath);
      const relation = Array.from(relDoc.getElementsByTagNameNS(REL_NS, "Relationship")).find(
        (entry) => entry.getAttribute("Id") === id,
      );
      if (!relation || relation.getAttribute("TargetMode") === "External") {
        fail(name, "embedded workbook must have an internal relationship");
      }
      const target = relation.getAttribute("Target");
      if (!target) fail(name, "missing workbook target");
      const workbook = zip.file(resolvePart(name, target));
      if (!workbook) fail(name, "embedded workbook part does not exist");
      if (changed || requireEditableData) {
        const book = await JSZip.loadAsync(await workbook.async("uint8array"));
        if (!book.file("xl/workbook.xml") || !book.file("[Content_Types].xml")) {
          fail(name, "embedded XLSX package is invalid");
        }
      }
    }
  }
}
