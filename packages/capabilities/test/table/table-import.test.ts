import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  cell,
  deck,
  font,
  paragraph,
  row,
  slide,
  slideNumber,
  span,
  table,
  type CellNode,
  type RowNode,
  type TableNode,
  type TableRowStyle,
} from "@pptxdk/capabilities/document";
import { defineTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import { tablePositiveScenario } from "./scenarios.ts";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
    ),
  },
]);

function importedSlides(doc: CommonNode): readonly CommonNode[] {
  return doc.children.filter(
    (child): child is CommonNode =>
      child instanceof CommonNode && child.type === "slide",
  );
}

function findFirstTable(slideNode: CommonNode): TableNode {
  const found = slideNode.children.find(
    (child): child is TableNode =>
      child instanceof CommonNode && child.type === "table",
  );
  assert.ok(found !== undefined, "Expected imported TableNode on slide");
  return found;
}

describe("native table import and editing roundtrip", () => {
  it("reconstructs exact EMU column widths, row heights, borders, fills, padding, and SlideNumber across buffer release", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "tbl-slide" }).append(
        table({
          id: "tbl",
          name: "MetricsGrid",
          x: 24,
          y: 36,
          width: 480,
          columns: [110, "auto", "1fr"],
          style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
        }).append(
          row({
            id: "hdr",
            style: { height: 38, backgroundColor: "#1E293B", color: "#FFFFFF" },
          }).append(
            cell(paragraph("Overview", { style: { align: "center" } }), {
              colSpan: 2,
              style: {
                fontWeight: "bold",
                padding: [6, 10],
                borderColor: "#0F172A",
                borderWidth: 2,
                borderStyle: "solid",
              },
            }),
            cell(
              paragraph(["Page ", slideNumber()], {
                style: { align: "right" },
              }),
              {
                style: {
                  backgroundColor: "#334155",
                  padding: [6, 10],
                  borderColor: "#0F172A",
                  borderWidth: 1,
                  borderStyle: "dashed",
                },
              },
            ),
          ),
          row({ id: "body" }).append(
            cell("AutoCol", {
              style: {
                verticalAlign: "top",
                padding: [4, 8, 6, 12],
                borderColor: "#2563EB",
                borderWidth: 2,
                borderStyle: "dotted",
              },
            }),
            cell("Metric", {
              style: {
                backgroundColor: "#EFF6FF",
                verticalAlign: "middle",
                padding: 6,
              },
            }),
            cell(
              paragraph([
                span("$99", {
                  style: { fontWeight: "bold", color: "#047857" },
                }),
                " / mo",
              ]),
              { style: { verticalAlign: "bottom", padding: 6 } },
            ),
          ),
        ),
      ),
    );

    const initialResolution = resolveDocument(doc);
    const initialEmitted = emit(initialResolution, { embedFonts: "full" });
    const initialValid = await validatePptx(initialEmitted.pptx);
    const initialObs = initialValid.slides[0]!.objects[0]!;
    const detachedBytes = new Uint8Array(initialEmitted.pptx);

    const imported = importDocument(detachedBytes);
    detachedBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedSlide = importedSlides(imported.document)[0]!;
    const importedTable = findFirstTable(importedSlide);
    assert.equal(importedTable.props.name, "MetricsGrid");
    assert.equal(importedTable.props.x, 24);
    assert.equal(importedTable.props.y, 36);
    assert.deepEqual(
      importedTable.props.columns,
      initialObs.table!.columnWidthsEmu.map((w) => w / 9525),
    );

    const importedRows = importedTable.children as readonly RowNode[];
    assert.equal(importedRows.length, 2);
    assert.equal(
      (importedRows[0]!.props["style"] as TableRowStyle | undefined)?.height,
      initialObs.table!.rowHeightsEmu[0]! / 9525,
    );
    assert.equal(
      (importedRows[1]!.props["style"] as TableRowStyle | undefined)?.height,
      initialObs.table!.rowHeightsEmu[1]! / 9525,
    );

    const reemittedResolution = resolveDocument(imported.document);
    const reemitted = emit(reemittedResolution, { embedFonts: "full" });
    const painted = paint(reemittedResolution);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);

    const roundtripValid = await validatePptx(reemitted.pptx);
    assert.deepEqual(roundtripValid.errors, []);

    const rtObs = roundtripValid.slides[0]!.objects[0]!;
    assert.equal(rtObs.kind, "table");
    assert.deepEqual(
      rtObs.table?.columnWidthsEmu,
      initialObs.table?.columnWidthsEmu,
    );
    assert.deepEqual(
      rtObs.table?.rowHeightsEmu,
      initialObs.table?.rowHeightsEmu,
    );
    assert.deepEqual(rtObs.table?.merges, initialObs.table?.merges);
    assert.deepEqual(rtObs.table?.cells, initialObs.table?.cells);
    assert.equal(
      rtObs.table?.cellTexts[1]?.text.paragraphs[0]?.runs[1]?.field,
      "slidenum",
    );
  });

  it("reconstructs 2x2 rectangular merges and external hMerge/vMerge-only continuation grids", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "merge-grid" }).append(
        table({
          id: "rect-merge",
          x: 20,
          y: 20,
          columns: [80, 80, 80],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ style: { height: 30 } }).append(
            cell("2x2 Anchor", { colSpan: 2, rowSpan: 2 }),
            cell("TopRight"),
          ),
          row({ style: { height: 30 } }).append(cell("MidRight")),
          row({ style: { height: 30 } }).append(
            cell("B0"),
            cell("B1"),
            cell("B2"),
          ),
        ),
      ),
    );

    const emitted = emit(resolveDocument(doc));
    const parts = unzipSync(emitted.pptx);
    // Simulate external PowerPoint/POI output where anchor omits gridSpan/rowSpan and relies on hMerge/vMerge.
    const slidePath = "ppt/slides/slide1.xml";
    const externalSlideXml = strFromU8(parts[slidePath]!)
      .replace('<a:tc gridSpan="2" rowSpan="2">', "<a:tc>")
      .replace('<a:tc hMerge="1" rowSpan="2">', '<a:tc hMerge="1">')
      .replace('<a:tc vMerge="1" gridSpan="2">', '<a:tc vMerge="1">');
    parts[slidePath] = strToU8(externalSlideXml);

    const imported = importDocument(zipSync(parts));
    assert.deepEqual(imported.report.diagnostics, []);

    const roundtrip = emit(resolveDocument(imported.document));
    const valid = await validatePptx(roundtrip.pptx);
    assert.deepEqual(valid.errors, []);
    assert.deepEqual(valid.slides[0]?.objects[0]?.table?.merges, [
      { row: 0, column: 0, rowSpan: 2, colSpan: 2 },
    ]);
    assert.deepEqual(valid.slides[0]?.objects[0]?.table?.cells, [
      ["2x2 Anchor", "", "TopRight"],
      ["", "", "MidRight"],
      ["B0", "B1", "B2"],
    ]);
  });

  it("imports dual-entry table scenario, edits one cell text and character format, and preserves merges, tracks, and adjacent cells", async () => {
    const { runs } = await runCapabilityScenario(tablePositiveScenario);
    const initialPptx = runs[0].pptx;
    const initialValid = runs[0].validation;
    const initialTable = initialValid.slides[0]!.objects[0]!.table!;

    const imported = importDocument(new Uint8Array(initialPptx));
    assert.deepEqual(
      imported.report.diagnostics.filter(
        (diag) => !(diag.code === "unconsumed-xml" && diag.localName === "bg"),
      ),
      [],
    );

    const importedSlide = importedSlides(imported.document)[0]!;
    const tbl = findFirstTable(importedSlide);
    const rows = tbl.children as readonly RowNode[];
    // Edit row 2, column 1 ("Auto Width" cell) text and character formatting.
    const targetCell = rows[1]!.children[1] as CellNode;
    targetCell.content("Edited Metric");
    targetCell.style({
      fontFamily: jostFont,
      color: "#DC2626",
      fontWeight: "bold",
      fontSize: "16pt",
    });

    // Attach Jost* for remaining unembedded cells on page 1 so preview renders PNG cleanly.
    imported.document.set({
      theme: defineTheme({ fonts: { major: jostFont, minor: jostFont } }),
    });
    const attachFonts = (node: CommonNode): void => {
      const style = node.props["style"] as
        | Readonly<Record<string, unknown>>
        | undefined;
      if (style?.["fontFamily"] === "Jost*") {
        (
          node as CommonNode<
            string,
            Record<string, unknown>,
            Record<string, unknown>
          >
        ).style({ fontFamily: jostFont });
      } else if (style?.["fontFamily"] === "Noto Sans CJK SC") {
        (
          node as CommonNode<
            string,
            Record<string, unknown>,
            Record<string, unknown>
          >
        ).style({ fontFamily: notoFont });
      }
      for (const child of node.children) {
        if (child instanceof CommonNode) attachFonts(child);
      }
    };
    attachFonts(importedSlide);

    const editedResolution = resolveDocument(imported.document);
    const editedEmitted = emit(editedResolution);
    const editedPainted = paint(editedResolution, { pages: [1] });
    assert.deepEqual(editedPainted.pages[0]?.diagnostics, []);
    assert.ok(editedPainted.pages[0]?.png instanceof Uint8Array);

    const editedValid = await validatePptx(editedEmitted.pptx);
    assert.deepEqual(editedValid.errors, []);
    const editedObs = editedValid.slides[0]!.objects[0]!;
    assert.equal(editedObs.kind, "table");
    assert.deepEqual(editedObs.table?.merges, initialTable.merges);
    assert.deepEqual(
      editedObs.table?.columnWidthsEmu,
      initialTable.columnWidthsEmu,
    );
    assert.deepEqual(
      editedObs.table?.rowHeightsEmu,
      initialTable.rowHeightsEmu,
    );
    assert.equal(editedObs.table?.cells[1]?.[1], "Edited Metric");
    assert.equal(editedObs.table?.cells[0]?.[0], "Q1 Plan Overview");
    assert.equal(editedObs.table?.cells[1]?.[0], "Pro\nTier");
    assert.equal(editedObs.table?.cells[1]?.[2], "$49 / seat");
    assert.equal(editedObs.table?.cells[2]?.[1], "Summary");
    assert.equal(editedObs.table?.cells[2]?.[2], "Includes SSO & audit");
    const editedRun = editedObs.table?.cellTexts.find(
      (entry) => entry.row === 1 && entry.column === 1,
    )?.text.paragraphs[0]?.runs[0];
    assert.equal(editedRun?.color, "#DC2626");
    assert.equal(editedRun?.bold, true);
    assert.equal(editedRun?.fontSizePt, 16);
  });

  it("imports external ph-populated-placeholders.pptx slide3 native table body and edits a cell cleanly", async () => {
    const fixtureBytes = new Uint8Array(
      readFileSync(
        new URL("../fixtures/ph-populated-placeholders.pptx", import.meta.url),
      ),
    );
    const imported = importDocument(fixtureBytes);
    fixtureBytes.fill(0);

    const slides = importedSlides(imported.document);
    const slide3 = slides[2]!;
    const tableNode = findFirstTable(slide3);
    assert.equal(tableNode.props.name, "Table Placeholder 2");
    assert.deepEqual(tableNode.props.columns, [192, 192]);
    const rows = tableNode.children as readonly RowNode[];
    assert.equal(rows.length, 3);
    for (const r of rows) {
      assert.equal(
        (r.props["style"] as TableRowStyle | undefined)?.height,
        370_840 / 9525,
      );
    }

    const firstCell = rows[0]!.children[0] as CellNode;
    firstCell.content("Imported Cell A1");
    firstCell.style({
      fontFamily: jostFont,
      fontSize: "14pt",
      color: "#1E293B",
    });

    const resolution = resolveDocument(imported.document);
    const emitted = emit(resolution);
    const valid = await validatePptx(emitted.pptx);
    assert.deepEqual(valid.errors, []);
    const validTable = valid.slides[2]!.objects.find(
      (obj) => obj.kind === "table",
    );
    assert.ok(validTable?.table !== undefined);
    assert.deepEqual(validTable.table.columnWidthsEmu, [1_828_800, 1_828_800]);
    assert.deepEqual(
      validTable.table.rowHeightsEmu,
      [370_840, 370_840, 370_840],
    );
    assert.equal(validTable.table.cells[0]?.[0], "Imported Cell A1");
  });
});
