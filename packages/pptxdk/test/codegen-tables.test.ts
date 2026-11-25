import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { build, codegen, importPptx } from "pptxdk";
import {
  Cell,
  Deck,
  P,
  Row,
  Slide,
  SlideLayout,
  SlideMaster,
  SlideNumber,
  Span,
  Table,
  Text,
  defineSlideLayout,
  expandDocument,
} from "pptxdk/react";
import {
  cell,
  deck,
  font,
  paragraph,
  row,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  table,
  text,
} from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  previewGeneratedSourceRelocated,
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

const emuPerPx = 9525;

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);
const zcoolBytes = new Uint8Array(
  readFileSync(
    new URL("./assets/fonts/ZCOOLXiaoWei_400Regular.ttf", import.meta.url),
  ),
);

const jostFont = font("Jost*", [
  { weight: 400, style: "normal", data: jostBookBytes },
  { weight: 700, style: "normal", data: jostBoldBytes },
]);

const notoFont = font("ZCOOL XiaoWei", [
  { weight: 400, style: "normal", data: zcoolBytes },
]);

const wrappingCopy =
  "Shared layout wraps long text cleanly inside constrained table cells.";

const tablePlaceholderLayout = defineSlideLayout({
  id: "tbl-layout",
  name: "Table Layout",
  placeholders: {
    title: { type: "title", x: 32, y: 20, width: 576, height: 40 },
    tbl: {
      type: "tbl",
      explicitType: true,
      idx: 1,
      x: 32,
      y: 76,
      width: 576,
      height: 220,
    },
  },
});

function samplePngPixel(png: PNG, pxX: number, pxY: number): readonly number[] {
  const idx = (Math.round(pxY) * png.width + Math.round(pxX)) * 4;
  return [...png.data.subarray(idx, idx + 4)];
}

function createTableAcceptanceVanillaDeck() {
  return deck({ width: 640, height: 360 }).append(
    slideMaster({ id: "main-master", name: "Main" }).append(
      slideLayout(tablePlaceholderLayout),
    ),
    slide({ id: "pricing-merges" }).append(
      table({
        id: "pricing-table",
        name: "PricingTable",
        x: 32,
        y: 36,
        width: 576,
        columns: [130, "auto", "1fr"],
        style: { fontFamily: jostFont, fontSize: 15, color: "#0F172A" },
      }).append(
        row({
          id: "header",
          style: { height: 46, backgroundColor: "#1E293B", color: "#FFFFFF" },
        }).append(
          cell(paragraph("Q1 Plan Overview", { style: { align: "center" } }), {
            id: "title-span",
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
            paragraph(["第 ", slideNumber(), " 页"], {
              style: { align: "right" },
            }),
            {
              id: "page-badge",
              style: {
                fontFamily: notoFont,
                backgroundColor: "#334155",
                color: "#E2E8F0",
                padding: [6, 10],
                borderColor: "#0F172A",
                borderWidth: 2,
                borderStyle: "dashed",
              },
            },
          ),
        ),
        row({ id: "tier-pro", style: { backgroundColor: "#FFFFFF" } }).append(
          cell(paragraph("Pro\nTier", { style: { align: "center" } }), {
            id: "tier-name",
            rowSpan: 2,
            style: {
              backgroundColor: "#EFF6FF",
              color: "#1D4ED8",
              fontWeight: "bold",
              verticalAlign: "middle",
              padding: [6, 10],
              borderColor: "#3B82F6",
              borderWidth: 2,
              borderStyle: "solid",
            },
          }),
          cell("Auto Width", {
            id: "pro-metric",
            style: {
              padding: [6, 10],
              borderColor: "#CBD5E1",
              borderWidth: 1,
              borderStyle: "dotted",
            },
          }),
          cell(
            paragraph(
              [
                span("$49", {
                  style: { fontWeight: "bold", color: "#047857" },
                }),
                " / seat",
              ],
              { style: { align: "right" } },
            ),
            {
              id: "pro-price",
              style: {
                backgroundColor: "#F1F5F9",
                verticalAlign: "bottom",
                padding: [6, 10],
                borderColor: "#CBD5E1",
                borderWidth: 1,
              },
            },
          ),
        ),
        row({ id: "tier-notes", style: { backgroundColor: "#FFFFFF" } }).append(
          cell("Summary", {
            id: "notes-label",
            style: {
              padding: [6, 10],
              borderColor: "#CBD5E1",
              borderWidth: 1,
            },
          }),
          cell("Includes SSO & audit", {
            id: "notes-detail",
            style: {
              backgroundColor: "#FEF3C7",
              color: "#92400E",
              verticalAlign: "top",
              padding: [6, 10],
              borderColor: "#D97706",
              borderWidth: 1,
            },
          }),
        ),
      ),
    ),
    slide({ id: "narrow-wrapping" }).append(
      table({
        id: "wide-table",
        x: 28,
        y: 44,
        columns: [270],
        style: { fontFamily: jostFont, fontSize: 15, color: "#065F46" },
      }).append(
        row({ id: "wide-row" }).append(
          cell(wrappingCopy, {
            id: "wide-cell",
            style: {
              padding: [10, 12],
              backgroundColor: "#ECFDF5",
              borderColor: "#10B981",
              borderWidth: 2,
            },
          }),
        ),
      ),
      table({
        id: "narrow-table",
        x: 332,
        y: 44,
        columns: [130],
        style: { fontFamily: jostFont, fontSize: 15, color: "#991B1B" },
      }).append(
        row({ id: "narrow-row" }).append(
          cell(wrappingCopy, {
            id: "narrow-cell",
            style: {
              padding: [10, 12],
              backgroundColor: "#FEF2F2",
              borderColor: "#EF4444",
              borderWidth: 2,
            },
          }),
        ),
      ),
    ),
    slide({ id: "placeholder-table", layout: tablePlaceholderLayout }).append(
      text("Quarterly Metrics", {
        placeholder: "title",
        style: { fontFamily: jostFont, fontSize: 20, fontWeight: "bold" },
      }),
      table({
        id: "ph-tbl",
        name: "PlaceholderMetrics",
        placeholder: "tbl",
        columns: ["1fr", "2fr"],
        style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
      }).append(
        row({
          style: { height: 36, backgroundColor: "#E2E8F0" },
        }).append(cell("Region"), cell("Notes")),
        row({ style: { height: 36 } }).append(
          cell("EMEA"),
          cell("Expanded coverage"),
        ),
      ),
    ),
  );
}

async function createTableAcceptanceReactDeck() {
  return expandDocument(
    createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "main-master", name: "Main" },
        createElement(SlideLayout, { layout: tablePlaceholderLayout }),
      ),
      createElement(
        Slide,
        { id: "pricing-merges" },
        createElement(
          Table,
          {
            id: "pricing-table",
            name: "PricingTable",
            x: 32,
            y: 36,
            width: 576,
            columns: [130, "auto", "1fr"],
            style: { fontFamily: jostFont, fontSize: 15, color: "#0F172A" },
          },
          createElement(
            Row,
            {
              id: "header",
              style: {
                height: 46,
                backgroundColor: "#1E293B",
                color: "#FFFFFF",
              },
            },
            createElement(
              Cell,
              {
                id: "title-span",
                colSpan: 2,
                style: {
                  fontWeight: "bold",
                  padding: [6, 10],
                  borderColor: "#0F172A",
                  borderWidth: 2,
                  borderStyle: "solid",
                },
              },
              createElement(
                P,
                { style: { align: "center" } },
                "Q1 Plan Overview",
              ),
            ),
            createElement(
              Cell,
              {
                id: "page-badge",
                style: {
                  fontFamily: notoFont,
                  backgroundColor: "#334155",
                  color: "#E2E8F0",
                  padding: [6, 10],
                  borderColor: "#0F172A",
                  borderWidth: 2,
                  borderStyle: "dashed",
                },
              },
              createElement(
                P,
                { style: { align: "right" } },
                "第 ",
                createElement(SlideNumber, null),
                " 页",
              ),
            ),
          ),
          createElement(
            Row,
            { id: "tier-pro", style: { backgroundColor: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "tier-name",
                rowSpan: 2,
                style: {
                  backgroundColor: "#EFF6FF",
                  color: "#1D4ED8",
                  fontWeight: "bold",
                  verticalAlign: "middle",
                  padding: [6, 10],
                  borderColor: "#3B82F6",
                  borderWidth: 2,
                  borderStyle: "solid",
                },
              },
              createElement(P, { style: { align: "center" } }, "Pro\nTier"),
            ),
            createElement(
              Cell,
              {
                id: "pro-metric",
                style: {
                  padding: [6, 10],
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                  borderStyle: "dotted",
                },
              },
              "Auto Width",
            ),
            createElement(
              Cell,
              {
                id: "pro-price",
                style: {
                  backgroundColor: "#F1F5F9",
                  verticalAlign: "bottom",
                  padding: [6, 10],
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                },
              },
              createElement(
                P,
                { style: { align: "right" } },
                createElement(
                  Span,
                  { style: { fontWeight: "bold", color: "#047857" } },
                  "$49",
                ),
                " / seat",
              ),
            ),
          ),
          createElement(
            Row,
            { id: "tier-notes", style: { backgroundColor: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "notes-label",
                style: {
                  padding: [6, 10],
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                },
              },
              "Summary",
            ),
            createElement(
              Cell,
              {
                id: "notes-detail",
                style: {
                  backgroundColor: "#FEF3C7",
                  color: "#92400E",
                  verticalAlign: "top",
                  padding: [6, 10],
                  borderColor: "#D97706",
                  borderWidth: 1,
                },
              },
              "Includes SSO & audit",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "narrow-wrapping" },
        createElement(
          Table,
          {
            id: "wide-table",
            x: 28,
            y: 44,
            columns: [270],
            style: { fontFamily: jostFont, fontSize: 15, color: "#065F46" },
          },
          createElement(
            Row,
            { id: "wide-row" },
            createElement(
              Cell,
              {
                id: "wide-cell",
                style: {
                  padding: [10, 12],
                  backgroundColor: "#ECFDF5",
                  borderColor: "#10B981",
                  borderWidth: 2,
                },
              },
              wrappingCopy,
            ),
          ),
        ),
        createElement(
          Table,
          {
            id: "narrow-table",
            x: 332,
            y: 44,
            columns: [130],
            style: { fontFamily: jostFont, fontSize: 15, color: "#991B1B" },
          },
          createElement(
            Row,
            { id: "narrow-row" },
            createElement(
              Cell,
              {
                id: "narrow-cell",
                style: {
                  padding: [10, 12],
                  backgroundColor: "#FEF2F2",
                  borderColor: "#EF4444",
                  borderWidth: 2,
                },
              },
              wrappingCopy,
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "placeholder-table", layout: tablePlaceholderLayout },
        createElement(
          Text,
          {
            placeholder: "title",
            style: { fontFamily: jostFont, fontSize: 20, fontWeight: "bold" },
          },
          "Quarterly Metrics",
        ),
        createElement(
          Table,
          {
            id: "ph-tbl",
            name: "PlaceholderMetrics",
            placeholder: "tbl",
            columns: ["1fr", "2fr"],
            style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
          },
          createElement(
            Row,
            { style: { height: 36, backgroundColor: "#E2E8F0" } },
            createElement(Cell, null, "Region"),
            createElement(Cell, null, "Notes"),
          ),
          createElement(
            Row,
            { style: { height: 36 } },
            createElement(Cell, null, "EMEA"),
            createElement(Cell, null, "Expanded coverage"),
          ),
        ),
      ),
    ),
  );
}

describe("pptxdk native table codegen acceptance", () => {
  it("preserves tracks, merges, borders, rich text, slide numbers, narrow wrapping, and placeholder bindings across direct and imported vanilla/react roundtrips", async () => {
    const vanillaDoc = createTableAcceptanceVanillaDeck();
    const reactDoc = await createTableAcceptanceReactDeck();

    const initialBuilt = await build(vanillaDoc, { embedFonts: "full" });
    assert.deepEqual(
      initialBuilt.report.diagnostics.filter((d) => d.severity === "error"),
      [],
    );
    const bytesCopy = new Uint8Array(initialBuilt.pptx);
    const imported = importPptx(bytesCopy);
    bytesCopy.fill(0);
    assert.deepEqual(
      imported.report.diagnostics.filter(
        (diag) => !(diag.code === "unconsumed-xml" && diag.localName === "bg"),
      ),
      [],
    );

    const sources = [
      { label: "direct-vanilla", doc: vanillaDoc },
      { label: "direct-react", doc: reactDoc },
      { label: "imported-pptx", doc: imported.document },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc }) => [
      {
        label: `${label}:vanilla`,
        result: codegen(doc, { target: "vanilla" }),
      },
      { label: `${label}:react`, result: codegen(doc, { target: "react" }) },
    ]);

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      const painted = await previewGeneratedSourceRelocated(item.result);

      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );
      assert.equal(file.slides.length, 3, `${item.label} slide count`);
      assert.equal(painted.pages.length, 3, `${item.label} preview page count`);
      for (const [pageIndex, page] of painted.pages.entries()) {
        assert.deepEqual(
          page.diagnostics,
          [],
          `${item.label} preview page ${pageIndex + 1} diagnostics`,
        );
        assert.ok(page.png instanceof Uint8Array);
      }

      // Slide 1: multi-column header, rowSpan=2, borders, rich text, slideNumber, and pixel layout checks
      const fileTable1 = file.slides[0]?.objects[0];
      assert.equal(fileTable1?.kind, "table", `${item.label} s1 table`);
      assert.equal(fileTable1?.name, "PricingTable", `${item.label} name`);
      assert.ok(fileTable1?.table !== undefined);
      const t1 = fileTable1.table;
      assert.deepEqual(
        t1.merges,
        [
          { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
          { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
        ],
        `${item.label} s1 merges`,
      );
      assert.deepEqual(
        t1.cells,
        [
          ["Q1 Plan Overview", "", "第 1 页"],
          ["Pro\nTier", "Auto Width", "$49 / seat"],
          ["", "Summary", "Includes SSO & audit"],
        ],
        `${item.label} s1 cells`,
      );
      assert.equal(t1.columnWidthsEmu[0], 130 * emuPerPx, item.label);
      const totalWidthEmu1 = t1.columnWidthsEmu.reduce((a, b) => a + b, 0);
      assert.ok(Math.abs(totalWidthEmu1 - 576 * emuPerPx) <= 2, item.label);
      assert.equal(t1.rowHeightsEmu[0], 46 * emuPerPx, item.label);

      const pageBadgeText = t1.cellTexts.find(
        (entry) => entry.row === 0 && entry.column === 2,
      )?.text;
      assert.equal(
        pageBadgeText?.paragraphs[0]?.runs[1]?.field,
        "slidenum",
        `${item.label} slideNumber field`,
      );
      assert.equal(
        pageBadgeText?.paragraphs[0]?.runs[1]?.text,
        "1",
        `${item.label} slideNumber text`,
      );
      const priceText = t1.cellTexts.find(
        (entry) => entry.row === 1 && entry.column === 2,
      )?.text;
      assert.equal(priceText?.paragraphs[0]?.runs[0]?.text, "$49");
      assert.equal(priceText?.paragraphs[0]?.runs[0]?.bold, true);
      assert.equal(priceText?.paragraphs[0]?.runs[0]?.color, "#047857");

      const parts = unzipSync(built.pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /gridSpan="2"/u);
      assert.match(slide1Xml, /rowSpan="2"/u);
      assert.match(slide1Xml, /hMerge="1"/u);
      assert.match(slide1Xml, /vMerge="1"/u);
      assert.match(slide1Xml, /prstDash val="dash"/u);
      assert.match(slide1Xml, /prstDash val="dot"/u);

      // Preview page 1: verify header navy cell, rowSpan=2 blue cell, and adjacent amber detail cell pixels
      const png1 = PNG.sync.read(Buffer.from(painted.pages[0]!.png!));
      const t1X = fileTable1.boxEmu.x / emuPerPx;
      const t1Y = fileTable1.boxEmu.y / emuPerPx;
      const col0W = t1.columnWidthsEmu[0] / emuPerPx;
      const col1W = t1.columnWidthsEmu[1]! / emuPerPx;
      const col2W = t1.columnWidthsEmu[2]! / emuPerPx;
      const row0H = t1.rowHeightsEmu[0] / emuPerPx;
      const row1H = t1.rowHeightsEmu[1]! / emuPerPx;
      const row2H = t1.rowHeightsEmu[2]! / emuPerPx;
      assert.deepEqual(
        samplePngPixel(png1, t1X + 8, t1Y + 5),
        [0x1e, 0x29, 0x3b, 255],
        `${item.label} header span cell fill`,
      );
      assert.deepEqual(
        samplePngPixel(png1, t1X + 8, t1Y + row0H + row1H + row2H - 6),
        [0xef, 0xf6, 0xff, 255],
        `${item.label} rowSpan=2 cell bottom fill`,
      );
      assert.deepEqual(
        samplePngPixel(
          png1,
          t1X + col0W + col1W + col2W - 6,
          t1Y + row0H + row1H + row2H - 5,
        ),
        [0xfe, 0xf3, 0xc7, 255],
        `${item.label} adjacent amber cell fill`,
      );

      // Slide 2: narrow column wrapping height & preview position checks
      const wideFileObj = file.slides[1]?.objects[0];
      const narrowFileObj = file.slides[1]?.objects[1];
      assert.ok(wideFileObj?.table !== undefined);
      assert.ok(narrowFileObj?.table !== undefined);
      assert.equal(wideFileObj.table.columnWidthsEmu[0], 270 * emuPerPx);
      assert.equal(narrowFileObj.table.columnWidthsEmu[0], 130 * emuPerPx);
      const wideRowH = wideFileObj.table.rowHeightsEmu[0]! / emuPerPx;
      const narrowRowH = narrowFileObj.table.rowHeightsEmu[0]! / emuPerPx;
      assert.ok(
        narrowRowH > wideRowH * 1.5,
        `${item.label}: expected narrow height (${narrowRowH}) > 1.5x wide height (${wideRowH})`,
      );
      const png2 = PNG.sync.read(Buffer.from(painted.pages[1]!.png!));
      assert.deepEqual(
        samplePngPixel(png2, 28 + 8, 44 + 6),
        [0xec, 0xfd, 0xf5, 255],
        `${item.label} wide green table top-left`,
      );
      assert.deepEqual(
        samplePngPixel(png2, 332 + 8, 44 + narrowRowH - 6),
        [0xfe, 0xf2, 0xf2, 255],
        `${item.label} narrow red table bottom-left`,
      );
      // Directly below wide table at same vertical band inside tall narrow table stays white
      assert.deepEqual(
        samplePngPixel(png2, 28 + 8, 44 + narrowRowH - 6),
        [0xff, 0xff, 0xff, 255],
        `${item.label} below wide table stays white`,
      );

      // Slide 3: placeholder-bound table retains placeholder identity and layout-derived width
      const s3 = file.slides[2]!;
      const s3Layout = file.masters
        ?.flatMap((master) => master.layouts)
        .find((layout) => layout.part === s3.layoutPart);
      assert.equal(s3Layout?.name, "Table Layout", item.label);
      const phTable = s3.objects.find((obj) => obj.kind === "table");
      assert.ok(
        phTable?.table !== undefined,
        `${item.label} placeholder table`,
      );
      assert.equal(phTable.name, "PlaceholderMetrics");
      assert.deepEqual(phTable.placeholder, { type: "tbl", idx: 1 });
      assert.deepEqual(phTable.table.cells, [
        ["Region", "Notes"],
        ["EMEA", "Expanded coverage"],
      ]);
      const phWidthEmu = phTable.table.columnWidthsEmu.reduce(
        (a, b) => a + b,
        0,
      );
      assert.ok(Math.abs(phWidthEmu - 576 * emuPerPx) <= 2, item.label);
    }
  }, 60_000);

  it("keeps auto and fr column recalculation live when editing direct-authored generated vanilla and react source", async () => {
    const vanillaDoc = createTableAcceptanceVanillaDeck();
    const reactDoc = await createTableAcceptanceReactDeck();

    const originalBuilt = await build(vanillaDoc, { embedFonts: "full" });
    const originalFile = await (
      await import("@pptxdk/validation")
    ).validatePptx(originalBuilt.pptx);
    const originalAutoWidth =
      originalFile.slides[0]!.objects[0]!.table!.columnWidthsEmu[1]!;

    const directCases = [
      {
        label: "direct-edit:vanilla",
        result: codegen(vanillaDoc, { target: "vanilla" }),
      },
      {
        label: "direct-edit:react",
        result: codegen(reactDoc, { target: "react" }),
      },
    ] as const;

    const editedCases = directCases.map(({ label, result }) => {
      const source = result.files[result.entry] as string;
      const updated = source
        .replace("Auto Width", "Much Wider Auto Column Label")
        .replace(/(id:\s*"pricing-table"[\s\S]*?)width:\s*576/u, "$1width: 660")
        .replace(/(id="pricing-table"[\s\S]*?)width=\{576\}/u, "$1width={660}");
      assert.notEqual(updated, source, `${label} edit applied`);
      return {
        label,
        result: {
          ...result,
          files: {
            ...result.files,
            [result.entry]: updated,
          },
        },
      };
    });

    await verifyGeneratedSourcesTypecheck(editedCases);

    for (const item of editedCases) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      const painted = await previewGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], item.label);
      assert.deepEqual(painted.pages[0]?.diagnostics, [], item.label);

      const tbl = file.slides[0]!.objects[0]!.table!;
      assert.equal(tbl.columnWidthsEmu[0], 130 * emuPerPx, item.label);
      assert.ok(
        tbl.columnWidthsEmu[1]! > originalAutoWidth,
        `${item.label}: expected auto column (${tbl.columnWidthsEmu[1]}) > original (${originalAutoWidth})`,
      );
      const totalWidth = tbl.columnWidthsEmu.reduce((a, b) => a + b, 0);
      assert.ok(Math.abs(totalWidth - 660 * emuPerPx) <= 2, item.label);
      assert.equal(tbl.cells[1]?.[1], "Much Wider Auto Column Label");
    }
  }, 60_000);

  it("reconstructs 2x2 hMerge/vMerge grids and supports independent merged-cell body edits while preserving merge range and siblings", async () => {
    const mergeDeck = deck({ width: 640, height: 360 }).append(
      slide({ id: "rect-merge-slide" }).append(
        table({
          id: "rect-merge",
          name: "MergeGrid",
          x: 24,
          y: 28,
          columns: [110, 110, 140],
          style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
        }).append(
          row({ id: "r0", style: { height: 34 } }).append(
            cell("2x2 Anchor", {
              id: "anchor",
              colSpan: 2,
              rowSpan: 2,
              style: {
                backgroundColor: "#DBEAFE",
                borderColor: "#2563EB",
                borderWidth: 2,
              },
            }),
            cell("TopRight", {
              id: "tr",
              style: { backgroundColor: "#F1F5F9" },
            }),
          ),
          row({ id: "r1", style: { height: 34 } }).append(
            cell("MidRight", {
              id: "mr",
              style: { backgroundColor: "#FEF3C7" },
            }),
          ),
          row({ id: "r2", style: { height: 34 } }).append(
            cell("B0"),
            cell("B1"),
            cell("B2"),
          ),
        ),
      ),
    );

    const built = await build(mergeDeck, { embedFonts: "full" });
    const parts = unzipSync(built.pptx);
    const slidePath = "ppt/slides/slide1.xml";
    // Simulate external PowerPoint/POI output where anchor omits gridSpan/rowSpan and relies on hMerge/vMerge.
    parts[slidePath] = strToU8(
      strFromU8(parts[slidePath]!)
        .replace('<a:tc gridSpan="2" rowSpan="2">', "<a:tc>")
        .replace('<a:tc hMerge="1" rowSpan="2">', '<a:tc hMerge="1">')
        .replace('<a:tc vMerge="1" gridSpan="2">', '<a:tc vMerge="1">'),
    );
    const externalBytes = new Uint8Array(zipSync(parts));
    const imported = importPptx(externalBytes);
    externalBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const generated = [
      {
        label: "rect-merge:vanilla",
        result: codegen(imported.document, { target: "vanilla" }),
      },
      {
        label: "rect-merge:react",
        result: codegen(imported.document, { target: "react" }),
      },
    ] as const;

    const edited = generated.map(({ label, result }) => {
      const source = result.files[result.entry] as string;
      const updated = source.replace("2x2 Anchor", "Updated 2x2 Anchor");
      assert.notEqual(updated, source, `${label} merged cell edit applied`);
      return {
        label: `${label}:edited`,
        result: {
          ...result,
          files: {
            ...result.files,
            [result.entry]: updated,
          },
        },
      };
    });

    await verifyGeneratedSourcesTypecheck([...generated, ...edited]);

    for (const item of generated) {
      const { built: rtBuilt, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      const painted = await previewGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], item.label);
      assert.deepEqual(painted.pages[0]?.diagnostics, [], item.label);
      const tbl = file.slides[0]!.objects[0]!.table!;
      assert.deepEqual(tbl.merges, [
        { row: 0, column: 0, rowSpan: 2, colSpan: 2 },
      ]);
      assert.deepEqual(tbl.cells, [
        ["2x2 Anchor", "", "TopRight"],
        ["", "", "MidRight"],
        ["B0", "B1", "B2"],
      ]);
      const xml = strFromU8(unzipSync(rtBuilt.pptx)[slidePath]!);
      assert.match(xml, /gridSpan="2"/u);
      assert.match(xml, /rowSpan="2"/u);
      assert.match(xml, /hMerge="1"/u);
      assert.match(xml, /vMerge="1"/u);
    }

    for (const item of edited) {
      const { built: editedBuilt, file } =
        await rebuildGeneratedSourceRelocated(item.result);
      const painted = await previewGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], item.label);
      assert.deepEqual(painted.pages[0]?.diagnostics, [], item.label);
      const tbl = file.slides[0]!.objects[0]!.table!;
      assert.deepEqual(tbl.merges, [
        { row: 0, column: 0, rowSpan: 2, colSpan: 2 },
      ]);
      assert.deepEqual(tbl.cells, [
        ["Updated 2x2 Anchor", "", "TopRight"],
        ["", "", "MidRight"],
        ["B0", "B1", "B2"],
      ]);
      assert.deepEqual(tbl.columnWidthsEmu, [
        110 * emuPerPx,
        110 * emuPerPx,
        140 * emuPerPx,
      ]);
      assert.deepEqual(tbl.rowHeightsEmu, [
        34 * emuPerPx,
        34 * emuPerPx,
        34 * emuPerPx,
      ]);

      const png = PNG.sync.read(Buffer.from(painted.pages[0].png!));
      assert.deepEqual(
        samplePngPixel(png, 24 + 8, 28 + 34 + 26),
        [0xdb, 0xea, 0xfe, 255],
        `${item.label} merged anchor span bottom-left stays blue`,
      );
      assert.deepEqual(
        samplePngPixel(png, 24 + 220 + 8, 28 + 34 + 6),
        [0xfe, 0xf3, 0xc7, 255],
        `${item.label} adjacent MidRight cell stays amber`,
      );

      const xml = strFromU8(unzipSync(editedBuilt.pptx)[slidePath]!);
      assert.match(xml, /gridSpan="2"/u);
      assert.match(xml, /rowSpan="2"/u);
      assert.match(xml, /hMerge="1"/u);
      assert.match(xml, /vMerge="1"/u);
    }
  }, 60_000);
});
