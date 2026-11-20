import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

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
} from "@react-pptx/capabilities/document";
import { Cell, Row, Table } from "@react-pptx/capabilities/table-elements";
import { P, SlideNumber, Span } from "@react-pptx/capabilities/text-elements";
import { Deck, Slide } from "@react-pptx/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

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

const arabicFont = font("Noto Naskh Arabic", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/NotoNaskhArabic.ttf", import.meta.url),
    ),
  },
]);

const wrappingCopy =
  "Shared layout wraps long text cleanly inside constrained table cells.";

export const tablePositiveScenario: CapabilityScenario = {
  id: "table-layout-and-preview",
  capability: "rp:table",
  description:
    "Dual-author Table scenario covering px/auto/fr columns, rowSpan/colSpan merges, fills, solid borders, rich text with slide number, and constrained-width wrapping increasing row height.",
  visualCheckpoints: [
    "Page 1 (tracks-merges-and-fills): light slate #F8FAFC canvas; 3x3 table with dark navy #1E293B colSpan=2 header ('Q1 Plan Overview', bold white centered) and dark slate #334155 top-right cell ('第 1 页', right-aligned); left rowSpan=2 blue #EFF6FF cell ('Pro / Tier', bold blue middle-aligned) with 2px #3B82F6 solid border; 'Auto Width' and '$49 / seat' cells in row 2; 'Summary' and amber #FEF3C7 'Includes SSO & audit' cell in row 3.",
    "Page 2 (wrapping-and-width-constraint): white canvas; left green #ECFDF5 table (270px wide) wraps copy into fewer lines with shorter row height; right red-tinted #FEF2F2 table (130px wide) wraps identical copy into visibly more lines with noticeably taller row height.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify both slides contain native editable <p:graphicFrame><a:tbl> tables preserving merges (gridSpan/rowSpan), cell fills, solid borders, and matching column widths/row heights.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "tracks-merges-and-fills", background: "#F8FAFC" },
        createElement(
          Table,
          {
            id: "pricing-table",
            x: 32,
            y: 36,
            width: 576,
            columns: [130, "auto", "1fr"],
            style: {
              fontFamily: jostFont,
              fontSize: 15,
              color: "#0F172A",
            },
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
        { id: "wrapping-and-width-constraint", background: "#FFFFFF" },
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
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "tracks-merges-and-fills", background: "#F8FAFC" }).append(
        table({
          id: "pricing-table",
          x: 32,
          y: 36,
          width: 576,
          columns: [130, "auto", "1fr"],
          style: {
            fontFamily: jostFont,
            fontSize: 15,
            color: "#0F172A",
          },
        }).append(
          row({
            id: "header",
            style: { height: 46, backgroundColor: "#1E293B", color: "#FFFFFF" },
          }).append(
            cell(
              paragraph("Q1 Plan Overview", { style: { align: "center" } }),
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
            ),
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
              },
            }),
            cell("Auto Width", {
              id: "pro-metric",
              style: {
                padding: [6, 10],
                borderColor: "#CBD5E1",
                borderWidth: 1,
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
          row({
            id: "tier-notes",
            style: { backgroundColor: "#FFFFFF" },
          }).append(
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
      slide({
        id: "wrapping-and-width-constraint",
        background: "#FFFFFF",
      }).append(
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
    );
  },
  assert({ author, resolution, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.deepEqual(preview.pages[0]?.diagnostics, [], author);
    assert.deepEqual(preview.pages[1]?.diagnostics, [], author);

    const p1 = preview.pages[0];
    const p2 = preview.pages[1];
    expect(p1.svg).toContain("<path");
    expect(p2.svg).toContain("<path");

    // Verify native PPTX table structure via independent validation reader.
    const fileTable1 = validation.slides[0]?.objects[0];
    assert.equal(fileTable1?.kind, "table", author);
    assert.deepEqual(
      fileTable1?.table?.merges,
      [
        { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
        { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
      ],
      author,
    );
    assert.deepEqual(
      fileTable1?.table?.cells,
      [
        ["Q1 Plan Overview", "", "第 1 页"],
        ["Pro\nTier", "Auto Width", "$49 / seat"],
        ["", "Summary", "Includes SSO & audit"],
      ],
      author,
    );

    // Verify sub-part locators match formal table layout geometry.
    const pricingQuery = resolution.query
      .getSlide("tracks-merges-and-fills")
      .get("pricing-table");
    const headerSpanBox = pricingQuery
      .get("header")
      .get("title-span")
      .observe().box!;
    const tierSpanBox = pricingQuery
      .get("tier-pro")
      .get("tier-name")
      .observe().box!;
    const notesDetailBox = pricingQuery
      .get("tier-notes")
      .get("notes-detail")
      .observe().box!;

    const reportTable1 = resolution.report.slides[0]?.objects[0]?.table;
    assert.ok(reportTable1 !== undefined, author);
    assert.equal(
      headerSpanBox.width,
      reportTable1.columnWidths[0]! + reportTable1.columnWidths[1]!,
      author,
    );
    assert.equal(headerSpanBox.height, 46, author);
    assert.equal(
      tierSpanBox.height,
      reportTable1.rowHeights[1]! + reportTable1.rowHeights[2]!,
      author,
    );
    assert.equal(notesDetailBox.width, reportTable1.columnWidths[2], author);

    // Sample rendered PNG pixels inside header navy cell and amber detail cell.
    const png1 = PNG.sync.read(Buffer.from(p1.png!));
    const headerSample =
      (Math.round(headerSpanBox.y + 5) * png1.width +
        Math.round(headerSpanBox.x + 8)) *
      4;
    assert.deepEqual(
      [...png1.data.subarray(headerSample, headerSample + 4)],
      [0x1e, 0x29, 0x3b, 255],
      author,
    );
    const amberSample =
      (Math.round(notesDetailBox.y + notesDetailBox.height - 4) * png1.width +
        Math.round(notesDetailBox.x + notesDetailBox.width - 6)) *
      4;
    assert.deepEqual(
      [...png1.data.subarray(amberSample, amberSample + 4)],
      [0xfe, 0xf3, 0xc7, 255],
      author,
    );

    // Page 2: constrained narrow column causes wrapping and noticeably taller row height.
    const slide2Query = resolution.query.getSlide(
      "wrapping-and-width-constraint",
    );
    const wideBox = slide2Query
      .get("wide-table")
      .get("wide-row")
      .get("wide-cell")
      .observe().box!;
    const narrowBox = slide2Query
      .get("narrow-table")
      .get("narrow-row")
      .get("narrow-cell")
      .observe().box!;
    assert.equal(wideBox.width, 270, author);
    assert.equal(narrowBox.width, 130, author);
    assert.ok(
      narrowBox.height > wideBox.height * 1.5,
      `${author}: expected narrow wrapped row height (${narrowBox.height}) > 1.5x wide height (${wideBox.height})`,
    );

    const fileWide = validation.slides[1]?.objects[0]?.table;
    const fileNarrow = validation.slides[1]?.objects[1]?.table;
    assert.equal(fileWide?.columnWidthsEmu[0], 270 * 9525, author);
    assert.equal(fileNarrow?.columnWidthsEmu[0], 130 * 9525, author);
    assert.ok(
      (fileNarrow?.rowHeightsEmu[0] ?? 0) > (fileWide?.rowHeightsEmu[0] ?? 0),
      author,
    );
  },
};

export const tableRichTextBordersAndOversetScenario: CapabilityScenario = {
  id: "table-rich-text-borders-and-overset",
  capability: "rp:table",
  description:
    "Dual-author Table scenario covering dashed/dotted borders, multiscript/BiDi, justify, bullet & numbered lists, underline/line-through, textShadow, letterSpacing, slideNumber, and fixed-height overset clipping with build warning and zero preview diagnostics.",
  visualCheckpoints: [
    "Page 1 (borders-and-rich-text): white canvas; 3-column table with dashed border cell (underline, letterSpacing, bullet & numbered lists), dotted border cell (line-through, textShadow, justify), and solid border Arabic RTL cell.",
    "Page 2 (overset-cell-clipping): white canvas; fixed-height (24px) row with 4 lines of red text clipped cleanly to cell bounds so pixels below the cell remain pure white.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify native <a:prstDash val='dash'/'dot'/> borders, rich text formatting, and preserved overset cell text with overset warning in build report.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "borders-and-rich-text", background: "#FFFFFF" },
        createElement(
          Table,
          {
            id: "rich-tbl",
            x: 24,
            y: 28,
            columns: [200, 200, 170],
            style: { fontFamily: jostFont, fontSize: 13, color: "#0F172A" },
          },
          createElement(
            Row,
            { id: "rich-row" },
            createElement(
              Cell,
              {
                id: "dashed-list-cell",
                style: {
                  padding: 8,
                  backgroundColor: "#F8FAFC",
                  borderColor: "#1E293B",
                  borderWidth: 2,
                  borderStyle: "dashed",
                },
              },
              createElement(
                P,
                {
                  style: {
                    textDecoration: "underline",
                    letterSpacing: 1.5,
                  },
                },
                "Page ",
                createElement(SlideNumber, null),
                " Dashed",
              ),
              createElement(
                P,
                { style: { list: { type: "bullet" } } },
                "Bullet alpha",
              ),
              createElement(
                P,
                { style: { list: { type: "number", startAt: 2 } } },
                "Number two",
              ),
            ),
            createElement(
              Cell,
              {
                id: "dotted-shadow-cell",
                style: {
                  padding: 8,
                  backgroundColor: "#EFF6FF",
                  borderColor: "#2563EB",
                  borderWidth: 2,
                  borderStyle: "dotted",
                },
              },
              createElement(
                P,
                {
                  style: {
                    textDecoration: "line-through",
                    textShadow: {
                      color: "#000000",
                      blur: 2,
                      distance: 2,
                      angle: 45,
                    },
                  },
                },
                "Strike Shadow",
              ),
              createElement(
                P,
                { style: { align: "justify" } },
                "Justified multi-word copy across cell box width",
              ),
            ),
            createElement(
              Cell,
              {
                id: "bidi-cell",
                style: {
                  fontFamily: arabicFont,
                  padding: 8,
                  borderColor: "#0F766E",
                  borderWidth: 1,
                  borderStyle: "solid",
                },
              },
              "مرحبا بالعالم",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "overset-cell-clipping", background: "#FFFFFF" },
        createElement(
          Table,
          {
            id: "overset-tbl",
            x: 32,
            y: 36,
            columns: [180],
            style: { fontFamily: jostFont, fontSize: 15, color: "#DC2626" },
          },
          createElement(
            Row,
            { id: "fixed-row", style: { height: 24 } },
            createElement(
              Cell,
              {
                id: "overset-cell",
                style: {
                  padding: 2,
                  borderColor: "#0F172A",
                  borderWidth: 1,
                },
              },
              "OVERFLOW_1\nOVERFLOW_2\nOVERFLOW_3\nOVERFLOW_4",
            ),
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "borders-and-rich-text", background: "#FFFFFF" }).append(
        table({
          id: "rich-tbl",
          x: 24,
          y: 28,
          columns: [200, 200, 170],
          style: { fontFamily: jostFont, fontSize: 13, color: "#0F172A" },
        }).append(
          row({ id: "rich-row" }).append(
            cell(
              [
                paragraph(["Page ", slideNumber(), " Dashed"], {
                  style: {
                    textDecoration: "underline",
                    letterSpacing: 1.5,
                  },
                }),
                paragraph("Bullet alpha", {
                  style: { list: { type: "bullet" } },
                }),
                paragraph("Number two", {
                  style: { list: { type: "number", startAt: 2 } },
                }),
              ],
              {
                id: "dashed-list-cell",
                style: {
                  padding: 8,
                  backgroundColor: "#F8FAFC",
                  borderColor: "#1E293B",
                  borderWidth: 2,
                  borderStyle: "dashed",
                },
              },
            ),
            cell(
              [
                paragraph("Strike Shadow", {
                  style: {
                    textDecoration: "line-through",
                    textShadow: {
                      color: "#000000",
                      blur: 2,
                      distance: 2,
                      angle: 45,
                    },
                  },
                }),
                paragraph("Justified multi-word copy across cell box width", {
                  style: { align: "justify" },
                }),
              ],
              {
                id: "dotted-shadow-cell",
                style: {
                  padding: 8,
                  backgroundColor: "#EFF6FF",
                  borderColor: "#2563EB",
                  borderWidth: 2,
                  borderStyle: "dotted",
                },
              },
            ),
            cell("مرحبا بالعالم", {
              id: "bidi-cell",
              style: {
                fontFamily: arabicFont,
                padding: 8,
                borderColor: "#0F766E",
                borderWidth: 1,
                borderStyle: "solid",
              },
            }),
          ),
        ),
      ),
      slide({ id: "overset-cell-clipping", background: "#FFFFFF" }).append(
        table({
          id: "overset-tbl",
          x: 32,
          y: 36,
          columns: [180],
          style: { fontFamily: jostFont, fontSize: 15, color: "#DC2626" },
        }).append(
          row({ id: "fixed-row", style: { height: 24 } }).append(
            cell("OVERFLOW_1\nOVERFLOW_2\nOVERFLOW_3\nOVERFLOW_4", {
              id: "overset-cell",
              style: {
                padding: 2,
                borderColor: "#0F172A",
                borderWidth: 1,
              },
            }),
          ),
        ),
      ),
    );
  },
  assert({ author, resolution, pptx, preview }) {
    assert.equal(preview.pages.length, 2, author);
    assert.deepEqual(preview.pages[0]?.diagnostics, [], author);
    assert.deepEqual(preview.pages[1]?.diagnostics, [], author);

    // Overset build warning continues to appear in resolution.report.diagnostics.
    assert.ok(
      resolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" &&
          diag.code === "overset" &&
          diag.path ===
            "deck/slide[overset-cell-clipping]/table[overset-tbl]/row[fixed-row]/cell[overset-cell]",
      ),
      author,
    );

    const p1 = preview.pages[0];
    expect(p1.svg).toContain('stroke-dasharray="8 6"');
    expect(p1.svg).toContain('stroke-dasharray="2 4"');
    expect(p1.svg).toContain("<feDropShadow");

    const zip = unzipSync(pptx);
    const slide1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
    assert.match(slide1Xml, /<a:prstDash val="dash"\/>/u, author);
    assert.match(slide1Xml, /<a:prstDash val="dot"\/>/u, author);

    const p2 = preview.pages[1];
    expect(p2.svg).toContain("<clipPath");
    const png2 = PNG.sync.read(Buffer.from(p2.png!));
    // Below fixed row (y=36, height=24 => bottom=60): pixels at y=74 remain pure white.
    for (let x = 40; x <= 190; x += 20) {
      const offset = (74 * png2.width + x) * 4;
      assert.deepEqual(
        [...png2.data.subarray(offset, offset + 4)],
        [255, 255, 255, 255],
        author,
      );
    }
  },
};

export const tableRejectionScenario: CapabilityScenario = {
  id: "table-preview-rejections",
  capability: "rp:table",
  kind: "rejection",
  description:
    "Dual-author Table rejection scenario verifying per-page error isolation strictly for cell missing-glyph and unmeasured font resources while valid sibling pages (with dashed borders & overset clipping) render.",
  visualCheckpoints: [
    "Pages 1-2 fail preview (no PNG) with located cell-level font/glyph error diagnostics.",
    "Page 3 (valid-table-sibling): renders clean 2-column table with dashed border and #EFF6FF header fill.",
  ],
  nativeCheckpoints: [
    "Verify exported PPTX remains schema-valid with all native tables intact.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "missing-glyph-cell-page" },
        createElement(
          Table,
          {
            id: "bad-text-tbl",
            x: 20,
            y: 20,
            columns: [140],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1" },
            createElement(Cell, { id: "missing-glyph-cell" }, "汉"),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "unsupported-font-cell-page" },
        createElement(
          Table,
          {
            id: "bad-font-tbl",
            x: 20,
            y: 20,
            columns: [140],
            style: { fontFamily: "Aptos", fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 28 } },
            createElement(Cell, { id: "unmeasured-font-cell" }, "No Font Data"),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "valid-table-sibling" },
        createElement(
          Table,
          {
            id: "ok-tbl",
            x: 24,
            y: 40,
            columns: [120, 120],
            style: { fontFamily: jostFont, fontSize: 15 },
          },
          createElement(
            Row,
            { id: "r1" },
            createElement(
              Cell,
              {
                style: {
                  backgroundColor: "#EFF6FF",
                  padding: [6, 10],
                  borderColor: "#2563EB",
                  borderWidth: 2,
                  borderStyle: "dashed",
                },
              },
              "Valid",
            ),
            createElement(
              Cell,
              {
                style: {
                  padding: [6, 10],
                  borderColor: "#2563EB",
                  borderWidth: 2,
                },
              },
              "Sibling",
            ),
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "missing-glyph-cell-page" }).append(
        table({
          id: "bad-text-tbl",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("汉", {
              id: "missing-glyph-cell",
            }),
          ),
        ),
      ),
      slide({ id: "unsupported-font-cell-page" }).append(
        table({
          id: "bad-font-tbl",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: "Aptos", fontSize: 14 },
        }).append(
          row({ id: "r1", style: { height: 28 } }).append(
            cell("No Font Data", {
              id: "unmeasured-font-cell",
            }),
          ),
        ),
      ),
      slide({ id: "valid-table-sibling" }).append(
        table({
          id: "ok-tbl",
          x: 24,
          y: 40,
          columns: [120, 120],
          style: { fontFamily: jostFont, fontSize: 15 },
        }).append(
          row({ id: "r1" }).append(
            cell("Valid", {
              style: {
                backgroundColor: "#EFF6FF",
                padding: [6, 10],
                borderColor: "#2563EB",
                borderWidth: 2,
                borderStyle: "dashed",
              },
            }),
            cell("Sibling", {
              style: {
                padding: [6, 10],
                borderColor: "#2563EB",
                borderWidth: 2,
              },
            }),
          ),
        ),
      ),
    );
  },
  assert({ author, preview }) {
    const expectedPaths = [
      /table\[bad-text-tbl\]\/row\[r1\]\/cell\[missing-glyph-cell\]/,
      /table\[bad-font-tbl\]\/row\[r1\]\/cell\[unmeasured-font-cell\]/,
    ];
    expectedPaths.forEach((pattern, index) => {
      const page = preview.pages[index]!;
      assert.equal(page.png, undefined, `${author}: page ${index + 1}`);
      assert.equal(page.diagnostics[0]?.severity, "error", author);
      assert.match(page.diagnostics[0]?.path ?? "", pattern, author);
    });
    const validPage = preview.pages[2]!;
    assert.deepEqual(validPage.diagnostics, [], author);
    assert.ok(validPage.png !== undefined, author);
  },
};
