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

export const tableRejectionScenario: CapabilityScenario = {
  id: "table-preview-rejections",
  capability: "rp:table",
  kind: "rejection",
  description:
    "Dual-author Table rejection scenario verifying per-page error isolation for dashed/dotted cell borders, fixed-height cell overflow, and unsupported cell text styles while valid sibling pages render.",
  visualCheckpoints: [
    "Pages 1-4 fail preview (no PNG) with located cell-level error diagnostics.",
    "Page 5 (valid-table-sibling): renders clean 2-column table with solid border and #EFF6FF header fill.",
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
        { id: "dashed-border-page" },
        createElement(
          Table,
          {
            id: "bad-border-tbl",
            x: 20,
            y: 20,
            columns: [140],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1" },
            createElement(
              Cell,
              {
                id: "dashed-cell",
                style: {
                  borderColor: "#1E293B",
                  borderWidth: 2,
                  borderStyle: "dashed",
                },
              },
              "Dashed border",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "dotted-border-page" },
        createElement(
          Table,
          {
            id: "bad-dot-tbl",
            x: 20,
            y: 20,
            columns: [140],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1" },
            createElement(
              Cell,
              {
                id: "dotted-cell",
                style: {
                  borderColor: "#1E293B",
                  borderWidth: 2,
                  borderStyle: "dotted",
                },
              },
              "Dotted border",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "cell-overflow-page" },
        createElement(
          Table,
          {
            id: "overflow-tbl",
            x: 20,
            y: 20,
            columns: [90],
            style: { fontFamily: jostFont, fontSize: 14 },
          },
          createElement(
            Row,
            { id: "r1", style: { height: 18 } },
            createElement(
              Cell,
              { id: "overflow-cell", style: { padding: 4 } },
              "Line 1\nLine 2\nLine 3",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "unsupported-cell-text-page" },
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
            createElement(
              Cell,
              { id: "underline-cell", style: { textDecoration: "underline" } },
              "Underlined",
            ),
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
      slide({ id: "dashed-border-page" }).append(
        table({
          id: "bad-border-tbl",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("Dashed border", {
              id: "dashed-cell",
              style: {
                borderColor: "#1E293B",
                borderWidth: 2,
                borderStyle: "dashed",
              },
            }),
          ),
        ),
      ),
      slide({ id: "dotted-border-page" }).append(
        table({
          id: "bad-dot-tbl",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("Dotted border", {
              id: "dotted-cell",
              style: {
                borderColor: "#1E293B",
                borderWidth: 2,
                borderStyle: "dotted",
              },
            }),
          ),
        ),
      ),
      slide({ id: "cell-overflow-page" }).append(
        table({
          id: "overflow-tbl",
          x: 20,
          y: 20,
          columns: [90],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1", style: { height: 18 } }).append(
            cell("Line 1\nLine 2\nLine 3", {
              id: "overflow-cell",
              style: { padding: 4 },
            }),
          ),
        ),
      ),
      slide({ id: "unsupported-cell-text-page" }).append(
        table({
          id: "bad-text-tbl",
          x: 20,
          y: 20,
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("Underlined", {
              id: "underline-cell",
              style: { textDecoration: "underline" },
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
      /table\[bad-border-tbl\]\/row\[r1\]\/cell\[dashed-cell\]/,
      /table\[bad-dot-tbl\]\/row\[r1\]\/cell\[dotted-cell\]/,
      /table\[overflow-tbl\]\/row\[r1\]\/cell\[overflow-cell\]/,
      /table\[bad-text-tbl\]\/row\[r1\]\/cell\[underline-cell\]/,
    ];
    expectedPaths.forEach((pattern, index) => {
      const page = preview.pages[index]!;
      assert.equal(page.png, undefined, `${author}: page ${index + 1}`);
      assert.equal(page.diagnostics[0]?.severity, "error", author);
      assert.match(page.diagnostics[0]?.path ?? "", pattern, author);
    });
    const validPage = preview.pages[4]!;
    assert.deepEqual(validPage.diagnostics, [], author);
    assert.ok(validPage.png !== undefined, author);
  },
};
