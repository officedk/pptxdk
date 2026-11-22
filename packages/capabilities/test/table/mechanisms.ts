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
} from "@pptxdk/capabilities/document";
import { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { tableMechanism } from "../../src/table/table.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../mechanism-verification.ts";

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

const wrappingCopy =
  "Shared layout wraps long text cleanly inside constrained table cells.";

export const tableRoundtripScenario: CapabilityScenario = {
  id: "mechanism-table-roundtrip",
  capability: "table",
  description:
    "Dual-author Transitional table scenario covering px/auto/fr columns, colSpan/rowSpan/2x2 merges, borders/fills/padding, wrapping height growth, and SlideNumber fields.",
  visualCheckpoints: [
    "Page 1 renders a 3x3 table with dark navy colSpan=2 header, SlideNumber badge, 2x2 rectangular merge with dotted border, and styled body cells.",
    "Page 2 renders wide green table and narrow red table demonstrating text wrapping increasing solved row height.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain <a:tbl> with exact EMU columnWidths/rowHeights, gridSpan/rowSpan/hMerge/vMerge, borders, fills, and <a:fld type='slidenum'>.",
  ],
  emitOptions: { embedFonts: "full" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "table-merges-and-fields" },
        createElement(
          Table,
          {
            id: "pricing-matrix",
            name: "PricingMatrix",
            x: 32,
            y: 32,
            width: 540,
            columns: [110, "auto", "1fr"],
            style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
          },
          createElement(
            Row,
            {
              id: "header",
              style: {
                height: 40,
                backgroundColor: "#1E293B",
                color: "#FFFFFF",
              },
            },
            createElement(
              Cell,
              {
                id: "hdr-span",
                colSpan: 2,
                style: {
                  fontWeight: "bold",
                  padding: [6, 10],
                  borderColor: "#0F172A",
                  borderWidth: 2,
                  borderStyle: "solid",
                },
              },
              createElement(P, { style: { align: "center" } }, "Q1 Revenue"),
            ),
            createElement(
              Cell,
              {
                id: "hdr-page",
                style: {
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
                "Slide ",
                createElement(SlideNumber, null),
              ),
            ),
          ),
          createElement(
            Row,
            { id: "r2", style: { backgroundColor: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "rect-anchor",
                colSpan: 2,
                rowSpan: 2,
                style: {
                  backgroundColor: "#EFF6FF",
                  color: "#1D4ED8",
                  fontWeight: "bold",
                  verticalAlign: "middle",
                  padding: [6, 10],
                  borderColor: "#2563EB",
                  borderWidth: 2,
                  borderStyle: "dotted",
                },
              },
              createElement(P, { style: { align: "center" } }, "Core\nBundle"),
            ),
            createElement(
              Cell,
              {
                id: "r2-c3",
                style: {
                  backgroundColor: "#F1F5F9",
                  padding: [6, 10],
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                },
              },
              "Auto Track",
            ),
          ),
          createElement(
            Row,
            { id: "r3", style: { backgroundColor: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "r3-c3",
                style: {
                  backgroundColor: "#FEF3C7",
                  verticalAlign: "bottom",
                  padding: [6, 10],
                  borderColor: "#D97706",
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
        ),
      ),
      createElement(
        Slide,
        { id: "table-wrapping" },
        createElement(
          Table,
          {
            id: "wide-tbl",
            name: "WideTbl",
            x: 28,
            y: 44,
            columns: [260],
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
            id: "narrow-tbl",
            name: "NarrowTbl",
            x: 332,
            y: 44,
            columns: [120],
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
      slide({ id: "table-merges-and-fields" }).append(
        table({
          id: "pricing-matrix",
          name: "PricingMatrix",
          x: 32,
          y: 32,
          width: 540,
          columns: [110, "auto", "1fr"],
          style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
        }).append(
          row({
            id: "header",
            style: { height: 40, backgroundColor: "#1E293B", color: "#FFFFFF" },
          }).append(
            cell(paragraph("Q1 Revenue", { style: { align: "center" } }), {
              id: "hdr-span",
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
              paragraph(["Slide ", slideNumber()], {
                style: { align: "right" },
              }),
              {
                id: "hdr-page",
                style: {
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
          row({ id: "r2", style: { backgroundColor: "#FFFFFF" } }).append(
            cell(paragraph("Core\nBundle", { style: { align: "center" } }), {
              id: "rect-anchor",
              colSpan: 2,
              rowSpan: 2,
              style: {
                backgroundColor: "#EFF6FF",
                color: "#1D4ED8",
                fontWeight: "bold",
                verticalAlign: "middle",
                padding: [6, 10],
                borderColor: "#2563EB",
                borderWidth: 2,
                borderStyle: "dotted",
              },
            }),
            cell("Auto Track", {
              id: "r2-c3",
              style: {
                backgroundColor: "#F1F5F9",
                padding: [6, 10],
                borderColor: "#CBD5E1",
                borderWidth: 1,
              },
            }),
          ),
          row({ id: "r3", style: { backgroundColor: "#FFFFFF" } }).append(
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
                id: "r3-c3",
                style: {
                  backgroundColor: "#FEF3C7",
                  verticalAlign: "bottom",
                  padding: [6, 10],
                  borderColor: "#D97706",
                  borderWidth: 1,
                },
              },
            ),
          ),
        ),
      ),
      slide({ id: "table-wrapping" }).append(
        table({
          id: "wide-tbl",
          name: "WideTbl",
          x: 28,
          y: 44,
          columns: [260],
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
          id: "narrow-tbl",
          name: "NarrowTbl",
          x: 332,
          y: 44,
          columns: [120],
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
  assert({ validation }) {
    assert.equal(validation.slides.length, 2);
    assert.equal(validation.slides[0]?.objects[0]?.kind, "table");
    assert.equal(validation.slides[1]?.objects.length, 2);
  },
};

export const tableMechanismVerifications = {
  table: {
    scenario: tableRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 2);

      const matrix = validation.slides[0]!.objects[0]!;
      assert.equal(matrix.kind, "table", `${stage} slide1 table kind`);
      assert.equal(matrix.name, "PricingMatrix", `${stage} table name`);
      assert.equal(matrix.table?.columnWidthsEmu.length, 3);
      assert.equal(matrix.table?.columnWidthsEmu[0], 110 * 9525);
      assert.equal(
        matrix.table?.columnWidthsEmu.reduce((sum, value) => sum + value, 0),
        540 * 9525,
      );
      assert.equal(matrix.table?.rowHeightsEmu[0], 40 * 9525);
      assert.deepEqual(matrix.table?.merges, [
        { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
        { row: 1, column: 0, rowSpan: 2, colSpan: 2 },
      ]);
      assert.deepEqual(matrix.table?.cells, [
        ["Q1 Revenue", "", "Slide 1"],
        ["Core\nBundle", "", "Auto Track"],
        ["", "", "$49 / seat"],
      ]);
      assert.equal(
        matrix.table?.cellTexts[1]?.text.paragraphs[0]?.runs[1]?.field,
        "slidenum",
      );

      const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:tbl>/u, `${stage} <a:tbl>`);
      assert.match(slide1Xml, /gridSpan="2"/u, `${stage} gridSpan`);
      assert.match(slide1Xml, /rowSpan="2"/u, `${stage} rowSpan`);
      assert.match(slide1Xml, /hMerge="1"/u, `${stage} hMerge`);
      assert.match(slide1Xml, /vMerge="1"/u, `${stage} vMerge`);
      assert.match(slide1Xml, /<a:prstDash val="dash"\/>/u, `${stage} dashed`);
      assert.match(slide1Xml, /<a:prstDash val="dot"\/>/u, `${stage} dotted`);
      assert.match(slide1Xml, /type="slidenum"/u, `${stage} slidenum`);

      const wide = validation.slides[1]!.objects[0]!;
      const narrow = validation.slides[1]!.objects[1]!;
      assert.equal(wide.table?.columnWidthsEmu[0], 260 * 9525);
      assert.equal(narrow.table?.columnWidthsEmu[0], 120 * 9525);
      assert.ok(
        (narrow.table?.rowHeightsEmu[0] ?? 0) >
          (wide.table?.rowHeightsEmu[0] ?? 0),
        `${stage} narrow wrapped row must be taller than wide row`,
      );
    },
    assertPaint({ initial, roundtrip }) {
      for (const preview of [initial, roundtrip]) {
        assert.equal(preview.pages.length, 2);
        for (const page of preview.pages) {
          assert.deepEqual(page.diagnostics, []);
          assert.ok(page.png instanceof Uint8Array);
        }
        const p1Svg = preview.pages[0]!.svg!;
        assert.ok(p1Svg.includes("#1E293B"));
        assert.ok(p1Svg.includes("#EFF6FF"));
        assert.ok(p1Svg.includes("stroke-dasharray="));
        assert.ok(p1Svg.includes("<path "));
      }
    },
  } satisfies MechanismVerificationEntry<typeof tableMechanism>,
} satisfies ExhaustiveMechanismVerificationMap<typeof tableMechanism>;
