import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  cell,
  deck,
  defineSlideLayout,
  font,
  paragraph,
  row,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  table,
} from "@pptxdk/capabilities/document";
import { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../../../test/mechanism-verification.ts";
import type { tableMechanism, tablePlaceholderMechanism } from "../table.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
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
            columns: [110, "auto", "1fr"],
            style: {
              position: "absolute",
              left: 32,
              top: 32,
              width: 540,
              fontFamily: jostFont,
              fontSize: 14,
              color: "#0F172A",
            },
          },
          createElement(
            Row,
            {
              id: "header",
              style: {
                height: 40,
                background: "#1E293B",
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
                  border: { color: "#0F172A", width: 2, style: "solid" },
                },
              },
              createElement(P, { style: { align: "center" } }, "Q1 Revenue"),
            ),
            createElement(
              Cell,
              {
                id: "hdr-page",
                style: {
                  background: "#334155",
                  color: "#E2E8F0",
                  padding: [6, 10],
                  border: { color: "#0F172A", width: 2, style: "dashed" },
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
            { id: "r2", style: { background: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "rect-anchor",
                colSpan: 2,
                rowSpan: 2,
                style: {
                  background: "#EFF6FF",
                  color: "#1D4ED8",
                  fontWeight: "bold",
                  verticalAlign: "middle",
                  padding: [6, 10],
                  border: { color: "#2563EB", width: 2, style: "dotted" },
                },
              },
              createElement(P, { style: { align: "center" } }, "Core\nBundle"),
            ),
            createElement(
              Cell,
              {
                id: "r2-c3",
                style: {
                  background: "#F1F5F9",
                  padding: [6, 10],
                  border: { color: "#CBD5E1", width: 1 },
                },
              },
              "Auto Track",
            ),
          ),
          createElement(
            Row,
            { id: "r3", style: { background: "#FFFFFF" } },
            createElement(
              Cell,
              {
                id: "r3-c3",
                style: {
                  background: "#FEF3C7",
                  verticalAlign: "bottom",
                  padding: [6, 10],
                  border: { color: "#D97706", width: 1 },
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
            columns: [260],
            style: {
              position: "absolute",
              left: 28,
              top: 44,
              fontFamily: jostFont,
              fontSize: 15,
              color: "#065F46",
            },
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
                  background: "#ECFDF5",
                  border: { color: "#10B981", width: 2 },
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
            columns: [120],
            style: {
              position: "absolute",
              left: 332,
              top: 44,
              fontFamily: jostFont,
              fontSize: 15,
              color: "#991B1B",
            },
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
                  background: "#FEF2F2",
                  border: { color: "#EF4444", width: 2 },
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
          columns: [110, "auto", "1fr"],
          style: {
            position: "absolute",
            left: 32,
            top: 32,
            width: 540,
            fontFamily: jostFont,
            fontSize: 14,
            color: "#0F172A",
          },
        }).append(
          row({
            id: "header",
            style: { height: 40, background: "#1E293B", color: "#FFFFFF" },
          }).append(
            cell(paragraph("Q1 Revenue", { style: { align: "center" } }), {
              id: "hdr-span",
              colSpan: 2,
              style: {
                fontWeight: "bold",
                padding: [6, 10],
                border: { color: "#0F172A", width: 2, style: "solid" },
              },
            }),
            cell(
              paragraph(["Slide ", slideNumber()], {
                style: { align: "right" },
              }),
              {
                id: "hdr-page",
                style: {
                  background: "#334155",
                  color: "#E2E8F0",
                  padding: [6, 10],
                  border: { color: "#0F172A", width: 2, style: "dashed" },
                },
              },
            ),
          ),
          row({ id: "r2", style: { background: "#FFFFFF" } }).append(
            cell(paragraph("Core\nBundle", { style: { align: "center" } }), {
              id: "rect-anchor",
              colSpan: 2,
              rowSpan: 2,
              style: {
                background: "#EFF6FF",
                color: "#1D4ED8",
                fontWeight: "bold",
                verticalAlign: "middle",
                padding: [6, 10],
                border: { color: "#2563EB", width: 2, style: "dotted" },
              },
            }),
            cell("Auto Track", {
              id: "r2-c3",
              style: {
                background: "#F1F5F9",
                padding: [6, 10],
                border: { color: "#CBD5E1", width: 1 },
              },
            }),
          ),
          row({ id: "r3", style: { background: "#FFFFFF" } }).append(
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
                  background: "#FEF3C7",
                  verticalAlign: "bottom",
                  padding: [6, 10],
                  border: { color: "#D97706", width: 1 },
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
          columns: [260],
          style: {
            position: "absolute",
            left: 28,
            top: 44,
            fontFamily: jostFont,
            fontSize: 15,
            color: "#065F46",
          },
        }).append(
          row({ id: "wide-row" }).append(
            cell(wrappingCopy, {
              id: "wide-cell",
              style: {
                padding: [10, 12],
                background: "#ECFDF5",
                border: { color: "#10B981", width: 2 },
              },
            }),
          ),
        ),
        table({
          id: "narrow-tbl",
          name: "NarrowTbl",
          columns: [120],
          style: {
            position: "absolute",
            left: 332,
            top: 44,
            fontFamily: jostFont,
            fontSize: 15,
            color: "#991B1B",
          },
        }).append(
          row({ id: "narrow-row" }).append(
            cell(wrappingCopy, {
              id: "narrow-cell",
              style: {
                padding: [10, 12],
                background: "#FEF2F2",
                border: { color: "#EF4444", width: 2 },
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

const tableHolderLayout = defineSlideLayout({
  id: "table-holder-layout",
  name: "Table Holder",
  placeholders: {
    grid: {
      type: "tbl",
      explicitType: true,
      idx: 11,
      style: {
        position: "absolute",
        left: 40,
        top: 60,
        width: 300,
        height: 140,
      },
      name: "Grid Slot",
    },
  },
});

export const tablePlaceholderRoundtripScenario: CapabilityScenario = {
  id: "mechanism-table-placeholder-roundtrip",
  capability: "table",
  description:
    "Dual-author table-placeholder mechanism scenario verifying inherited bounds for 1fr columns and <p:ph> binding roundtrip.",
  visualCheckpoints: [
    "Page 1 paints a 2x2 table sized to inherited layout width (300px) at (40, 60).",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain <p:nvPr><p:ph type='tbl' idx='11'/></p:nvPr> and <p:xfrm> with x=40px, y=60px, width=300px.",
  ],
  emitOptions: { embedFonts: "full" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "tbl-master", name: "Table Master" },
        createElement(SlideLayout, { layout: tableHolderLayout }),
      ),
      createElement(Slide, {
        id: "tbl-ph-slide",
        layout: tableHolderLayout,
        placeholders: {
          grid: createElement(
            Table,
            {
              id: "ph-table",
              columns: ["1fr", "1fr"],
              style: { fontFamily: jostFont, fontSize: 12, color: "#0F172A" },
            },
            createElement(
              Row,
              null,
              createElement(Cell, null, "Item"),
              createElement(Cell, null, "Amount"),
            ),
            createElement(
              Row,
              null,
              createElement(Cell, null, "Seats"),
              createElement(Cell, null, "42"),
            ),
          ),
        },
      }),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "tbl-master", name: "Table Master" }).append(
        slideLayout(tableHolderLayout),
      ),
      slide({
        id: "tbl-ph-slide",
        layout: tableHolderLayout,
        placeholders: {
          grid: table({
            id: "ph-table",
            columns: ["1fr", "1fr"],
            style: { fontFamily: jostFont, fontSize: 12, color: "#0F172A" },
          }).append(
            row().append(cell().content("Item"), cell().content("Amount")),
            row().append(cell().content("Seats"), cell().content("42")),
          ),
        },
      }),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 1);
    assert.equal(validation.slides[0]?.objects[0]?.kind, "table");
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
          assert.deepEqual(page.errors, []);
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
  "table-placeholder": {
    scenario: tablePlaceholderRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const tbl = validation.slides[0]?.objects[0];
      assert.equal(tbl?.kind, "table", `${stage} table kind`);
      assert.deepEqual(tbl?.placeholder, { type: "tbl", idx: 11 });
      assert.equal(tbl?.hasLocalTransform, true);
      assert.equal(tbl?.boxEmu.x, 40 * 9525);
      assert.equal(tbl?.boxEmu.y, 60 * 9525);
      assert.equal(tbl?.boxEmu.width, 300 * 9525);
      assert.deepEqual(tbl?.table?.columnWidthsEmu, [150 * 9525, 150 * 9525]);
      const slideXml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slideXml, /<p:ph type="tbl" idx="11"\/>/u);
    },
    assertPaint({ initial, roundtrip }) {
      for (const preview of [initial, roundtrip]) {
        assert.equal(preview.pages.length, 1);
        assert.deepEqual(preview.pages[0]?.errors, []);
        assert.ok(preview.pages[0]?.png instanceof Uint8Array);
      }
    },
  } satisfies MechanismVerificationEntry<typeof tablePlaceholderMechanism>,
} satisfies ExhaustiveMechanismVerificationMap<
  typeof tableMechanism | typeof tablePlaceholderMechanism
>;
