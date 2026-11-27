import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  cell,
  deck,
  font,
  paragraph,
  row,
  shape,
  slide,
  slideNumber,
  span,
  table,
} from "@pptxdk/capabilities/document";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { tableCapability } from "@pptxdk/capabilities/table";
import { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
import { type CapabilityModule, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { stripVariableFontTables } from "../../../test/sfnt-test-helpers.ts";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      slides: resolution.slides,
      slide: (id: string) => resolution.slide(id),
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Table static preview positive and negative behaviors", () => {
  const jostRegularBytes = readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  );
  const jostBoldBytes = readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  );
  const jostFont = font("Jost*", [
    { weight: 400, data: jostRegularBytes },
    { weight: 700, data: jostBoldBytes },
  ]);
  const notoFont = font("Noto Sans CJK SC", [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
          import.meta.url,
        ),
      ),
    },
  ]);
  const arabicFont = font("Noto Naskh Arabic", [
    {
      weight: 400,
      data: stripVariableFontTables(
        readFileSync(
          new URL(
            "../../../test/assets/fonts/NotoNaskhArabic.ttf",
            import.meta.url,
          ),
        ),
      ),
    },
  ]);

  it("renders multi-row/multi-column table fills, solid borders, merges, rich text, stretched layout, and isolates mutations", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const editableCell = cell(
      paragraph("North Q1", { style: { align: "center" } }),
      {
        id: "hero-cell",
        colSpan: 2,
        style: {
          fontFamily: jostFont,
          background: "#1E3A8A",
          color: "#FFFFFF",
          fontWeight: "bold",
          border: { color: "#0F172A", width: 2, style: "solid" },
        },
      },
    );
    const doc = deck({ width: 520, height: 300 }).append(
      slide({
        id: "table-slide",
        style: {
          background: "#F8FAFC",
          display: "flex",
          padding: 20,
          alignItems: "stretch",
        },
      }).append(
        table({
          id: "pricing",
          columns: [110, "auto", "1fr"],
          style: {
            width: 440,
            fontFamily: notoFont,
            fontSize: 15,
            color: "#0F172A",
          },
        }).append(
          row({ id: "header-row", style: { height: 42 } }).append(
            editableCell,
            cell(
              [
                paragraph(["第 ", slideNumber(), " 页"], {
                  style: { align: "right" },
                }),
              ],
              {
                id: "page-cell",
                style: {
                  background: "#DBEAFE",
                  border: { color: "#1D4ED8", width: 2 },
                },
              },
            ),
          ),
          row({ id: "body-row", style: { background: "#FFFFFF" } }).append(
            cell("Merged\nRows", {
              id: "span-cell",
              rowSpan: 2,
              style: {
                fontFamily: jostFont,
                verticalAlign: "bottom",
                background: "#FEF3C7",
                border: { color: "#D97706", width: 1 },
              },
            }),
            cell("AutoCol", {
              id: "auto-cell",
              style: {
                fontFamily: jostFont,
                border: { color: "#CBD5E1", width: 1 },
              },
            }),
            cell(
              paragraph([
                "窄列自动换行 ",
                span("WRAP", {
                  style: { fontFamily: jostFont, color: "#B91C1C" },
                }),
              ]),
              {
                id: "wrap-cell",
                style: {
                  border: { color: "#CBD5E1", width: 1 },
                },
              },
            ),
          ),
          row({ id: "tail-row" }).append(
            cell("", {
              id: "empty-cell",
              style: {
                background: "#ECFDF5",
                border: { color: "#10B981", width: 1 },
              },
            }),
            cell("Tail", {
              id: "tail-cell",
              style: {
                fontFamily: jostFont,
                verticalAlign: "top",
                border: { color: "#CBD5E1", width: 1 },
              },
            }),
          ),
        ),
      ),
    );

    const first = await previewDoc(doc);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const page = first.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain("<path");

    const rendered = PNG.sync.read(Buffer.from(page.png));
    const tableObs = first.slide("table-slide").get("pricing")
      .objects[0]?.table;
    assert.ok(tableObs !== undefined);
    // Stretched height should expand auto rows to fill 260px (300 - 40 padding).
    const totalHeight = tableObs.rows.reduce((a, r) => a + r.height, 0);
    expect(totalHeight).toBeCloseTo(260, 1);

    // Sample pixel inside hero-cell (#1E3A8A) and span-cell (#FEF3C7) and empty-cell (#ECFDF5).
    const heroBox = first
      .slide("table-slide")
      .get("pricing")
      .get("header-row")
      .get("hero-cell").box!;
    const heroPixel =
      (Math.round(heroBox.y + 4) * rendered.width + Math.round(heroBox.x + 6)) *
      4;
    assert.deepEqual(
      [...rendered.data.subarray(heroPixel, heroPixel + 4)],
      [0x1e, 0x3a, 0x8a, 255],
    );

    const emptyBox = first
      .slide("table-slide")
      .get("pricing")
      .get("tail-row")
      .get("empty-cell").box!;
    const emptyPixel =
      (Math.round(emptyBox.y + emptyBox.height / 2) * rendered.width +
        Math.round(emptyBox.x + emptyBox.width / 2)) *
      4;
    assert.deepEqual(
      [...rendered.data.subarray(emptyPixel, emptyPixel + 4)],
      [0xec, 0xfd, 0xf5, 255],
    );

    // Mutate after preview and verify prior result remains immutable.
    const prevPng = page.png.slice();
    const prevSvg = page.svg;
    const prevPptx = first.pptx.slice();
    editableCell.style({ background: "#047857" });
    editableCell.content("South Q2");

    const second = await previewDoc(doc);
    assert.deepEqual(first.pages[0]?.png, prevPng);
    assert.equal(first.pages[0]?.svg, prevSvg);
    assert.deepEqual(first.pptx, prevPptx);
    assert.notEqual(second.pages[0]?.svg, prevSvg);
    const secondRendered = PNG.sync.read(Buffer.from(second.pages[0]!.png!));
    assert.deepEqual(
      [...secondRendered.data.subarray(heroPixel, heroPixel + 4)],
      [0x04, 0x78, 0x57, 255],
    );

    // JSX entry produces identical table layout & valid PNG.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "jsx-tbl", style: { background: "#FFFFFF" } },
          createElement(
            Table,
            {
              id: "t1",
              columns: [100, "1fr"],
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 240,
                fontFamily: jostFont,
                fontSize: 14,
              },
            },
            createElement(
              Row,
              { id: "r1", style: { height: 34, background: "#E0F2FE" } },
              createElement(
                Cell,
                {
                  id: "c1",
                  style: { border: { color: "#0284C7", width: 1 } },
                },
                "JSX Cell",
              ),
              createElement(
                Cell,
                {
                  id: "c2",
                  style: { border: { color: "#0284C7", width: 1 } },
                },
                "Value",
              ),
            ),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.errors, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
  });

  it("renders cropped [shapeCapability, tableCapability] build + preview without loading text.ts or registering textCapability", async () => {
    const croppedPreview = runDocPreview([shapeCapability, tableCapability]);
    const doc = deck({ width: 360, height: 200 }).append(
      slide({ id: "cropped" }).append(
        shape({
          id: "badge",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 80,
            height: 40,
            background: "#2563EB",
          },
        }),
        table({
          id: "tbl",
          columns: [120, "1fr"],
          style: {
            position: "absolute",
            left: 20,
            top: 80,
            width: 260,
            fontFamily: jostFont,
            fontSize: 14,
          },
        }).append(
          row({ id: "r1" }).append(
            cell("Metric", {
              style: {
                background: "#F1F5F9",
                border: { color: "#334155", width: 1 },
              },
            }),
            cell("100%", {
              style: {
                border: { color: "#334155", width: 1 },
              },
            }),
          ),
        ),
      ),
    );
    const result = await croppedPreview(doc);
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    expect(result.pages[0]?.svg).toContain("<path");
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    const packageRoot = fileURLToPath(new URL("../../../", import.meta.url));
    const isolated = spawnSync(
      process.execPath,
      [
        "--conditions=source",
        "--input-type=module",
        "--eval",
        `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier.endsWith("/text.ts") || specifier.endsWith("/text.js") || specifier.endsWith("/text")) {
      throw new Error("Cropped table preview loaded text.ts: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});
const { createDeckNode, createSlideNode } = await import("@pptxdk/core/document");
const { createShapeNode, shapeCapability } = await import("@pptxdk/capabilities/shape");
const { createCellNode, createRowNode, createTableNode, tableCapability } = await import("@pptxdk/capabilities/table");
const { font } = await import("@pptxdk/capabilities/font");
const { createResolver } = await import("@pptxdk/core/build");
const { paint } = await import("@pptxdk/preview");
const jost = font("Jost*", [{ weight: 400, data: readFileSync("./test/assets/fonts/Jost-400-Book.ttf") }]);
const resolveTrimmed = createResolver([shapeCapability, tableCapability]);
const res = paint(resolveTrimmed(
  createDeckNode({ width: 300, height: 160 }).append(
    createSlideNode({ id: "s1" }).append(
      createShapeNode({ id: "s", style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#DC2626" } }),
      createTableNode({ id: "t", columns: [90, 90], style: { position: "absolute", left: 60, top: 10, width: 180, fontFamily: jost, fontSize: 14 } }).append(
        createRowNode({}).append(
          createCellNode({ style: { border: { color: "#000000" } } }).content("A"),
          createCellNode({}).content("B"),
        )
      )
    )
  )
));
if (!res.pages[0]?.png || res.pages[0].errors.length !== 0) {
  throw new Error(JSON.stringify(res.pages[0]?.errors));
}
process.stdout.write("ok");
`,
      ],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(isolated.status, 0, isolated.stderr);
    assert.equal(isolated.stdout, "ok");
  });

  it("renders dashed/dotted cell borders, full cell rich text (BiDi, justify, lists, decorations, textShadow, letterSpacing), and clips fixed-height overset cells with zero preview errors", async () => {
    const resolveWith = createResolver(documentCapabilities);
    const doc = deck({ width: 520, height: 300 }).append(
      slide({ id: "table-parity", style: { background: "#FFFFFF" } }).append(
        table({
          id: "parity-tbl",
          columns: [160, 160, 140],
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            fontFamily: jostFont,
            fontSize: 13,
            color: "#0F172A",
          },
        }).append(
          row({ id: "styled-row" }).append(
            cell(
              [
                paragraph("Dashed cell", {
                  style: {
                    textDecoration: "underline",
                    letterSpacing: 1.5,
                  },
                }),
                paragraph("List item 1", {
                  style: { list: { type: "bullet" } },
                }),
                paragraph("List item 2", {
                  style: { list: { type: "number", startAt: 2 } },
                }),
              ],
              {
                id: "dashed-cell",
                style: {
                  padding: 6,
                  border: { color: "#1E293B", width: 2, style: "dashed" },
                },
              },
            ),
            cell(
              [
                paragraph("Shadow strike", {
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
                paragraph("Justified sample copy across cell width", {
                  style: { align: "justify" },
                }),
              ],
              {
                id: "dotted-cell",
                style: {
                  padding: 6,
                  border: { color: "#2563EB", width: 2, style: "dotted" },
                },
              },
            ),
            cell("مرحبا بالعالم", {
              id: "arabic-cell",
              style: {
                fontFamily: arabicFont,
                padding: 6,
                border: { color: "#0F766E", width: 1, style: "solid" },
              },
            }),
          ),
        ),
        table({
          id: "overset-tbl",
          columns: [140],
          style: {
            position: "absolute",
            left: 20,
            top: 160,
            fontFamily: jostFont,
            fontSize: 15,
            color: "#DC2626",
          },
        }).append(
          row({ id: "short-row", style: { height: 22 } }).append(
            cell("LINE1\nLINE2\nLINE3\nLINE4\nLINE5", {
              id: "overset-cell",
              style: {
                padding: 2,
                border: { color: "#0F172A", width: 1 },
              },
            }),
          ),
        ),
      ),
    );

    const resolution = resolveWith(doc);
    const oversetCellMetrics = resolution
      .slide("table-parity")
      .get("overset-tbl")
      .get("short-row")
      .get("overset-cell").textMetrics;
    assert.equal(oversetCellMetrics?.status, "measured");
    assert.equal(oversetCellMetrics.fits, false);
    assert.ok(
      oversetCellMetrics.content.height > oversetCellMetrics.available.height,
    );
    const { pptx } = emit(resolution);
    assert.deepEqual((await validatePptx(pptx)).errors, []);

    const preview = paint(resolution);
    const page = preview.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain('stroke-dasharray="8 6"');
    expect(page.svg).toContain('stroke-dasharray="2 4"');
    expect(page.svg).toContain("<feDropShadow");
    expect(page.svg).toContain("<clipPath");

    // Overflowing lines below y=160+22=182 must be clipped; pixels at y=195 under overset-tbl must remain white.
    const rendered = PNG.sync.read(Buffer.from(page.png));
    for (let x = 28; x <= 140; x += 12) {
      const offset = (195 * rendered.width + x) * 4;
      assert.deepEqual(
        [...rendered.data.subarray(offset, offset + 4)],
        [255, 255, 255, 255],
      );
    }
  });

  it("rejects only missing glyphs and unmeasured fonts on fixed-height cells with cell-located errors while preserving sibling pages", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const cases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof table>;
      readonly expectedCode: "preview-missing-glyph" | "preview-missing-font";
      readonly expectedMessage: RegExp;
    }[] = [
      {
        label: "missing-glyph-arabic-cell",
        node: table({
          id: "bad-tbl",
          columns: [120],
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            fontFamily: jostFont,
            fontSize: 14,
          },
        }).append(
          row({ id: "bad-row", style: { height: 30 } }).append(
            cell("مرحبا", {
              id: "bad-cell",
            }),
          ),
        ),
        expectedCode: "preview-missing-glyph",
        expectedMessage: /glyph/i,
      },
      {
        label: "unmeasured-font-on-fixed-row",
        node: table({
          id: "bad-tbl",
          columns: [120],
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            fontFamily: "Aptos",
            fontSize: 14,
          },
        }).append(
          row({ id: "bad-row", style: { height: 30 } }).append(
            cell("Missing face", {
              id: "bad-cell",
            }),
          ),
        ),
        expectedCode: "preview-missing-font",
        expectedMessage: /font/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 260, height: 140 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "good-page" }).append(
            table({
              id: "ok-tbl",
              columns: [120],
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                fontFamily: jostFont,
                fontSize: 14,
              },
            }).append(
              row().append(
                cell("Valid", {
                  style: { border: { color: "#0F172A", width: 1 } },
                }),
              ),
            ),
          ),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.errors[0]?.code,
        item.expectedCode,
        item.label,
      );
      assert.match(
        result.pages[0]?.errors[0]?.path ?? "",
        /table\[bad-tbl\]\/row\[bad-row\]\/cell\[bad-cell\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.errors[0]?.message ?? "",
        item.expectedMessage,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.errors, [], item.label);
    }
  });

  it("applies CSS table target height and min-content height without sibling overlap and paints cleanly", async () => {
    const resolveWith = createResolver(documentCapabilities);
    const doc = deck({ width: 320, height: 280 }).append(
      slide({
        id: "height-sizing",
        style: {
          background: "#FFFFFF",
          display: "flex",
          flexDirection: "column",
          gap: 8,
          padding: 16,
        },
      }).append(
        table({
          id: "underflow-auto-tbl",
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 13, height: 90 },
        }).append(
          row({ id: "auto-r" }).append(
            cell("Expands", {
              style: { border: { color: "#0F172A", width: 1 } },
            }),
          ),
        ),
        table({
          id: "underflow-explicit-tbl",
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 13, height: 90 },
        }).append(
          row({ style: { height: 28 } }).append(
            cell("Short", {
              style: { border: { color: "#0F172A", width: 1 } },
            }),
          ),
        ),
        table({
          id: "overflow-tbl",
          columns: [140],
          style: { fontFamily: jostFont, fontSize: 13, height: 20 },
        }).append(
          row({ style: { height: 50 } }).append(
            cell("Tall", { style: { border: { color: "#0F172A", width: 1 } } }),
          ),
        ),
        shape({
          id: "after-overflow",
          style: { width: 140, height: 20, background: "#2563EB" },
        }),
      ),
    );
    const resolution = resolveWith(doc);
    const pageLocator = resolution.slide("height-sizing");
    const autoBox = pageLocator.get("underflow-auto-tbl").box!;
    const explicitBox = pageLocator.get("underflow-explicit-tbl").box!;
    const overflowBox = pageLocator.get("overflow-tbl").box!;
    const siblingBox = pageLocator.get("after-overflow").box!;

    assert.equal(autoBox.height, 90);
    assert.equal(explicitBox.height, 28);
    assert.equal(overflowBox.height, 50);
    assert.ok(
      siblingBox.y >= overflowBox.y + overflowBox.height,
      `Sibling at y=${siblingBox.y} must not overlap expanded table ending at ${overflowBox.y + overflowBox.height}`,
    );

    const emitted = emit(resolution);
    assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png !== undefined);
  });
});
