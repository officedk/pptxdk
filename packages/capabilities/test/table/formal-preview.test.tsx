import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  cell,
  deck,
  documentCapabilities,
  font,
  paragraph,
  row,
  shape,
  slide,
  slideNumber,
  span,
  table,
} from "@react-pptx/capabilities/document";
import { shapeCapability } from "@react-pptx/capabilities/shape";
import { tableCapability } from "@react-pptx/capabilities/table";
import { Cell, Row, Table } from "@react-pptx/capabilities/table-elements";
import { type CapabilityModule, CommonNode } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Slice 4 (Ticket 09): Table static preview positive and negative behaviors", () => {
  const jostRegularBytes = readFileSync(
    new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
  );
  const jostBoldBytes = readFileSync(
    new URL("../assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  );
  const jostFont = font("Jost*", [
    { weight: 400, data: jostRegularBytes },
    { weight: 700, data: jostBoldBytes },
  ]);
  const notoFont = font("Noto Sans CJK SC", [
    {
      weight: 400,
      data: readFileSync(
        new URL("../assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
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
          backgroundColor: "#1E3A8A",
          color: "#FFFFFF",
          fontWeight: "bold",
          borderColor: "#0F172A",
          borderWidth: 2,
          borderStyle: "solid",
        },
      },
    );
    const doc = deck({ width: 520, height: 300 }).append(
      slide({
        id: "table-slide",
        background: "#F8FAFC",
        style: { display: "flex", padding: 20, alignItems: "stretch" },
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
                  backgroundColor: "#DBEAFE",
                  borderColor: "#1D4ED8",
                  borderWidth: 2,
                },
              },
            ),
          ),
          row({ id: "body-row", style: { backgroundColor: "#FFFFFF" } }).append(
            cell("Merged\nRows", {
              id: "span-cell",
              rowSpan: 2,
              style: {
                fontFamily: jostFont,
                verticalAlign: "bottom",
                backgroundColor: "#FEF3C7",
                borderColor: "#D97706",
                borderWidth: 1,
              },
            }),
            cell("AutoCol", {
              id: "auto-cell",
              style: {
                fontFamily: jostFont,
                borderColor: "#CBD5E1",
                borderWidth: 1,
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
                  borderColor: "#CBD5E1",
                  borderWidth: 1,
                },
              },
            ),
          ),
          row({ id: "tail-row" }).append(
            cell("", {
              id: "empty-cell",
              style: {
                backgroundColor: "#ECFDF5",
                borderColor: "#10B981",
                borderWidth: 1,
              },
            }),
            cell("Tail", {
              id: "tail-cell",
              style: {
                fontFamily: jostFont,
                verticalAlign: "top",
                borderColor: "#CBD5E1",
                borderWidth: 1,
              },
            }),
          ),
        ),
      ),
    );

    const first = await previewDoc(doc);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain("<path");

    const rendered = PNG.sync.read(Buffer.from(page.png));
    const tableObs = first.report.slides[0]?.objects[0]?.table;
    assert.ok(tableObs !== undefined);
    // Stretched height should expand auto rows to fill 260px (300 - 40 padding).
    const totalHeight = tableObs.rowHeights.reduce((a, b) => a + b, 0);
    expect(totalHeight).toBeCloseTo(260, 1);

    // Sample pixel inside hero-cell (#1E3A8A) and span-cell (#FEF3C7) and empty-cell (#ECFDF5).
    const heroBox = first.query
      .getSlide("table-slide")
      .get("pricing")
      .get("header-row")
      .get("hero-cell")
      .observe().box!;
    const heroPixel =
      (Math.round(heroBox.y + 4) * rendered.width + Math.round(heroBox.x + 6)) *
      4;
    assert.deepEqual(
      [...rendered.data.subarray(heroPixel, heroPixel + 4)],
      [0x1e, 0x3a, 0x8a, 255],
    );

    const emptyBox = first.query
      .getSlide("table-slide")
      .get("pricing")
      .get("tail-row")
      .get("empty-cell")
      .observe().box!;
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
    editableCell.style({ backgroundColor: "#047857" });
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
          { id: "jsx-tbl", background: "#FFFFFF" },
          createElement(
            Table,
            {
              id: "t1",
              x: 20,
              y: 20,
              width: 240,
              columns: [100, "1fr"],
              style: { fontFamily: jostFont, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1", style: { height: 34, backgroundColor: "#E0F2FE" } },
              createElement(
                Cell,
                {
                  id: "c1",
                  style: { borderColor: "#0284C7", borderWidth: 1 },
                },
                "JSX Cell",
              ),
              createElement(
                Cell,
                {
                  id: "c2",
                  style: { borderColor: "#0284C7", borderWidth: 1 },
                },
                "Value",
              ),
            ),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
  });

  it("renders cropped [shapeCapability, tableCapability] build + preview without loading text.ts or registering textCapability", async () => {
    const croppedPreview = runDocPreview([shapeCapability, tableCapability]);
    const doc = deck({ width: 360, height: 200 }).append(
      slide({ id: "cropped" }).append(
        shape({
          id: "badge",
          x: 20,
          y: 20,
          width: 80,
          height: 40,
          fill: "#2563EB",
        }),
        table({
          id: "tbl",
          x: 20,
          y: 80,
          width: 260,
          columns: [120, "1fr"],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "r1" }).append(
            cell("Metric", {
              style: {
                backgroundColor: "#F1F5F9",
                borderColor: "#334155",
                borderWidth: 1,
              },
            }),
            cell("100%", {
              style: {
                borderColor: "#334155",
                borderWidth: 1,
              },
            }),
          ),
        ),
      ),
    );
    const result = await croppedPreview(doc);
    assert.deepEqual(result.pages[0]?.diagnostics, []);
    assert.ok(result.pages[0]?.png !== undefined);
    expect(result.pages[0]?.svg).toContain("<path");
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    const packageRoot = fileURLToPath(new URL("../../", import.meta.url));
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
const { createDeckNode, createSlideNode } = await import("@react-pptx/core/document");
const { createShapeNode, shapeCapability } = await import("@react-pptx/capabilities/shape");
const { createCellNode, createRowNode, createTableNode, tableCapability } = await import("@react-pptx/capabilities/table");
const { font } = await import("@react-pptx/capabilities/font");
const { createResolver } = await import("@react-pptx/core/build");
const { paint } = await import("@react-pptx/preview");
const jost = font("Jost*", [{ weight: 400, data: readFileSync("./test/assets/fonts/Jost-400-Book.ttf") }]);
const resolveTrimmed = createResolver([shapeCapability, tableCapability]);
const res = paint(resolveTrimmed(
  createDeckNode({ width: 300, height: 160 }).append(
    createSlideNode({ id: "s1" }).append(
      createShapeNode({ id: "s", x: 10, y: 10, width: 40, height: 40, fill: "#DC2626" }),
      createTableNode({ id: "t", x: 60, y: 10, width: 180, columns: [90, 90], style: { fontFamily: jost, fontSize: 14 } }).append(
        createRowNode({}).append(
          createCellNode({ style: { borderColor: "#000000" } }).content("A"),
          createCellNode({}).content("B"),
        )
      )
    )
  )
));
if (!res.pages[0]?.png || res.pages[0].diagnostics.length !== 0) {
  throw new Error(JSON.stringify(res.pages[0]?.diagnostics));
}
process.stdout.write("ok");
`,
      ],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(isolated.status, 0, isolated.stderr);
    assert.equal(isolated.stdout, "ok");
  });

  it("rejects non-solid cell borders, cell overflow, and unsupported cell text features with cell-located errors while preserving sibling pages", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const cases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof table>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "dashed-cell-border",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [120],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "bad-row" }).append(
            cell("Dashed", {
              id: "bad-cell",
              style: {
                borderColor: "#1E293B",
                borderWidth: 2,
                borderStyle: "dashed",
              },
            }),
          ),
        ),
        expected: /non-solid/i,
      },
      {
        label: "dotted-cell-border",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [120],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "bad-row" }).append(
            cell("Dotted", {
              id: "bad-cell",
              style: {
                borderColor: "#1E293B",
                borderWidth: 1,
                borderStyle: "dotted",
              },
            }),
          ),
        ),
        expected: /non-solid/i,
      },
      {
        label: "fixed-height-cell-overflow",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [80],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "bad-row", style: { height: 16 } }).append(
            cell("Line one\nLine two\nLine three", {
              id: "bad-cell",
              style: { padding: 4 },
            }),
          ),
        ),
        expected: /overflow/i,
      },
      {
        label: "text-decoration-underline",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [120],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "bad-row" }).append(
            cell("Underlined", {
              id: "bad-cell",
              style: { textDecoration: "underline" },
            }),
          ),
        ),
        expected: /decoration/i,
      },
      {
        label: "unsupported-arabic-script",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [120],
          style: { fontFamily: jostFont, fontSize: 14 },
        }).append(
          row({ id: "bad-row" }).append(
            cell("مرحبا", {
              id: "bad-cell",
            }),
          ),
        ),
        expected: /Latin and Han/i,
      },
      {
        label: "unmeasured-font-on-fixed-row",
        node: table({
          id: "bad-tbl",
          x: 10,
          y: 10,
          columns: [120],
          style: { fontFamily: "Aptos", fontSize: 14 },
        }).append(
          row({ id: "bad-row", style: { height: 30 } }).append(
            cell("Missing face", {
              id: "bad-cell",
            }),
          ),
        ),
        expected: /measurement|font/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 260, height: 140 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "good-page" }).append(
            table({
              id: "ok-tbl",
              x: 10,
              y: 10,
              columns: [120],
              style: { fontFamily: jostFont, fontSize: 14 },
            }).append(
              row().append(
                cell("Valid", {
                  style: { borderColor: "#0F172A", borderWidth: 1 },
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
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /table\[bad-tbl\]\/row\[bad-row\]\/cell\[bad-cell\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
    }
  });
});
