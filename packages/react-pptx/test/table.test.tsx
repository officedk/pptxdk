import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  Box,
  build,
  BuildError,
  Cell,
  Deck,
  font,
  P,
  Row,
  Slide,
  Table,
  tableCapability,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, type ReactNode } from "react";
import { describe, it, vi } from "vitest";

const notoSansScRegular = new Uint8Array(
  readFileSync(
    new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
  ),
);

function tableDeck(content: ReactNode): ReactNode {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "invalid" }, content),
  );
}

describe("native tables", () => {
  it("writes tables through shared text service even when Text capability is stripped", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx-table-trim-"));
    try {
      const loader = join(directory, "reject-text-box.mjs");
      await writeFile(
        loader,
        [
          "export async function resolve(specifier, context, nextResolve) {",
          "  const resolved = await nextResolve(specifier, context);",
          '  if (resolved.url.endsWith("/capabilities/src/text.ts")) { throw new Error("text-box implementation was loaded"); }',
          "  return resolved;",
          "}",
        ].join("\n"),
      );
      const child = spawnSync(process.execPath, [
        "--conditions=source",
        "--loader",
        loader,
        "--input-type=module",
        "--eval",
        [
          'import { createElement } from "react";',
          'import { Deck, Slide, expandDocument } from "@react-pptx/core";',
          'import { createResolver, emit } from "@react-pptx/core/build";',
          'import { Cell, Row, Table } from "@react-pptx/capabilities/table-elements";',
          'import { tableCapability } from "@react-pptx/capabilities/table";',
          'import { validatePptx } from "@react-pptx/validation";',
          "const resolve = createResolver([tableCapability]);",
          'const expanded = await expandDocument(createElement(Deck, { width: 160, height: 96 }, createElement(Slide, { id: "trimmed" }, createElement(Table, { id: "only-table", columns: [80] }, createElement(Row, { style: { height: 40 } }, createElement(Cell, null, "table survives Text crop"))))));',
          "const result = emit(resolve(expanded));",
          "const file = await validatePptx(result.pptx);",
          'if (file.errors.length !== 0 || file.slides[0]?.objects.length !== 1 || file.slides[0]?.objects[0]?.kind !== "table") throw new Error(JSON.stringify(file.errors));',
          "process.stdout.write(JSON.stringify({ kind: file.slides[0]?.objects[0]?.kind, text: file.slides[0]?.objects[0]?.table?.cells[0]?.[0] }));",
        ].join("\n"),
      ]);
      assert.equal(
        child.status,
        0,
        `cropped table process failed: ${child.stderr.toString()}`,
      );
      assert.deepEqual(JSON.parse(child.stdout.toString()), {
        kind: "table",
        text: "table survives Text crop",
      });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000); // Loading Open XML SDK WASM runtime in a separate process takes ~10 seconds on first run.

  it("shares kerning, case, decoration, and shadow services across table and cell text styles", async () => {
    const result = await build(
      tableDeck(
        createElement(
          Table,
          {
            id: "text-style-table",
            columns: [160],
            style: {
              fontFamily: "Aptos",
              fontSize: "12pt",
              letterSpacing: 1.25,
              textTransform: "uppercase",
              textDecoration: "underline",
              textShadow: { color: "accent3", blur: 2 },
            },
          },
          createElement(
            Row,
            { style: { height: 28 } },
            createElement(Cell, null, "ßeta"),
          ),
          createElement(
            Row,
            { style: { height: 28 } },
            createElement(
              Cell,
              {
                style: {
                  letterSpacing: 2,
                  textTransform: "uppercase",
                  textDecoration: "underline",
                  textShadow: { color: "accent4", blur: 1 },
                },
              },
              "cell style",
            ),
          ),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /spc="94"/u);
    assert.match(slideXml, /spc="150"/u);
    assert.equal((slideXml.match(/cap="all"/gu) ?? []).length, 2);
    assert.equal((slideXml.match(/u="sng"/gu) ?? []).length, 2);
    assert.equal((slideXml.match(/<a:outerShdw/gu) ?? []).length, 2);
  });

  it("identifies structural gaps and conflicting merges at table, row, or cell paths", async () => {
    const incomplete = createElement(
      Table,
      { id: "short", columns: [60, 60] },
      createElement(
        Row,
        { id: "missing" },
        createElement(Cell, { id: "only" }, "One"),
      ),
    );
    await assert.rejects(build(tableDeck(incomplete)), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.deepEqual(
        error.diagnostics.map(({ code, path }) => ({ code, path })),
        [
          {
            code: "invalid-table-structure",
            path: "deck/slide[invalid]/table[short]/row[missing]",
          },
        ],
      );
      return true;
    });

    const outOfBounds = createElement(
      Table,
      { id: "wide", columns: [60, 60] },
      createElement(
        Row,
        { id: "overflow" },
        createElement(Cell, { id: "too-wide", colSpan: 3 }, "Too wide"),
      ),
    );
    await assert.rejects(build(tableDeck(outOfBounds)), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "invalid-table-merge");
      assert.equal(
        error.diagnostics[0]?.path,
        "deck/slide[invalid]/table[wide]/row[overflow]/cell[too-wide]",
      );
      return true;
    });

    const conflicting = createElement(
      Table,
      { id: "conflict", columns: [60, 60, 60] },
      createElement(
        Row,
        { id: "first" },
        createElement(Cell, { rowSpan: 2 }, "Left"),
        createElement(Cell, null, "Centre"),
        createElement(Cell, { rowSpan: 2 }, "Right"),
      ),
      createElement(
        Row,
        { id: "second" },
        createElement(Cell, { id: "overlap", colSpan: 2 }, "Conflict"),
      ),
    );
    await assert.rejects(build(tableDeck(conflicting)), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "conflicting-table-merge");
      assert.equal(
        error.diagnostics[0]?.path,
        "deck/slide[invalid]/table[conflict]/row[second]/cell[overlap]",
      );
      return true;
    });

    const trailingRowSpan = createElement(
      Table,
      { id: "trailing", columns: [60] },
      createElement(
        Row,
        { id: "last" },
        createElement(Cell, { id: "too-tall", rowSpan: 2 }, "Too tall"),
      ),
    );
    await assert.rejects(
      build(tableDeck(trailingRowSpan)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-table-merge");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[invalid]/table[trailing]/row[last]/cell[too-tall]",
        );
        return true;
      },
    );
  });

  it("expands auto row height using actual text measurement while preserving explicitly small row height with warning", async () => {
    const style = {
      fontFamily: font("Noto Sans CJK SC", [
        { weight: 400, data: notoSansScRegular },
      ]),
      fontSize: 16,
    } as const;
    const automatic = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "automatic" },
          createElement(
            Table,
            { id: "automatic-row", columns: [30], style },
            createElement(
              Row,
              null,
              createElement(
                Cell,
                { style: { padding: 2 } },
                createElement(P, null, "中文中文"),
                createElement(P, null, "自动行高第二段"),
              ),
            ),
          ),
        ),
      ),
    );
    const autoHeight =
      automatic.report.slides[0]?.objects[0]?.table?.rowHeights[0] ?? 0;
    assert.ok(
      autoHeight > 80,
      `measured content height should wrap: ${autoHeight}px`,
    );
    assert.deepEqual(automatic.report.diagnostics, []);
    const automaticFile = await validatePptx(automatic.pptx);
    assert.deepEqual(automaticFile.errors, []);
    const automaticParagraphs =
      automaticFile.slides[0]?.objects[0]?.table?.cellTexts[0]?.text
        .paragraphs ?? [];
    assert.equal(automaticParagraphs.length, 2);
    assert.deepEqual(
      automaticParagraphs.map(({ lineSpacingPoints }) => lineSpacingPoints),
      [17.38, 17.38],
    );

    const spanningAutomatic = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "spanning-automatic" },
          createElement(
            Table,
            { id: "spanning-automatic", columns: [32, 32], style },
            createElement(
              Row,
              null,
              createElement(
                Cell,
                { rowSpan: 2, style: { padding: 2 } },
                createElement(P, null, "跨行中文中文"),
                createElement(P, null, "跨行自动高度"),
              ),
              createElement(Cell, null, "首行"),
            ),
            createElement(Row, null, createElement(Cell, null, "次行")),
          ),
        ),
      ),
    );
    const spanningFile = await validatePptx(spanningAutomatic.pptx);
    assert.deepEqual(spanningFile.errors, []);
    assert.deepEqual(spanningFile.slides[0]?.objects[0]?.table?.merges, [
      { row: 0, column: 0, rowSpan: 2, colSpan: 1 },
    ]);
    const spanningParagraphs =
      spanningFile.slides[0]?.objects[0]?.table?.cellTexts[0]?.text
        .paragraphs ?? [];
    assert.equal(spanningParagraphs.length, 2);
    assert.deepEqual(
      spanningParagraphs.map(({ lineSpacingPoints }) => lineSpacingPoints),
      [17.38, 17.38],
    );

    const fixed = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "fixed" },
          createElement(
            Table,
            { id: "fixed-row", columns: [30], style },
            createElement(
              Row,
              { id: "short", style: { height: 20 } },
              createElement(
                Cell,
                { id: "long", style: { padding: 2 } },
                "中文中文",
              ),
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(
      fixed.report.slides[0]?.objects[0]?.table?.rowHeights,
      [20],
    );
    assert.deepEqual(
      fixed.report.diagnostics.map(({ severity, code, path }) => ({
        severity,
        code,
        path,
      })),
      [
        {
          severity: "warning",
          code: "overset",
          path: "deck/slide[fixed]/table[fixed-row]/row[short]/cell[long]",
        },
      ],
    );
    const file = await validatePptx(fixed.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.table?.cells[0]?.[0], "中文中文");
    assert.equal(file.slides[0]?.objects[0]?.table?.rowHeightsEmu[0], 190_500);
    assert.equal(
      file.slides[0]?.objects[0]?.table?.cellTexts[0]?.text.paragraphs[0]
        ?.lineSpacingPoints,
      undefined,
    );

    const explicitLineHeight = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "explicit-line-height" },
          createElement(
            Table,
            { id: "explicit-line-height", columns: [60], style },
            createElement(
              Row,
              { style: { height: 40 } },
              createElement(Cell, { style: { lineHeight: 18 } }, "显式行距"),
            ),
          ),
        ),
      ),
    );
    const explicitLineHeightFile = await validatePptx(explicitLineHeight.pptx);
    assert.deepEqual(explicitLineHeightFile.errors, []);
    assert.equal(
      explicitLineHeightFile.slides[0]?.objects[0]?.table?.cellTexts[0]?.text
        .paragraphs[0]?.lineSpacingPoints,
      13.5,
    );
  });

  it("writes span on first-row and first-column continuation cells for rectangular merges and rejects missing spans in standalone reader", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "rectangle-merge" },
          createElement(
            Table,
            { id: "rectangle-merge", columns: [60, 60] },
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, { colSpan: 2, rowSpan: 2 }, "矩形合并"),
            ),
            createElement(Row, { style: { height: 24 } }),
          ),
        ),
      ),
    );

    const slidePath = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const slideXml = strFromU8(parts[slidePath]!);
    assert.deepEqual(
      [...slideXml.matchAll(/<a:tc(?=[\s>])([^>]*)>/gu)].map(
        ([, attributes]) => attributes?.trim() ?? "",
      ),
      [
        'gridSpan="2" rowSpan="2"',
        'hMerge="1" rowSpan="2"',
        'vMerge="1" gridSpan="2"',
        'hMerge="1" vMerge="1"',
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.table?.merges, [
      { row: 0, column: 0, rowSpan: 2, colSpan: 2 },
    ]);

    const malformedSlide = slideXml
      .replace('hMerge="1" rowSpan="2"', 'hMerge="1"')
      .replace('vMerge="1" gridSpan="2"', 'vMerge="1"');
    assert.notEqual(malformedSlide, slideXml);
    parts[slidePath] = strToU8(malformedSlide);
    const malformedFile = await validatePptx(zipSync(parts));
    assert.equal(
      malformedFile.errors.filter(({ code }) => code === "invalid-table-merge")
        .length,
      2,
      JSON.stringify(malformedFile.errors),
    );
  });

  it("keeps the raw table frame and grid when a native save normalizes the frame", async () => {
    const result = await build(
      tableDeck(
        createElement(
          Table,
          { columns: [100, 100] },
          createElement(
            Row,
            { style: { height: 24 } },
            createElement(Cell, null, "A"),
            createElement(Cell, null, "B"),
          ),
        ),
      ),
    );
    const baseline = await validatePptx(result.pptx);
    assert.deepEqual(baseline.errors, []);
    const table = baseline.slides[0]?.objects[0];
    assert.equal(table?.kind, "table");
    const rowHeight = table?.table?.rowHeightsEmu.reduce(
      (sum, height) => sum + height,
      0,
    );
    assert.equal(rowHeight, 24 * 9525);
    const nativeFrameHeight = rowHeight + 151511;
    const slidePath = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts[slidePath]!);
    const normalized = slide.replace(
      /(<p:graphicFrame>[\s\S]*?<a:ext cx="\d+" cy=")\d+("[^>]*>[\s\S]*?<a:tbl>)/u,
      (_match, prefix: string, suffix: string) =>
        `${prefix}${nativeFrameHeight}${suffix}`,
    );
    assert.notEqual(normalized, slide);
    parts[slidePath] = strToU8(normalized);

    const file = await validatePptx(zipSync(parts));
    assert.deepEqual(file.errors, []);
    const observed = file.slides[0]?.objects[0];
    assert.equal(observed?.boxEmu.height, nativeFrameHeight);
    assert.equal(
      observed?.table?.rowHeightsEmu.reduce((sum, height) => sum + height, 0),
      rowHeight,
    );
  });

  it("still rejects invalid table frame extents and grid dimensions", async () => {
    const result = await build(
      tableDeck(
        createElement(
          Table,
          { columns: [100, 100] },
          createElement(
            Row,
            { style: { height: 24 } },
            createElement(Cell, null, "A"),
            createElement(Cell, null, "B"),
          ),
        ),
      ),
    );
    const slidePath = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts[slidePath]!);
    const withSlide = (xml: string) =>
      validatePptx(zipSync({ ...parts, [slidePath]: strToU8(xml) }));

    const zeroHeight = slide.replace(
      /(<p:graphicFrame>[\s\S]*?<a:ext cx="\d+" cy=")\d+("[^>]*>[\s\S]*?<a:tbl>)/u,
      (_match, prefix: string, suffix: string) => `${prefix}0${suffix}`,
    );
    assert.notEqual(zeroHeight, slide);
    assert.equal(
      (await withSlide(zeroHeight)).errors.filter(
        ({ code }) => code === "invalid-table-transform",
      ).length,
      1,
    );

    const zeroWidth = slide.replace(
      /(<p:graphicFrame>[\s\S]*?<a:ext cx=")\d+(" cy="\d+"[^>]*>[\s\S]*?<a:tbl>)/u,
      (_match, prefix: string, suffix: string) => `${prefix}0${suffix}`,
    );
    assert.notEqual(zeroWidth, slide);
    assert.equal(
      (await withSlide(zeroWidth)).errors.filter(
        ({ code }) => code === "invalid-table-transform",
      ).length,
      1,
    );

    const zeroColumn = slide.replace(
      /<a:gridCol w="\d+"\/>/u,
      '<a:gridCol w="0"/>',
    );
    assert.notEqual(zeroColumn, slide);
    assert.equal(
      (await withSlide(zeroColumn)).errors.filter(
        ({ code }) => code === "invalid-table-grid",
      ).length,
      1,
    );

    const zeroRow = slide.replace(/<a:tr h="\d+">/u, '<a:tr h="0">');
    assert.notEqual(zeroRow, slide);
    assert.equal(
      (await withSlide(zeroRow)).errors.filter(
        ({ code }) => code === "invalid-table-grid",
      ).length,
      1,
    );
  });

  it("writes horizontal and vertical merge anchors and continuation cells", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "merged" },
          createElement(
            Table,
            {
              id: "merged-table",
              columns: [50, 50, 50],
              style: {
                fontFamily: font("Noto Sans CJK SC", [
                  { weight: 400, data: notoSansScRegular },
                ]),
                fontSize: 14,
              },
            },
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, { colSpan: 2 }, "Header"),
              createElement(Cell, null, "Top"),
            ),
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, { rowSpan: 2 }, "Side"),
              createElement(Cell, null, "Middle"),
              createElement(Cell, null, "Right"),
            ),
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, null, "Bottom"),
              createElement(Cell, null, "End"),
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual(result.report.slides[0]?.objects[0]?.table?.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
      { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
    ]);
    assert.deepEqual(result.report.slides[0]?.objects[0]?.table?.cells, [
      ["Header", "", "Top"],
      ["Side", "Middle", "Right"],
      ["", "Bottom", "End"],
    ]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.table?.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
      { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
    ]);
    assert.deepEqual(file.slides[0]?.objects[0]?.table?.cells, [
      ["Header", "", "Top"],
      ["Side", "Middle", "Right"],
      ["", "Bottom", "End"],
    ]);
  });

  it("allocates native grid widths across px, Nfr, and auto column tracks", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "tracks" },
          createElement(
            Table,
            {
              id: "tracks",
              x: 10,
              y: 12,
              width: 180,
              columns: [40, "2fr", "auto"],
              style: {
                fontFamily: font("Noto Sans CJK SC", [
                  { weight: 400, data: notoSansScRegular },
                ]),
                fontSize: 14,
              },
            },
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, null, "Fixed"),
              createElement(Cell, null, "Fraction"),
              createElement(Cell, null, "Auto"),
            ),
          ),
        ),
      ),
    );

    const table = result.report.slides[0]?.objects[0]?.table;
    assert.deepEqual(result.report.diagnostics, []);
    assert.equal(table?.columnWidths[0], 40);
    assert.ok((table?.columnWidths[1] ?? 0) > (table?.columnWidths[2] ?? 0));
    assert.ok(
      Math.abs(
        (table?.columnWidths.reduce((sum, width) => sum + width, 0) ?? 0) - 180,
      ) < 0.01,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.table?.cells, [
      ["Fixed", "Fraction", "Auto"],
    ]);
    assert.equal(
      file.slides[0]?.objects[0]?.table?.columnWidthsEmu[0],
      381_000,
    );
  });

  it("requires real font measurement for auto columns and preserves px tracks with warning for overly narrow widths", async () => {
    const autoWithoutMetrics = createElement(
      Table,
      { id: "measured-auto", columns: ["auto"] },
      createElement(
        Row,
        { style: { height: 24 } },
        createElement(Cell, null, "Needs real metrics"),
      ),
    );
    await assert.rejects(
      build(tableDeck(autoWithoutMetrics)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(
          error.diagnostics[0]?.code,
          "font-measurement-unavailable",
        );
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[invalid]/table[measured-auto]/row[1]/cell[1]",
        );
        return true;
      },
    );

    const narrow = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "narrow" },
          createElement(
            Table,
            {
              id: "preserved-px",
              x: 0,
              y: 0,
              width: 120,
              columns: [80, 80],
              style: {
                fontFamily: font("Noto Sans CJK SC", [
                  { weight: 400, data: notoSansScRegular },
                ]),
                fontSize: 14,
              },
            },
            createElement(
              Row,
              { style: { height: 24 } },
              createElement(Cell, null, "A"),
              createElement(Cell, null, "B"),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual(
      narrow.report.slides[0]?.objects[0]?.table?.columnWidths,
      [80, 80],
    );
    assert.deepEqual(
      narrow.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "table-width-overflow",
          path: "deck/slide[narrow]/table[preserved-px]",
        },
      ],
    );
  });

  it("writes graphicFrame tables preserving styles, row/column dimensions, and text", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "comparison" },
          createElement(
            Table,
            {
              id: "plans",
              name: "plans",
              x: 12,
              y: 8,
              columns: [80, 80],
              style: {
                fontFamily: font("Noto Sans CJK SC", [
                  { weight: 400, data: notoSansScRegular },
                ]),
                fontSize: 16,
              },
            },
            createElement(
              Row,
              { style: { height: 28, backgroundColor: "#336699" } },
              createElement(Cell, { style: { color: "#FFFFFF" } }, "方案"),
              createElement(Cell, null, "说明"),
            ),
            createElement(
              Row,
              { style: { height: 28 } },
              createElement(Cell, null, "团队"),
              createElement(Cell, null, "可编辑"),
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual(result.report.slides[0]?.objects[0], {
      id: 2,
      kind: "table",
      name: "plans",
      authorId: "plans",
      box: { x: 12, y: 8, width: 160, height: 56 },
      table: {
        columnWidths: [80, 80],
        rowHeights: [28, 28],
        cells: [
          ["方案", "说明"],
          ["团队", "可编辑"],
        ],
        merges: [],
      },
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const table = file.slides[0]?.objects[0];
    assert.equal(table?.kind, "table");
    assert.equal(table?.name, "plans");
    assert.deepEqual(table?.table?.columnWidthsEmu, [762_000, 762_000]);
    assert.deepEqual(table?.table?.rowHeightsEmu, [266_700, 266_700]);
    assert.deepEqual(table?.table?.cells, [
      ["方案", "说明"],
      ["团队", "可编辑"],
    ]);
    assert.deepEqual(table?.table?.merges, []);
    assert.deepEqual(table?.table?.cellTexts[0]?.text.paragraphs[0]?.runs[0], {
      text: "方案",
      fontFamily: "Noto Sans CJK SC",
      fontSizePt: 12,
      bold: false,
      italic: false,
      color: "#FFFFFF",
    });
  });

  it("escapes table names once and reports build errors on invalid XML characters", async () => {
    const testId = 'Q1 & "R&D" <board>';
    const result = await build(
      tableDeck(
        createElement(
          Table,
          { id: testId, name: testId, columns: [60] },
          createElement(
            Row,
            { style: { height: 40 } },
            createElement(Cell, null, "name"),
          ),
        ),
      ),
    );
    assert.equal(result.report.slides[0]?.objects[0]?.name, testId);
    assert.equal(result.report.slides[0]?.objects[0]?.authorId, testId);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.name, testId);

    await assert.rejects(
      build(
        tableDeck(
          createElement(
            Table,
            { name: "invalid\u0001name", columns: [60] },
            createElement(
              Row,
              { style: { height: 40 } },
              createElement(Cell, null, "name"),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-xml-character"),
    );
  });

  it("executes tableCapability.bind exactly once per Table node during build and encapsulates parsed state", async () => {
    assert.equal("bind" in tableCapability, true);
    assert.equal("parse" in tableCapability, false);
    assert.equal("measureLayout" in tableCapability, false);

    const bindSpy = vi.spyOn(tableCapability, "bind");

    try {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: "single-table-parse" },
            createElement(
              Box,
              { style: { display: "flex", width: 400 } },
              createElement(
                Table,
                {
                  id: "piloted-table",
                  columns: [120, 100],
                  style: { fontSize: 14, color: "#333333" },
                },
                createElement(
                  Row,
                  { style: { height: 30 } },
                  createElement(Cell, null, "Header 1"),
                  createElement(Cell, null, "Header 2"),
                ),
              ),
            ),
          ),
        ),
      );

      assert.equal(
        bindSpy.mock.calls.length,
        1,
        `Expected tableCapability.bind to be called exactly 1 time, got ${bindSpy.mock.calls.length}`,
      );

      const boundHandle = bindSpy.mock.results[0]?.value;
      assert.ok(boundHandle);
      assert.equal(boundHandle.contract.nativeKind, "graphicFrame");
      assert.equal("parsed" in boundHandle, false);

      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("rejects unknown parameters and unknown style keys on Table during parse", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: "invalid-table-param" },
            createElement(
              Table,
              // @ts-expect-error -- testing unknown parameter
              { columns: [100], unknownParam: "forbidden" },
              createElement(Row, null, createElement(Cell, null, "Text")),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
        return true;
      },
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: "invalid-table-style" },
            createElement(
              Table,
              {
                columns: [100],
                // @ts-expect-error -- testing unknown style property
                style: { unknownStyleKey: "forbidden" },
              },
              createElement(Row, null, createElement(Cell, null, "Text")),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-style");
        return true;
      },
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: "invalid-table-padding" },
            createElement(
              Table,
              {
                columns: [100],
                style: { padding: 10 },
              },
              createElement(Row, null, createElement(Cell, null, "Text")),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-layout-context");
        return true;
      },
    );
  });
});
