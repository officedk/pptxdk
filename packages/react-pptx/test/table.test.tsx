import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  build,
  BuildError,
  Cell,
  Deck,
  font,
  P,
  Row,
  Slide,
  Table,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, type ReactNode } from "react";
import { describe, it } from "vitest";

const notoSansScRegular = new Uint8Array(
  readFileSync(
    new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
  ),
);

function tableDeck(content: ReactNode): ReactNode {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { slideId: "invalid" }, content),
  );
}

describe("原生表格", () => {
  it("裁剪 Text 能力时仍通过共享文字服务写出表格", async () => {
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
          'import { Deck, Slide, createBuilder } from "@react-pptx/core";',
          'import { Cell, Row, Table, tableCapability } from "@react-pptx/capabilities/table";',
          'import { validatePptx } from "@react-pptx/validation";',
          "const build = createBuilder([tableCapability]);",
          'const result = await build(createElement(Deck, { width: 160, height: 96 }, createElement(Slide, { slideId: "trimmed" }, createElement(Table, { testId: "only-table", columns: [80] }, createElement(Row, { style: { height: 40 } }, createElement(Cell, null, "table survives Text crop"))))));',
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
  }, 20_000); // 独立进程首次加载 Open XML SDK WASM runtime 约需 10 秒。

  it("让表格和单元格文字样式共用字距、大小写、装饰和阴影服务", async () => {
    const result = await build(
      tableDeck(
        createElement(
          Table,
          {
            testId: "text-style-table",
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

  it("在表格、行或单元格路径定位结构缺口和冲突合并", async () => {
    const incomplete = createElement(
      Table,
      { testId: "short", columns: [60, 60] },
      createElement(
        Row,
        { testId: "missing" },
        createElement(Cell, { testId: "only" }, "One"),
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
      { testId: "wide", columns: [60, 60] },
      createElement(
        Row,
        { testId: "overflow" },
        createElement(Cell, { testId: "too-wide", colSpan: 3 }, "Too wide"),
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
      { testId: "conflict", columns: [60, 60, 60] },
      createElement(
        Row,
        { testId: "first" },
        createElement(Cell, { rowSpan: 2 }, "Left"),
        createElement(Cell, null, "Centre"),
        createElement(Cell, { rowSpan: 2 }, "Right"),
      ),
      createElement(
        Row,
        { testId: "second" },
        createElement(Cell, { testId: "overlap", colSpan: 2 }, "Conflict"),
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
      { testId: "trailing", columns: [60] },
      createElement(
        Row,
        { testId: "last" },
        createElement(Cell, { testId: "too-tall", rowSpan: 2 }, "Too tall"),
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

  it("自动行高按真实文字度量展开，显式过小行高则保留并警告", async () => {
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
          { slideId: "automatic" },
          createElement(
            Table,
            { testId: "automatic-row", columns: [30], style },
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
          { slideId: "spanning-automatic" },
          createElement(
            Table,
            { testId: "spanning-automatic", columns: [32, 32], style },
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
          { slideId: "fixed" },
          createElement(
            Table,
            { testId: "fixed-row", columns: [30], style },
            createElement(
              Row,
              { testId: "short", style: { height: 20 } },
              createElement(
                Cell,
                { testId: "long", style: { padding: 2 } },
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
          { slideId: "explicit-line-height" },
          createElement(
            Table,
            { testId: "explicit-line-height", columns: [60], style },
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

  it("矩形合并在首行和首列续格写出跨度，独立读取器拒绝缺失跨度", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "rectangle-merge" },
          createElement(
            Table,
            { testId: "rectangle-merge", columns: [60, 60] },
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

  it("写入横向和纵向合并锚点及续格", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "merged" },
          createElement(
            Table,
            {
              testId: "merged-table",
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

  it("按 px、Nfr 与 auto 列轨道分配原生网格宽度", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "tracks" },
          createElement(
            Table,
            {
              testId: "tracks",
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

  it("auto 列要求真实字体度量，过窄宽度保留 px 轨道并警告", async () => {
    const autoWithoutMetrics = createElement(
      Table,
      { testId: "measured-auto", columns: ["auto"] },
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
          { slideId: "narrow" },
          createElement(
            Table,
            {
              testId: "preserved-px",
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

  it("写入保留样式、行列尺寸和文字的 graphicFrame 表格", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "comparison" },
          createElement(
            Table,
            {
              testId: "plans",
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
      testId: "plans",
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

  it("表格名称只转义一次并在非法 XML 字符处报告构建错误", async () => {
    const testId = 'Q1 & "R&D" <board>';
    const result = await build(
      tableDeck(
        createElement(
          Table,
          { testId, columns: [60] },
          createElement(
            Row,
            { style: { height: 40 } },
            createElement(Cell, null, "name"),
          ),
        ),
      ),
    );
    assert.equal(result.report.slides[0]?.objects[0]?.name, testId);
    assert.equal(result.report.slides[0]?.objects[0]?.testId, testId);
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
});
