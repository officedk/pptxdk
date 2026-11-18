import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { BoundObjectCapabilityModule } from "@react-pptx/core";
import {
  build,
  BuildError,
  boxCapability,
  Box,
  Cell,
  Chart,
  Deck,
  createBuilder,
  font,
  P,
  Row,
  Shape,
  shapeCapability,
  Slide,
  SlideNumber,
  Span,
  Table,
  Text,
  targets,
  transitions,
  type BuildOptions,
  type FontSize,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import * as fontkit from "fontkit";
import * as harfbuzz from "harfbuzzjs";
import {
  createElement,
  useEffect,
  useLayoutEffect,
  type ComponentType,
} from "react";
import { describe, it } from "vitest";

function deck(shape: ReturnType<typeof createElement>) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "cover" }, shape),
  );
}

function testFont(path: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(path, import.meta.url)));
}

function parsedFont(bytes: Uint8Array): fontkit.Font {
  const parsed = fontkit.create(Buffer.from(bytes));
  if (parsed.type !== "TTF") throw new Error("expected static TrueType font");
  return parsed;
}

function shapedGlyphs(
  bytes: Uint8Array,
  text: string,
): readonly {
  readonly cluster: number;
  readonly path: string;
  readonly xAdvance: number;
  readonly yAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
}[] {
  const blob = new harfbuzz.Blob(bytes);
  const face = new harfbuzz.Face(blob);
  const font = new harfbuzz.Font(face);
  const buffer = new harfbuzz.Buffer();
  buffer.addText(text);
  buffer.guessSegmentProperties();
  harfbuzz.shape(font, buffer);
  const parsed = parsedFont(bytes);
  const infos = buffer.getGlyphInfos();
  const positions = buffer.getGlyphPositions();
  return infos.map((info, index) => {
    const position = positions[index]!;
    return {
      cluster: info.cluster,
      path: parsed.getGlyph(info.codepoint).path.toSVG(),
      xAdvance: position.xAdvance,
      yAdvance: position.yAdvance,
      xOffset: position.xOffset,
      yOffset: position.yOffset,
    };
  });
}

function embeddedFont(
  parts: Record<string, Uint8Array | null>,
  presentation: string,
  family: string,
  style: "regular" | "bold" | "italic" | "boldItalic",
): Uint8Array {
  const record = [
    ...presentation.matchAll(
      /<p:embeddedFont><p:font typeface="([^"]+)"[^>]*\/><p:(regular|bold|italic|boldItalic) r:id="([^"]+)"\/><\/p:embeddedFont>/gu,
    ),
  ].find(
    ([, typeface, fontStyle]) => typeface === family && fontStyle === style,
  );
  assert.ok(record?.[3]);
  const relationships = strFromU8(parts["ppt/_rels/presentation.xml.rels"]!);
  const relation = [
    ...relationships.matchAll(
      /<Relationship Id="([^"]+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font" Target="([^"]+)"\/>/gu,
    ),
  ].find(([, id]) => id === record[3]);
  assert.ok(relation?.[2]);
  const wrapped = parts[`ppt/${relation[2]}`]!;
  const dataLength = new DataView(
    wrapped.buffer,
    wrapped.byteOffset,
    wrapped.byteLength,
  ).getUint32(4, true);
  return wrapped.subarray(wrapped.byteLength - dataLength);
}

function withFontFsType(bytes: Uint8Array, fsType: number): Uint8Array {
  const copy = new Uint8Array(bytes);
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  const tableCount = view.getUint16(4);
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...copy.subarray(record, record + 4));
    if (tag === "OS/2") {
      view.setUint16(view.getUint32(record + 8) + 8, fsType);
      return copy;
    }
  }
  throw new Error("Jost fixture must contain an OS/2 table");
}

const notoSansScRegular = testFont("./assets/fonts/NotoSansSC-Regular.woff2");
const jostBook = testFont("./assets/fonts/Jost-400-Book.ttf");
const jostBold = testFont("./assets/fonts/Jost-700-Bold.ttf");
const jostItalic = testFont("./assets/fonts/Jost-400-Italic.ttf");
const jostBoldItalic = testFont("./assets/fonts/Jost-700-BoldItalic.ttf");
const zcoolXiaoWei = testFont("./assets/fonts/ZCOOLXiaoWei_400Regular.ttf");
const darkerGrotesqueSemiBold = testFont(
  "./assets/fonts/DarkerGrotesque-SemiBold.ttf",
);

describe("public build entry", () => {
  it("embeds TrueType subsets based on actual characters in Text, Table, Chart, text transform casing, and lists", async () => {
    const cjkFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: zcoolXiaoWei },
    ]);
    const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "font-subset" },
          createElement(
            Text,
            {
              id: "subset-text",
              x: 8,
              y: 8,
              width: 640,
              height: 180,
              style: { fontFamily: cjkFont },
            },
            "中文",
            createElement(
              Span,
              { style: { textTransform: "uppercase" } },
              "aß",
            ),
            createElement(
              P,
              { style: { list: { type: "bullet", character: "•" } } },
              "列表项",
            ),
            createElement(
              P,
              { style: { list: { type: "number", startAt: 7 } } },
              "编号项",
            ),
          ),
          createElement(
            Text,
            {
              id: "subset-latin",
              x: 8,
              y: 80,
              width: 300,
              height: 36,
              style: { fontFamily: latinFont },
            },
            "office AV",
          ),
          createElement(
            Table,
            {
              id: "subset-table",
              x: 8,
              y: 128,
              width: 240,
              columns: [240],
              style: { fontFamily: cjkFont },
            },
            createElement(
              Row,
              { style: { height: 28 } },
              createElement(Cell, null, "表格"),
            ),
          ),
          createElement(Chart, {
            id: "subset-chart",
            x: 280,
            y: 128,
            width: 320,
            height: 220,
            style: { fontFamily: cjkFont },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["销售图"] },
                series: [{ name: "金额", values: [12] }],
              },
            ],
          }),
        ),
      ),
      { embedFonts: "subset" },
    );

    assert.deepEqual(result.report.diagnostics, []);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects.length, 4);

    const parts = unzipSync(result.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    const cjkSubsetBytes = embeddedFont(
      parts,
      presentation,
      "ZCOOL XiaoWei",
      "regular",
    );
    const latinSubsetBytes = embeddedFont(
      parts,
      presentation,
      "Jost*",
      "regular",
    );
    const cjkFull = parsedFont(zcoolXiaoWei);
    const cjkSubset = parsedFont(cjkSubsetBytes);
    const latinFull = parsedFont(jostBook);
    const latinSubset = parsedFont(latinSubsetBytes);
    assert.ok(cjkSubset.numGlyphs < cjkFull.numGlyphs);
    assert.ok(latinSubset.numGlyphs < latinFull.numGlyphs);
    assert.ok(cjkSubset.hasGlyphForCodePoint("中".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("表".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("销".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("•".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("7".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint(".".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("A".codePointAt(0)!));
    assert.ok(cjkSubset.hasGlyphForCodePoint("S".codePointAt(0)!));
    assert.equal(cjkSubset.hasGlyphForCodePoint("ß".codePointAt(0)!), false);
    assert.equal(cjkSubset.hasGlyphForCodePoint("a".codePointAt(0)!), false);
    assert.equal(cjkSubset.hasGlyphForCodePoint("Z".codePointAt(0)!), false);
    const expectedCjkCharacters = [
      ...new Set(
        Array.from(
          "中文ASS•列表项0123456789.,-+Ee编号项表格销售图金额",
          (character) => character.codePointAt(0)!,
        ),
      ),
      // The ending segment of cmap format 4 must retain the U+FFFF sentinel.
      0xffff,
    ].sort((left, right) => left - right);
    assert.deepEqual(
      [...cjkSubset.characterSet].sort((left, right) => left - right),
      expectedCjkCharacters,
    );
    assert.ok(latinSubset.availableFeatures.includes("kern"));
    assert.deepEqual(
      shapedGlyphs(zcoolXiaoWei, "中文A7"),
      shapedGlyphs(cjkSubsetBytes, "中文A7"),
    );
    assert.deepEqual(
      shapedGlyphs(jostBook, "office AV"),
      shapedGlyphs(latinSubsetBytes, "office AV"),
    );
    assert.deepEqual(
      Object.keys(parts).filter((path) => path.startsWith("ppt/fonts/")),
      ["ppt/fonts/font1.fntdata", "ppt/fonts/font2.fntdata"],
    );
  });

  it("covers scale ticks, format literals, percentages, and style inheritance for chart-exclusive font subsets", async () => {
    const chartFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: zcoolXiaoWei },
    ]);
    const chartCases = [
      {
        id: "subset-chart-default-axis",
        chart: createElement(Chart, {
          id: "subset-chart-default-axis",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789"),
        absent: [],
      },
      {
        id: "subset-chart-percent-axis",
        chart: createElement(Chart, {
          id: "subset-chart-percent-axis",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "column",
              stacking: "percent",
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789%"),
        absent: [],
      },
      {
        id: "subset-chart-formatted-axis",
        chart: createElement(Chart, {
          id: "subset-chart-formatted-axis",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          axes: {
            category: { kind: "category" },
            value: {
              kind: "value",
              min: 0,
              max: 20,
              majorUnit: 5,
              numberFormat: '#,##0"元"',
              style: { fontSize: 14 },
            },
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789,元"),
        absent: ["#", '"'],
      },
      {
        id: "subset-chart-formatted-label",
        chart: createElement(Chart, {
          id: "subset-chart-formatted-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: '0"元"',
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲", "乙"] },
              series: [{ name: "丙", values: [1, 2] }],
            },
          ],
        }),
        expected: Array.from("0123456789元"),
        absent: ['"'],
      },
      {
        id: "subset-chart-percent-label",
        chart: createElement(Chart, {
          id: "subset-chart-percent-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: { percent: true, style: { fontSize: 14 } },
              data: { kind: "shared-categories", categories: ["甲", "乙"] },
              series: [{ name: "丙", values: [1, 2] }],
            },
          ],
        }),
        expected: Array.from("0123456789%"),
        absent: [],
      },
      {
        id: "subset-chart-locale-only-label",
        chart: createElement(Chart, {
          id: "subset-chart-locale-only-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "[$-409]0",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789"),
        absent: ["$", "[", "]"],
      },
      {
        id: "subset-chart-currency-literal-label",
        chart: createElement(Chart, {
          id: "subset-chart-currency-literal-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "[$￥-804]#,##0.00",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789￥,."),
        absent: ["$", "#", "[", "]"],
      },
      {
        id: "subset-chart-condition-color-label",
        chart: createElement(Chart, {
          id: "subset-chart-condition-color-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: '[Red][>=100]0"元";[Blue][<100]0"分"',
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789元分"),
        absent: Array.from("[]<>RedBlu"),
      },
      {
        id: "subset-chart-escaped-literal-label",
        chart: createElement(Chart, {
          id: "subset-chart-escaped-literal-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "0\\元",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789元"),
        absent: ["\\"],
      },
      {
        id: "subset-chart-chinese-date-name-label",
        chart: createElement(Chart, {
          id: "subset-chart-chinese-date-name-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "[$-804]dddd",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789星期日一二三四五六"),
        absent: ["A", "J", "S"],
      },
      {
        id: "subset-chart-question-placeholder-label",
        chart: createElement(Chart, {
          id: "subset-chart-question-placeholder-label",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "??",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
        expected: Array.from("0123456789 "),
        absent: ["?"],
      },
    ] as const;

    for (const { id, chart, expected, absent } of chartCases) {
      const result = await build(deck(chart), { embedFonts: "subset" });
      assert.deepEqual(result.report.diagnostics, []);
      assert.deepEqual((await validatePptx(result.pptx)).errors, []);

      const parts = unzipSync(result.pptx);
      const subset = parsedFont(
        embeddedFont(
          parts,
          strFromU8(parts["ppt/presentation.xml"]!),
          "ZCOOL XiaoWei",
          "regular",
        ),
      );
      for (const character of expected) {
        assert.ok(
          subset.hasGlyphForCodePoint(character.codePointAt(0)!),
          `${id} subset should cover ${character}`,
        );
      }
      for (const character of absent) {
        assert.equal(
          subset.hasGlyphForCodePoint(character.codePointAt(0)!),
          false,
          `${id} should not treat format control character ${character} as display glyph`,
        );
      }
    }

    const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    await assert.rejects(
      build(
        deck(
          createElement(Chart, {
            id: "subset-chart-label-style-missing-glyph",
            x: 0,
            y: 0,
            width: 320,
            height: 220,
            style: { fontFamily: chartFont },
            groups: [
              {
                kind: "pie",
                dataLabels: {
                  value: true,
                  numberFormat: '0"元"',
                  style: { fontFamily: latinFont },
                },
                data: { kind: "shared-categories", categories: ["甲", "乙"] },
                series: [{ name: "丙", values: [1, 2] }],
              },
            ],
          }),
        ),
        { embedFonts: "subset" },
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, message }) =>
            code === "font-missing-glyph" && message.includes("元"),
        ),
    );
  });

  it("ignores the fourth text-only section for numeric labels", async () => {
    const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    const result = await build(
      deck(
        createElement(Chart, {
          id: "subset-chart-fourth-text-section",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: latinFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: '0;0;0;"元"@',
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "B", values: [12] }],
            },
          ],
        }),
      ),
      { embedFonts: "subset" },
    );
    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    const parts = unzipSync(result.pptx);
    const subset = parsedFont(
      embeddedFont(
        parts,
        strFromU8(parts["ppt/presentation.xml"]!),
        "Jost*",
        "regular",
      ),
    );
    assert.equal(subset.hasGlyphForCodePoint("元".codePointAt(0)!), false);
  });

  it("does not require currency glyphs for locale-only numeric formats", async () => {
    const chartFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: zcoolXiaoWei },
    ]);
    const axisResult = await build(
      deck(
        createElement(Chart, {
          id: "subset-chart-locale-only-source",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
      ),
      { embedFonts: "subset" },
    );
    const axisParts = unzipSync(axisResult.pptx);
    const narrowFontBytes = embeddedFont(
      axisParts,
      strFromU8(axisParts["ppt/presentation.xml"]!),
      "ZCOOL XiaoWei",
      "regular",
    );
    assert.equal(
      parsedFont(narrowFontBytes).hasGlyphForCodePoint("$".codePointAt(0)!),
      false,
    );

    const narrowFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: narrowFontBytes },
    ]);
    const result = await build(
      deck(
        createElement(Chart, {
          id: "subset-chart-locale-only-label-narrow-font",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: narrowFont },
          groups: [
            {
              kind: "pie",
              dataLabels: {
                value: true,
                numberFormat: "[$-409]0",
                style: { fontSize: 14 },
              },
              data: { kind: "shared-categories", categories: ["甲"] },
              series: [{ name: "乙", values: [12] }],
            },
          ],
        }),
      ),
      { embedFonts: "subset" },
    );
    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    const parts = unzipSync(result.pptx);
    const subset = parsedFont(
      embeddedFont(
        parts,
        strFromU8(parts["ppt/presentation.xml"]!),
        "ZCOOL XiaoWei",
        "regular",
      ),
    );
    assert.equal(subset.hasGlyphForCodePoint("$".codePointAt(0)!), false);
    for (const character of "0123456789") {
      assert.ok(subset.hasGlyphForCodePoint(character.codePointAt(0)!));
    }
  });

  it("does not treat newline and tab control characters as missing glyphs in character subsets", async () => {
    const jost = font("Jost*", [{ weight: 400, data: jostBook }]);
    for (const [testId, text] of [
      ["subset-lf", "A\nB"],
      ["subset-crlf", "A\r\nB"],
      ["subset-tab", "A\tB"],
    ] as const) {
      const result = await build(
        deck(
          createElement(
            Text,
            {
              id: testId,
              x: 0,
              y: 0,
              width: 320,
              height: 300,
              style: { fontFamily: jost },
            },
            text,
          ),
        ),
        { embedFonts: "subset" },
      );
      assert.deepEqual(result.report.diagnostics, []);
      assert.deepEqual((await validatePptx(result.pptx)).errors, []);

      const parts = unzipSync(result.pptx);
      const subset = parsedFont(
        embeddedFont(
          parts,
          strFromU8(parts["ppt/presentation.xml"]!),
          "Jost*",
          "regular",
        ),
      );
      assert.ok(subset.hasGlyphForCodePoint("A".codePointAt(0)!));
      assert.ok(subset.hasGlyphForCodePoint("B".codePointAt(0)!));
    }
  });

  it("measures Shape text via shared text service, emits page number fields, and registers actual characters", async () => {
    const pageNumberClaims: string[] = [];
    const observedShapeCapability: BoundObjectCapabilityModule = {
      declaration: shapeCapability.declaration,
      bind(input, context) {
        const bound = shapeCapability.bind(input, context);
        return {
          ...bound,
          compile(compileContext) {
            const output = bound.compile(compileContext);
            pageNumberClaims.push(
              ...(output.resourceClaims ?? []).flatMap(({ characters }) =>
                characters === undefined ? [] : [characters],
              ),
            );
            return output;
          },
        };
      },
    };
    const shapeBuild = createBuilder([boxCapability, observedShapeCapability]);
    const measurementFont = font("Jost", [{ weight: 400, data: jostBook }], {
      substitute: true,
    });
    const pageShapes = (pageNumber: number) =>
      createElement(
        Box,
        {
          style: {
            display: "flex",
            alignItems: "flex-start",
            width: 250,
            height: 60,
          },
        },
        createElement(
          Shape,
          {
            id: `field-shape-${pageNumber}`,
            textStyle: {
              fontFamily: measurementFont,
              fontSize: 20,
              lineHeight: 28,
            },
            style: { width: "auto", height: "auto" },
          },
          createElement(
            P,
            { style: { lineHeight: 28 } },
            "Page ",
            createElement(
              Span,
              { style: { color: "#336699" } },
              createElement(SlideNumber, {}),
            ),
          ),
        ),
        createElement(Shape, {
          id: `literal-shape-${pageNumber}`,
          text: `Page ${pageNumber}`,
          textStyle: {
            fontFamily: measurementFont,
            fontSize: 20,
            lineHeight: 28,
          },
          style: { width: "auto", height: "auto" },
        }),
      );
    const result = await shapeBuild(
      createElement(
        Deck,
        { width: 320, height: 100 },
        ...Array.from({ length: 10 }, (_, index) => {
          const pageNumber = index + 1;
          return createElement(
            Slide,
            { key: pageNumber, id: `shape-field-${pageNumber}` },
            pageNumber === 1 || pageNumber === 10
              ? pageShapes(pageNumber)
              : null,
          );
        }),
      ),
    );

    const fieldShape = (pageNumber: number) =>
      result.report.slides[pageNumber - 1]?.objects.find(
        ({ authorId: testId }) => testId === `field-shape-${pageNumber}`,
      );
    const literalShape = (pageNumber: number) =>
      result.report.slides[pageNumber - 1]?.objects.find(
        ({ authorId: testId }) => testId === `literal-shape-${pageNumber}`,
      );
    assert.equal(fieldShape(1)?.text, "Page 1");
    assert.equal(fieldShape(10)?.text, "Page 10");
    assert.deepEqual(
      [fieldShape(1)?.box?.width, fieldShape(1)?.box?.height],
      [literalShape(1)?.box?.width, literalShape(1)?.box?.height],
    );
    assert.deepEqual(
      [fieldShape(10)?.box?.width, fieldShape(10)?.box?.height],
      [literalShape(10)?.box?.width, literalShape(10)?.box?.height],
    );
    assert.ok(
      (fieldShape(10)?.box?.width ?? 0) > (fieldShape(1)?.box?.width ?? 0),
    );
    assert.ok(
      pageNumberClaims.some(
        (characters) =>
          characters.includes("Page 1") && characters.includes("0123456789"),
      ),
    );
    assert.ok(
      pageNumberClaims.some(
        (characters) =>
          characters.includes("Page 10") && characters.includes("0123456789"),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      [1, 10].map((pageNumber) => {
        const field = file.slides[pageNumber - 1]?.objects[0]?.text;
        const paragraph = field?.paragraphs[0];
        const run = paragraph?.runs.at(-1);
        return [
          field?.text,
          paragraph?.lineSpacingPoints,
          run?.text,
          run?.field,
          run?.color,
        ];
      }),
      [
        ["Page 1", 21, "1", "slidenum", "#336699"],
        ["Page 10", 21, "10", "slidenum", "#336699"],
      ],
    );
    for (const pageNumber of [1, 10]) {
      const xml = strFromU8(
        unzipSync(result.pptx)[`ppt/slides/slide${pageNumber}.xml`]!,
      );
      assert.match(xml, /<a:fld id="\{[0-9A-F-]+\}" type="slidenum">/u);
    }

    await assert.rejects(
      shapeBuild(
        deck(
          createElement(
            Shape,
            {
              id: "conflicting-shape-body",
              x: 10,
              y: 10,
              width: 100,
              height: 40,
              text: "literal",
            },
            createElement(SlideNumber, {}),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "conflicting-text-content" &&
        error.diagnostics[0]?.path ===
          "deck/slide[cover]/shape[conflicting-shape-body]",
    );
  });

  it("includes Text, Shape, and Table page number fields into the same Chart character subset", async () => {
    const pageFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    const textFields = (count: number) =>
      Array.from({ length: count }, (_, index) =>
        createElement(SlideNumber, { key: index }),
      );
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { id: "field-first" },
          createElement(
            Text,
            {
              id: "text-first",
              x: 0,
              y: 0,
              width: 100,
              height: 30,
              style: { fontFamily: pageFont },
            },
            ...textFields(3),
          ),
          createElement(
            Shape,
            {
              id: "shape-first",
              x: 0,
              y: 40,
              width: 100,
              height: 30,
              textStyle: { fontFamily: pageFont },
            },
            createElement(SlideNumber, {}),
          ),
          createElement(
            Table,
            {
              id: "table-first",
              x: 0,
              y: 80,
              columns: [100],
              style: { fontFamily: pageFont },
            },
            createElement(
              Row,
              { style: { height: 30 } },
              createElement(Cell, {}, createElement(SlideNumber, {})),
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "field-second" },
          createElement(
            Text,
            {
              id: "text-second",
              x: 0,
              y: 0,
              width: 100,
              height: 30,
              style: { fontFamily: pageFont },
            },
            ...textFields(2),
          ),
          createElement(
            Shape,
            {
              id: "shape-second",
              x: 0,
              y: 40,
              width: 100,
              height: 30,
              textStyle: { fontFamily: pageFont },
            },
            createElement(SlideNumber, {}),
          ),
          createElement(
            Table,
            {
              id: "table-second",
              x: 0,
              y: 80,
              columns: [50, 50],
              style: { fontFamily: pageFont },
            },
            createElement(
              Row,
              { style: { height: 30 } },
              createElement(Cell, {}, createElement(SlideNumber, {})),
              createElement(Cell, {}, createElement(SlideNumber, {})),
            ),
          ),
          createElement(Chart, {
            id: "page-field-chart",
            x: 110,
            y: 0,
            width: 200,
            height: 220,
            style: { fontFamily: pageFont },
            groups: [
              {
                kind: "pie",
                dataLabels: {
                  value: true,
                  numberFormat: '0;0;0;"元"@',
                },
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "B", values: [12] }],
              },
            ],
          }),
        ),
      ),
      { embedFonts: "subset" },
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const parts = unzipSync(result.pptx);
    const subset = parsedFont(
      embeddedFont(
        parts,
        strFromU8(parts["ppt/presentation.xml"]!),
        "Jost*",
        "regular",
      ),
    );
    assert.ok(subset.numGlyphs < parsedFont(jostBook).numGlyphs);
    for (const character of "0123456789") {
      assert.ok(subset.hasGlyphForCodePoint(character.codePointAt(0)!));
    }
    assert.equal(subset.hasGlyphForCodePoint("元".codePointAt(0)!), false);

    const allIds = file.slides.flatMap((slide, index) => {
      const xml = strFromU8(
        unzipSync(result.pptx)[`ppt/slides/slide${index + 1}.xml`]!,
      );
      return [...xml.matchAll(/<a:fld id="([^"]+)" type="slidenum"/gu)].map(
        ([, id]) => id!,
      );
    });
    assert.equal(allIds.length, 10);
    assert.equal(new Set(allIds).size, allIds.length);
    assert.ok(
      allIds.every((id) =>
        /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/u.test(
          id,
        ),
      ),
    );

    const fieldsInObject = (
      object: (typeof file.slides)[number]["objects"][number],
    ) => [
      ...(object.text?.paragraphs.flatMap((paragraph) =>
        paragraph.runs
          .filter(({ field }) => field === "slidenum")
          .map(({ text }) => text),
      ) ?? []),
      ...(object.table?.cellTexts.flatMap(
        (cell) =>
          cell.text?.paragraphs.flatMap((paragraph) =>
            paragraph.runs
              .filter(({ field }) => field === "slidenum")
              .map(({ text }) => text),
          ) ?? [],
      ) ?? []),
    ];
    assert.deepEqual(
      file.slides.map((slide) => slide.objects.flatMap(fieldsInObject)),
      [Array(5).fill("1"), Array(5).fill("2")],
    );
  });

  it("rejects content-height text with a build error when measurement font is missing", async () => {
    let failure: unknown;
    try {
      await build(
        deck(
          createElement(Text, {
            id: "missing-content-metrics",
            x: 24,
            y: 18,
            width: 240,
            height: "auto",
            style: { fontFamily: "Aptos" },
            children: "需要明确度量字体",
          }),
        ),
      );
    } catch (error) {
      failure = error;
    }

    assert.ok(failure instanceof BuildError);
    assert.deepEqual(
      failure.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "font-measurement-unavailable",
          path: "deck/slide[cover]/text[missing-content-metrics]",
        },
      ],
    );
  });

  it("computes content height from explicit WOFF2 font metrics and emits identical fixed line spacing", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          id: "measured-mixed-text",
          x: 24,
          y: 18,
          width: 400,
          height: "auto",
          style: {
            fontFamily: font("Noto Sans CJK SC", [
              { weight: 400, data: notoSansScRegular },
            ]),
            fontSize: 20,
          },
          children: "中文 A é",
        }),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 24,
      y: 18,
      width: 400,
      height: 28.96,
    });
    assert.deepEqual(result.report.diagnostics, []);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text?.paragraphs[0], {
      text: "中文 A é",
      lineSpacingPoints: 21.72,
      runs: [
        {
          text: "中文 A é",
          fontFamily: "Noto Sans CJK SC",
          fontSizePt: 15,
          bold: false,
          italic: false,
          color: "#000000",
        },
      ],
    });
    assert.deepEqual(
      Object.keys(unzipSync(result.pptx)).filter((path) =>
        path.startsWith("ppt/fonts/"),
      ),
      [],
    );
  });

  it("covers actually used text box, shape, table, and chart font family styles for full embedding", async () => {
    const installedFont = font("Jost*", [
      { weight: 400, data: jostBook },
      { weight: 700, data: jostBold },
      { weight: 600, data: new Uint8Array([1]) },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "font-embedding" },
          createElement(
            Text,
            {
              id: "embedded-font-bold-first",
              x: 24,
              y: 18,
              width: 240,
              height: 36,
              style: { fontFamily: installedFont, fontWeight: "bold" },
            },
            "Bold first",
          ),
          createElement(
            Text,
            {
              id: "embedded-font-text",
              x: 24,
              y: 18,
              width: 240,
              height: 36,
              style: { fontFamily: installedFont },
            },
            "Regular ",
            createElement(Span, { style: { fontWeight: "bold" } }, "bold"),
          ),
          createElement(Shape, {
            id: "embedded-font-shape",
            x: 24,
            y: 64,
            width: 240,
            height: 36,
            text: "Shape text",
            textStyle: { fontFamily: installedFont },
          }),
          createElement(
            Table,
            {
              id: "embedded-font-table",
              x: 24,
              y: 110,
              width: 240,
              columns: [240],
              style: { fontFamily: installedFont, fontSize: 16 },
            },
            createElement(
              Row,
              { style: { height: 28 } },
              createElement(
                Cell,
                { style: { fontWeight: "bold" } },
                "Table text",
              ),
            ),
          ),
          createElement(Chart, {
            id: "embedded-font-chart",
            x: 300,
            y: 18,
            width: 400,
            height: 240,
            style: { fontFamily: installedFont },
            axes: {
              category: { kind: "category" },
              value: { kind: "value", style: { fontWeight: "bold" } },
            },
            groups: [
              {
                kind: "line",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "Sales", values: [12, 18] }],
              },
            ],
          }),
        ),
      ),
      { embedFonts: "full" },
    );

    const parts = unzipSync(result.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    const presentationRelationships = strFromU8(
      parts["ppt/_rels/presentation.xml.rels"]!,
    );
    assert.match(presentation, /embedTrueTypeFonts="1"/u);
    assert.equal((presentation.match(/<p:embeddedFont>/gu) ?? []).length, 1);
    assert.match(
      presentation,
      /<p:font typeface="Jost\*" pitchFamily="2" charset="0"\/>/u,
    );
    assert.match(
      presentation,
      /<p:regular r:id="rId\d+"\/><p:bold r:id="rId\d+"\/>/u,
    );
    assert.doesNotMatch(presentation, /<p:italic(?:\s|\/)/u);
    const fontRelationships = [
      ...presentationRelationships.matchAll(
        /<Relationship Id="(rId\d+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font" Target="([^"]+)"\/>/gu,
      ),
    ];
    assert.equal(fontRelationships.length, 2);
    for (const [style, expectedBytes, expectedWeight] of [
      ["regular", jostBook, 400],
      ["bold", jostBold, 700],
    ] as const) {
      const embeddedStyle = new RegExp(
        `<p:${style} r:id="(rId\\d+)"\\/>`,
        "u",
      ).exec(presentation);
      assert.ok(embeddedStyle?.[1]);
      const relationship = fontRelationships.find(
        ([, relationshipId]) => relationshipId === embeddedStyle[1],
      );
      assert.ok(relationship?.[2]);
      const fontPart = parts[`ppt/${relationship[2]}`]!;
      const view = new DataView(
        fontPart.buffer,
        fontPart.byteOffset,
        fontPart.byteLength,
      );
      const fontDataLength = view.getUint32(4, true);
      assert.equal(view.getUint32(8, true), 0x00020002);
      assert.equal(view.getUint32(28, true), expectedWeight);
      assert.equal(view.getUint16(32, true), 0);
      assert.equal(fontDataLength, expectedBytes.byteLength);
      assert.deepEqual(
        fontPart.subarray(fontPart.byteLength - fontDataLength),
        expectedBytes,
      );
    }
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("writes 600 weight SemiBold into regular slot according to OS/2 style linking", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "embedded-semibold",
            x: 24,
            y: 18,
            width: 240,
            height: 36,
            style: {
              fontFamily: font("Darker Grotesque SemiBold", [
                { weight: 600, data: darkerGrotesqueSemiBold },
              ]),
              fontWeight: 600,
            },
          },
          "SemiBold text",
        ),
      ),
      { embedFonts: "full" },
    );
    const parts = unzipSync(result.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    assert.match(
      presentation,
      /<p:embeddedFont><p:font typeface="Darker Grotesque SemiBold"[^>]*\/><p:regular r:id="rId\d+"\/><\/p:embeddedFont>/u,
    );
    assert.doesNotMatch(presentation, /<p:bold r:id/u);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.bold,
      false,
    );
    const fontPartPath = Object.keys(parts).find((path) =>
      path.startsWith("ppt/fonts/"),
    );
    assert.ok(fontPartPath);
    const fontPart = parts[fontPartPath]!;
    const view = new DataView(
      fontPart.buffer,
      fontPart.byteOffset,
      fontPart.byteLength,
    );
    assert.equal(view.getUint32(28, true), 600);
    assert.deepEqual(
      fontPart.subarray(fontPart.byteLength - view.getUint32(4, true)),
      darkerGrotesqueSemiBold,
    );
  });

  it("maps real italic and bold-italic fonts to corresponding style slots for full embedding", async () => {
    const jostItalicFaces = font("Jost", [
      { weight: 400, style: "italic", data: jostItalic },
      { weight: 700, style: "italic", data: jostBoldItalic },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "italic-font-slots" },
          createElement(
            Text,
            {
              id: "italic-font",
              x: 24,
              y: 18,
              width: 240,
              height: 36,
              style: { fontFamily: jostItalicFaces, fontStyle: "italic" },
            },
            "Italic",
          ),
          createElement(
            Text,
            {
              id: "bold-italic-font",
              x: 24,
              y: 60,
              width: 240,
              height: 36,
              style: {
                fontFamily: jostItalicFaces,
                fontWeight: "bold",
                fontStyle: "italic",
              },
            },
            "Bold italic",
          ),
        ),
      ),
      { embedFonts: "full" },
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) => ({
        run: text?.paragraphs[0]?.runs[0]
          ? {
              text: text.paragraphs[0].runs[0].text,
              bold: text.paragraphs[0].runs[0].bold,
              italic: text.paragraphs[0].runs[0].italic,
            }
          : undefined,
      })),
      [
        {
          run: { text: "Italic", bold: false, italic: true },
        },
        {
          run: { text: "Bold italic", bold: true, italic: true },
        },
      ],
    );

    const parts = unzipSync(result.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    assert.match(presentation, /embedTrueTypeFonts="1"/u);
    assert.match(presentation, /<p:italic r:id="rId\d+"\/>/u);
    assert.match(presentation, /<p:boldItalic r:id="rId\d+"\/>/u);
    const presentationRelationships = strFromU8(
      parts["ppt/_rels/presentation.xml.rels"]!,
    );
    const fontRelationships = [
      ...presentationRelationships.matchAll(
        /<Relationship Id="(rId\d+)" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/font" Target="([^"]+)"\/>/gu,
      ),
    ];
    assert.equal(fontRelationships.length, 2);
    for (const [style, expectedBytes, expectedWeight] of [
      ["italic", jostItalic, 400],
      ["boldItalic", jostBoldItalic, 700],
    ] as const) {
      const embeddedStyle = new RegExp(
        `<p:${style} r:id="(rId\\d+)"\\/>`,
        "u",
      ).exec(presentation);
      assert.ok(embeddedStyle?.[1]);
      const relationship = fontRelationships.find(
        ([, relationshipId]) => relationshipId === embeddedStyle[1],
      );
      assert.ok(relationship?.[2]);
      const fontPart = parts[`ppt/${relationship[2]}`]!;
      const view = new DataView(
        fontPart.buffer,
        fontPart.byteOffset,
        fontPart.byteLength,
      );
      const fontDataLength = view.getUint32(4, true);
      assert.equal(view.getUint32(8, true), 0x00020002);
      assert.equal(view.getUint32(28, true), expectedWeight);
      assert.equal(fontDataLength, expectedBytes.byteLength);
      assert.deepEqual(
        fontPart.subarray(fontPart.byteLength - fontDataLength),
        expectedBytes,
      );
    }
  });

  it("keeps embedding restrictions in identical text unaffected by same-byte deduplication and run order", async () => {
    const policies = [
      {
        label: "not-embeddable",
        options: { embeddable: false },
        code: "font-not-embeddable",
      },
      {
        label: "substitute",
        options: { substitute: true },
        code: "font-substitute-not-embeddable",
      },
    ] as const;

    for (const { label, options, code } of policies) {
      const allowed = font("Jost*", [{ weight: 400, data: jostBook }]);
      const restricted = font(
        "Jost*",
        [{ weight: 400, data: jostBook }],
        options,
      );
      for (const [order, fontValues] of [
        ["allowed-first", [allowed, restricted]],
        ["restricted-first", [restricted, allowed]],
      ] as const) {
        const testId = `font-policy-${label}-${order}`;
        const spans = fontValues.map((fontValue, index) =>
          createElement(
            Span,
            { key: index, style: { fontFamily: fontValue } },
            `run ${index}`,
          ),
        );
        await assert.rejects(
          build(
            deck(
              createElement(
                Text,
                {
                  id: testId,
                  x: 24,
                  y: 18,
                  width: 240,
                  height: 36,
                },
                ...spans,
              ),
            ),
            { embedFonts: "full" },
          ),
          (error: unknown) => {
            assert.ok(error instanceof BuildError);
            assert.equal(error.diagnostics[0]?.code, code);
            assert.equal(
              error.diagnostics[0]?.path,
              `deck/slide[cover]/text[${testId}]`,
            );
            return true;
          },
        );
      }
    }
  });

  it("still merges identical allowed font bytes, while default build still does not embed fonts", async () => {
    const firstAllowed = font("Jost*", [{ weight: 400, data: jostBook }]);
    const secondAllowed = font("Jost*", [{ weight: 400, data: jostBook }]);
    const full = await build(
      deck(
        createElement(
          Text,
          { id: "same-font-bytes", x: 24, y: 18, width: 240, height: 36 },
          createElement(Span, { style: { fontFamily: firstAllowed } }, "first"),
          createElement(
            Span,
            { style: { fontFamily: secondAllowed } },
            "second",
          ),
        ),
      ),
      { embedFonts: "full" },
    );
    const parts = unzipSync(full.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    assert.equal((presentation.match(/<p:embeddedFont>/gu) ?? []).length, 1);
    assert.equal((presentation.match(/<p:regular r:id=/gu) ?? []).length, 1);
    assert.equal(
      Object.keys(parts).filter((path) => path.startsWith("ppt/fonts/")).length,
      1,
    );

    const notRequested = await build(
      deck(
        createElement(
          Text,
          {
            id: "default-font-policy",
            x: 24,
            y: 18,
            width: 240,
            height: 36,
          },
          createElement(
            Span,
            { style: { fontFamily: firstAllowed } },
            "allowed",
          ),
          createElement(
            Span,
            {
              style: {
                fontFamily: font("Jost*", [{ weight: 400, data: jostBook }], {
                  embeddable: false,
                }),
              },
            },
            "not embedded",
          ),
        ),
      ),
    );
    assert.deepEqual(
      Object.keys(unzipSync(notRequested.pptx)).filter((path) =>
        path.startsWith("ppt/fonts/"),
      ),
      [],
    );
    assert.deepEqual((await validatePptx(notRequested.pptx)).errors, []);
  });

  it("accepts Editable and no-subsetting flags for full embedding, rejecting subsetting requests by license", async () => {
    const original = jostBook;
    const noSubsetting = withFontFsType(original, 0x0108);
    const result = await build(
      <Deck>
        <Slide id="font-no-subsetting">
          <Text
            x={0}
            y={0}
            width={100}
            height={32}
            style={{
              fontFamily: font("Jost*", [{ weight: 400, data: noSubsetting }]),
            }}
          >
            Full font
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );
    const parts = unzipSync(result.pptx);
    const embedded = parts["ppt/fonts/font1.fntdata"];
    assert.ok(embedded);
    const view = new DataView(embedded.buffer, embedded.byteOffset);
    const payloadStart = view.getUint32(0, true) - original.byteLength;
    assert.equal(view.getUint16(32, true), 0x0108);
    assert.deepEqual(embedded.subarray(payloadStart), noSubsetting);

    await assert.rejects(
      build(
        <Deck>
          <Slide id="font-no-subsetting-subset">
            <Text
              id="font-no-subsetting-subset"
              x={0}
              y={0}
              width={100}
              height={32}
              style={{
                fontFamily: font("Jost*", [
                  { weight: 400, data: noSubsetting },
                ]),
              }}
            >
              Subset font
            </Text>
          </Slide>
        </Deck>,
        { embedFonts: "subset" },
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "font-subsetting-forbidden");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[font-no-subsetting-subset]/text[font-no-subsetting-subset]",
        );
        return true;
      },
    );
  });

  it("does not allow full embedding requests to bypass missing resources, unsupported formats, or restricted permissions", async () => {
    const cases = [
      {
        label: "font-value-opt-out",
        fontValue: font("Jost*", [{ weight: 400, data: jostBook }], {
          embeddable: false,
        }),
        code: "font-not-embeddable",
      },
      {
        label: "missing-resource",
        fontValue: font("Jost*"),
        code: "font-embedding-resource-missing",
      },
      {
        label: "substitute-resource",
        fontValue: font("Other Family", [{ weight: 400, data: jostBook }], {
          substitute: true,
        }),
        code: "font-substitute-not-embeddable",
      },
      {
        label: "unsupported-woff2",
        fontValue: font("Noto Sans CJK SC", [
          { weight: 400, data: notoSansScRegular },
        ]),
        code: "unsupported-font-format",
      },
      {
        label: "wrong-weight-resource",
        fontValue: font("Jost*", [{ weight: 400, data: jostBold }]),
        code: "font-style-mismatch",
      },
      {
        label: "restricted-license",
        fontValue: font("Jost*", [
          { weight: 400, data: withFontFsType(jostBook, 0x0002) },
        ]),
        code: "font-embedding-restricted",
      },
      {
        label: "preview-print-license",
        fontValue: font("Jost*", [
          { weight: 400, data: withFontFsType(jostBook, 0x0004) },
        ]),
        code: "font-embedding-preview-print",
      },
      {
        label: "bitmap-only-license",
        fontValue: font("Jost*", [
          { weight: 400, data: withFontFsType(jostBook, 0x0200) },
        ]),
        code: "font-embedding-bitmap-only",
      },
      {
        label: "invalid-license-flags",
        fontValue: font("Jost*", [
          { weight: 400, data: withFontFsType(jostBook, 0x000a) },
        ]),
        code: "invalid-font-permissions",
      },
    ] as const;

    for (const { label, fontValue, code } of cases) {
      await assert.rejects(
        build(
          deck(
            createElement(
              Text,
              {
                id: label,
                x: 24,
                y: 18,
                width: 240,
                height: 36,
                style: { fontFamily: fontValue },
              },
              "Embed me",
            ),
          ),
          { embedFonts: "full" },
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          assert.equal(
            error.diagnostics[0]?.path,
            `deck/slide[cover]/text[${label}]`,
          );
          return true;
        },
      );
    }

    await assert.rejects(
      build(
        deck(
          createElement(
            Text,
            {
              id: "subset-unsupported-woff2",
              x: 24,
              y: 18,
              width: 240,
              height: 36,
              style: {
                fontFamily: font("Noto Sans CJK SC", [
                  { weight: 400, data: notoSansScRegular },
                ]),
              },
            },
            "中文",
          ),
        ),
        { embedFonts: "subset" },
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-font-format");
        return true;
      },
    );

    assert.equal(parsedFont(zcoolXiaoWei).hasGlyphForCodePoint(0x1f984), false);
    await assert.rejects(
      build(
        deck(
          createElement(
            Text,
            {
              id: "subset-missing-glyph",
              x: 24,
              y: 18,
              width: 240,
              height: 36,
              style: {
                fontFamily: font("ZCOOL XiaoWei", [
                  { weight: 400, data: zcoolXiaoWei },
                ]),
              },
            },
            "🦄",
          ),
        ),
        { embedFonts: "subset" },
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "font-missing-glyph");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[cover]/text[subset-missing-glyph]",
        );
        return true;
      },
    );

    const notRequested = await build(
      deck(
        createElement(
          Text,
          {
            id: "font-value-opt-out-default",
            x: 24,
            y: 18,
            width: 240,
            height: 36,
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }], {
                embeddable: false,
              }),
            },
          },
          "Measure only",
        ),
      ),
    );
    assert.deepEqual(
      Object.keys(unzipSync(notRequested.pptx)).filter((path) =>
        path.startsWith("ppt/fonts/"),
      ),
      [],
    );
    await assert.rejects(
      build(deck(createElement(Shape, { x: 0, y: 0, width: 32, height: 32 })), {
        embedFonts: "partial" as "full",
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-font-embedding");
        return true;
      },
    );
  });

  it("measures Chinese and overly long words according to Unicode line break opportunities and grapheme boundaries", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(Text, {
            id: "cjk-breaks",
            x: 0,
            y: 0,
            width: 20,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "中文",
          }),
          createElement(Text, {
            id: "long-word-breaks",
            x: 40,
            y: 0,
            width: 1,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "abcdef",
          }),
          createElement(Text, {
            id: "combining-grapheme",
            x: 80,
            y: 0,
            width: 1,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "é",
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        height: box?.height,
      })),
      [
        { id: "cjk-breaks", height: 57.92 },
        { id: "long-word-breaks", height: 173.76 },
        { id: "combining-grapheme", height: 28.96 },
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        { code: "overset", path: "deck/slide[cover]/text[long-word-breaks]" },
        { code: "overset", path: "deck/slide[cover]/text[combining-grapheme]" },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) => text?.text),
      ["中文", "abcdef", "é"],
    );
  });

  it("counts ultra-wide graphemes and explicit newlines as single lines each and preserves trailing empty lines", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(Text, {
            id: "wide-grapheme-hard-break",
            x: 0,
            y: 0,
            width: 1,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "a\nb",
          }),
          createElement(Text, {
            id: "trailing-hard-break",
            x: 40,
            y: 0,
            width: 100,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "a\n",
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ box }) => box?.height),
      [57.92, 57.92],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) => text?.text),
      ["a\nb", "a\n"],
    );
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) =>
        text?.paragraphs.map(({ lineSpacingPoints }) => lineSpacingPoints),
      ),
      [[21.72], [21.72]],
    );
  });

  it("requires explicit proxy font declaration and preserves original output font family name", async () => {
    const alternateFace = font(
      "Deck Display Name",
      [{ weight: 400, data: notoSansScRegular }],
      { substitute: true },
    );
    const result = await build(
      deck(
        createElement(Text, {
          id: "substituted-font",
          x: 0,
          y: 0,
          width: 100,
          height: "auto",
          style: { fontFamily: alternateFace, fontSize: 20 },
          children: "中文",
        }),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontFamily,
      "Deck Display Name",
    );

    await assert.rejects(
      build(
        deck(
          createElement(Text, {
            x: 0,
            y: 0,
            width: 100,
            height: "auto",
            style: {
              fontFamily: font("Deck Display Name", [
                { weight: 400, data: notoSansScRegular },
              ]),
              fontSize: 20,
            },
            children: "中文",
          }),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "font-substitution-required",
        ),
    );
  });

  it("content-sized text records warning without embedFonts, and fails with font-missing-glyph when embedding", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          id: "content-sized-missing-glyph",
          x: 0,
          y: 0,
          width: 100,
          height: "auto",
          style: {
            fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
            fontSize: 20,
          },
          children: "中文",
        }),
      ),
    );
    assert.ok(result.pptx.byteLength > 0);
    const warning = result.report.diagnostics.find(
      ({ code }) => code === "font-missing-glyph",
    );
    assert.ok(
      warning !== undefined,
      "expected font-missing-glyph warning in report",
    );
    assert.equal(warning.severity, "warning");
    assert.equal(
      warning.path,
      "deck/slide[cover]/text[content-sized-missing-glyph]",
    );
    assert.ok(warning.message.includes("中"));
    assert.ok(warning.message.includes("文"));

    await assert.rejects(
      build(
        deck(
          createElement(Text, {
            id: "content-sized-missing-glyph-full",
            x: 0,
            y: 0,
            width: 100,
            height: "auto",
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
              fontSize: 20,
            },
            children: "中文",
          }),
        ),
        { embedFonts: "full" },
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );

    await assert.rejects(
      build(
        deck(
          createElement(Text, {
            id: "content-sized-missing-glyph-subset",
            x: 0,
            y: 0,
            width: 100,
            height: "auto",
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
              fontSize: 20,
            },
            children: "中文",
          }),
        ),
        { embedFonts: "subset" },
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );
  });

  describe("unified glyph coverage check", () => {
    it("reports font-missing-glyph error for fixed-height Text and Shape when embedFonts is full or subset", async () => {
      // Fixed-height Text with full embedding
      await assert.rejects(
        build(
          deck(
            createElement(Text, {
              id: "fixed-text-full-missing-glyph",
              x: 0,
              y: 0,
              width: 100,
              height: 50,
              style: {
                fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
                fontSize: 20,
              },
              children: "中文",
            }),
          ),
          { embedFonts: "full" },
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            ({ code, path }) =>
              code === "font-missing-glyph" &&
              path === "deck/slide[cover]/text[fixed-text-full-missing-glyph]",
          ),
      );

      // Fixed-height Shape with full embedding
      await assert.rejects(
        build(
          deck(
            createElement(Shape, {
              id: "fixed-shape-full-missing-glyph",
              x: 0,
              y: 0,
              width: 100,
              height: 50,
              style: {
                fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
                fontSize: 20,
              },
              text: "中文",
            }),
          ),
          { embedFonts: "full" },
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            ({ code, path }) =>
              code === "font-missing-glyph" &&
              path ===
                "deck/slide[cover]/shape[fixed-shape-full-missing-glyph]",
          ),
      );

      // Fixed-height Shape with subset embedding
      await assert.rejects(
        build(
          deck(
            createElement(Shape, {
              id: "fixed-shape-subset-missing-glyph",
              x: 0,
              y: 0,
              width: 100,
              height: 50,
              style: {
                fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
                fontSize: 20,
              },
              text: "中文",
            }),
          ),
          { embedFonts: "subset" },
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            ({ code, path }) =>
              code === "font-missing-glyph" &&
              path ===
                "deck/slide[cover]/shape[fixed-shape-subset-missing-glyph]",
          ),
      );
    });

    it("records font-missing-glyph warning and produces PPTX when embedding is not enabled", async () => {
      // Fixed-height Text without embedFonts
      const textResult = await build(
        deck(
          createElement(Text, {
            id: "fixed-text-warn-missing-glyph",
            x: 0,
            y: 0,
            width: 100,
            height: 50,
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
              fontSize: 20,
            },
            children: "中文",
          }),
        ),
      );
      assert.ok(textResult.pptx.byteLength > 0);
      const textWarning = textResult.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.ok(textWarning !== undefined);
      assert.equal(
        textWarning.path,
        "deck/slide[cover]/text[fixed-text-warn-missing-glyph]",
      );
      assert.ok(textWarning.message.includes("中"));
      assert.ok(textWarning.message.includes("文"));

      // Fixed-height Shape without embedFonts
      const shapeResult = await build(
        deck(
          createElement(Shape, {
            id: "fixed-shape-warn-missing-glyph",
            x: 0,
            y: 0,
            width: 100,
            height: 50,
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
              fontSize: 20,
            },
            text: "中文",
          }),
        ),
      );
      assert.ok(shapeResult.pptx.byteLength > 0);
      const shapeWarning = shapeResult.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.ok(shapeWarning !== undefined);
      assert.equal(
        shapeWarning.path,
        "deck/slide[cover]/shape[fixed-shape-warn-missing-glyph]",
      );
      assert.ok(shapeWarning.message.includes("中"));
      assert.ok(shapeWarning.message.includes("文"));
    });

    it("does not report font-missing-glyph when only font family name is declared without font bytes", async () => {
      const result = await build(
        deck(
          createElement(Text, {
            id: "name-only-text",
            x: 0,
            y: 0,
            width: 100,
            height: 50,
            style: {
              fontFamily: "CustomFamilyName",
              fontSize: 20,
            },
            children: "中文",
          }),
        ),
      );
      assert.ok(result.pptx.byteLength > 0);
      const missingGlyph = result.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.equal(missingGlyph, undefined);
      const unavailable = result.report.diagnostics.find(
        ({ code }) => code === "font-measurement-unavailable",
      );
      assert.ok(unavailable !== undefined);
    });

    it("does not trigger font-missing-glyph for control characters and Default_Ignorable_Code_Point", async () => {
      // Text with control characters (\r, \n, \t) and ignorable characters (ZWS \u200B, SHY \u00AD, BOM \uFEFF, VS16 \uFE0F)
      const result = await build(
        deck(
          createElement(Text, {
            id: "ignorable-chars-text",
            x: 0,
            y: 0,
            width: 200,
            height: 50,
            style: {
              fontFamily: font("Jost*", [{ weight: 400, data: jostBook }]),
              fontSize: 20,
            },
            children: "Hello\r\n\tWorld\u200B\u00AD\uFEFF\uFE0F",
          }),
        ),
        { embedFonts: "full" },
      );
      assert.ok(result.pptx.byteLength > 0);
      const missingGlyph = result.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.equal(missingGlyph, undefined);
    });

    it("reports font-missing-glyph error for Table when embedFonts is full, and warning without embedding", async () => {
      const jost = font("Jost*", [{ weight: 400, data: jostBook }]);

      await assert.rejects(
        build(
          deck(
            createElement(
              Table,
              {
                id: "table-full-missing-glyph",
                x: 0,
                y: 0,
                columns: [100],
                style: { fontFamily: jost, fontSize: 14 },
              },
              createElement(
                Row,
                { id: "r1" },
                createElement(Cell, { id: "c1" }, "中文"),
              ),
            ),
          ),
          { embedFonts: "full" },
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            ({ code, path }) =>
              code === "font-missing-glyph" &&
              path ===
                "deck/slide[cover]/table[table-full-missing-glyph]/row[r1]/cell[c1]",
          ),
      );

      const tableResult = await build(
        deck(
          createElement(
            Table,
            {
              id: "table-warn-missing-glyph",
              x: 0,
              y: 0,
              columns: [100],
              style: { fontFamily: jost, fontSize: 14 },
            },
            createElement(
              Row,
              { id: "r1" },
              createElement(Cell, { id: "c1" }, "中文"),
            ),
          ),
        ),
      );
      assert.ok(tableResult.pptx.byteLength > 0);
      const tableWarning = tableResult.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.ok(tableWarning !== undefined);
      assert.equal(
        tableWarning.path,
        "deck/slide[cover]/table[table-warn-missing-glyph]/row[r1]/cell[c1]",
      );
    });

    it("reports font-missing-glyph error for Chart when embedFonts is full, and warning without embedding", async () => {
      const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);

      await assert.rejects(
        build(
          deck(
            createElement(Chart, {
              id: "chart-full-missing-glyph",
              x: 0,
              y: 0,
              width: 320,
              height: 220,
              style: { fontFamily: latinFont },
              groups: [
                {
                  kind: "pie",
                  data: { kind: "shared-categories", categories: ["甲", "乙"] },
                  series: [{ name: "丙", values: [1, 2] }],
                },
              ],
            }),
          ),
          { embedFonts: "full" },
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            ({ code, path }) =>
              code === "font-missing-glyph" &&
              path === "deck/slide[cover]/chart[chart-full-missing-glyph]",
          ),
      );

      const chartResult = await build(
        deck(
          createElement(Chart, {
            id: "chart-warn-missing-glyph",
            x: 0,
            y: 0,
            width: 320,
            height: 220,
            style: { fontFamily: latinFont },
            groups: [
              {
                kind: "pie",
                data: { kind: "shared-categories", categories: ["甲", "乙"] },
                series: [{ name: "丙", values: [1, 2] }],
              },
            ],
          }),
        ),
      );
      assert.ok(chartResult.pptx.byteLength > 0);
      const chartWarning = chartResult.report.diagnostics.find(
        ({ code }) => code === "font-missing-glyph",
      );
      assert.ok(chartWarning !== undefined);
      assert.equal(
        chartWarning.path,
        "deck/slide[cover]/chart[chart-warn-missing-glyph]",
      );
    });
  });

  it("selects measurement file by exact weight and requires actual italic weight resources", async () => {
    const jost = font("Jost*", [
      { weight: 400, data: jostBook },
      { weight: 700, data: jostBold },
    ]);
    const result = await build(
      deck(
        createElement(
          Text,
          {
            x: 0,
            y: 0,
            width: 240,
            height: "auto",
            style: { fontFamily: jost, fontSize: 20 },
          },
          "Book ",
          createElement(Span, { style: { fontWeight: "bold" } }, "Bold"),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs.map(
        ({ fontFamily, bold }) => ({ fontFamily, bold }),
      ),
      [
        { fontFamily: "Jost*", bold: false },
        { fontFamily: "Jost*", bold: true },
      ],
    );

    await assert.rejects(
      build(
        deck(
          createElement(Text, {
            x: 0,
            y: 0,
            width: 240,
            height: "auto",
            style: { fontFamily: jost, fontStyle: "italic", fontSize: 20 },
            children: "Italic",
          }),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "font-measurement-unavailable",
        ),
    );
  });

  it("writes weight mapping according to real regular style of standalone SemiBold family", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          id: "semibold-family-regular-face",
          x: 24,
          y: 18,
          width: 240,
          height: "auto",
          style: {
            fontFamily: font("Darker Grotesque SemiBold", [
              { weight: 600, data: darkerGrotesqueSemiBold },
            ]),
            fontSize: 20,
            fontWeight: 600,
          },
          children: "SemiBold family",
        }),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0], {
      text: "SemiBold family",
      fontFamily: "Darker Grotesque SemiBold",
      fontSizePt: 15,
      bold: false,
      italic: false,
      color: "#000000",
    });
  });

  it("accepts non-RIBBI weights via typographic family and preserves native weight mapping", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          id: "preferred-family-face",
          x: 24,
          y: 18,
          width: 240,
          height: "auto",
          style: {
            fontFamily: font("Darker Grotesque", [
              { weight: 600, data: darkerGrotesqueSemiBold },
            ]),
            fontSize: 20,
            fontWeight: 600,
          },
          children: "SemiBold preferred family",
        }),
      ),
      { embedFonts: "full" },
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontFamily,
      "Darker Grotesque",
    );
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.bold,
      false,
    );
  });

  it("warns when measured text overflows fixed bounding box, preserving box, content, and native default line spacing", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "fixed-overset",
            x: 24,
            y: 18,
            width: 20,
            height: 20,
            style: {
              fontFamily: font("Noto Sans CJK SC", [
                { weight: 400, data: notoSansScRegular },
              ]),
              fontSize: 20,
            },
          },
          "中文中文",
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 24,
      y: 18,
      width: 20,
      height: 20,
    });
    assert.equal(result.report.slides[0]?.objects[0]?.text, "中文中文");
    assert.deepEqual(
      result.report.diagnostics.map(({ code }) => code),
      ["overset"],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.text?.text, "中文中文");
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.lineSpacingPoints,
      undefined,
    );
  });

  it("shares fixed bounding box overflow diagnostics across shapes with text and text boxes", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(Text, {
            id: "overset-text",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "中文中文",
          }),
          createElement(Shape, {
            id: "overset-shape",
            x: 40,
            y: 0,
            width: 20,
            height: 20,
            text: "中文中文",
            textStyle: { fontFamily: measureFont, fontSize: 20 },
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        { code: "overset", path: "deck/slide[cover]/text[overset-text]" },
        { code: "overset", path: "deck/slide[cover]/shape[overset-shape]" },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ box }) => box),
      [
        { x: 0, y: 0, width: 20, height: 20 },
        { x: 40, y: 0, width: 20, height: 20 },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text }) => text),
      ["中文中文", "中文中文"],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ text }) => text?.text),
      ["中文中文", "中文中文"],
    );
  });
  it("emits slide color map overrides prior to basic and extended transitions", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "fade-order",
          transition: transitions.fade(),
        }),
        createElement(Slide, {
          id: "flip-order",
          transition: transitions.flip({ fallback: transitions.fade() }),
        }),
        createElement(Slide, {
          id: "morph-order",
          transition: transitions.morph({ fallback: transitions.push() }),
        }),
      ),
    );
    const parts = unzipSync(result.pptx);

    for (let index = 0; index < 3; index += 1) {
      const part = parts[`ppt/slides/slide${index + 1}.xml`];
      assert.ok(part);
      const xml = strFromU8(part);
      const slideContentEnd = xml.indexOf("</p:cSld>");
      const colorMapStart = xml.indexOf("<p:clrMapOvr>");
      const transitionStart = xml.indexOf("<p:transition");
      assert.ok(slideContentEnd >= 0);
      assert.ok(colorMapStart > slideContentEnd);
      assert.ok(transitionStart > colorMapStart);
    }
  });

  it("concurrently validates normal shape and text artifacts", async () => {
    const shapeResult = await build(
      deck(
        createElement(Shape, {
          id: "concurrent-shape",
          x: 0,
          y: 0,
          width: 120,
          height: 72,
        }),
      ),
    );
    const textResult = await build(
      deck(
        createElement(
          Text,
          {
            x: 0,
            y: 0,
            width: 240,
            height: 96,
          },
          "并发验证",
        ),
      ),
    );

    const files = await Promise.all([
      validatePptx(shapeResult.pptx),
      validatePptx(textResult.pptx),
    ]);

    assert.deepEqual(
      files.map(({ errors }) => errors),
      [[], []],
    );
    assert.deepEqual(
      files.map(({ presentation }) => presentation.slideCount),
      [1, 1],
    );

    const sequentialFile = await validatePptx(shapeResult.pptx);
    assert.deepEqual(sequentialFile.errors, []);

    const invalidFile = await validatePptx(new Uint8Array([0]));
    assert.ok(
      invalidFile.errors.some(({ code }) => code === "schema-validator-failed"),
    );
    assert.ok(
      invalidFile.errors.some(({ code }) => code === "invalid-package"),
    );

    const fileAfterFailure = await validatePptx(textResult.pptx);
    assert.deepEqual(fileAfterFailure.errors, []);
  });

  it("generates fixed bounding box text and preserves Chinese, English, and Unicode content", async () => {
    const content = "固定区域文字 — React + 简体中文 😀";
    const result = await build(
      deck(
        createElement(Text, {
          id: "mixed-text",
          x: 32,
          y: 24,
          width: 360,
          height: 96,
          style: { fontFamily: "Aptos" },
          children: content,
        }),
      ),
    );

    assert.equal(result.report.slides[0]?.objects[0]?.kind, "text");
    assert.equal(result.report.slides[0]?.objects[0]?.text, content);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text, {
      text: content,
      paragraphs: [
        {
          text: content,
          runs: [
            {
              text: content,
              fontFamily: "Aptos",
              fontSizePt: 18,
              bold: false,
              italic: false,
              color: "#000000",
            },
          ],
        },
      ],
    });
    assert.ok(result.pptx.byteLength > 0);
  });

  it("emits run styles, fixed line spacing, and lists as native paragraph semantics", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "formatted-text",
            x: 24,
            y: 32,
            width: 420,
            height: 180,
            style: { lineHeight: "28px" },
          },
          createElement(
            P,
            {
              style: {
                align: "center",
                lineHeight: "24pt",
                list: { type: "bullet", character: "•" },
              },
            },
            "默认片段 ",
            createElement(
              Span,
              {
                style: {
                  fontFamily: font("Example Sans", [
                    { weight: 700, data: new Uint8Array([1, 2, 3]) },
                  ]),
                  fontSize: "12pt",
                  fontWeight: "bold",
                  color: "#334455",
                },
              },
              "中英 Text",
            ),
          ),
          createElement(
            P,
            { style: { list: { type: "number", startAt: 3 } } },
            "第二段",
          ),
        ),
      ),
    );
    assert.equal(
      result.report.slides[0]?.objects[0]?.text,
      "默认片段 中英 Text\n第二段",
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.text?.paragraphs, [
      {
        text: "默认片段 中英 Text",
        alignment: "ctr",
        lineSpacingPoints: 24,
        list: { type: "bullet", character: "•" },
        runs: [
          {
            text: "默认片段 ",
            fontFamily: "Aptos",
            fontSizePt: 18,
            bold: false,
            italic: false,
            color: "#000000",
          },
          {
            text: "中英 Text",
            fontFamily: "Example Sans",
            fontSizePt: 12,
            bold: true,
            italic: false,
            color: "#334455",
          },
        ],
      },
      {
        text: "第二段",
        lineSpacingPoints: 21,
        list: { type: "number", style: "arabicPeriod", startAt: 3 },
        runs: [
          {
            text: "第二段",
            fontFamily: "Aptos",
            fontSizePt: 18,
            bold: false,
            italic: false,
            color: "#000000",
          },
        ],
      },
    ]);
  });

  it("inherits text box styling for paragraphs and overrides only explicitly declared fields", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "paragraph-inherited-style",
            x: 24,
            y: 32,
            width: 420,
            height: 180,
            style: {
              fontFamily: "Paragraph Base",
              fontSize: "32pt",
              fontWeight: "bold",
              fontStyle: "italic",
              color: "#123456",
            },
          },
          createElement(P, {}, "继承全部文字样式"),
          createElement(P, { style: { color: "#ABCDEF" } }, "只覆盖颜色"),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects[0]?.text?.paragraphs.map(({ runs }) =>
        runs.map(({ fontFamily, fontSizePt, bold, italic, color }) => ({
          fontFamily,
          fontSizePt,
          bold,
          italic,
          color,
        })),
      ),
      [
        [
          {
            fontFamily: "Paragraph Base",
            fontSizePt: 32,
            bold: true,
            italic: true,
            color: "#123456",
          },
        ],
        [
          {
            fontFamily: "Paragraph Base",
            fontSizePt: 32,
            bold: true,
            italic: true,
            color: "#ABCDEF",
          },
        ],
      ],
    );
  });

  it("reports origins when conflicting resources are referenced for the same family, weight, and style", async () => {
    const element = deck(
      createElement(
        Text,
        { id: "font-conflict", x: 12, y: 12, width: 300, height: 80 },
        createElement(
          P,
          {},
          createElement(
            Span,
            {
              style: {
                fontFamily: font("Collision Sans", [
                  { data: new Uint8Array([1, 2]) },
                ]),
              },
            },
            "first",
          ),
          createElement(
            Span,
            {
              style: {
                fontFamily: font("Collision Sans", [
                  { data: new Uint8Array([3, 4]) },
                ]),
              },
            },
            "second",
          ),
        ),
      ),
    );

    await assert.rejects(build(element), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "conflicting-resource");
      assert.match(error.diagnostics[0]?.path ?? "", /text\[font-conflict\]/u);
      return true;
    });
  });

  it("fails when conflicting resources are referenced for actually used duplicate weight and style within same font value", async () => {
    const sharedValue = font("Duplicate Face", [
      { data: new Uint8Array([1, 2]) },
      { weight: 400, style: "normal", data: new Uint8Array([3, 4]) },
    ]);

    await assert.rejects(
      build(
        deck(
          createElement(
            Text,
            { id: "duplicate-face", x: 12, y: 12, width: 300, height: 80 },
            createElement(
              Span,
              { style: { fontFamily: sharedValue } },
              "实际使用默认字重和样式",
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "conflicting-resource");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /text\[duplicate-face\]/u,
        );
        return true;
      },
    );
  });

  it("preserves measurement hints and builds successfully when name-only references coexist with unique same-named font resource", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "font-name-with-resource",
            x: 12,
            y: 12,
            width: 300,
            height: 80,
          },
          createElement(
            P,
            {},
            createElement(
              Span,
              { style: { fontFamily: "Shared" } },
              "名称引用",
            ),
            createElement(
              Span,
              {
                style: {
                  fontFamily: font("Shared", [
                    { data: new Uint8Array([1, 2]) },
                    { weight: 700, data: new Uint8Array([3, 4]) },
                  ]),
                },
              },
              "资源引用",
            ),
          ),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs.map(
        ({ fontFamily }) => fontFamily,
      ),
      ["Shared", "Shared"],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ path, code }) => ({ path, code })),
      [
        {
          path: "deck/slide[cover]/text[font-name-with-resource]",
          code: "font-measurement-unavailable",
        },
      ],
    );
  });

  it("allows fixed bounding box to declare font family name only, reporting measurement limits without embedding fonts by default", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            id: "missing-metrics",
            x: 0,
            y: 0,
            width: 280,
            height: 60,
            style: { fontFamily: "Local Only" },
          },
          "不要求字体文件的固定文字",
        ),
      ),
    );

    assert.deepEqual(
      result.report.diagnostics.map(({ severity, path, code }) => ({
        severity,
        path,
        code,
      })),
      [
        {
          severity: "warning",
          path: "deck/slide[cover]/text[missing-metrics]",
          code: "font-measurement-unavailable",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontFamily,
      "Local Only",
    );
    const parts = unzipSync(result.pptx);
    assert.deepEqual(
      Object.keys(parts).filter((path) => path.startsWith("ppt/fonts/")),
      [],
    );
    assert.doesNotMatch(
      strFromU8(parts["ppt/presentation.xml"]!),
      /embeddedFont/u,
    );
  });

  it("preserves native line spacing for percentage-height fixed bounding boxes and measures per height convention", async () => {
    const jost = font("Jost*", [
      { weight: 400, data: testFont("./assets/fonts/Jost-400-Book.ttf") },
    ]);
    const buildRegion = (
      height: number | "100%" | "auto",
      fontFamily: string | typeof jost,
      lineHeight?: FontSize,
    ) =>
      build(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "fixed-text-region" },
            createElement(
              Box,
              { style: { width: 200, height: 80 } },
              createElement(
                Text,
                {
                  id: "fixed-text",
                  style: {
                    width: "100%",
                    height,
                    fontFamily,
                    fontSize: 20,
                    ...(lineHeight === undefined ? {} : { lineHeight }),
                  },
                },
                "Fixed region",
              ),
            ),
          ),
        ),
      );

    const namePercent = await buildRegion("100%", "Arial");
    const namePixels = await buildRegion(80, "Arial");
    assert.deepEqual(namePercent.pptx, namePixels.pptx);
    assert.deepEqual(
      namePercent.report.diagnostics.map(({ severity, code }) => ({
        severity,
        code,
      })),
      [{ severity: "warning", code: "font-measurement-unavailable" }],
    );
    for (const result of [namePercent, namePixels]) {
      assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
        x: 0,
        y: 0,
        width: 200,
        height: 80,
      });
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
      assert.doesNotMatch(xml, /<a:lnSpc>/u);
    }

    const metricsPercent = await buildRegion("100%", jost);
    const metricsPixels = await buildRegion(80, jost);
    assert.deepEqual(metricsPercent.pptx, metricsPixels.pptx);
    for (const result of [metricsPercent, metricsPixels]) {
      assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
        x: 0,
        y: 0,
        width: 200,
        height: 80,
      });
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const paragraph = file.slides[0]?.objects[0]?.text?.paragraphs[0];
      assert.equal(paragraph?.lineSpacingPoints, undefined);
      assert.equal(paragraph?.runs[0]?.fontFamily, "Jost*");
      const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
      assert.doesNotMatch(xml, /<a:lnSpc>/u);
    }

    const explicitLineHeight = await buildRegion("100%", jost, "22pt");
    const explicitParagraph = (await validatePptx(explicitLineHeight.pptx))
      .slides[0]?.objects[0]?.text?.paragraphs[0];
    assert.equal(explicitParagraph?.lineSpacingPoints, 22);
    assert.match(
      strFromU8(unzipSync(explicitLineHeight.pptx)["ppt/slides/slide1.xml"]!),
      /<a:lnSpc><a:spcPts val="2200"\/><\/a:lnSpc>/u,
    );

    const autoContent = await buildRegion("auto", jost);
    const autoParagraph = (await validatePptx(autoContent.pptx)).slides[0]
      ?.objects[0]?.text?.paragraphs[0];
    assert.equal(typeof autoParagraph?.lineSpacingPoints, "number");
    assert.match(
      strFromU8(unzipSync(autoContent.pptx)["ppt/slides/slide1.xml"]!),
      /<a:lnSpc>/u,
    );
  });

  it("generates plain shapes and shapes with text in standalone trimmed process without loading text box implementation", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx-text-trim-"));
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
      const child = spawnSync(
        process.execPath,
        [
          "--conditions=source",
          "--loader",
          loader,
          "--input-type=module",
          "--eval",
          [
            'import { createElement } from "react";',
            'import { BuildError } from "@react-pptx/core";',
            'import { Text as TextDeclaration } from "@react-pptx/capabilities/text-elements";',
            'import { build, Deck, Shape, Slide } from "@react-pptx/react-pptx/shape";',
            'import { validatePptx } from "@react-pptx/validation";',
            'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { id: "trimmed" }, createElement(Shape, { id: "plain", name: "plain", x: 0, y: 0, width: 20, height: 20 }), createElement(Shape, { id: "labeled", name: "labeled", x: 24, y: 0, width: 60, height: 20, text: "形状文字 — Shape text" }), createElement(Shape, { id: "empty-label", name: "empty-label", x: 90, y: 0, width: 20, height: 20, text: "" }))));',
            "const file = await validatePptx(result.pptx);",
            'if (file.errors.length !== 0 || result.report.slides[0]?.objects[1]?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[1]?.text?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[2]?.text !== undefined) throw new Error(JSON.stringify(file.errors));',
            'function NestedText() { return createElement(TextDeclaration, { id: "missing-text", x: 0, y: 0, width: 20, height: 20 }, "not available"); }',
            "let missingText;",
            'try { await build(createElement(Deck, {}, createElement(Slide, { id: "nested" }, createElement(NestedText)))); } catch (error) { missingText = error; }',
            'if (!(missingText instanceof BuildError) || missingText.diagnostics[0]?.code !== "missing-capability" || !missingText.diagnostics[0]?.path.includes("text[missing-text]")) throw new Error("nested Text did not produce a located missing-capability error");',
            'process.stdout.write(JSON.stringify({ kinds: file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]), missing: missingText.diagnostics[0], loaded: false }));',
          ].join("\n"),
        ],
        { cwd: process.cwd(), encoding: "utf8" },
      );
      assert.equal(child.status, 0, child.stderr || child.stdout);
      const output = JSON.parse(child.stdout);
      assert.deepEqual(output.kinds, [
        ["plain", ""],
        ["labeled", "形状文字 — Shape text"],
        ["empty-label", ""],
      ]);
      assert.equal(output.missing.severity, "error");
      assert.equal(
        output.missing.path,
        "deck/slide[nested]/text[missing-text]",
      );
      assert.equal(output.missing.code, "missing-capability");
      assert.ok(output.missing.message.includes("rp:text"));
      assert.equal(output.loaded, false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000); // Standalone process takes about 10s to load SDK WASM runtime on first run.

  it("generates Transitional PPTX from real React shapes and preserves read-only identity observation", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          id: "badge",
          name: "badge",
          geometry: "ellipse",
          x: 32,
          y: 48,
          width: 120,
          height: 72,
          fill: "#1A2B3C",
        }),
      ),
    );

    assert.ok(result.pptx instanceof Uint8Array);
    assert.ok(result.pptx.byteLength > 0);
    assert.equal(
      result.report.target.format,
      "ECMA-376 5th edition Transitional",
    );
    assert.equal(result.report.evidence, "generation");
    assert.deepEqual(result.report.slides[0]?.objects, [
      {
        id: 2,
        kind: "shape",
        name: "badge",
        authorId: "badge",
        geometry: "ellipse",
        box: { x: 32, y: 48, width: 120, height: 72 },
      },
    ]);
    assert.ok(Object.isFrozen(result.report));
    assert.ok(Object.isFrozen(result.report.slides));
    assert.ok(Object.isFrozen(result.report.slides[0]?.objects[0]));

    const file = await validatePptx(result.pptx);
    assert.equal(file.profile.standard, result.report.target.format);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 1);
    assert.equal(file.slides[0]?.id, "cover");
    assert.deepEqual(file.slides[0]?.objects[0], {
      id: 2,
      name: "badge",
      kind: "shape",
      geometry: "ellipse",
      boxEmu: { x: 304800, y: 457200, width: 1143000, height: 685800 },
    });
  });

  it("performs SDK structural validation on the same publicly built PPTX and preserves locations", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          id: "schema-gate",
          x: 0,
          y: 0,
          width: 120,
          height: 72,
        }),
      ),
    );
    const validFile = await validatePptx(result.pptx);
    assert.deepEqual(validFile.errors, []);
    assert.deepEqual(validFile.profile.schemaValidator, {
      package: "@ooxml-tools/validate",
      packageVersion: "0.4.6",
      sdkVersion: "3.0.2",
      officeVersion: "Microsoft365",
    });
    assert.equal(validFile.profile.applicationEvidence, "not-validated");
    assert.match(validFile.profile.scope, /MC/iu);

    const masterPart = "ppt/slideMasters/slideMaster1.xml";
    const masterParts = unzipSync(result.pptx);
    const master = strFromU8(masterParts[masterPart]!);
    const presentation = strFromU8(masterParts["ppt/presentation.xml"]!);
    const masterIds = Array.from(
      presentation.matchAll(/<p:sldMasterId\b[^>]*\bid="(\d+)"/gu),
      (match) => Number(match[1]),
    );
    const layoutIds = Array.from(
      master.matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/gu),
      (match) => Number(match[1]),
    );
    assert.equal(masterIds.length, 1);
    assert.equal(layoutIds.length, 1);
    for (const id of [...masterIds, ...layoutIds]) {
      assert.ok(id >= 2_147_483_648 && id <= 4_294_967_295);
    }
    assert.equal(
      new Set([...masterIds, ...layoutIds]).size,
      masterIds.length + layoutIds.length,
    );

    const layoutIdMatch = master.match(
      /(<p:sldLayoutId\b[^>]*\bid=")(\d+)(")/u,
    );
    assert.ok(layoutIdMatch);
    masterParts[masterPart] = strToU8(
      master.replace(
        layoutIdMatch[0],
        `${layoutIdMatch[1]}1${layoutIdMatch[3]}`,
      ),
    );
    const invalidMaster = await validatePptx(zipSync(masterParts));
    assert.ok(
      invalidMaster.errors.some(
        ({ code, part, xpath }) =>
          code === "Sch_AttributeValueDataTypeDetailed" &&
          part === "/ppt/slideMasters/slideMaster1.xml" &&
          xpath !== undefined,
      ),
    );

    const textResult = await build(
      deck(
        createElement(
          Text,
          {
            x: 0,
            y: 0,
            width: 240,
            height: 96,
            style: { lineHeight: "24px" },
          },
          "标准结构边界",
        ),
      ),
    );
    const slidePart = "ppt/slides/slide1.xml";
    const textParts = unzipSync(textResult.pptx);
    const textSlide = strFromU8(textParts[slidePart]!);
    assert.match(textSlide, /sz="1800"/u);
    assert.match(textSlide, /<a:spcPts val="1800"/u);
    textParts[slidePart] = strToU8(
      textSlide
        .replace('sz="1800"', 'sz="99"')
        .replace('<a:spcPts val="1800"', '<a:spcPts val="158401"'),
    );
    const invalidText = await validatePptx(zipSync(textParts));
    assert.ok(
      invalidText.errors.some(
        ({ code, part, xpath }) =>
          code === "Sch_AttributeValueDataTypeDetailed" &&
          part === "/ppt/slides/slide1.xml" &&
          xpath?.includes("/a:rPr") === true,
      ),
    );
    assert.ok(
      invalidText.errors.some(
        ({ code, part, xpath }) =>
          code === "Sch_AttributeValueDataTypeDetailed" &&
          part === "/ppt/slides/slide1.xml" &&
          xpath?.includes("/a:spcPts") === true,
      ),
    );

    const invalidOrderParts = unzipSync(result.pptx);
    const slide = strFromU8(invalidOrderParts[slidePart]!);
    assert.match(slide, /<\/p:cSld><p:clrMapOvr>/u);
    invalidOrderParts[slidePart] = strToU8(
      slide.replace(
        "</p:cSld><p:clrMapOvr>",
        "</p:cSld><p:transition><p:fade/></p:transition><p:clrMapOvr>",
      ),
    );
    const invalidOrder = await validatePptx(zipSync(invalidOrderParts));
    assert.ok(
      invalidOrder.errors.some(
        ({ code, part, xpath }) =>
          code === "Sch_UnexpectedElementContentExpectingComplex" &&
          part === "/ppt/slides/slide1.xml" &&
          xpath === "/p:sld[1]",
      ),
    );
  });

  it("preserves newlines, carriage returns, and tabs in slide identifiers when read back from PPTX", async () => {
    const slideIds = [
      "alpha\nbeta",
      "alpha\rbeta",
      "alpha\tbeta",
      "alpha beta",
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: slideIds[0] },
          createElement(Shape, {
            id: "object-name",
            name: "object\nname\r\tend",
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(Shape, {
            id: "test\nid",
            name: "test\nid",
            x: 50,
            y: 0,
            width: 40,
            height: 30,
          }),
        ),
        ...slideIds.slice(1).map((id) => createElement(Slide, { id })),
      ),
    );

    const file = await validatePptx(result.pptx);

    assert.deepEqual(
      result.report.slides.map(({ id }) => id),
      slideIds,
    );
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ id }) => id),
      slideIds,
    );
    assert.deepEqual(
      file.slides[0]?.objects.map(({ name }) => name),
      ["object\nname\r\tend", "test\nid"],
    );
  });

  it("rejects object names and slide identifiers that cannot be represented in XML 1.0 and locates source", async () => {
    const invalidObjectName = deck(
      createElement(Shape, {
        id: "invalid-name",
        name: "hello\u0000world",
        x: 0,
        y: 0,
        width: 40,
        height: 30,
      }),
    );
    const invalidText = deck(
      createElement(
        Text,
        {
          id: "invalid-text",
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        },
        "hello\u0000world",
      ),
    );
    const invalidSlideId = createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(Slide, { id: "alpha\u0000beta" }),
    );

    for (const [element, path] of [
      [invalidObjectName, /slide\[cover\].*shape\[invalid-name\]/u],
      [invalidText, /slide\[cover\].*text\[invalid-text\]/u],
      [invalidSlideId, /slide\[1\]/u],
    ] as const) {
      await assert.rejects(build(element), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-xml-character");
        assert.match(error.diagnostics[0]?.path ?? "", path);
        return true;
      });
    }
  });

  it("rejects shape dimensions that cannot be expressed within closed EMU interval", async () => {
    for (const size of [
      { width: 0.00001, height: 20 },
      { width: 40, height: 0.00001 },
      { width: 1_000_000, height: 20 },
      { width: 40, height: 1_000_000 },
    ]) {
      await assert.rejects(
        build(
          deck(
            createElement(Shape, {
              id: "too-small",
              x: 0,
              y: 0,
              ...size,
            }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-coordinate");
          assert.match(
            error.diagnostics[0]?.path ?? "",
            /slide\[cover\].*shape\[too-small\]/u,
          );
          return true;
        },
      );
    }
  });

  it("accepts final closed EMU interval boundaries for shape dimensions in standalone reader", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 5376, height: 5376 },
        createElement(
          Slide,
          { id: "extent-boundary" },
          createElement(Shape, {
            id: "extent-boundary",
            x: 0,
            y: 0,
            // This pixel value rounds to the shape size upper bound after standard conversion.
            width: 225457.60073490813,
            height: 0.00006,
          }),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.width, 2_147_483_647);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.height, 1);
  });

  it("rejects shape dimensions exceeding format upper limit in standalone validator", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          id: "oversized",
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const slidePart = "ppt/slides/slide1.xml";
    const slide = parts[slidePart];
    assert.ok(slide);
    const slideXml = strFromU8(slide);
    const invalidSlideXml = slideXml.replace(
      '<a:ext cx="381000" cy="285750"',
      '<a:ext cx="2147483648" cy="285750"',
    );
    assert.notEqual(invalidSlideXml, slideXml);
    parts[slidePart] = strToU8(invalidSlideXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === slidePart && code === "invalid-shape-transform",
      ),
    );
  });

  it("constrains slide dimensions to final EMU limits and preserves fractional pixels within boundaries", async () => {
    const boundary = await build(
      createElement(
        Deck,
        { width: 95.99995, height: 5376.00005 },
        createElement(Slide, { id: "boundary" }),
      ),
    );
    const boundaryFile = await validatePptx(boundary.pptx);
    assert.deepEqual(boundaryFile.errors, []);
    assert.equal(boundaryFile.presentation.widthEmu, 914400);
    assert.equal(boundaryFile.presentation.heightEmu, 51206400);

    for (const size of [
      { width: 1, height: 1 },
      { width: 95.9999, height: 100 },
      { width: 100, height: 95.9999 },
      { width: 5376.0001, height: 100 },
      { width: 100, height: 5376.0001 },
      { width: 10000, height: 720 },
    ]) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            size,
            createElement(Slide, { id: "invalid-size" }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-dimension");
          assert.equal(error.diagnostics[0]?.path, "deck");
          return true;
        },
      );
    }
  });

  it("rejects actual slide dimensions exceeding format range in standalone validator", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 96, height: 5376 },
        createElement(Slide, { id: "size" }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const presentationPart = "ppt/presentation.xml";
    const presentation = parts[presentationPart];
    assert.ok(presentation);
    const presentationXml = strFromU8(presentation);
    const invalidPresentationXml = presentationXml
      .replace('cx="914400"', 'cx="914399"')
      .replace('cy="51206400"', 'cy="51206401"');
    assert.notEqual(invalidPresentationXml, presentationXml);
    parts[presentationPart] = strToU8(invalidPresentationXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === presentationPart && code === "invalid-slide-size",
      ),
    );
  });

  it("rejects master layout identifier below format lower limit in standalone validator", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          id: "badge",
          name: "badge",
          geometry: "ellipse",
          x: 32,
          y: 48,
          width: 120,
          height: 72,
        }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const masterPart = "ppt/slideMasters/slideMaster1.xml";
    const master = parts[masterPart];
    assert.ok(master);
    const masterXml = strFromU8(master);
    const invalidMasterXml = masterXml.replace(
      /(<p:sldLayoutId id=")\d+(")/u,
      (_match, prefix: string, suffix: string) =>
        `${prefix}2147483647${suffix}`,
    );
    assert.notEqual(invalidMasterXml, masterXml);
    parts[masterPart] = strToU8(invalidMasterXml);

    const file = await validatePptx(zipSync(parts));

    assert.ok(
      file.errors.some(
        ({ part, code }) =>
          part === masterPart && code === "invalid-slide-layout-id",
      ),
    );
  });

  it("awaits layout and passive Effect commits of real components and completes root unmount cleanup", async () => {
    const events: string[] = [];
    function EffectShape() {
      useLayoutEffect(() => {
        events.push("layout-commit");
        return () => {
          events.push("layout-cleanup");
        };
      }, []);
      useEffect(() => {
        events.push("passive-commit");
        return () => {
          events.push("passive-cleanup");
        };
      }, []);
      return createElement(Shape, {
        id: "from-component",
        x: 8,
        y: 9,
        width: 40,
        height: 30,
      });
    }

    const result = await build(deck(createElement(EffectShape)));

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId }) => testId),
      ["from-component"],
    );
    assert.deepEqual(events.sort(), [
      "layout-cleanup",
      "layout-commit",
      "passive-cleanup",
      "passive-commit",
    ]);
  });

  it("runs the shared formal build only after React unmount cleanup", async () => {
    let cleaned = 0;
    const cleanupSeenAtParse: number[] = [];
    const cleanupSeenAtCompile: number[] = [];
    const observedShape: BoundObjectCapabilityModule = {
      declaration: shapeCapability.declaration,
      bind(input, context) {
        cleanupSeenAtParse.push(cleaned);
        const bound = shapeCapability.bind(input, context);
        return {
          ...bound,
          compile(compileContext) {
            cleanupSeenAtCompile.push(cleaned);
            return bound.compile(compileContext);
          },
        };
      },
    };

    function CleanupBeforeBuild() {
      useLayoutEffect(
        () => () => {
          cleaned += 1;
        },
        [],
      );
      useEffect(
        () => () => {
          cleaned += 1;
        },
        [],
      );
      return createElement(Shape, {
        id: "cleanup-before-build",
        x: 0,
        y: 0,
        width: 240,
        height: 48,
        text: "Cleanup before build",
      });
    }

    const result = await createBuilder([observedShape])(
      deck(createElement(CleanupBeforeBuild)),
    );

    assert.equal(cleaned, 2);
    assert.deepEqual(cleanupSeenAtParse, [2]);
    assert.deepEqual(cleanupSeenAtCompile, [2]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.text?.text,
      "Cleanup before build",
    );
  });

  it("keeps original build options visible to capability parse and compile", async () => {
    const controller = new AbortController();
    const parseOptions: BuildOptions[] = [];
    const compileOptions: BuildOptions[] = [];
    const observedShape: BoundObjectCapabilityModule = {
      declaration: shapeCapability.declaration,
      bind(input, context) {
        parseOptions.push(context.options);
        const bound = shapeCapability.bind(input, context);
        return {
          ...bound,
          compile(compileContext) {
            compileOptions.push(compileContext.options);
            return bound.compile(compileContext);
          },
        };
      },
    };

    const result = await createBuilder([observedShape])(
      deck(
        createElement(Shape, {
          id: "options-visible-to-capabilities",
          geometry: "triangle",
          x: 0,
          y: 0,
          width: 80,
          height: 60,
        }),
      ),
      {
        signal: controller.signal,
        target: targets.baselineOnly,
        embedFonts: "full",
      },
    );

    assert.equal(parseOptions.length, 1);
    assert.equal(compileOptions.length, 1);
    for (const options of [parseOptions[0], compileOptions[0]]) {
      assert.equal(options?.signal, controller.signal);
      assert.equal(options?.target, targets.baselineOnly);
      assert.equal(options?.embedFonts, "full");
    }
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });

  it("reports unknown shape parameters and trimmed capabilities at declaration path", async () => {
    const UncheckedShape = Shape as ComponentType<Record<string, unknown>>;
    const unknownParameter = deck(
      createElement(UncheckedShape, {
        id: "unknown-prop",
        x: 0,
        y: 0,
        width: 20,
        height: 20,
        mystery: true,
      }),
    );

    await assert.rejects(build(unknownParameter), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
      assert.match(
        error.diagnostics[0]?.path ?? "",
        /slide\[cover\].*shape\[unknown-prop\]/u,
      );
      return true;
    });

    const clippedBuild = createBuilder([]);
    await assert.rejects(
      clippedBuild(
        deck(
          createElement(Shape, {
            id: "trimmed-shape",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /slide\[cover\].*shape\[trimmed-shape\]/u,
        );
        return true;
      },
    );
  });

  it("adheres to basic Transitional target along the same pipeline for explicit shape module", async () => {
    const result = await createBuilder([shapeCapability])(
      deck(
        createElement(Shape, {
          id: "triangle",
          geometry: "triangle",
          x: 0,
          y: 0,
          width: 80,
          height: 60,
          fill: "#FFAA00",
        }),
      ),
      { target: targets.baselineOnly },
    );

    assert.deepEqual(result.report.target.extensions, []);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.geometry, "triangle");
  });

  it("generates and reads back basic slide transitions from the same PPTX", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(Slide, {
          id: "fade",
          transition: transitions.fade(),
        }),
        createElement(Slide, {
          id: "strips",
          transition: transitions.strips({ direction: "leftUp" }),
        }),
        createElement(Slide, {
          id: "push",
          transition: transitions.push({ from: "top" }),
        }),
        createElement(Slide, {
          id: "strips-default",
          transition: transitions.strips(),
        }),
        createElement(Slide, {
          id: "push-default",
          transition: transitions.push(),
        }),
      ),
    );

    assert.deepEqual(result.report.extensions, []);
    assert.deepEqual(
      result.report.slides.map(
        ({ capabilities }) => capabilities?.["transition"],
      ),
      [
        { kind: "fade", speed: "slow" },
        { kind: "strips", direction: "leftUp", speed: "slow" },
        { kind: "push", from: "top", speed: "slow" },
        { kind: "strips", direction: "rightDown", speed: "slow" },
        { kind: "push", from: "bottom", speed: "slow" },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ transition }) => transition),
      [
        { kind: "fade", speed: "slow" },
        { kind: "strips", speed: "slow", direction: "lu" },
        { kind: "push", speed: "slow", direction: "u" },
        { kind: "strips", speed: "slow", direction: "rd" },
        { kind: "push", speed: "slow", direction: "d" },
      ],
    );
  });

  it("emits each direction parameter as corresponding Transitional direction code", async () => {
    const declarations = [
      ...(
        [
          ["leftUp", "lu"],
          ["rightUp", "ru"],
          ["leftDown", "ld"],
          ["rightDown", "rd"],
        ] as const
      ).map(([direction], index) =>
        createElement(Slide, {
          id: `strips-${index}`,
          transition: transitions.strips({ direction }),
        }),
      ),
      ...(
        [
          ["left", "l"],
          ["right", "r"],
          ["top", "u"],
          ["bottom", "d"],
        ] as const
      ).map(([from], index) =>
        createElement(Slide, {
          id: `push-${index}`,
          transition: transitions.push({ from }),
        }),
      ),
    ];
    const result = await build(createElement(Deck, {}, ...declarations));
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ transition }) => [
        transition?.kind,
        transition?.direction,
      ]),
      [
        ["strips", "lu"],
        ["strips", "ru"],
        ["strips", "ld"],
        ["strips", "rd"],
        ["push", "l"],
        ["push", "r"],
        ["push", "u"],
        ["push", "d"],
      ],
    );
  });

  it("selects basic speed based on duration bounds and emits rounded exact milliseconds", async () => {
    const timings = [
      [600, "fast", 600],
      [600.6, "med", 601],
      [601, "med", 601],
      [875, "med", 875],
      [876, "slow", 876],
    ] as const;
    const result = await build(
      createElement(
        Deck,
        {},
        ...timings.map(([duration], index) =>
          createElement(Slide, {
            id: `duration-${index}`,
            transition: transitions.fade({ duration }),
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ transition }) => [
        transition?.speed,
        transition?.durationMs,
      ]),
      timings.map(([, speed, duration]) => [speed, duration]),
    );
  });

  it("detects Choice-omitted p14 duration dependency via standalone file check", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "morph-duration",
          transition: transitions.morph({
            duration: 1200,
            fallback: transitions.fade(),
          }),
        }),
      ),
    );
    const parts = unzipSync(result.pptx);
    const slidePart = parts["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const xml = strFromU8(slidePart);
    assert.match(xml, /Requires="p14 p159"/u);
    parts["ppt/slides/slide1.xml"] = strToU8(
      xml.replace('Requires="p14 p159"', 'Requires="p159"'),
    );

    const file = await validatePptx(
      zipSync(parts, { mtime: new Date("2020-01-01T00:00:00Z") }),
      { allowedExtensions: ["p159"] },
    );
    assert.ok(
      file.errors.some(({ code }) => code === "missing-extension-requirement"),
    );
    assert.ok(file.errors.some(({ code }) => code === "forbidden-extension"));
  });

  it("writes exact duration into p14 with justified Transitional speed fallback", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "timed",
          transition: transitions.fade({ duration: 777 }),
        }),
      ),
    );

    assert.deepEqual(result.report.extensions, ["p14"]);
    assert.deepEqual(result.report.slides[0]?.capabilities?.["transition"], {
      kind: "fade",
      speed: "med",
      durationMs: 777,
      fallback: { kind: "fade", speed: "med" },
    });
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, ["p14"]);
    assert.deepEqual(file.slides[0]?.transition, {
      kind: "fade",
      speed: "med",
      durationMs: 777,
      fallback: { kind: "fade", speed: "med" },
    });

    const disallowed = await validatePptx(result.pptx, {
      allowedExtensions: [],
    });
    assert.ok(
      disallowed.errors.some(({ code }) => code === "forbidden-extension"),
    );
    await assert.rejects(
      build(
        createElement(
          Deck,
          {},
          createElement(Slide, {
            id: "baseline-only",
            transition: transitions.fade({ duration: 777 }),
          }),
        ),
        { target: targets.baselineOnly },
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        assert.equal(error.diagnostics[0]?.path, "deck/slide[baseline-only]");
        return true;
      },
    );
  });

  it("includes explicit compatibility branches for Flip and Morph and verifies combination mechanism requirements", async () => {
    const declarations = createElement(
      Deck,
      {},
      createElement(Slide, {
        id: "flip",
        transition: transitions.flip({
          direction: "left",
          fallback: transitions.fade(),
        }),
      }),
      createElement(Slide, {
        id: "morph",
        transition: transitions.morph({
          by: "character",
          duration: 1200,
          fallback: transitions.push({ from: "right" }),
        }),
      }),
    );
    const result = await build(declarations);

    assert.deepEqual(result.report.extensions, ["p14", "p159"]);
    assert.deepEqual(
      result.report.slides.map(
        ({ capabilities }) => capabilities?.["transition"],
      ),
      [
        {
          kind: "flip",
          direction: "left",
          speed: "slow",
          fallback: { kind: "fade", speed: "slow" },
        },
        {
          kind: "morph",
          by: "character",
          speed: "slow",
          durationMs: 1200,
          fallback: { kind: "push", from: "right", speed: "slow" },
        },
      ],
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, ["p14", "p159"]);
    assert.deepEqual(
      file.slides.map(({ transition }) => transition),
      [
        {
          kind: "flip",
          speed: "slow",
          direction: "l",
          fallback: { kind: "fade", speed: "slow" },
        },
        {
          kind: "morph",
          speed: "slow",
          durationMs: 1200,
          by: "character",
          fallback: { kind: "push", speed: "slow", direction: "r" },
        },
      ],
    );

    await assert.rejects(
      build(declarations, { target: targets.baselineOnly }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        return true;
      },
    );
    await assert.rejects(
      build(declarations, {
        target: {
          ...targets.transitional,
          name: "p159-only",
          extensions: ["p159"],
        },
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        assert.ok(error.diagnostics[0]?.message.includes("p14"));
        return true;
      },
    );
    await assert.rejects(
      build(declarations, {
        target: {
          ...targets.transitional,
          name: "p14-only",
          extensions: ["p14"],
        },
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        assert.ok(error.diagnostics[0]?.message.includes("p159"));
        return true;
      },
    );
  });

  it("uses extended transition default options and preserves basic fallback declaration", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "flip-default",
          transition: transitions.flip({ fallback: transitions.strips() }),
        }),
        createElement(Slide, {
          id: "morph-default",
          transition: transitions.morph({ fallback: transitions.fade() }),
        }),
      ),
    );
    const file = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map(({ transition }) => transition),
      [
        {
          kind: "flip",
          speed: "slow",
          direction: "r",
          fallback: { kind: "strips", speed: "slow", direction: "rd" },
        },
        {
          kind: "morph",
          speed: "slow",
          by: "object",
          fallback: { kind: "fade", speed: "slow" },
        },
      ],
    );
  });

  it("locates omitted Slide declarations after trimming transition capabilities", async () => {
    await assert.rejects(
      createBuilder([shapeCapability])(
        createElement(
          Deck,
          {},
          createElement(Slide, {
            id: "missing-transition-module",
            transition: transitions.fade(),
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[missing-transition-module]",
        );
        return true;
      },
    );
  });

  it("rejects invalid duration, direction, options, and fallback at Slide path", async () => {
    const invalidTransitions: readonly {
      readonly id: string;
      readonly transition: unknown;
      readonly code: string;
    }[] = [
      {
        id: "negative-duration",
        transition: transitions.fade({ duration: -1 }),
        code: "invalid-transition-duration",
      },
      {
        id: "non-finite-duration",
        transition: transitions.fade({ duration: Number.POSITIVE_INFINITY }),
        code: "invalid-transition-duration",
      },
      {
        id: "unknown-parameter",
        transition: transitions.fade({
          duration: 600,
          surprise: true,
        } as never),
        code: "unknown-parameter",
      },
      {
        id: "null-strips-direction",
        transition: transitions.strips({ direction: null as never }),
        code: "invalid-transition-direction",
      },
      {
        id: "invalid-strips-direction",
        transition: transitions.strips({ direction: "diagonal" as never }),
        code: "invalid-transition-direction",
      },
      {
        id: "null-push-direction",
        transition: transitions.push({ from: null as never }),
        code: "invalid-transition-direction",
      },
      {
        id: "invalid-push-direction",
        transition: transitions.push({ from: "center" as never }),
        code: "invalid-transition-direction",
      },
      {
        id: "null-flip-direction",
        transition: transitions.flip({
          direction: null as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-transition-direction",
      },
      {
        id: "missing-flip-fallback",
        transition: transitions.flip({} as never),
        code: "missing-transition-fallback",
      },
      {
        id: "null-morph-option",
        transition: transitions.morph({
          by: null as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-morph-option",
      },
      {
        id: "invalid-morph-option",
        transition: transitions.morph({
          by: "sentence" as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-morph-option",
      },
      {
        id: "helper-kind-and-unknown-option",
        transition: transitions.fade({ kind: "push", from: "left" } as never),
        code: "unknown-parameter",
      },
      {
        id: "extension-fallback",
        transition: transitions.flip({
          fallback: transitions.morph({
            fallback: transitions.fade(),
          }) as never,
        }),
        code: "invalid-transition-fallback",
      },
    ];

    for (const invalid of invalidTransitions) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(Slide, {
              id: invalid.id,
              transition: invalid.transition as never,
            }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, invalid.code);
          assert.equal(error.diagnostics[0]?.path, `deck/slide[${invalid.id}]`);
          return true;
        },
      );
    }
  });

  it("preserves original shapes on out-of-bounds and returns warnings in generated evidence", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 100, height: 100 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(Shape, {
            id: "overflowing",
            x: 90,
            y: 70,
            width: 30,
            height: 20,
          }),
        ),
      ),
    );

    assert.equal(result.report.diagnostics.length, 1);
    assert.equal(result.report.diagnostics[0]?.severity, "warning");
    assert.equal(
      result.report.diagnostics[0]?.path,
      "deck/slide[cover]/shape[overflowing]",
    );
    assert.equal(result.report.diagnostics[0]?.code, "overflow");
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.boxEmu, {
      x: 857250,
      y: 666750,
      width: 285750,
      height: 190500,
    });
  });

  it("explicitly rejects undeclared shapes and unwriteable dimensions", async () => {
    await assert.rejects(
      build(
        deck(
          createElement(Shape, {
            id: "unknown-geometry",
            geometry: "madeUpShape",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unsupported-geometry");
        return true;
      },
    );

    await assert.rejects(
      build(
        deck(
          createElement(Shape, {
            id: "zero-size",
            x: 0,
            y: 0,
            width: 0,
            height: 20,
          }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-coordinate");
        return true;
      },
    );
  });
});
