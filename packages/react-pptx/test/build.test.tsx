import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { ObjectCapabilityModule } from "@react-pptx/core";
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
    createElement(Slide, { slideId: "cover" }, shape),
  );
}

function testFont(path: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(path, import.meta.url)));
}

function parsedFont(bytes: Uint8Array): fontkit.Font {
  const parsed = fontkit.create(Buffer.from(bytes));
  if (parsed.type !== "TTF") throw new Error("预期静态 TrueType 字体");
  return parsed as fontkit.Font;
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

describe("公开构建入口", () => {
  it("按 Text、Table、Chart、大小写和列表的实际字符嵌入 TrueType 子集", async () => {
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
          { slideId: "font-subset" },
          createElement(
            Text,
            {
              testId: "subset-text",
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
              testId: "subset-latin",
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
              testId: "subset-table",
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
            testId: "subset-chart",
            kind: "column",
            x: 280,
            y: 128,
            width: 320,
            height: 220,
            style: { fontFamily: cjkFont },
            data: {
              categories: ["销售图"],
              series: [{ name: "金额", values: [12] }],
            },
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
      // cmap format 4 的结束段必须保留 U+FFFF 哨兵。
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

  it("图表独占字体子集覆盖刻度、格式字面量、百分比和样式继承", async () => {
    const chartFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: zcoolXiaoWei },
    ]);
    const chartCases = [
      {
        testId: "subset-chart-default-axis",
        chart: createElement(Chart, {
          testId: "subset-chart-default-axis",
          kind: "column",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789"),
        absent: [],
      },
      {
        testId: "subset-chart-percent-axis",
        chart: createElement(Chart, {
          testId: "subset-chart-percent-axis",
          kind: "column",
          stacking: "percent",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789%"),
        absent: [],
      },
      {
        testId: "subset-chart-formatted-axis",
        chart: createElement(Chart, {
          testId: "subset-chart-formatted-axis",
          kind: "column",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          valueAxis: {
            min: 0,
            max: 20,
            majorUnit: 5,
            numberFormat: '#,##0"元"',
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789,元"),
        absent: ["#", '"'],
      },
      {
        testId: "subset-chart-formatted-label",
        chart: createElement(Chart, {
          testId: "subset-chart-formatted-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: '0"元"',
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲", "乙"],
            series: [{ name: "丙", values: [1, 2] }],
          },
        }),
        expected: Array.from("0123456789元"),
        absent: ['"'],
      },
      {
        testId: "subset-chart-percent-label",
        chart: createElement(Chart, {
          testId: "subset-chart-percent-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: { percent: true, style: { fontSize: 14 } },
          data: {
            categories: ["甲", "乙"],
            series: [{ name: "丙", values: [1, 2] }],
          },
        }),
        expected: Array.from("0123456789%"),
        absent: [],
      },
      {
        testId: "subset-chart-locale-only-label",
        chart: createElement(Chart, {
          testId: "subset-chart-locale-only-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: "[$-409]0",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789"),
        absent: ["$", "[", "]"],
      },
      {
        testId: "subset-chart-currency-literal-label",
        chart: createElement(Chart, {
          testId: "subset-chart-currency-literal-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: "[$￥-804]#,##0.00",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789￥,."),
        absent: ["$", "#", "[", "]"],
      },
      {
        testId: "subset-chart-condition-color-label",
        chart: createElement(Chart, {
          testId: "subset-chart-condition-color-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: '[Red][>=100]0"元";[Blue][<100]0"分"',
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789元分"),
        absent: Array.from("[]<>RedBlu"),
      },
      {
        testId: "subset-chart-escaped-literal-label",
        chart: createElement(Chart, {
          testId: "subset-chart-escaped-literal-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: "0\\元",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789元"),
        absent: ["\\"],
      },
      {
        testId: "subset-chart-chinese-date-name-label",
        chart: createElement(Chart, {
          testId: "subset-chart-chinese-date-name-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: "[$-804]dddd",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789星期日一二三四五六"),
        absent: ["A", "J", "S"],
      },
      {
        testId: "subset-chart-question-placeholder-label",
        chart: createElement(Chart, {
          testId: "subset-chart-question-placeholder-label",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          dataLabels: {
            value: true,
            numberFormat: "??",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
        }),
        expected: Array.from("0123456789 "),
        absent: ["?"],
      },
    ] as const;

    for (const { testId, chart, expected, absent } of chartCases) {
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
          `${testId} 子集应覆盖 ${character}`,
        );
      }
      for (const character of absent) {
        assert.equal(
          subset.hasGlyphForCodePoint(character.codePointAt(0)!),
          false,
          `${testId} 不应把格式控制符 ${character} 当作显示字形`,
        );
      }
    }

    const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    await assert.rejects(
      build(
        deck(
          createElement(Chart, {
            testId: "subset-chart-label-style-missing-glyph",
            kind: "pie",
            x: 0,
            y: 0,
            width: 320,
            height: 220,
            style: { fontFamily: chartFont },
            dataLabels: {
              value: true,
              numberFormat: '0"元"',
              style: { fontFamily: latinFont },
            },
            data: {
              categories: ["甲", "乙"],
              series: [{ name: "丙", values: [1, 2] }],
            },
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

  it("数值标签忽略专用于文本的第四节", async () => {
    const latinFont = font("Jost*", [{ weight: 400, data: jostBook }]);
    const result = await build(
      deck(
        createElement(Chart, {
          testId: "subset-chart-fourth-text-section",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: latinFont },
          dataLabels: {
            value: true,
            numberFormat: '0;0;0;"元"@',
            style: { fontSize: 14 },
          },
          data: {
            categories: ["A"],
            series: [{ name: "B", values: [12] }],
          },
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

  it("locale-only 数值格式不要求货币字形", async () => {
    const chartFont = font("ZCOOL XiaoWei", [
      { weight: 400, data: zcoolXiaoWei },
    ]);
    const axisResult = await build(
      deck(
        createElement(Chart, {
          testId: "subset-chart-locale-only-source",
          kind: "column",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: chartFont },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
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
          testId: "subset-chart-locale-only-label-narrow-font",
          kind: "pie",
          x: 0,
          y: 0,
          width: 320,
          height: 220,
          style: { fontFamily: narrowFont },
          dataLabels: {
            value: true,
            numberFormat: "[$-409]0",
            style: { fontSize: 14 },
          },
          data: {
            categories: ["甲"],
            series: [{ name: "乙", values: [12] }],
          },
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

  it("字符子集不把换行和制表控制符当作缺失字形", async () => {
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
              testId,
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

  it("Shape 正文经共享文字服务度量、写出页码字段并登记实际字符", async () => {
    const pageNumberClaims: string[] = [];
    const shapeModule = shapeCapability as ObjectCapabilityModule;
    const observedShapeCapability: ObjectCapabilityModule = {
      ...shapeModule,
      compile(input, context) {
        const output = shapeModule.compile(input, context);
        pageNumberClaims.push(
          ...(output.resourceClaims ?? []).flatMap(({ characters }) =>
            characters === undefined ? [] : [characters],
          ),
        );
        return output;
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
            testId: `field-shape-${pageNumber}`,
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
          testId: `literal-shape-${pageNumber}`,
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
            { key: pageNumber, slideId: `shape-field-${pageNumber}` },
            pageNumber === 1 || pageNumber === 10
              ? pageShapes(pageNumber)
              : null,
          );
        }),
      ),
    );

    const fieldShape = (pageNumber: number) =>
      result.report.slides[pageNumber - 1]?.objects.find(
        ({ testId }) => testId === `field-shape-${pageNumber}`,
      );
    const literalShape = (pageNumber: number) =>
      result.report.slides[pageNumber - 1]?.objects.find(
        ({ testId }) => testId === `literal-shape-${pageNumber}`,
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
              testId: "conflicting-shape-body",
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

  it("Text、Shape、Table 页码字段进入同一 Chart 字符子集", async () => {
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
          { slideId: "field-first" },
          createElement(
            Text,
            {
              testId: "text-first",
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
              testId: "shape-first",
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
              testId: "table-first",
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
          { slideId: "field-second" },
          createElement(
            Text,
            {
              testId: "text-second",
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
              testId: "shape-second",
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
              testId: "table-second",
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
            testId: "page-field-chart",
            kind: "pie",
            x: 110,
            y: 0,
            width: 200,
            height: 220,
            style: { fontFamily: pageFont },
            dataLabels: {
              value: true,
              numberFormat: '0;0;0;"元"@',
            },
            data: {
              categories: ["A"],
              series: [{ name: "B", values: [12] }],
            },
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

  it("内容定高文字缺少度量字体时以构建错误拒绝", async () => {
    let failure: unknown;
    try {
      await build(
        deck(
          createElement(Text, {
            testId: "missing-content-metrics",
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

  it("按明确 WOFF2 字体指标计算内容高度并写入相同固定行距", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          testId: "measured-mixed-text",
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

  it("完整嵌入覆盖实际使用的文本框、形状、表格和图表字族样式", async () => {
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
          { slideId: "font-embedding" },
          createElement(
            Text,
            {
              testId: "embedded-font-bold-first",
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
              testId: "embedded-font-text",
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
            testId: "embedded-font-shape",
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
              testId: "embedded-font-table",
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
            testId: "embedded-font-chart",
            kind: "line",
            x: 300,
            y: 18,
            width: 400,
            height: 240,
            style: { fontFamily: installedFont },
            valueAxis: { style: { fontWeight: "bold" } },
            data: {
              categories: ["Q1", "Q2"],
              series: [{ name: "Sales", values: [12, 18] }],
            },
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

  it("按 OS/2 样式链接把 600 字重 SemiBold 写入 regular 槽位", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "embedded-semibold",
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

  it("完整嵌入把真实斜体和粗斜体字体映射到对应样式槽", async () => {
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
          { slideId: "italic-font-slots" },
          createElement(
            Text,
            {
              testId: "italic-font",
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
              testId: "bold-italic-font",
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

  it("同一文字中的嵌入限制不受相同字节去重和片段顺序影响", async () => {
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
                  testId,
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

  it("允许的相同字体字节仍合并，默认构建仍不嵌入字体", async () => {
    const firstAllowed = font("Jost*", [{ weight: 400, data: jostBook }]);
    const secondAllowed = font("Jost*", [{ weight: 400, data: jostBook }]);
    const full = await build(
      deck(
        createElement(
          Text,
          { testId: "same-font-bytes", x: 24, y: 18, width: 240, height: 36 },
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
            testId: "default-font-policy",
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

  it("完整嵌入接受 Editable 与 no-subsetting 标志，子集请求按许可拒绝", async () => {
    const original = jostBook;
    const noSubsetting = withFontFsType(original, 0x0108);
    const result = await build(
      <Deck>
        <Slide slideId="font-no-subsetting">
          <Text
            x={0}
            y={0}
            width={100}
            height={32}
            style={{
              fontFamily: font("Jost*", [{ weight: 400, data: noSubsetting }]),
            }}
          >
            完整字体
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );
    const parts = unzipSync(result.pptx);
    const embedded = parts["ppt/fonts/font1.fntdata"];
    assert.ok(embedded);
    const view = new DataView(embedded!.buffer, embedded!.byteOffset);
    const payloadStart = view.getUint32(0, true) - original.byteLength;
    assert.equal(view.getUint16(32, true), 0x0108);
    assert.deepEqual(embedded.subarray(payloadStart), noSubsetting);

    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="font-no-subsetting-subset">
            <Text
              testId="font-no-subsetting-subset"
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
              子集字体
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

  it("完整嵌入请求不能跳过缺少资源、格式或权限受限的实际字族", async () => {
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
                testId: label,
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
              testId: "subset-unsupported-woff2",
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
              testId: "subset-missing-glyph",
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
            testId: "font-value-opt-out-default",
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

  it("按 Unicode 断行机会和字素边界测量中文与超长词", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "cover" },
          createElement(Text, {
            testId: "cjk-breaks",
            x: 0,
            y: 0,
            width: 20,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "中文",
          }),
          createElement(Text, {
            testId: "long-word-breaks",
            x: 40,
            y: 0,
            width: 1,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "abcdef",
          }),
          createElement(Text, {
            testId: "combining-grapheme",
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
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        height: box?.height,
      })),
      [
        { testId: "cjk-breaks", height: 57.92 },
        { testId: "long-word-breaks", height: 173.76 },
        { testId: "combining-grapheme", height: 28.96 },
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

  it("超宽字素与显式换行各自只计一行且保留尾部空行", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "cover" },
          createElement(Text, {
            testId: "wide-grapheme-hard-break",
            x: 0,
            y: 0,
            width: 1,
            height: "auto",
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "a\nb",
          }),
          createElement(Text, {
            testId: "trailing-hard-break",
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

  it("代测字体必须显式声明并保留原输出字体名", async () => {
    const alternateFace = font(
      "Deck Display Name",
      [{ weight: 400, data: notoSansScRegular }],
      { substitute: true },
    );
    const result = await build(
      deck(
        createElement(Text, {
          testId: "substituted-font",
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

  it("内容定高拒绝没有覆盖所需字符的字体文件", async () => {
    await assert.rejects(
      build(
        deck(
          createElement(Text, {
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
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );
  });

  it("按准确字重选择度量文件，并要求实际斜体字重资源", async () => {
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

  it("按独立 SemiBold 字族的真实常规样式写入字重映射", async () => {
    const result = await build(
      deck(
        createElement(Text, {
          testId: "semibold-family-regular-face",
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

  it("度量到固定区域装不下时警告并保留区域、内容和原生默认行距", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "fixed-overset",
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

  it("带文字形状与文本框共用固定区域装不下诊断", async () => {
    const measureFont = font("Noto Sans CJK SC", [
      { weight: 400, data: notoSansScRegular },
    ]);
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "cover" },
          createElement(Text, {
            testId: "overset-text",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
            style: { fontFamily: measureFont, fontSize: 20 },
            children: "中文中文",
          }),
          createElement(Shape, {
            testId: "overset-shape",
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
  it("在基础及扩展转场前写出页面色彩映射覆盖", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          slideId: "fade-order",
          transition: transitions.fade(),
        }),
        createElement(Slide, {
          slideId: "flip-order",
          transition: transitions.flip({ fallback: transitions.fade() }),
        }),
        createElement(Slide, {
          slideId: "morph-order",
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

  it("并发验证正常的形状与文字产物", async () => {
    const shapeResult = await build(
      deck(
        createElement(Shape, {
          testId: "concurrent-shape",
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

  it("生成固定区域文字并保留中英与 Unicode 内容", async () => {
    const content = "固定区域文字 — React + 简体中文 😀";
    const result = await build(
      deck(
        createElement(Text, {
          testId: "mixed-text",
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

  it("把分段样式、固定行距和列表写成原生段落语义", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "formatted-text",
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

  it("段落继承文本框文字样式并只覆盖显式声明的字段", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "paragraph-inherited-style",
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

  it("同名字族、字重和样式引用冲突资源时报告来源", async () => {
    const element = deck(
      createElement(
        Text,
        { testId: "font-conflict", x: 12, y: 12, width: 300, height: 80 },
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

  it("同一字体值中实际使用的重复字重样式冲突资源会失败", async () => {
    const sharedValue = font("Duplicate Face", [
      { data: new Uint8Array([1, 2]) },
      { weight: 400, style: "normal", data: new Uint8Array([3, 4]) },
    ]);

    await assert.rejects(
      build(
        deck(
          createElement(
            Text,
            { testId: "duplicate-face", x: 12, y: 12, width: 300, height: 80 },
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

  it("仅名称引用与唯一同名字体资源共存时保留度量提示且成功构建", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "font-name-with-resource",
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

  it("固定区域可只声明字体名，报告度量限制且默认不嵌字体", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            testId: "missing-metrics",
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

  it("百分比高度的固定区域保留原生行距并按高度约定度量", async () => {
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
            { slideId: "fixed-text-region" },
            createElement(
              Box,
              { style: { width: 200, height: 80 } },
              createElement(
                Text,
                {
                  testId: "fixed-text",
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

  it("独立裁剪进程可在未加载文本框实现时生成纯形状和带文字形状", async () => {
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
            'const result = await build(createElement(Deck, { width: 120, height: 96 }, createElement(Slide, { slideId: "trimmed" }, createElement(Shape, { testId: "plain", x: 0, y: 0, width: 20, height: 20 }), createElement(Shape, { testId: "labeled", x: 24, y: 0, width: 60, height: 20, text: "形状文字 — Shape text" }), createElement(Shape, { testId: "empty-label", x: 90, y: 0, width: 20, height: 20, text: "" }))));',
            "const file = await validatePptx(result.pptx);",
            'if (file.errors.length !== 0 || result.report.slides[0]?.objects[1]?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[1]?.text?.text !== "形状文字 — Shape text" || file.slides[0]?.objects[2]?.text !== undefined) throw new Error(JSON.stringify(file.errors));',
            'function NestedText() { return createElement(TextDeclaration, { testId: "missing-text", x: 0, y: 0, width: 20, height: 20 }, "not available"); }',
            "let missingText;",
            'try { await build(createElement(Deck, {}, createElement(Slide, { slideId: "nested" }, createElement(NestedText)))); } catch (error) { missingText = error; }',
            'if (!(missingText instanceof BuildError) || missingText.diagnostics[0]?.code !== "missing-capability" || !missingText.diagnostics[0]?.path.includes("text[missing-text]")) throw new Error("nested Text did not produce a located missing-capability error");',
            'process.stdout.write(JSON.stringify({ kinds: file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]), missing: missingText.diagnostics[0], loaded: false }));',
          ].join("\n"),
        ],
        { cwd: process.cwd(), encoding: "utf8" },
      );
      assert.equal(child.status, 0, child.stderr || child.stdout);
      assert.deepEqual(JSON.parse(child.stdout), {
        kinds: [
          ["plain", ""],
          ["labeled", "形状文字 — Shape text"],
          ["empty-label", ""],
        ],
        missing: {
          severity: "error",
          path: "deck/slide[nested]/text[missing-text]",
          code: "missing-capability",
          message: "没有注册 rp:text 的能力模块",
        },
        loaded: false,
      });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000); // 独立进程首次加载 SDK WASM runtime 约需 10 秒。

  it("把真实 React 形状生成 Transitional PPTX 并保留只读身份观察", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "badge",
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
        testId: "badge",
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

  it("对公开构建的同一份 PPTX 执行 SDK 结构校验并保留定位", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "schema-gate",
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

  it("从 PPTX 实际读回时保留页面标识中的换行、回车和制表符", async () => {
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
          { slideId: slideIds[0] },
          createElement(Shape, {
            testId: "object-name",
            name: "object\nname\r\tend",
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(Shape, {
            testId: "test\nid",
            x: 50,
            y: 0,
            width: 40,
            height: 30,
          }),
        ),
        ...slideIds
          .slice(1)
          .map((slideId) => createElement(Slide, { slideId })),
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

  it("拒绝 XML 1.0 无法表示的对象名称和页面标识并定位来源", async () => {
    const invalidObjectName = deck(
      createElement(Shape, {
        testId: "invalid-name",
        name: "hello\u0000world",
        x: 0,
        y: 0,
        width: 40,
        height: 30,
      }),
    );
    const invalidTestId = deck(
      createElement(Shape, {
        testId: "hello\u0000world",
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
          testId: "invalid-text",
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
      createElement(Slide, { slideId: "alpha\u0000beta" }),
    );

    for (const [element, path] of [
      [invalidObjectName, /slide\[cover\].*shape\[invalid-name\]/u],
      [invalidTestId, /slide\[cover\].*shape\[/u],
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

  it("拒绝无法在 EMU 闭区间内表达的形状宽高", async () => {
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
              testId: "too-small",
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

  it("独立读取器接受形状宽高的最终 EMU 闭区间边界", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 5376, height: 5376 },
        createElement(
          Slide,
          { slideId: "extent-boundary" },
          createElement(Shape, {
            testId: "extent-boundary",
            x: 0,
            y: 0,
            // 该像素值按标准换算后舍入到形状尺寸上界。
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

  it("独立校验器拒绝超过格式上限的形状尺寸", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "oversized",
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

  it("按最终 EMU 限制页面尺寸并保留边界内的小数像素", async () => {
    const boundary = await build(
      createElement(
        Deck,
        { width: 95.99995, height: 5376.00005 },
        createElement(Slide, { slideId: "boundary" }),
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
            createElement(Slide, { slideId: "invalid-size" }),
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

  it("独立校验器拒绝超出格式范围的实际页面尺寸", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 96, height: 5376 },
        createElement(Slide, { slideId: "size" }),
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

  it("独立校验器拒绝低于格式下限的母版版式标识", async () => {
    const result = await build(
      deck(
        createElement(Shape, {
          testId: "badge",
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

  it("等待真实组件的布局与被动 Effect 提交并完成 root 卸载清理", async () => {
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
        testId: "from-component",
        x: 8,
        y: 9,
        width: 40,
        height: 30,
      });
    }

    const result = await build(deck(createElement(EffectShape)));

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId }) => testId),
      ["from-component"],
    );
    assert.deepEqual(events.sort(), [
      "layout-cleanup",
      "layout-commit",
      "passive-cleanup",
      "passive-commit",
    ]);
  });

  it("在声明路径报告未知形状参数与裁剪后缺失的能力", async () => {
    const UncheckedShape = Shape as ComponentType<Record<string, unknown>>;
    const unknownParameter = deck(
      createElement(UncheckedShape, {
        testId: "unknown-prop",
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
            testId: "trimmed-shape",
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

  it("显式形状模块沿同一链路遵守基础 Transitional 目标", async () => {
    const result = await createBuilder([shapeCapability])(
      deck(
        createElement(Shape, {
          testId: "triangle",
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

  it("生成并从同次 PPTX 读回基础页面转场", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(Slide, {
          slideId: "fade",
          transition: transitions.fade(),
        }),
        createElement(Slide, {
          slideId: "strips",
          transition: transitions.strips({ direction: "leftUp" }),
        }),
        createElement(Slide, {
          slideId: "push",
          transition: transitions.push({ from: "top" }),
        }),
        createElement(Slide, {
          slideId: "strips-default",
          transition: transitions.strips(),
        }),
        createElement(Slide, {
          slideId: "push-default",
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

  it("将每个方向参数写为对应 Transitional 方向代码", async () => {
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
          slideId: `strips-${index}`,
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
          slideId: `push-${index}`,
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

  it("按时长边界选择基础速度并写入取整后的精确毫秒", async () => {
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
            slideId: `duration-${index}`,
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

  it("独立文件检查发现 Choice 漏报的 p14 时长依赖", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          slideId: "morph-duration",
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

  it("精确时长写入 p14 与有依据的 Transitional 速度回退", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          slideId: "timed",
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
            slideId: "baseline-only",
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

  it("Flip 与 Morph 显式携带兼容分支并核对联合机制需求", async () => {
    const declarations = createElement(
      Deck,
      {},
      createElement(Slide, {
        slideId: "flip",
        transition: transitions.flip({
          direction: "left",
          fallback: transitions.fade(),
        }),
      }),
      createElement(Slide, {
        slideId: "morph",
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
        assert.equal(error.diagnostics[0]?.message, "输出目标禁止 p14");
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
        assert.equal(error.diagnostics[0]?.message, "输出目标禁止 p159");
        return true;
      },
    );
  });

  it("使用扩展转场默认选项并保留基础 fallback 声明", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(Slide, {
          slideId: "flip-default",
          transition: transitions.flip({ fallback: transitions.strips() }),
        }),
        createElement(Slide, {
          slideId: "morph-default",
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

  it("裁剪转场能力后定位被省略的 Slide 声明", async () => {
    await assert.rejects(
      createBuilder([shapeCapability])(
        createElement(
          Deck,
          {},
          createElement(Slide, {
            slideId: "missing-transition-module",
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

  it("按 Slide 路径拒绝非法时长、方向、选项和 fallback", async () => {
    const invalidTransitions: readonly {
      readonly slideId: string;
      readonly transition: unknown;
      readonly code: string;
    }[] = [
      {
        slideId: "negative-duration",
        transition: transitions.fade({ duration: -1 }),
        code: "invalid-transition-duration",
      },
      {
        slideId: "non-finite-duration",
        transition: transitions.fade({ duration: Number.POSITIVE_INFINITY }),
        code: "invalid-transition-duration",
      },
      {
        slideId: "unknown-parameter",
        transition: transitions.fade({
          duration: 600,
          surprise: true,
        } as never),
        code: "unknown-parameter",
      },
      {
        slideId: "null-strips-direction",
        transition: transitions.strips({ direction: null as never }),
        code: "invalid-transition-direction",
      },
      {
        slideId: "invalid-strips-direction",
        transition: transitions.strips({ direction: "diagonal" as never }),
        code: "invalid-transition-direction",
      },
      {
        slideId: "null-push-direction",
        transition: transitions.push({ from: null as never }),
        code: "invalid-transition-direction",
      },
      {
        slideId: "invalid-push-direction",
        transition: transitions.push({ from: "center" as never }),
        code: "invalid-transition-direction",
      },
      {
        slideId: "null-flip-direction",
        transition: transitions.flip({
          direction: null as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-transition-direction",
      },
      {
        slideId: "missing-flip-fallback",
        transition: transitions.flip({} as never),
        code: "missing-transition-fallback",
      },
      {
        slideId: "null-morph-option",
        transition: transitions.morph({
          by: null as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-morph-option",
      },
      {
        slideId: "invalid-morph-option",
        transition: transitions.morph({
          by: "sentence" as never,
          fallback: transitions.fade(),
        }),
        code: "invalid-morph-option",
      },
      {
        slideId: "helper-kind-and-unknown-option",
        transition: transitions.fade({ kind: "push", from: "left" } as never),
        code: "unknown-parameter",
      },
      {
        slideId: "extension-fallback",
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
              slideId: invalid.slideId,
              transition: invalid.transition as never,
            }),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, invalid.code);
          assert.equal(
            error.diagnostics[0]?.path,
            `deck/slide[${invalid.slideId}]`,
          );
          return true;
        },
      );
    }
  });

  it("越界保留原始形状并在生成证据中返回警告", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 100, height: 100 },
        createElement(
          Slide,
          { slideId: "cover" },
          createElement(Shape, {
            testId: "overflowing",
            x: 90,
            y: 70,
            width: 30,
            height: 20,
          }),
        ),
      ),
    );

    assert.deepEqual(result.report.diagnostics, [
      {
        severity: "warning",
        path: "deck/slide[cover]/shape[overflowing]",
        code: "overflow",
        message: "对象超出页面画布，构建保留原始声明",
      },
    ]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.boxEmu, {
      x: 857250,
      y: 666750,
      width: 285750,
      height: 190500,
    });
  });

  it("明确拒绝声明外形状和不能写入的尺寸", async () => {
    await assert.rejects(
      build(
        deck(
          createElement(Shape, {
            testId: "unknown-geometry",
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
            testId: "zero-size",
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
