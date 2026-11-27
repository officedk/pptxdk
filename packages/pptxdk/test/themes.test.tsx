import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  BuildError,
  CapabilityError,
  codegen,
  darker,
  defaultColorMap,
  defineSlideLayout,
  defineTheme,
  font,
  importPptx,
  lighter,
  schemeColorNames,
  type FontReference,
  type HexColor,
  type ThemeColor,
  type ThemeFontSlot,
  type ThemeInput,
  type ThemeOverride,
} from "@officedk/pptxdk";
import {
  Box,
  Chart,
  Deck,
  Group,
  MotionScope,
  P,
  Slide,
  SlideLayout,
  SlideMaster,
  Shape,
  Span,
  Table,
  Row,
  Cell,
  Text,
  linearGradient,
  useMotionScope,
} from "@officedk/pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import type { ReactNode } from "react";
import { describe, it } from "vitest";

const cjkFont = font(
  "Theme Metrics Sans",
  [
    {
      weight: 400,
      data: new Uint8Array(
        readFileSync(
          new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
        ),
      ),
    },
  ],
  { substitute: true },
);

const localFont = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
]);

const italicFont = font("Jost", [
  {
    weight: 400,
    style: "italic",
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Italic.ttf", import.meta.url),
      ),
    ),
  },
]);

function TestMotionScope({ children }: { readonly children?: ReactNode }) {
  const scope = useMotionScope();
  return <MotionScope scope={scope}>{children}</MotionScope>;
}

type NamedReportObject = {
  readonly id: number;
  readonly authorId?: string;
  readonly members?: readonly NamedReportObject[];
};

function findAuthorObject(
  objects: readonly NamedReportObject[] | undefined,
  authorId: string,
): NamedReportObject | undefined {
  const pending = [...(objects ?? [])];
  while (pending.length > 0) {
    const object = pending.shift();
    if (object === undefined) break;
    if (object.authorId === authorId) return object;
    pending.push(...(object.members ?? []));
  }
  return undefined;
}

describe("public theme entry", () => {
  it("participates in content measurement with theme fonts and preserves major font and native color slot references", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Theme Contract",
          colors: { accent1: "#336699" },
          fonts: { major: cjkFont, minor: localFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "accent1" }}
      >
        <Slide id="theme">
          <Text
            id="measured"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 400,
              height: "auto",
            }}
          >
            中文 A é
          </Text>
          <Text
            id="minor-font"
            style={{
              position: "absolute",
              left: 0,
              top: 60,
              width: 400,
              height: "auto",
              fontFamily: "minor",
              fontSize: 18,
            }}
          >
            Minor font
          </Text>
        </Slide>
      </Deck>,
    );

    assert.equal(result.report.slides[0]?.objects[0]?.box?.height, 28.96);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const parts = unzipSync(result.pptx);
    const theme = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.match(theme, /<a:theme[^>]+name="Theme Contract"/u);
    assert.match(theme, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/u);
    assert.match(
      theme,
      /<a:majorFont><a:latin typeface="Theme Metrics Sans"\/><a:ea typeface="Theme Metrics Sans"\/><a:cs typeface="Theme Metrics Sans"\/><\/a:majorFont>/u,
    );
    assert.match(
      theme,
      /<a:minorFont><a:latin typeface="Jost\*"\/><a:ea typeface="Jost\*"\/><a:cs typeface="Jost\*"\/>/u,
    );
    assert.match(
      slide,
      /<a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/>/u,
    );
    assert.match(
      slide,
      /<a:latin typeface="\+mn-lt"\/><a:ea typeface="\+mn-ea"\/><a:cs typeface="\+mn-cs"\/>/u,
    );
    assert.match(slide, /<a:schemeClr val="accent1"\/>/u);
  });

  it("applies Deck minor/dk1 defaults to normal text, shapes, tables, and charts", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Theme Defaults",
          colors: { dk1: "#223344", accent1: "#446688" },
          fonts: { minor: localFont },
        })}
      >
        <Slide id="theme-defaults">
          <Text
            id="default-text"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 220,
              height: "auto",
            }}
          >
            Default text
          </Text>
          <Text
            id="explicit-minor"
            style={{
              position: "absolute",
              left: 0,
              top: 50,
              width: 220,
              height: "auto",
              fontFamily: "minor",
            }}
          >
            Default text
          </Text>
          <Shape
            id="default-shape-text"
            style={{
              position: "absolute",
              left: 0,
              top: 100,
              width: 180,
              height: 48,
            }}
          >
            Shape text
          </Shape>
          <Table
            id="default-table"
            columns={[160]}
            style={{ position: "absolute", left: 0, top: 160 }}
          >
            <Row>
              <Cell>Cell text</Cell>
            </Row>
          </Table>
          <Chart
            id="default-chart"
            style={{
              position: "absolute",
              left: 200,
              top: 160,
              width: 300,
              height: 180,
            }}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Jan"] },
                series: [{ name: "Revenue", values: [2] }],
                dataLabels: { value: true },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const reportObjects = result.report.slides[0]?.objects ?? [];
    const defaultMeasurement = reportObjects.find(
      (object) => object.authorId === "default-text",
    )?.box?.height;
    const explicitMeasurement = reportObjects.find(
      (object) => object.authorId === "explicit-minor",
    )?.box?.height;
    assert.equal(defaultMeasurement, explicitMeasurement);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objects = file.slides[0]?.objects ?? [];
    const run = (testId: string, text: string) => {
      const id = findAuthorObject(result.report.slides[0]?.objects, testId)?.id;
      const object = objects.find((candidate) => candidate.id === id);
      assert.ok(object, `missing ${testId}`);
      const paragraphs =
        object.text?.paragraphs ?? object.table?.cellTexts[0]?.text.paragraphs;
      const found = paragraphs
        ?.flatMap((paragraph) => paragraph.runs)
        .find((item) => item.text === text);
      assert.ok(found, `missing ${text} in ${testId}`);
      assert.equal(found.fontFamily, "Jost*");
      assert.equal(found.color, "#223344");
    };
    run("default-text", "Default text");
    run("default-shape-text", "Shape text");
    run("default-table", "Cell text");

    const parts = unzipSync(result.pptx);
    const theme = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const chart = strFromU8(parts["ppt/charts/chart1.xml"]!);
    assert.match(theme, /<a:dk1><a:srgbClr val="223344"\/><\/a:dk1>/u);
    assert.match(theme, /<a:minorFont><a:latin typeface="Jost\*"/u);
    assert.match(slide, /<a:latin typeface="\+mn-lt"\/>/u);
    assert.match(slide, /<a:schemeClr val="tx1"\/>/u);
    assert.match(chart, /<c:txPr>.*?<a:latin typeface="\+mn-lt"\/>/su);
    assert.match(chart, /<a:schemeClr val="tx1"\/>/u);
  });

  it("inherits style fonts and colors along Box, Group, and transparent MotionScope and applies native Slide themeOverride", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: {
            dk1: "#336699",
            accent1: "#336699",
            accent2: "#C0504D",
          },
          fonts: { major: cjkFont, minor: cjkFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "dk1" }}
      >
        <Slide id="scopes">
          <Box
            id="local-box"
            style={{
              width: 600,
              height: 400,
              display: "flex",
              flexDirection: "column",
              fontSize: 19,
              fontFamily: localFont,
              color: "#C00000",
            }}
          >
            <TestMotionScope>
              <Group id="local-group" style={{ fontSize: 18 }}>
                <Text
                  id="local-text"
                  style={{ width: 300, height: "auto", fontSize: 20 }}
                >
                  Local scope
                </Text>
                <Shape
                  id="local-shape"
                  style={{
                    position: "absolute",
                    left: 0,
                    top: 60,
                    width: 180,
                    height: 60,
                    background: "#C00000",
                  }}
                >
                  Shape text
                </Shape>
              </Group>
            </TestMotionScope>
            <Table
              id="local-table"
              columns={[180]}
              style={{ position: "absolute", left: 0, top: 150 }}
            >
              <Row>
                <Cell
                  style={{
                    color: "#C00000",
                    background: "#C00000",
                    border: { color: "accent1" },
                  }}
                >
                  Cell text
                </Cell>
              </Row>
            </Table>
            <Chart
              id="local-chart"
              style={{
                position: "absolute",
                left: 200,
                top: 150,
                width: 220,
                height: 120,
                fontFamily: localFont,
                color: "#C00000",
              }}
              axes={{
                category: {
                  kind: "category",
                  style: { fontFamily: localFont, color: "#C00000" },
                  line: { color: "#C00000" },
                  gridlines: { color: "accent1" },
                },
                value: {
                  kind: "value",
                  style: { fontFamily: localFont, color: "#C00000" },
                  line: { color: { color: "#C00000", alpha: 0.5 } },
                  gridlines: { color: "accent1" },
                },
              }}
              groups={[
                {
                  kind: "column",
                  data: { kind: "shared-categories", categories: ["A"] },
                  series: [
                    {
                      name: "Local",
                      values: [1],
                      color: "#C00000",
                      points: {
                        0: { color: "accent1", labelColor: "#C00000" },
                      },
                    },
                  ],
                  dataLabels: {
                    value: true,
                    style: { fontFamily: localFont, color: "#C00000" },
                  },
                },
              ]}
            />
            <Box
              id="nested-box"
              style={{
                width: 180,
                height: 60,
                fontFamily: "major",
                color: "#008040",
              }}
            >
              <Text id="nested-text" style={{ width: 150, height: "auto" }}>
                Nested scope
              </Text>
              <Shape
                id="nested-shape"
                style={{
                  position: "absolute",
                  left: 0,
                  top: 30,
                  width: 40,
                  height: 20,
                  background: "#C00000",
                }}
              />
            </Box>
            <Text id="outer-after" style={{ width: 300, height: "auto" }}>
              Outer scope again
            </Text>
          </Box>
          <Text
            id="sibling-text"
            style={{
              position: "absolute",
              left: 0,
              top: 300,
              width: 300,
              height: "auto",
            }}
          >
            Global scope
          </Text>
        </Slide>
        <Slide
          id="slide-override"
          themeOverride={{
            colors: { dk1: "#123456" },
            fonts: { major: localFont, minor: cjkFont },
          }}
          style={{ fontSize: 18 }}
        >
          <Text
            id="slide-text"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 300,
              height: "auto",
            }}
          >
            Slide defaults
          </Text>
          <Text
            id="slide-minor"
            style={{
              position: "absolute",
              left: 0,
              top: 60,
              width: 300,
              height: "auto",
              fontFamily: "minor",
            }}
          >
            Slide minor default
          </Text>
        </Slide>
      </Deck>,
    );

    const textObjects = result.report.slides[0]?.objects ?? [];
    const box = result.report.slides[0]?.layouts?.find(
      (layout) => layout.authorId === "local-box",
    );
    assert.ok(box);
    const localHeight = textObjects
      .find((object) => object.authorId === "local-group")
      ?.members?.find((object) => object.authorId === "local-text")
      ?.box?.height;
    const siblingHeight = textObjects.find(
      (object) => object.authorId === "sibling-text",
    )?.box?.height;
    assert.equal(typeof localHeight, "number");
    assert.equal(typeof siblingHeight, "number");
    assert.notEqual(localHeight, siblingHeight);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const chart = strFromU8(parts["ppt/charts/chart1.xml"]!);
    const tableId = findAuthorObject(
      result.report.slides[0]?.objects,
      "local-table",
    )?.id;
    const tableStart = slide.indexOf(`id="${tableId}"`);
    const tableEnd = slide.indexOf("</p:graphicFrame>", tableStart);
    const table = slide.slice(tableStart, tableEnd);
    const categoryAxis = /<c:catAx>.*?<\/c:catAx>/su.exec(chart)?.[0];
    const valueAxis = /<c:valAx>.*?<\/c:valAx>/su.exec(chart)?.[0];
    const dataLabels = /<c:dLbls>.*?<\/c:dLbls>/su.exec(chart)?.[0];
    const objectXml = (name: string) => {
      const id = findAuthorObject(result.report.slides[0]?.objects, name)?.id;
      const start = slide.indexOf(`id="${id}"`);
      assert.notEqual(start, -1, `missing native object ${name}`);
      return slide.slice(start, slide.indexOf("</p:sp>", start));
    };
    assert.match(slide, /<a:latin typeface="Jost\*"\/>/u);
    assert.match(
      slide,
      /<a:solidFill><a:srgbClr val="C00000"\/><\/a:solidFill>/u,
    );
    assert.match(
      slide,
      /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill>/u,
    );
    assert.match(objectXml("nested-text"), /<a:srgbClr val="008040"\/>/u);
    assert.match(objectXml("nested-text"), /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(objectXml("nested-shape"), /<a:srgbClr val="C00000"\/>/u);
    assert.match(objectXml("outer-after"), /<a:srgbClr val="C00000"\/>/u);
    assert.match(objectXml("outer-after"), /<a:latin typeface="Jost\*"\/>/u);
    assert.match(objectXml("outer-after"), /sz="1425"/u);
    assert.match(objectXml("local-shape"), /sz="1350"/u);
    assert.match(table, /<a:srgbClr val="C00000"\/>/u);
    assert.match(table, /<a:latin typeface="Jost\*"\/>/u);
    assert.match(table, /sz="1425"/u);
    assert.match(table, /<a:schemeClr val="accent1"\/>/u);
    assert.match(chart, /<a:srgbClr val="C00000"\/>/u);
    assert.match(chart, /<a:latin typeface="Jost\*"\/>/u);
    assert.match(chart, /<a:schemeClr val="accent1"\/>/u);
    assert.match(categoryAxis ?? "", /<a:srgbClr val="C00000"\/>/u);
    assert.match(categoryAxis ?? "", /<a:latin typeface="Jost\*"\/>/u);
    assert.match(categoryAxis ?? "", /<a:schemeClr val="accent1"\/>/u);
    assert.match(valueAxis ?? "", /<a:alpha val="50000"\/>/u);
    assert.match(valueAxis ?? "", /<a:schemeClr val="accent1"\/>/u);
    assert.match(dataLabels ?? "", /<a:srgbClr val="C00000"\/>/u);
    assert.match(dataLabels ?? "", /<a:latin typeface="Jost\*"\/>/u);
    assert.match(chart, /<c:txPr>/u);
    const theme = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slideOverrideXml = strFromU8(parts["ppt/slides/slide2.xml"]!);
    const themeOverrideXml = strFromU8(parts["ppt/theme/themeOverride1.xml"]!);
    assert.match(theme, /<a:majorFont><a:latin typeface="Theme Metrics Sans"/u);
    assert.match(theme, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/u);
    assert.match(
      themeOverrideXml,
      /<a:dk1><a:srgbClr val="123456"\/><\/a:dk1>/u,
    );
    assert.match(themeOverrideXml, /<a:majorFont><a:latin typeface="Jost\*"/u);
    assert.match(slideOverrideXml, /<a:schemeClr val="dk1"\/>/u);
    assert.match(slideOverrideXml, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(slideOverrideXml, /<a:latin typeface="\+mn-lt"\/>/u);
    assert.match(slideOverrideXml, /sz="1350"/u);
  });

  it("preserves native slots for host ThemeShade constants, opacity, and #RGBA style hex values", async () => {
    const pale = lighter("accent1", 0.4);
    const deep = darker("accent1", 0.2);
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: {
            accent1: "#336699",
          },
        })}
      >
        <Slide id="colors">
          <Shape
            id="alias"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 80,
              height: 40,
              background: "accent1",
            }}
          />
          <Shape
            id="pale"
            style={{
              position: "absolute",
              left: 90,
              top: 0,
              width: 80,
              height: 40,
              background: pale,
            }}
          />
          <Shape
            id="deep"
            style={{
              position: "absolute",
              left: 180,
              top: 0,
              width: 80,
              height: 40,
              background: deep,
            }}
          />
          <Shape
            id="alpha"
            style={{
              position: "absolute",
              left: 270,
              top: 0,
              width: 80,
              height: 40,
              background: { color: "accent1", alpha: 0.5 },
            }}
          />
          <Shape
            id="short-alpha"
            style={{
              position: "absolute",
              left: 360,
              top: 0,
              width: 80,
              height: 40,
              background: "#1234",
            }}
          />
        </Slide>
      </Deck>,
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const slide = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(slide, /<a:schemeClr val="accent1"\/>/u);
    assert.match(slide, /<a:lumMod val="60000"\/><a:lumOff val="40000"\/>/u);
    assert.match(slide, /<a:lumMod val="80000"\/>/u);
    assert.match(slide, /<a:alpha val="50000"\/>/u);
    assert.match(
      slide,
      /<a:srgbClr val="112233"><a:alpha val="26667"\/><\/a:srgbClr>/u,
    );
  });

  it("applies Slide themeOverride and colorMapOverride consistently across Text, Shape, Table, and Chart and roundtrips #RGB normalization", async () => {
    const theme = defineTheme({
      name: "Normalized Brand",
      colors: {
        dk1: "#abc",
        lt1: "#ffffff",
        dk2: "#1a2b3c",
        lt2: "#fed",
        accent1: "#123456",
        accent2: "#234567",
        accent3: "#345678",
        accent4: "#456789",
        accent5: "#56789a",
        accent6: "#6789ab",
        hlink: "#00f",
        folHlink: "#808",
      },
      fonts: { major: "Georgia", minor: "Aptos" },
    });
    assert.equal(theme.colors.dk1, "#AABBCC");
    assert.equal(theme.colors.dk2, "#1A2B3C");
    assert.equal(theme.colors.lt2, "#FFEEDD");
    assert.equal(theme.colors.hlink, "#0000FF");
    assert.equal(theme.colors.folHlink, "#880088");

    const result = await build(
      <Deck theme={theme}>
        <Slide
          id="native-override"
          themeOverride={{ colors: { accent1: "#c00", dk1: "#112233" } }}
          colorMapOverride={{ ...defaultColorMap, tx1: "accent1" }}
        >
          <Text
            id="override-text"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 180,
              height: 40,
              color: "tx1",
            }}
          >
            Override text
          </Text>
          <Shape
            id="override-shape"
            style={{
              position: "absolute",
              left: 0,
              top: 50,
              width: 100,
              height: 30,
              background: "accent1",
              color: "dk1",
            }}
          >
            Override shape
          </Shape>
          <Table
            id="override-table"
            columns={[140]}
            style={{ position: "absolute", left: 0, top: 90 }}
          >
            <Row style={{ height: 36 }}>
              <Cell style={{ color: "tx1", background: "accent1" }}>
                Override cell
              </Cell>
            </Row>
          </Table>
          <Chart
            id="override-chart"
            style={{
              position: "absolute",
              left: 160,
              top: 90,
              width: 240,
              height: 150,
              color: "tx1",
            }}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [
                  { name: "Override series", values: [1], color: "accent1" },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const validated = await validatePptx(result.pptx);
    assert.deepEqual(validated.errors, []);
    const parts = unzipSync(result.pptx);
    const themeXml = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const overrideXml = strFromU8(parts["ppt/theme/themeOverride1.xml"]!);
    assert.match(themeXml, /<a:dk1><a:srgbClr val="AABBCC"\/><\/a:dk1>/u);
    assert.match(themeXml, /<a:dk2><a:srgbClr val="1A2B3C"\/><\/a:dk2>/u);
    assert.match(
      overrideXml,
      /<a:accent1><a:srgbClr val="CC0000"\/><\/a:accent1>/u,
    );
    assert.match(overrideXml, /<a:dk1><a:srgbClr val="112233"\/><\/a:dk1>/u);

    const slideFile = validated.slides[0];
    const overrideTextId = findAuthorObject(
      result.report.slides[0]?.objects,
      "override-text",
    )?.id;
    const textRun = slideFile?.objects.find(({ id }) => id === overrideTextId)
      ?.text?.paragraphs[0]?.runs[0];
    assert.equal(textRun?.color, "#CC0000");

    const imported = importPptx(result.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const reactCode = codegen(imported.document, { target: "react" });
    const vanillaCode = codegen(imported.document, { target: "vanilla" });
    assert.match(String(reactCode.files["deck.tsx"]), /#AABBCC/u);
    assert.match(String(reactCode.files["deck.tsx"]), /#CC0000/u);
    assert.match(String(vanillaCode.files["deck.ts"]), /#AABBCC/u);
    assert.match(String(vanillaCode.files["deck.ts"]), /#CC0000/u);
  });

  it("resolves system font families directly across Text, P, and Span without dictionary lookups", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          fonts: { major: "Georgia", minor: "Aptos" },
        })}
      >
        <Slide id="system-font-direct">
          <Text
            id="system-font-direct"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 300,
              height: 120,
              fontFamily: "Aptos",
            }}
          >
            Bare text
            <P style={{ fontFamily: "Courier New" }}>Paragraph text</P>
            <Span style={{ fontFamily: "minor" }}>Span text</Span>
          </Text>
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fontId = findAuthorObject(
      result.report.slides[0]?.objects,
      "system-font-direct",
    )?.id;
    const object = file.slides[0]?.objects.find(({ id }) => id === fontId);
    const runs = object?.text?.paragraphs.flatMap(({ runs: paragraphRuns }) =>
      paragraphRuns.map(({ text, fontFamily }) => ({ text, fontFamily })),
    );
    assert.deepEqual(
      runs?.map(({ text, fontFamily }) => [text.trim(), fontFamily]),
      [
        ["Bare text", "Aptos"],
        ["Paragraph text", "Courier New"],
        ["Span text", "Aptos"],
      ],
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:latin typeface="Aptos"\/>/u);
    assert.match(slideXml, /<a:latin typeface="Courier New"\/>/u);
    assert.match(slideXml, /<a:latin typeface="\+mn-lt"\/>/u);
  });

  it("participates in measurement with paragraph defaults on Deck, Slide, Box, and Group and allows descendant overrides", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({ fonts: { minor: localFont } })}
        style={{ align: "center", spacingAfter: 16 }}
      >
        <Slide id="deck-paragraph-defaults">
          <Text
            id="deck-auto"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 280,
              height: "auto",
            }}
          >
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Text
            id="deck-zero-gap"
            style={{
              position: "absolute",
              left: 300,
              top: 0,
              width: 280,
              height: "auto",
              spacingAfter: 0,
            }}
          >
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Chart
            id="deck-paragraph-chart"
            style={{
              position: "absolute",
              left: 0,
              top: 160,
              width: 280,
              height: 160,
            }}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "Series", values: [1] }],
              },
            ]}
          />
        </Slide>
        <Slide
          id="slide-paragraph-defaults"
          style={{ align: "right", spacingAfter: 12 }}
        >
          <Text
            id="slide-default"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 280,
              height: 100,
            }}
          >
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
        </Slide>
        <Slide id="box-paragraph-defaults">
          <Box
            id="paragraph-box"
            style={{ width: 600, height: 260, align: "left", spacingAfter: 8 }}
          >
            <Text
              id="box-auto"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 280,
                height: "auto",
              }}
            >
              <P>First paragraph</P>
              <P style={{ align: "right", spacingAfter: 4 }}>
                Second paragraph override
              </P>
            </Text>
            <Text
              id="box-zero-gap"
              style={{
                position: "absolute",
                left: 300,
                top: 0,
                width: 280,
                height: "auto",
                spacingAfter: 0,
              }}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Box>
          <Text
            id="box-sibling"
            style={{
              position: "absolute",
              left: 0,
              top: 280,
              width: 280,
              height: 100,
            }}
          >
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Box
            id="second-paragraph-box"
            style={{
              width: 600,
              height: 140,
              align: "right",
              spacingAfter: 10,
            }}
          >
            <Text
              id="second-box-default"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 280,
                height: 100,
              }}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Box>
        </Slide>
        <Slide id="group-paragraph-defaults">
          <Group
            id="paragraph-group"
            style={{ align: "justify", spacingAfter: 20 }}
          >
            <Text
              id="group-default"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 280,
                height: 100,
              }}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Group>
        </Slide>
      </Deck>,
    );

    const measuredHeight = (testId: string, slideIndex: number) => {
      const object = result.report.slides[slideIndex]?.objects.find(
        ({ authorId: objectTestId }) => objectTestId === testId,
      );
      const height = object?.box?.height;
      if (typeof height !== "number") assert.fail(`missing ${testId}`);
      return height;
    };
    assert.equal(
      Math.round(
        (measuredHeight("deck-auto", 0) - measuredHeight("deck-zero-gap", 0)) *
          100,
      ) / 100,
      16,
    );
    assert.equal(
      Math.round(
        (measuredHeight("box-auto", 2) - measuredHeight("box-zero-gap", 2)) *
          100,
      ) / 100,
      8,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideXml = (index: number) =>
      strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index + 1}.xml`]!);
    const objectXml = (index: number, testId: string) => {
      const xml = slideXml(index);
      const id = findAuthorObject(
        result.report.slides[index]?.objects,
        testId,
      )?.id;
      const start = xml.indexOf(`id="${id}"`);
      assert.notEqual(start, -1, `missing ${testId}`);
      const end = xml.indexOf("</p:sp>", start);
      return xml.slice(start, end);
    };
    assert.match(objectXml(0, "deck-auto"), /algn="ctr"/u);
    assert.match(objectXml(0, "deck-auto"), /<a:spcPts val="1200"\/>/u);
    const chart = strFromU8(unzipSync(result.pptx)["ppt/charts/chart1.xml"]!);
    assert.match(
      chart,
      /<a:pPr algn="ctr">.*?<a:spcAft><a:spcPts val="1200"/su,
    );
    assert.match(objectXml(1, "slide-default"), /algn="r"/u);
    assert.match(objectXml(1, "slide-default"), /<a:spcPts val="900"\/>/u);
    assert.match(objectXml(2, "box-auto"), /algn="l"/u);
    assert.match(objectXml(2, "box-auto"), /<a:spcPts val="600"\/>/u);
    assert.match(objectXml(2, "box-auto"), /algn="r"/u);
    assert.match(objectXml(2, "box-auto"), /<a:spcPts val="300"\/>/u);
    assert.match(objectXml(2, "box-sibling"), /algn="ctr"/u);
    assert.match(objectXml(2, "second-box-default"), /algn="r"/u);
    assert.match(objectXml(2, "second-box-default"), /<a:spcPts val="750"\/>/u);
    assert.match(objectXml(3, "group-default"), /algn="just"/u);
    assert.match(objectXml(3, "group-default"), /<a:spcPts val="1500"\/>/u);
  });

  it("does not change original layout participation criteria with text-only defaults on Group", async () => {
    const result = await build(
      <Deck>
        <Slide id="text-only-group-style">
          <Group
            id="text-only-group"
            style={{ align: "center", spacingAfter: 8 }}
          >
            <Text
              id="group-child"
              style={{
                position: "absolute",
                left: 10,
                top: 20,
                width: 120,
                height: 40,
              }}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Group>
        </Slide>
      </Deck>,
    );

    const group = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "text-only-group",
    );
    const child = group?.members?.find(
      ({ authorId: testId }) => testId === "group-child",
    );
    assert.deepEqual(child?.box, { x: 10, y: 20, width: 120, height: 40 });
    const slide = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    const groupId = findAuthorObject(
      result.report.slides[0]?.objects,
      "text-only-group",
    )?.id;
    assert.match(slide, new RegExp(`id="${groupId}"[\\s\\S]*?algn="ctr"`, "u"));
    assert.match(slide, /<a:spcPts val="600"\/>/u);
  });

  it("passes theme and text defaults of transparent Group through layout nodes while isolating external siblings", async () => {
    const leaf = (style?: {
      readonly width?: number;
      readonly position?: "relative";
    }) => (
      <Text
        id="leaf"
        style={{
          position: "absolute",
          left: 10,
          top: 20,
          width: 300,
          height: "auto",
          ...style,
        }}
      >
        <P>First paragraph</P>
        <P>Second paragraph</P>
      </Text>
    );
    const variants: { readonly id: string; readonly child: ReactNode }[] = [
      { id: "direct", child: leaf() },
      {
        id: "box",
        child: <Box style={{ width: 400 }}>{leaf()}</Box>,
      },
      {
        id: "nested-layout-group",
        child: <Group style={{ display: "flex", width: 400 }}>{leaf()}</Group>,
      },
      { id: "styled-leaf", child: leaf({ width: 300 }) },
      { id: "relative-leaf", child: leaf({ position: "relative" }) },
      {
        id: "motion-box",
        child: (
          <TestMotionScope>
            <Box style={{ width: 400 }}>{leaf()}</Box>
          </TestMotionScope>
        ),
      },
    ];
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: { dk1: "#112233" },
          fonts: { minor: cjkFont },
        })}
      >
        {variants.map(({ id: slideId, child }) => (
          <Slide id={slideId} key={slideId}>
            <Group
              id="outer-group"
              style={{
                fontFamily: localFont,
                color: "#C00000",
                align: "center",
                spacingAfter: 12,
              }}
            >
              {child}
            </Group>
            <Text
              id="sibling"
              style={{
                position: "absolute",
                left: 10,
                top: 120,
                width: 300,
                height: "auto",
                color: "dk1",
              }}
            >
              <P>Outside the Group</P>
            </Text>
          </Slide>
        ))}
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const findReportObject = (
      objects: (typeof result.report.slides)[number]["objects"],
      testId: string,
    ) => {
      const pending = [...objects];
      while (pending.length > 0) {
        const object = pending.shift();
        if (object === undefined) break;
        if (object.authorId === testId) return object;
        pending.push(...(object.members ?? []));
      }
      assert.fail(`missing report object ${testId}`);
    };
    const findFileObject = (slideIndex: number, testId: string) => {
      const id = findReportObject(
        result.report.slides[slideIndex]?.objects ?? [],
        testId,
      ).id;
      const pending = [...(file.slides[slideIndex]?.objects ?? [])];
      while (pending.length > 0) {
        const object = pending.shift();
        if (object === undefined) break;
        if (object.id === id) return object;
        pending.push(...(object.members ?? []));
      }
      assert.fail(`missing file object ${testId}`);
    };

    for (const [index, { id: slideId }] of variants.entries()) {
      const reportSlide = result.report.slides[index];
      const fileSlide = file.slides[index];
      assert.ok(reportSlide);
      assert.ok(fileSlide);
      const text = findReportObject(reportSlide.objects, "leaf");
      assert.equal(text.box?.height, 81.36, `${slideId} measured height`);
      if (slideId === "direct") {
        assert.deepEqual(text.box, { x: 10, y: 20, width: 300, height: 81.36 });
      }

      const output = findFileObject(index, "leaf");
      const runs = output.text?.paragraphs.flatMap(
        ({ runs: paragraphRuns }) => paragraphRuns,
      );
      assert.ok(runs);
      assert.ok(runs.length > 0);
      for (const run of runs) {
        assert.equal(run.fontFamily, "Jost*");
        assert.equal(run.color, "#C00000");
      }

      const sibling = findFileObject(index, "sibling");
      const siblingRun = sibling.text?.paragraphs
        .flatMap(({ runs: paragraphRuns }) => paragraphRuns)
        .find(({ text: value }) => value.includes("Outside the Group"));
      assert.equal(siblingRun?.fontFamily, "Theme Metrics Sans");
      assert.equal(siblingRun?.color, "#112233");

      const slide = strFromU8(
        unzipSync(result.pptx)[`ppt/slides/slide${index + 1}.xml`]!,
      );
      const leafId = findAuthorObject(
        result.report.slides[index]?.objects,
        "leaf",
      )?.id;
      const start = slide.indexOf(`id="${leafId}"`);
      assert.notEqual(start, -1);
      const objectXml = slide.slice(start, slide.indexOf("</p:sp>", start));
      assert.match(objectXml, /algn="ctr"/u);
      assert.match(objectXml, /<a:spcPts val="900"\/>/u);
    }
  });

  it("enforces compile-time and runtime closure on ThemeInput, ThemeOverride, ThemeColor, FontReference, and Slide/Box/Group props", async () => {
    // Compile-time type checks for narrowed public types
    const validColor: ThemeColor = "bg1";
    const validHex: HexColor = "#1a2b3c";
    const validFontSlot: ThemeFontSlot = "major";
    const validFontRef: FontReference = validFontSlot;
    assert.equal(validColor, "bg1");
    assert.equal(validHex, "#1a2b3c");
    assert.equal(validFontRef, "major");

    // @ts-expect-error unknown color slot in ThemeInput
    const badThemeColorKey: ThemeInput = { colors: { followHlink: "#112233" } };
    const badThemeShadeColor: ThemeInput = {
      colors: {
        // @ts-expect-error ThemeShade rejected in ThemeInput.colors
        dk1: lighter("accent1", 0.2),
      },
    };
    // @ts-expect-error scheme slot string rejected in ThemeInput.colors
    const badThemeSlotRef: ThemeInput = { colors: { dk1: "accent2" } };
    // @ts-expect-error unknown font slot in ThemeInput.fonts
    const badThemeFontKey: ThemeInput = { fonts: { minro: "Aptos" } };
    const badOverrideColorKey: ThemeOverride = {
      colors: {
        // @ts-expect-error unknown color slot in ThemeOverride
        tileAccent: "#112233",
      },
    };
    // @ts-expect-error unknown color string rejected on ThemeColor
    const badStyleColor: ThemeColor = "red";
    // @ts-expect-error misspelled scheme slot rejected on ThemeColor
    const badSchemeColor: ThemeColor = "acent1";
    assert.ok(badThemeColorKey);
    assert.ok(badThemeShadeColor);
    assert.ok(badThemeSlotRef);
    assert.ok(badThemeFontKey);
    assert.ok(badOverrideColorKey);
    assert.ok(badStyleColor);
    assert.ok(badSchemeColor);

    for (const invalidTheme of [
      { colors: { followHlink: "#112233" } },
      { colors: { tileAccent: "#112233" } },
      { colors: { accent1: "#11223380" } },
      { colors: { accent1: "#1234" } },
      { colors: { dk1: "accent2" } },
      { colors: { accent1: lighter("accent1", 0.2) } },
      { fonts: { minro: "Aptos" } },
      { fonts: { mono: "Consolas" } },
      { fonts: { major: "major" } },
      { fonts: { minor: "minor" } },
      { fonts: { major: "+mj-lt" } },
      { fonts: { minor: "+mn-ea" } },
    ]) {
      assert.throws(
        () => defineTheme(invalidTheme as never),
        (error: unknown) => {
          assert.ok(error instanceof CapabilityError);
          assert.equal(error.code, "invalid-theme");
          return true;
        },
      );

      await assert.rejects(
        build(
          <Deck theme={invalidTheme as never}>
            <Slide id="s1" />
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-prop");
          assert.equal(error.diagnostics[0]?.path, "deck");
          return true;
        },
      );

      const blank = defineSlideLayout({ id: "blank", type: "blank" });
      await assert.rejects(
        build(
          <Deck>
            <SlideMaster id="m1" theme={invalidTheme as never}>
              <SlideLayout layout={blank} />
            </SlideMaster>
            <Slide id="s1" layout={blank} />
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-prop");
          assert.equal(error.diagnostics[0]?.path, "deck/slide-master[m1]");
          return true;
        },
      );
    }

    // Rejects Slide/Box/Group local theme prop at compile time and runtime
    await assert.rejects(
      build(
        <Deck>
          {/* @ts-expect-error Slide.theme removed */}
          <Slide id="s1" theme={{ colors: { dk1: "#112233" } }} />
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unknown-prop" &&
        error.diagnostics[0]?.path === "deck/slide[s1]",
    );
    await assert.rejects(
      build(
        <Deck>
          <Slide id="s1">
            {/* @ts-expect-error Box.theme removed */}
            <Box id="b1" theme={{ colors: { dk1: "#112233" } }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unknown-prop" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/box[b1]",
    );
    await assert.rejects(
      build(
        <Deck>
          <Slide id="s1">
            {/* @ts-expect-error Group.theme removed */}
            <Group id="g1" theme={{ colors: { dk1: "#112233" } }}>
              <Shape id="sh1" style={{ width: 20, height: 20 }} />
            </Group>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unknown-prop" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/group[g1]",
    );

    // Rejects invalid style/border/shadow/gradient colors consistently
    for (const invalidColor of ["red", "tileAccent", "acent1"] as const) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id="bad-color">
              <Text
                id="t1"
                style={{ width: 100, height: 40, color: invalidColor as never }}
              >
                Sample
              </Text>
            </Slide>
          </Deck>,
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-style",
      );
      await assert.rejects(
        build(
          <Deck>
            <Slide id="bad-bg">
              <Shape
                id="sh1"
                style={{
                  width: 100,
                  height: 40,
                  background: invalidColor as never,
                }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-style",
      );
      await assert.rejects(
        build(
          <Deck>
            <Slide id="bad-border">
              <Shape
                id="sh1"
                style={{
                  width: 100,
                  height: 40,
                  border: { color: invalidColor as never, width: 1 },
                }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-style",
      );
      await assert.rejects(
        build(
          <Deck>
            <Slide id="bad-gradient">
              <Shape
                id="sh1"
                style={{
                  width: 100,
                  height: 40,
                  background: linearGradient(
                    90,
                    invalidColor as never,
                    "#FFFFFF",
                  ),
                }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-style",
      );
    }

    // All 16 SchemeColorNames are accepted on styles
    const allSchemeDeck = await build(
      <Deck>
        <Slide id="all-scheme">
          {schemeColorNames.map((name, idx) => (
            <Shape
              key={name}
              id={`s-${name}`}
              style={{
                position: "absolute",
                left: (idx % 4) * 40,
                top: Math.floor(idx / 4) * 30,
                width: 30,
                height: 20,
                background: name,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    assert.deepEqual((await validatePptx(allSchemeDeck.pptx)).errors, []);
  });

  it("includes Slide themeOverride fonts in font resource records of the same deck", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({ fonts: { major: "Aptos Display" } })}
        style={{ fontFamily: "major", fontSize: 20 }}
      >
        <Slide
          id="override-font"
          themeOverride={{ fonts: { major: italicFont } }}
          style={{ fontStyle: "italic" }}
        >
          <Text
            id="italic"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 300,
              height: "auto",
            }}
          >
            Override font resource{" "}
            <Span style={{ fontFamily: "major" }}>inline resource</Span>
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const presentation = strFromU8(parts["ppt/presentation.xml"]!);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const themeOverride = strFromU8(parts["ppt/theme/themeOverride1.xml"]!);
    assert.match(presentation, /<p:embeddedFontLst>/u);
    assert.match(presentation, /<p:italic r:id="rId\d+"\/>/u);
    assert.ok(Object.keys(parts).some((path) => path.startsWith("ppt/fonts/")));
    assert.match(themeOverride, /<a:majorFont><a:latin typeface="Jost"/u);
    assert.match(slide, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(slide, /i="1"/u);
    const italicId = findAuthorObject(
      result.report.slides[0]?.objects,
      "italic",
    )?.id;
    const inline = file.slides[0]?.objects
      .find(({ id }) => id === italicId)
      ?.text?.paragraphs.flatMap(({ runs }) => runs)
      .find(({ text }) => text.includes("inline resource"));
    assert.equal(inline?.fontFamily, "Jost");
    assert.equal(inline?.italic, true);
    assert.match(
      strFromU8(parts["ppt/theme/theme1.xml"]!),
      /<a:majorFont><a:latin typeface="Aptos Display"/u,
    );
  });

  it("binds latin, ea, and cs font slots for theme fonts and direct font families", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Three Font Slots Theme",
          fonts: { major: cjkFont, minor: localFont },
        })}
      >
        <Slide id="font-slots">
          <Text
            id="theme-major"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 300,
              height: "auto",
              fontFamily: "major",
            }}
          >
            Major Run
          </Text>
          <Text
            id="theme-minor"
            style={{
              position: "absolute",
              left: 0,
              top: 50,
              width: 300,
              height: "auto",
              fontFamily: "minor",
            }}
          >
            Minor Run
          </Text>
          <Text
            id="direct-family"
            style={{
              position: "absolute",
              left: 0,
              top: 100,
              width: 300,
              height: 40,
              fontFamily: "Direct Family Face",
            }}
          >
            Direct Family Run
          </Text>
        </Slide>
      </Deck>,
    );

    const parts = unzipSync(result.pptx);
    const themeXml = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);

    assert.match(
      themeXml,
      /<a:majorFont><a:latin typeface="Theme Metrics Sans"\/><a:ea typeface="Theme Metrics Sans"\/><a:cs typeface="Theme Metrics Sans"\/><\/a:majorFont>/u,
    );
    assert.match(
      themeXml,
      /<a:minorFont><a:latin typeface="Jost\*"\/><a:ea typeface="Jost\*"\/><a:cs typeface="Jost\*"\/>/u,
    );

    assert.match(
      slideXml,
      /<a:latin typeface="\+mj-lt"\/><a:ea typeface="\+mj-ea"\/><a:cs typeface="\+mj-cs"\/>/u,
    );
    assert.match(
      slideXml,
      /<a:latin typeface="\+mn-lt"\/><a:ea typeface="\+mn-ea"\/><a:cs typeface="\+mn-cs"\/>/u,
    );

    assert.match(
      slideXml,
      /<a:latin typeface="Direct Family Face"\/><a:ea typeface="Direct Family Face"\/><a:cs typeface="Direct Family Face"\/>/u,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });
});
