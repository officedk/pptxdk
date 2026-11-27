import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  BuildError,
  darker,
  defineTheme,
  font,
  lighter,
  type ThemeScope,
} from "@officedk/pptxdk";
import {
  Box,
  Chart,
  Deck,
  Group,
  MotionScope,
  P,
  Slide,
  Shape,
  Span,
  Table,
  Row,
  Cell,
  Text,
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
          colors: { ink: "accent1", accent1: "#336699" },
          fonts: { major: cjkFont, minor: localFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
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

  it("inherits local fonts and colors along Box, Group, and transparent MotionScope while isolating sibling instances", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: {
            ink: "accent1",
            accent1: "#336699",
            brand: "accent2",
          },
          fonts: { major: cjkFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
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
            }}
            theme={{ colors: { ink: "#C00000" }, fonts: { major: localFont } }}
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
                    background: "ink",
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
                    color: "ink",
                    background: "ink",
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
                fontFamily: "major",
                color: "ink",
              }}
              axes={{
                category: {
                  kind: "category",
                  style: { fontFamily: "major", color: "ink" },
                  line: { color: "ink" },
                  gridlines: { color: "accent1" },
                },
                value: {
                  kind: "value",
                  style: { fontFamily: "major", color: "ink" },
                  line: { color: { color: "ink", alpha: 0.5 } },
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
                      color: "ink",
                      points: { 0: { color: "accent1", labelColor: "ink" } },
                    },
                  ],
                  dataLabels: {
                    value: true,
                    style: { fontFamily: "major", color: "ink" },
                  },
                },
              ]}
            />
            <Box
              id="nested-box"
              style={{ width: 180, height: 60 }}
              theme={{
                colors: { ink: "#008040", brand: "ink" },
                fonts: { major: cjkFont },
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
                  background: "brand",
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
          id="slide-scope"
          theme={{
            colors: { ink: "#123456" },
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
    assert.match(categoryAxis ?? "", /<a:srgbClr val="C00000"\/>/u);
    assert.match(valueAxis ?? "", /<a:alpha val="50000"\/>/u);
    assert.match(valueAxis ?? "", /<a:schemeClr val="accent1"\/>/u);
    assert.match(dataLabels ?? "", /<a:srgbClr val="C00000"\/>/u);
    assert.match(dataLabels ?? "", /<a:latin typeface="Jost\*"\/>/u);
    assert.match(chart, /<c:txPr>/u);
    const theme = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slideScope = strFromU8(parts["ppt/slides/slide2.xml"]!);
    assert.match(theme, /<a:majorFont><a:latin typeface="Theme Metrics Sans"/u);
    assert.match(theme, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/u);
    assert.match(slideScope, /<a:srgbClr val="123456"\/>/u);
    assert.match(slideScope, /<a:latin typeface="Jost\*"\/>/u);
    assert.match(slideScope, /<a:latin typeface="Theme Metrics Sans"\/>/u);
    assert.match(slideScope, /sz="1350"/u);
  });

  it("preserves native slots for aliases, tint/shade variants, and opacity along the same color resolver", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: {
            accent1: "#336699",
            brand: "accent1",
            pale: lighter("accent1", 0.4),
            deep: darker("accent1", 0.2),
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
              background: "brand",
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
              background: "pale",
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
              background: "deep",
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
              background: { color: "brand", alpha: 0.5 },
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

  it("resolves nearest local overrides for deck color aliases at usage sites while preserving outer references of the same name", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: {
            ink: "#112233",
            caption: "ink",
            chain: "caption",
          },
        })}
      >
        <Slide id="local-alias" theme={{ colors: { ink: "#C00000" } }}>
          <Text
            id="alias-text"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 180,
              height: 40,
              color: "chain",
            }}
          >
            Alias text
          </Text>
          <Shape
            id="alias-shape"
            style={{
              position: "absolute",
              left: 0,
              top: 50,
              width: 100,
              height: 30,
              background: "chain",
              color: "chain",
            }}
          >
            Alias shape
          </Shape>
          <Table
            id="alias-table"
            columns={[140]}
            style={{ position: "absolute", left: 0, top: 90 }}
          >
            <Row style={{ height: 36 }}>
              <Cell style={{ color: "chain", background: "chain" }}>
                Alias cell
              </Cell>
            </Row>
          </Table>
          <Chart
            id="alias-chart"
            style={{
              position: "absolute",
              left: 160,
              top: 90,
              width: 240,
              height: 150,
              color: "chain",
            }}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "Alias series", values: [1], color: "chain" }],
              },
            ]}
          />
          <Box
            id="same-name-scope"
            style={{ width: 100, height: 40 }}
            theme={{ colors: { ink: "ink" } }}
          >
            <Shape
              id="same-name-shape"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 80,
                height: 30,
                background: "ink",
              }}
            />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const chart = strFromU8(parts["ppt/charts/chart1.xml"]!);
    for (const authorId of ["alias-text", "alias-shape", "alias-table"]) {
      const id = findAuthorObject(
        result.report.slides[0]?.objects,
        authorId,
      )?.id;
      assert.match(
        slide,
        new RegExp(`id="${id}"[\\s\\S]*?<a:srgbClr val="C00000"/>`, "u"),
      );
    }
    assert.match(chart, /<a:srgbClr val="C00000"\/>/u);
    const sameNameId = findAuthorObject(
      result.report.slides[0]?.objects,
      "same-name-shape",
    )?.id;
    const sameNameStart = slide.indexOf(`id="${sameNameId}"`);
    assert.notEqual(sameNameStart, -1);
    assert.match(
      slide.slice(sameNameStart, sameNameStart + 1200),
      /val="C00000"/u,
    );
  });

  it("resolves the same font token only once across Text, P, and Span", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          fonts: { body: "Aptos", Aptos: "Courier New" },
        })}
      >
        <Slide id="font-token-once">
          <Text
            id="font-token-once"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 300,
              height: 120,
              fontFamily: "body",
            }}
          >
            Bare text
            <P style={{ fontFamily: "body" }}>Paragraph text</P>
            <Span style={{ fontFamily: "body" }}>Span text</Span>
          </Text>
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fontTokenId = findAuthorObject(
      result.report.slides[0]?.objects,
      "font-token-once",
    )?.id;
    const object = file.slides[0]?.objects.find(({ id }) => id === fontTokenId);
    const runs = object?.text?.paragraphs.flatMap(({ runs: paragraphRuns }) =>
      paragraphRuns.map(({ text, fontFamily }) => ({ text, fontFamily })),
    );
    assert.deepEqual(
      runs?.map(({ text, fontFamily }) => [text.trim(), fontFamily]),
      [
        ["Bare text", "Aptos"],
        ["Paragraph text", "Aptos"],
        ["Span text", "Aptos"],
      ],
    );
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
          colors: { ink: "#112233", caption: "ink" },
          fonts: { minor: cjkFont, body: cjkFont },
        })}
      >
        {variants.map(({ id: slideId, child }) => (
          <Slide id={slideId} key={slideId}>
            <Group
              id="outer-group"
              theme={{
                colors: { ink: "#C00000" },
                fonts: { minor: localFont, body: localFont },
              }}
              style={{ color: "caption", align: "center", spacingAfter: 12 }}
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
                color: "caption",
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

  it("rejects unknown color and font tokens in local themes", async () => {
    const theme = defineTheme({
      colors: { ink: "accent1" },
      fonts: { major: "Aptos Display" },
    });
    const scopedDeck = (scope: unknown) => (
      <Deck theme={theme}>
        <Slide id="invalid-scope">
          <Box
            id="scope"
            style={{ width: 100, height: 50 }}
            theme={scope as ThemeScope}
          >
            <Text
              id="sample"
              style={{ width: 80, height: "auto", color: "ink" }}
            >
              Sample
            </Text>
          </Box>
        </Slide>
      </Deck>
    );

    for (const scope of [
      { colors: { unknown: "#FF0000" } },
      { colors: { accent1: "#FF0000" } },
      { fonts: { heading: "Aptos Display" } },
    ]) {
      await assert.rejects(build(scopedDeck(scope)), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-prop");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[invalid-scope]/box[scope]",
        );
        return true;
      });
    }

    await assert.rejects(
      build(scopedDeck({ colors: { ink: "missing-token" } })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-style");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[invalid-scope]/box[scope]/text[sample]",
        );
        return true;
      },
    );
  });

  it("keeps local tokens isolated without cross-contamination during concurrent builds sharing a deck theme", async () => {
    const theme = defineTheme({ colors: { ink: "accent1" } });
    const deck = (slideId: string, color: string) => (
      <Deck theme={theme}>
        <Slide id={slideId}>
          <Box
            id="scope"
            style={{ width: 100, height: 50 }}
            theme={{ colors: { ink: color } }}
          >
            <Shape
              id="sample"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 80,
                height: 40,
                background: "ink",
              }}
            />
          </Box>
        </Slide>
      </Deck>
    );
    const [red, green] = await Promise.all([
      build(deck("red", "#CC0000")),
      build(deck("green", "#00AA00")),
    ]);
    const redSlide = strFromU8(unzipSync(red.pptx)["ppt/slides/slide1.xml"]!);
    const greenSlide = strFromU8(
      unzipSync(green.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(redSlide, /<a:srgbClr val="CC0000"\/>/u);
    assert.doesNotMatch(redSlide, /<a:srgbClr val="00AA00"\/>/u);
    assert.match(greenSlide, /<a:srgbClr val="00AA00"\/>/u);
    assert.doesNotMatch(greenSlide, /<a:srgbClr val="CC0000"\/>/u);
    assert.deepEqual((await validatePptx(red.pptx)).errors, []);
    assert.deepEqual((await validatePptx(green.pptx)).errors, []);
  });

  it("includes local theme fonts in font resource records of the same deck", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({ fonts: { major: "Aptos Display" } })}
        style={{ fontFamily: "major", fontSize: 20 }}
      >
        <Slide
          id="local-font"
          theme={{ fonts: { major: italicFont } }}
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
            Local font resource{" "}
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
    assert.match(presentation, /<p:embeddedFontLst>/u);
    assert.match(presentation, /<p:italic r:id="rId\d+"\/>/u);
    assert.ok(Object.keys(parts).some((path) => path.startsWith("ppt/fonts/")));
    assert.match(slide, /<a:latin typeface="Jost"\/>/u);
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
