import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Box,
  BuildError,
  Chart,
  darker,
  Deck,
  defineTheme,
  font,
  Group,
  lighter,
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
} from "@react-pptx/react-pptx";
import type { ThemeScope } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
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

describe("公开主题入口", () => {
  it("主题字体参与内容度量，文字保留 major 字体与原生颜色槽引用", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Theme Contract",
          colors: { ink: "accent1", accent1: "#336699" },
          fonts: { major: cjkFont, minor: localFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
      >
        <Slide slideId="theme">
          <Text testId="measured" x={0} y={0} width={400} height="auto">
            中文 A é
          </Text>
          <Text
            testId="minor-font"
            x={0}
            y={60}
            width={400}
            height="auto"
            style={{ fontFamily: "minor", fontSize: 18 }}
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
    assert.match(theme, /<a:majorFont><a:latin typeface="Theme Metrics Sans"/u);
    assert.match(theme, /<a:minorFont><a:latin typeface="Jost\*"/u);
    assert.match(slide, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(slide, /<a:latin typeface="\+mn-lt"\/>/u);
    assert.match(slide, /<a:schemeClr val="accent1"\/>/u);
  });

  it("Deck minor/dk1 默认值应用于普通文字、形状、表格和图表", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Theme Defaults",
          colors: { dk1: "#223344", accent1: "#446688" },
          fonts: { minor: localFont },
        })}
      >
        <Slide slideId="theme-defaults">
          <Text testId="default-text" x={0} y={0} width={220} height="auto">
            Default text
          </Text>
          <Text
            testId="explicit-minor"
            x={0}
            y={50}
            width={220}
            height="auto"
            style={{ fontFamily: "minor" }}
          >
            Default text
          </Text>
          <Shape
            testId="default-shape-text"
            x={0}
            y={100}
            width={180}
            height={48}
            text="Shape text"
          />
          <Table testId="default-table" columns={[160]} x={0} y={160}>
            <Row>
              <Cell>Cell text</Cell>
            </Row>
          </Table>
          <Chart
            testId="default-chart"
            kind="column"
            x={200}
            y={160}
            width={300}
            height={180}
            data={{
              categories: ["Jan"],
              series: [{ name: "Revenue", values: [2] }],
            }}
            dataLabels={{ value: true }}
          />
        </Slide>
      </Deck>,
    );

    const reportObjects = result.report.slides[0]?.objects ?? [];
    const defaultMeasurement = reportObjects.find(
      (object) => object.testId === "default-text",
    )?.box?.height;
    const explicitMeasurement = reportObjects.find(
      (object) => object.testId === "explicit-minor",
    )?.box?.height;
    assert.equal(defaultMeasurement, explicitMeasurement);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objects = file.slides[0]?.objects ?? [];
    const run = (testId: string, text: string) => {
      const object = objects.find(({ name }) => name === testId);
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

  it("局部字体和颜色沿 Box、Group、透明 MotionScope 继承且兄弟实例隔离", async () => {
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
        <Slide slideId="scopes">
          <Box
            testId="local-box"
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
              <Group testId="local-group" style={{ fontSize: 18 }}>
                <Text
                  testId="local-text"
                  width={300}
                  height="auto"
                  style={{ fontSize: 20 }}
                >
                  Local scope
                </Text>
                <Shape
                  testId="local-shape"
                  x={0}
                  y={60}
                  width={180}
                  height={60}
                  fill="ink"
                  text="Shape text"
                />
              </Group>
            </TestMotionScope>
            <Table testId="local-table" columns={[180]} x={0} y={150}>
              <Row>
                <Cell
                  style={{
                    color: "ink",
                    backgroundColor: "ink",
                    borderColor: "accent1",
                  }}
                >
                  Cell text
                </Cell>
              </Row>
            </Table>
            <Chart
              testId="local-chart"
              kind="column"
              x={200}
              y={150}
              width={220}
              height={120}
              style={{ fontFamily: "major", color: "ink" }}
              data={{
                categories: ["A"],
                series: [
                  {
                    name: "Local",
                    values: [1],
                    color: "ink",
                    points: { 0: { color: "accent1", labelColor: "ink" } },
                  },
                ],
              }}
              dataLabels={{
                value: true,
                style: { fontFamily: "major", color: "ink" },
              }}
              categoryAxis={{
                style: { fontFamily: "major", color: "ink" },
                line: { color: "ink" },
                gridlines: { color: "accent1" },
              }}
              valueAxis={{
                style: { fontFamily: "major", color: "ink" },
                line: { color: { color: "ink", alpha: 0.5 } },
                gridlines: { color: "accent1" },
              }}
            />
            <Box
              testId="nested-box"
              style={{ width: 180, height: 60 }}
              theme={{
                colors: { ink: "#008040", brand: "ink" },
                fonts: { major: cjkFont },
              }}
            >
              <Text testId="nested-text" width={150} height="auto">
                Nested scope
              </Text>
              <Shape
                testId="nested-shape"
                x={0}
                y={30}
                width={40}
                height={20}
                fill="brand"
              />
            </Box>
            <Text testId="outer-after" width={300} height="auto">
              Outer scope again
            </Text>
          </Box>
          <Text testId="sibling-text" x={0} y={300} width={300} height="auto">
            Global scope
          </Text>
        </Slide>
        <Slide
          slideId="slide-scope"
          theme={{
            colors: { ink: "#123456" },
            fonts: { major: localFont, minor: cjkFont },
          }}
          style={{ fontSize: 18 }}
        >
          <Text testId="slide-text" x={0} y={0} width={300} height="auto">
            Slide defaults
          </Text>
          <Text
            testId="slide-minor"
            x={0}
            y={60}
            width={300}
            height="auto"
            style={{ fontFamily: "minor" }}
          >
            Slide minor default
          </Text>
        </Slide>
      </Deck>,
    );

    const textObjects = result.report.slides[0]?.objects ?? [];
    const box = result.report.slides[0]?.layouts?.find(
      (layout) => layout.testId === "local-box",
    );
    assert.ok(box);
    const localHeight = textObjects
      .find((object) => object.testId === "local-group")
      ?.members?.find((object) => object.testId === "local-text")?.box?.height;
    const siblingHeight = textObjects.find(
      (object) => object.testId === "sibling-text",
    )?.box?.height;
    assert.equal(typeof localHeight, "number");
    assert.equal(typeof siblingHeight, "number");
    assert.notEqual(localHeight, siblingHeight);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const chart = strFromU8(parts["ppt/charts/chart1.xml"]!);
    const tableStart = slide.indexOf('name="local-table"');
    const tableEnd = slide.indexOf("</p:graphicFrame>", tableStart);
    const table = slide.slice(tableStart, tableEnd);
    const categoryAxis = /<c:catAx>.*?<\/c:catAx>/su.exec(chart)?.[0];
    const valueAxis = /<c:valAx>.*?<\/c:valAx>/su.exec(chart)?.[0];
    const dataLabels = /<c:dLbls>.*?<\/c:dLbls>/su.exec(chart)?.[0];
    const objectXml = (name: string) => {
      const start = slide.indexOf(`name="${name}"`);
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

  it("别名、明暗变体和透明度沿同一颜色解析器保留原生槽位", async () => {
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
        <Slide slideId="colors">
          <Shape
            testId="alias"
            x={0}
            y={0}
            width={80}
            height={40}
            fill="brand"
          />
          <Shape
            testId="pale"
            x={90}
            y={0}
            width={80}
            height={40}
            fill="pale"
          />
          <Shape
            testId="deep"
            x={180}
            y={0}
            width={80}
            height={40}
            fill="deep"
          />
          <Shape
            testId="alpha"
            x={270}
            y={0}
            width={80}
            height={40}
            fill={{ color: "brand", alpha: 0.5 }}
          />
          <Shape
            testId="short-alpha"
            x={360}
            y={0}
            width={80}
            height={40}
            fill="#1234"
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

  it("文稿颜色别名在使用处解析最近局部覆盖并保留同名外层引用", async () => {
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
        <Slide slideId="local-alias" theme={{ colors: { ink: "#C00000" } }}>
          <Text
            testId="alias-text"
            x={0}
            y={0}
            width={180}
            height={40}
            style={{ color: "chain" }}
          >
            Alias text
          </Text>
          <Shape
            testId="alias-shape"
            x={0}
            y={50}
            width={100}
            height={30}
            fill="chain"
            text="Alias shape"
            textStyle={{ color: "chain" }}
          />
          <Table testId="alias-table" columns={[140]} x={0} y={90}>
            <Row style={{ height: 36 }}>
              <Cell style={{ color: "chain", backgroundColor: "chain" }}>
                Alias cell
              </Cell>
            </Row>
          </Table>
          <Chart
            testId="alias-chart"
            kind="column"
            x={160}
            y={90}
            width={240}
            height={150}
            style={{ color: "chain" }}
            data={{
              categories: ["A"],
              series: [{ name: "Alias series", values: [1], color: "chain" }],
            }}
          />
          <Box
            testId="same-name-scope"
            style={{ width: 100, height: 40 }}
            theme={{ colors: { ink: "ink" } }}
          >
            <Shape
              testId="same-name-shape"
              x={0}
              y={0}
              width={80}
              height={30}
              fill="ink"
            />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const chart = strFromU8(parts["ppt/charts/chart1.xml"]!);
    assert.match(slide, /name="alias-text"[\s\S]*?<a:srgbClr val="C00000"\/>/u);
    assert.match(
      slide,
      /name="alias-shape"[\s\S]*?<a:srgbClr val="C00000"\/>/u,
    );
    assert.match(
      slide,
      /name="alias-table"[\s\S]*?<a:srgbClr val="C00000"\/>/u,
    );
    assert.match(chart, /<a:srgbClr val="C00000"\/>/u);
    const sameNameStart = slide.indexOf('name="same-name-shape"');
    assert.notEqual(sameNameStart, -1);
    assert.match(
      slide.slice(sameNameStart, sameNameStart + 1200),
      /val="C00000"/u,
    );
  });

  it("Text、P 和 Span 对同一字体令牌只解析一次", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          fonts: { body: "Aptos", Aptos: "Courier New" },
        })}
      >
        <Slide slideId="font-token-once">
          <Text
            testId="font-token-once"
            x={0}
            y={0}
            width={300}
            height={120}
            style={{ fontFamily: "body" }}
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
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "font-token-once",
    );
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

  it("Deck、Slide、Box 和 Group 的段落默认值参与度量并允许后代覆盖", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({ fonts: { minor: localFont } })}
        style={{ align: "center", spacingAfter: 16 }}
      >
        <Slide slideId="deck-paragraph-defaults">
          <Text testId="deck-auto" x={0} y={0} width={280} height="auto">
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Text
            testId="deck-zero-gap"
            x={300}
            y={0}
            width={280}
            height="auto"
            style={{ spacingAfter: 0 }}
          >
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Chart
            testId="deck-paragraph-chart"
            kind="column"
            x={0}
            y={160}
            width={280}
            height={160}
            data={{
              categories: ["A"],
              series: [{ name: "Series", values: [1] }],
            }}
          />
        </Slide>
        <Slide
          slideId="slide-paragraph-defaults"
          style={{ align: "right", spacingAfter: 12 }}
        >
          <Text testId="slide-default" x={0} y={0} width={280} height={100}>
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
        </Slide>
        <Slide slideId="box-paragraph-defaults">
          <Box
            testId="paragraph-box"
            style={{ width: 600, height: 260, align: "left", spacingAfter: 8 }}
          >
            <Text testId="box-auto" x={0} y={0} width={280} height="auto">
              <P>First paragraph</P>
              <P style={{ align: "right", spacingAfter: 4 }}>
                Second paragraph override
              </P>
            </Text>
            <Text
              testId="box-zero-gap"
              x={300}
              y={0}
              width={280}
              height="auto"
              style={{ spacingAfter: 0 }}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Box>
          <Text testId="box-sibling" x={0} y={280} width={280} height={100}>
            <P>First paragraph</P>
            <P>Second paragraph</P>
          </Text>
          <Box
            testId="second-paragraph-box"
            style={{
              width: 600,
              height: 140,
              align: "right",
              spacingAfter: 10,
            }}
          >
            <Text
              testId="second-box-default"
              x={0}
              y={0}
              width={280}
              height={100}
            >
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Box>
        </Slide>
        <Slide slideId="group-paragraph-defaults">
          <Group
            testId="paragraph-group"
            style={{ align: "justify", spacingAfter: 20 }}
          >
            <Text testId="group-default" x={0} y={0} width={280} height={100}>
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Group>
        </Slide>
      </Deck>,
    );

    const measuredHeight = (testId: string, slideIndex: number) => {
      const object = result.report.slides[slideIndex]?.objects.find(
        ({ testId: objectTestId }) => objectTestId === testId,
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
      const start = xml.indexOf(`name="${testId}"`);
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

  it("Group 纯文字默认值不改变原有布局参与条件", async () => {
    const result = await build(
      <Deck>
        <Slide slideId="text-only-group-style">
          <Group
            testId="text-only-group"
            style={{ align: "center", spacingAfter: 8 }}
          >
            <Text testId="group-child" x={10} y={20} width={120} height={40}>
              <P>First paragraph</P>
              <P>Second paragraph</P>
            </Text>
          </Group>
        </Slide>
      </Deck>,
    );

    const group = result.report.slides[0]?.objects.find(
      ({ testId }) => testId === "text-only-group",
    );
    const child = group?.members?.find(
      ({ testId }) => testId === "group-child",
    );
    assert.deepEqual(child?.box, { x: 10, y: 20, width: 120, height: 40 });
    const slide = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(slide, /name="text-only-group"[\s\S]*?algn="ctr"/u);
    assert.match(slide, /<a:spcPts val="600"\/>/u);
  });

  it("透明 Group 的主题与文字默认值穿过布局节点并隔离外部兄弟", async () => {
    const leaf = (style?: {
      readonly width?: number;
      readonly position?: "relative";
    }) => (
      <Text
        testId="leaf"
        x={10}
        y={20}
        width={300}
        height="auto"
        {...(style === undefined ? {} : { style })}
      >
        <P>First paragraph</P>
        <P>Second paragraph</P>
      </Text>
    );
    const variants: { readonly slideId: string; readonly child: ReactNode }[] =
      [
        { slideId: "direct", child: leaf() },
        {
          slideId: "box",
          child: <Box style={{ width: 400 }}>{leaf()}</Box>,
        },
        {
          slideId: "nested-layout-group",
          child: (
            <Group style={{ display: "flex", width: 400 }}>{leaf()}</Group>
          ),
        },
        { slideId: "styled-leaf", child: leaf({ width: 300 }) },
        { slideId: "relative-leaf", child: leaf({ position: "relative" }) },
        {
          slideId: "motion-box",
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
        {variants.map(({ slideId, child }) => (
          <Slide slideId={slideId} key={slideId}>
            <Group
              testId="outer-group"
              theme={{
                colors: { ink: "#C00000" },
                fonts: { minor: localFont, body: localFont },
              }}
              style={{ color: "caption", align: "center", spacingAfter: 12 }}
            >
              {child}
            </Group>
            <Text
              testId="sibling"
              x={10}
              y={120}
              width={300}
              height="auto"
              style={{ color: "caption" }}
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
        if (object.testId === testId) return object;
        pending.push(...(object.members ?? []));
      }
      assert.fail(`missing report object ${testId}`);
    };
    const findFileObject = (
      objects: NonNullable<(typeof file.slides)[number]>["objects"],
      testId: string,
    ) => {
      const pending = [...objects];
      while (pending.length > 0) {
        const object = pending.shift();
        if (object === undefined) break;
        if (object.name === testId) return object;
        pending.push(...(object.members ?? []));
      }
      assert.fail(`missing file object ${testId}`);
    };

    for (const [index, { slideId }] of variants.entries()) {
      const reportSlide = result.report.slides[index];
      const fileSlide = file.slides[index];
      assert.ok(reportSlide);
      assert.ok(fileSlide);
      const text = findReportObject(reportSlide.objects, "leaf");
      assert.equal(text.box?.height, 81.36, `${slideId} measured height`);
      if (slideId === "direct") {
        assert.deepEqual(text.box, { x: 10, y: 20, width: 300, height: 81.36 });
      }

      const output = findFileObject(fileSlide.objects, "leaf");
      const runs = output.text?.paragraphs.flatMap(
        ({ runs: paragraphRuns }) => paragraphRuns,
      );
      assert.ok(runs);
      assert.ok(runs.length > 0);
      for (const run of runs) {
        assert.equal(run.fontFamily, "Jost*");
        assert.equal(run.color, "#C00000");
      }

      const sibling = findFileObject(fileSlide.objects, "sibling");
      const siblingRun = sibling.text?.paragraphs
        .flatMap(({ runs: paragraphRuns }) => paragraphRuns)
        .find(({ text: value }) => value.includes("Outside the Group"));
      assert.equal(siblingRun?.fontFamily, "Theme Metrics Sans");
      assert.equal(siblingRun?.color, "#112233");

      const slide = strFromU8(
        unzipSync(result.pptx)[`ppt/slides/slide${index + 1}.xml`]!,
      );
      const start = slide.indexOf('name="leaf"');
      assert.notEqual(start, -1);
      const objectXml = slide.slice(start, slide.indexOf("</p:sp>", start));
      assert.match(objectXml, /algn="ctr"/u);
      assert.match(objectXml, /<a:spcPts val="900"\/>/u);
    }
  });

  it("局部主题拒绝未知颜色与字体令牌", async () => {
    const theme = defineTheme({
      colors: { ink: "accent1" },
      fonts: { major: "Aptos Display" },
    });
    const scopedDeck = (scope: unknown) => (
      <Deck theme={theme}>
        <Slide slideId="invalid-scope">
          <Box
            testId="scope"
            style={{ width: 100, height: 50 }}
            theme={scope as ThemeScope}
          >
            <Text
              testId="sample"
              width={80}
              height="auto"
              style={{ color: "ink" }}
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
        assert.equal(error.diagnostics[0]?.code, "invalid-theme-scope");
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
        assert.equal(error.diagnostics[0]?.code, "invalid-color");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[invalid-scope]/box[scope]/text[sample]",
        );
        return true;
      },
    );
  });

  it("并发构建共享文稿主题时保持局部令牌互不污染", async () => {
    const theme = defineTheme({ colors: { ink: "accent1" } });
    const deck = (slideId: string, color: string) => (
      <Deck theme={theme}>
        <Slide slideId={slideId}>
          <Box
            testId="scope"
            style={{ width: 100, height: 50 }}
            theme={{ colors: { ink: color } }}
          >
            <Shape
              testId="sample"
              x={0}
              y={0}
              width={80}
              height={40}
              fill="ink"
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

  it("局部主题字体进入同一文稿的字型资源记录", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({ fonts: { major: "Aptos Display" } })}
        style={{ fontFamily: "major", fontSize: 20 }}
      >
        <Slide
          slideId="local-font"
          theme={{ fonts: { major: italicFont } }}
          style={{ fontStyle: "italic" }}
        >
          <Text testId="italic" x={0} y={0} width={300} height="auto">
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
    const inline = file.slides[0]?.objects
      .find(({ name }) => name === "italic")
      ?.text?.paragraphs.flatMap(({ runs }) => runs)
      .find(({ text }) => text.includes("inline resource"));
    assert.equal(inline?.fontFamily, "Jost");
    assert.equal(inline?.italic, true);
    assert.match(
      strFromU8(parts["ppt/theme/theme1.xml"]!),
      /<a:majorFont><a:latin typeface="Aptos Display"/u,
    );
  });
});
