import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Box,
  Chart,
  Deck,
  defineTheme,
  font,
  Group,
  MotionScope,
  Slide,
  Shape,
  Table,
  Row,
  Cell,
  Text,
  useMotionScope,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import type { ReactNode } from "react";
import { describe, it } from "vitest";

const cjkFont = font("Theme Metrics Sans", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
], { substitute: true });

const localFont = font(
  "Local Jost",
  [
    {
      weight: 400,
      data: new Uint8Array(
        readFileSync(
          new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      ),
    },
  ],
  { substitute: true },
);

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
          fonts: { major: cjkFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
      >
        <Slide slideId="theme">
          <Text testId="measured" x={0} y={0} width={400} height="auto">
            中文 A é
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
    assert.match(slide, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(slide, /<a:schemeClr val="accent1"><\/a:schemeClr>/u);
  });

  it("局部字体和颜色沿 Box、Group、透明 MotionScope 继承且兄弟实例隔离", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          colors: { ink: "accent1", accent1: "#336699" },
          fonts: { major: cjkFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
      >
        <Slide slideId="scopes">
          <Box
            testId="local-box"
            style={{ width: 600, height: 280, display: "flex", flexDirection: "column" }}
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
                <Cell style={{ color: "ink", backgroundColor: "ink" }}>
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
              categoryAxis={{ style: { fontFamily: "major", color: "ink" } }}
            />
          </Box>
          <Text
            testId="sibling-text"
            x={0}
            y={300}
            width={300}
            height="auto"
          >
            Global scope
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
    assert.match(slide, /<a:latin typeface="Local Jost"\/>/u);
    assert.match(slide, /<a:solidFill><a:srgbClr val="C00000"><\/a:srgbClr><\/a:solidFill>/u);
    assert.match(slide, /<a:solidFill><a:schemeClr val="accent1"><\/a:schemeClr><\/a:solidFill>/u);
    assert.match(chart, /<a:srgbClr val="C00000"><\/a:srgbClr>/u);
    assert.match(chart, /<a:latin typeface="Local Jost"\/>/u);
  });
});
