import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  Box,
  build,
  BuildError,
  Chart,
  Cell,
  Deck,
  fade,
  font,
  Group,
  Image,
  Shape,
  Slide,
  Row,
  Table,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const cjkFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
]);
const wideImage = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

describe("public layout entry", () => {
  it("uses Slide as a canvas grid container, omits hidden subtrees, and changes only paint order with zIndex", async () => {
    const result = await build(
      <Deck width={320} height={180}>
        <Slide
          slideId="slide-grid"
          style={{
            display: "grid",
            gridTemplateColumns: "2fr 1fr",
            gridTemplateRows: "100px",
          }}
        >
          <Shape
            testId="front"
            height={30}
            enter={fade()}
            style={{ aspectRatio: 2, zIndex: 1 }}
          />
          <Shape
            testId="back"
            width={100}
            height={40}
            enter={fade()}
            style={{ gridColumn: 2, gridRow: 1, zIndex: -1 }}
          />
          <Shape
            testId="hidden"
            style={{ display: "none", width: 500, height: 500 }}
          />
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ id, testId, box }) => ({
        id,
        testId,
        box,
      })),
      [
        {
          id: 3,
          testId: "back",
          box: { x: 213.33, y: 0, width: 100, height: 40 },
        },
        { id: 2, testId: "front", box: { x: 0, y: 0, width: 60, height: 30 } },
      ],
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(
      slideXml.indexOf('name="back"') < slideXml.indexOf('name="front"'),
    );
    assert.equal(slideXml.includes('name="hidden"'), false);
    const motion = result.report.slides[0]?.capabilities?.["motion"] as {
      readonly steps: readonly {
        readonly effects: readonly {
          readonly target: { readonly testId?: string };
        }[];
      }[];
    };
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target }) => target.testId),
      ["front", "back"],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("uses respective native sizing models for Table and Chart inside Grid", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="grid-native-layout">
          <Box
            testId="grid-native"
            style={{
              display: "grid",
              width: 300,
              height: 180,
              gridTemplateColumns: "1fr 1fr",
              gridTemplateRows: "1fr",
              columnGap: 10,
            }}
          >
            <Table
              testId="grid-table"
              columns={["1fr"]}
              style={{ fontFamily: cjkFont, fontSize: 12, gridColumn: 1 }}
            >
              <Row>
                <Cell>Table</Cell>
              </Row>
            </Table>
            <Chart
              testId="grid-chart"
              kind="line"
              style={{ gridColumn: 2 }}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            />
          </Box>
        </Slide>
      </Deck>,
    );
    const objects = result.report.slides[0]?.objects ?? [];
    assert.deepEqual(
      objects.map(({ kind, testId, box }) => ({ kind, testId, box })),
      [
        {
          kind: "table",
          testId: "grid-table",
          box: { x: 0, y: 0, width: 145, height: 180 },
        },
        {
          kind: "chart",
          testId: "grid-chart",
          box: { x: 155, y: 0, width: 145, height: 180 },
        },
      ],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("starts Grid tracks after container padding", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="grid-padding">
          <Box
            style={{
              display: "grid",
              width: 200,
              height: 100,
              padding: [10, 20],
              gridTemplateColumns: "1fr",
              gridTemplateRows: "1fr",
            }}
          >
            <Shape testId="inside-padding" />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 20,
      y: 10,
      width: 160,
      height: 80,
    });
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("respects text min-content for auto grid tracks and matches native object geometry", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="grid-min-content">
          <Box
            style={{
              display: "grid",
              width: 140,
              height: 50,
              gridTemplateColumns: "auto 1fr",
              gridTemplateRows: "1fr",
              columnGap: 10,
            }}
          >
            <Text
              testId="min-content"
              style={{ fontFamily: cjkFont, fontSize: 20 }}
            >
              WWWW
            </Text>
            <Shape testId="after-min-content" width={20} height={20} />
          </Box>
        </Slide>
      </Deck>,
    );
    const objects = result.report.slides[0]?.objects ?? [];
    const text = objects.find(({ testId }) => testId === "min-content");
    const afterText = objects.find(
      ({ testId }) => testId === "after-min-content",
    );
    assert.ok(text?.box !== undefined && text.box.width > 0);
    assert.equal(afterText?.box?.x, text.box.x + text.box.width + 10);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const nativeText = file.slides[0]?.objects.find(
      ({ name }) => name === "min-content",
    );
    assert.deepEqual(
      {
        x: (nativeText?.boxEmu.x ?? 0) / 9525,
        y: (nativeText?.boxEmu.y ?? 0) / 9525,
        width: (nativeText?.boxEmu.width ?? 0) / 9525,
        height: (nativeText?.boxEmu.height ?? 0) / 9525,
      },
      text?.box,
    );
  });

  it("uses nearest explicitly positioned ancestor for absolute descendants within Grid items", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="grid-absolute-descendant">
          <Box style={{ position: "relative", width: 200, height: 120 }}>
            <Box
              style={{
                display: "grid",
                width: 100,
                height: 80,
                marginLeft: 30,
                marginTop: 40,
                gridTemplateColumns: "1fr",
                gridTemplateRows: "1fr",
              }}
            >
              <Box style={{ width: 50, height: 40 }}>
                <Shape
                  testId="absolute-descendant"
                  style={{
                    position: "absolute",
                    left: 10,
                    top: 20,
                    width: 10,
                    height: 10,
                  }}
                />
              </Box>
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 10,
      y: 20,
      width: 10,
      height: 10,
    });
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("rejects unimplemented placement for absolute grid items and unsupported grid baseline", async () => {
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide slideId="absolute-grid-placement">
            <Box style={{ display: "grid", width: 200, height: 100 }}>
              <Shape
                testId="misplaced"
                style={{ position: "absolute", gridColumn: 1 }}
              />
            </Box>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unsupported-grid-positioning",
    );
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide slideId="grid-baseline">
            <Box style={{ display: "grid", alignItems: "baseline" }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unsupported-grid-alignment",
    );
  });

  it("participates in parent flow layout as a native object for Group with layout style", async () => {
    const result = await build(
      <Deck width={320} height={140}>
        <Slide slideId="group-flow">
          <Box style={{ display: "flex", width: 300, height: 100, gap: 10 }}>
            <Group
              testId="flow-group"
              style={{
                display: "flex",
                flexDirection: "row",
                gap: 5,
                width: 100,
                height: 60,
              }}
            >
              <Shape testId="group-a" width={20} height={20} />
              <Shape testId="group-b" width={20} height={20} />
            </Group>
            <Shape testId="after-group" width={30} height={30} />
          </Box>
        </Slide>
      </Deck>,
    );
    const group = result.report.slides[0]?.objects.find(
      ({ testId }) => testId === "flow-group",
    );
    assert.deepEqual(group?.box, { x: 0, y: 0, width: 100, height: 60 });
    assert.deepEqual(
      group?.members?.map(({ testId, box }) => ({ testId, box })),
      [
        { testId: "group-a", box: { x: 0, y: 0, width: 20, height: 20 } },
        { testId: "group-b", box: { x: 25, y: 0, width: 20, height: 20 } },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.find(
        ({ testId }) => testId === "after-group",
      )?.box,
      { x: 110, y: 0, width: 30, height: 30 },
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("writes padding into bodyPr and participates in text height and baseline for Text and Shape.text", async () => {
    const result = await build(
      <Deck width={320} height={160}>
        <Slide slideId="text-insets">
          <Box style={{ display: "flex", alignItems: "baseline", width: 300 }}>
            <Text
              testId="plain-text"
              width={80}
              height="auto"
              style={{ fontFamily: cjkFont, fontSize: 20 }}
            >
              AB
            </Text>
            <Text
              testId="padded-text"
              width={80}
              height="auto"
              style={{ fontFamily: cjkFont, fontSize: 20, padding: [10, 4] }}
            >
              AB
            </Text>
            <Shape
              testId="padded-shape"
              width={100}
              height={50}
              text="中文中文中文"
              textStyle={{
                fontFamily: cjkFont,
                fontSize: 20,
                verticalAlign: "top",
              }}
              style={{
                paddingTop: 2,
                paddingRight: 3,
                paddingBottom: 4,
                paddingLeft: 5,
              }}
            />
          </Box>
        </Slide>
      </Deck>,
    );
    const objects = result.report.slides[0]?.objects ?? [];
    const plain = objects.find(({ testId }) => testId === "plain-text")?.box;
    const padded = objects.find(({ testId }) => testId === "padded-text")?.box;
    assert.equal(padded?.height, (plain?.height ?? 0) + 20);
    assert.equal(padded?.y, (plain?.y ?? 0) - 10);
    const shape = objects.find(({ testId }) => testId === "padded-shape")?.box;
    assert.deepEqual(shape, { x: 160, y: 8, width: 100, height: 50 });
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const bodyPrFor = (testId: string) => {
      const id = objects.find(
        ({ testId: objectTestId }) => objectTestId === testId,
      )?.id;
      const start = slideXml.indexOf(`id="${id}"`);
      const body = slideXml.indexOf("<a:bodyPr", start);
      return slideXml.slice(body, slideXml.indexOf("</a:bodyPr>", body));
    };
    assert.match(
      bodyPrFor("padded-text"),
      /lIns="38100" tIns="95250" rIns="38100" bIns="95250"/u,
    );
    assert.match(
      bodyPrFor("padded-shape"),
      /lIns="47625" tIns="19050" rIns="28575" bIns="38100"/u,
    );
    assert.ok(
      result.report.diagnostics.some(
        ({ code, path }) =>
          code === "overset" && path.endsWith("shape[padded-shape]"),
      ),
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("places content according to repeat/minmax grid tracks and satisfies cross-column constraints", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="grid-tracks">
          <Box
            testId="grid"
            style={{
              display: "grid",
              width: 200,
              height: 100,
              gridTemplateColumns: "repeat(2, minmax(60px, 1fr))",
              gridTemplateRows: "auto auto",
              columnGap: 10,
              rowGap: 5,
            }}
          >
            <Shape testId="first" height={20} />
            <Shape testId="span" height={10} style={{ gridColumn: "span 2" }} />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        { testId: "first", box: { x: 0, y: 0, width: 95, height: 20 } },
        { testId: "span", box: { x: 0, y: 25, width: 200, height: 10 } },
      ],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("evaluates relative and absolute positioning by horizontal and vertical percentages of containing block respectively", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="position-percent">
          <Box
            testId="anchor"
            style={{ position: "relative", width: 200, height: 100 }}
          >
            <Box style={{ width: 100, height: 40 }}>
              <Shape
                testId="absolute"
                style={{
                  position: "absolute",
                  left: "10%",
                  top: "20%",
                  width: "25%",
                  height: 10,
                }}
              />
            </Box>
            <Shape
              testId="relative"
              width={50}
              height={20}
              style={{ position: "relative", left: "10%", top: "50%" }}
            />
            <Shape
              testId="absolute-native-coordinates"
              x={70}
              y={30}
              width={10}
              height={10}
              style={{ position: "absolute" }}
            />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        { testId: "absolute", box: { x: 20, y: 20, width: 50, height: 10 } },
        { testId: "relative", box: { x: 20, y: 90, width: 50, height: 20 } },
        {
          testId: "absolute-native-coordinates",
          box: { x: 70, y: 30, width: 10, height: 10 },
        },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.find(
        ({ testId }) => testId === "absolute-native-coordinates",
      )?.box,
      { x: 70, y: 30, width: 10, height: 10 },
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("uses slide as containing block for percentage positioning of direct Slide children, with diagnostics for static offset", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide slideId="page-positioning">
          <Shape
            testId="page-absolute"
            style={{
              position: "absolute",
              right: "10%",
              bottom: "10%",
              width: "25%",
              height: 10,
            }}
          />
          <Shape
            testId="page-relative"
            x={10}
            y={20}
            width={20}
            height={10}
            style={{ position: "relative", left: "10%", top: "10%" }}
          />
        </Slide>
      </Deck>,
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "page-absolute",
          box: { x: 208, y: 206, width: 80, height: 10 },
        },
        {
          testId: "page-relative",
          box: { x: 42, y: 44, width: 20, height: 10 },
        },
      ],
    );
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide slideId="static-offset">
            <Shape
              testId="invalid-offset"
              width={20}
              height={10}
              style={{ left: 10 }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-layout-context",
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("preserves block layout with explicit display:block, and rejects flex-specific properties on block containers", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "explicit-block-layout" },
          createElement(
            Box,
            {
              testId: "block",
              style: { display: "block", width: 100, height: 100 },
            },
            createElement(Shape, { testId: "member", height: 20 }),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 20,
    });
    assert.deepEqual(result.report.slides[0]?.layouts?.[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    });
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "invalid-block-flex-style" },
            createElement(Box, {
              testId: "block",
              style: { display: "block", justifyContent: "center" },
            }),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-layout-context" &&
            path === "deck/slide[invalid-block-flex-style]/box[block]",
        ),
    );
  });

  it("stacks vertically and stretches horizontally in block containers, keeping off-canvas explicit geometry in place", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "block-layout" },
          createElement(
            Box,
            { testId: "column", style: { width: 200, height: 10 } },
            createElement(Shape, {
              testId: "first",
              height: 20,
            }),
            createElement(Shape, {
              testId: "second",
              height: 10,
            }),
          ),
          createElement(Shape, {
            testId: "explicit",
            x: 24,
            y: 32,
            width: 50,
            height: 40,
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "first",
          box: { x: 0, y: 0, width: 200, height: 20 },
        },
        {
          testId: "second",
          box: { x: 0, y: 20, width: 200, height: 10 },
        },
        {
          testId: "explicit",
          box: { x: 24, y: 32, width: 50, height: 40 },
        },
      ],
    );
    assert.equal(result.report.slides[0]?.layouts?.[0]?.testId, "column");
    assert.deepEqual(result.report.slides[0]?.layouts?.[0]?.box, {
      x: 0,
      y: 0,
      width: 200,
      height: 10,
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 200, height: 20 },
        { x: 0, y: 20, width: 200, height: 10 },
        { x: 24, y: 32, width: 50, height: 40 },
      ],
    );
  });

  it("overrides explicit dimensions of native objects with auto dimensions and lays out according to content", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "auto-layout-size" },
          createElement(
            Box,
            { style: { width: 100, height: 100 } },
            createElement(Shape, {
              testId: "auto-text-shape",
              width: 20,
              height: 10,
              text: "中文中文中文",
              textStyle: { fontFamily: cjkFont, fontSize: 20 },
              style: { width: "auto", height: "auto" },
            }),
          ),
        ),
      ),
    );

    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 57.92,
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.boxEmu, {
      x: 0,
      y: 0,
      width: 100 * 9525,
      height: 57.92 * 9525,
    });
  });

  it("stretches lines by default in wrapping flex containers, aligning from line start with explicit flex-start", async () => {
    const buildWrapped = (alignContent?: "flex-start") =>
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            {
              slideId:
                alignContent === undefined ? "wrap-stretch" : "wrap-start",
            },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  flexWrap: "wrap",
                  width: 100,
                  height: 100,
                  ...(alignContent === undefined ? {} : { alignContent }),
                },
              },
              createElement(Shape, { testId: "first", width: 60, height: 20 }),
              createElement(Shape, { testId: "second", width: 60, height: 20 }),
            ),
          ),
        ),
      );

    const stretched = await buildWrapped();
    assert.deepEqual(
      stretched.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 50],
    );
    const stretchedFile = await validatePptx(stretched.pptx);
    assert.deepEqual(stretchedFile.errors, []);
    assert.deepEqual(
      stretchedFile.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [0, 50],
    );

    const rowStart = await buildWrapped("flex-start");
    assert.deepEqual(
      rowStart.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 20],
    );
    const rowStartFile = await validatePptx(rowStart.pptx);
    assert.deepEqual(rowStartFile.errors, []);
    assert.deepEqual(
      rowStartFile.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [0, 20],
    );
  });

  it("shrinks flex items to content min-width, and removes this floor with minWidth: 0", async () => {
    const makeText = (testId: string, minWidth?: number) =>
      createElement(Text, {
        testId,
        width: 100,
        height: "auto",
        style: {
          fontFamily: cjkFont,
          fontSize: 20,
          ...(minWidth === undefined ? {} : { minWidth }),
        },
        children: "中文中文中文",
      });
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 500 },
        createElement(
          Slide,
          { slideId: "flex-min-content" },
          createElement(
            Box,
            {
              testId: "constrained",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("constrained-a"),
            makeText("constrained-b"),
          ),
          createElement(
            Box,
            {
              testId: "released",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("released-a", 0),
            makeText("released-b", 0),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "constrained-a",
          box: { x: 0, y: 0, width: 20, height: 200 },
        },
        {
          testId: "constrained-b",
          box: { x: 20, y: 0, width: 20, height: 200 },
        },
        {
          testId: "released-a",
          box: { x: 0, y: 200, width: 15, height: 200 },
        },
        {
          testId: "released-b",
          box: { x: 15, y: 200, width: 15, height: 200 },
        },
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "overset",
          path: "deck/slide[flex-min-content]/box[released]/text[released-a]",
        },
        {
          code: "overset",
          path: "deck/slide[flex-min-content]/box[released]/text[released-b]",
        },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 20, height: 200 },
        { x: 20, y: 0, width: 20, height: 200 },
        { x: 0, y: 200, width: 15, height: 200 },
        { x: 15, y: 200, width: 15, height: 200 },
      ],
    );
  });

  it("uses descendant dimensions and minWidth contributions for nested auto min-width, while explicit zero remains shrinkable", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "nested-minimums" },
          createElement(
            Box,
            {
              testId: "fixed-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "fixed-content" },
              createElement(Shape, {
                testId: "fixed-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "min-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "min-content" },
              createElement(Shape, {
                testId: "min-shape",
                width: 20,
                height: 20,
                style: { minWidth: 150 },
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "released-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "released-content", style: { minWidth: 0 } },
              createElement(Shape, {
                testId: "released-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              testId: "margin-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { testId: "margin-content" },
              createElement(Shape, {
                testId: "margin-shape",
                width: 100,
                height: 20,
                style: { marginLeft: 10, marginRight: 15 },
              }),
            ),
          ),
        ),
      ),
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    const layoutWidth = (testId: string) =>
      layouts.find((layout) => layout.testId === testId)?.box.width;
    assert.equal(layoutWidth("fixed-content"), 200);
    assert.equal(layoutWidth("min-content"), 150);
    assert.equal(layoutWidth("released-content"), 100);
    assert.equal(layoutWidth("margin-content"), 125);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        x: box?.x,
        width: box?.width,
      })),
      [
        { testId: "fixed-shape", x: 0, width: 200 },
        { testId: "min-shape", x: 0, width: 150 },
        { testId: "released-shape", x: 0, width: 200 },
        { testId: "margin-shape", x: 10, width: 100 },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        width: boxEmu.width / 9525,
      })),
      [
        { x: 0, width: 200 },
        { x: 0, width: 150 },
        { x: 0, width: 200 },
        { x: 10, width: 100 },
      ],
    );
  });

  it("preserves instance identities for nested layout containers, solving each independently when reusing identical elements", async () => {
    const card = createElement(
      Box,
      { testId: "card", style: { flex: 1, height: 100 } },
      createElement(Shape, { testId: "card-shape", height: 20 }),
    );
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "reused-layout" },
          createElement(
            Box,
            {
              testId: "row",
              style: { display: "flex", width: 300, height: 120 },
            },
            card,
            card,
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "card-shape",
          box: { x: 0, y: 0, width: 150, height: 20 },
        },
        {
          testId: "card-shape",
          box: { x: 150, y: 0, width: 150, height: 20 },
        },
      ],
    );
    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.deepEqual(
      layouts.map(({ testId }) => testId),
      ["row", "card", "card"],
    );
    const row = layouts.find(({ testId }) => testId === "row");
    const cards = layouts.filter(({ testId }) => testId === "card");
    assert.equal(row?.box.width, 300);
    assert.equal(cards.length, 2);
    assert.deepEqual(
      row?.children.map(({ kind }) => kind),
      ["layout", "layout"],
    );
    assert.notEqual(row?.children[0]?.id, row?.children[1]?.id);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 0, width: 150, height: 20 },
        { x: 150, y: 0, width: 150, height: 20 },
      ],
    );
  });

  it("preserves member structure in native groups and solves transparent layout containers inside groups", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "group-with-layout" },
          createElement(
            Group,
            { testId: "native-group" },
            createElement(
              Box,
              {
                testId: "inner-box",
                style: { width: 100, height: 100 },
              },
              createElement(Shape, { testId: "group-member", height: 20 }),
            ),
          ),
        ),
      ),
    );

    const group = result.report.slides[0]?.objects[0];
    const member = group?.members?.[0];
    assert.equal(group?.kind, "group");
    assert.equal(group?.testId, "native-group");
    assert.deepEqual(member?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 20,
    });
    const layout = result.report.slides[0]?.layouts?.[0];
    assert.equal(layout?.testId, "inner-box");
    assert.deepEqual(layout?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    });
    assert.deepEqual(layout?.children, [{ kind: "object", id: member?.id }]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.kind, "group");
    assert.equal(
      file.slides[0]?.objects[0]?.members?.[0]?.name,
      "group-member",
    );
    assert.deepEqual(file.slides[0]?.objects[0]?.members?.[0]?.boxEmu, {
      x: 0,
      y: 0,
      width: 100 * 9525,
      height: 20 * 9525,
    });
  });

  it("rejects flex-specific properties on block containers and ignores flex item properties under block parents", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "invalid-context" },
            createElement(Box, { style: { justifyContent: "center" } }),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-layout-context"),
    );

    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "ignored-item-style" },
          createElement(
            Box,
            { style: { width: 100 } },
            createElement(Shape, {
              testId: "block-item",
              width: 30,
              height: 10,
              style: { flex: 1, alignSelf: "baseline" },
            }),
          ),
        ),
      ),
    );
    assert.deepEqual(result.report.slides[0]?.objects[0]?.box, {
      x: 0,
      y: 0,
      width: 30,
      height: 10,
    });
  });

  it("preserves resolved dimensions and returns warnings when containers and child objects overflow", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "overflow" },
          createElement(
            Box,
            { testId: "wide-container", style: { width: 400, height: 50 } },
            createElement(Shape, { testId: "wide-shape", height: 10 }),
          ),
        ),
      ),
    );

    assert.equal(result.report.slides[0]?.layouts?.[0]?.box.width, 400);
    assert.equal(result.report.slides[0]?.objects[0]?.box?.width, 400);
    assert.deepEqual(
      result.report.diagnostics.map(({ path, code }) => ({ path, code })),
      [
        { path: "deck/slide[overflow]/box[wide-container]", code: "overflow" },
        {
          path: "deck/slide[overflow]/box[wide-container]/shape[wide-shape]",
          code: "overflow",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.boxEmu.width, 400 * 9525);
  });

  it("writes native objects for images and shapes according to flex-aligned positions", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "aligned-objects" },
          createElement(
            Box,
            {
              testId: "toolbar",
              style: {
                display: "flex",
                width: 100,
                height: 100,
                justifyContent: "space-between",
                alignItems: "center",
              },
            },
            createElement(Image, {
              testId: "logo",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              testId: "action",
              width: 20,
              height: 20,
            }),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ kind, testId, box }) => ({
        kind,
        testId,
        box,
      })),
      [
        {
          kind: "image",
          testId: "logo",
          box: { x: 0, y: 40, width: 30, height: 20 },
        },
        {
          kind: "shape",
          testId: "action",
          box: { x: 80, y: 40, width: 20, height: 20 },
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        x: boxEmu.x / 9525,
        y: boxEmu.y / 9525,
        width: boxEmu.width / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { x: 0, y: 40, width: 30, height: 20 },
        { x: 80, y: 40, width: 20, height: 20 },
      ],
    );
  });

  it("aligns baseline according to text baselines of measurable fonts and rejects unfounded text combinations", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "text-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                alignItems: "baseline",
                width: 300,
                height: 100,
              },
            },
            createElement(
              Text,
              {
                testId: "small-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
            createElement(
              Text,
              {
                testId: "large-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 40 },
              },
              "中",
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "small-text", y: 23.2 },
        { testId: "large-text", y: 0 },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [23.2, 0],
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "unmeasured-baseline" },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                  height: 100,
                },
              },
              createElement(
                Text,
                { testId: "unmeasured-text", width: 100, height: 60 },
                "text",
              ),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "baseline-measurement-unavailable" &&
            path ===
              "deck/slide[unmeasured-baseline]/box[1]/text[unmeasured-text]",
        ),
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { slideId: "wrapped-text-baseline" },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "baseline",
                  width: 100,
                  height: 100,
                },
              },
              createElement(
                Text,
                {
                  testId: "wrapped-text",
                  width: 60,
                  height: 20,
                  style: { fontFamily: cjkFont, fontSize: 20 },
                },
                "中",
              ),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "unsupported-baseline-context",
        ),
    );
  });

  it("uses Yoga synthesized edge baseline for pure boxes and column-direction baseline", async () => {
    const wrapped = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "synthetic-wrap-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexWrap: "wrap",
                alignItems: "baseline",
                width: 100,
                height: 100,
              },
            },
            createElement(Shape, {
              testId: "wrap-first",
              width: 60,
              height: 20,
            }),
            createElement(Shape, {
              testId: "wrap-second",
              width: 60,
              height: 20,
            }),
          ),
        ),
      ),
    );
    assert.deepEqual(
      wrapped.report.slides[0]?.objects.map(({ box }) => box?.y),
      [0, 50],
    );
    assert.deepEqual((await validatePptx(wrapped.pptx)).errors, []);

    const column = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "synthetic-column-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "column",
                alignItems: "baseline",
                width: 120,
                height: 120,
              },
            },
            createElement(
              Text,
              {
                testId: "column-small-text",
                width: 100,
                height: 40,
                style: { fontFamily: "Aptos", fontSize: 20 },
              },
              "中",
            ),
            createElement(Image, {
              testId: "column-image",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              testId: "column-shape",
              width: 80,
              height: 20,
            }),
          ),
        ),
      ),
    );
    assert.equal(
      column.report.diagnostics.some(
        ({ code }) => code === "baseline-measurement-unavailable",
      ),
      false,
    );
    assert.deepEqual((await validatePptx(column.pptx)).errors, []);
  });

  it("includes shifted text boxes in adaptive baseline line height and maintains subsequent block positions", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { slideId: "natural-baseline-height" },
          createElement(
            Box,
            { style: { width: 300 } },
            createElement(
              Box,
              {
                testId: "baseline-row",
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                },
              },
              createElement(
                Text,
                {
                  testId: "small-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 20 },
                },
                "中",
              ),
              createElement(
                Text,
                {
                  testId: "large-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 40 },
                },
                "中",
              ),
            ),
            createElement(Shape, {
              testId: "following-block",
              width: 100,
              height: 20,
            }),
          ),
        ),
      ),
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.equal(
      layouts.find(({ testId }) => testId === "baseline-row")?.box.height,
      83.2,
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "small-text", y: 23.2 },
        { testId: "large-text", y: 0 },
        { testId: "following-block", y: 83.2 },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      [23.2, 0, 83.2],
    );
  });

  it("participates in automatic minimums, permits explicit release, and passes through simple Box for baseline reference", async () => {
    const baselineText = (testId: string, fontSize: number) =>
      createElement(
        Text,
        {
          testId,
          width: 100,
          height: 60,
          style: { fontFamily: cjkFont, fontSize },
        },
        "中",
      );
    const buildColumn = (minHeight?: number) =>
      build(
        createElement(
          Deck,
          { width: 320, height: 120 },
          createElement(
            Slide,
            {
              slideId:
                minHeight === 0 ? "baseline-min-zero" : "baseline-auto-min",
            },
            createElement(
              Box,
              {
                testId: "column",
                style: {
                  display: "flex",
                  flexDirection: "column",
                  width: 300,
                  height: 70,
                },
              },
              createElement(
                Box,
                {
                  testId: "baseline-row",
                  style: {
                    display: "flex",
                    alignItems: "baseline",
                    width: 300,
                    ...(minHeight === undefined ? {} : { minHeight }),
                  },
                },
                baselineText("small-text", 20),
                baselineText("large-text", 40),
              ),
              createElement(Shape, {
                testId: "following-block",
                width: 100,
                height: 20,
                style: { flexShrink: 0 },
              }),
            ),
          ),
        ),
      );

    const autoMin = await buildColumn();
    const autoLayouts = autoMin.report.slides[0]?.layouts ?? [];
    assert.equal(
      autoLayouts.find(({ testId }) => testId === "baseline-row")?.box.height,
      83.2,
    );
    assert.equal(
      autoMin.report.slides[0]?.objects.find(
        ({ testId }) => testId === "following-block",
      )?.box?.y,
      83.2,
    );
    const autoFile = await validatePptx(autoMin.pptx);
    assert.deepEqual(autoFile.errors, []);
    assert.deepEqual(
      autoFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [23.2, 0, 83.2],
    );
    assert.equal(
      autoFile.slides[0]?.objects.find(({ name }) => name === "following-block")
        ?.boxEmu.y,
      83.2 * 9525,
    );

    const released = await buildColumn(0);
    const releasedLayouts = released.report.slides[0]?.layouts ?? [];
    assert.equal(
      releasedLayouts.find(({ testId }) => testId === "baseline-row")?.box
        .height,
      50,
    );
    assert.deepEqual(
      released.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "small-text", y: 23.2 },
        { testId: "large-text", y: 0 },
        { testId: "following-block", y: 50 },
      ],
    );
    const releasedFile = await validatePptx(released.pptx);
    assert.deepEqual(releasedFile.errors, []);
    assert.deepEqual(
      releasedFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [23.2, 0, 50],
    );

    const nested = await build(
      createElement(
        Deck,
        { width: 320, height: 120 },
        createElement(
          Slide,
          { slideId: "nested-text-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                alignItems: "baseline",
                width: 300,
                height: 100,
              },
            },
            createElement(
              Box,
              {
                testId: "small-wrapper",
                style: { display: "flex", width: 100, height: 60 },
              },
              baselineText("wrapped-small-text", 20),
            ),
            createElement(
              Box,
              {
                testId: "large-wrapper",
                style: { display: "flex", width: 100, height: 60 },
              },
              baselineText("wrapped-large-text", 40),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual(
      nested.report.slides[0]?.layouts
        ?.filter(
          ({ testId }) =>
            testId === "small-wrapper" || testId === "large-wrapper",
        )
        .map(({ testId, box }) => ({ testId, y: box.y })),
      [
        { testId: "small-wrapper", y: 23.2 },
        { testId: "large-wrapper", y: 0 },
      ],
    );
    assert.deepEqual(
      nested.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "wrapped-small-text", y: 23.2 },
        { testId: "wrapped-large-text", y: 0 },
      ],
    );
    const nestedFile = await validatePptx(nested.pptx);
    assert.deepEqual(nestedFile.errors, []);
    assert.equal(nested.report.slides[0]?.objects.length, 2);
    assert.equal(nested.report.slides[0]?.layouts?.length, 3);
    assert.deepEqual(
      nestedFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [23.2, 0],
    );

    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 120 },
          createElement(
            Slide,
            { slideId: "complex-text-baseline-wrapper" },
            createElement(
              Box,
              {
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                  height: 100,
                },
              },
              createElement(
                Box,
                {
                  style: { display: "flex", width: 200, height: 60 },
                },
                baselineText("wrapped-first", 20),
                baselineText("wrapped-second", 40),
              ),
            ),
          ),
        ),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "unsupported-baseline-context",
        ),
    );
  });

  it("preserves flex stretch and content-sized height for non-baseline text", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 120 },
        createElement(
          Slide,
          { slideId: "text-stretch-without-baseline" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                alignItems: "stretch",
                width: 300,
                height: 100,
              },
            },
            createElement(
              Text,
              {
                testId: "fixed-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
            createElement(
              Text,
              {
                testId: "content-sized-text",
                width: 100,
                height: "auto",
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
        height: box?.height,
      })),
      [
        { testId: "fixed-text", y: 0, height: 60 },
        { testId: "content-sized-text", y: 0, height: 100 },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        y: boxEmu.y / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { y: 0, height: 60 },
        { y: 0, height: 100 },
      ],
    );
  });

  it("uses child baselines of text-less flex containers for baseline automatic minimum height", async () => {
    const makeText = () =>
      createElement(
        Text,
        {
          testId: "measured-text",
          width: 100,
          height: 60,
          style: { fontFamily: cjkFont, fontSize: 20 },
        },
        "中",
      );
    const makeRow = () =>
      createElement(
        Box,
        {
          testId: "baseline-row",
          style: { display: "flex", alignItems: "baseline", width: 300 },
        },
        makeText(),
        createElement(
          Box,
          {
            testId: "shape-container",
            style: {
              display: "flex",
              alignItems: "flex-start",
              width: 100,
              height: 100,
            },
          },
          createElement(Shape, {
            testId: "inner-shape",
            width: 100,
            height: 20,
          }),
        ),
      );
    const makeFollowing = () =>
      createElement(Shape, {
        testId: "following-shape",
        width: 100,
        height: 20,
        style: { flexShrink: 0 },
      });

    const natural = await build(
      createElement(
        Deck,
        { width: 400, height: 400 },
        createElement(
          Slide,
          { slideId: "mixed-baseline-natural" },
          createElement(
            Box,
            { style: { width: 300 } },
            makeRow(),
            makeFollowing(),
          ),
        ),
      ),
    );
    assert.equal(
      natural.report.slides[0]?.layouts?.find(
        ({ testId }) => testId === "baseline-row",
      )?.box.height,
      103.2,
    );
    assert.deepEqual(
      natural.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "measured-text", y: 0 },
        { testId: "inner-shape", y: 3.2 },
        { testId: "following-shape", y: 103.2 },
      ],
    );
    const naturalFile = await validatePptx(natural.pptx);
    assert.deepEqual(naturalFile.errors, []);
    assert.deepEqual(
      naturalFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [0, 3.2, 103.2],
    );

    const automatic = await build(
      createElement(
        Deck,
        { width: 400, height: 400 },
        createElement(
          Slide,
          { slideId: "mixed-baseline-auto-min" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "column",
                width: 300,
                height: 200,
              },
            },
            makeRow(),
            makeFollowing(),
          ),
        ),
      ),
    );
    assert.equal(
      automatic.report.slides[0]?.layouts?.find(
        ({ testId }) => testId === "baseline-row",
      )?.box.height,
      103.2,
    );
    assert.deepEqual(
      automatic.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "measured-text", y: 0 },
        { testId: "inner-shape", y: 3.2 },
        { testId: "following-shape", y: 103.2 },
      ],
    );
    const automaticFile = await validatePptx(automatic.pptx);
    assert.deepEqual(automaticFile.errors, []);
    assert.deepEqual(
      automaticFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [0, 3.2, 103.2],
    );
  });

  it("selects baseline child items adopting Yoga resolved alignment for column", async () => {
    const makeText = () =>
      createElement(
        Text,
        {
          testId: "selection-text",
          width: 100,
          height: 60,
          style: { fontFamily: cjkFont, fontSize: 20 },
        },
        "中",
      );
    const makeRow = () =>
      createElement(
        Box,
        {
          testId: "selection-row",
          style: { display: "flex", alignItems: "baseline", width: 300 },
        },
        makeText(),
        createElement(
          Box,
          {
            testId: "selection-column",
            style: {
              display: "flex",
              flexDirection: "column",
              alignItems: "flex-start",
              width: 100,
              height: 100,
            },
          },
          createElement(Shape, {
            testId: "selection-first-shape",
            width: 100,
            height: 20,
          }),
          createElement(Shape, {
            testId: "selection-second-shape",
            width: 100,
            height: 80,
            style: { alignSelf: "baseline" },
          }),
        ),
      );
    const makeFollowing = () =>
      createElement(Shape, {
        testId: "selection-following-shape",
        width: 100,
        height: 20,
        style: { flexShrink: 0 },
      });
    const expectedObjects = [
      { testId: "selection-text", y: 0 },
      { testId: "selection-first-shape", y: 3.2 },
      { testId: "selection-second-shape", y: 23.2 },
      { testId: "selection-following-shape", y: 103.2 },
    ];

    const natural = await build(
      createElement(
        Deck,
        { width: 400, height: 400 },
        createElement(
          Slide,
          { slideId: "baseline-selection-natural" },
          createElement(
            Box,
            { style: { width: 300 } },
            makeRow(),
            makeFollowing(),
          ),
        ),
      ),
    );
    assert.equal(
      natural.report.slides[0]?.layouts?.find(
        ({ testId }) => testId === "selection-row",
      )?.box.height,
      103.2,
    );
    assert.deepEqual(
      natural.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      expectedObjects,
    );
    const naturalFile = await validatePptx(natural.pptx);
    assert.deepEqual(naturalFile.errors, []);
    assert.deepEqual(
      naturalFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [0, 3.2, 23.2, 103.2],
    );

    const automatic = await build(
      createElement(
        Deck,
        { width: 400, height: 400 },
        createElement(
          Slide,
          { slideId: "baseline-selection-auto-min" },
          createElement(
            Box,
            {
              style: {
                display: "flex",
                flexDirection: "column",
                width: 300,
                height: 200,
              },
            },
            makeRow(),
            makeFollowing(),
          ),
        ),
      ),
    );
    assert.equal(
      automatic.report.slides[0]?.layouts?.find(
        ({ testId }) => testId === "selection-row",
      )?.box.height,
      103.2,
    );
    assert.deepEqual(
      automatic.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      expectedObjects,
    );
    const automaticFile = await validatePptx(automatic.pptx);
    assert.deepEqual(automaticFile.errors, []);
    assert.deepEqual(
      automaticFile.slides[0]?.objects.map(({ boxEmu }) =>
        Number((boxEmu.y / 9525).toFixed(2)),
      ),
      [0, 3.2, 23.2, 103.2],
    );
  });

  it("preserves text main-axis content height in flex column, allowing shrink with minHeight: 0", async () => {
    const columnText = (testId: string, minHeight?: number) =>
      createElement(Text, {
        testId,
        width: 20,
        height: "auto",
        style: {
          fontFamily: cjkFont,
          fontSize: 20,
          ...(minHeight === undefined ? {} : { minHeight }),
        },
        children: "中文中文中文",
      });
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 700 },
        createElement(
          Slide,
          { slideId: "column-min-content" },
          createElement(
            Box,
            {
              testId: "constrained-column",
              style: {
                display: "flex",
                flexDirection: "column",
                width: 100,
                height: 300,
              },
            },
            columnText("column-a"),
            columnText("column-b"),
          ),
          createElement(
            Box,
            {
              testId: "released-column",
              style: {
                display: "flex",
                flexDirection: "column",
                width: 100,
                height: 300,
              },
            },
            columnText("released-column-a", 0),
            columnText("released-column-b", 0),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        box,
      })),
      [
        {
          testId: "column-a",
          box: { x: 0, y: 0, width: 20, height: 173.76 },
        },
        {
          testId: "column-b",
          box: { x: 0, y: 173.76, width: 20, height: 173.76 },
        },
        {
          testId: "released-column-a",
          box: { x: 0, y: 300, width: 20, height: 150 },
        },
        {
          testId: "released-column-b",
          box: { x: 0, y: 450, width: 20, height: 150 },
        },
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "overset",
          path: "deck/slide[column-min-content]/box[released-column]/text[released-column-a]",
        },
        {
          code: "overset",
          path: "deck/slide[column-min-content]/box[released-column]/text[released-column-b]",
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ boxEmu }) => ({
        y: boxEmu.y / 9525,
        height: boxEmu.height / 9525,
      })),
      [
        { y: 0, height: 173.76 },
        { y: 173.76, height: 173.76 },
        { y: 300, height: 150 },
        { y: 450, height: 150 },
      ],
    );
  });

  it("treats only numeric minmax lower bounds as constraints for fr tracks, processing row content by lower-bound type", async () => {
    const result = await build(
      <Deck width={400} height={400}>
        <Slide slideId="fr-minimum-100">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "minmax(100px, 1fr) 1fr",
              gridTemplateRows: "1fr",
            }}
          >
            <Shape testId="lower-100-first" />
            <Shape testId="lower-100-second" />
          </Box>
        </Slide>
        <Slide slideId="fr-minimum-200">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "minmax(200px, 1fr) 1fr",
              gridTemplateRows: "1fr",
            }}
          >
            <Shape testId="lower-200-first" />
            <Shape testId="lower-200-second" />
          </Box>
        </Slide>
        <Slide slideId="fr-row-numeric-minimum">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(20px, 1fr) 1fr",
            }}
          >
            <Shape testId="row-minimum-first" />
            <Shape testId="row-minimum-second" />
          </Box>
        </Slide>
        <Slide slideId="fr-row-zero-minimum-overflow">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(0px, 1fr) 1fr",
            }}
          >
            <Shape testId="row-zero-first" width={20} height={200} />
            <Shape testId="row-zero-second" width={20} height={10} />
          </Box>
        </Slide>
        <Slide slideId="fr-row-auto-content-minimum">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(auto, 1fr) 1fr",
            }}
          >
            <Shape testId="row-auto-first" width={20} height={200} />
            <Shape testId="row-auto-second" width={20} height={10} />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides.map((slide) => slide?.objects.map(({ box }) => box)),
      [
        [
          { x: 0, y: 0, width: 150, height: 100 },
          { x: 150, y: 0, width: 150, height: 100 },
        ],
        [
          { x: 0, y: 0, width: 200, height: 100 },
          { x: 200, y: 0, width: 100, height: 100 },
        ],
        [
          { x: 0, y: 0, width: 100, height: 50 },
          { x: 0, y: 50, width: 100, height: 50 },
        ],
        [
          { x: 0, y: 0, width: 20, height: 200 },
          { x: 0, y: 50, width: 20, height: 10 },
        ],
        [
          { x: 0, y: 0, width: 20, height: 200 },
          { x: 0, y: 200, width: 20, height: 10 },
        ],
      ],
    );
    assert.deepEqual(
      result.report.diagnostics.map(({ code }) => code),
      ["grid-item-overflow", "grid-overflow", "grid-item-overflow"],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) =>
        slide?.objects.map(({ boxEmu }) => [
          boxEmu.x,
          boxEmu.y,
          boxEmu.width,
          boxEmu.height,
        ]),
      ),
      [
        [
          [0, 0, 1_428_750, 952_500],
          [1_428_750, 0, 1_428_750, 952_500],
        ],
        [
          [0, 0, 1_905_000, 952_500],
          [1_905_000, 0, 952_500, 952_500],
        ],
        [
          [0, 0, 952_500, 476_250],
          [0, 476_250, 952_500, 476_250],
        ],
        [
          [0, 0, 190_500, 1_905_000],
          [0, 476_250, 190_500, 95_250],
        ],
        [
          [0, 0, 190_500, 1_905_000],
          [0, 1_905_000, 190_500, 95_250],
        ],
      ],
    );
  });

  it("preserves stretch or absolutely positioned height provided by parent layout in auto-height Grid", async () => {
    const stretchedGrid = (testId: string) => (
      <Box
        testId={testId}
        style={{
          display: "grid",
          width: 100,
          gridTemplateColumns: "100px",
          gridTemplateRows: "1fr 1fr",
        }}
      >
        <Shape testId={`${testId}-first`} width={20} height={20} />
        <Shape testId={`${testId}-second`} width={20} height={20} />
      </Box>
    );
    const result = await build(
      <Deck width={400} height={300}>
        <Slide slideId="grid-in-flex-stretch">
          <Box style={{ display: "flex", width: 300, height: 100 }}>
            {stretchedGrid("flex-inner-grid")}
          </Box>
        </Slide>
        <Slide slideId="grid-in-grid-stretch">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "100px",
              gridTemplateRows: "100px",
            }}
          >
            {stretchedGrid("grid-inner-grid")}
          </Box>
        </Slide>
        <Slide slideId="grid-in-absolute-inset">
          <Box style={{ position: "relative", width: 300, height: 100 }}>
            <Box
              testId="absolute-inner-grid"
              style={{
                display: "grid",
                position: "absolute",
                inset: 0,
                gridTemplateColumns: "1fr",
                gridTemplateRows: "1fr 1fr",
              }}
            >
              <Shape
                testId="absolute-inner-grid-first"
                width={20}
                height={20}
              />
              <Shape
                testId="absolute-inner-grid-second"
                width={20}
                height={20}
              />
            </Box>
          </Box>
        </Slide>
        <Slide slideId="grid-content-auto-height">
          <Box style={{ display: "block", width: 100 }}>
            <Box
              testId="content-inner-grid"
              style={{
                display: "grid",
                width: 100,
                gridTemplateColumns: "100px",
                gridTemplateRows: "auto auto",
              }}
            >
              <Shape testId="content-inner-grid-first" width={20} height={20} />
              <Shape
                testId="content-inner-grid-second"
                width={20}
                height={20}
              />
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides.map((slide, index) => {
        const layoutId = [
          "flex-inner-grid",
          "grid-inner-grid",
          "absolute-inner-grid",
          "content-inner-grid",
        ][index];
        return slide?.layouts?.find(({ testId }) => testId === layoutId)?.box;
      }),
      [
        { x: 0, y: 0, width: 100, height: 100 },
        { x: 0, y: 0, width: 100, height: 100 },
        { x: 0, y: 0, width: 300, height: 100 },
        { x: 0, y: 0, width: 100, height: 40 },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) =>
        slide?.objects.map(({ boxEmu }) => boxEmu.y / 9525),
      ),
      [
        [0, 50],
        [0, 50],
        [0, 50],
        [0, 20],
      ],
    );
  });

  it("evaluates percentage dimensions and relative offsets of Grid items based on their assigned area", async () => {
    const result = await build(
      <Deck width={400} height={300}>
        <Slide slideId="grid-area-percent-size">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "100px 200px",
              gridTemplateRows: "100px",
            }}
          >
            <Shape
              testId="percent-width-in-area"
              height={20}
              style={{ width: "50%", margin: [0, 10] }}
            />
            <Shape
              testId="percent-height-in-area"
              width={20}
              style={{ height: "50%", margin: [10, 0] }}
            />
          </Box>
        </Slide>
        <Slide slideId="grid-area-relative-offset">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "100px 100px 100px",
              gridTemplateRows: "20px 30px 50px",
            }}
          >
            <Shape
              testId="relative-first-cell"
              width={20}
              height={10}
              style={{ position: "relative", left: "10%", top: "10%" }}
            />
            <Shape
              testId="relative-spanning-area"
              width={20}
              height={10}
              style={{
                position: "relative",
                left: "10%",
                top: "10%",
                gridColumn: "2 / span 2",
                gridRow: "2 / span 2",
              }}
            />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides.map((slide) => slide?.objects.map(({ box }) => box)),
      [
        [
          { x: 10, y: 0, width: 50, height: 20 },
          { x: 100, y: 10, width: 20, height: 50 },
        ],
        [
          { x: 10, y: 2, width: 20, height: 10 },
          { x: 120, y: 28, width: 20, height: 10 },
        ],
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) =>
        slide?.objects.map(({ boxEmu }) => [
          boxEmu.x,
          boxEmu.y,
          boxEmu.width,
          boxEmu.height,
        ]),
      ),
      [
        [
          [95_250, 0, 476_250, 190_500],
          [952_500, 95_250, 190_500, 476_250],
        ],
        [
          [95_250, 19_050, 190_500, 95_250],
          [1_143_000, 266_700, 190_500, 95_250],
        ],
      ],
    );
  });

  it("distinguishes upper-bound measurement of natural width from precise layout allocation in Table", async () => {
    const table = (testId: string, style: { readonly width?: number } = {}) => (
      <Table
        testId={testId}
        columns={[80]}
        style={{ fontFamily: cjkFont, fontSize: 12, ...style }}
      >
        <Row style={{ height: 40 }}>
          <Cell>中</Cell>
        </Row>
      </Table>
    );
    const result = await build(
      <Deck width={600} height={400}>
        <Slide slideId="table-natural-control">{table("natural")}</Slide>
        <Slide slideId="table-flex-natural">
          <Box
            style={{
              display: "flex",
              alignItems: "flex-start",
              width: 400,
              height: 100,
              gap: 10,
            }}
          >
            {table("flex-natural")}
            <Shape testId="flex-after-natural" width={100} height={20} />
          </Box>
        </Slide>
        <Slide slideId="table-block-stretch">
          <Box style={{ display: "block", width: 200 }}>
            {table("stretched")}
          </Box>
        </Slide>
        <Slide slideId="table-flex-explicit-width">
          <Box
            style={{
              display: "flex",
              alignItems: "flex-start",
              width: 400,
              height: 100,
              gap: 10,
            }}
          >
            {table("explicit", { width: 120 })}
            <Shape testId="flex-after-explicit" width={100} height={20} />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides.map((slide) =>
        slide?.objects.map(({ testId, box, table }) => ({
          testId,
          box,
          columns: table?.columnWidths,
          rows: table?.rowHeights,
        })),
      ),
      [
        [
          {
            testId: "natural",
            box: { x: 0, y: 0, width: 80, height: 40 },
            columns: [80],
            rows: [40],
          },
        ],
        [
          {
            testId: "flex-natural",
            box: { x: 0, y: 0, width: 80, height: 40 },
            columns: [80],
            rows: [40],
          },
          {
            testId: "flex-after-natural",
            box: { x: 90, y: 0, width: 100, height: 20 },
            columns: undefined,
            rows: undefined,
          },
        ],
        [
          {
            testId: "stretched",
            box: { x: 0, y: 0, width: 200, height: 40 },
            columns: [200],
            rows: [40],
          },
        ],
        [
          {
            testId: "explicit",
            box: { x: 0, y: 0, width: 120, height: 40 },
            columns: [120],
            rows: [40],
          },
          {
            testId: "flex-after-explicit",
            box: { x: 130, y: 0, width: 100, height: 20 },
            columns: undefined,
            rows: undefined,
          },
        ],
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) =>
        slide?.objects.map((object) => ({
          kind: object.kind,
          box: [
            object.boxEmu.x,
            object.boxEmu.y,
            object.boxEmu.width,
            object.boxEmu.height,
          ],
          columns: object.table?.columnWidthsEmu,
          rows: object.table?.rowHeightsEmu,
        })),
      ),
      [
        [
          {
            kind: "table",
            box: [0, 0, 762_000, 381_000],
            columns: [762_000],
            rows: [381_000],
          },
        ],
        [
          {
            kind: "table",
            box: [0, 0, 762_000, 381_000],
            columns: [762_000],
            rows: [381_000],
          },
          {
            kind: "shape",
            box: [857_250, 0, 952_500, 190_500],
            columns: undefined,
            rows: undefined,
          },
        ],
        [
          {
            kind: "table",
            box: [0, 0, 1_905_000, 381_000],
            columns: [1_905_000],
            rows: [381_000],
          },
        ],
        [
          {
            kind: "table",
            box: [0, 0, 1_143_000, 381_000],
            columns: [1_143_000],
            rows: [381_000],
          },
          {
            kind: "shape",
            box: [1_238_250, 0, 952_500, 190_500],
            columns: undefined,
            rows: undefined,
          },
        ],
      ],
    );
  });

  it("defaults Shape.text to middle anchor and provides text baseline only for valid top anchor", async () => {
    const ordinary = await build(
      <Deck width={400} height={200}>
        <Slide slideId="shape-text-effective-anchors">
          <Box style={{ display: "flex", width: 210, height: 100, gap: 10 }}>
            <Shape
              testId="default-middle-shape"
              width={100}
              height={100}
              text="中"
              textStyle={{ fontFamily: cjkFont, fontSize: 20 }}
              style={{ padding: 10 }}
            />
            <Shape
              testId="explicit-top-shape"
              width={100}
              height={100}
              text="中"
              textStyle={{
                fontFamily: cjkFont,
                fontSize: 20,
                verticalAlign: "top",
              }}
              style={{ padding: 10 }}
            />
          </Box>
        </Slide>
      </Deck>,
    );
    assert.deepEqual(ordinary.report.diagnostics, []);
    const ordinaryXml = strFromU8(
      unzipSync(ordinary.pptx)["ppt/slides/slide1.xml"]!,
    );
    const anchorFor = (testId: string) => {
      const id = ordinary.report.slides[0]?.objects.find(
        ({ testId: objectTestId }) => objectTestId === testId,
      )?.id;
      const start = ordinaryXml.indexOf(`id="${id}"`);
      const body = ordinaryXml.indexOf("<a:bodyPr", start);
      const end = ordinaryXml.indexOf("</a:bodyPr>", body);
      return ordinaryXml.slice(body, end);
    };
    assert.match(anchorFor("default-middle-shape"), /anchor="ctr"/u);
    assert.match(anchorFor("explicit-top-shape"), /anchor="t"/u);
    assert.deepEqual((await validatePptx(ordinary.pptx)).errors, []);

    const topBaseline = await build(
      <Deck width={400} height={200}>
        <Slide slideId="shape-text-top-baseline">
          <Box style={{ display: "flex", alignItems: "baseline", width: 300 }}>
            <Text
              testId="top-baseline-text"
              width={100}
              height={100}
              style={{
                fontFamily: cjkFont,
                fontSize: 20,
                verticalAlign: "top",
              }}
            >
              中
            </Text>
            <Shape
              testId="top-baseline-shape"
              width={100}
              height={100}
              text="中"
              textStyle={{
                fontFamily: cjkFont,
                fontSize: 40,
                verticalAlign: "top",
              }}
            />
          </Box>
        </Slide>
      </Deck>,
    );
    assert.deepEqual(
      topBaseline.report.slides[0]?.objects.map(({ testId, box }) => ({
        testId,
        y: box?.y,
      })),
      [
        { testId: "top-baseline-text", y: 23.2 },
        { testId: "top-baseline-shape", y: 0 },
      ],
    );
    const topBaselineFile = await validatePptx(topBaseline.pptx);
    assert.deepEqual(topBaselineFile.errors, []);
    assert.deepEqual(
      topBaselineFile.slides[0]?.objects.map(({ boxEmu }) => boxEmu.y),
      [220_980, 0],
    );

    await assert.rejects(
      build(
        <Deck width={400} height={200}>
          <Slide slideId="shape-text-middle-baseline">
            <Box
              style={{ display: "flex", alignItems: "baseline", width: 300 }}
            >
              <Text
                testId="top-text"
                width={100}
                height={100}
                style={{
                  fontFamily: cjkFont,
                  fontSize: 20,
                  verticalAlign: "top",
                }}
              >
                中
              </Text>
              <Shape
                testId="middle-shape"
                width={100}
                height={100}
                text="中"
                textStyle={{ fontFamily: cjkFont, fontSize: 40 }}
              />
            </Box>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "baseline-measurement-unavailable" &&
            path.endsWith("shape[middle-shape]"),
        ),
    );
  });
});
