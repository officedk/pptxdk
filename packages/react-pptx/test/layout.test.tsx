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

type ReportObject = {
  readonly id: number;
  readonly authorId?: string;
  readonly members?: readonly ReportObject[];
};

function authorIdsByObjectId(
  objects: readonly ReportObject[] | undefined,
  ids = new Map<number, string>(),
): Map<number, string> {
  for (const object of objects ?? []) {
    if (object.authorId !== undefined) ids.set(object.id, object.authorId);
    authorIdsByObjectId(object.members, ids);
  }
  return ids;
}

function nativeByAuthor(
  result: Awaited<ReturnType<typeof build>>,
  file: Awaited<ReturnType<typeof validatePptx>>,
  slideIndex: number,
) {
  const ids = authorIdsByObjectId(result.report.slides[slideIndex]?.objects);
  return new Map(
    (file.slides[slideIndex]?.objects ?? []).flatMap((object) => {
      const authorId = ids.get(object.id);
      return authorId === undefined ? [] : [[authorId, object] as const];
    }),
  );
}

function nativeBoxByAuthor(
  result: Awaited<ReturnType<typeof build>>,
  file: Awaited<ReturnType<typeof validatePptx>>,
  slideIndex: number,
) {
  return new Map(
    [...nativeByAuthor(result, file, slideIndex)].map(
      ([authorId, object]) => [authorId, object.boxEmu] as const,
    ),
  );
}

describe("public layout entry", () => {
  it("uses Slide as a canvas grid container, omits hidden subtrees, and changes only paint order with zIndex", async () => {
    const result = await build(
      <Deck width={320} height={180}>
        <Slide
          id="slide-grid"
          style={{
            display: "grid",
            gridTemplateColumns: "2fr 1fr",
            gridTemplateRows: "100px",
          }}
        >
          <Shape
            id="front"
            name="front"
            height={30}
            enter={fade()}
            style={{ aspectRatio: 2, zIndex: 1 }}
          />
          <Shape
            id="back"
            name="back"
            width={100}
            height={40}
            enter={fade()}
            style={{ gridColumn: 2, gridRow: 1, zIndex: -1 }}
          />
          <Shape
            id="hidden"
            style={{ display: "none", width: 500, height: 500 }}
          />
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ id, authorId, box }) => ({
        id,
        authorId,
        box,
      })),
      [
        {
          id: 3,
          authorId: "back",
          box: { x: 213.33, y: 0, width: 100, height: 40 },
        },
        {
          id: 2,
          authorId: "front",
          box: { x: 0, y: 0, width: 60, height: 30 },
        },
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
          readonly target: { readonly authorId?: string };
        }[];
      }[];
    };
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target }) => target.authorId),
      ["front", "back"],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("uses respective native sizing models for Table and Chart inside Grid", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="grid-native-layout">
          <Box
            id="grid-native"
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
              id="grid-table"
              columns={["1fr"]}
              style={{ fontFamily: cjkFont, fontSize: 12, gridColumn: 1 }}
            >
              <Row>
                <Cell>Table</Cell>
              </Row>
            </Table>
            <Chart
              id="grid-chart"
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
      objects.map(({ kind, authorId, box }) => ({ kind, authorId, box })),
      [
        {
          kind: "table",
          authorId: "grid-table",
          box: { x: 0, y: 0, width: 145, height: 180 },
        },
        {
          kind: "chart",
          authorId: "grid-chart",
          box: { x: 155, y: 0, width: 145, height: 180 },
        },
      ],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("starts Grid tracks after container padding", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="grid-padding">
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
            <Shape id="inside-padding" />
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

  it("offsets auto-placed grid items by the flow grid origin", async () => {
    const result = await build(
      <Deck width={1280} height={720}>
        <Slide id="grid-origin">
          <Box
            id="frame"
            style={{
              display: "flex",
              flexDirection: "row",
              width: 1280,
              height: 720,
              padding: 48,
              gap: 24,
            }}
          >
            <Box
              id="sidebar"
              style={{ width: 300, height: 624, backgroundColor: "#172033" }}
            />
            <Box
              id="grid"
              style={{
                display: "grid",
                width: 820,
                height: 512,
                gridTemplateColumns: "1fr 1fr",
                gridAutoRows: "1fr",
                gap: 18,
                backgroundColor: "#EFF3F8",
              }}
            >
              <Box
                id="card-a"
                style={{ backgroundColor: "#FFFFFF", padding: 18 }}
              >
                <Text
                  id="card-a-text"
                  style={{ fontFamily: cjkFont, fontSize: 20 }}
                >
                  A
                </Text>
              </Box>
              <Box
                id="card-b"
                style={{ backgroundColor: "#FFFFFF", padding: 18 }}
              >
                <Text
                  id="card-b-text"
                  style={{ fontFamily: cjkFont, fontSize: 20 }}
                >
                  B
                </Text>
              </Box>
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileBox = (name: string) => {
      const id = result.report.slides[0]?.objects.find(
        (candidate) => candidate.authorId === name,
      )?.id;
      const object = file.slides[0]?.objects.find(
        (candidate) => candidate.id === id,
      );
      assert.ok(object, `missing object ${name}`);
      return {
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      };
    };

    assert.deepEqual(fileBox("grid"), {
      x: 372,
      y: 48,
      width: 820,
      height: 512,
    });
    assert.deepEqual(fileBox("card-a"), {
      x: 372,
      y: 48,
      width: 401,
      height: 512,
    });
    assert.deepEqual(fileBox("card-b"), {
      x: 791,
      y: 48,
      width: 401,
      height: 512,
    });
    assert.deepEqual(
      { x: fileBox("card-a-text").x, y: fileBox("card-a-text").y },
      { x: 390, y: 66 },
    );
    assert.deepEqual(
      { x: fileBox("card-b-text").x, y: fileBox("card-b-text").y },
      { x: 809, y: 66 },
    );
  });

  it("places nested grid tracks from a nonzero flow grid origin", async () => {
    const result = await build(
      <Deck width={1280} height={720}>
        <Slide id="grid-nested-origin">
          <Box
            id="frame"
            style={{
              display: "flex",
              flexDirection: "row",
              width: 1280,
              height: 720,
              padding: 48,
              gap: 24,
            }}
          >
            <Box
              id="sidebar"
              style={{ width: 300, height: 624, backgroundColor: "#172033" }}
            />
            <Box
              id="grid"
              style={{
                display: "grid",
                width: 820,
                height: 512,
                padding: 20,
                gap: 18,
                gridTemplateColumns: "1fr 1fr",
                gridTemplateRows: "100px",
                gridAutoRows: "auto",
                backgroundColor: "#EFF3F8",
              }}
            >
              <Box
                id="spanning-card"
                style={{
                  gridColumn: "span 2",
                  backgroundColor: "#FFFFFF",
                }}
              />
              <Box
                id="nested-grid"
                style={{
                  display: "grid",
                  gridColumn: 1,
                  gridRow: 2,
                  width: 200,
                  height: 150,
                  padding: 10,
                  gap: 10,
                  gridTemplateColumns: "1fr 1fr",
                  gridTemplateRows: "1fr",
                  backgroundColor: "#DDE5EF",
                }}
              >
                <Box
                  id="nested-cell-a"
                  style={{ backgroundColor: "#FFFFFF" }}
                />
                <Box
                  id="nested-cell-b"
                  style={{ backgroundColor: "#FFFFFF" }}
                />
              </Box>
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objectBox = (name: string) => {
      const id = result.report.slides[0]?.objects.find(
        (candidate) => candidate.authorId === name,
      )?.id;
      const object = file.slides[0]?.objects.find(
        (candidate) => candidate.id === id,
      );
      assert.ok(object, `missing object ${name}`);
      return {
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      };
    };

    assert.deepEqual(objectBox("spanning-card"), {
      x: 392,
      y: 68,
      width: 780,
      height: 100,
    });
    assert.deepEqual(objectBox("nested-grid"), {
      x: 392,
      y: 186,
      width: 200,
      height: 150,
    });
    assert.deepEqual(objectBox("nested-cell-a"), {
      x: 402,
      y: 196,
      width: 85,
      height: 130,
    });
    assert.deepEqual(objectBox("nested-cell-b"), {
      x: 497,
      y: 196,
      width: 85,
      height: 130,
    });
  });

  it("keeps author absolute descendants anchored to the nearest positioned ancestor", async () => {
    const result = await build(
      <Deck width={640} height={480}>
        <Slide id="grid-absolute-origin">
          <Box
            id="anchor"
            style={{
              position: "relative",
              width: 400,
              height: 300,
              padding: 40,
              backgroundColor: "#EFF3F8",
            }}
          >
            <Box
              id="static-grid"
              style={{
                display: "grid",
                width: 300,
                height: 200,
                padding: 20,
                gridTemplateColumns: "1fr",
                gridTemplateRows: "1fr",
                backgroundColor: "#DDE5EF",
              }}
            >
              <Box id="static-grid-item" style={{ backgroundColor: "#FFFFFF" }}>
                <Shape
                  id="item-absolute"
                  style={{
                    position: "absolute",
                    left: 12,
                    top: 14,
                    width: 10,
                    height: 10,
                  }}
                />
              </Box>
              <Shape
                id="grid-absolute"
                style={{
                  position: "absolute",
                  left: 30,
                  top: 40,
                  width: 10,
                  height: 10,
                }}
              />
            </Box>
          </Box>
          <Box
            id="relative-anchor"
            style={{
              position: "absolute",
              left: 100,
              top: 50,
              width: 300,
              height: 200,
              backgroundColor: "#EFF3F8",
            }}
          >
            <Box
              id="relative-grid"
              style={{
                position: "relative",
                display: "grid",
                width: 200,
                height: 100,
                gridTemplateColumns: "1fr",
                gridTemplateRows: "1fr",
                backgroundColor: "#DDE5EF",
              }}
            >
              <Shape
                id="relative-grid-absolute"
                style={{
                  position: "absolute",
                  left: 5,
                  top: 6,
                  width: 10,
                  height: 10,
                }}
              />
            </Box>
          </Box>
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objectBox = (name: string) => {
      const id = result.report.slides[0]?.objects.find(
        (candidate) => candidate.authorId === name,
      )?.id;
      const object = file.slides[0]?.objects.find(
        (candidate) => candidate.id === id,
      );
      assert.ok(object, `missing object ${name}`);
      return {
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      };
    };

    assert.deepEqual(objectBox("static-grid"), {
      x: 40,
      y: 40,
      width: 300,
      height: 200,
    });
    assert.deepEqual(objectBox("item-absolute"), {
      x: 12,
      y: 14,
      width: 10,
      height: 10,
    });
    assert.deepEqual(objectBox("grid-absolute"), {
      x: 30,
      y: 40,
      width: 10,
      height: 10,
    });
    assert.deepEqual(objectBox("relative-grid"), {
      x: 100,
      y: 50,
      width: 200,
      height: 100,
    });
    assert.deepEqual(objectBox("relative-grid-absolute"), {
      x: 105,
      y: 56,
      width: 10,
      height: 10,
    });
  });

  it("respects text min-content for auto grid tracks and matches native object geometry", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="grid-min-content">
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
              id="min-content"
              style={{ fontFamily: cjkFont, fontSize: 20 }}
            >
              WWWW
            </Text>
            <Shape id="after-min-content" width={20} height={20} />
          </Box>
        </Slide>
      </Deck>,
    );
    const objects = result.report.slides[0]?.objects ?? [];
    const text = objects.find(
      ({ authorId: testId }) => testId === "min-content",
    );
    const afterText = objects.find(
      ({ authorId: testId }) => testId === "after-min-content",
    );
    assert.ok(text?.box !== undefined && text.box.width > 0);
    assert.equal(afterText?.box?.x, text.box.x + text.box.width + 10);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const nativeText = file.slides[0]?.objects.find(
      ({ id }) => id === text?.id,
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
        <Slide id="grid-absolute-descendant">
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
                  id="absolute-descendant"
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
          <Slide id="absolute-grid-placement">
            <Box style={{ display: "grid", width: 200, height: 100 }}>
              <Shape
                id="misplaced"
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
          <Slide id="grid-baseline">
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
        <Slide id="group-flow">
          <Box style={{ display: "flex", width: 300, height: 100, gap: 10 }}>
            <Group
              id="flow-group"
              style={{
                display: "flex",
                flexDirection: "row",
                gap: 5,
                width: 100,
                height: 60,
              }}
            >
              <Shape id="group-a" width={20} height={20} />
              <Shape id="group-b" width={20} height={20} />
            </Group>
            <Shape id="after-group" width={30} height={30} />
          </Box>
        </Slide>
      </Deck>,
    );
    const group = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "flow-group",
    );
    assert.deepEqual(group?.box, { x: 0, y: 0, width: 100, height: 60 });
    assert.deepEqual(
      group?.members?.map(({ authorId: testId, box }) => ({
        authorId: testId,
        box,
      })),
      [
        { authorId: "group-a", box: { x: 0, y: 0, width: 20, height: 20 } },
        { authorId: "group-b", box: { x: 25, y: 0, width: 20, height: 20 } },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.find(
        ({ authorId: testId }) => testId === "after-group",
      )?.box,
      { x: 110, y: 0, width: 30, height: 30 },
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("writes padding into bodyPr and participates in text height and baseline for Text and Shape.text", async () => {
    const result = await build(
      <Deck width={320} height={160}>
        <Slide id="text-insets">
          <Box style={{ display: "flex", alignItems: "baseline", width: 300 }}>
            <Text
              id="plain-text"
              width={80}
              height="auto"
              style={{ fontFamily: cjkFont, fontSize: 20 }}
            >
              AB
            </Text>
            <Text
              id="padded-text"
              width={80}
              height="auto"
              style={{ fontFamily: cjkFont, fontSize: 20, padding: [10, 4] }}
            >
              AB
            </Text>
            <Shape
              id="padded-shape"
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
    const plain = objects.find(
      ({ authorId: testId }) => testId === "plain-text",
    )?.box;
    const padded = objects.find(
      ({ authorId: testId }) => testId === "padded-text",
    )?.box;
    assert.equal(padded?.height, (plain?.height ?? 0) + 20);
    assert.equal(padded?.y, (plain?.y ?? 0) - 10);
    const shape = objects.find(
      ({ authorId: testId }) => testId === "padded-shape",
    )?.box;
    assert.deepEqual(shape, { x: 160, y: 8, width: 100, height: 50 });
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const bodyPrFor = (testId: string) => {
      const id = objects.find(
        ({ authorId: objectTestId }) => objectTestId === testId,
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
        <Slide id="grid-tracks">
          <Box
            id="grid"
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
            <Shape id="first" height={20} />
            <Shape id="span" height={10} style={{ gridColumn: "span 2" }} />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        { id: "first", box: { x: 0, y: 0, width: 95, height: 20 } },
        { id: "span", box: { x: 0, y: 25, width: 200, height: 10 } },
      ],
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("lays out spanning cards across explicit and implicit rows configured by gridAutoRows", async () => {
    const notes =
      "Service revenue grew in every region this quarter, led by managed services.";
    const cardGrid = (name: string, width: number) => (
      <Box
        id={`${name}-grid`}
        style={{
          display: "grid",
          width,
          height: 300,
          gridTemplateColumns: "1fr 1fr 1fr",
          gridTemplateRows: "48px",
          gridAutoRows: "auto 72px",
          columnGap: 12,
          rowGap: 12,
          backgroundColor: "#ffffff",
        }}
      >
        <Box
          id={`${name}-heading`}
          style={{
            gridColumn: "2 / span 2",
            gridRow: 1,
            backgroundColor: "#2b579a",
          }}
        >
          <Text style={{ fontFamily: cjkFont, fontSize: 18, color: "#ffffff" }}>
            Quarterly review
          </Text>
        </Box>
        <Box
          id={`${name}-feature`}
          style={{
            gridColumn: 1,
            gridRow: "1 / span 3",
            backgroundColor: "#e8f4fd",
          }}
        >
          <Text style={{ fontFamily: cjkFont, fontSize: 14 }}>Lead</Text>
        </Box>
        <Box
          id={`${name}-metric`}
          style={{
            gridColumn: 2,
            gridRow: 2,
            backgroundColor: "#e8f4fd",
          }}
        >
          <Text style={{ fontFamily: cjkFont, fontSize: 14 }}>North 42%</Text>
        </Box>
        <Box
          id={`${name}-notes`}
          style={{
            gridColumn: 3,
            gridRow: 2,
            backgroundColor: "#e8f4fd",
          }}
        >
          <Text
            id={`${name}-notes-text`}
            style={{ fontFamily: cjkFont, fontSize: 14 }}
          >
            {notes}
          </Text>
        </Box>
        <Box
          id={`${name}-summary`}
          style={{
            gridColumn: "2 / span 2",
            gridRow: 3,
            backgroundColor: "#e8f4fd",
          }}
        />
      </Box>
    );

    const result = await build(
      <Deck width={800} height={400}>
        <Slide id="grid-auto-wide">{cardGrid("wide", 600)}</Slide>
        <Slide id="grid-auto-narrow">{cardGrid("narrow", 360)}</Slide>
      </Deck>,
    );

    const layoutMap = (slideIndex: number) =>
      new Map(
        (result.report.slides[slideIndex]?.layouts ?? []).flatMap(
          ({ authorId: testId, box }) =>
            testId === undefined ? [] : [[testId, box] as const],
        ),
      );
    const wide = layoutMap(0);
    const narrow = layoutMap(1);
    const wideRow = wide.get("wide-metric")?.height ?? 0;
    const narrowRow = narrow.get("narrow-metric")?.height ?? 0;

    assert.deepEqual(wide.get("wide-heading"), {
      x: 204,
      y: 0,
      width: 396,
      height: 48,
    });
    assert.deepEqual(wide.get("wide-feature"), {
      x: 0,
      y: 0,
      width: 192,
      height: wideRow + 144,
    });
    assert.deepEqual(wide.get("wide-metric"), {
      x: 204,
      y: 60,
      width: 192,
      height: wideRow,
    });
    assert.deepEqual(wide.get("wide-notes"), {
      x: 408,
      y: 60,
      width: 192,
      height: wideRow,
    });
    assert.deepEqual(wide.get("wide-summary"), {
      x: 204,
      y: 60 + wideRow + 12,
      width: 396,
      height: 72,
    });
    assert.ok(
      wideRow > 40,
      "wrapped card text should make the auto implicit row taller than one line",
    );
    assert.ok(
      narrowRow > wideRow,
      "narrower fr columns should wrap card text into a taller implicit row",
    );
    assert.deepEqual(narrow.get("narrow-feature"), {
      x: 0,
      y: 0,
      width: 112,
      height: narrowRow + 144,
    });
    assert.deepEqual(narrow.get("narrow-summary"), {
      x: 124,
      y: 60 + narrowRow + 12,
      width: 236,
      height: 72,
    });
    assert.deepEqual(result.report.diagnostics, []);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const [index, name] of [
      [0, "wide"],
      [1, "narrow"],
    ] as const) {
      const native = nativeByAuthor(result, file, index);
      for (const card of ["heading", "feature", "metric", "notes", "summary"]) {
        const object = native.get(`${name}-${card}`);
        assert.ok(object !== undefined, `missing native card ${name}-${card}`);
        assert.deepEqual(
          {
            x: object.boxEmu.x / 9525,
            y: object.boxEmu.y / 9525,
            width: object.boxEmu.width / 9525,
            height: object.boxEmu.height / 9525,
          },
          layoutMap(index).get(`${name}-${card}`),
        );
      }
      assert.equal(native.get(`${name}-notes-text`)?.text?.text, notes);
    }
  });

  it("distributes implicit fr rows from gridAutoRows and retains fixed implicit rows across container heights", async () => {
    const board = (name: string, height: number) => (
      <Box
        id={`${name}-board`}
        style={{
          display: "grid",
          width: 320,
          height,
          gridTemplateColumns: "100px 100px 100px",
          gridTemplateRows: "20px",
          gridAutoRows: "minmax(40px, 1fr) 60px",
          columnGap: 10,
          rowGap: 10,
        }}
      >
        <Shape
          id={`${name}-header`}
          style={{ gridColumn: "span 3", gridRow: 1 }}
        />
        <Shape id={`${name}-top`} style={{ gridColumn: "span 3" }} />
        <Shape id={`${name}-middle`} style={{ gridColumn: "span 3" }} />
        <Shape id={`${name}-bottom`} style={{ gridColumn: "span 3" }} />
      </Box>
    );
    const result = await build(
      <Deck width={400} height={600}>
        <Slide id="grid-auto-tall">{board("tall", 300)}</Slide>
        <Slide id="grid-auto-medium">{board("medium", 220)}</Slide>
        <Slide id="grid-auto-constrained">{board("constrained", 140)}</Slide>
      </Deck>,
    );

    const objects = (slideIndex: number) =>
      new Map(
        (result.report.slides[slideIndex]?.objects ?? []).flatMap(
          ({ authorId: testId, box }) =>
            testId === undefined || box === undefined
              ? []
              : [[testId, box] as const],
        ),
      );
    const tall = objects(0);
    assert.deepEqual(tall.get("tall-header"), {
      x: 0,
      y: 0,
      width: 320,
      height: 20,
    });
    assert.deepEqual(tall.get("tall-top"), {
      x: 0,
      y: 30,
      width: 320,
      height: 95,
    });
    assert.deepEqual(tall.get("tall-middle"), {
      x: 0,
      y: 135,
      width: 320,
      height: 60,
    });
    assert.deepEqual(tall.get("tall-bottom"), {
      x: 0,
      y: 205,
      width: 320,
      height: 95,
    });

    const medium = objects(1);
    assert.deepEqual(medium.get("medium-top"), {
      x: 0,
      y: 30,
      width: 320,
      height: 55,
    });
    assert.deepEqual(medium.get("medium-middle"), {
      x: 0,
      y: 95,
      width: 320,
      height: 60,
    });
    assert.deepEqual(medium.get("medium-bottom"), {
      x: 0,
      y: 165,
      width: 320,
      height: 55,
    });

    const constrained = objects(2);
    assert.deepEqual(constrained.get("constrained-top"), {
      x: 0,
      y: 30,
      width: 320,
      height: 40,
    });
    assert.deepEqual(constrained.get("constrained-middle"), {
      x: 0,
      y: 80,
      width: 320,
      height: 60,
    });
    assert.deepEqual(constrained.get("constrained-bottom"), {
      x: 0,
      y: 150,
      width: 320,
      height: 40,
    });
    assert.deepEqual(
      result.report.diagnostics.map(({ code }) => code),
      ["grid-overflow"],
    );
    assert.equal(result.report.slides[2]?.layouts?.[0]?.box.height, 140);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const [index, name] of [
      [0, "tall"],
      [1, "medium"],
      [2, "constrained"],
    ] as const) {
      const native = nativeBoxByAuthor(result, file, index);
      for (const part of ["header", "top", "middle", "bottom"]) {
        const object = native.get(`${name}-${part}`);
        const reported = objects(index).get(`${name}-${part}`);
        assert.ok(object !== undefined && reported !== undefined);
        assert.deepEqual(
          {
            x: object.x / 9525,
            y: object.y / 9525,
            width: object.width / 9525,
            height: object.height / 9525,
          },
          reported,
        );
      }
    }
  });

  it("keeps zero-minimum fr implicit rows on their share when spanning content overflows", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="grid-auto-zero-fr-span">
          <Box
            id="zero-fr-span-grid"
            style={{
              display: "grid",
              width: 200,
              height: 40,
              gridTemplateColumns: "100px 100px",
              gridAutoRows: "minmax(0px, 1fr)",
            }}
          >
            <Shape
              id="zero-fr-spanning"
              height={100}
              style={{ gridColumn: 1, gridRow: "1 / span 2" }}
            />
            <Shape id="zero-fr-second" style={{ gridColumn: 2, gridRow: 2 }} />
          </Box>
        </Slide>
      </Deck>,
    );

    const objects = new Map(
      (result.report.slides[0]?.objects ?? []).flatMap(
        ({ authorId: testId, box }) =>
          testId === undefined || box === undefined
            ? []
            : [[testId, box] as const],
      ),
    );
    assert.deepEqual(objects.get("zero-fr-spanning"), {
      x: 0,
      y: 0,
      width: 100,
      height: 100,
    });
    assert.deepEqual(objects.get("zero-fr-second"), {
      x: 100,
      y: 20,
      width: 100,
      height: 20,
    });
    assert.deepEqual(
      result.report.diagnostics.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "grid-item-overflow",
          path: "deck/slide[grid-auto-zero-fr-span]/box[zero-fr-span-grid]/shape[zero-fr-spanning]",
        },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const native = nativeBoxByAuthor(result, file, 0);
    for (const id of ["zero-fr-spanning", "zero-fr-second"]) {
      const object = native.get(id);
      const reported = objects.get(id);
      assert.ok(object !== undefined && reported !== undefined);
      assert.deepEqual(
        {
          x: object.x / 9525,
          y: object.y / 9525,
          width: object.width / 9525,
          height: object.height / 9525,
        },
        reported,
      );
    }
  });

  it("grows auto and content-minimum fr rows but not fixed rows for spanning content", async () => {
    const result = await build(
      <Deck width={320} height={400}>
        <Slide id="grid-auto-span-growth">
          <Box
            id="growth-grid"
            style={{
              display: "grid",
              width: 200,
              height: 300,
              gridTemplateColumns: "100px 100px",
              gridTemplateRows: "auto",
              gridAutoRows: "auto minmax(auto, 1fr) 30px",
            }}
          >
            <Shape
              id="growth-span"
              height={210}
              style={{ gridColumn: 1, gridRow: "1 / span 4" }}
            />
            <Shape
              id="growth-auto-explicit"
              height={10}
              style={{ gridColumn: 2, gridRow: 1 }}
            />
            <Shape
              id="growth-auto-implicit"
              height={10}
              style={{ gridColumn: 2, gridRow: 2 }}
            />
            <Shape
              id="growth-fr-implicit"
              height={10}
              style={{ gridColumn: 2, gridRow: 3 }}
            />
            <Shape
              id="growth-fixed-implicit"
              height={10}
              style={{ gridColumn: 2, gridRow: 4 }}
            />
          </Box>
        </Slide>
      </Deck>,
    );

    const objects = new Map(
      (result.report.slides[0]?.objects ?? []).flatMap(
        ({ authorId: testId, box }) =>
          testId === undefined || box === undefined
            ? []
            : [[testId, box] as const],
      ),
    );
    assert.deepEqual(objects.get("growth-span"), {
      x: 0,
      y: 0,
      width: 100,
      height: 210,
    });
    assert.deepEqual(objects.get("growth-auto-explicit"), {
      x: 100,
      y: 0,
      width: 100,
      height: 10,
    });
    assert.deepEqual(objects.get("growth-auto-implicit"), {
      x: 100,
      y: 60,
      width: 100,
      height: 10,
    });
    assert.deepEqual(objects.get("growth-fr-implicit"), {
      x: 100,
      y: 120,
      width: 100,
      height: 10,
    });
    assert.deepEqual(objects.get("growth-fixed-implicit"), {
      x: 100,
      y: 270,
      width: 100,
      height: 10,
    });
    assert.deepEqual(result.report.diagnostics, []);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const native = nativeBoxByAuthor(result, file, 0);
    for (const id of [
      "growth-span",
      "growth-auto-explicit",
      "growth-auto-implicit",
      "growth-fr-implicit",
      "growth-fixed-implicit",
    ]) {
      const object = native.get(id);
      const reported = objects.get(id);
      assert.ok(object !== undefined && reported !== undefined);
      assert.deepEqual(
        {
          x: object.x / 9525,
          y: object.y / 9525,
          width: object.width / 9525,
          height: object.height / 9525,
        },
        reported,
      );
    }
  });

  it("retains fixed implicit rows and reports grid and text constraint failures", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="grid-auto-constraints">
          <Box
            id="constraint-grid"
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "300px",
              gridTemplateRows: "20px",
              gridAutoRows: "30px",
            }}
          >
            <Shape id="constraint-header" />
            <Shape
              id="overset-card"
              width={60}
              height={30}
              text="中文中文中文中文"
              textStyle={{
                fontFamily: cjkFont,
                fontSize: 20,
                verticalAlign: "top",
              }}
            />
            <Shape id="tall-card" height={50} />
          </Box>
        </Slide>
        <Slide id="grid-auto-overflow">
          <Box
            id="overflow-grid"
            style={{
              display: "grid",
              width: 300,
              height: 40,
              gridTemplateColumns: "300px",
              gridTemplateRows: "20px",
              gridAutoRows: "30px",
            }}
          >
            <Shape id="overflow-header" />
            <Shape id="overflow-card" height={10} />
            <Shape id="overflow-extra" height={10} />
          </Box>
        </Slide>
      </Deck>,
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.deepEqual(layouts[0]?.box, {
      x: 0,
      y: 0,
      width: 300,
      height: 100,
    });
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "constraint-header",
          box: { x: 0, y: 0, width: 300, height: 20 },
        },
        {
          id: "overset-card",
          box: { x: 0, y: 20, width: 60, height: 30 },
        },
        {
          id: "tall-card",
          box: { x: 0, y: 50, width: 300, height: 50 },
        },
      ],
    );
    assert.deepEqual(
      result.report.slides[1]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "overflow-header",
          box: { x: 0, y: 0, width: 300, height: 20 },
        },
        {
          id: "overflow-card",
          box: { x: 0, y: 20, width: 300, height: 10 },
        },
        {
          id: "overflow-extra",
          box: { x: 0, y: 50, width: 300, height: 10 },
        },
      ],
    );
    const diagnostic = (code: string) =>
      result.report.diagnostics.find((item) => item.code === code);
    assert.deepEqual(
      {
        severity: diagnostic("overset")?.severity,
        path: diagnostic("overset")?.path,
      },
      {
        severity: "warning",
        path: "deck/slide[grid-auto-constraints]/box[constraint-grid]/shape[overset-card]",
      },
    );
    assert.deepEqual(
      {
        severity: diagnostic("grid-item-overflow")?.severity,
        path: diagnostic("grid-item-overflow")?.path,
      },
      {
        severity: "warning",
        path: "deck/slide[grid-auto-constraints]/box[constraint-grid]/shape[tall-card]",
      },
    );
    assert.deepEqual(
      {
        severity: diagnostic("grid-overflow")?.severity,
        path: diagnostic("grid-overflow")?.path,
      },
      {
        severity: "warning",
        path: "deck/slide[grid-auto-overflow]/box[overflow-grid]",
      },
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const [index, ids] of [
      [0, ["constraint-header", "overset-card", "tall-card"]],
      [1, ["overflow-header", "overflow-card", "overflow-extra"]],
    ] as const) {
      const native = nativeBoxByAuthor(result, file, index);
      for (const id of ids) {
        const object = native.get(id);
        const reported = result.report.slides[index]?.objects.find(
          ({ authorId: testId }) => testId === id,
        )?.box;
        assert.ok(object !== undefined && reported !== undefined);
        assert.deepEqual(
          {
            x: object.x / 9525,
            y: object.y / 9525,
            width: object.width / 9525,
            height: object.height / 9525,
          },
          reported,
        );
      }
    }
  });

  it("rejects unsupported gridAutoRows declarations instead of ignoring them", async () => {
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide id="unsupported-grid-auto-row">
            <Box
              style={{ display: "grid", gridAutoRows: "fit-content(100px)" }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-grid-tracks" &&
        error.diagnostics[0]?.message.includes("fit-content"),
    );
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide id="empty-grid-auto-row">
            <Box style={{ display: "grid", gridAutoRows: "" }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-grid-tracks",
    );
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide id="block-grid-auto-row">
            <Box style={{ gridAutoRows: "20px" }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-layout-context",
    );
  });

  it("evaluates relative and absolute positioning by horizontal and vertical percentages of containing block respectively", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="position-percent">
          <Box
            id="anchor"
            style={{ position: "relative", width: 200, height: 100 }}
          >
            <Box style={{ width: 100, height: 40 }}>
              <Shape
                id="absolute"
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
              id="relative"
              width={50}
              height={20}
              style={{ position: "relative", left: "10%", top: "50%" }}
            />
            <Shape
              id="absolute-native-coordinates"
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
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        { id: "absolute", box: { x: 20, y: 20, width: 50, height: 10 } },
        { id: "relative", box: { x: 20, y: 90, width: 50, height: 20 } },
        {
          id: "absolute-native-coordinates",
          box: { x: 70, y: 30, width: 10, height: 10 },
        },
      ],
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.find(
        ({ authorId: testId }) => testId === "absolute-native-coordinates",
      )?.box,
      { x: 70, y: 30, width: 10, height: 10 },
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("uses slide as containing block for percentage positioning of direct Slide children, with diagnostics for static offset", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="page-positioning">
          <Shape
            id="page-absolute"
            style={{
              position: "absolute",
              right: "10%",
              bottom: "10%",
              width: "25%",
              height: 10,
            }}
          />
          <Shape
            id="page-relative"
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
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "page-absolute",
          box: { x: 208, y: 206, width: 80, height: 10 },
        },
        {
          id: "page-relative",
          box: { x: 42, y: 44, width: 20, height: 10 },
        },
      ],
    );
    await assert.rejects(
      build(
        <Deck width={320} height={240}>
          <Slide id="static-offset">
            <Shape
              id="invalid-offset"
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
          { id: "explicit-block-layout" },
          createElement(
            Box,
            {
              id: "block",
              style: { display: "block", width: 100, height: 100 },
            },
            createElement(Shape, { id: "member", height: 20 }),
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
            { id: "invalid-block-flex-style" },
            createElement(Box, {
              id: "block",
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
          { id: "block-layout" },
          createElement(
            Box,
            { id: "column", style: { width: 200, height: 10 } },
            createElement(Shape, {
              id: "first",
              height: 20,
            }),
            createElement(Shape, {
              id: "second",
              height: 10,
            }),
          ),
          createElement(Shape, {
            id: "explicit",
            x: 24,
            y: 32,
            width: 50,
            height: 40,
          }),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "first",
          box: { x: 0, y: 0, width: 200, height: 20 },
        },
        {
          id: "second",
          box: { x: 0, y: 20, width: 200, height: 10 },
        },
        {
          id: "explicit",
          box: { x: 24, y: 32, width: 50, height: 40 },
        },
      ],
    );
    assert.equal(result.report.slides[0]?.layouts?.[0]?.authorId, "column");
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
          { id: "auto-layout-size" },
          createElement(
            Box,
            { style: { width: 100, height: 100 } },
            createElement(Shape, {
              id: "auto-text-shape",
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
              id: alignContent === undefined ? "wrap-stretch" : "wrap-start",
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
              createElement(Shape, { id: "first", width: 60, height: 20 }),
              createElement(Shape, { id: "second", width: 60, height: 20 }),
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
        id: testId,
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
          { id: "flex-min-content" },
          createElement(
            Box,
            {
              id: "constrained",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("constrained-a"),
            makeText("constrained-b"),
          ),
          createElement(
            Box,
            {
              id: "released",
              style: { display: "flex", width: 30, height: 200 },
            },
            makeText("released-a", 0),
            makeText("released-b", 0),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "constrained-a",
          box: { x: 0, y: 0, width: 20, height: 200 },
        },
        {
          id: "constrained-b",
          box: { x: 20, y: 0, width: 20, height: 200 },
        },
        {
          id: "released-a",
          box: { x: 0, y: 200, width: 15, height: 200 },
        },
        {
          id: "released-b",
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
          { id: "nested-minimums" },
          createElement(
            Box,
            {
              id: "fixed-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { id: "fixed-content" },
              createElement(Shape, {
                id: "fixed-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              id: "min-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { id: "min-content" },
              createElement(Shape, {
                id: "min-shape",
                width: 20,
                height: 20,
                style: { minWidth: 150 },
              }),
            ),
          ),
          createElement(
            Box,
            {
              id: "released-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { id: "released-content", style: { minWidth: 0 } },
              createElement(Shape, {
                id: "released-shape",
                width: 200,
                height: 20,
              }),
            ),
          ),
          createElement(
            Box,
            {
              id: "margin-parent",
              style: { display: "flex", width: 100, height: 60 },
            },
            createElement(
              Box,
              { id: "margin-content" },
              createElement(Shape, {
                id: "margin-shape",
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
      layouts.find((layout) => layout.authorId === testId)?.box.width;
    assert.equal(layoutWidth("fixed-content"), 200);
    assert.equal(layoutWidth("min-content"), 150);
    assert.equal(layoutWidth("released-content"), 100);
    assert.equal(layoutWidth("margin-content"), 125);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        x: box?.x,
        width: box?.width,
      })),
      [
        { id: "fixed-shape", x: 0, width: 200 },
        { id: "min-shape", x: 0, width: 150 },
        { id: "released-shape", x: 0, width: 200 },
        { id: "margin-shape", x: 10, width: 100 },
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
    const card = (id: string) =>
      createElement(
        Box,
        { id, style: { flex: 1, height: 100 } },
        createElement(Shape, { id: "card-shape", height: 20 }),
      );
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { id: "reused-layout" },
          createElement(
            Box,
            {
              id: "row",
              style: { display: "flex", width: 300, height: 120 },
            },
            card("card-a"),
            card("card-b"),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "card-shape",
          box: { x: 0, y: 0, width: 150, height: 20 },
        },
        {
          id: "card-shape",
          box: { x: 150, y: 0, width: 150, height: 20 },
        },
      ],
    );
    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.deepEqual(
      layouts.map(({ authorId: testId }) => testId),
      ["row", "card-a", "card-b"],
    );
    const row = layouts.find(({ authorId: testId }) => testId === "row");
    const cards = layouts.filter(({ authorId: testId }) =>
      testId?.startsWith("card-"),
    );
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
          { id: "group-with-layout" },
          createElement(
            Group,
            { id: "native-group" },
            createElement(
              Box,
              {
                id: "inner-box",
                style: { width: 100, height: 100 },
              },
              createElement(Shape, {
                id: "group-member",
                name: "group-member",
                height: 20,
              }),
            ),
          ),
        ),
      ),
    );

    const group = result.report.slides[0]?.objects[0];
    const member = group?.members?.[0];
    assert.equal(group?.kind, "group");
    assert.equal(group?.authorId, "native-group");
    assert.deepEqual(member?.box, {
      x: 0,
      y: 0,
      width: 100,
      height: 20,
    });
    const layout = result.report.slides[0]?.layouts?.[0];
    assert.equal(layout?.authorId, "inner-box");
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
            { id: "invalid-context" },
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
          { id: "ignored-item-style" },
          createElement(
            Box,
            { style: { width: 100 } },
            createElement(Shape, {
              id: "block-item",
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
          { id: "overflow" },
          createElement(
            Box,
            { id: "wide-container", style: { width: 400, height: 50 } },
            createElement(Shape, { id: "wide-shape", height: 10 }),
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
          { id: "aligned-objects" },
          createElement(
            Box,
            {
              id: "toolbar",
              style: {
                display: "flex",
                width: 100,
                height: 100,
                justifyContent: "space-between",
                alignItems: "center",
              },
            },
            createElement(Image, {
              id: "logo",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              id: "action",
              width: 20,
              height: 20,
            }),
          ),
        ),
      ),
    );

    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ kind, authorId, box }) => ({
        kind,
        authorId,
        box,
      })),
      [
        {
          kind: "image",
          authorId: "logo",
          box: { x: 0, y: 40, width: 30, height: 20 },
        },
        {
          kind: "shape",
          authorId: "action",
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
          { id: "text-baseline" },
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
                id: "small-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
            createElement(
              Text,
              {
                id: "large-text",
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
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        y: box?.y,
      })),
      [
        { id: "small-text", y: 23.2 },
        { id: "large-text", y: 0 },
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
            { id: "unmeasured-baseline" },
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
                { id: "unmeasured-text", width: 100, height: 60 },
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
            { id: "wrapped-text-baseline" },
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
                  id: "wrapped-text",
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

  it("rejects textless shapes in wrapping baseline container and resolves column-direction baseline", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 240 },
          createElement(
            Slide,
            { id: "synthetic-wrap-baseline" },
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
                id: "wrap-first",
                width: 60,
                height: 20,
              }),
              createElement(Shape, {
                id: "wrap-second",
                width: 60,
                height: 20,
              }),
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

    const column = await build(
      createElement(
        Deck,
        { width: 320, height: 240 },
        createElement(
          Slide,
          { id: "synthetic-column-baseline" },
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
                id: "column-small-text",
                width: 100,
                height: 40,
                style: { fontFamily: "Aptos", fontSize: 20 },
              },
              "中",
            ),
            createElement(Image, {
              id: "column-image",
              src: wideImage,
              width: 30,
              height: 20,
            }),
            createElement(Shape, {
              id: "column-shape",
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
          { id: "natural-baseline-height" },
          createElement(
            Box,
            { style: { width: 300 } },
            createElement(
              Box,
              {
                id: "baseline-row",
                style: {
                  display: "flex",
                  alignItems: "baseline",
                  width: 300,
                },
              },
              createElement(
                Text,
                {
                  id: "small-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 20 },
                },
                "中",
              ),
              createElement(
                Text,
                {
                  id: "large-text",
                  width: 100,
                  height: 60,
                  style: { fontFamily: cjkFont, fontSize: 40 },
                },
                "中",
              ),
            ),
            createElement(Shape, {
              id: "following-block",
              width: 100,
              height: 20,
            }),
          ),
        ),
      ),
    );

    const layouts = result.report.slides[0]?.layouts ?? [];
    assert.equal(
      layouts.find(({ authorId: testId }) => testId === "baseline-row")?.box
        .height,
      83.2,
    );
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        y: box?.y,
      })),
      [
        { id: "small-text", y: 23.2 },
        { id: "large-text", y: 0 },
        { id: "following-block", y: 83.2 },
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
          id: testId,
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
              id: minHeight === 0 ? "baseline-min-zero" : "baseline-auto-min",
            },
            createElement(
              Box,
              {
                id: "column",
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
                  id: "baseline-row",
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
                id: "following-block",
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
      autoLayouts.find(({ authorId: testId }) => testId === "baseline-row")?.box
        .height,
      83.2,
    );
    assert.equal(
      autoMin.report.slides[0]?.objects.find(
        ({ authorId: testId }) => testId === "following-block",
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
    const followingId = autoMin.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "following-block",
    )?.id;
    assert.equal(
      autoFile.slides[0]?.objects.find(({ id }) => id === followingId)?.boxEmu
        .y,
      83.2 * 9525,
    );

    const released = await buildColumn(0);
    const releasedLayouts = released.report.slides[0]?.layouts ?? [];
    assert.equal(
      releasedLayouts.find(({ authorId: testId }) => testId === "baseline-row")
        ?.box.height,
      50,
    );
    assert.deepEqual(
      released.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        y: box?.y,
      })),
      [
        { id: "small-text", y: 23.2 },
        { id: "large-text", y: 0 },
        { id: "following-block", y: 50 },
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
          { id: "nested-text-baseline" },
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
                id: "small-wrapper",
                style: { display: "flex", width: 100, height: 60 },
              },
              baselineText("wrapped-small-text", 20),
            ),
            createElement(
              Box,
              {
                id: "large-wrapper",
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
          ({ authorId: testId }) =>
            testId === "small-wrapper" || testId === "large-wrapper",
        )
        .map(({ authorId: testId, box }) => ({ authorId: testId, y: box.y })),
      [
        { authorId: "small-wrapper", y: 23.2 },
        { authorId: "large-wrapper", y: 0 },
      ],
    );
    assert.deepEqual(
      nested.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        y: box?.y,
      })),
      [
        { id: "wrapped-small-text", y: 23.2 },
        { id: "wrapped-large-text", y: 0 },
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
            { id: "complex-text-baseline-wrapper" },
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
          { id: "text-stretch-without-baseline" },
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
                id: "fixed-text",
                width: 100,
                height: 60,
                style: { fontFamily: cjkFont, fontSize: 20 },
              },
              "中",
            ),
            createElement(
              Text,
              {
                id: "content-sized-text",
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
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        y: box?.y,
        height: box?.height,
      })),
      [
        { id: "fixed-text", y: 0, height: 60 },
        { id: "content-sized-text", y: 0, height: 100 },
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

  it("preserves text main-axis content height in flex column, allowing shrink with minHeight: 0", async () => {
    const columnText = (testId: string, minHeight?: number) =>
      createElement(Text, {
        id: testId,
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
          { id: "column-min-content" },
          createElement(
            Box,
            {
              id: "constrained-column",
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
              id: "released-column",
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
      result.report.slides[0]?.objects.map(({ authorId: testId, box }) => ({
        id: testId,
        box,
      })),
      [
        {
          id: "column-a",
          box: { x: 0, y: 0, width: 20, height: 173.76 },
        },
        {
          id: "column-b",
          box: { x: 0, y: 173.76, width: 20, height: 173.76 },
        },
        {
          id: "released-column-a",
          box: { x: 0, y: 300, width: 20, height: 150 },
        },
        {
          id: "released-column-b",
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
        <Slide id="fr-minimum-100">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "minmax(100px, 1fr) 1fr",
              gridTemplateRows: "1fr",
            }}
          >
            <Shape id="lower-100-first" />
            <Shape id="lower-100-second" />
          </Box>
        </Slide>
        <Slide id="fr-minimum-200">
          <Box
            style={{
              display: "grid",
              width: 300,
              height: 100,
              gridTemplateColumns: "minmax(200px, 1fr) 1fr",
              gridTemplateRows: "1fr",
            }}
          >
            <Shape id="lower-200-first" />
            <Shape id="lower-200-second" />
          </Box>
        </Slide>
        <Slide id="fr-row-numeric-minimum">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(20px, 1fr) 1fr",
            }}
          >
            <Shape id="row-minimum-first" />
            <Shape id="row-minimum-second" />
          </Box>
        </Slide>
        <Slide id="fr-row-zero-minimum-overflow">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(0px, 1fr) 1fr",
            }}
          >
            <Shape id="row-zero-first" width={20} height={200} />
            <Shape id="row-zero-second" width={20} height={10} />
          </Box>
        </Slide>
        <Slide id="fr-row-auto-content-minimum">
          <Box
            style={{
              display: "grid",
              width: 100,
              height: 100,
              gridTemplateColumns: "1fr",
              gridTemplateRows: "minmax(auto, 1fr) 1fr",
            }}
          >
            <Shape id="row-auto-first" width={20} height={200} />
            <Shape id="row-auto-second" width={20} height={10} />
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
        id={testId}
        style={{
          display: "grid",
          width: 100,
          gridTemplateColumns: "100px",
          gridTemplateRows: "1fr 1fr",
        }}
      >
        <Shape id={`${testId}-first`} width={20} height={20} />
        <Shape id={`${testId}-second`} width={20} height={20} />
      </Box>
    );
    const result = await build(
      <Deck width={400} height={300}>
        <Slide id="grid-in-flex-stretch">
          <Box style={{ display: "flex", width: 300, height: 100 }}>
            {stretchedGrid("flex-inner-grid")}
          </Box>
        </Slide>
        <Slide id="grid-in-grid-stretch">
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
        <Slide id="grid-in-absolute-inset">
          <Box style={{ position: "relative", width: 300, height: 100 }}>
            <Box
              id="absolute-inner-grid"
              style={{
                display: "grid",
                position: "absolute",
                inset: 0,
                gridTemplateColumns: "1fr",
                gridTemplateRows: "1fr 1fr",
              }}
            >
              <Shape id="absolute-inner-grid-first" width={20} height={20} />
              <Shape id="absolute-inner-grid-second" width={20} height={20} />
            </Box>
          </Box>
        </Slide>
        <Slide id="grid-content-auto-height">
          <Box style={{ display: "block", width: 100 }}>
            <Box
              id="content-inner-grid"
              style={{
                display: "grid",
                width: 100,
                gridTemplateColumns: "100px",
                gridTemplateRows: "auto auto",
              }}
            >
              <Shape id="content-inner-grid-first" width={20} height={20} />
              <Shape id="content-inner-grid-second" width={20} height={20} />
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
        return slide?.layouts?.find(
          ({ authorId: testId }) => testId === layoutId,
        )?.box;
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
        <Slide id="grid-area-percent-size">
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
              id="percent-width-in-area"
              height={20}
              style={{ width: "50%", margin: [0, 10] }}
            />
            <Shape
              id="percent-height-in-area"
              width={20}
              style={{ height: "50%", margin: [10, 0] }}
            />
          </Box>
        </Slide>
        <Slide id="grid-area-relative-offset">
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
              id="relative-first-cell"
              width={20}
              height={10}
              style={{ position: "relative", left: "10%", top: "10%" }}
            />
            <Shape
              id="relative-spanning-area"
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
        id={testId}
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
        <Slide id="table-natural-control">{table("natural")}</Slide>
        <Slide id="table-flex-natural">
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
            <Shape id="flex-after-natural" width={100} height={20} />
          </Box>
        </Slide>
        <Slide id="table-block-stretch">
          <Box style={{ display: "block", width: 200 }}>
            {table("stretched")}
          </Box>
        </Slide>
        <Slide id="table-flex-explicit-width">
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
            <Shape id="flex-after-explicit" width={100} height={20} />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(
      result.report.slides.map((slide) =>
        slide?.objects.map(({ authorId: testId, box, table }) => ({
          id: testId,
          box,
          columns: table?.columnWidths,
          rows: table?.rowHeights,
        })),
      ),
      [
        [
          {
            id: "natural",
            box: { x: 0, y: 0, width: 80, height: 40 },
            columns: [80],
            rows: [40],
          },
        ],
        [
          {
            id: "flex-natural",
            box: { x: 0, y: 0, width: 80, height: 40 },
            columns: [80],
            rows: [40],
          },
          {
            id: "flex-after-natural",
            box: { x: 90, y: 0, width: 100, height: 20 },
            columns: undefined,
            rows: undefined,
          },
        ],
        [
          {
            id: "stretched",
            box: { x: 0, y: 0, width: 200, height: 40 },
            columns: [200],
            rows: [40],
          },
        ],
        [
          {
            id: "explicit",
            box: { x: 0, y: 0, width: 120, height: 40 },
            columns: [120],
            rows: [40],
          },
          {
            id: "flex-after-explicit",
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
        <Slide id="shape-text-effective-anchors">
          <Box style={{ display: "flex", width: 210, height: 100, gap: 10 }}>
            <Shape
              id="default-middle-shape"
              width={100}
              height={100}
              text="中"
              textStyle={{ fontFamily: cjkFont, fontSize: 20 }}
              style={{ padding: 10 }}
            />
            <Shape
              id="explicit-top-shape"
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
        ({ authorId: objectTestId }) => objectTestId === testId,
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
        <Slide id="shape-text-top-baseline">
          <Box style={{ display: "flex", alignItems: "baseline", width: 300 }}>
            <Text
              id="top-baseline-text"
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
              id="top-baseline-shape"
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
      topBaseline.report.slides[0]?.objects.map(
        ({ authorId: testId, box }) => ({
          id: testId,
          y: box?.y,
        }),
      ),
      [
        { id: "top-baseline-text", y: 23.2 },
        { id: "top-baseline-shape", y: 0 },
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
          <Slide id="shape-text-middle-baseline">
            <Box
              style={{ display: "flex", alignItems: "baseline", width: 300 }}
            >
              <Text
                id="top-text"
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
                id="middle-shape"
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
