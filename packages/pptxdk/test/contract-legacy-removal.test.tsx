import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import * as pptxdkMain from "@officedk/pptxdk";
import { BuildError, build, font, resolve } from "@officedk/pptxdk";
import * as pptxdkBuild from "@officedk/pptxdk/build";
import * as pptxdkNode from "@officedk/pptxdk/node";
import * as pptxdkPreview from "@officedk/pptxdk/preview";
import { preview } from "@officedk/pptxdk/preview";
import * as pptxdkReact from "@officedk/pptxdk/react";
import {
  Box,
  Cell,
  Chart,
  Deck,
  Group,
  Image,
  Placeholder,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
  defineSlideLayout,
  expandDocument,
} from "@officedk/pptxdk/react";
import * as pptxdkVanilla from "@officedk/pptxdk/vanilla";
import {
  box,
  cell,
  chart,
  deck,
  group,
  image,
  placeholder,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import type { ReactNode } from "react";
import { describe, it } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
]);

async function expectBuildErrorCode(
  input: ReactNode | Parameters<typeof resolve>[0],
  expectedCode: "unknown-prop" | "invalid-style",
): Promise<void> {
  let caught: unknown;
  try {
    await build(input);
  } catch (error) {
    caught = error;
  }
  assert.ok(
    caught instanceof BuildError,
    `Expected BuildError, got ${String(caught)}`,
  );
  assert.equal(caught.diagnostics[0]?.code, expectedCode);
}

function expectSyncBuildErrorCode(
  fn: () => unknown,
  expectedCode: "unknown-prop" | "invalid-style",
): void {
  let caught: unknown;
  try {
    fn();
  } catch (error) {
    caught = error;
  }
  assert.ok(
    caught instanceof BuildError,
    `Expected BuildError, got ${String(caught)}`,
  );
  assert.equal(caught.diagnostics[0]?.code, expectedCode);
}

describe("contract: legacy surface removal", () => {
  it("does not export removed runtime or schema symbols from any public @officedk/pptxdk subpath", () => {
    const subpaths: readonly (readonly [string, Record<string, unknown>])[] = [
      ["@officedk/pptxdk", pptxdkMain],
      ["@officedk/pptxdk/build", pptxdkBuild],
      ["@officedk/pptxdk/preview", pptxdkPreview],
      ["@officedk/pptxdk/react", pptxdkReact],
      ["@officedk/pptxdk/vanilla", pptxdkVanilla],
      ["@officedk/pptxdk/node", pptxdkNode],
    ];
    const removedSymbols = [
      "observeMotion",
      "diagnosticSchema",
      "diagnosticSeveritySchema",
      "previewSeveritySchema",
      "previewDiagnosticSchema",
      "placeholderSzSchema",
      "groupBoundsSchema",
      "slideLayoutObservationSchema",
      "slideMasterObservationSchema",
    ] as const;

    for (const [subpath, exportsObj] of subpaths) {
      for (const symbol of removedSymbols) {
        assert.equal(
          exportsObj[symbol],
          undefined,
          `${subpath} must not export ${symbol}`,
        );
      }
    }
  });

  it("omits removed properties on successful resolve, build, and preview outputs and locators", async () => {
    const sampleDeck = (
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Shape
            id="sh1"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 180,
              height: 60,
              fontFamily: jost,
              fontSize: 16,
              background: "#E2E8F0",
            }}
          >
            Hello
          </Shape>
          <Table
            id="tbl1"
            columns={[120, 120]}
            style={{ position: "absolute", left: 40, top: 140, width: 240 }}
          >
            <Row style={{ height: 40 }}>
              <Cell style={{ fontFamily: jost, fontSize: 12 }}>A</Cell>
              <Cell style={{ fontFamily: jost, fontSize: 12 }}>B</Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>
    );

    const expanded = await expandDocument(sampleDeck);
    const resolution = resolve(expanded);
    assert.equal("query" in resolution, false);
    assert.equal("diagnostics" in resolution.report, false);

    const buildResult = await build(sampleDeck);
    assert.equal("query" in buildResult, false);
    assert.equal("diagnostics" in buildResult.report, false);

    const previewResult = await preview(sampleDeck);
    assert.equal("query" in previewResult, false);
    const page = previewResult.pages[0];
    assert.ok(page);
    assert.equal("diagnostics" in page, false);
    const previewNode = page.nodes[0];
    assert.ok(previewNode);
    assert.equal("screenBox" in previewNode, false);

    const slideLocator = resolution.slide("s1");
    assert.equal("observation" in slideLocator, false);
    assert.equal("authorId" in slideLocator, false);

    const shapeLocator = slideLocator.get("sh1");
    assert.equal("observe" in shapeLocator, false);
    assert.equal("authorId" in shapeLocator, false);

    const tableObj = slideLocator.objects.find(
      (obj) => obj.authorId === "tbl1",
    )?.table;
    assert.ok(tableObj);
    assert.equal("columnWidths" in tableObj, false);
    assert.equal("rowHeights" in tableObj, false);
    assert.equal("parts" in tableObj, false);
  });

  it("rejects removed top-level geometry props (x, y, width, height) at compile time and runtime across JSX and vanilla", async () => {
    const dummyChartGroup = {
      kind: "column" as const,
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1"],
      },
      series: [{ name: "S1", values: [10] }],
    };

    // Compile-time type rejection checks via @ts-expect-error on explicit literal keys
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error legacy top-level geometry removed from Shape */}
          <Shape x={10} y={20} width={100} height={50} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error legacy top-level geometry removed from shape
          shape({ x: 10, y: 20, width: 100, height: 50 }),
        ),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error legacy top-level geometry removed from Text */}
          <Text x={10} y={20} width={100} height={50}>
            Hello
          </Text>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error legacy top-level geometry removed from text
          text("Hello", { x: 10, y: 20, width: 100, height: 50 }),
        ),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error legacy top-level geometry removed from Box */}
          <Box x={10} y={20} width={100} height={50} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error legacy top-level geometry removed from box
          box({ x: 10, y: 20, width: 100, height: 50 }),
        ),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error legacy top-level geometry removed from Image */}
          <Image src={png} x={10} y={20} width={100} height={50} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error legacy top-level geometry removed from image
          image({ src: png, x: 10, y: 20, width: 100, height: 50 }),
        ),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error legacy top-level geometry removed from Table */}
          <Table columns={[100]} x={10} y={20} width={100}>
            <Row style={{ height: 40 }}>
              <Cell>A</Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error legacy top-level geometry removed from table
          table({ columns: [100], x: 10, y: 20, width: 100 }).append(
            row({ style: { height: 40 } }).append(cell("A")),
          ),
        ),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Chart
            groups={[dummyChartGroup]}
            // @ts-expect-error legacy top-level geometry removed from Chart
            x={10}
            y={20}
            width={200}
            height={120}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          chart({
            groups: [dummyChartGroup],
            // @ts-expect-error legacy top-level geometry removed from chart
            x: 10,
            y: 20,
            width: 200,
            height: 120,
          }),
        ),
      ),
      "unknown-prop",
    );
    const placeholderTestLayout = defineSlideLayout({
      id: "l-ph",
      placeholders: { body: { type: "body" } },
    });
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <SlideMaster id="m1">
          <SlideLayout layout={placeholderTestLayout} />
        </SlideMaster>
        <Slide id="s1" layout={placeholderTestLayout}>
          <Placeholder
            placeholder="body"
            // @ts-expect-error legacy top-level geometry removed from Placeholder
            x={10}
            y={20}
            width={100}
            height={50}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slideMaster({ id: "m1" }).append(slideLayout(placeholderTestLayout)),
        slide({ id: "s1", layout: placeholderTestLayout }).append(
          placeholder({
            placeholder: "body",
            // @ts-expect-error legacy top-level geometry removed from placeholder
            x: 10,
            y: 20,
            width: 100,
            height: 50,
          }),
        ),
      ),
      "unknown-prop",
    );
    expectSyncBuildErrorCode(
      () =>
        defineSlideLayout({
          id: "l-typed",
          placeholders: {
            title: { type: "title", x: 10, y: 20, width: 100, height: 50 },
          },
        }),
      "unknown-prop",
    );
    expectSyncBuildErrorCode(
      () =>
        defineSlideLayout({
          id: "l-footer-typed",
          footerPlaceholders: {
            // @ts-expect-error legacy top-level geometry removed from FooterPlaceholderSpec
            footer: { x: 10, y: 20, width: 100, height: 50 },
          },
        }),
      "unknown-prop",
    );

    // Exhaustive per-prop runtime checks
    for (const prop of ["x", "y", "width", "height"] as const) {
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Shape {...({ [prop]: 10 } as Record<string, unknown>)} />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Text {...({ [prop]: 10 } as Record<string, unknown>)}>Hello</Text>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Box {...({ [prop]: 10 } as Record<string, unknown>)} />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Image src={png} {...({ [prop]: 10 } as Record<string, unknown>)} />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Chart
              groups={[dummyChartGroup]}
              {...({ [prop]: 10 } as Record<string, unknown>)}
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <SlideMaster id="m1">
            <SlideLayout layout={placeholderTestLayout} />
          </SlideMaster>
          <Slide id="s1" layout={placeholderTestLayout}>
            <Placeholder
              placeholder="body"
              {...({ [prop]: 10 } as Record<string, unknown>)}
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      expectSyncBuildErrorCode(
        () =>
          defineSlideLayout({
            id: "l-prop",
            placeholders: {
              title: {
                type: "title",
                ...({ [prop]: 10 } as Record<string, unknown>),
              } as never,
            },
          }),
        "unknown-prop",
      );
      expectSyncBuildErrorCode(
        () =>
          defineSlideLayout({
            id: "l-footer-prop",
            footerPlaceholders: {
              footer: {
                ...({ [prop]: 10 } as Record<string, unknown>),
              } as never,
            },
          }),
        "unknown-prop",
      );
    }

    for (const prop of ["x", "y", "width"] as const) {
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Table
              columns={[100]}
              {...({ [prop]: 10 } as Record<string, unknown>)}
            >
              <Row style={{ height: 40 }}>
                <Cell>A</Cell>
              </Row>
            </Table>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
    }
  });

  it("rejects Group.bounds, Shape.fill/text/textStyle, top-level flipH/flipV, and placeholder sz", async () => {
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error Group.bounds removed */}
          <Group bounds={{ x: 0, y: 0, width: 100, height: 100 }}>
            <Box style={{ width: 20, height: 20 }} />
          </Group>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          // @ts-expect-error group.bounds removed
          group({ bounds: { x: 0, y: 0, width: 100, height: 100 } }).append(
            box({ style: { width: 20, height: 20 } }),
          ),
        ),
      ),
      "unknown-prop",
    );

    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error Shape.fill removed */}
          <Shape fill="#FF0000" style={{ width: 100, height: 50 }} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error Shape.text removed */}
          <Shape text="Legacy" style={{ width: 100, height: 50 }} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Shape
            // @ts-expect-error Shape.textStyle removed
            textStyle={{ fontSize: 14 }}
            style={{ width: 100, height: 50 }}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        // @ts-expect-error shape.fill removed
        slide({ id: "s1" }).append(shape({ fill: "#FF0000" })),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        // @ts-expect-error shape.text removed
        slide({ id: "s1" }).append(shape({ text: "Legacy" })),
      ),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        // @ts-expect-error shape.textStyle removed
        slide({ id: "s1" }).append(shape({ textStyle: { fontSize: 14 } })),
      ),
      "unknown-prop",
    );

    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error top-level flipH/flipV removed on Shape */}
          <Shape flipH flipV />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error top-level flipH/flipV removed on Image */}
          <Image src={png} flipH flipV />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error top-level flipH/flipV removed on Group */}
          <Group flipH flipV>
            <Box style={{ width: 20, height: 20 }} />
          </Group>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    expectSyncBuildErrorCode(
      () =>
        defineSlideLayout({
          id: "l-flip-typed",
          placeholders: {
            title: { type: "title", flipH: true, flipV: true },
          },
        }),
      "unknown-prop",
    );

    for (const flipProp of ["flipH", "flipV"] as const) {
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Shape {...({ [flipProp]: true } as Record<string, unknown>)} />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Image
              src={png}
              {...({ [flipProp]: true } as Record<string, unknown>)}
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Group {...({ [flipProp]: true } as Record<string, unknown>)}>
              <Box style={{ width: 20, height: 20 }} />
            </Group>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      expectSyncBuildErrorCode(
        () =>
          defineSlideLayout({
            id: "l-flip",
            placeholders: {
              title: {
                type: "title",
                ...({ [flipProp]: true } as Record<string, unknown>),
              } as never,
            },
          }),
        "unknown-prop",
      );
    }

    expectSyncBuildErrorCode(
      () =>
        defineSlideLayout({
          id: "l-sz",
          placeholders: {
            title: { type: "title", sz: "full" },
          },
        }),
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          {/* @ts-expect-error nativePlaceholder.sz removed; use size */}
          <Placeholder nativePlaceholder={{ type: "body", sz: "half" }} />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
  });

  it("rejects removed legacy style keys across Shape, Box, Text, Image, Slide, Row, Cell, and placeholder style", async () => {
    // Compile-time type rejection checks for removed style keys
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Shape
            style={{
              width: 100,
              height: 50,
              // @ts-expect-error backgroundColor removed from Shape.style
              backgroundColor: "#FF0000",
            }}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Shape
            style={{
              width: 100,
              height: 50,
              // @ts-expect-error startArrow removed from Shape.style top-level (belongs inside style.border)
              startArrow: "triangle",
            }}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Box
            style={{
              width: 100,
              height: 50,
              // @ts-expect-error borderColor removed from Box.style
              borderColor: "#000000",
            }}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Text
            style={{
              width: 100,
              height: 50,
              // @ts-expect-error borderWidth removed from Text.style
              borderWidth: 2,
            }}
          >
            Hi
          </Text>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Image
            src={png}
            style={{
              width: 100,
              height: 50,
              // @ts-expect-error borderStyle removed from Image.style
              borderStyle: "solid",
            }}
          />
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide
          id="s1"
          style={{
            // @ts-expect-error backgroundColor removed from Slide.style
            backgroundColor: "#FFFFFF",
          }}
        />
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Table columns={[100]} style={{ width: 100 }}>
            <Row
              style={{
                height: 40,
                // @ts-expect-error backgroundColor removed from Row.style
                backgroundColor: "#EEEEEE",
              }}
            >
              <Cell>A</Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <Slide id="s1">
          <Table columns={[100]} style={{ width: 100 }}>
            <Row style={{ height: 40 }}>
              <Cell
                style={{
                  // @ts-expect-error borderColor removed from Cell.style
                  borderColor: "#111111",
                }}
              >
                A
              </Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
      "unknown-prop",
    );
    expectSyncBuildErrorCode(
      () =>
        defineSlideLayout({
          id: "l-style-typed",
          placeholders: {
            title: {
              type: "title",
              style: {
                // @ts-expect-error backgroundColor removed from placeholder style
                backgroundColor: "#111111",
              },
            },
          },
        }),
      "unknown-prop",
    );

    const legacyStyleKeys = [
      ["backgroundColor", "#FF0000"],
      [
        "backgroundImage",
        {
          type: "linear-gradient",
          angle: 0,
          stops: [
            { position: 0, color: "#000000" },
            { position: 1, color: "#FFFFFF" },
          ],
        },
      ],
      ["borderColor", "#000000"],
      ["borderWidth", 2],
      ["borderStyle", "solid"],
      ["startArrow", "triangle"],
      ["endArrow", "arrow"],
      ["borderStartArrow", "triangle"],
      ["borderEndArrow", "arrow"],
    ] as const;

    for (const [key, val] of legacyStyleKeys) {
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Shape
              style={
                { width: 100, height: 50, [key]: val } as Record<
                  string,
                  unknown
                >
              }
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Box
              style={
                { width: 100, height: 50, [key]: val } as Record<
                  string,
                  unknown
                >
              }
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Text
              style={
                { width: 100, height: 50, [key]: val } as Record<
                  string,
                  unknown
                >
              }
            >
              Hi
            </Text>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Image
              src={png}
              style={
                { width: 100, height: 50, [key]: val } as Record<
                  string,
                  unknown
                >
              }
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1" style={{ [key]: val } as Record<string, unknown>} />
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Table columns={[100]} style={{ width: 100 }}>
              <Row
                style={{ height: 40, [key]: val } as Record<string, unknown>}
              >
                <Cell>A</Cell>
              </Row>
            </Table>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <Slide id="s1">
            <Table columns={[100]} style={{ width: 100 }}>
              <Row style={{ height: 40 }}>
                <Cell style={{ [key]: val } as Record<string, unknown>}>A</Cell>
              </Row>
            </Table>
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      const placeholderTestLayout = defineSlideLayout({
        id: "l-ph-style",
        placeholders: { body: { type: "body" } },
      });
      await expectBuildErrorCode(
        <Deck width={960} height={540}>
          <SlideMaster id="m1">
            <SlideLayout layout={placeholderTestLayout} />
          </SlideMaster>
          <Slide id="s1" layout={placeholderTestLayout}>
            <Placeholder
              placeholder="body"
              style={{ [key]: val } as Record<string, unknown>}
            />
          </Slide>
        </Deck>,
        "unknown-prop",
      );
      expectSyncBuildErrorCode(
        () =>
          defineSlideLayout({
            id: "l-style",
            placeholders: {
              title: {
                type: "title",
                style: { [key]: val } as Record<string, unknown>,
              },
            },
          }),
        "unknown-prop",
      );
    }

    // Inline SlideLayout / slideLayout placeholder specs during resolve
    await expectBuildErrorCode(
      <Deck width={960} height={540}>
        <SlideMaster id="m1">
          <SlideLayout
            id="l-inline"
            // @ts-expect-error legacy top-level geometry removed on inline SlideLayout placeholder spec
            placeholders={{ title: { type: "title", x: 10 } }}
          />
        </SlideMaster>
      </Deck>,
      "unknown-prop",
    );
    await expectBuildErrorCode(
      deck({ width: 960, height: 540 }).append(
        slideMaster({ id: "m1" }).append(
          slideLayout({
            id: "l-inline",
            placeholders: { title: { type: "title", x: 10 } },
          }),
        ),
      ),
      "unknown-prop",
    );
  });
});
