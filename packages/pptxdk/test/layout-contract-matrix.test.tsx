import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { deflateSync } from "node:zlib";

import { BuildError, build, font, resolve } from "@officedk/pptxdk";
import { preview } from "@officedk/pptxdk/preview";
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
  expandDocument,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  chart,
  deck,
  defineSlideLayout,
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
import { validatePptx } from "@pptxdk/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ ((crc & 1) !== 0 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.byteLength, 0);
  const body = Buffer.concat([typeBuf, Buffer.from(data)]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([lenBuf, body, crcBuf]);
}

function createSolidPng(width: number, height: number): Uint8Array {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const rowSize = 1 + width * 4;
  const raw = Buffer.alloc(height * rowSize);
  for (let y = 0; y < height; y += 1) {
    raw[y * rowSize] = 0;
  }
  const idat = deflateSync(raw);
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      pngChunk("IHDR", ihdr),
      pngChunk("IDAT", idat),
      pngChunk("IEND", new Uint8Array(0)),
    ]),
  );
}

const largePng1600x900 = createSolidPng(1600, 900);
const freePng120x80 = createSolidPng(120, 80);

const sampleChartGroups = [
  {
    kind: "column" as const,
    data: { kind: "shared-categories" as const, categories: ["Q1", "Q2"] },
    series: [{ name: "Revenue", values: [120, 180] }],
  },
] as const;

describe("Ticket 01: layout contract and min-content propagation", () => {
  it("supports Chart and Group across flex, percentage, auto, and CSS Grid 1fr tracks", async () => {
    const jsxDoc = createElement(
      Deck,
      { width: 800, height: 500, style: { fontFamily: jostFont } },
      createElement(
        Slide,
        {
          id: "chart-and-group-layout",
          style: {
            display: "flex",
            flexDirection: "column",
            padding: 20,
            gap: 20,
          },
        },
        createElement(
          Box,
          {
            id: "flex-row",
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 20,
              width: 760,
              height: 180,
            },
          },
          createElement(Chart, {
            id: "flex-chart",
            groups: sampleChartGroups,
            style: { flex: 1, height: "100%" },
          }),
          createElement(
            Group,
            {
              id: "flex-group",
              style: { flex: 1, minWidth: 0 },
            },
            createElement(Shape, {
              id: "fg-shape",
              style: { width: 80, height: 40, background: "#2563EB" },
            }),
          ),
        ),
        createElement(
          Box,
          {
            id: "grid-box",
            style: {
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gridTemplateRows: "200px",
              gap: 20,
              width: 760,
              height: 200,
            },
          },
          createElement(Chart, {
            id: "grid-chart",
            groups: sampleChartGroups,
            style: { width: "100%", height: "auto" },
          }),
          createElement(
            Group,
            {
              id: "grid-group",
            },
            createElement(Shape, {
              id: "gg-shape",
              style: { width: 90, height: 50, background: "#16A34A" },
            }),
          ),
        ),
      ),
    );

    const vanillaDoc = deck({
      width: 800,
      height: 500,
      style: { fontFamily: jostFont },
    }).append(
      slide({
        id: "chart-and-group-layout",
        style: {
          display: "flex",
          flexDirection: "column",
          padding: 20,
          gap: 20,
        },
      }).append(
        box({
          id: "flex-row",
          style: {
            display: "flex",
            flexDirection: "row",
            gap: 20,
            width: 760,
            height: 180,
          },
        }).append(
          chart({
            id: "flex-chart",
            groups: sampleChartGroups,
            style: { flex: 1, height: "100%" },
          }),
          group({
            id: "flex-group",
            style: { flex: 1, minWidth: 0 },
          }).append(
            shape({
              id: "fg-shape",
              style: { width: 80, height: 40, background: "#2563EB" },
            }),
          ),
        ),
        box({
          id: "grid-box",
          style: {
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gridTemplateRows: "200px",
            gap: 20,
            width: 760,
            height: 200,
          },
        }).append(
          chart({
            id: "grid-chart",
            groups: sampleChartGroups,
            style: { width: "100%", height: "auto" },
          }),
          group({
            id: "grid-group",
          }).append(
            shape({
              id: "gg-shape",
              style: { width: 90, height: 50, background: "#16A34A" },
            }),
          ),
        ),
      ),
    );

    const jsxExpanded = await expandDocument(jsxDoc);
    const jsxResolution = resolve(jsxExpanded);
    const vanillaResolution = resolve(vanillaDoc);

    for (const res of [jsxResolution, vanillaResolution]) {
      const s = res.slide("chart-and-group-layout");
      const flexRow = s.get("flex-row");
      const gridBox = s.get("grid-box");
      assert.deepEqual(flexRow.get("flex-chart").box, {
        x: 20,
        y: 20,
        width: 370,
        height: 180,
      });
      assert.deepEqual(flexRow.get("flex-group").box, {
        x: 410,
        y: 20,
        width: 370,
        height: 180,
      });
      assert.deepEqual(gridBox.get("grid-chart").box, {
        x: 20,
        y: 220,
        width: 370,
        height: 200,
      });
      assert.deepEqual(gridBox.get("grid-group").box, {
        x: 410,
        y: 220,
        width: 370,
        height: 200,
      });
    }

    const built = await build(jsxExpanded);
    const validated = await validatePptx(built.pptx);
    assert.deepEqual(validated.errors, []);
    const previewed = await preview(jsxExpanded);
    assert.equal(previewed.pages.length, 1);
  });

  it("prevents percentage-sized high-resolution Image in flex column from squeezing sibling Chart while preserving intrinsic size on free images", async () => {
    const doughnutGroups = [
      {
        kind: "doughnut" as const,
        data: { kind: "shared-categories" as const, categories: ["A", "B"] },
        series: [{ name: "Share", values: [60, 40] }],
      },
    ] as const;

    const jsxDoc = createElement(
      Deck,
      { width: 840, height: 500, style: { fontFamily: jostFont } },
      createElement(
        Slide,
        { id: "image-flex-sibling", style: { padding: 20 } },
        createElement(
          Box,
          {
            id: "split-row",
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 20,
              width: 800,
              height: 300,
            },
          },
          createElement(Chart, {
            id: "fixed-chart",
            groups: doughnutGroups,
            style: { width: 280, height: 280 },
          }),
          createElement(
            Box,
            { id: "flex-image-col", style: { flex: 1 } },
            createElement(Image, {
              id: "percent-hero-image",
              src: largePng1600x900,
              style: { width: "100%" },
            }),
          ),
        ),
        createElement(Image, {
          id: "free-intrinsic-image",
          src: freePng120x80,
        }),
      ),
    );

    const vanillaDoc = deck({
      width: 840,
      height: 500,
      style: { fontFamily: jostFont },
    }).append(
      slide({ id: "image-flex-sibling", style: { padding: 20 } }).append(
        box({
          id: "split-row",
          style: {
            display: "flex",
            flexDirection: "row",
            gap: 20,
            width: 800,
            height: 300,
          },
        }).append(
          chart({
            id: "fixed-chart",
            groups: doughnutGroups,
            style: { width: 280, height: 280 },
          }),
          box({ id: "flex-image-col", style: { flex: 1 } }).append(
            image({
              id: "percent-hero-image",
              src: largePng1600x900,
              style: { width: "100%" },
            }),
          ),
        ),
        image({
          id: "free-intrinsic-image",
          src: freePng120x80,
        }),
      ),
    );

    const jsxRes = resolve(await expandDocument(jsxDoc));
    const vanillaRes = resolve(vanillaDoc);

    for (const res of [jsxRes, vanillaRes]) {
      const s = res.slide("image-flex-sibling");
      const splitRow = s.get("split-row");
      assert.deepEqual(splitRow.get("fixed-chart").box, {
        x: 20,
        y: 20,
        width: 280,
        height: 280,
      });
      assert.deepEqual(splitRow.get("flex-image-col").box, {
        x: 320,
        y: 20,
        width: 500,
        height: 300,
      });
      assert.deepEqual(
        splitRow.get("flex-image-col").get("percent-hero-image").box,
        {
          x: 320,
          y: 20,
          width: 500,
          height: 281.25,
        },
      );
      assert.deepEqual(s.get("free-intrinsic-image").box, {
        x: 20,
        y: 320,
        width: 120,
        height: 80,
      });
    }
  });

  it("reports explicit 0px collapse diagnostics across Box, Shape, Text, Image, Group, Table, and Chart while keeping EMU overflow distinct", async () => {
    const collapseCases = [
      {
        label: "visible Box",
        doc: () =>
          deck({ width: 640, height: 360 }).append(
            slide({ id: "s1" }).append(
              box({ style: { width: 200, background: "#E2E8F0" } }),
            ),
          ),
        expectedText: /0px/u,
      },
      {
        label: "Shape",
        doc: () =>
          deck({ width: 640, height: 360 }).append(
            slide({ id: "s1" }).append(
              shape({ style: { width: 0, height: 40, background: "#2563EB" } }),
            ),
          ),
        expectedText: /width.*0px/iu,
      },
      {
        label: "Text",
        doc: () =>
          deck({
            width: 640,
            height: 360,
            style: { fontFamily: jostFont },
          }).append(
            slide({ id: "s1" }).append(text("Hello", { style: { width: 0 } })),
          ),
        expectedText: /width.*0px/iu,
      },
      {
        label: "Image",
        doc: () =>
          deck({ width: 640, height: 360 }).append(
            slide({ id: "s1" }).append(
              image({ src: freePng120x80, style: { width: 0, height: 40 } }),
            ),
          ),
        expectedText: /width.*0px/iu,
      },
      {
        label: "Group",
        doc: () =>
          deck({ width: 640, height: 360 }).append(
            slide({ id: "s1" }).append(
              group({ style: { width: 0, height: 40 } }).append(
                shape({
                  style: {
                    position: "absolute",
                    left: 0,
                    top: 0,
                    width: 40,
                    height: 40,
                  },
                }),
              ),
            ),
          ),
        expectedText: /width.*0px/iu,
      },
      {
        label: "Table",
        doc: () =>
          deck({
            width: 640,
            height: 360,
            style: { fontFamily: jostFont },
          }).append(
            slide({ id: "s1" }).append(
              table({
                columns: [100],
                style: { width: 0 },
              }).append(row().append(cell("Cell"))),
            ),
          ),
        expectedText: /0px/iu,
      },
      {
        label: "Chart",
        doc: () =>
          deck({
            width: 640,
            height: 360,
            style: { fontFamily: jostFont },
          }).append(
            slide({ id: "s1" }).append(
              chart({
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [10] }],
                  },
                ],
                style: { width: 0, height: 120 },
              }),
            ),
          ),
        expectedText: /width.*0px/iu,
      },
    ];

    for (const item of collapseCases) {
      assert.throws(
        () => resolve(item.doc()),
        (error: unknown) => {
          assert(
            error instanceof BuildError,
            `${item.label} should throw BuildError`,
          );
          assert.match(
            error.message,
            item.expectedText,
            `${item.label} message should mention 0px, got: ${error.message}`,
          );
          return true;
        },
      );
    }

    // Oversized dimensions should still report EMU range overflow rather than 0px.
    assert.throws(
      () =>
        resolve(
          deck({ width: 640, height: 360 }).append(
            slide({ id: "s1" }).append(
              shape({ style: { width: 500_000, height: 40 } }),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.doesNotMatch(error.message, /0px/u);
        assert.match(error.message, /EMU|range/iu);
        return true;
      },
    );
  });

  it("evaluates Box, Group, Shape, Text, Image, Table, Chart, and bound Placeholder identically across percentage, auto, flex, Grid 1fr, and opposing offsets in JSX and Vanilla", async () => {
    const sharedPad4 = [9, 12, 9, 12] as const;
    const sharedPad2 = [9, 12] as const;

    type CapabilityKind =
      | "box"
      | "group"
      | "shape"
      | "text"
      | "image"
      | "table"
      | "chart";

    const buildJsxCapability = (
      kind: CapabilityKind,
      id: string,
      style: Record<string, unknown>,
    ) => {
      switch (kind) {
        case "box":
          return createElement(Box, {
            id,
            style: { ...style, background: "#E2E8F0", padding: sharedPad4 },
          });
        case "group":
          return createElement(
            Group,
            { id, style },
            createElement(Shape, {
              id: `${id}-member`,
              style: { width: 40, height: 24, background: "#2563EB" },
            }),
          );
        case "shape":
          return createElement(
            Shape,
            {
              id,
              style: { ...style, background: "#1D4ED8", padding: sharedPad2 },
            },
            "S",
          );
        case "text":
          return createElement(
            Text,
            { id, style: { ...style, padding: sharedPad4 } },
            "Hi",
          );
        case "image":
          return createElement(Image, {
            id,
            src: freePng120x80,
            style,
          });
        case "table":
          return createElement(
            Table,
            { id, columns: ["1fr"], style },
            createElement(
              Row,
              null,
              createElement(Cell, { style: { padding: sharedPad4 } }, "C"),
            ),
          );
        case "chart":
          return createElement(Chart, {
            id,
            groups: sampleChartGroups,
            style,
          });
      }
    };

    const buildVanillaCapability = (
      kind: CapabilityKind,
      id: string,
      style: Record<string, unknown>,
    ) => {
      switch (kind) {
        case "box":
          return box({
            id,
            style: { ...style, background: "#E2E8F0", padding: sharedPad4 },
          });
        case "group":
          return group({ id, style }).append(
            shape({
              id: `${id}-member`,
              style: { width: 40, height: 24, background: "#2563EB" },
            }),
          );
        case "shape":
          return shape("S", {
            id,
            style: { ...style, background: "#1D4ED8", padding: sharedPad2 },
          });
        case "text":
          return text("Hi", { id, style: { ...style, padding: sharedPad4 } });
        case "image":
          return image({ id, src: freePng120x80, style });
        case "table":
          return table({ id, columns: ["1fr"], style }).append(
            row().append(cell("C", { style: { padding: sharedPad4 } })),
          );
        case "chart":
          return chart({ id, groups: sampleChartGroups, style });
      }
    };

    const capabilities: readonly CapabilityKind[] = [
      "box",
      "group",
      "shape",
      "text",
      "image",
      "table",
      "chart",
    ];

    const phLayout = defineSlideLayout({
      id: "matrix-ph-layout",
      placeholders: {
        slotPct: {
          type: "obj",
          style: { left: 20, top: 20, width: 200, height: 100 },
        },
        slotInset: {
          type: "obj",
          style: { inset: 20 },
        },
        slotSides: {
          type: "obj",
          style: { left: 20, right: 20, top: 20, bottom: 20 },
        },
      },
    });

    const jsxSlides = [
      ...capabilities.map((kind) =>
        createElement(
          Slide,
          { id: `matrix-${kind}` },
          createElement(
            Box,
            {
              id: "pct-host",
              style: {
                position: "absolute",
                left: 0,
                top: 0,
                width: 400,
                height: 200,
              },
            },
            buildJsxCapability(kind, "pct-item", {
              width: "50%",
              height: "50%",
            }),
          ),
          createElement(
            Box,
            {
              id: "auto-host",
              style: {
                position: "absolute",
                left: 420,
                top: 0,
                width: 400,
                height: 200,
                display: "flex",
                flexDirection: "row",
              },
            },
            buildJsxCapability(kind, "auto-item", {
              width: "auto",
              height: "auto",
            }),
          ),
          createElement(
            Box,
            {
              id: "flex-host",
              style: {
                position: "absolute",
                left: 0,
                top: 220,
                width: 400,
                height: 200,
                display: "flex",
                flexDirection: "row",
                gap: 20,
              },
            },
            createElement(Box, { style: { width: 100, height: 200 } }),
            buildJsxCapability(kind, "flex-item", { flex: 1, minWidth: 0 }),
          ),
          createElement(
            Box,
            {
              id: "grid-host",
              style: {
                position: "absolute",
                left: 420,
                top: 220,
                width: 400,
                height: 160,
                display: "grid",
                gridTemplateColumns: "100px 1fr",
                gridTemplateRows: "160px",
                gap: 20,
              },
            },
            createElement(Box, { style: { height: 160 } }),
            buildJsxCapability(kind, "grid-item", {}),
          ),
          createElement(
            Box,
            {
              id: "inset-host",
              style: {
                position: "absolute",
                left: 0,
                top: 440,
                width: 400,
                height: 200,
              },
            },
            buildJsxCapability(kind, "inset-item", {
              position: "absolute",
              inset: 20,
            }),
          ),
          createElement(
            Box,
            {
              id: "sides-host",
              style: {
                position: "absolute",
                left: 420,
                top: 440,
                width: 400,
                height: 200,
              },
            },
            buildJsxCapability(kind, "sides-item", {
              position: "absolute",
              left: 30,
              right: 50,
              top: 20,
              bottom: 40,
            }),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "matrix-placeholder", layout: phLayout },
        createElement(Placeholder, {
          id: "ph-pct",
          placeholder: "slotPct",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: "50%",
            height: "50%",
          },
        }),
        createElement(Placeholder, {
          id: "ph-inset",
          placeholder: "slotInset",
          style: { position: "absolute", inset: 40 },
        }),
        createElement(Placeholder, {
          id: "ph-sides",
          placeholder: "slotSides",
          style: {
            position: "absolute",
            left: 60,
            right: 140,
            top: 50,
            bottom: 70,
          },
        }),
      ),
    ];

    const vanillaSlides = [
      ...capabilities.map((kind) =>
        slide({ id: `matrix-${kind}` }).append(
          box({
            id: "pct-host",
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 400,
              height: 200,
            },
          }).append(
            buildVanillaCapability(kind, "pct-item", {
              width: "50%",
              height: "50%",
            }),
          ),
          box({
            id: "auto-host",
            style: {
              position: "absolute",
              left: 420,
              top: 0,
              width: 400,
              height: 200,
              display: "flex",
              flexDirection: "row",
            },
          }).append(
            buildVanillaCapability(kind, "auto-item", {
              width: "auto",
              height: "auto",
            }),
          ),
          box({
            id: "flex-host",
            style: {
              position: "absolute",
              left: 0,
              top: 220,
              width: 400,
              height: 200,
              display: "flex",
              flexDirection: "row",
              gap: 20,
            },
          }).append(
            box({ style: { width: 100, height: 200 } }),
            buildVanillaCapability(kind, "flex-item", {
              flex: 1,
              minWidth: 0,
            }),
          ),
          box({
            id: "grid-host",
            style: {
              position: "absolute",
              left: 420,
              top: 220,
              width: 400,
              height: 160,
              display: "grid",
              gridTemplateColumns: "100px 1fr",
              gridTemplateRows: "160px",
              gap: 20,
            },
          }).append(
            box({ style: { height: 160 } }),
            buildVanillaCapability(kind, "grid-item", {}),
          ),
          box({
            id: "inset-host",
            style: {
              position: "absolute",
              left: 0,
              top: 440,
              width: 400,
              height: 200,
            },
          }).append(
            buildVanillaCapability(kind, "inset-item", {
              position: "absolute",
              inset: 20,
            }),
          ),
          box({
            id: "sides-host",
            style: {
              position: "absolute",
              left: 420,
              top: 440,
              width: 400,
              height: 200,
            },
          }).append(
            buildVanillaCapability(kind, "sides-item", {
              position: "absolute",
              left: 30,
              right: 50,
              top: 20,
              bottom: 40,
            }),
          ),
        ),
      ),
      slide({ id: "matrix-placeholder", layout: phLayout }).append(
        placeholder({
          id: "ph-pct",
          placeholder: "slotPct",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: "50%",
            height: "50%",
          },
        }),
        placeholder({
          id: "ph-inset",
          placeholder: "slotInset",
          style: { position: "absolute", inset: 40 },
        }),
        placeholder({
          id: "ph-sides",
          placeholder: "slotSides",
          style: {
            position: "absolute",
            left: 60,
            right: 140,
            top: 50,
            bottom: 70,
          },
        }),
      ),
    ];

    const jsxDoc = createElement(
      Deck,
      { width: 840, height: 680, style: { fontFamily: jostFont } },
      createElement(
        SlideMaster,
        { id: "matrix-master" },
        createElement(SlideLayout, { layout: phLayout }),
      ),
      ...jsxSlides,
    );

    const vanillaDoc = deck({
      width: 840,
      height: 680,
      style: { fontFamily: jostFont },
    }).append(
      slideMaster({ id: "matrix-master" }).append(slideLayout(phLayout)),
      ...vanillaSlides,
    );

    const jsxExpanded = await expandDocument(jsxDoc);
    const jsxRes = resolve(jsxExpanded);
    const vanillaRes = resolve(vanillaDoc);

    for (const res of [jsxRes, vanillaRes]) {
      for (const kind of capabilities) {
        const s = res.slide(`matrix-${kind}`);
        assert.deepEqual(s.get("pct-host").get("pct-item").box, {
          x: 0,
          y: 0,
          width: 200,
          height: 100,
        });
        assert.ok((s.get("auto-host").get("auto-item").box?.width ?? 0) > 0);
        assert.deepEqual(s.get("flex-host").get("flex-item").box, {
          x: 120,
          y: 220,
          width: 280,
          height: 200,
        });
        assert.deepEqual(s.get("grid-host").get("grid-item").box, {
          x: 540,
          y: 220,
          width: 280,
          height: 160,
        });
        assert.deepEqual(s.get("inset-host").get("inset-item").box, {
          x: 20,
          y: 460,
          width: 360,
          height: 160,
        });
        assert.deepEqual(s.get("sides-host").get("sides-item").box, {
          x: 450,
          y: 460,
          width: 320,
          height: 140,
        });
      }

      const phSlide = res.slide("matrix-placeholder");
      assert.deepEqual(phSlide.get("ph-pct").box, {
        x: 10,
        y: 10,
        width: 420,
        height: 340,
      });
      assert.deepEqual(phSlide.get("ph-inset").box, {
        x: 40,
        y: 40,
        width: 760,
        height: 600,
      });
      assert.deepEqual(phSlide.get("ph-sides").box, {
        x: 60,
        y: 50,
        width: 640,
        height: 560,
      });
    }

    const jsxBuilt = await build(jsxExpanded);
    const vanillaBuilt = await build(vanillaDoc);
    assert.deepEqual((await validatePptx(jsxBuilt.pptx)).errors, []);
    assert.deepEqual((await validatePptx(vanillaBuilt.pptx)).errors, []);

    const jsxPreview = await preview(jsxExpanded);
    const vanillaPreview = await preview(vanillaDoc);
    assert.equal(jsxPreview.pages.length, capabilities.length + 1);
    assert.equal(vanillaPreview.pages.length, capabilities.length + 1);
  });
});
