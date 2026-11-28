import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  boxCapability,
  chartCapability,
  groupCapability,
  imageCapability,
  motionCapability,
  notesCapability,
  shapeCapability,
  slideBackgroundCapability,
  tableCapability,
  textCapability,
  transitionCapability,
} from "@pptxdk/capabilities";
import { documentCapabilities } from "@pptxdk/capabilities/all";
import { Box } from "@pptxdk/capabilities/box-elements";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  box,
  cell,
  chart,
  createMotionScopeHandle,
  deck,
  defineSlideLayout,
  font,
  group,
  image,
  motionScope,
  notes,
  paragraph,
  placeholder,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  step,
  table,
  text,
  transitions,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Image } from "@pptxdk/capabilities/image-elements";
import { breathe, fade, spin } from "@pptxdk/capabilities/motion";
import { MotionScope, Step } from "@pptxdk/capabilities/motion-elements";
import { placeholderCapability } from "@pptxdk/capabilities/placeholder";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Cell, Row, Table } from "@pptxdk/capabilities/table-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import {
  codegenDiagnosticSchema,
  type CodegenDiagnostic,
} from "@pptxdk/codegen";
import {
  BuildError,
  backgroundValueSchema,
  borderObjectSchema,
  borderValueSchema,
  buildDiagnosticSchema,
  createMorphScope,
  layoutStyleSchema,
  placeholderSizeSchema,
  previewErrorSchema,
  previewPlaceholderSchema,
  visualStyleSchema,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);
const jostRegular = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostRegular }]);

const resolveAll = createResolver(documentCapabilities);

describe("Capability pure-data schema closure: Box, Group, and Shape", () => {
  it("resolves valid Box, Group, and Shape contracts and rejects invalid props and styles identically", async () => {
    const morphScope = createMorphScope();
    const heroKey = morphScope.handle("hero");

    const validDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { major: jostFont, minor: jostFont } },
    }).append(
      slide({ id: "s1" }).append(
        box({
          id: "b1",
          morphKey: heroKey,
          style: {
            display: "flex",
            width: 300,
            padding: 12,
            fontSize: 18,
            background: "#F8FAFC",
            border: { color: "#CBD5E1", width: 1, style: "solid" },
          },
          enter: fade(),
          motion: true,
        }),
        group({
          id: "g1",
          name: "HeaderGroup",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: 60,
            display: "flex",
            fontSize: 16,
            rotate: 15,
          },
          enter: fade(),
          emphasis: spin(),
          motion: true,
        }).append(
          shape({
            id: "sh1",
            name: "Callout",
            geometry: "roundRect",
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 120,
              height: 60,
              background: "#2563EB",
              border: {
                color: "#1D4ED8",
                width: 2,
                startArrow: "none",
                endArrow: "triangle",
              },
              fontSize: "16pt",
              fontWeight: "bold",
            },
            enter: fade(),
            emphasis: spin(),
            motion: true,
          }),
        ),
      ),
    );
    assert.equal(resolveAll(validDoc).slides.length, 1);

    const resolveSubset = createResolver([
      slideBackgroundCapability,
      boxCapability,
      groupCapability,
      shapeCapability,
      motionCapability,
    ]);

    const invalidCases = [
      {
        label: "blank Box id",
        expectedCode: "invalid-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Box, { id: "  " }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(box({ id: "  " })),
          ),
      },
      {
        label: "unknown style prop on Box",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Box, { style: { unknownStyleProp: 1 } as never }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              box({ style: { unknownStyleProp: 1 } as never }),
            ),
          ),
      },
      {
        label: "emphasis effect passed to Box.enter",
        expectedCode: "invalid-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Box, { enter: spin() as never }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(box({ enter: spin() as never })),
          ),
      },
      {
        label: "invalid width type on Shape style inside Box",
        expectedCode: "invalid-style",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Box,
                { style: { display: "flex", width: 300 } },
                createElement(Shape, {
                  id: "bad-w",
                  style: { width: true as never, height: 40 },
                }),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              box({ style: { display: "flex", width: 300 } }).append(
                shape({
                  id: "bad-w",
                  style: { width: true as never, height: 40 },
                }),
              ),
            ),
          ),
      },
      {
        label: "invalid border value on Shape",
        expectedCode: "invalid-style",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Shape,
                {
                  id: "bad-border",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 100,
                    height: 40,
                    border: "solid" as never,
                  },
                },
                "Hi",
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape("Hi", {
                id: "bad-border",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                  border: "solid" as never,
                },
              }),
            ),
          ),
      },
      {
        label: "background on Group",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Group,
                {
                  id: "bad-g",
                  style: { background: "#FFFFFF" } as never,
                },
                createElement(Shape, {
                  id: "m1",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 40,
                    height: 40,
                  },
                }),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              group({
                id: "bad-g",
                style: { background: "#FFFFFF" } as never,
              }).append(
                shape({
                  id: "m1",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 40,
                    height: 40,
                  },
                }),
              ),
            ),
          ),
      },
    ] as const;

    for (const item of invalidCases) {
      const jsxDoc = await expandDocument(item.jsx());
      const vanillaDoc = item.vanilla();

      for (const [entry, doc] of [
        ["jsx", jsxDoc],
        ["vanilla", vanillaDoc],
      ] as const) {
        assert.throws(
          () => resolveSubset(doc),
          (error: unknown) =>
            error instanceof BuildError &&
            error.diagnostics.some((d) => d.code === item.expectedCode),
          `${item.label} (${entry}) should fail with ${item.expectedCode}`,
        );
      }
    }
  });
});

describe("Capability pure-data schema closure: Text, Image, Table, Chart, Notes, and Font", () => {
  it("enforces public contracts on Text, Image, Table, Chart, Notes, and Font across JSX and imperative entries", async () => {
    const morphScope = createMorphScope();
    const heroKey = morphScope.handle("hero");

    assert.throws(() => font("   "), TypeError);
    assert.equal(
      font("Jost*", [{ weight: 450, data: png1x1 }]).faces?.[0]?.weight,
      450,
    );
    assert.throws(
      () => font("Jost*", [{ weight: 0, data: png1x1 }]),
      TypeError,
    );
    assert.throws(
      () => font("Jost*", [{ weight: 1001, data: png1x1 }]),
      TypeError,
    );
    assert.throws(
      () => font("Jost*", [{ weight: 400.5, data: png1x1 }]),
      TypeError,
    );
    assert.throws(
      () => font("Jost*", [{ weight: 400, data: new Uint8Array(0) }]),
      TypeError,
    );

    const validDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { major: jostFont, minor: jostFont } },
    }).append(
      slide({ id: "s1" }).append(
        text("Headline", {
          id: "t1",
          morphKey: heroKey,
          style: {
            position: "absolute",
            left: 20,
            top: 30,
            width: 300,
            height: "auto",
            fontFamily: jostFont,
            fontSize: "24pt",
            color: "#0F172A",
            padding: 8,
          },
          enter: fade(),
          emphasis: spin(),
          motion: true,
        }),
        image({
          id: "img1",
          src: png1x1,
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 20,
            top: 100,
            width: 40,
            height: 40,
          },
          enter: fade(),
        }),
        table({
          columns: [120, "1fr"],
          style: { position: "absolute", left: 20, top: 160, width: 240 },
          enter: fade(),
          emphasis: spin(),
          motion: true,
        }).append(row().append(cell("A"), cell("B"))),
        chart({
          id: "ch1",
          style: {
            position: "absolute",
            left: 340,
            top: 40,
            width: 320,
            height: 200,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S1", values: [1] }],
            },
          ],
          enter: fade(),
          emphasis: spin(),
        }),
        notes("Valid note", { id: "n1" }),
      ),
    );
    assert.equal(resolveAll(validDoc).slides.length, 1);

    const resolve = createResolver([
      slideBackgroundCapability,
      boxCapability,
      textCapability,
      imageCapability,
      tableCapability,
      chartCapability,
      notesCapability,
      motionCapability,
    ]);

    for (const [label, invalidDoc, expectedCode] of [
      [
        "empty Text id",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(text("Hi", { id: "" })),
        ),
        "invalid-prop",
      ],
      [
        "blank Text name",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(text("Hi", { name: "   " })),
        ),
        "invalid-prop",
      ],
      [
        "invalid morphKey on Text",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            text("Hi", { morphKey: "not-a-key" as never }),
          ),
        ),
        "invalid-prop",
      ],
      [
        "blank P id",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            text([paragraph("Hi", { id: "  " })], {
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 100,
                height: 40,
              },
            }),
          ),
        ),
        "invalid-prop",
      ],
      [
        "blank Span id",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            text([paragraph([span("Hi", { id: "  " })])], {
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 100,
                height: 40,
              },
            }),
          ),
        ),
        "invalid-prop",
      ],
      [
        "blank SlideNumber id",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            text([paragraph([slideNumber({ id: "  " })])], {
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 100,
                height: 40,
              },
            }),
          ),
        ),
        "invalid-prop",
      ],
      [
        "invalid Image preset geometry",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            image({ src: png1x1, geometry: "not-a-preset" as never }),
          ),
        ),
        "invalid-prop",
      ],
      [
        "blank Notes id",
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(notes("Body", { id: "  " })),
        ),
        "invalid-prop",
      ],
    ] as const) {
      assert.throws(
        () => resolve(invalidDoc),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === expectedCode),
        `${label} should fail with ${expectedCode}`,
      );
    }

    const invalidCases = [
      {
        label: "invalid width type on Text inside Box",
        expectedCode: "invalid-style",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Box,
                { style: { display: "flex", width: 300 } },
                createElement(
                  Text,
                  {
                    id: "bad-tw",
                    style: { width: true as never, height: 40 },
                  },
                  "Hello",
                ),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              box({ style: { display: "flex", width: 300 } }).append(
                text("Hello", {
                  id: "bad-tw",
                  style: { width: true as never, height: 40 },
                }),
              ),
            ),
          ),
      },
      {
        label: "unsupported align on Table.style",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Table,
                {
                  id: "bad-tbl",
                  columns: [120],
                  style: { align: "center" } as never,
                },
                createElement(Row, null, createElement(Cell, null, "A")),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              table({
                id: "bad-tbl",
                columns: [120],
                style: { align: "center" } as never,
              }).append(row().append(cell("A"))),
            ),
          ),
      },
      {
        label: "unsupported textWrap on Cell.style",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Table,
                { id: "bad-cell-tbl", columns: [120] },
                createElement(
                  Row,
                  null,
                  createElement(
                    Cell,
                    { style: { textWrap: "nowrap" } as never },
                    "A",
                  ),
                ),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              table({ id: "bad-cell-tbl", columns: [120] }).append(
                row().append(
                  cell("A", { style: { textWrap: "nowrap" } as never }),
                ),
              ),
            ),
          ),
      },
      {
        label: "unknown prop on Image",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Image, {
                id: "bad-img",
                src: png1x1,
                extraProp: true,
              } as never),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              image({
                id: "bad-img",
                src: png1x1,
                extraProp: true,
              } as never),
            ),
          ),
      },
      {
        label: "unknown prop on Chart",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Chart, {
                id: "bad-ch",
                extraProp: 1,
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              } as never),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              chart({
                id: "bad-ch",
                extraProp: 1,
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              } as never),
            ),
          ),
      },
    ] as const;

    for (const item of invalidCases) {
      const jsxDoc = await expandDocument(item.jsx());
      const vanillaDoc = item.vanilla();

      for (const [entry, doc] of [
        ["jsx", jsxDoc],
        ["vanilla", vanillaDoc],
      ] as const) {
        assert.throws(
          () => resolve(doc),
          (error: unknown) =>
            error instanceof BuildError &&
            error.diagnostics.some((d) => d.code === item.expectedCode),
          `${item.label} (${entry}) should fail with ${item.expectedCode}`,
        );
      }
    }
  });
});

describe("Capability pure-data schema closure: Transition and Motion", () => {
  it("validates Transition, MotionScope, Step, Spin, and Breathe options consistently across JSX and imperative entries", async () => {
    const morphTransition = transitions.morph({
      by: "word",
      fallback: transitions.fade(),
    });
    assert.equal(morphTransition.kind, "morph");
    assert.equal(spin({ by: 180 }).preset, "spin");
    assert.equal(spin({ from: 0, to: 90 }).preset, "spin");

    const resolve = createResolver([
      slideBackgroundCapability,
      shapeCapability,
      motionCapability,
      transitionCapability,
    ]);
    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({
              id: "s1",
              transition: {
                type: "transition",
                kind: "morph",
                by: "word",
              } as never,
            }),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-prop"),
    );
    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "bad-spin",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 60,
                  height: 40,
                },
                emphasis: spin({ by: 90, from: 0 } as never),
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-prop"),
    );
    const scopeHandle = createMotionScopeHandle();

    const invalidCases = [
      {
        label: "unknown easing key",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "sh",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 60,
                  height: 40,
                },
                enter: fade({ easing: { accelerate: 0.2, extra: 1 } as never }),
              }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "sh",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 60,
                  height: 40,
                },
                enter: fade({ easing: { accelerate: 0.2, extra: 1 } as never }),
              }),
            ),
          ),
      },
      {
        label: "unknown breathe parameter",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "sh",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 60,
                  height: 40,
                  background: "#112233",
                },
                emphasis: breathe({
                  period: 1000,
                  from: "#112233",
                  to: "#AABBCC",
                  unexpected: true,
                } as never),
              }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "sh",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 60,
                  height: 40,
                  background: "#112233",
                },
                emphasis: breathe({
                  period: 1000,
                  from: "#112233",
                  to: "#AABBCC",
                  unexpected: true,
                } as never),
              }),
            ),
          ),
      },
      {
        label: "unknown MotionScope parameter",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                MotionScope,
                { scope: scopeHandle, bogus: true } as never,
                createElement(Shape, {
                  id: "sh",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 60,
                    height: 40,
                  },
                  enter: fade(),
                }),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              motionScope(scopeHandle, { bogus: true } as never).append(
                shape({
                  id: "sh",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 60,
                    height: 40,
                  },
                  enter: fade(),
                }),
              ),
            ),
          ),
      },
      {
        label: "unknown Step parameter",
        expectedCode: "unknown-prop",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(
                Step,
                { n: 1, bogus: true } as never,
                createElement(Shape, {
                  id: "sh",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 60,
                    height: 40,
                  },
                  enter: fade(),
                }),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              step({ n: 1, bogus: true } as never).append(
                shape({
                  id: "sh",
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 60,
                    height: 40,
                  },
                  enter: fade(),
                }),
              ),
            ),
          ),
      },
      {
        label: "unsupported transition kind",
        expectedCode: "invalid-prop",
        jsx: () =>
          createElement(Deck, { width: 480, height: 270 }, [
            createElement(Slide, {
              key: "s1",
              id: "s1",
              transition: { type: "transition", kind: "blinds" },
            }),
          ]),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({
              id: "s1",
              transition: { type: "transition", kind: "blinds" } as never,
            }),
          ),
      },
    ] as const;

    for (const item of invalidCases) {
      const jsxDoc = await expandDocument(item.jsx());
      const vanillaDoc = item.vanilla();

      for (const [entry, doc] of [
        ["jsx", jsxDoc],
        ["vanilla", vanillaDoc],
      ] as const) {
        assert.throws(
          () => resolve(doc),
          (error: unknown) =>
            error instanceof BuildError &&
            error.diagnostics.some((d) => d.code === item.expectedCode),
          `${item.label} (${entry}) should fail with ${item.expectedCode}`,
        );
      }
    }
  });

  it("enforces exported core and codegen visual, layout, placeholderSize, and diagnostic Zod schemas", () => {
    assert.equal(backgroundValueSchema.safeParse("none").success, true);
    assert.equal(backgroundValueSchema.safeParse("#2563EB").success, true);
    assert.equal(
      backgroundValueSchema.safeParse({
        type: "linear-gradient",
        angle: 90,
        stops: [
          { color: "#112233", position: 0 },
          { color: "#445566", position: 1 },
        ],
      }).success,
      true,
    );
    assert.equal(backgroundValueSchema.safeParse(123).success, false);

    assert.equal(borderValueSchema.safeParse("none").success, true);
    assert.equal(
      borderObjectSchema.safeParse({
        color: "#1E293B",
        width: 2,
        style: "dashed",
      }).success,
      true,
    );
    assert.equal(
      borderValueSchema.safeParse({ color: "#1E293B", startArrow: "triangle" })
        .success,
      false,
    );

    assert.equal(
      visualStyleSchema.safeParse({
        background: "none",
        border: "none",
        boxShadow: "none",
        flipH: true,
        flipV: false,
      }).success,
      true,
    );
    assert.equal(
      layoutStyleSchema.safeParse({
        position: "absolute",
        left: 16,
        top: 24,
        width: 120,
        height: "auto",
      }).success,
      true,
    );
    assert.equal(placeholderSizeSchema.safeParse("full").success, true);
    assert.equal(placeholderSizeSchema.safeParse("half").success, true);
    assert.equal(placeholderSizeSchema.safeParse("quarter").success, true);
    assert.equal(placeholderSizeSchema.safeParse("third").success, false);

    assert.equal(
      buildDiagnosticSchema.safeParse({
        code: "invalid-prop",
        path: "slide[0].shape[0]",
        field: "emphasis",
        message: "spin() cannot combine from/to and by",
      }).success,
      true,
    );
    assert.equal(
      buildDiagnosticSchema.safeParse({
        code: "invalid-prop",
        path: "slide[0].shape[0]",
        message: "legacy",
        severity: "error",
      }).success,
      false,
    );
    assert.equal(
      previewErrorSchema.safeParse({
        code: "preview-missing-font",
        path: "slide[0].text[0]",
        message: "Missing font",
        font: { family: "Inter", weight: 400, style: "normal" },
      }).success,
      true,
    );
    assert.equal(
      previewPlaceholderSchema.safeParse({
        id: "hero-img",
        path: "slide[0].image[0]",
        url: "https://example.com/a.png",
      }).success,
      true,
    );
    const codegenDiag: CodegenDiagnostic = {
      code: "unexpressible-value",
      path: "slide[0].shape[0]",
      field: "style.background",
      message: "Cannot serialize value",
    };
    assert.equal(codegenDiagnosticSchema.safeParse(codegenDiag).success, true);
  });

  it("shares readonly non-negative padding tuples across Box, Text, Shape, Table Cell, and validates Placeholder schemas", () => {
    const pad2 = [9, 12] as const;
    const pad4 = [9, 12, 9, 12] as const;
    const layoutDef = defineSlideLayout({
      id: "l1",
      name: "L1",
      placeholders: {
        body: {
          type: "body",
          style: { left: 10, top: 20, width: 200, height: 100 },
        },
      },
    });

    const validDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { major: jostFont, minor: jostFont } },
    }).append(
      slideMaster({ id: "m1" }).append(slideLayout(layoutDef)),
      slide({ id: "s1", layout: layoutDef }).append(
        box({ style: { padding: pad4 } }),
        text("T", { style: { padding: pad2 } }),
        shape("S", { style: { padding: pad4 } }),
        table({ columns: [120, 120] }).append(
          row().append(
            cell("C1", { style: { padding: pad4 } }),
            cell("C2", { style: { padding: pad2 } }),
          ),
        ),
        placeholder({
          placeholder: "body",
          style: { position: "absolute", left: 10, top: 20 },
        }),
      ),
    );
    assert.equal(resolveAll(validDoc).slides.length, 1);

    const resolveTable = createResolver([tableCapability]);
    assert.throws(
      () =>
        resolveTable(
          deck({ width: 960, height: 540 }).append(
            slide({ id: "s1" }).append(
              table({ columns: [120, 120] }).append(
                row().append(cell("Bad", { style: { padding: [9, -1] } })),
              ),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-style"),
    );

    const resolvePlaceholder = createResolver([
      slideBackgroundCapability,
      placeholderCapability,
    ]);
    assert.throws(
      () =>
        resolvePlaceholder(
          deck({ width: 960, height: 540 }).append(
            slide({ id: "s1" }).append(
              placeholder({ unknownProp: true } as never),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "unknown-prop"),
    );
  });
});
