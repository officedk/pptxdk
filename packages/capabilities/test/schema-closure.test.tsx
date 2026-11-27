import assert from "node:assert/strict";

import {
  boxCapability,
  chartCapability,
  groupCapability,
  imageCapability,
  motionCapability,
  shapeCapability,
  slideBackgroundCapability,
  tableCapability,
  textCapability,
  transitionCapability,
} from "@pptxdk/capabilities";
import { Box } from "@pptxdk/capabilities/box-elements";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  box,
  cell,
  chart,
  createMotionScopeHandle,
  deck,
  group,
  image,
  motionScope,
  row,
  shape,
  slide,
  step,
  table,
  text,
  transitions,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Image } from "@pptxdk/capabilities/image-elements";
import { breathe, fade, spin } from "@pptxdk/capabilities/motion";
import { MotionScope, Step } from "@pptxdk/capabilities/motion-elements";
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

import {
  boxDeclarationPropsSchema,
  boxStyleSchema,
  type BoxStyle,
} from "../src/box/box-declaration.ts";
import { chartPropsSchema } from "../src/chart/chart-schemas.ts";
import {
  groupDeclarationPropsSchema,
  groupStyleSchema,
} from "../src/group/group-declaration.ts";
import { imageDeclarationPropsSchema } from "../src/image/image-schemas.ts";
import { spinOptionsSchema } from "../src/motion/motion-schemas.ts";
import { notesDeclarationPropsSchema } from "../src/notes/notes-schemas.ts";
import {
  placeholderDeclarationPropsSchema,
  placeholderObjectStyleSchema,
} from "../src/placeholder/placeholder-schemas.ts";
import {
  shapeDeclarationPropsSchema,
  shapeStyleSchema,
  type ShapeStyle,
} from "../src/shape/shape-declaration.ts";
import {
  paragraphAuthorPropsSchema,
  slideNumberAuthorPropsSchema,
  spanAuthorPropsSchema,
  textAuthorPropsSchema,
  textBoxStyleSchema,
  type TextBoxStyle,
} from "../src/shared/text-schemas.ts";
import {
  shapeBorderObjectSchema,
  shapeBorderValueSchema,
} from "../src/shared/visual.ts";
import {
  tableCellStyleSchema,
  tableDeclarationPropsSchema,
  type TableCellStyle,
} from "../src/table/table-schemas.ts";
import {
  transitionKindSchema,
  transitionSchema,
} from "../src/transition/transition-api.ts";

describe("Capability pure-data schema closure: Box, Group, and Shape", () => {
  it("enforces complete pure-data contracts on boxDeclarationPropsSchema, groupDeclarationPropsSchema, and shapeDeclarationPropsSchema", () => {
    const morphScope = createMorphScope();
    const heroKey = morphScope.handle("hero");

    assert.equal(
      boxDeclarationPropsSchema.safeParse({
        id: "b1",
        morphKey: heroKey,
        style: {
          display: "flex",
          padding: 12,
          fontSize: 18,
          background: "#F8FAFC",
          border: { color: "#CBD5E1", width: 1, style: "solid" },
        },
        enter: fade(),
        motion: true,
      }).success,
      true,
    );
    // Rejects unknown/invalid style fields, invalid id, invalid enter, and invalid motion.
    assert.equal(
      boxDeclarationPropsSchema.safeParse({ id: "  " }).success,
      false,
    );
    assert.equal(
      boxStyleSchema.safeParse({ unknownStyleProp: 1 }).success,
      false,
    );
    assert.equal(
      boxStyleSchema.safeParse({ startArrow: "triangle" }).success,
      false,
    );
    assert.equal(
      boxDeclarationPropsSchema.safeParse({ enter: spin() }).success,
      false,
    );
    assert.equal(
      boxDeclarationPropsSchema.safeParse({ motion: "yes" }).success,
      false,
    );

    assert.equal(
      groupDeclarationPropsSchema.safeParse({
        id: "g1",
        name: "HeaderGroup",
        morphKey: heroKey,
        style: { display: "flex", fontSize: 16, rotate: 15 },
        enter: fade(),
        emphasis: spin(),
        motion: true,
      }).success,
      true,
    );
    // Group supports rotate/flipH/flipV but not background or blank name.
    assert.equal(
      groupDeclarationPropsSchema.safeParse({ name: "   " }).success,
      false,
    );
    assert.equal(
      groupStyleSchema.safeParse({ background: "#FFFFFF" }).success,
      false,
    );
    assert.equal(
      groupDeclarationPropsSchema.safeParse({ emphasis: fade() }).success,
      false,
    );

    assert.equal(
      shapeDeclarationPropsSchema.safeParse({
        id: "sh1",
        name: "Callout",
        morphKey: heroKey,
        geometry: "roundRect",
        style: {
          position: "absolute",
          left: 10,
          top: 20,
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
      }).success,
      true,
    );
    assert.equal(
      shapeDeclarationPropsSchema.safeParse({ name: "  " }).success,
      false,
    );
    assert.equal(
      shapeStyleSchema.safeParse({ unknownProp: true }).success,
      false,
    );
    assert.equal(
      shapeDeclarationPropsSchema.safeParse({ enter: "fade" }).success,
      false,
    );
  });

  it("rejects invalid Shape style.width, invalid Box border, and Group background identically across JSX and imperative entries", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      boxCapability,
      groupCapability,
      shapeCapability,
    ]);

    const invalidCases = [
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

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("Capability pure-data schema closure: Text, Image, Table, Chart, Notes, and Font", () => {
  it("enforces complete pure-data contracts on Text, Image, Table, Chart, and Notes declaration schemas", () => {
    const morphScope = createMorphScope();
    const heroKey = morphScope.handle("hero");

    assert.equal(
      textAuthorPropsSchema.safeParse({
        id: "t1",
        name: "Headline",
        morphKey: heroKey,
        style: {
          position: "absolute",
          left: 20,
          top: 30,
          width: 300,
          height: "auto",
          fontSize: "24pt",
          color: "#0F172A",
          padding: 8,
        },
        enter: fade(),
        emphasis: spin(),
        motion: true,
      }).success,
      true,
    );
    assert.equal(textAuthorPropsSchema.safeParse({ id: "" }).success, false);
    assert.equal(
      textAuthorPropsSchema.safeParse({ name: "   " }).success,
      false,
    );
    assert.equal(
      textAuthorPropsSchema.safeParse({ morphKey: "not-a-key" }).success,
      false,
    );
    assert.equal(
      textBoxStyleSchema.safeParse({ startArrow: "triangle" }).success,
      false,
    );
    assert.equal(
      textAuthorPropsSchema.safeParse({ enter: spin() }).success,
      false,
    );
    assert.equal(
      textAuthorPropsSchema.safeParse({ emphasis: fade() }).success,
      false,
    );
    assert.equal(
      textAuthorPropsSchema.safeParse({ motion: "yes" }).success,
      false,
    );
    assert.equal(
      paragraphAuthorPropsSchema.safeParse({ id: "  " }).success,
      false,
    );
    assert.equal(spanAuthorPropsSchema.safeParse({ id: "  " }).success, false);
    assert.equal(
      slideNumberAuthorPropsSchema.safeParse({ id: "  " }).success,
      false,
    );

    assert.equal(
      imageDeclarationPropsSchema.safeParse({
        id: "img1",
        src: png1x1,
        geometry: "roundRect",
        enter: fade(),
      }).success,
      true,
    );
    assert.equal(
      imageDeclarationPropsSchema.safeParse({ id: "" }).success,
      false,
    );
    assert.equal(
      imageDeclarationPropsSchema.safeParse({
        src: png1x1,
        geometry: "not-a-preset",
      }).success,
      false,
    );
    assert.equal(
      imageDeclarationPropsSchema.safeParse({
        src: png1x1,
        enter: spin(),
      }).success,
      false,
    );

    assert.equal(
      tableDeclarationPropsSchema.safeParse({
        columns: [120, "1fr"],
        enter: fade(),
        emphasis: spin(),
        motion: true,
      }).success,
      true,
    );
    assert.equal(
      tableDeclarationPropsSchema.safeParse({
        columns: [120],
        enter: "bad",
      }).success,
      false,
    );

    const chartBase = {
      groups: [
        {
          kind: "column" as const,
          data: { kind: "shared-categories" as const, categories: ["A"] },
          series: [{ name: "S1", values: [1] }],
        },
      ],
    };
    assert.equal(
      chartPropsSchema.safeParse({
        ...chartBase,
        id: "ch1",
        enter: fade(),
        emphasis: spin(),
      }).success,
      true,
    );
    assert.equal(
      chartPropsSchema.safeParse({ ...chartBase, id: "  " }).success,
      false,
    );
    assert.equal(
      chartPropsSchema.safeParse({ ...chartBase, enter: spin() }).success,
      false,
    );

    assert.equal(
      notesDeclarationPropsSchema.safeParse({ id: "  " }).success,
      false,
    );
  });

  it("enforces Table.style schema constraints, Text.width type constraints, and Font constructor validation across public entries", async () => {
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

    const resolve = createResolver([
      slideBackgroundCapability,
      boxCapability,
      textCapability,
      imageCapability,
      tableCapability,
      chartCapability,
    ]);

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
        label: "unknown prop on Image and Chart",
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
  it("validates Transition, MotionScope, Step, and Breathe options consistently across JSX and imperative entries", async () => {
    assert.equal(transitionKindSchema.safeParse("morph").success, true);
    assert.equal(transitionKindSchema.safeParse("blinds").success, false);
    assert.equal(
      transitionSchema.safeParse(
        transitions.morph({ by: "word", fallback: transitions.fade() }),
      ).success,
      true,
    );
    assert.equal(
      transitionSchema.safeParse({
        type: "transition",
        kind: "morph",
        by: "word",
      }).success,
      false,
    );

    const resolve = createResolver([
      slideBackgroundCapability,
      shapeCapability,
      motionCapability,
      transitionCapability,
    ]);
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

  it("enforces updated visual, layout, placeholderSize, spinOptions, and stage error Zod schemas", () => {
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
      shapeBorderObjectSchema.safeParse({
        color: "#1E293B",
        width: 2,
        startArrow: "none",
        endArrow: "triangle",
      }).success,
      true,
    );
    assert.equal(shapeBorderValueSchema.safeParse("none").success, true);

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

    assert.equal(spinOptionsSchema.safeParse({ by: 180 }).success, true);
    assert.equal(
      spinOptionsSchema.safeParse({ from: 0, to: 90 }).success,
      true,
    );
    assert.equal(
      spinOptionsSchema.safeParse({ by: 90, from: 0 }).success,
      false,
    );

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

    const boxStyle: BoxStyle = { padding: pad4 };
    const textStyle: TextBoxStyle = { padding: pad2 };
    const shapeStyle: ShapeStyle = { padding: pad4 };
    const cellStyle: TableCellStyle = { padding: pad4 };
    const cellStyle2: TableCellStyle = { padding: pad2 };

    assert.equal(boxStyleSchema.safeParse(boxStyle).success, true);
    assert.equal(textBoxStyleSchema.safeParse(textStyle).success, true);
    assert.equal(shapeStyleSchema.safeParse(shapeStyle).success, true);
    assert.equal(tableCellStyleSchema.safeParse(cellStyle).success, true);
    assert.equal(tableCellStyleSchema.safeParse(cellStyle2).success, true);
    assert.equal(
      tableCellStyleSchema.safeParse({ padding: [9, -1] }).success,
      false,
    );

    assert.equal(
      placeholderObjectStyleSchema.safeParse({ width: 200, height: 100 })
        .success,
      true,
    );
    assert.equal(
      placeholderDeclarationPropsSchema.safeParse({
        placeholder: "body",
        style: { left: 10, top: 20 },
      }).success,
      true,
    );
    assert.equal(
      placeholderDeclarationPropsSchema.safeParse({ unknownProp: true })
        .success,
      false,
    );
  });
});
