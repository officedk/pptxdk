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
import { BuildError, createMorphScope } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  boxDeclarationPropsSchema,
  boxStyleSchema,
} from "../src/box/box-declaration.ts";
import { ChartPropsSchema } from "../src/chart/chart-schemas.ts";
import {
  groupDeclarationPropsSchema,
  groupStyleSchema,
} from "../src/group/group-declaration.ts";
import { imageDeclarationPropsSchema } from "../src/image/image-schemas.ts";
import { notesDeclarationPropsSchema } from "../src/notes/notes-schemas.ts";
import {
  shapeDeclarationPropsSchema,
  shapeStyleSchema,
} from "../src/shape/shape-declaration.ts";
import {
  paragraphAuthorPropsSchema,
  slideNumberAuthorPropsSchema,
  spanAuthorPropsSchema,
  textAuthorPropsSchema,
  textBoxStyleSchema,
} from "../src/shared/text-schemas.ts";
import { tableDeclarationPropsSchema } from "../src/table/table-schemas.ts";
import {
  TransitionKindSchema,
  TransitionSchema,
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
          backgroundColor: "#F8FAFC",
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
    // Group supports rotate but not backgroundColor or blank name.
    assert.equal(
      groupDeclarationPropsSchema.safeParse({ name: "   " }).success,
      false,
    );
    assert.equal(
      groupStyleSchema.safeParse({ backgroundColor: "#FFFFFF" }).success,
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
        x: 10,
        y: 20,
        width: 120,
        height: 60,
        style: {
          backgroundColor: "#2563EB",
          borderColor: "#1D4ED8",
          borderWidth: 2,
          startArrow: "none",
          endArrow: "triangle",
          fontSize: "16pt",
        },
        text: "Hello",
        textStyle: { fontWeight: "bold" },
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

  it("rejects invalid Shape width/height and non-object Shape.textStyle identically across JSX and imperative entries", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      boxCapability,
      groupCapability,
      shapeCapability,
    ]);

    const invalidCases = [
      {
        label: "string width on Shape inside Box",
        expectedCode: "invalid-coordinate",
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
                  width: "200" as never,
                  height: 40,
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
                  width: "200" as never,
                  height: 40,
                }),
              ),
            ),
          ),
      },
      {
        label: "non-object textStyle on Shape",
        expectedCode: "invalid-style",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "bad-ts",
                x: 10,
                y: 10,
                width: 100,
                height: 40,
                text: "Hi",
                textStyle: "bold" as never,
              }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "bad-ts",
                x: 10,
                y: 10,
                width: 100,
                height: 40,
                text: "Hi",
                textStyle: "bold" as never,
              }),
            ),
          ),
      },
      {
        label: "backgroundColor on Group",
        expectedCode: "unknown-style",
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
                  style: { backgroundColor: "#FFFFFF" } as never,
                },
                createElement(Shape, {
                  id: "m1",
                  x: 10,
                  y: 10,
                  width: 40,
                  height: 40,
                }),
              ),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              group({
                id: "bad-g",
                style: { backgroundColor: "#FFFFFF" } as never,
              }).append(
                shape({ id: "m1", x: 10, y: 10, width: 40, height: 40 }),
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
        x: 20,
        y: 30,
        width: 300,
        height: "auto",
        style: { fontSize: "24pt", color: "#0F172A", padding: 8 },
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
      ChartPropsSchema.safeParse({
        ...chartBase,
        id: "ch1",
        enter: fade(),
        emphasis: spin(),
      }).success,
      true,
    );
    assert.equal(
      ChartPropsSchema.safeParse({ ...chartBase, id: "  " }).success,
      false,
    );
    assert.equal(
      ChartPropsSchema.safeParse({ ...chartBase, enter: spin() }).success,
      false,
    );

    assert.equal(
      notesDeclarationPropsSchema.safeParse({ id: "  " }).success,
      false,
    );
  });

  it("enforces Table.style schema constraints, Text.width type constraints, and Font constructor validation across public entries", async () => {
    assert.throws(() => font("   "), TypeError);
    assert.throws(
      () => font("Jost*", [{ weight: 450, data: png1x1 }]),
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
        label: "string width on Text inside Box",
        expectedCode: "invalid-text-box",
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
                  { id: "bad-tw", width: "200" as never, height: 40 },
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
                  width: "200" as never,
                  height: 40,
                }),
              ),
            ),
          ),
      },
      {
        label: "unsupported align on Table.style",
        expectedCode: "unknown-style",
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
        expectedCode: "unknown-parameter",
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
        expectedCode: "unknown-parameter",
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
        expectedCode: "unknown-parameter",
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
    assert.equal(TransitionKindSchema.safeParse("morph").success, true);
    assert.equal(TransitionKindSchema.safeParse("blinds").success, false);
    assert.equal(
      TransitionSchema.safeParse(
        transitions.morph({ by: "word", fallback: transitions.fade() }),
      ).success,
      true,
    );
    assert.equal(
      TransitionSchema.safeParse({
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
        expectedCode: "unknown-parameter",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "sh",
                x: 10,
                y: 10,
                width: 60,
                height: 40,
                enter: fade({ easing: { accelerate: 0.2, extra: 1 } as never }),
              }),
            ),
          ),
        vanilla: () =>
          deck({ width: 480, height: 270 }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "sh",
                x: 10,
                y: 10,
                width: 60,
                height: 40,
                enter: fade({ easing: { accelerate: 0.2, extra: 1 } as never }),
              }),
            ),
          ),
      },
      {
        label: "unknown breathe parameter",
        expectedCode: "unknown-parameter",
        jsx: () =>
          createElement(
            Deck,
            { width: 480, height: 270 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "sh",
                x: 10,
                y: 10,
                width: 60,
                height: 40,
                fill: "#112233",
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
                x: 10,
                y: 10,
                width: 60,
                height: 40,
                fill: "#112233",
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
        expectedCode: "unknown-parameter",
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
                  x: 10,
                  y: 10,
                  width: 60,
                  height: 40,
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
                  x: 10,
                  y: 10,
                  width: 60,
                  height: 40,
                  enter: fade(),
                }),
              ),
            ),
          ),
      },
      {
        label: "unknown Step parameter",
        expectedCode: "unknown-parameter",
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
                  x: 10,
                  y: 10,
                  width: 60,
                  height: 40,
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
                  x: 10,
                  y: 10,
                  width: 60,
                  height: 40,
                  enter: fade(),
                }),
              ),
            ),
          ),
      },
      {
        label: "unsupported transition kind",
        expectedCode: "unsupported-transition",
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
});
