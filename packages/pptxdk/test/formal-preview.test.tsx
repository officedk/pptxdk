import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { BuildError, build, font } from "@officedk/pptxdk";
import {
  PreviewPageRangeError,
  preview as previewDocument,
  preview as previewJsx,
} from "@officedk/pptxdk/preview";
import {
  Box,
  Chart,
  Deck,
  Group,
  Shape,
  Slide,
  Text,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  chart,
  deck,
  group,
  row,
  shape,
  slide,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runIsolatedScript(script: string) {
  const packageRoot = fileURLToPath(new URL("../", import.meta.url));
  const child = spawnSync(
    process.execPath,
    ["--conditions=source", "--input-type=module", "--eval", script],
    { cwd: packageRoot, encoding: "utf8" },
  );
  assert.equal(child.status, 0, child.stderr);
  return child.stdout;
}

describe("Slice 3: formal static preview public entrypoints and runtime isolation", () => {
  it("supports JSX preview and plain document preview from the same formal pipeline", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const jsxResult = await previewJsx(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "cover", background: "#EEEEEE" },
          createElement(Shape, {
            id: "card",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 120,
              height: 60,
              background: "#336699",
            },
          }),
          createElement(
            Text,
            {
              id: "caption",
              style: {
                position: "absolute",
                left: 20,
                top: 96,
                width: 160,
                fontFamily: jostFont,
                fontSize: 18,
                color: "#111827",
              },
            },
            "Shared preview",
          ),
        ),
      ),
    );
    const docResult = await previewDocument(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "cover", background: "#EEEEEE" }).append(
          shape({
            id: "card",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 120,
              height: 60,
              background: "#336699",
            },
          }),
          text("Shared preview", {
            id: "caption",
            style: {
              position: "absolute",
              left: 20,
              top: 96,
              width: 160,
              fontFamily: jostFont,
              fontSize: 18,
              color: "#111827",
            },
          }),
        ),
      ),
    );

    assert.equal(jsxResult.pages.length, 1);
    assert.equal(docResult.pages.length, 1);
    assert.deepEqual(jsxResult.report, docResult.report);
    assert.deepEqual(jsxResult.pages[0]?.errors, []);
    assert.deepEqual(docResult.pages[0]?.errors, []);
    assert.deepEqual(jsxResult.pages[0]?.placeholders, []);
    assert.deepEqual(docResult.pages[0]?.placeholders, []);
    expect(jsxResult.pages[0]?.svg).toContain("<path");
    assert.deepEqual(jsxResult.pages[0]?.png, docResult.pages[0]?.png);
  });

  it("validates 1-based pages, skips drawing unselected pages, and still validates full-deck build", async () => {
    const validTwoPages = deck({ width: 320, height: 180 }).append(
      slide({ id: "unselected-unmigrated" }).append(
        group({ id: "g-unselected" }).append(
          shape({
            id: "s-unselected",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 80,
              height: 40,
              background: "#334155",
            },
          }),
        ),
      ),
      slide({ id: "selected", background: "#123456" }).append(
        shape({
          id: "s",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 80,
            height: 40,
            background: "#FFFFFF",
          },
        }),
      ),
    );

    const selectedOnly = await previewDocument(validTwoPages, { pages: [2] });
    assert.deepEqual(
      selectedOnly.pages.map((p) => p.id),
      ["selected"],
    );
    assert.deepEqual(selectedOnly.pages[0]?.errors, []);
    assert.ok(selectedOnly.pages[0]?.png !== undefined);
    assert.equal(selectedOnly.report.slides.length, 2);

    const noneSelected = await previewDocument(validTwoPages, { pages: [] });
    assert.equal(noneSelected.pages.length, 0);
    assert.equal(noneSelected.report.slides.length, 2);

    for (const invalidPage of [0, -1, 1.5]) {
      await assert.rejects(
        () => previewDocument(validTwoPages, { pages: [invalidPage] }),
        RangeError,
      );
      await assert.rejects(
        () =>
          previewJsx(
            createElement(
              Deck,
              { width: 320, height: 180 },
              createElement(Slide, { id: "one" }),
            ),
            { pages: [invalidPage] },
          ),
        RangeError,
      );
    }
    await assert.rejects(
      () => previewDocument(validTwoPages, { pages: [1, 3] }),
      (error: unknown) =>
        error instanceof RangeError &&
        error instanceof PreviewPageRangeError &&
        error.totalSlides === 2 &&
        error.requestedPages.length === 2 &&
        error.requestedPages[0] === 1 &&
        error.requestedPages[1] === 3,
    );

    // Full-deck formal build validation still catches build errors on unselected pages.
    const invalidUnselected = deck({ width: 320, height: 180 }).append(
      slide({ id: "broken-unselected" }).append(
        shape({
          id: "bad-x",
          style: {
            position: "absolute",
            left: Number.NaN,
            top: 0,
            width: 50,
            height: 50,
          },
        }),
      ),
      slide({ id: "valid-selected" }),
    );
    await assert.rejects(
      () => previewDocument(invalidUnselected, { pages: [2] }),
      BuildError,
    );
  });

  it("rejects missing glyphs on selected pages while rendering all migrated object capabilities without silently omitting them", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "ok-table" }).append(
        table({
          id: "tb1",
          columns: [120],
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 120,
            fontFamily: jostFont,
            fontSize: 12,
          },
        }).append(
          row({ style: { height: 30 } }).append(cell("Migrated table")),
        ),
      ),
      slide({ id: "with-group" }).append(
        group({ id: "g1" }).append(
          shape({
            id: "m1",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 40,
              height: 40,
              background: "#FF0000",
            },
          }),
        ),
      ),
      slide({ id: "with-missing-glyph" }).append(
        text("汉", {
          id: "t1",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 200,
            height: 60,
            fontFamily: jostFont,
            fontSize: 16,
          },
        }),
      ),
      slide({ id: "ok-box" }).append(
        box({
          id: "b1",
          style: {
            width: 120,
            height: 60,
            padding: 8,
            background: "#E2E8F0",
          },
        }).append(
          shape({
            id: "ok",
            style: {
              width: 60,
              height: 40,
              background: "#00AA00",
            },
          }),
        ),
      ),
      slide({ id: "ok-column-chart" }).append(
        chart({
          id: "col1",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 280,
            height: 150,
            fontFamily: jostFont,
            fontSize: 12,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [3, 5] }],
            },
          ],
        }),
      ),
    );
    const result = await previewDocument(doc);
    expect(result.pages[0]?.png).toBeDefined();
    expect(result.pages[0]?.errors).toEqual([]);
    expect(result.pages[1]?.png).toBeDefined();
    expect(result.pages[1]?.errors).toEqual([]);
    expect(result.pages[2]?.png).toBeUndefined();
    expect(result.pages[2]?.svg).toBeUndefined();
    expect(result.pages[2]?.errors[0]?.code).toBe("preview-missing-glyph");
    expect(result.pages[2]?.errors[0]?.message).toMatch(/visible glyph/i);
    expect(result.pages[2]?.errors[0]?.path).toContain("text[t1]");
    expect(result.pages[3]?.png).toBeDefined();
    expect(result.pages[3]?.errors).toEqual([]);
    expect(result.pages[4]?.png).toBeDefined();
    expect(result.pages[4]?.errors).toEqual([]);

    const jsxResult = await previewJsx(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "with-group" },
          createElement(
            Group,
            { id: "g1", style: { rotate: 90 } },
            createElement(Shape, {
              id: "m1",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 60,
                height: 20,
                background: "#FF0000",
              },
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "with-jsx-box" },
          createElement(
            Box,
            { id: "jb1", style: { width: 80, height: 40, padding: 4 } },
            createElement(Shape, {
              style: { width: 40, height: 20, background: "#0000FF" },
            }),
          ),
        ),
      ),
    );
    expect(jsxResult.pages[0]?.png).toBeDefined();
    expect(jsxResult.pages[0]?.errors).toEqual([]);
    const rotatedGroupNode = jsxResult.pages[0]?.nodes.find(
      (node) => node.authorId === "g1",
    );
    expect(rotatedGroupNode?.bounds?.width).toBeCloseTo(20);
    expect(rotatedGroupNode?.bounds?.height).toBeCloseTo(60);
    expect(jsxResult.pages[1]?.png).toBeDefined();
    expect(jsxResult.pages[1]?.errors).toEqual([]);
  });

  it("computes exact screen corners and bounds for rotated groups, offset children, rotated/flipped members, and nested groups", async () => {
    const result = await previewJsx(
      createElement(
        Deck,
        { width: 480, height: 320 },
        createElement(
          Slide,
          { id: "rotated-group-members" },
          createElement(
            Group,
            { id: "outer-group", style: { rotate: 90 } },
            createElement(Shape, {
              id: "offset-rect",
              style: {
                position: "absolute",
                left: 100,
                top: 100,
                width: 60,
                height: 20,
                background: "#2563EB",
              },
            }),
            createElement(Shape, {
              id: "rotated-child",
              style: {
                position: "absolute",
                left: 180,
                top: 120,
                width: 40,
                height: 40,
                background: "#16A34A",
                rotate: 45,
              },
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "nested-rotated-groups" },
          createElement(
            Group,
            { id: "parent-group", style: { rotate: 90 } },
            createElement(
              Group,
              { id: "child-group", style: { rotate: 90 } },
              createElement(Shape, {
                id: "leaf-rect",
                style: {
                  position: "absolute",
                  left: 100,
                  top: 100,
                  width: 60,
                  height: 20,
                  background: "#DC2626",
                },
              }),
            ),
            createElement(Shape, {
              id: "anchor-rect",
              style: {
                position: "absolute",
                left: 200,
                top: 140,
                width: 20,
                height: 20,
                background: "#0EA5E9",
              },
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "arbitrary-angle-and-flip" },
          createElement(
            Group,
            { id: "angled-group", style: { rotate: 30 } },
            createElement(Shape, {
              id: "flipped-child",
              style: {
                position: "absolute",
                left: 100,
                top: 80,
                width: 80,
                height: 40,
                flipH: true,
                background: "#9333EA",
                rotate: 60,
              },
            }),
            createElement(Shape, {
              id: "second-member",
              style: {
                position: "absolute",
                left: 220,
                top: 120,
                width: 40,
                height: 40,
                background: "#F59E0B",
              },
            }),
          ),
        ),
      ),
    );

    const assertCornersClose = (
      actual: readonly { readonly x: number; readonly y: number }[] | undefined,
      expected: readonly { readonly x: number; readonly y: number }[],
    ) => {
      assert.equal(actual?.length, expected.length);
      expected.forEach((point, index) => {
        expect(actual[index]!.x).toBeCloseTo(point.x, 6);
        expect(actual[index]!.y).toBeCloseTo(point.y, 6);
      });
    };

    const assertBoundsClose = (
      actual:
        | {
            readonly x: number;
            readonly y: number;
            readonly width: number;
            readonly height: number;
          }
        | undefined,
      expected: {
        readonly x: number;
        readonly y: number;
        readonly width: number;
        readonly height: number;
      },
    ) => {
      assert.ok(actual !== undefined);
      expect(actual.x).toBeCloseTo(expected.x, 6);
      expect(actual.y).toBeCloseTo(expected.y, 6);
      expect(actual.width).toBeCloseTo(expected.width, 6);
      expect(actual.height).toBeCloseTo(expected.height, 6);
    };

    // Page 1: 90° rotated group [100..220, 100..160] (center 160, 130)
    const p1 = result.pages[0]!;
    assert.deepEqual(p1.errors, []);
    assert.ok(p1.png !== undefined);

    const outerGroup = p1.nodes.find((n) => n.authorId === "outer-group");
    assertCornersClose(outerGroup?.corners, [
      { x: 190, y: 70 },
      { x: 190, y: 190 },
      { x: 130, y: 190 },
      { x: 130, y: 70 },
    ]);
    assertBoundsClose(outerGroup?.bounds, {
      x: 130,
      y: 70,
      width: 60,
      height: 120,
    });

    const offsetRect = p1.nodes.find((n) => n.authorId === "offset-rect");
    assertCornersClose(offsetRect?.corners, [
      { x: 190, y: 70 },
      { x: 190, y: 130 },
      { x: 170, y: 130 },
      { x: 170, y: 70 },
    ]);
    assertBoundsClose(offsetRect?.bounds, {
      x: 170,
      y: 70,
      width: 20,
      height: 60,
    });

    const halfDiag = 20 * Math.SQRT2;
    const rotatedChild = p1.nodes.find((n) => n.authorId === "rotated-child");
    assertCornersClose(rotatedChild?.corners, [
      { x: 150 + halfDiag, y: 170 },
      { x: 150, y: 170 + halfDiag },
      { x: 150 - halfDiag, y: 170 },
      { x: 150, y: 170 - halfDiag },
    ]);
    assertBoundsClose(rotatedChild?.bounds, {
      x: 150 - halfDiag,
      y: 170 - halfDiag,
      width: 2 * halfDiag,
      height: 2 * halfDiag,
    });

    // Page 2: nested 90° groups with distinct centers ((130, 110) inside (160, 130))
    const p2 = result.pages[1]!;
    assert.deepEqual(p2.errors, []);
    assert.ok(p2.png !== undefined);

    const childGroup = p2.nodes.find((n) => n.authorId === "child-group");
    const leafRect = p2.nodes.find((n) => n.authorId === "leaf-rect");
    const expectedNestedCorners = [
      { x: 210, y: 110 },
      { x: 150, y: 110 },
      { x: 150, y: 90 },
      { x: 210, y: 90 },
    ];
    const expectedNestedBox = { x: 150, y: 90, width: 60, height: 20 };
    assertCornersClose(childGroup?.corners, expectedNestedCorners);
    assertBoundsClose(childGroup?.bounds, expectedNestedBox);
    assertCornersClose(leafRect?.corners, expectedNestedCorners);
    assertBoundsClose(leafRect?.bounds, expectedNestedBox);

    // Page 3: 30° group [100..260, 80..160] (center 180, 120) with flipH + 60° child [100..180, 80..120] (center 140, 100)
    const p3 = result.pages[2]!;
    assert.deepEqual(p3.errors, []);
    assert.ok(p3.png !== undefined);

    const flippedChild = p3.nodes.find((n) => n.authorId === "flipped-child");
    // Local flipH swaps left/right around (140, 100), then +60° + 30° = +90° total rotation,
    // while child center (140, 100) rotates +30° around group center (180, 120):
    // dc = (-40, -20) -> (180 - 40*cos30 + 20*sin30, 120 - 40*sin30 - 20*cos30)
    const cos30 = Math.cos(Math.PI / 6);
    const sin30 = 0.5;
    const childCenterX = 180 + (-40 * cos30 - -20 * sin30);
    const childCenterY = 120 + (-40 * sin30 + -20 * cos30);
    // Original corners after flipH around (140, 100): (+40,-20), (-40,-20), (-40,+20), (+40,+20)
    // Rotated by 90° CW: (dx, dy) -> (-dy, dx): (+20,+40), (+20,-40), (-20,-40), (-20,+40)
    assertCornersClose(flippedChild?.corners, [
      { x: childCenterX + 20, y: childCenterY + 40 },
      { x: childCenterX + 20, y: childCenterY - 40 },
      { x: childCenterX - 20, y: childCenterY - 40 },
      { x: childCenterX - 20, y: childCenterY + 40 },
    ]);
    assertBoundsClose(flippedChild?.bounds, {
      x: childCenterX - 20,
      y: childCenterY - 40,
      width: 40,
      height: 80,
    });
  });

  it("never loads @resvg/resvg-js, pngjs, or jpeg-js when calling JSX build, document build, or resolve + emit with PNG and JPEG images", () => {
    const pngBase64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=";
    const jpegBase64 =
      "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAFR/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIRAxEAPwD/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2Q==";
    const jsxOut = runIsolatedScript(`
import { registerHooks } from "node:module";
import { createElement } from "react";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("JSX build loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build } = await import("@officedk/pptxdk");
const { Deck, Image, Slide, Shape } = await import("@officedk/pptxdk/react");
const pngBytes = new Uint8Array(Buffer.from("${pngBase64}", "base64"));
const jpegBytes = new Uint8Array(Buffer.from("${jpegBase64}", "base64"));
const result = await build(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "s1" },
      createElement(Shape, { style: { position: "absolute", left: 10, top: 10, width: 40, height: 40 } }),
      createElement(Image, { style: { position: "absolute", left: 60, top: 10, width: 40, height: 40 }, src: pngBytes, fit: "cover" }),
      createElement(Image, { style: { position: "absolute", left: 110, top: 10, width: 40, height: 40, flipH: true, rotate: 15, opacity: 0.8 }, src: jpegBytes, crop: { left: 0.2 } }),
    ),
  ),
);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output");
}
process.stdout.write("ok");
`);
    assert.equal(jsxOut, "ok");

    const docBuildOut = runIsolatedScript(`
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("Document build loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build } = await import("@officedk/pptxdk");
const { deck, image, slide, shape } = await import("@officedk/pptxdk/vanilla");
const { createResolver, emit } = await import("@officedk/pptxdk/build");
void createResolver;
void emit;
const pngBytes = new Uint8Array(Buffer.from("${pngBase64}", "base64"));
const jpegBytes = new Uint8Array(Buffer.from("${jpegBase64}", "base64"));
const result = await build(
  deck({ width: 320, height: 180 }).append(
    slide({ id: "s1" }).append(
      shape({ style: { position: "absolute", left: 10, top: 10, width: 40, height: 40 } }),
      image({ style: { position: "absolute", left: 60, top: 10, width: 40, height: 40 }, src: pngBytes, fit: "contain" }),
      image({ style: { position: "absolute", left: 110, top: 10, width: 40, height: 40, flipV: true, rotate: 30 }, src: jpegBytes }),
    ),
  ),
);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output");
}
process.stdout.write("ok");
`);
    assert.equal(docBuildOut, "ok");

    const resolveEmitOut = runIsolatedScript(`
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("resolve + emit loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { imageCapability } = await import("@pptxdk/capabilities/image");
const { createResolver } = await import("@pptxdk/core/build");
const { createDeckNode, createSlideNode } = await import("@pptxdk/core/document");
const { emit } = await import("@pptxdk/emit");
const { createImageNode } = await import("@pptxdk/capabilities/image");

const pngBytes = new Uint8Array(Buffer.from("${pngBase64}", "base64"));
const jpegBytes = new Uint8Array(Buffer.from("${jpegBase64}", "base64"));
const d = createDeckNode({ width: 320, height: 180 });
const s = createSlideNode({ id: "s1" });
s.append(
  createImageNode({ style: { position: "absolute", left: 10, top: 10, width: 80, height: 40 }, src: pngBytes, fit: "contain" }),
  createImageNode({ style: { position: "absolute", left: 100, top: 10, width: 80, height: 40, flipH: true }, src: jpegBytes, crop: { top: 0.25 } }),
);
d.append(s);

const resolution = createResolver([imageCapability])({ children: [d] });
const emitted = emit(resolution);
if (!(emitted.pptx instanceof Uint8Array) || emitted.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output from resolve + emit");
}
process.stdout.write("ok");
`);
    assert.equal(resolveEmitOut, "ok");
  });

  it("never loads react or react-reconciler when calling plain document preview", () => {
    const out = runIsolatedScript(`
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("Plain document preview loaded React or test dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { deck, slide, shape } = await import("@officedk/pptxdk/vanilla");
const { preview } = await import("@officedk/pptxdk/preview");
const result = await preview(
  deck({ width: 320, height: 180 }).append(
    slide({ id: "s1", background: "#112233" }).append(
      shape({ style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#FF0000" } }),
    ),
  ),
);
if (!result.pages[0]?.png || result.pages[0].png.byteLength === 0) {
  throw new Error("Expected preview PNG");
}
process.stdout.write("ok");
`);
    assert.equal(out, "ok");
  });

  it("never loads test scenarios, vitest, or validation from production preview entrypoints", () => {
    const out = runIsolatedScript(`
import { registerHooks } from "node:module";
import { createElement } from "react";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@pptxdk/validation") ||
      specifier === "vitest" ||
      specifier.includes("/capabilities/test/")
    ) {
      throw new Error("Production entrypoint loaded test/validation dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { Deck, Slide, Shape } = await import("@officedk/pptxdk/react");
const { preview } = await import("@officedk/pptxdk/preview");
const result = await preview(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "s1" },
      createElement(Shape, { style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#2563EB" } }),
    ),
  ),
);
if (!result.pages[0]?.png || result.pages[0].png.byteLength === 0) {
  throw new Error("Expected preview PNG");
}
process.stdout.write("ok");
`);
    assert.equal(out, "ok");
  });

  it("renders dual-axis rotated-title combo and decimal-scale area charts without plotArea clipping across JSX and Vanilla preview", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const chartStyle = {
      position: "absolute" as const,
      left: 20,
      top: 20,
      width: 560,
      height: 300,
      fontFamily: jostFont,
      fontSize: 13,
    };
    const comboAxes = {
      quarter: { kind: "category" as const, title: "Quarter" },
      primary: {
        kind: "value" as const,
        position: "left" as const,
        title: "Revenue ($M)",
        min: 0,
        max: 1200,
        majorUnit: 300,
        numberFormat: "#,##0",
      },
      secondary: {
        kind: "value" as const,
        position: "right" as const,
        title: "Margin (%)",
        min: 0,
        max: 1.2,
        majorUnit: 0.3,
        numberFormat: "0%",
      },
    };
    const comboGroups = [
      {
        kind: "column" as const,
        axes: { category: "quarter", value: "primary" },
        data: { kind: "shared-categories" as const, categories: ["Q1", "Q2"] },
        series: [{ name: "Revenue", values: [400, 1050] }],
      },
      {
        kind: "line" as const,
        axes: { category: "quarter", value: "secondary" },
        data: { kind: "shared-categories" as const, categories: ["Q1", "Q2"] },
        series: [{ name: "Margin", values: [0.35, 1.05] }],
      },
    ];
    const decimalAxes = {
      x: { kind: "category" as const, title: "Batch" },
      y: {
        kind: "value" as const,
        position: "left" as const,
        title: "Density (kg/L)",
        min: 0,
        max: 3.5,
        majorUnit: 0.5,
        numberFormat: "0.0",
      },
    };
    const decimalGroups = [
      {
        kind: "area" as const,
        axes: { category: "x", value: "y" },
        data: {
          kind: "shared-categories" as const,
          categories: ["B1", "B2", "B3"],
        },
        series: [{ name: "Density", values: [0.5, 2.2, 3.4] }],
      },
    ];

    const vanillaDeck = deck({ width: 640, height: 360 }).append(
      slide({ id: "combo", background: "#FFFFFF" }).append(
        chart({
          id: "combo-chart",
          style: chartStyle,
          axes: comboAxes,
          groups: comboGroups,
        }),
      ),
      slide({ id: "area-decimal", background: "#FFFFFF" }).append(
        chart({
          id: "decimal-chart",
          style: chartStyle,
          axes: decimalAxes,
          groups: decimalGroups,
        }),
      ),
    );
    const jsxNode = createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "combo", background: "#FFFFFF" },
        createElement(Chart, {
          id: "combo-chart",
          style: chartStyle,
          axes: comboAxes,
          groups: comboGroups,
        }),
      ),
      createElement(
        Slide,
        { id: "area-decimal", background: "#FFFFFF" },
        createElement(Chart, {
          id: "decimal-chart",
          style: chartStyle,
          axes: decimalAxes,
          groups: decimalGroups,
        }),
      ),
    );

    const [vanillaPreview, jsxPreview, built] = await Promise.all([
      previewDocument(vanillaDeck),
      previewJsx(jsxNode),
      build(vanillaDeck),
    ]);

    const files = unzipSync(built.pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).not.toContain(
      "<c:manualLayout>",
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).not.toContain(
      "<c:manualLayout>",
    );

    for (const pageIndex of [0, 1] as const) {
      const vPage = vanillaPreview.pages[pageIndex]!;
      const jPage = jsxPreview.pages[pageIndex]!;
      assert.deepEqual(vPage.errors, []);
      assert.deepEqual(jPage.errors, []);
      assert.ok(vPage.png !== undefined);
      assert.ok(vPage.svg !== undefined);
      assert.deepEqual(jPage.png, vPage.png);
      expect(vPage.svg.match(/<clipPath\b/gu) ?? []).toHaveLength(2);
      const png = PNG.sync.read(Buffer.from(vPage.png));
      for (let y = 40; y < 280; y += 1) {
        const leftIdx = (y * png.width + 20) * 4;
        expect(png.data[leftIdx]).toBe(255);
      }
    }
  });
});
