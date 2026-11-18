import assert from "node:assert/strict";

import { Deck, Slide, expandDocument } from "@react-pptx/core";
import { createDocumentPreview } from "@react-pptx/core/preview";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import {
  deck,
  documentCapabilities,
  image,
  notes,
  notesCapability,
  paragraph,
  slide,
  shape,
} from "../src/document.ts";
import { blockArc, path, preset } from "../src/geometry.ts";
import { Image } from "../src/image-elements.ts";
import { motionCapability } from "../src/motion.ts";
import { Shape } from "../src/shape-elements.ts";
import { shapeCapability } from "../src/shape.ts";
import { slideBackgroundCapability } from "../src/slide-background.ts";
import { P } from "../src/text-elements.ts";
import { transitionCapability, transitions } from "../src/transition.ts";
import { linearGradient } from "../src/visual.ts";

describe("Slice 1: core bound slide preview & slideBackgroundCapability", () => {
  it("renders default white and solid backgrounds across both authoring entries using bound slide role", async () => {
    const previewDoc = createDocumentPreview([
      slideBackgroundCapability,
      shapeCapability,
      notesCapability,
      transitionCapability,
      motionCapability,
    ]);

    const jsxRoot = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "blank" }),
        createElement(
          Slide,
          {
            id: "solid",
            background: "#123456",
            transition: transitions.fade(),
          },
          createElement(Shape, {
            id: "dummy-ignored-for-bg-check",
            x: 0,
            y: 0,
            width: 10,
            height: 10,
          }),
        ),
      ),
    );
    // Remove dummy shape for pure background slide test in Slice 1 if needed, or test blank/solid slides only:
    const pureJsx = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "blank" }),
        createElement(Slide, {
          id: "solid",
          background: "#123456",
          transition: transitions.fade(),
        }),
      ),
    );
    void jsxRoot;

    const impDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "blank" }),
      slide({
        id: "solid",
        background: "#123456",
        transition: transitions.fade(),
      }).append(notes("Speaker note stays off canvas")),
    );
    void shape;

    for (const [label, doc] of [
      ["jsx", pureJsx],
      ["document", impDoc],
    ] as const) {
      const result = await previewDoc(doc);
      assert.equal(result.pages.length, 2, label);
      const validation = await validatePptx(result.build.pptx);
      assert.deepEqual(validation.errors, [], label);

      const blankPage = result.pages[0]!;
      assert.equal(blankPage.id, "blank");
      assert.equal(blankPage.page, 1);
      assert.deepEqual(blankPage.diagnostics, []);
      assert.ok(blankPage.png !== undefined);
      const blankPng = PNG.sync.read(Buffer.from(blankPage.png));
      assert.equal(blankPng.width, 320);
      assert.equal(blankPng.height, 180);
      assert.deepEqual([...blankPng.data.subarray(0, 4)], [255, 255, 255, 255]);

      const solidPage = result.pages[1]!;
      assert.equal(solidPage.id, "solid");
      assert.equal(solidPage.page, 2);
      assert.deepEqual(solidPage.diagnostics, []);
      assert.ok(solidPage.png !== undefined);
      const solidPng = PNG.sync.read(Buffer.from(solidPage.png));
      assert.deepEqual(
        [...solidPng.data.subarray(0, 4)],
        [0x12, 0x34, 0x56, 255],
      );
    }
  });

  it("rejects gradient background on its page with located error while keeping valid pages and PPTX", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const doc = deck({ width: 320, height: 180 }).append(
      slide({
        id: "grad",
        background: linearGradient(45, "#112233", "#445566"),
      }),
      slide({ id: "ok", background: "#ABCDEF" }),
    );
    const result = await previewDoc(doc);
    const validation = await validatePptx(result.build.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(result.pages[0]?.png, undefined);
    assert.equal(result.pages[0]?.diagnostics.length, 1);
    assert.equal(result.pages[0]?.diagnostics[0]?.severity, "error");
    assert.equal(result.pages[0]?.diagnostics[0]?.page, 1);
    assert.match(
      result.pages[0]?.diagnostics[0]?.path ?? "",
      /^deck\/slide\[grad\]/,
    );
    assert.ok(result.pages[1]?.png !== undefined);
  });
});

describe("Slice 2: Shape static preview", () => {
  it("renders rect, ellipse, preset, custom path, transforms, zIndex, and screenGeometry", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const custom = {
      type: "custom" as const,
      width: 10,
      height: 10,
      paths: [
        {
          commands: [
            { op: "M" as const, x: 0, y: 0 },
            { op: "L" as const, x: 10, y: 0 },
            { op: "L" as const, x: 10, y: 10 },
            { op: "L" as const, x: 0, y: 10 },
            { op: "Z" as const },
          ],
          fill: true,
          stroke: false,
        },
        {
          commands: [
            { op: "M" as const, x: 0, y: 0 },
            { op: "Q" as const, x1: 5, y1: 10, x: 10, y: 0 },
            { op: "C" as const, x1: 8, y1: 4, x2: 2, y2: 4, x: 0, y: 0 },
          ],
          fill: false,
          stroke: true,
        },
      ],
    };
    const mutableShape = shape({
      id: "hero-rect",
      x: 20,
      y: 20,
      width: 100,
      height: 50,
      fill: "#FF0000",
      style: {
        borderColor: "#000000",
        borderWidth: 2,
        borderStyle: "solid",
        rotate: 90,
        zIndex: 2,
      },
    });
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "shapes" }).append(
        mutableShape,
        shape({
          id: "underlay",
          geometry: preset("ellipse"),
          x: 20,
          y: 20,
          width: 100,
          height: 50,
          style: { backgroundColor: "#0000FF", zIndex: 1 },
        }),
        shape({
          id: "custom-path",
          geometry: custom,
          x: 160,
          y: 20,
          width: 100,
          height: 50,
          fill: "accent1",
          flipH: true,
          flipV: true,
          style: {
            borderColor: "#112233",
            borderWidth: 2,
            opacity: 0.5,
          },
        }),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // zIndex 1 underlay must appear before zIndex 2 hero-rect in SVG paint order.
    expect(page.svg.indexOf('fill="#0000FF"')).toBeLessThan(
      page.svg.indexOf('fill="#FF0000"'),
    );
    expect(page.svg).toContain("<ellipse");
    expect(page.svg).toContain('fill="none"');
    expect(page.svg).toContain('stroke="none"');
    expect(page.svg).toContain('stroke-width="2"');

    const heroNode = page.nodes.find((node) => node.authorId === "hero-rect")!;
    expect(heroNode.screenBox?.x).toBeCloseTo(45);
    expect(heroNode.screenBox?.y).toBeCloseTo(-5);
    expect(heroNode.screenBox?.width).toBeCloseTo(50);
    expect(heroNode.screenBox?.height).toBeCloseTo(100);

    // Mutate after preview and rebuild: first result stays untouched.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    mutableShape.set({ fill: "#00AA00" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.notEqual(second.pages[0]?.svg, firstSvg);
    expect(second.pages[0]?.svg).toContain('fill="#00AA00"');
  });

  it("rejects unsupported shape styles, arrows, inline text, and geometries with located errors", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const cases: readonly {
      label: string;
      node: ReturnType<typeof shape>;
      expected: RegExp;
    }[] = [
      {
        label: "startArrow",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          style: { borderColor: "#000000", startArrow: "triangle" },
        }),
        expected: /startArrow/i,
      },
      {
        label: "endArrow",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          style: { borderColor: "#000000", endArrow: "arrow" },
        }),
        expected: /endArrow/i,
      },
      {
        label: "dashed",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          style: { borderColor: "#000000", borderStyle: "dashed" },
        }),
        expected: /border|dash/i,
      },
      {
        label: "borderRadius",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
          style: { borderRadius: 8 },
        }),
        expected: /borderRadius/i,
      },
      {
        label: "boxShadow",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        expected: /boxShadow/i,
      },
      {
        label: "softEdge",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
          style: { softEdge: 4 },
        }),
        expected: /softEdge/i,
      },
      {
        label: "gradient",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          style: { backgroundImage: linearGradient(90, "#111111", "#999999") },
        }),
        expected: /backgroundImage|gradient/i,
      },
      {
        label: "inline-text",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
          text: "Inline",
        }),
        expected: /text/i,
      },
      {
        label: "rich-children-text",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
        }).append(paragraph("Child")),
        expected: /text/i,
      },
      {
        label: "star5-preset",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 60,
          geometry: "star5",
          fill: "#112233",
        }),
        expected: /geometry/i,
      },
      {
        label: "adjusted-preset",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 60,
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          fill: "#112233",
        }),
        expected: /geometry/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "s1" }).append(item.node),
        ),
      );
      const validation = await validatePptx(result.build.pptx);
      assert.deepEqual(validation.errors, [], item.label);
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /shape\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
    }

    // Verify JSX shape with children text also rejects cleanly.
    const jsxWithText = await expandDocument(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "jsx-text" },
          createElement(
            Shape,
            { id: "bad-jsx", x: 10, y: 10, width: 60, height: 20 },
            createElement(P, null, "Hello"),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxWithText);
    assert.equal(jsxResult.pages[0]?.png, undefined);
    assert.match(
      jsxResult.pages[0]?.diagnostics[0]?.path ?? "",
      /shape\[bad-jsx\]/,
    );
  });
});

function createTwoToneSource(): PNG {
  // Left half (x < 10) is pure red, right half (x >= 10) is pure blue.
  const source = new PNG({ width: 20, height: 10 });
  for (let y = 0; y < 10; y++) {
    for (let x = 0; x < 20; x++) {
      source.data.set(
        x < 10 ? [255, 0, 0, 255] : [0, 0, 255, 255],
        (y * 20 + x) * 4,
      );
    }
  }
  return source;
}

describe("Slice 3 (Ticket 04): Image static preview positive behaviors", () => {
  it("renders PNG and JPEG with shared imagePlacement (contain, cover, explicit crop), transforms, opacity, and screenGeometry", async () => {
    const { encode } = await import("jpeg-js");
    const source = createTwoToneSource();
    const pngBytes = new Uint8Array(PNG.sync.write(source));
    const jpegBytes = new Uint8Array(encode(source, 100).data);

    const previewDoc = createDocumentPreview(documentCapabilities);

    for (const [format, bytes] of [
      ["png", pngBytes],
      ["jpeg", jpegBytes],
    ] as const) {
      const mutable = image({
        id: "cropped",
        src: bytes,
        x: 10,
        y: 10,
        width: 40,
        height: 40,
        crop: { left: 0.5 },
      });
      const doc = deck({ width: 320, height: 180 }).append(
        slide({ id: `img-${format}`, background: "#FFFFFF" }).append(
          mutable,
          image({
            id: "contained",
            src: bytes,
            x: 80,
            y: 10,
            width: 80,
            height: 80,
            fit: "contain",
            style: { rotate: 90 },
          }),
          image({
            id: "covered",
            src: bytes,
            x: 180,
            y: 10,
            width: 40,
            height: 80,
            fit: "cover",
            flipH: true,
            style: { opacity: 0.5 },
          }),
        ),
      );

      const first = await previewDoc(doc);
      const page = first.pages[0]!;
      assert.deepEqual(page.diagnostics, [], format);
      assert.ok(page.png !== undefined, format);
      assert.ok(page.svg !== undefined, format);
      expect(page.svg).toContain("<clipPath");
      expect(page.svg).toContain("<image");

      // Explicit crop: left: 0.5 removes red half, center of (10..50, 10..50) = (30, 30) must be blue.
      const rendered = PNG.sync.read(Buffer.from(page.png));
      const cropCenter = (30 * rendered.width + 30) * 4;
      assert.ok((rendered.data[cropCenter] ?? 255) < 15, format);
      assert.ok((rendered.data[cropCenter + 2] ?? 0) > 240, format);

      // Contain 20x10 inside 80x80 at (80, 10) -> box { x: 80, y: 30, width: 80, height: 40 }.
      // Rotated 90deg -> screenBox width ~ 40, height ~ 80.
      const containedNode = page.nodes.find((n) => n.authorId === "contained")!;
      assert.deepEqual(containedNode.box, {
        x: 80,
        y: 30,
        width: 80,
        height: 40,
      });
      expect(containedNode.screenBox?.width).toBeCloseTo(40, 1);
      expect(containedNode.screenBox?.height).toBeCloseTo(80, 1);

      // Mutate image after preview and verify earlier build & preview remain untouched.
      const prevPng = page.png.slice();
      const prevSvg = page.svg;
      const prevPptx = first.build.pptx.slice();
      mutable.set({ crop: { right: 0.5 } });
      const second = await previewDoc(doc);
      assert.deepEqual(first.pages[0]?.png, prevPng);
      assert.equal(first.pages[0]?.svg, prevSvg);
      assert.deepEqual(first.build.pptx, prevPptx);
      assert.notEqual(second.pages[0]?.svg, prevSvg);
      const updatedRendered = PNG.sync.read(Buffer.from(second.pages[0]!.png!));
      assert.ok((updatedRendered.data[cropCenter] ?? 0) > 240, format);
      assert.ok((updatedRendered.data[cropCenter + 2] ?? 255) < 15, format);
    }

    // JSX entry produces valid image preview via same formal pipeline.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "jsx-img" },
          createElement(Image, {
            id: "photo",
            src: pngBytes,
            x: 20,
            y: 20,
            width: 60,
            height: 60,
            fit: "contain",
          }),
        ),
      ),
    );
    const jsxPreview = await previewDoc(jsxDoc);
    assert.deepEqual(jsxPreview.pages[0]?.diagnostics, []);
    assert.deepEqual(
      jsxPreview.pages[0]?.nodes.find((n) => n.authorId === "photo")?.box,
      { x: 20, y: 35, width: 60, height: 30 },
    );
  });

  it("rejects GIF, non-rect geometry masks, image border/shadow styles, and corrupt PNG/JPEG streams with located diagnostics while preserving valid sibling pages", async () => {
    const { encode } = await import("jpeg-js");
    const previewDoc = createDocumentPreview(documentCapabilities);

    const validSource = createTwoToneSource();
    const validPng = new Uint8Array(PNG.sync.write(validSource));
    const validJpeg = new Uint8Array(encode(validSource, 90).data);

    // Minimal valid 1x1 GIF89a with one image block and trailer.
    const validGif = new Uint8Array(
      Buffer.from(
        "47494638396101000100800000ffffff0000002c00000000010001000002024401003b",
        "hex",
      ),
    );

    // Valid PNG header/chunks with corrupted IDAT compressed data.
    const corruptPng = new Uint8Array(validPng);
    const idatOffset = Buffer.from(corruptPng).indexOf(Buffer.from("IDAT"));
    corruptPng[idatOffset + 6] = (corruptPng[idatOffset + 6] ?? 0) ^ 0xff;
    corruptPng[idatOffset + 7] = (corruptPng[idatOffset + 7] ?? 0) ^ 0xff;

    // Valid JPEG framing (SOI, SOF, SOS, EOI) with corrupted entropy-coded scan stream.
    const corruptJpeg = new Uint8Array(validJpeg);
    const sosOffset = Buffer.from(corruptJpeg).indexOf(
      Buffer.from([0xff, 0xda]),
    );
    const sosLength =
      ((corruptJpeg[sosOffset + 2] ?? 0) << 8) +
      (corruptJpeg[sosOffset + 3] ?? 0);
    const scanStart = sosOffset + 2 + sosLength;
    for (let i = scanStart; i < corruptJpeg.length - 4; i++) {
      corruptJpeg[i] = 0x55;
    }

    const cases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof image>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "gif",
        node: image({ id: "bad", src: validGif, x: 10, y: 10 }),
        expected: /PNG and JPEG/i,
      },
      {
        label: "ellipse-mask",
        node: image({
          id: "bad",
          src: validPng,
          geometry: "ellipse",
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
        expected: /shaped images/i,
      },
      {
        label: "roundRect-mask",
        node: image({
          id: "bad",
          src: validPng,
          geometry: preset("roundRect"),
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
        expected: /shaped images/i,
      },
      {
        label: "custom-path-mask",
        node: image({
          id: "bad",
          src: validPng,
          geometry: path("M0 0 L10 0 L5 10 Z", { viewBox: [10, 10] }),
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
        expected: /shaped images/i,
      },
      {
        label: "borderColor",
        node: image({
          id: "bad",
          src: validPng,
          x: 10,
          y: 10,
          style: { borderColor: "#000000" },
        }),
        expected: /borderColor/i,
      },
      {
        label: "borderWidth",
        node: image({
          id: "bad",
          src: validPng,
          x: 10,
          y: 10,
          style: { borderWidth: 2 },
        }),
        expected: /borderWidth/i,
      },
      {
        label: "borderStyle",
        node: image({
          id: "bad",
          src: validPng,
          x: 10,
          y: 10,
          style: { borderStyle: "solid" },
        }),
        expected: /borderStyle/i,
      },
      {
        label: "boxShadow",
        node: image({
          id: "bad",
          src: validPng,
          x: 10,
          y: 10,
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        expected: /boxShadow/i,
      },
      {
        label: "softEdge",
        node: image({
          id: "bad",
          src: validPng,
          x: 10,
          y: 10,
          style: { softEdge: 4 },
        }),
        expected: /softEdge/i,
      },
      {
        label: "corrupt-png",
        node: image({ id: "bad", src: corruptPng, x: 10, y: 10 }),
        expected: /decode PNG/i,
      },
      {
        label: "corrupt-jpeg",
        node: image({ id: "bad", src: corruptJpeg, x: 10, y: 10 }),
        expected: /decode JPEG/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "ok-sibling" }).append(
            image({
              id: "ok-img",
              src: validPng,
              x: 20,
              y: 20,
              width: 60,
              height: 30,
            }),
          ),
        ),
      );
      const validation = await validatePptx(result.build.pptx);
      assert.deepEqual(validation.errors, [], item.label);
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.equal(result.pages[0]?.diagnostics[0]?.page, 1, item.label);
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /image\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
    }
  });
});
