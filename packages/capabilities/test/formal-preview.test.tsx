import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  BuildError,
  Deck,
  Slide,
  expandDocument,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
} from "@react-pptx/core";
import {
  parseCapabilityLayoutContract,
  textStyleKeys,
} from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";
import { createDocumentPreview } from "@react-pptx/core/preview";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { Box } from "../src/box-elements.ts";
import {
  box,
  deck,
  documentCapabilities,
  font,
  group,
  image,
  notes,
  notesCapability,
  paragraph,
  slide,
  slideNumber,
  shape,
  span,
  text,
} from "../src/document.ts";
import { blockArc, path, preset } from "../src/geometry.ts";
import { Group } from "../src/group-elements.ts";
import { Image } from "../src/image-elements.ts";
import { motionCapability } from "../src/motion.ts";
import { Shape } from "../src/shape-elements.ts";
import { shapeCapability } from "../src/shape.ts";
import { slideBackgroundCapability } from "../src/slide-background.ts";
import { P, SlideNumber, Span, Text } from "../src/text-elements.ts";
import {
  parseTextContent,
  renderTextOutlinesSvg,
  textInsets,
  writeTextBodyFromParagraphs,
} from "../src/text-service.ts";
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

describe("Slice 3: Text static preview & shared text-service outlines", () => {
  const jostFont = font("Jost*", [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
          import.meta.url,
        ),
      ),
    },
    {
      weight: 700,
      data: readFileSync(
        new URL(
          "../../react-pptx/test/assets/fonts/Jost-700-Bold.ttf",
          import.meta.url,
        ),
      ),
    },
  ]);
  const notoFont = font("Noto Sans CJK SC", [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../react-pptx/test/assets/fonts/NotoSansSC-Regular.woff2",
          import.meta.url,
        ),
      ),
    },
  ]);

  it("renders wrapped CJK/Latin text, alignments, uppercase, slide number, and box visuals across JSX and document entries", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);

    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "text-main", background: "#F8FAFC" },
          createElement(
            Text,
            {
              id: "mixed",
              x: 20,
              y: 20,
              width: 240,
              style: {
                fontFamily: notoFont,
                fontSize: 18,
                padding: 10,
                backgroundColor: "#EEF2FF",
                borderColor: "#4F46E5",
                borderWidth: 2,
              },
            },
            createElement(
              P,
              { style: { lineHeight: 26, spacingAfter: 8 } },
              "共享排版 ",
              createElement(
                Span,
                {
                  style: {
                    fontFamily: jostFont,
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    textTransform: "uppercase",
                  },
                },
                "shared 2026",
              ),
              " 第",
              createElement(SlideNumber, null),
              "页",
            ),
            createElement(
              P,
              { style: { align: "right", lineHeight: 24 } },
              createElement(Span, { style: { fontFamily: jostFont } }, "Right"),
            ),
          ),
          createElement(
            Text,
            {
              id: "centered",
              x: 280,
              y: 40,
              width: 160,
              height: 120,
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                verticalAlign: "bottom",
                align: "center",
                backgroundColor: "#FEF3C7",
                opacity: 0.9,
                rotate: 90,
              },
            },
            "Bottom center",
          ),
        ),
      ),
    );

    const impDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "text-main", background: "#F8FAFC" }).append(
        text(
          [
            paragraph(
              [
                "共享排版 ",
                span("shared 2026", {
                  style: {
                    fontFamily: jostFont,
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    textTransform: "uppercase",
                  },
                }),
                " 第",
                slideNumber(),
                "页",
              ],
              { style: { lineHeight: 26, spacingAfter: 8 } },
            ),
            paragraph([span("Right", { style: { fontFamily: jostFont } })], {
              style: { align: "right", lineHeight: 24 },
            }),
          ],
          {
            id: "mixed",
            x: 20,
            y: 20,
            width: 240,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              padding: 10,
              backgroundColor: "#EEF2FF",
              borderColor: "#4F46E5",
              borderWidth: 2,
            },
          },
        ),
        text("Bottom center", {
          id: "centered",
          x: 280,
          y: 40,
          width: 160,
          height: 120,
          style: {
            fontFamily: jostFont,
            fontSize: 16,
            verticalAlign: "bottom",
            align: "center",
            backgroundColor: "#FEF3C7",
            opacity: 0.9,
            rotate: 90,
          },
        }),
      ),
    );

    const jsxPreview = await previewDoc(jsxDoc);
    const impPreview = await previewDoc(impDoc);
    assert.deepEqual((await validatePptx(jsxPreview.build.pptx)).errors, []);
    assert.deepEqual((await validatePptx(impPreview.build.pptx)).errors, []);

    assert.deepEqual(jsxPreview.pages[0]?.diagnostics, []);
    assert.deepEqual(impPreview.pages[0]?.diagnostics, []);
    assert.ok(jsxPreview.pages[0]?.png !== undefined);
    assert.deepEqual(jsxPreview.pages[0]?.png, impPreview.pages[0]?.png);
    expect(jsxPreview.pages[0]?.svg).toContain("<path");
    expect(jsxPreview.pages[0]?.svg).toContain('fill="#EEF2FF"');
    expect(jsxPreview.pages[0]?.svg).toContain("rotate(90");

    const centeredNode = jsxPreview.pages[0]?.nodes.find(
      (node) => node.authorId === "centered",
    );
    expect(centeredNode?.screenBox?.width).toBeCloseTo(120);
    expect(centeredNode?.screenBox?.height).toBeCloseTo(160);
    assert.equal(
      jsxPreview.build.query.getSlide("text-main").get("mixed").observe().text,
      "共享排版 shared 2026 第1页\nRight",
    );
  });

  it("rejects unsupported scripts, text styles, missing glyphs/fonts, color tables, and visual effects while keeping valid pages", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const jostBytes = readFileSync(
      new URL(
        "../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
        import.meta.url,
      ),
    );
    const withSyntheticTable = (source: Buffer, tag: string): Uint8Array => {
      const numTables = source.readUInt16BE(4);
      const headerSize = 12 + numTables * 16;
      const nextNumTables = numTables + 1;
      const nextHeaderSize = 12 + nextNumTables * 16;
      const records: {
        tag: string;
        checkSum: number;
        offset: number;
        length: number;
      }[] = [];
      for (let i = 0; i < numTables; i++) {
        const off = 12 + i * 16;
        records.push({
          tag: source.subarray(off, off + 4).toString("ascii"),
          checkSum: source.readUInt32BE(off + 4),
          offset: source.readUInt32BE(off + 8) + 16,
          length: source.readUInt32BE(off + 12),
        });
      }
      records.push({
        tag,
        checkSum: 0,
        offset: source.length + 16,
        length: 16,
      });
      records.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
      const out = Buffer.alloc(source.length + 32);
      source.copy(out, 0, 0, 12);
      out.writeUInt16BE(nextNumTables, 4);
      const entrySelector = Math.floor(Math.log2(nextNumTables));
      const searchRange = (1 << entrySelector) * 16;
      out.writeUInt16BE(searchRange, 6);
      out.writeUInt16BE(entrySelector, 8);
      out.writeUInt16BE(nextNumTables * 16 - searchRange, 10);
      for (let i = 0; i < records.length; i++) {
        const off = 12 + i * 16;
        out.write(records[i]!.tag, off, 4, "ascii");
        out.writeUInt32BE(records[i]!.checkSum, off + 4);
        out.writeUInt32BE(records[i]!.offset, off + 8);
        out.writeUInt32BE(records[i]!.length, off + 12);
      }
      source.copy(out, nextHeaderSize, headerSize);
      return new Uint8Array(out);
    };

    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof text>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "complex-script",
        node: text("مرحبا", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont },
        }),
        expected: /Latin and Han/i,
      },
      {
        label: "underline",
        node: text("Underlined", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont, textDecoration: "underline" },
        }),
        expected: /decoration/i,
      },
      {
        label: "textShadow",
        node: text("Shadowed", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont, textShadow: { blur: 2, distance: 1 } },
        }),
        expected: /shadow/i,
      },
      {
        label: "letterSpacing",
        node: text("Tracked", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont, letterSpacing: 2 },
        }),
        expected: /letter spacing/i,
      },
      {
        label: "justify",
        node: text("Justified paragraph", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont, align: "justify" },
        }),
        expected: /justified/i,
      },
      {
        label: "list",
        node: text(
          [paragraph("Bullet item", { style: { list: { type: "bullet" } } })],
          {
            id: "bad",
            x: 10,
            y: 10,
            width: 160,
            height: 50,
            style: { fontFamily: jostFont },
          },
        ),
        expected: /lists/i,
      },
      {
        label: "missing-glyph",
        node: text("汉", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont },
        }),
        expected: /glyph/i,
      },
      {
        label: "missing-font-file",
        node: text("No font file", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: font("Missing Preview Face") },
        }),
        expected: /font file/i,
      },
      {
        label: "colr-font",
        node: text("Color font", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: {
            fontFamily: font(
              "ColrJost",
              [{ weight: 400, data: withSyntheticTable(jostBytes, "COLR") }],
              { substitute: true },
            ),
          },
        }),
        expected: /monochrome outline/i,
      },
      {
        label: "sbix-font",
        node: text("Bitmap font", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: {
            fontFamily: font(
              "SbixJost",
              [{ weight: 400, data: withSyntheticTable(jostBytes, "sbix") }],
              { substitute: true },
            ),
          },
        }),
        expected: /monochrome outline/i,
      },
      {
        label: "boxShadow-on-text",
        node: text("Box shadow", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: {
            fontFamily: jostFont,
            boxShadow: { blur: 4, distance: 2 },
          },
        }),
        expected: /boxShadow/i,
      },
      {
        label: "dashed-border-on-text",
        node: text("Dashed", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: {
            fontFamily: jostFont,
            borderColor: "#111827",
            borderStyle: "dashed",
          },
        }),
        expected: /non-solid/i,
      },
    ];

    for (const item of rejectionCases) {
      const result = await previewDoc(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "good-page" }).append(
            text("Valid", {
              id: "ok",
              x: 10,
              y: 10,
              width: 120,
              style: { fontFamily: jostFont },
            }),
          ),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.build.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /text\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
    }
  });

  it("shares text-service layout and glyph outlines without registering textCapability or loading text.ts", async () => {
    const customLabelCapability: BoundObjectCapabilityModule = {
      declaration: "rp:custom-label",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set(["width", "height", ...textStyleKeys]),
          allowPadding: false,
          path: parseContext.path,
        });
        const insets = textInsets(input.props["style"]);
        const { paragraphs, style } = parseTextContent(
          input.children,
          input.props["style"],
          parseContext.theme,
          { slideNumber: parseContext.slideNumber },
        );
        let boundResult:
          | {
              readonly box: {
                x: number;
                y: number;
                width: number;
                height: number;
              };
              readonly measuredLayout: ReturnType<
                typeof writeTextBodyFromParagraphs
              >["measuredLayout"];
            }
          | undefined;
        return {
          contract,
          compile(compileContext) {
            const box = compileContext.layoutBox ?? {
              x: Number(input.props["x"] ?? 0),
              y: Number(input.props["y"] ?? 0),
              width: Number(input.props["width"] ?? 120),
              height: Number(input.props["height"] ?? 40),
            };
            const body = writeTextBodyFromParagraphs(
              paragraphs,
              style,
              compileContext.theme,
              "middle",
              { width: box.width, height: box.height, insets },
              "text-box",
              compileContext.options.embedFonts,
              compileContext.createTextFieldId,
            );
            boundResult = { box, measuredLayout: body.measuredLayout };
            return {
              xml: `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="CustomLabel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${body.xml}</p:sp>`,
              observation: {
                id: compileContext.id,
                kind: "shape",
                name: "CustomLabel",
                geometry: "rect",
                box,
                text: body.text,
              },
              extensions: [],
              resourceClaims: body.resourceClaims,
            };
          },
          preview(previewContext) {
            const box = previewContext.box ?? boundResult!.box;
            return {
              svg: renderTextOutlinesSvg({
                paragraphs,
                box,
                insets,
                verticalAlign: "middle",
                theme: previewContext.theme,
                measuredLayout: boundResult?.measuredLayout,
              }),
            };
          },
        };
      },
    };

    const previewWithoutTextCapability = createDocumentPreview([
      slideBackgroundCapability,
      customLabelCapability,
    ]);
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "shared-only" }).append(
        new RichTextContainer("custom-label", "rp:custom-label", {
          id: "lbl",
          x: 20,
          y: 20,
          width: 180,
          height: 50,
          style: { fontFamily: jostFont, fontSize: 18, color: "#1D4ED8" },
        }).content("Shared glyphs"),
      ),
    );
    const result = await previewWithoutTextCapability(doc);
    assert.deepEqual(result.pages[0]?.diagnostics, []);
    expect(result.pages[0]?.svg).toContain("<path");
    assert.ok(result.pages[0]?.png !== undefined);

    // Also verify text-service never pulls in text.ts at runtime.
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const isolated = spawnSync(
      process.execPath,
      [
        "--conditions=source",
        "--input-type=module",
        "--eval",
        `
import { registerHooks } from "node:module";
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier.endsWith("/text.ts") || specifier.endsWith("/text.js") || specifier.endsWith("/text")) {
      throw new Error("text-service loaded text capability module: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});
await import("./src/text-service.ts");
process.stdout.write("ok");
`,
      ],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(isolated.status, 0, isolated.stderr);
    assert.equal(isolated.stdout, "ok");
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

describe("Slice 4 (Ticket 05): Box static preview positive and negative behaviors", () => {
  it("renders transparent and visible Box containers with solid fill/border, opacity, rotate, theme overrides, and post-preview isolation", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const mutableCard = box({
      id: "visible-card",
      theme: { colors: { cardTone: "#1D4ED8" } },
      style: {
        width: 140,
        height: 80,
        padding: 12,
        backgroundColor: "cardTone",
        borderColor: "#0F172A",
        borderWidth: 2,
        borderStyle: "solid",
        opacity: 0.85,
        rotate: 90,
      },
    }).append(
      shape({
        id: "card-dot",
        width: 24,
        height: 24,
        fill: "#FACC15",
      }),
    );

    const doc = deck({
      width: 360,
      height: 200,
      theme: { colors: { cardTone: "#991B1B" } },
    }).append(
      slide({ id: "boxes", background: "#FFFFFF" }).append(
        box({
          id: "transparent-row",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            display: "flex",
            flexDirection: "row",
            gap: 20,
            padding: 8,
          },
        }).append(
          mutableCard,
          box({
            id: "border-only",
            style: {
              width: 100,
              height: 60,
              padding: 8,
              borderColor: "#16A34A",
              borderWidth: 3,
            },
          }).append(
            shape({
              id: "inner-rect",
              width: 30,
              height: 20,
              fill: "#0EA5E9",
            }),
          ),
        ),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // Transparent outer box does not emit a native shape object; only visible boxes and child shapes become objects.
    assert.equal(page.nodes.length, 4);
    assert.equal(
      page.nodes.some((node) => node.authorId === "transparent-row"),
      false,
    );
    assert.equal(page.layouts[0]?.authorId, "transparent-row");
    expect(page.svg).toContain('fill="#1D4ED8"');
    expect(page.svg).toContain('stroke="#16A34A"');
    // Visible box background appears before its child dot in SVG paint order.
    expect(page.svg.indexOf('fill="#1D4ED8"')).toBeLessThan(
      page.svg.indexOf('fill="#FACC15"'),
    );
    const rotatedCard = page.nodes.find((n) => n.authorId === "visible-card")!;
    expect(rotatedCard.screenBox?.width).toBeCloseTo(80);
    expect(rotatedCard.screenBox?.height).toBeCloseTo(140);

    // Mutate after preview and rebuild: prior PNG, SVG, and PPTX remain untouched.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    const firstPptx = first.build.pptx.slice();
    mutableCard.style({ backgroundColor: "#059669" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.build.pptx, firstPptx);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // JSX entry renders transparent and visible Box identically.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 240, height: 140 },
        createElement(
          Slide,
          { id: "jsx-box", background: "#FFFFFF" },
          createElement(
            Box,
            {
              id: "jsx-transparent",
              style: { display: "flex", padding: 10 },
            },
            createElement(
              Box,
              {
                id: "jsx-visible",
                style: {
                  width: 120,
                  height: 60,
                  padding: 8,
                  backgroundColor: "#E2E8F0",
                  borderColor: "#334155",
                  borderWidth: 2,
                },
              },
              createElement(Shape, {
                id: "jsx-child",
                width: 40,
                height: 20,
                fill: "#EF4444",
              }),
            ),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    expect(jsxResult.pages[0]?.svg).toContain('fill="#E2E8F0"');
  });

  it("rejects unsupported visible Box preview styles while retaining valid sibling pages and PPTX", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof box>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "backgroundImage",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundImage: linearGradient(90, "#1D4ED8", "#9333EA"),
          },
        }),
        expected: /backgroundImage/i,
      },
      {
        label: "borderRadius",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            borderRadius: 12,
          },
        }),
        expected: /borderRadius/i,
      },
      {
        label: "boxShadow",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            boxShadow: { blur: 4, distance: 2 },
          },
        }),
        expected: /boxShadow/i,
      },
      {
        label: "softEdge",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            softEdge: 4,
          },
        }),
        expected: /softEdge/i,
      },
      {
        label: "dashed-border",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            borderColor: "#0F172A",
            borderStyle: "dashed",
          },
        }),
        expected: /non-solid/i,
      },
      {
        label: "dotted-border",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            borderColor: "#0F172A",
            borderStyle: "dotted",
          },
        }),
        expected: /non-solid/i,
      },
    ];

    for (const item of rejectionCases) {
      const result = await previewDoc(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "bad-box-page" }).append(item.node),
          slide({ id: "good-box-page" }).append(
            box({
              id: "ok",
              style: { width: 100, height: 50, backgroundColor: "#10B981" },
            }),
          ),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.build.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /box\[bad\]/,
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

describe("Slice 5 (Ticket 08): Group static preview, screenGeometry, and member failure propagation", () => {
  it("renders nested rotated Group trees with flipped members, Box children, screenGeometry, and post-preview isolation", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const mutableShape = shape({
      id: "inner-leaf",
      x: 110,
      y: 70,
      width: 60,
      height: 20,
      fill: "#1D4ED8",
      flipH: true,
    });
    const doc = deck({
      width: 400,
      height: 240,
      theme: { colors: { groupAccent: "#111111" } },
    }).append(
      slide({ id: "groups-preview", background: "#FFFFFF" }).append(
        group({
          id: "outer",
          theme: { colors: { groupAccent: "#0F766E" } },
          style: { rotate: 90 },
        }).append(
          shape({
            id: "base-rect",
            x: 70,
            y: 40,
            width: 80,
            height: 40,
            fill: "groupAccent",
          }),
          box({
            id: "group-box",
            style: {
              position: "absolute",
              left: 150,
              top: 40,
              width: 60,
              height: 40,
              padding: 4,
              backgroundColor: "#E2E8F0",
            },
          }).append(
            shape({
              id: "box-child",
              width: 30,
              height: 20,
              fill: "#F59E0B",
            }),
          ),
          group({
            id: "inner",
            style: { rotate: -90 },
          }).append(mutableShape),
        ),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain('transform="rotate(90 ');
    expect(page.svg).toContain('transform="rotate(-90 ');
    expect(page.svg).toContain('fill="#0F766E"');
    expect(page.svg).toContain('fill="#1D4ED8"');

    const outerNode = page.nodes.find((n) => n.authorId === "outer")!;
    // Outer bounds: x=[70..210] (width 140), y=[40..90] (height 50), center=(140, 65).
    // Rotated 90deg -> width=50, height=140, x=115, y=-5.
    assert.deepEqual(outerNode.box, { x: 70, y: 40, width: 140, height: 50 });
    expect(outerNode.screenBox?.x).toBeCloseTo(115);
    expect(outerNode.screenBox?.y).toBeCloseTo(-5);
    expect(outerNode.screenBox?.width).toBeCloseTo(50);
    expect(outerNode.screenBox?.height).toBeCloseTo(140);
    assert.equal(outerNode.corners?.length, 4);

    // Inner group is rotated -90 around (140, 80), then +90 around (140, 65) => net 0deg rotation!
    // Center (140, 80) rotated +90deg around (140, 65) goes to (125, 65).
    // Since net rotation on inner-leaf is 0deg, its screenBox has width=60, height=20 centered at (125, 65).
    const innerLeafNode = page.nodes.find((n) => n.authorId === "inner-leaf")!;
    expect(innerLeafNode.screenBox?.x).toBeCloseTo(95);
    expect(innerLeafNode.screenBox?.y).toBeCloseTo(55);
    expect(innerLeafNode.screenBox?.width).toBeCloseTo(60);
    expect(innerLeafNode.screenBox?.height).toBeCloseTo(20);

    // Mutate after preview: prior outputs remain unchanged.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    const firstPptx = first.build.pptx.slice();
    mutableShape.set({ fill: "#DC2626" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.build.pptx, firstPptx);
    expect(second.pages[0]?.svg).toContain('fill="#DC2626"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // JSX entry renders equivalent rotated Group.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 240, height: 160 },
        createElement(
          Slide,
          { id: "jsx-group", background: "#FFFFFF" },
          createElement(
            Group,
            { id: "jg", style: { rotate: 90 } },
            createElement(Shape, {
              id: "js",
              x: 40,
              y: 30,
              width: 80,
              height: 40,
              fill: "#2563EB",
            }),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    const jsxGroupNode = jsxResult.pages[0]?.nodes.find(
      (n) => n.authorId === "jg",
    );
    expect(jsxGroupNode?.screenBox?.width).toBeCloseTo(40);
    expect(jsxGroupNode?.screenBox?.height).toBeCloseTo(80);
  });

  it("propagates member preview failures out of Group without omitting failed members or breaking sibling pages", async () => {
    const previewDoc = createDocumentPreview(documentCapabilities);
    const result = await previewDoc(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "group-with-bad-member" }).append(
          group({ id: "g", style: { rotate: 15 } }).append(
            shape({
              id: "ok-member",
              x: 20,
              y: 20,
              width: 50,
              height: 30,
              fill: "#10B981",
            }),
            shape({
              id: "bad-member",
              x: 80,
              y: 20,
              width: 50,
              height: 30,
              fill: "#EF4444",
              style: { boxShadow: { blur: 4, distance: 2 } },
            }),
          ),
        ),
        slide({ id: "valid-sibling" }).append(
          group({ id: "ok-group" }).append(
            shape({
              id: "ok-shape",
              x: 20,
              y: 20,
              width: 60,
              height: 40,
              fill: "#2563EB",
            }),
          ),
        ),
      ),
    );

    assert.deepEqual((await validatePptx(result.build.pptx)).errors, []);
    const badPage = result.pages[0]!;
    assert.equal(badPage.png, undefined);
    assert.equal(badPage.svg, undefined);
    assert.ok(
      badPage.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /group\[g\]\/shape\[bad-member\]/u.test(d.path) &&
          /boxShadow/i.test(d.message),
      ),
    );
    // Nodes and transformed screenBoxes are still reported for locator inspection.
    assert.ok(
      badPage.nodes.find((n) => n.authorId === "bad-member")?.screenBox !==
        undefined,
    );

    const goodPage = result.pages[1]!;
    assert.deepEqual(goodPage.diagnostics, []);
    assert.ok(goodPage.png !== undefined);

    // Non-finite rotation on Group fails formal build validation.
    await assert.rejects(
      () =>
        previewDoc(
          deck({ width: 320, height: 180 }).append(
            slide({ id: "bad-rot" }).append(
              group({ id: "g", style: { rotate: Number.NaN } }).append(
                shape({ id: "s", x: 10, y: 10, width: 20, height: 20 }),
              ),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "invalid-visual-style"),
        );
        return true;
      },
    );
  });
});
