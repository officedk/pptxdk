import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
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

import {
  deck,
  documentCapabilities,
  font,
  notes,
  notesCapability,
  paragraph,
  slide,
  slideNumber,
  shape,
  span,
  text,
} from "../src/document.ts";
import { blockArc, preset } from "../src/geometry.ts";
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
