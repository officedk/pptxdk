import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  font,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
} from "@pptxdk/capabilities/document";
import { blockArc, linearGradient, preset } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import type { CapabilityModule, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

const jostRegularBytes = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [
  { weight: 400, data: jostRegularBytes },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

function withSyntheticFontTable(source: Buffer, tag: string): Uint8Array {
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
}

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Slice 2: Shape static preview", () => {
  it("renders rect, ellipse, preset, custom path, transforms, zIndex, and screenGeometry", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
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
    // Uniform path evaluation renders all shapes (including ellipse) as SVG <path> elements.
    expect(page.svg).toContain("<path ");
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

  it("renders standard preset geometries (including adjust, shaded subpaths, local path w/h, and connectors) with zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 480, height: 240 }).append(
      slide({ id: "presets" }).append(
        shape({
          id: "star",
          geometry: "star5",
          x: 20,
          y: 20,
          width: 80,
          height: 80,
          fill: "#F59E0B",
        }),
        shape({
          id: "arc",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 120,
          y: 20,
          width: 80,
          height: 80,
          fill: "#10B981",
        }),
        shape({
          id: "bevel-box",
          geometry: preset("bevel", { adj: 20000 }),
          x: 220,
          y: 20,
          width: 100,
          height: 60,
          fill: "#3B82F6",
          style: { borderColor: "#1E3A8A", borderWidth: 2 },
        }),
        shape({
          id: "cloud-shape",
          geometry: "cloud",
          x: 340,
          y: 20,
          width: 110,
          height: 70,
          fill: "#E0F2FE",
          style: { borderColor: "#0284C7", borderWidth: 2 },
        }),
        shape({
          id: "donut-ring",
          geometry: "donut",
          x: 20,
          y: 120,
          width: 80,
          height: 80,
          fill: "#EC4899",
        }),
        shape({
          id: "connector",
          geometry: preset("bentConnector3", { adj1: 30000 }),
          x: 130,
          y: 130,
          width: 120,
          height: 60,
          style: { borderColor: "#0F172A", borderWidth: 2 },
        }),
      ),
    );
    const result = await previewDoc(doc);
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // Bevel produces shaded fill variants alongside the base #3B82F6 face.
    expect(page.svg).toContain('fill="#3B82F6"');
    expect(page.svg).not.toEqual(
      page.svg.replace(/fill="#(?!3B82F6)[0-9A-F]{6}"/u, ""),
    );
  });

  it("renders gradients, image fills, borderRadius, dashed/dotted borders, open-path arrows, boxShadow, and softEdge with zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const { PNG } = await import("pngjs");
    const { arrowTypes } = await import("@pptxdk/capabilities");
    const { path } = await import("@pptxdk/capabilities/shape");

    const samplePng = new PNG({ width: 4, height: 4 });
    for (let i = 0; i < samplePng.data.length; i += 4) {
      samplePng.data.set([0x16, 0xa3, 0x4a, 255], i);
    }
    const pngBytes = new Uint8Array(PNG.sync.write(samplePng));

    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "visuals", background: "#FFFFFF" }).append(
        shape({
          id: "grad-shape",
          x: 10,
          y: 10,
          width: 90,
          height: 40,
          style: {
            backgroundImage: linearGradient(90, "#111111", "#888888"),
            borderRadius: 8,
            borderColor: "#0F172A",
            borderWidth: 2,
            borderStyle: "dashed",
          },
        }),
        shape({
          id: "img-shape",
          geometry: "ellipse",
          x: 120,
          y: 10,
          width: 80,
          height: 40,
          style: {
            backgroundImage: { image: pngBytes },
            borderColor: "#064E3B",
            borderWidth: 2,
            borderStyle: "dotted",
          },
        }),
        shape({
          id: "shadow-shape",
          x: 220,
          y: 15,
          width: 70,
          height: 35,
          fill: "#112233",
          style: {
            rotate: 45,
            boxShadow: { blur: 4, distance: 6, angle: 90 },
          },
        }),
        shape({
          id: "soft-shape",
          x: 320,
          y: 10,
          width: 80,
          height: 40,
          fill: "#1D4ED8",
          style: { softEdge: 6 },
        }),
        ...arrowTypes.map((arrowType, idx) =>
          shape({
            id: `arrow-${arrowType}`,
            geometry:
              idx % 2 === 0
                ? "line"
                : path("M 0 12 Q 50 -6 100 12", {
                    viewBox: [100, 18],
                    fill: false,
                  }),
            x: 15 + idx * 75,
            y: 100,
            width: 60,
            height: 16,
            style: {
              borderColor: "#0F172A",
              borderWidth: 2.5,
              startArrow: arrowType,
              endArrow: arrowType,
            },
          }),
        ),
      ),
    );
    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    expect(page.svg).toContain("<linearGradient");
    expect(page.svg).toContain("<pattern");
    expect(page.svg).toContain("<feDropShadow");
    expect(page.svg).toContain("<feGaussianBlur");
    expect(page.svg).toContain("stroke-dasharray");
  });

  it("renders inline rich text inside shapes with textRect, flipH/flipV, rotate, opacity, and decorations", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 520, height: 280 }).append(
      slide({ id: "shape-text", background: "#FFFFFF" }).append(
        shape({
          id: "opacity-shape-text",
          x: 20,
          y: 20,
          width: 180,
          height: 80,
          fill: "#000000",
          style: { opacity: 0.5, padding: 8 },
          text: "IIII",
          textStyle: {
            fontFamily: jostFont,
            fontSize: 36,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "middle",
          },
        }),
        shape({
          id: "flipped-arrow-text",
          geometry: "rightArrow",
          x: 220,
          y: 20,
          width: 240,
          height: 90,
          fill: "#1D4ED8",
          flipH: true,
          flipV: true,
          style: { rotate: 15, borderColor: "#0F172A", borderWidth: 2 },
        }).append(
          paragraph([
            span("Page ").style({ fontWeight: 700, color: "#FFFFFF" }),
            slideNumber().style({ color: "#FDE047" }),
          ]).style({ fontFamily: jostFont, fontSize: 18, align: "center" }),
        ),
        shape({
          id: "diamond-wrapped",
          geometry: "diamond",
          x: 20,
          y: 130,
          width: 180,
          height: 120,
          style: {
            backgroundImage: linearGradient(90, "#0F766E", "#14B8A6"),
            boxShadow: { blur: 4, distance: 4, angle: 90 },
            borderColor: "#115E59",
            borderWidth: 2,
            borderStyle: "dashed",
            padding: 4,
          },
        }).append(
          paragraph("Diamond inner text wraps neatly").style({
            fontFamily: jostFont,
            fontSize: 14,
            color: "#FFFFFF",
            align: "center",
            textDecoration: "underline",
          }),
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);

    // Shape opacity (0.5 over #000000 on white) produces gray ~128 fill while text remains solid white (#FFFFFF).
    const pixels = PNG.sync.read(Buffer.from(page.png));
    const fillCorner = (30 * pixels.width + 30) * 4;
    expect(pixels.data[fillCorner]).toBeGreaterThan(115);
    expect(pixels.data[fillCorner]).toBeLessThan(140);
    // At least one pixel in the middle of opacity-shape-text is pure white (#FFFFFF) from the un-faded glyph.
    let foundPureWhiteText = false;
    for (let y = 45; y <= 75; y++) {
      for (let x = 90; x <= 130; x++) {
        const idx = (y * pixels.width + x) * 4;
        if (
          pixels.data[idx] === 255 &&
          pixels.data[idx + 1] === 255 &&
          pixels.data[idx + 2] === 255
        ) {
          foundPureWhiteText = true;
          break;
        }
      }
      if (foundPureWhiteText) break;
    }
    assert.equal(foundPureWhiteText, true);

    // Flipped arrow geometry uses scale(-1 -1), while its text layer must only rotate (no negative scale).
    expect(page.svg).toContain("scale(-1 -1)");
    expect(page.svg).toContain('transform="rotate(15 340 65)"');
    // SlideNumber resolved to "1" inside the shape text.
    expect(
      page.nodes.find((n) => n.authorId === "flipped-arrow-text")?.text,
    ).toBe("Page 1");

    // JSX equivalent with Shape + P/Span/SlideNumber also renders cleanly with zero diagnostics.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "jsx-shape-text", background: "#FFFFFF" },
          createElement(
            Shape,
            {
              id: "jsx-rich",
              geometry: "roundRect",
              x: 30,
              y: 30,
              width: 240,
              height: 100,
              fill: "#1E293B",
              style: { padding: 12, verticalAlign: "bottom" },
            },
            createElement(
              P,
              {
                style: { fontFamily: jostFont, fontSize: 16, color: "#FFFFFF" },
              },
              createElement(Span, { style: { fontWeight: 700 } }, "Bold "),
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
  });

  it("rejects only underlying font resource errors on shape inline text with located diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const colrFont = font(
      "ColrJost",
      [{ weight: 400, data: withSyntheticFontTable(jostRegularBytes, "COLR") }],
      { substitute: true },
    );

    const cases: readonly {
      label: string;
      node: ReturnType<typeof shape>;
      expected: RegExp;
    }[] = [
      {
        label: "missing-glyph-inline",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 120,
          height: 40,
          fill: "#112233",
          text: "汉",
          textStyle: { fontFamily: jostFont },
        }),
        expected: /glyph/i,
      },
      {
        label: "color-font-children",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 120,
          height: 40,
          fill: "#112233",
        }).append(paragraph("Color").style({ fontFamily: colrFont })),
        expected: /monochrome/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "s1" }).append(item.node),
          slide({ id: "s2" }).append(
            shape({
              id: "ok-sibling",
              x: 10,
              y: 10,
              width: 80,
              height: 40,
              fill: "#16A34A",
              text: "OK",
              textStyle: { fontFamily: jostFont, color: "#FFFFFF" },
            }),
          ),
        ),
      );
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, [], item.label);
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
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
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
      assert.ok(result.pages[1]?.png !== undefined, item.label);
    }

    // Verify JSX shape with missing glyph also rejects cleanly with located path.
    const jsxMissingGlyph = await expandDocument(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "jsx-text" },
          createElement(
            Shape,
            { id: "bad-jsx", x: 10, y: 10, width: 120, height: 40 },
            createElement(P, { style: { fontFamily: jostFont } }, "汉"),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxMissingGlyph);
    assert.equal(jsxResult.pages[0]?.png, undefined);
    assert.match(
      jsxResult.pages[0]?.diagnostics[0]?.path ?? "",
      /shape\[bad-jsx\]/,
    );
  });

  it("measures non-rectangular shape inline text against the inscribed textRect so diamond overflow emits overset", async () => {
    const resolveWith = createResolver(documentCapabilities);
    const label = "One two three four";
    const doc = deck({ width: 400, height: 200 }).append(
      slide({ id: "diamond-overset", background: "#FFFFFF" }).append(
        shape({
          id: "rect-box",
          geometry: preset("rect"),
          x: 20,
          y: 20,
          width: 160,
          height: 52,
          fill: "#E2E8F0",
          text: label,
          textStyle: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
        }),
        shape({
          id: "diamond-box",
          geometry: preset("diamond"),
          x: 210,
          y: 20,
          width: 160,
          height: 52,
          fill: "#FEF3C7",
          text: label,
          textStyle: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
        }),
      ),
    );

    const resolution = resolveWith(doc);
    assert.equal(
      resolution.report.diagnostics.some(
        (diag) =>
          diag.code === "overset" &&
          diag.path === "deck/slide[diamond-overset]/shape[rect-box]",
      ),
      false,
    );
    assert.ok(
      resolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" &&
          diag.code === "overset" &&
          diag.path === "deck/slide[diamond-overset]/shape[diamond-box]",
      ),
    );

    const { pptx } = emit(resolution);
    assert.deepEqual((await validatePptx(pptx)).errors, []);

    const preview = paint(resolution);
    const page = preview.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.svg !== undefined && page.svg.length > 0);
    assert.ok(page.png !== undefined);
    const decoded = PNG.sync.read(Buffer.from(page.png));
    assert.equal(decoded.width, 400);
    assert.equal(decoded.height, 200);
  });
});
