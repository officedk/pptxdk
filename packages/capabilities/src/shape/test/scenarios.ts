import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  font,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
} from "@pptxdk/capabilities/document";
import {
  blockArc,
  linearGradient,
  path,
  preset,
  presetGeometryTuple,
  radialGradient,
  type CustomGeometry,
  type PresetGeometryName,
} from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide } from "@pptxdk/react";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import { stripVariableFontTables } from "../../../test/sfnt-test-helpers.ts";

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

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

const notoArabicFont = font("Noto Naskh Arabic", [
  {
    weight: 400,
    data: stripVariableFontTables(
      readFileSync(
        new URL(
          "../../../test/assets/fonts/NotoNaskhArabic.ttf",
          import.meta.url,
        ),
      ),
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

const colrJostFont = font(
  "ColrJost",
  [{ weight: 400, data: withSyntheticFontTable(jostRegularBytes, "COLR") }],
  { substitute: true },
);

const missingPreviewFont = font("Missing Shape Preview Face");

const sampleFillPng = (() => {
  const img = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < img.data.length; i += 4) {
    img.data.set([0x1e, 0x3a, 0x8a, 255], i);
  }
  return new Uint8Array(PNG.sync.write(img));
})();

const dualSubpathCustom: CustomGeometry = {
  type: "custom",
  width: 100,
  height: 100,
  paths: [
    {
      commands: [
        { op: "M", x: 50, y: 0 },
        { op: "L", x: 100, y: 50 },
        { op: "L", x: 50, y: 100 },
        { op: "L", x: 0, y: 50 },
        { op: "Z" },
      ],
      fill: true,
      stroke: false,
    },
    {
      commands: [
        { op: "M", x: 15, y: 50 },
        { op: "Q", x1: 50, y1: 15, x: 85, y: 50 },
      ],
      fill: false,
      stroke: true,
    },
  ],
};

const ribbonPath = path("M0 0 L80 0 L100 50 L80 100 L0 100 L20 50 Z", {
  viewBox: [100, 100],
});

export const shapePositiveScenario: CapabilityScenario = {
  id: "shape-primitives-and-transforms",
  capability: "shape",
  description:
    "Dual-author Shape scenario covering rect, ellipse, preset, custom path, fill/border, opacity, rotate, flip, theme tokens, and zIndex.",
  visualCheckpoints: [
    "Page 1 (geometry-and-styles): light gray #F4F6F8 canvas; blue rectangle (#2563EB) with 3px dark border (#1E293B) at top-left; semi-transparent red ellipse (#EF4444, opacity 0.7) overlapping the rectangle's bottom-right corner; 45-degree rotated emerald diamond (#10B981 via Slide themeOverride accent1) in the center; horizontally flipped purple chevron (#8B5CF6) and orange/navy custom diamond+arc shape on the right.",
    "Page 2 (layer-order): warm amber rectangle (#F59E0B, zIndex 5) rendered in front of dark slate rectangle (#334155, zIndex 1) despite being declared earlier; adjacent borderless teal rectangle (borderWidth: 0) has no dark outline.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all shapes remain editable native DrawingML <p:sp> objects with matching geometry, rotation, flipH, and z-order.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        theme: { colors: { accent1: "#059669" } },
      },
      createElement(
        Slide,
        {
          id: "geometry-and-styles",
          background: "#F4F6F8",
          themeOverride: { colors: { accent1: "#10B981" } },
        },
        createElement(Shape, {
          id: "base-rect",
          style: {
            position: "absolute",
            left: 32,
            top: 36,
            width: 150,
            height: 95,
            background: "#2563EB",
            border: { color: "#1E293B", width: 3, style: "solid" },
            zIndex: 1,
          },
        }),
        createElement(Shape, {
          id: "overlap-ellipse",
          geometry: preset("ellipse"),
          style: {
            position: "absolute",
            left: 110,
            top: 80,
            width: 150,
            height: 95,
            background: "#EF4444",
            border: { color: "#7F1D1D", width: 2 },
            opacity: 0.7,
            zIndex: 2,
          },
        }),
        createElement(Shape, {
          id: "rotated-square",
          geometry: preset("rect"),
          style: {
            position: "absolute",
            left: 300,
            top: 70,
            width: 84,
            height: 84,
            background: "accent1",
            border: { color: "#065F46", width: 2 },
            rotate: 45,
          },
        }),
        createElement(Shape, {
          id: "flipped-chevron",
          geometry: ribbonPath,
          style: {
            position: "absolute",
            left: 430,
            top: 45,
            width: 140,
            height: 80,
            background: "#8B5CF6",
            flipH: true,
            border: { color: "#4C1D95", width: 2 },
          },
        }),
        createElement(Shape, {
          id: "custom-dual-path",
          geometry: dualSubpathCustom,
          style: {
            position: "absolute",
            left: 440,
            top: 165,
            width: 120,
            height: 110,
            background: "#F97316",
            flipV: true,
            border: { color: "#0F172A", width: 3 },
          },
        }),
      ),
      createElement(
        Slide,
        { id: "layer-order", background: "#FFFFFF" },
        createElement(Shape, {
          id: "front-gold",
          style: {
            position: "absolute",
            left: 80,
            top: 70,
            width: 180,
            height: 110,
            background: "#F59E0B",
            zIndex: 5,
          },
        }),
        createElement(Shape, {
          id: "back-slate",
          style: {
            position: "absolute",
            left: 140,
            top: 110,
            width: 180,
            height: 110,
            background: "#334155",
            zIndex: 1,
          },
        }),
        createElement(Shape, {
          id: "borderless-teal",
          style: {
            position: "absolute",
            left: 380,
            top: 90,
            width: 160,
            height: 100,
            background: "#0D9488",
            border: { color: "#000000", width: 0 },
          },
        }),
      ),
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      theme: { colors: { accent1: "#059669" } },
    }).append(
      slide({
        id: "geometry-and-styles",
        background: "#F4F6F8",
        themeOverride: { colors: { accent1: "#10B981" } },
      }).append(
        shape({
          id: "base-rect",
          style: {
            position: "absolute",
            left: 32,
            top: 36,
            width: 150,
            height: 95,
            background: "#2563EB",
            border: { color: "#1E293B", width: 3, style: "solid" },
            zIndex: 1,
          },
        }),
        shape({
          id: "overlap-ellipse",
          geometry: preset("ellipse"),
          style: {
            position: "absolute",
            left: 110,
            top: 80,
            width: 150,
            height: 95,
            background: "#EF4444",
            border: { color: "#7F1D1D", width: 2 },
            opacity: 0.7,
            zIndex: 2,
          },
        }),
        shape({
          id: "rotated-square",
          geometry: preset("rect"),
          style: {
            position: "absolute",
            left: 300,
            top: 70,
            width: 84,
            height: 84,
            background: "accent1",
            border: { color: "#065F46", width: 2 },
            rotate: 45,
          },
        }),
        shape({
          id: "flipped-chevron",
          geometry: ribbonPath,
          style: {
            position: "absolute",
            left: 430,
            top: 45,
            width: 140,
            height: 80,
            background: "#8B5CF6",
            flipH: true,
            border: { color: "#4C1D95", width: 2 },
          },
        }),
        shape({
          id: "custom-dual-path",
          geometry: dualSubpathCustom,
          style: {
            position: "absolute",
            left: 440,
            top: 165,
            width: 120,
            height: 110,
            background: "#F97316",
            flipV: true,
            border: { color: "#0F172A", width: 3 },
          },
        }),
      ),
      slide({ id: "layer-order", background: "#FFFFFF" }).append(
        shape({
          id: "front-gold",
          style: {
            position: "absolute",
            left: 80,
            top: 70,
            width: 180,
            height: 110,
            background: "#F59E0B",
            zIndex: 5,
          },
        }),
        shape({
          id: "back-slate",
          style: {
            position: "absolute",
            left: 140,
            top: 110,
            width: 180,
            height: 110,
            background: "#334155",
            zIndex: 1,
          },
        }),
        shape({
          id: "borderless-teal",
          style: {
            position: "absolute",
            left: 380,
            top: 90,
            width: 160,
            height: 100,
            background: "#0D9488",
            border: { color: "#000000", width: 0 },
          },
        }),
      ),
    );
  },
  assert({ author, resolution, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.equal(validation.slides.length, 2, author);
    assert.equal(validation.slides[0]?.objects.length, 5, author);

    const p1 = preview.pages[0]!;
    assert.deepEqual(p1.errors, [], author);
    assert.ok(p1.png !== undefined, author);
    const rotatedNode = p1.nodes.find((n) => n.authorId === "rotated-square")!;
    expect(rotatedNode.bounds?.width).toBeCloseTo(84 * Math.SQRT2, 1);
    expect(rotatedNode.bounds?.height).toBeCloseTo(84 * Math.SQRT2, 1);
    assert.equal(
      resolution.slide("geometry-and-styles").get("overlap-ellipse").objects[0]
        ?.geometry,
      "ellipse",
    );

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.errors, [], author);
    assert.ok(p2.png !== undefined, author);
    const pixels = PNG.sync.read(Buffer.from(p2.png));
    // Overlap point (180, 140) must display gold #F59E0B on top of slate.
    const idx = (140 * pixels.width + 180) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(idx, idx + 4)],
      [0xf5, 0x9e, 0x0b, 255],
    );
  },
};

const catalogPalette = [
  "#2563EB",
  "#059669",
  "#D97706",
  "#7C3AED",
  "#DB2777",
  "#0891B2",
] as const;

const multiSubpathBezierCustom: CustomGeometry = {
  type: "custom",
  width: 100,
  height: 100,
  paths: [
    {
      commands: [
        { op: "M", x: 10, y: 80 },
        { op: "C", x1: 10, y1: 15, x2: 90, y2: 15, x: 90, y: 80 },
        { op: "Q", x1: 50, y1: 95, x: 10, y: 80 },
        { op: "Z" },
      ],
      fill: true,
      stroke: false,
    },
    {
      commands: [
        { op: "M", x: 15, y: 20 },
        { op: "L", x: 85, y: 20 },
        { op: "Q", x1: 50, y1: 60, x: 15, y: 85 },
      ],
      fill: false,
      stroke: true,
    },
  ],
};

const evenOddRingPath = path(
  "M 0 0 L 100 0 L 100 100 L 0 100 Z M 30 30 L 70 30 L 70 70 L 30 70 Z",
  { viewBox: [100, 100], fillRule: "evenodd" },
);

function chunkPresets(
  names: readonly PresetGeometryName[],
  chunkSize: number,
): readonly (readonly PresetGeometryName[])[] {
  const chunks: PresetGeometryName[][] = [];
  for (let i = 0; i < names.length; i += chunkSize) {
    chunks.push(names.slice(i, i + chunkSize));
  }
  return chunks;
}

const presetPages = chunkPresets(presetGeometryTuple, 50);

function pixelRgb(
  pixels: PNG,
  x: number,
  y: number,
): readonly [number, number, number] {
  const idx = (Math.round(y) * pixels.width + Math.round(x)) * 4;
  return [
    pixels.data[idx] ?? 0,
    pixels.data[idx + 1] ?? 0,
    pixels.data[idx + 2] ?? 0,
  ];
}

export const shapePresetCatalogScenario: CapabilityScenario = {
  id: "shape-187-presets-and-custom-parity",
  capability: "shape",
  description:
    "Data-driven dual-author scenario iterating all 187 ECMA-376 preset geometries plus adjust overrides and transformed multi-subpath custom geometries.",
  visualCheckpoints: [
    "Pages 1-4 (presets-page-1..4): 10x5 grid displaying all 187 standard ECMA-376 preset geometries with alternating non-square aspect ratios and crisp fills/strokes.",
    "Page 5 (adjust-and-custom-parity): blockArc, star5 (sharp adj), roundRect, wedgeRoundRectCallout, bentConnector3, donut (hollow center), chevron, circularArrow, bevel/cube (shaded faces), plus multi-subpath bezier custom shapes scaled, rotated, flipped, and evenodd-holed.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all 187 preset names appear as native <a:prstGeom> shapes alongside <a:custGeom> bezier shapes.",
  ],
  jsx() {
    const gridSlides = presetPages.map((group, pageIndex) =>
      createElement(
        Slide,
        {
          key: `presets-${pageIndex + 1}`,
          id: `presets-page-${pageIndex + 1}`,
          background: "#F8FAFC",
        },
        ...group.map((name, itemIndex) => {
          const col = itemIndex % 10;
          const row = Math.floor(itemIndex / 10);
          const wide = itemIndex % 2 === 0;
          return createElement(Shape, {
            key: name,
            id: `preset-${name}`,
            geometry: name,
            style: {
              position: "absolute",
              left: 24 + col * 92 + (wide ? 0 : 8),
              top: 22 + row * 102 + (wide ? 6 : 0),
              width: wide ? 76 : 60,
              height: wide ? 60 : 76,
              background:
                catalogPalette[
                  (pageIndex * 50 + itemIndex) % catalogPalette.length
                ]!,
              border: { color: "#0F172A", width: 1.5 },
            },
          });
        }),
      ),
    );

    const detailSlide = createElement(
      Slide,
      {
        key: "adjust-and-custom-parity",
        id: "adjust-and-custom-parity",
        background: "#F8FAFC",
      },
      createElement(Shape, {
        id: "adj-block-arc",
        geometry: blockArc({ startAngle: 30, endAngle: 250, thickness: 0.35 }),
        style: {
          position: "absolute",
          left: 30,
          top: 30,
          width: 120,
          height: 90,
          background: "#10B981",
          border: { color: "#064E3B", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "adj-star5",
        geometry: preset("star5", { adj: 35000 }),
        style: {
          position: "absolute",
          left: 175,
          top: 30,
          width: 100,
          height: 100,
          background: "#F59E0B",
          border: { color: "#78350F", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "adj-round-rect",
        geometry: preset("roundRect", { adj: 25000 }),
        style: {
          position: "absolute",
          left: 300,
          top: 30,
          width: 130,
          height: 85,
          background: "#3B82F6",
          border: { color: "#1E3A8A", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "adj-callout",
        geometry: preset("wedgeRoundRectCallout", {
          adj1: 28000,
          adj2: 42000,
          adj3: 16667,
        }),
        style: {
          position: "absolute",
          left: 460,
          top: 30,
          width: 140,
          height: 90,
          background: "#8B5CF6",
          border: { color: "#4C1D95", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "adj-connector",
        geometry: preset("bentConnector3", { adj1: 25000 }),
        style: {
          position: "absolute",
          left: 630,
          top: 35,
          width: 120,
          height: 80,
          border: { color: "#DC2626", width: 3 },
        },
      }),
      createElement(Shape, {
        id: "adj-donut",
        geometry: preset("donut", { adj: 30000 }),
        style: {
          position: "absolute",
          left: 780,
          top: 30,
          width: 100,
          height: 100,
          background: "#EC4899",
          border: { color: "#831843", width: 2 },
        },
      }),
      createElement(Shape, {
        id: "adj-chevron",
        geometry: preset("chevron", { adj: 30000 }),
        style: {
          position: "absolute",
          left: 30,
          top: 160,
          width: 130,
          height: 80,
          background: "#0EA5E9",
        },
      }),
      createElement(Shape, {
        id: "adj-circular-arrow",
        geometry: preset("circularArrow", {
          adj1: 15000,
          adj2: 1500000,
          adj3: 20000000,
          adj4: 9000000,
          adj5: 15000,
        }),
        style: {
          position: "absolute",
          left: 190,
          top: 155,
          width: 100,
          height: 100,
          background: "#6366F1",
          border: { color: "#312E81", width: 1.5 },
        },
      }),
      createElement(Shape, {
        id: "shaded-bevel",
        geometry: preset("bevel", { adj: 22000 }),
        style: {
          position: "absolute",
          left: 320,
          top: 165,
          width: 120,
          height: 80,
          background: "#2563EB",
          border: { color: "#0F172A", width: 1.5 },
        },
      }),
      createElement(Shape, {
        id: "shaded-cube",
        geometry: preset("cube", { adj: 25000 }),
        style: {
          position: "absolute",
          left: 470,
          top: 165,
          width: 110,
          height: 85,
          background: "#059669",
          border: { color: "#064E3B", width: 1.5 },
        },
      }),
      createElement(Shape, {
        id: "custom-scaled-bezier",
        geometry: multiSubpathBezierCustom,
        style: {
          position: "absolute",
          left: 30,
          top: 290,
          width: 180,
          height: 90,
          background: "#F97316",
          border: { color: "#0F172A", width: 3 },
        },
      }),
      createElement(Shape, {
        id: "custom-flipped-rotated",
        geometry: multiSubpathBezierCustom,
        style: {
          position: "absolute",
          left: 260,
          top: 290,
          width: 140,
          height: 90,
          background: "#14B8A6",
          flipH: true,
          flipV: true,
          border: { color: "#0F172A", width: 3 },
          rotate: 90,
        },
      }),
      createElement(Shape, {
        id: "custom-evenodd-hole",
        geometry: evenOddRingPath,
        style: {
          position: "absolute",
          left: 460,
          top: 290,
          width: 100,
          height: 100,
          background: "#E11D48",
          border: { color: "#881337", width: 2 },
        },
      }),
    );

    return createElement(
      Deck,
      { width: 960, height: 540 },
      ...gridSlides,
      detailSlide,
    );
  },
  document() {
    const gridSlides = presetPages.map((group, pageIndex) =>
      slide({
        id: `presets-page-${pageIndex + 1}`,
        background: "#F8FAFC",
      }).append(
        ...group.map((name, itemIndex) => {
          const col = itemIndex % 10;
          const row = Math.floor(itemIndex / 10);
          const wide = itemIndex % 2 === 0;
          return shape({
            id: `preset-${name}`,
            geometry: name,
            style: {
              position: "absolute",
              left: 24 + col * 92 + (wide ? 0 : 8),
              top: 22 + row * 102 + (wide ? 6 : 0),
              width: wide ? 76 : 60,
              height: wide ? 60 : 76,
              background:
                catalogPalette[
                  (pageIndex * 50 + itemIndex) % catalogPalette.length
                ]!,
              border: { color: "#0F172A", width: 1.5 },
            },
          });
        }),
      ),
    );

    const detailSlide = slide({
      id: "adjust-and-custom-parity",
      background: "#F8FAFC",
    }).append(
      shape({
        id: "adj-block-arc",
        geometry: blockArc({ startAngle: 30, endAngle: 250, thickness: 0.35 }),
        style: {
          position: "absolute",
          left: 30,
          top: 30,
          width: 120,
          height: 90,
          background: "#10B981",
          border: { color: "#064E3B", width: 2 },
        },
      }),
      shape({
        id: "adj-star5",
        geometry: preset("star5", { adj: 35000 }),
        style: {
          position: "absolute",
          left: 175,
          top: 30,
          width: 100,
          height: 100,
          background: "#F59E0B",
          border: { color: "#78350F", width: 2 },
        },
      }),
      shape({
        id: "adj-round-rect",
        geometry: preset("roundRect", { adj: 25000 }),
        style: {
          position: "absolute",
          left: 300,
          top: 30,
          width: 130,
          height: 85,
          background: "#3B82F6",
          border: { color: "#1E3A8A", width: 2 },
        },
      }),
      shape({
        id: "adj-callout",
        geometry: preset("wedgeRoundRectCallout", {
          adj1: 28000,
          adj2: 42000,
          adj3: 16667,
        }),
        style: {
          position: "absolute",
          left: 460,
          top: 30,
          width: 140,
          height: 90,
          background: "#8B5CF6",
          border: { color: "#4C1D95", width: 2 },
        },
      }),
      shape({
        id: "adj-connector",
        geometry: preset("bentConnector3", { adj1: 25000 }),
        style: {
          position: "absolute",
          left: 630,
          top: 35,
          width: 120,
          height: 80,
          border: { color: "#DC2626", width: 3 },
        },
      }),
      shape({
        id: "adj-donut",
        geometry: preset("donut", { adj: 30000 }),
        style: {
          position: "absolute",
          left: 780,
          top: 30,
          width: 100,
          height: 100,
          background: "#EC4899",
          border: { color: "#831843", width: 2 },
        },
      }),
      shape({
        id: "adj-chevron",
        geometry: preset("chevron", { adj: 30000 }),
        style: {
          position: "absolute",
          left: 30,
          top: 160,
          width: 130,
          height: 80,
          background: "#0EA5E9",
        },
      }),
      shape({
        id: "adj-circular-arrow",
        geometry: preset("circularArrow", {
          adj1: 15000,
          adj2: 1500000,
          adj3: 20000000,
          adj4: 9000000,
          adj5: 15000,
        }),
        style: {
          position: "absolute",
          left: 190,
          top: 155,
          width: 100,
          height: 100,
          background: "#6366F1",
          border: { color: "#312E81", width: 1.5 },
        },
      }),
      shape({
        id: "shaded-bevel",
        geometry: preset("bevel", { adj: 22000 }),
        style: {
          position: "absolute",
          left: 320,
          top: 165,
          width: 120,
          height: 80,
          background: "#2563EB",
          border: { color: "#0F172A", width: 1.5 },
        },
      }),
      shape({
        id: "shaded-cube",
        geometry: preset("cube", { adj: 25000 }),
        style: {
          position: "absolute",
          left: 470,
          top: 165,
          width: 110,
          height: 85,
          background: "#059669",
          border: { color: "#064E3B", width: 1.5 },
        },
      }),
      shape({
        id: "custom-scaled-bezier",
        geometry: multiSubpathBezierCustom,
        style: {
          position: "absolute",
          left: 30,
          top: 290,
          width: 180,
          height: 90,
          background: "#F97316",
          border: { color: "#0F172A", width: 3 },
        },
      }),
      shape({
        id: "custom-flipped-rotated",
        geometry: multiSubpathBezierCustom,
        style: {
          position: "absolute",
          left: 260,
          top: 290,
          width: 140,
          height: 90,
          background: "#14B8A6",
          flipH: true,
          flipV: true,
          border: { color: "#0F172A", width: 3 },
          rotate: 90,
        },
      }),
      shape({
        id: "custom-evenodd-hole",
        geometry: evenOddRingPath,
        style: {
          position: "absolute",
          left: 460,
          top: 290,
          width: 100,
          height: 100,
          background: "#E11D48",
          border: { color: "#881337", width: 2 },
        },
      }),
    );

    return deck({ width: 960, height: 540 }).append(...gridSlides, detailSlide);
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 5, author);
    assert.equal(validation.slides.length, 5, author);

    const emittedPresets = new Set<string>();
    for (let i = 0; i < 4; i += 1) {
      for (const obj of validation.slides[i]!.objects) {
        if (obj.geometry !== undefined) emittedPresets.add(obj.geometry);
      }
      assert.deepEqual(preview.pages[i]?.errors, [], author);
    }
    assert.equal(
      emittedPresets.size,
      presetGeometryTuple.length,
      `${author}: expected all ${presetGeometryTuple.length} preset geometries`,
    );
    for (const name of presetGeometryTuple) {
      assert.ok(emittedPresets.has(name), `${author}: missing preset ${name}`);
    }

    const detailPage = preview.pages[4]!;
    assert.deepEqual(detailPage.errors, [], author);
    assert.ok(detailPage.png !== undefined, author);
    assert.ok(detailPage.svg !== undefined, author);
    const pixels = PNG.sync.read(Buffer.from(detailPage.png));

    // Donut center (830, 80) is hollow (canvas #F8FAFC), top ring (830, 45) is filled (#EC4899).
    assert.deepEqual(pixelRgb(pixels, 830, 80), [0xf8, 0xfa, 0xfc], author);
    assert.deepEqual(pixelRgb(pixels, 830, 45), [0xec, 0x48, 0x99], author);

    // Even-odd custom ring center (510, 340) is hollow (#F8FAFC), border band (475, 340) is filled (#E11D48).
    assert.deepEqual(pixelRgb(pixels, 510, 340), [0xf8, 0xfa, 0xfc], author);
    assert.deepEqual(pixelRgb(pixels, 475, 340), [0xe1, 0x1d, 0x48], author);

    // Star5 center (225, 80) is filled (#F59E0B), bounding-box corner (180, 35) is canvas (#F8FAFC).
    assert.deepEqual(pixelRgb(pixels, 225, 80), [0xf5, 0x9e, 0x0b], author);
    assert.deepEqual(pixelRgb(pixels, 180, 35), [0xf8, 0xfa, 0xfc], author);

    // Bevel center (380, 205) is base #2563EB while top bevel edge (380, 172) is lightened.
    const bevelCenter = pixelRgb(pixels, 380, 205);
    const bevelTop = pixelRgb(pixels, 380, 172);
    assert.deepEqual(bevelCenter, [0x25, 0x63, 0xeb], author);
    assert.notDeepEqual(bevelTop, bevelCenter, author);
    assert.ok(bevelTop[0] > bevelCenter[0], author);

    // Rotated 90deg custom bezier reports swapped 90x140 bounds and draws filled pixels at its center (330, 335).
    const rotatedCustom = detailPage.nodes.find(
      (n) => n.authorId === "custom-flipped-rotated",
    )!;
    expect(rotatedCustom.bounds?.width).toBeCloseTo(90, 1);
    expect(rotatedCustom.bounds?.height).toBeCloseTo(140, 1);
    assert.deepEqual(pixelRgb(pixels, 330, 335), [0x14, 0xb8, 0xa6], author);
  },
};

export const shapeTextParityScenario: CapabilityScenario = {
  id: "shape-inline-rich-text-and-textrect",
  capability: "shape",
  description:
    "Dual-author Shape inline rich text scenario covering text prop, P/Span/SlideNumber children, non-rect textRect, flipH/rotate, opacity, gradient/image fill, softEdge, boxShadow, multi-script/BiDi, justify, lists, decorations, and verticalAlign.",
  visualCheckpoints: [
    "Page 1 (shape-rich-text-and-presets): #F8FAFC canvas; top-left blue gradient roundRect with shadow and centered 'Round Gradient'; center purple dashed diamond wrapping 'Diamond Rich' (gold underlined) + 'Wraps inner text' (struck-through) inside its inscribed textRect; right horizontally-flipped & +12deg rotated blue rightArrow with upright 'Arrow 1'; bottom-left semi-transparent teal callout (opacity 0.55) with crisp opaque white bottom-aligned 'Opacity Bottom'; bottom-center navy image-filled soft-edged rect with shadowed 'Image Soft'.",
    "Page 2 (shape-multiscript-lists-and-align): white canvas; top-left blue-bordered card with justified CJK paragraph and RTL Arabic text; top-right dotted green ellipse with bullet and numbered list items inside its inscribed textRect; bottom row of three roundRect shapes demonstrating top, middle, and bottom verticalAlign with 'Slide 2'.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all shapes preserve native <p:txBody> rich paragraphs, <a:fld type='slidenum'>, list bullets/numbering, textShadow, flipH, rotate, and preset geometry text placement.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        theme: { colors: { accent1: "#7C3AED", accent2: "#FDE047" } },
      },
      createElement(
        Slide,
        { id: "shape-rich-text-and-presets", background: "#F8FAFC" },
        createElement(
          Shape,
          {
            id: "round-grad-text",
            geometry: "roundRect",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 180,
              height: 76,
              background: linearGradient(135, "#1D4ED8", "#0EA5E9"),
              boxShadow: { blur: 4, distance: 4, angle: 90 },
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
              align: "center",
              verticalAlign: "middle",
            },
          },
          "Round Gradient",
        ),
        createElement(
          Shape,
          {
            id: "diamond-rich-text",
            geometry: "diamond",
            style: {
              position: "absolute",
              left: 224,
              top: 20,
              width: 190,
              height: 130,
              background: "accent1",
              border: { color: "#4C1D95", width: 2, style: "dashed" },
              padding: 4,
            },
          },
          createElement(
            P,
            { style: { fontFamily: jostFont, fontSize: 14, align: "center" } },
            createElement(
              Span,
              { style: { fontWeight: 700, color: "#FFFFFF" } },
              "Diamond ",
            ),
            createElement(
              Span,
              { style: { textDecoration: "underline", color: "accent2" } },
              "Rich",
            ),
          ),
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#FFFFFF",
                textDecoration: "line-through",
                letterSpacing: 1,
                align: "center",
              },
            },
            "Wraps inner text",
          ),
        ),
        createElement(
          Shape,
          {
            id: "flipped-arrow-text",
            geometry: "rightArrow",
            style: {
              position: "absolute",
              left: 436,
              top: 32,
              width: 180,
              height: 90,
              background: "#2563EB",
              flipH: true,
              rotate: 12,
              border: { color: "#0F172A", width: 2 },
            },
          },
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                color: "#FFFFFF",
                align: "center",
              },
            },
            createElement(Span, null, "Arrow "),
            createElement(SlideNumber, null),
          ),
        ),
        createElement(
          Shape,
          {
            id: "opacity-callout",
            geometry: "wedgeRoundRectCallout",
            style: {
              position: "absolute",
              left: 24,
              top: 170,
              width: 220,
              height: 130,
              background: "#0F766E",
              opacity: 0.55,
              padding: [6, 10, 8, 10],
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
              align: "center",
              verticalAlign: "bottom",
            },
          },
          "Opacity Bottom",
        ),
        createElement(
          Shape,
          {
            id: "image-softedge-text",
            style: {
              position: "absolute",
              left: 276,
              top: 180,
              width: 170,
              height: 100,
              background: { image: sampleFillPng },
              softEdge: 5,
              fontFamily: jostFont,
              fontSize: 18,
              color: "#FFFFFF",
              textShadow: { blur: 3, distance: 2, angle: 90, color: "#000000" },
              align: "center",
              verticalAlign: "middle",
            },
          },
          "Image Soft",
        ),
      ),
      createElement(
        Slide,
        { id: "shape-multiscript-lists-and-align", background: "#FFFFFF" },
        createElement(
          Shape,
          {
            id: "cjk-bidi-justify",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 280,
              height: 150,
              background: "#EFF6FF",
              border: { color: "#2563EB", width: 2 },
              padding: 12,
              verticalAlign: "top",
            },
          },
          createElement(
            P,
            {
              style: {
                fontFamily: notoFont,
                fontSize: 13,
                color: "#0F172A",
                align: "justify",
              },
            },
            "形状内嵌富文本支持多语言排版与两端对齐，确保文本在容器区域内自然折行。",
          ),
          createElement(
            P,
            {
              style: {
                fontFamily: notoArabicFont,
                fontSize: 15,
                color: "#1E293B",
                align: "right",
              },
            },
            "الشكل 2026 نص",
          ),
        ),
        createElement(
          Shape,
          {
            id: "ellipse-list",
            geometry: "ellipse",
            style: {
              position: "absolute",
              left: 330,
              top: 24,
              width: 280,
              height: 170,
              background: "#ECFDF5",
              border: { color: "#059669", width: 2, style: "dotted" },
              padding: 6,
            },
          },
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#064E3B",
                list: { type: "bullet" },
              },
            },
            "Bullet item inside ellipse",
          ),
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#064E3B",
                list: { type: "number", startAt: 1 },
              },
            },
            "Numbered item one",
          ),
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                color: "#064E3B",
                list: { type: "number", startAt: 1 },
              },
            },
            "Numbered item two",
          ),
        ),
        ...(["top", "middle", "bottom"] as const).map((vAlign, idx) =>
          createElement(
            Shape,
            {
              key: vAlign,
              id: `valign-${vAlign}`,
              geometry: "roundRect",
              style: {
                position: "absolute",
                left: 24 + idx * 198,
                top: 224,
                width: 176,
                height: 96,
                background: "#F1F5F9",
                border: { color: "#475569", width: 1.5 },
                padding: 8,
                verticalAlign: vAlign,
              },
            },
            createElement(
              P,
              {
                style: {
                  fontFamily: jostFont,
                  fontSize: 14,
                  color: "#0F172A",
                  align: "center",
                },
              },
              createElement(Span, null, `${vAlign} `),
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      theme: { colors: { accent1: "#7C3AED", accent2: "#FDE047" } },
    }).append(
      slide({
        id: "shape-rich-text-and-presets",
        background: "#F8FAFC",
      }).append(
        shape("Round Gradient", {
          id: "round-grad-text",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 180,
            height: 76,
            background: linearGradient(135, "#1D4ED8", "#0EA5E9"),
            boxShadow: { blur: 4, distance: 4, angle: 90 },
            fontFamily: jostFont,
            fontSize: 16,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "middle",
          },
        }),
        shape({
          id: "diamond-rich-text",
          geometry: "diamond",
          style: {
            position: "absolute",
            left: 224,
            top: 20,
            width: 190,
            height: 130,
            background: "accent1",
            border: { color: "#4C1D95", width: 2, style: "dashed" },
            padding: 4,
          },
        }).append(
          paragraph([
            span("Diamond ").style({ fontWeight: 700, color: "#FFFFFF" }),
            span("Rich").style({
              textDecoration: "underline",
              color: "accent2",
            }),
          ]).style({ fontFamily: jostFont, fontSize: 14, align: "center" }),
          paragraph("Wraps inner text").style({
            fontFamily: jostFont,
            fontSize: 13,
            color: "#FFFFFF",
            textDecoration: "line-through",
            letterSpacing: 1,
            align: "center",
          }),
        ),
        shape({
          id: "flipped-arrow-text",
          geometry: "rightArrow",
          style: {
            position: "absolute",
            left: 436,
            top: 32,
            width: 180,
            height: 90,
            background: "#2563EB",
            flipH: true,
            rotate: 12,
            border: { color: "#0F172A", width: 2 },
          },
        }).append(
          paragraph([span("Arrow "), slideNumber()]).style({
            fontFamily: jostFont,
            fontSize: 16,
            color: "#FFFFFF",
            align: "center",
          }),
        ),
        shape("Opacity Bottom", {
          id: "opacity-callout",
          geometry: "wedgeRoundRectCallout",
          style: {
            position: "absolute",
            left: 24,
            top: 170,
            width: 220,
            height: 130,
            background: "#0F766E",
            opacity: 0.55,
            padding: [6, 10, 8, 10],
            fontFamily: jostFont,
            fontSize: 16,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "bottom",
          },
        }),
        shape("Image Soft", {
          id: "image-softedge-text",
          style: {
            position: "absolute",
            left: 276,
            top: 180,
            width: 170,
            height: 100,
            background: { image: sampleFillPng },
            softEdge: 5,
            fontFamily: jostFont,
            fontSize: 18,
            color: "#FFFFFF",
            textShadow: { blur: 3, distance: 2, angle: 90, color: "#000000" },
            align: "center",
            verticalAlign: "middle",
          },
        }),
      ),
      slide({
        id: "shape-multiscript-lists-and-align",
        background: "#FFFFFF",
      }).append(
        shape({
          id: "cjk-bidi-justify",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 280,
            height: 150,
            background: "#EFF6FF",
            border: { color: "#2563EB", width: 2 },
            padding: 12,
            verticalAlign: "top",
          },
        }).append(
          paragraph(
            "形状内嵌富文本支持多语言排版与两端对齐，确保文本在容器区域内自然折行。",
          ).style({
            fontFamily: notoFont,
            fontSize: 13,
            color: "#0F172A",
            align: "justify",
          }),
          paragraph("الشكل 2026 نص").style({
            fontFamily: notoArabicFont,
            fontSize: 15,
            color: "#1E293B",
            align: "right",
          }),
        ),
        shape({
          id: "ellipse-list",
          geometry: "ellipse",
          style: {
            position: "absolute",
            left: 330,
            top: 24,
            width: 280,
            height: 170,
            background: "#ECFDF5",
            border: { color: "#059669", width: 2, style: "dotted" },
            padding: 6,
          },
        }).append(
          paragraph("Bullet item inside ellipse").style({
            fontFamily: jostFont,
            fontSize: 13,
            color: "#064E3B",
            list: { type: "bullet" },
          }),
          paragraph("Numbered item one").style({
            fontFamily: jostFont,
            fontSize: 13,
            color: "#064E3B",
            list: { type: "number", startAt: 1 },
          }),
          paragraph("Numbered item two").style({
            fontFamily: jostFont,
            fontSize: 13,
            color: "#064E3B",
            list: { type: "number", startAt: 1 },
          }),
        ),
        ...(["top", "middle", "bottom"] as const).map((vAlign, idx) =>
          shape({
            id: `valign-${vAlign}`,
            geometry: "roundRect",
            style: {
              position: "absolute",
              left: 24 + idx * 198,
              top: 224,
              width: 176,
              height: 96,
              background: "#F1F5F9",
              border: { color: "#475569", width: 1.5 },
              padding: 8,
              verticalAlign: vAlign,
            },
          }).append(
            paragraph([span(`${vAlign} `), slideNumber()]).style({
              fontFamily: jostFont,
              fontSize: 14,
              color: "#0F172A",
              align: "center",
            }),
          ),
        ),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.equal(validation.slides.length, 2, author);

    const p1 = preview.pages[0]!;
    assert.deepEqual(p1.errors, [], author);
    assert.ok(p1.png !== undefined, author);
    assert.ok(p1.svg !== undefined, author);
    assert.equal(
      p1.nodes.find((n) => n.authorId === "flipped-arrow-text")?.text,
      "Arrow 1",
      author,
    );
    const arrowNode = p1.nodes.find(
      (n) => n.authorId === "flipped-arrow-text",
    )!;
    assert.ok((arrowNode.bounds?.height ?? 0) > 90, author);
    expect(p1.svg).toContain("<linearGradient");
    expect(p1.svg).toContain("<pattern");
    expect(p1.svg).toContain("<feDropShadow");
    expect(p1.svg).toContain("<feGaussianBlur");
    expect(p1.svg).toContain('transform="rotate(12 526 77)"');

    // Opacity callout fill blends with #F8FAFC while white text stays pure [255, 255, 255].
    const p1Pixels = PNG.sync.read(Buffer.from(p1.png));
    const fillSample = pixelRgb(p1Pixels, 40, 185);
    assert.notDeepEqual(fillSample, [0x0f, 0x76, 0x6e], author);
    let hasWhiteCalloutText = false;
    for (let y = 250; y <= 295; y += 1) {
      for (let x = 50; x <= 210; x += 1) {
        const rgb = pixelRgb(p1Pixels, x, y);
        if (rgb[0] === 255 && rgb[1] === 255 && rgb[2] === 255) {
          hasWhiteCalloutText = true;
          break;
        }
      }
      if (hasWhiteCalloutText) break;
    }
    assert.equal(hasWhiteCalloutText, true, author);

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.errors, [], author);
    assert.ok(p2.png !== undefined, author);
    assert.equal(
      p2.nodes.find((n) => n.authorId === "valign-middle")?.text,
      "middle 2",
      author,
    );
  },
};

export const shapeRejectionScenario: CapabilityScenario = {
  id: "shape-preview-rejections",
  capability: "shape",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying Shape inline text enforces only permanent underlying font resource boundaries (missing glyph and color/bitmap font tables) while valid PPTX and valid sibling pages survive.",
  visualCheckpoints: [
    "Page 1 (shape-font-resource-limits): no PNG emitted (preview rejected cleanly with located errors for missing glyph and color font table).",
    "Page 2 (valid-sibling-page): clean white canvas with centered emerald rectangle (#16A34A) and white 'OK' inline text.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 retains native <p:txBody> inline text on all three shapes while Page 2 stays valid.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "shape-font-resource-limits" },
        createElement(
          Shape,
          {
            id: "missing-glyph-inline",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 140,
              height: 50,
              background: "#1E293B",
              fontFamily: jostFont,
              color: "#FFFFFF",
            },
          },
          "汉",
        ),
        createElement(
          Shape,
          {
            id: "missing-glyph-children",
            style: {
              position: "absolute",
              left: 180,
              top: 20,
              width: 140,
              height: 50,
              background: "#1E293B",
            },
          },
          createElement(
            P,
            { style: { fontFamily: jostFont, color: "#FFFFFF" } },
            "☃",
          ),
        ),
        createElement(
          Shape,
          {
            id: "color-font-shape",
            style: {
              position: "absolute",
              left: 20,
              top: 90,
              width: 140,
              height: 50,
              background: "#1E293B",
              fontFamily: colrJostFont,
              color: "#FFFFFF",
            },
          },
          "Color",
        ),
        createElement(
          Shape,
          {
            id: "missing-font-shape",
            style: {
              position: "absolute",
              left: 180,
              top: 90,
              width: 140,
              height: 50,
              background: "#1E293B",
              fontFamily: missingPreviewFont,
              color: "#FFFFFF",
            },
          },
          "No font",
        ),
      ),
      createElement(
        Slide,
        { id: "valid-sibling-page" },
        createElement(
          Shape,
          {
            id: "ok-box",
            style: {
              position: "absolute",
              left: 140,
              top: 75,
              width: 200,
              height: 120,
              background: "#16A34A",
              fontFamily: jostFont,
              fontSize: 22,
              color: "#FFFFFF",
              align: "center",
              verticalAlign: "middle",
            },
          },
          "OK",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "shape-font-resource-limits" }).append(
        shape("汉", {
          id: "missing-glyph-inline",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 140,
            height: 50,
            background: "#1E293B",
            fontFamily: jostFont,
            color: "#FFFFFF",
          },
        }),
        shape({
          id: "missing-glyph-children",
          style: {
            position: "absolute",
            left: 180,
            top: 20,
            width: 140,
            height: 50,
            background: "#1E293B",
          },
        }).append(
          paragraph("☃").style({ fontFamily: jostFont, color: "#FFFFFF" }),
        ),
        shape("Color", {
          id: "color-font-shape",
          style: {
            position: "absolute",
            left: 20,
            top: 90,
            width: 140,
            height: 50,
            background: "#1E293B",
            fontFamily: colrJostFont,
            color: "#FFFFFF",
          },
        }),
        shape("No font", {
          id: "missing-font-shape",
          style: {
            position: "absolute",
            left: 180,
            top: 90,
            width: 140,
            height: 50,
            background: "#1E293B",
            fontFamily: missingPreviewFont,
            color: "#FFFFFF",
          },
        }),
      ),
      slide({ id: "valid-sibling-page" }).append(
        shape("OK", {
          id: "ok-box",
          style: {
            position: "absolute",
            left: 140,
            top: 75,
            width: 200,
            height: 120,
            background: "#16A34A",
            fontFamily: jostFont,
            fontSize: 22,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "middle",
          },
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(validation.slides[0]?.objects.length, 4, author);
    const p1 = preview.pages[0]!;
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.svg, undefined, author);
    assert.equal(p1.errors.length, 4, author);
    for (const [id, code] of [
      ["missing-glyph-inline", "preview-missing-glyph"],
      ["missing-glyph-children", "preview-missing-glyph"],
      ["color-font-shape", "preview-unsupported-font"],
      ["missing-font-shape", "preview-missing-font"],
    ] as const) {
      assert.ok(
        p1.errors.some(
          (d) => d.code === code && d.path.includes(`shape[${id}]`),
        ),
        `${author}: missing ${code} error for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.errors, [], author);
    assert.ok(p2.png !== undefined, author);
    const pixels = PNG.sync.read(Buffer.from(p2.png));
    const corner = (82 * pixels.width + 150) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(corner, corner + 4)],
      [0x16, 0xa3, 0x4a, 255],
    );
  },
};

const cubicOpenArrowPath = path("M 10 80 C 30 10 70 10 90 50 S 130 90 150 30", {
  viewBox: [160, 100],
});

const quadAndArcPath = path(
  "M 10 50 Q 30 10 50 50 T 90 50 A 30 20 0 1 0 10 50",
  { viewBox: [100, 100] },
);

const curvedEvenOddCutout = path(
  "M 10 10 C 40 0 60 0 90 10 L 90 90 Q 50 100 10 90 Z M 30 30 C 45 25 55 25 70 30 L 70 70 Q 50 75 30 70 Z",
  { viewBox: [100, 100], fillRule: "evenodd" },
);

export const shapeShadedFillsAndPathsScenario: CapabilityScenario = {
  id: "shape-shaded-fills-open-beziers-and-flex",
  capability: "shape",
  description:
    "Dual-author Shape scenario covering shaded multi-facet presets with gradient/image fills, open cubic/quadratic/arc paths with fill and start/end arrows, curved evenodd cutouts, capsule borderRadius clamp, large softEdge, and flex-measured inline text.",
  visualCheckpoints: [
    "Page 1 (shaded-fills-and-paths): bevel/cube/can/foldedCorner with gradient and image fills, open cubic Bezier with arrows & fill, quadratic+arc path, curved evenodd cutout, capsule borderRadius, and large softEdge.",
    "Page 2 (flex-inline-text): flex row container holding an auto-sized Shape with top-aligned inline text.",
  ],
  nativeCheckpoints: [
    "Verify native preset geometries, gradients, picture fills, custGeom cubic/quadratic curves, arrows, and flex-resolved shape bounds in PPTX.",
  ],
  jsx() {
    const gradA = linearGradient(45, "#1D4ED8", "#38BDF8");
    const gradB = radialGradient("#F59E0B", "#B45309");
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "shaded-fills-and-paths", background: "#F8FAFC" },
        createElement(Shape, {
          id: "bevel-grad",
          geometry: preset("bevel", { adj: 20000 }),
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 110,
            height: 80,
            background: gradA,
            border: { color: "#0F172A", width: 1.5 },
          },
        }),
        createElement(Shape, {
          id: "cube-radial",
          geometry: preset("cube", { adj: 25000 }),
          style: {
            position: "absolute",
            left: 150,
            top: 20,
            width: 110,
            height: 80,
            background: gradB,
            border: { color: "#78350F", width: 1.5 },
          },
        }),
        createElement(Shape, {
          id: "can-image",
          geometry: preset("can", { adj: 25000 }),
          style: {
            position: "absolute",
            left: 280,
            top: 20,
            width: 100,
            height: 90,
            background: { image: sampleFillPng },
            border: { color: "#1E3A8A", width: 1.5 },
          },
        }),
        createElement(Shape, {
          id: "folded-grad",
          geometry: preset("foldedCorner", { adj: 22000 }),
          style: {
            position: "absolute",
            left: 400,
            top: 20,
            width: 110,
            height: 80,
            background: gradA,
            border: { color: "#0F172A", width: 1.5 },
          },
        }),
        createElement(Shape, {
          id: "open-cubic-arrows",
          geometry: cubicOpenArrowPath,
          style: {
            position: "absolute",
            left: 20,
            top: 130,
            width: 180,
            height: 100,
            background: "#DBEAFE",
            border: {
              color: "#1E293B",
              width: 2,
              startArrow: "triangle",
              endArrow: "stealth",
            },
          },
        }),
        createElement(Shape, {
          id: "quad-arc-path",
          geometry: quadAndArcPath,
          style: {
            position: "absolute",
            left: 220,
            top: 130,
            width: 120,
            height: 100,
            background: "#DCFCE7",
            border: {
              color: "#166534",
              width: 2,
              startArrow: "oval",
              endArrow: "diamond",
            },
          },
        }),
        createElement(Shape, {
          id: "curved-evenodd",
          geometry: curvedEvenOddCutout,
          style: {
            position: "absolute",
            left: 360,
            top: 130,
            width: 100,
            height: 100,
            background: "#E11D48",
            border: { color: "#881337", width: 2 },
          },
        }),
        createElement(Shape, {
          id: "capsule-radius",
          style: {
            position: "absolute",
            left: 480,
            top: 130,
            width: 130,
            height: 44,
            background: "#7C3AED",
            borderRadius: 40,
          },
        }),
        createElement(Shape, {
          id: "large-softedge",
          style: {
            position: "absolute",
            left: 480,
            top: 195,
            width: 100,
            height: 60,
            background: "#0EA5E9",
            softEdge: 40,
          },
        }),
      ),
      createElement(
        Slide,
        { id: "flex-inline-text", background: "#FFFFFF" },
        createElement(
          Box,
          {
            id: "flex-row",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 560,
              display: "flex",
              flexDirection: "row",
              gap: 16,
              padding: 12,
              background: "#F1F5F9",
            },
          },
          createElement(
            Shape,
            {
              id: "flex-auto-shape",
              style: {
                background: "#1D4ED8",
                padding: 8,
                fontFamily: jostFont,
                fontSize: 16,
                color: "#FFFFFF",
                verticalAlign: "top",
              },
            },
            "Flex Shape Top",
          ),
          createElement(Shape, {
            id: "flex-empty-shape",
            style: { width: 80, height: 40, background: "#059669" },
          }),
        ),
      ),
    );
  },
  document() {
    const gradA = linearGradient(45, "#1D4ED8", "#38BDF8");
    const gradB = radialGradient("#F59E0B", "#B45309");
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "shaded-fills-and-paths", background: "#F8FAFC" }).append(
        shape({
          id: "bevel-grad",
          geometry: preset("bevel", { adj: 20000 }),
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 110,
            height: 80,
            background: gradA,
            border: { color: "#0F172A", width: 1.5 },
          },
        }),
        shape({
          id: "cube-radial",
          geometry: preset("cube", { adj: 25000 }),
          style: {
            position: "absolute",
            left: 150,
            top: 20,
            width: 110,
            height: 80,
            background: gradB,
            border: { color: "#78350F", width: 1.5 },
          },
        }),
        shape({
          id: "can-image",
          geometry: preset("can", { adj: 25000 }),
          style: {
            position: "absolute",
            left: 280,
            top: 20,
            width: 100,
            height: 90,
            background: { image: sampleFillPng },
            border: { color: "#1E3A8A", width: 1.5 },
          },
        }),
        shape({
          id: "folded-grad",
          geometry: preset("foldedCorner", { adj: 22000 }),
          style: {
            position: "absolute",
            left: 400,
            top: 20,
            width: 110,
            height: 80,
            background: gradA,
            border: { color: "#0F172A", width: 1.5 },
          },
        }),
        shape({
          id: "open-cubic-arrows",
          geometry: cubicOpenArrowPath,
          style: {
            position: "absolute",
            left: 20,
            top: 130,
            width: 180,
            height: 100,
            background: "#DBEAFE",
            border: {
              color: "#1E293B",
              width: 2,
              startArrow: "triangle",
              endArrow: "stealth",
            },
          },
        }),
        shape({
          id: "quad-arc-path",
          geometry: quadAndArcPath,
          style: {
            position: "absolute",
            left: 220,
            top: 130,
            width: 120,
            height: 100,
            background: "#DCFCE7",
            border: {
              color: "#166534",
              width: 2,
              startArrow: "oval",
              endArrow: "diamond",
            },
          },
        }),
        shape({
          id: "curved-evenodd",
          geometry: curvedEvenOddCutout,
          style: {
            position: "absolute",
            left: 360,
            top: 130,
            width: 100,
            height: 100,
            background: "#E11D48",
            border: { color: "#881337", width: 2 },
          },
        }),
        shape({
          id: "capsule-radius",
          style: {
            position: "absolute",
            left: 480,
            top: 130,
            width: 130,
            height: 44,
            background: "#7C3AED",
            borderRadius: 40,
          },
        }),
        shape({
          id: "large-softedge",
          style: {
            position: "absolute",
            left: 480,
            top: 195,
            width: 100,
            height: 60,
            background: "#0EA5E9",
            softEdge: 40,
          },
        }),
      ),
      slide({ id: "flex-inline-text", background: "#FFFFFF" }).append(
        box({
          id: "flex-row",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 560,
            display: "flex",
            flexDirection: "row",
            gap: 16,
            padding: 12,
            background: "#F1F5F9",
          },
        }).append(
          shape("Flex Shape Top", {
            id: "flex-auto-shape",
            style: {
              background: "#1D4ED8",
              padding: 8,
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
              verticalAlign: "top",
            },
          }),
          shape({
            id: "flex-empty-shape",
            style: { width: 80, height: 40, background: "#059669" },
          }),
        ),
      ),
    );
  },
  assert({ author, resolution, preview }) {
    assert.equal(preview.pages.length, 2, author);
    assert.deepEqual(preview.pages[0]?.errors, [], author);
    assert.deepEqual(preview.pages[1]?.errors, [], author);
    const p1Svg = preview.pages[0]?.svg ?? "";
    expect(p1Svg).toContain("<linearGradient");
    expect(p1Svg).toContain("<radialGradient");
    expect(p1Svg).toContain("<pattern");
    expect(p1Svg).toContain("<mask");
    const p1Pixels = PNG.sync.read(Buffer.from(preview.pages[0].png!));
    // Curved evenodd cutout center at (410, 180) remains #F8FAFC background.
    assert.deepEqual(pixelRgb(p1Pixels, 410, 180), [0xf8, 0xfa, 0xfc], author);
    const flexShapeBox = resolution
      .slide("flex-inline-text")
      .get("flex-row")
      .get("flex-auto-shape").box;
    assert.ok((flexShapeBox?.width ?? 0) > 80, author);
    assert.ok((flexShapeBox?.height ?? 0) > 20, author);
  },
};
