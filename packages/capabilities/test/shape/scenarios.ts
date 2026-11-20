import assert from "node:assert/strict";

import {
  deck,
  paragraph,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import {
  blockArc,
  linearGradient,
  path,
  preset,
  presetGeometryTuple,
  type CustomGeometry,
  type PresetGeometryName,
} from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { P } from "@react-pptx/capabilities/text-elements";
import { Deck, Slide } from "@react-pptx/react";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

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
  capability: "rp:shape",
  description:
    "Dual-author Shape scenario covering rect, ellipse, preset, custom path, fill/border, opacity, rotate, flip, theme tokens, and zIndex.",
  visualCheckpoints: [
    "Page 1 (geometry-and-styles): light gray #F4F6F8 canvas; blue rectangle (#2563EB) with 3px dark border (#1E293B) at top-left; semi-transparent red ellipse (#EF4444, opacity 0.7) overlapping the rectangle's bottom-right corner; 45-degree rotated emerald diamond (#10B981 via ThemeScope accent1) in the center; horizontally flipped purple chevron (#8B5CF6) and orange/navy custom diamond+arc shape on the right.",
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
        theme: { colors: { diamondFill: "#059669" } },
      },
      createElement(
        Slide,
        {
          id: "geometry-and-styles",
          background: "#F4F6F8",
          theme: { colors: { diamondFill: "#10B981" } },
        },
        createElement(Shape, {
          id: "base-rect",
          x: 32,
          y: 36,
          width: 150,
          height: 95,
          fill: "#2563EB",
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "solid",
            zIndex: 1,
          },
        }),
        createElement(Shape, {
          id: "overlap-ellipse",
          geometry: preset("ellipse"),
          x: 110,
          y: 80,
          width: 150,
          height: 95,
          style: {
            backgroundColor: "#EF4444",
            borderColor: "#7F1D1D",
            borderWidth: 2,
            opacity: 0.7,
            zIndex: 2,
          },
        }),
        createElement(Shape, {
          id: "rotated-square",
          geometry: preset("rect"),
          x: 300,
          y: 70,
          width: 84,
          height: 84,
          fill: "diamondFill",
          style: {
            borderColor: "#065F46",
            borderWidth: 2,
            rotate: 45,
          },
        }),
        createElement(Shape, {
          id: "flipped-chevron",
          geometry: ribbonPath,
          x: 430,
          y: 45,
          width: 140,
          height: 80,
          fill: "#8B5CF6",
          flipH: true,
          style: {
            borderColor: "#4C1D95",
            borderWidth: 2,
          },
        }),
        createElement(Shape, {
          id: "custom-dual-path",
          geometry: dualSubpathCustom,
          x: 440,
          y: 165,
          width: 120,
          height: 110,
          fill: "#F97316",
          flipV: true,
          style: {
            borderColor: "#0F172A",
            borderWidth: 3,
          },
        }),
      ),
      createElement(
        Slide,
        { id: "layer-order", background: "#FFFFFF" },
        createElement(Shape, {
          id: "front-gold",
          x: 80,
          y: 70,
          width: 180,
          height: 110,
          fill: "#F59E0B",
          style: { zIndex: 5 },
        }),
        createElement(Shape, {
          id: "back-slate",
          x: 140,
          y: 110,
          width: 180,
          height: 110,
          fill: "#334155",
          style: { zIndex: 1 },
        }),
        createElement(Shape, {
          id: "borderless-teal",
          x: 380,
          y: 90,
          width: 160,
          height: 100,
          fill: "#0D9488",
          style: { borderColor: "#000000", borderWidth: 0 },
        }),
      ),
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      theme: { colors: { diamondFill: "#059669" } },
    }).append(
      slide({
        id: "geometry-and-styles",
        background: "#F4F6F8",
        theme: { colors: { diamondFill: "#10B981" } },
      }).append(
        shape({
          id: "base-rect",
          x: 32,
          y: 36,
          width: 150,
          height: 95,
          fill: "#2563EB",
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "solid",
            zIndex: 1,
          },
        }),
        shape({
          id: "overlap-ellipse",
          geometry: preset("ellipse"),
          x: 110,
          y: 80,
          width: 150,
          height: 95,
          style: {
            backgroundColor: "#EF4444",
            borderColor: "#7F1D1D",
            borderWidth: 2,
            opacity: 0.7,
            zIndex: 2,
          },
        }),
        shape({
          id: "rotated-square",
          geometry: preset("rect"),
          x: 300,
          y: 70,
          width: 84,
          height: 84,
          fill: "diamondFill",
          style: {
            borderColor: "#065F46",
            borderWidth: 2,
            rotate: 45,
          },
        }),
        shape({
          id: "flipped-chevron",
          geometry: ribbonPath,
          x: 430,
          y: 45,
          width: 140,
          height: 80,
          fill: "#8B5CF6",
          flipH: true,
          style: {
            borderColor: "#4C1D95",
            borderWidth: 2,
          },
        }),
        shape({
          id: "custom-dual-path",
          geometry: dualSubpathCustom,
          x: 440,
          y: 165,
          width: 120,
          height: 110,
          fill: "#F97316",
          flipV: true,
          style: {
            borderColor: "#0F172A",
            borderWidth: 3,
          },
        }),
      ),
      slide({ id: "layer-order", background: "#FFFFFF" }).append(
        shape({
          id: "front-gold",
          x: 80,
          y: 70,
          width: 180,
          height: 110,
          fill: "#F59E0B",
          style: { zIndex: 5 },
        }),
        shape({
          id: "back-slate",
          x: 140,
          y: 110,
          width: 180,
          height: 110,
          fill: "#334155",
          style: { zIndex: 1 },
        }),
        shape({
          id: "borderless-teal",
          x: 380,
          y: 90,
          width: 160,
          height: 100,
          fill: "#0D9488",
          style: { borderColor: "#000000", borderWidth: 0 },
        }),
      ),
    );
  },
  assert({ author, resolution, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.equal(validation.slides.length, 2, author);
    assert.equal(validation.slides[0]?.objects.length, 5, author);

    const p1 = preview.pages[0]!;
    assert.deepEqual(p1.diagnostics, [], author);
    assert.ok(p1.png !== undefined, author);
    const rotatedNode = p1.nodes.find((n) => n.authorId === "rotated-square")!;
    expect(rotatedNode.screenBox?.width).toBeCloseTo(84 * Math.SQRT2, 1);
    expect(rotatedNode.screenBox?.height).toBeCloseTo(84 * Math.SQRT2, 1);
    assert.equal(
      resolution.query
        .getSlide("geometry-and-styles")
        .get("overlap-ellipse")
        .observe().objects[0]?.geometry,
      "ellipse",
    );

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
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
  capability: "rp:shape",
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
            x: 24 + col * 92 + (wide ? 0 : 8),
            y: 22 + row * 102 + (wide ? 6 : 0),
            width: wide ? 76 : 60,
            height: wide ? 60 : 76,
            fill: catalogPalette[
              (pageIndex * 50 + itemIndex) % catalogPalette.length
            ]!,
            style: { borderColor: "#0F172A", borderWidth: 1.5 },
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
        x: 30,
        y: 30,
        width: 120,
        height: 90,
        fill: "#10B981",
        style: { borderColor: "#064E3B", borderWidth: 2 },
      }),
      createElement(Shape, {
        id: "adj-star5",
        geometry: preset("star5", { adj: 35000 }),
        x: 175,
        y: 30,
        width: 100,
        height: 100,
        fill: "#F59E0B",
        style: { borderColor: "#78350F", borderWidth: 2 },
      }),
      createElement(Shape, {
        id: "adj-round-rect",
        geometry: preset("roundRect", { adj: 25000 }),
        x: 300,
        y: 30,
        width: 130,
        height: 85,
        fill: "#3B82F6",
        style: { borderColor: "#1E3A8A", borderWidth: 2 },
      }),
      createElement(Shape, {
        id: "adj-callout",
        geometry: preset("wedgeRoundRectCallout", {
          adj1: 28000,
          adj2: 42000,
          adj3: 16667,
        }),
        x: 460,
        y: 30,
        width: 140,
        height: 90,
        fill: "#8B5CF6",
        style: { borderColor: "#4C1D95", borderWidth: 2 },
      }),
      createElement(Shape, {
        id: "adj-connector",
        geometry: preset("bentConnector3", { adj1: 25000 }),
        x: 630,
        y: 35,
        width: 120,
        height: 80,
        style: { borderColor: "#DC2626", borderWidth: 3 },
      }),
      createElement(Shape, {
        id: "adj-donut",
        geometry: preset("donut", { adj: 30000 }),
        x: 780,
        y: 30,
        width: 100,
        height: 100,
        fill: "#EC4899",
        style: { borderColor: "#831843", borderWidth: 2 },
      }),
      createElement(Shape, {
        id: "adj-chevron",
        geometry: preset("chevron", { adj: 30000 }),
        x: 30,
        y: 160,
        width: 130,
        height: 80,
        fill: "#0EA5E9",
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
        x: 190,
        y: 155,
        width: 100,
        height: 100,
        fill: "#6366F1",
        style: { borderColor: "#312E81", borderWidth: 1.5 },
      }),
      createElement(Shape, {
        id: "shaded-bevel",
        geometry: preset("bevel", { adj: 22000 }),
        x: 320,
        y: 165,
        width: 120,
        height: 80,
        fill: "#2563EB",
        style: { borderColor: "#0F172A", borderWidth: 1.5 },
      }),
      createElement(Shape, {
        id: "shaded-cube",
        geometry: preset("cube", { adj: 25000 }),
        x: 470,
        y: 165,
        width: 110,
        height: 85,
        fill: "#059669",
        style: { borderColor: "#064E3B", borderWidth: 1.5 },
      }),
      createElement(Shape, {
        id: "custom-scaled-bezier",
        geometry: multiSubpathBezierCustom,
        x: 30,
        y: 290,
        width: 180,
        height: 90,
        fill: "#F97316",
        style: { borderColor: "#0F172A", borderWidth: 3 },
      }),
      createElement(Shape, {
        id: "custom-flipped-rotated",
        geometry: multiSubpathBezierCustom,
        x: 260,
        y: 290,
        width: 140,
        height: 90,
        fill: "#14B8A6",
        flipH: true,
        flipV: true,
        style: { borderColor: "#0F172A", borderWidth: 3, rotate: 90 },
      }),
      createElement(Shape, {
        id: "custom-evenodd-hole",
        geometry: evenOddRingPath,
        x: 460,
        y: 290,
        width: 100,
        height: 100,
        fill: "#E11D48",
        style: { borderColor: "#881337", borderWidth: 2 },
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
            x: 24 + col * 92 + (wide ? 0 : 8),
            y: 22 + row * 102 + (wide ? 6 : 0),
            width: wide ? 76 : 60,
            height: wide ? 60 : 76,
            fill: catalogPalette[
              (pageIndex * 50 + itemIndex) % catalogPalette.length
            ]!,
            style: { borderColor: "#0F172A", borderWidth: 1.5 },
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
        x: 30,
        y: 30,
        width: 120,
        height: 90,
        fill: "#10B981",
        style: { borderColor: "#064E3B", borderWidth: 2 },
      }),
      shape({
        id: "adj-star5",
        geometry: preset("star5", { adj: 35000 }),
        x: 175,
        y: 30,
        width: 100,
        height: 100,
        fill: "#F59E0B",
        style: { borderColor: "#78350F", borderWidth: 2 },
      }),
      shape({
        id: "adj-round-rect",
        geometry: preset("roundRect", { adj: 25000 }),
        x: 300,
        y: 30,
        width: 130,
        height: 85,
        fill: "#3B82F6",
        style: { borderColor: "#1E3A8A", borderWidth: 2 },
      }),
      shape({
        id: "adj-callout",
        geometry: preset("wedgeRoundRectCallout", {
          adj1: 28000,
          adj2: 42000,
          adj3: 16667,
        }),
        x: 460,
        y: 30,
        width: 140,
        height: 90,
        fill: "#8B5CF6",
        style: { borderColor: "#4C1D95", borderWidth: 2 },
      }),
      shape({
        id: "adj-connector",
        geometry: preset("bentConnector3", { adj1: 25000 }),
        x: 630,
        y: 35,
        width: 120,
        height: 80,
        style: { borderColor: "#DC2626", borderWidth: 3 },
      }),
      shape({
        id: "adj-donut",
        geometry: preset("donut", { adj: 30000 }),
        x: 780,
        y: 30,
        width: 100,
        height: 100,
        fill: "#EC4899",
        style: { borderColor: "#831843", borderWidth: 2 },
      }),
      shape({
        id: "adj-chevron",
        geometry: preset("chevron", { adj: 30000 }),
        x: 30,
        y: 160,
        width: 130,
        height: 80,
        fill: "#0EA5E9",
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
        x: 190,
        y: 155,
        width: 100,
        height: 100,
        fill: "#6366F1",
        style: { borderColor: "#312E81", borderWidth: 1.5 },
      }),
      shape({
        id: "shaded-bevel",
        geometry: preset("bevel", { adj: 22000 }),
        x: 320,
        y: 165,
        width: 120,
        height: 80,
        fill: "#2563EB",
        style: { borderColor: "#0F172A", borderWidth: 1.5 },
      }),
      shape({
        id: "shaded-cube",
        geometry: preset("cube", { adj: 25000 }),
        x: 470,
        y: 165,
        width: 110,
        height: 85,
        fill: "#059669",
        style: { borderColor: "#064E3B", borderWidth: 1.5 },
      }),
      shape({
        id: "custom-scaled-bezier",
        geometry: multiSubpathBezierCustom,
        x: 30,
        y: 290,
        width: 180,
        height: 90,
        fill: "#F97316",
        style: { borderColor: "#0F172A", borderWidth: 3 },
      }),
      shape({
        id: "custom-flipped-rotated",
        geometry: multiSubpathBezierCustom,
        x: 260,
        y: 290,
        width: 140,
        height: 90,
        fill: "#14B8A6",
        flipH: true,
        flipV: true,
        style: { borderColor: "#0F172A", borderWidth: 3, rotate: 90 },
      }),
      shape({
        id: "custom-evenodd-hole",
        geometry: evenOddRingPath,
        x: 460,
        y: 290,
        width: 100,
        height: 100,
        fill: "#E11D48",
        style: { borderColor: "#881337", borderWidth: 2 },
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
      assert.deepEqual(preview.pages[i]?.diagnostics, [], author);
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
    assert.deepEqual(detailPage.diagnostics, [], author);
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

    // Rotated 90deg custom bezier reports swapped 90x140 screenBox and draws filled pixels at its center (330, 335).
    const rotatedCustom = detailPage.nodes.find(
      (n) => n.authorId === "custom-flipped-rotated",
    )!;
    expect(rotatedCustom.screenBox?.width).toBeCloseTo(90, 1);
    expect(rotatedCustom.screenBox?.height).toBeCloseTo(140, 1);
    assert.deepEqual(pixelRgb(pixels, 330, 335), [0x14, 0xb8, 0xa6], author);
  },
};

export const shapeRejectionScenario: CapabilityScenario = {
  id: "shape-preview-rejections",
  capability: "rp:shape",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying unsupported Shape features emit located page errors while valid PPTX and valid sibling pages survive.",
  visualCheckpoints: [
    "Page 1 (unsupported-shape-features): no PNG emitted (preview rejected cleanly with 8 located errors).",
    "Page 2 (valid-sibling-page): clean white canvas with centered emerald rectangle (#16A34A).",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 still retains arrows, dashed lines, rounded rect, shadow, gradient, and inline text.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "unsupported-shape-features" },
        createElement(Shape, {
          id: "arrow-start",
          x: 10,
          y: 10,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", startArrow: "triangle" },
        }),
        createElement(Shape, {
          id: "arrow-end",
          x: 10,
          y: 40,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", endArrow: "arrow" },
        }),
        createElement(Shape, {
          id: "dashed",
          x: 10,
          y: 70,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", borderStyle: "dashed" },
        }),
        createElement(Shape, {
          id: "rounded",
          x: 10,
          y: 100,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { borderRadius: 6 },
        }),
        createElement(Shape, {
          id: "shadow",
          x: 120,
          y: 10,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        createElement(Shape, {
          id: "soft",
          x: 120,
          y: 40,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { softEdge: 4 },
        }),
        createElement(Shape, {
          id: "gradient",
          x: 120,
          y: 70,
          width: 80,
          height: 20,
          style: { backgroundImage: linearGradient(90, "#111111", "#888888") },
        }),
        createElement(
          Shape,
          { id: "inline-text", x: 120, y: 100, width: 80, height: 30 },
          createElement(P, null, "Text"),
        ),
      ),
      createElement(
        Slide,
        { id: "valid-sibling-page" },
        createElement(Shape, {
          id: "ok-box",
          x: 140,
          y: 75,
          width: 200,
          height: 120,
          fill: "#16A34A",
        }),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "unsupported-shape-features" }).append(
        shape({
          id: "arrow-start",
          x: 10,
          y: 10,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", startArrow: "triangle" },
        }),
        shape({
          id: "arrow-end",
          x: 10,
          y: 40,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", endArrow: "arrow" },
        }),
        shape({
          id: "dashed",
          x: 10,
          y: 70,
          width: 80,
          height: 20,
          style: { borderColor: "#000000", borderStyle: "dashed" },
        }),
        shape({
          id: "rounded",
          x: 10,
          y: 100,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { borderRadius: 6 },
        }),
        shape({
          id: "shadow",
          x: 120,
          y: 10,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        shape({
          id: "soft",
          x: 120,
          y: 40,
          width: 80,
          height: 20,
          fill: "#112233",
          style: { softEdge: 4 },
        }),
        shape({
          id: "gradient",
          x: 120,
          y: 70,
          width: 80,
          height: 20,
          style: { backgroundImage: linearGradient(90, "#111111", "#888888") },
        }),
        shape({
          id: "inline-text",
          x: 120,
          y: 100,
          width: 80,
          height: 30,
        }).append(paragraph("Text")),
      ),
      slide({ id: "valid-sibling-page" }).append(
        shape({
          id: "ok-box",
          x: 140,
          y: 75,
          width: 200,
          height: 120,
          fill: "#16A34A",
        }),
      ),
    );
  },
  assert({ author, preview }) {
    const p1 = preview.pages[0]!;
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.diagnostics.length, 8, author);
    for (const id of [
      "arrow-start",
      "arrow-end",
      "dashed",
      "rounded",
      "shadow",
      "soft",
      "gradient",
      "inline-text",
    ]) {
      assert.ok(
        p1.diagnostics.some(
          (d) => d.severity === "error" && d.path.includes(`shape[${id}]`),
        ),
        `${author}: missing diagnostic for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    assert.ok(p2.png !== undefined, author);
    const pixels = PNG.sync.read(Buffer.from(p2.png));
    const center = (135 * pixels.width + 240) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(center, center + 4)],
      [0x16, 0xa3, 0x4a, 255],
    );
  },
};
