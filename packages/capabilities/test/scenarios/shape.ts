import assert from "node:assert/strict";

import { Deck, Slide } from "@react-pptx/core";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import { deck, paragraph, shape, slide } from "../../src/document.ts";
import {
  blockArc,
  path,
  preset,
  type CustomGeometry,
} from "../../src/geometry.ts";
import { Shape } from "../../src/shape-elements.ts";
import { P } from "../../src/text-elements.ts";
import { linearGradient } from "../../src/visual.ts";
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
  assert({ author, preview, validation }) {
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
      preview.build.query
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

export const shapeRejectionScenario: CapabilityScenario = {
  id: "shape-preview-rejections",
  capability: "rp:shape",
  description:
    "Dual-author negative scenario verifying unsupported Shape features emit located page errors while valid PPTX and valid sibling pages survive.",
  visualCheckpoints: [
    "Page 1 (unsupported-shape-features): no PNG emitted (preview rejected cleanly with 9 located errors).",
    "Page 2 (valid-sibling-page): clean white canvas with centered emerald rectangle (#16A34A).",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 still retains arrows, dashed lines, rounded rect, shadow, gradient, inline text, and star5.",
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
        createElement(Shape, {
          id: "star",
          geometry: "star5",
          x: 240,
          y: 10,
          width: 60,
          height: 60,
          fill: "#F59E0B",
        }),
        createElement(Shape, {
          id: "arc",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 240,
          y: 90,
          width: 60,
          height: 60,
          fill: "#10B981",
        }),
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
        shape({
          id: "star",
          geometry: "star5",
          x: 240,
          y: 10,
          width: 60,
          height: 60,
          fill: "#F59E0B",
        }),
        shape({
          id: "arc",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 240,
          y: 90,
          width: 60,
          height: 60,
          fill: "#10B981",
        }),
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
    assert.equal(p1.diagnostics.length, 10, author);
    for (const id of [
      "arrow-start",
      "arrow-end",
      "dashed",
      "rounded",
      "shadow",
      "soft",
      "gradient",
      "inline-text",
      "star",
      "arc",
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
