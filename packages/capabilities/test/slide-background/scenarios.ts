import assert from "node:assert/strict";

import { deck, shape, slide } from "@react-pptx/capabilities/document";
import { linearGradient } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Deck, Slide } from "@react-pptx/react";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";

function solidSamplePng(): Uint8Array {
  const image = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < image.data.length; i += 4) {
    image.data[i] = 0x22;
    image.data[i + 1] = 0x44;
    image.data[i + 2] = 0x88;
    image.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(image));
}

export const slideBackgroundPositiveScenario: CapabilityScenario = {
  id: "slide-background-variants",
  capability: "slide-background",
  description:
    "Dual-author slide-background scenario covering default white canvas, solid hex background, and local ThemeScope background token.",
  visualCheckpoints: [
    "Page 1 (default-white): clean white canvas with centered dark slate square (#0F172A).",
    "Page 2 (solid-navy): deep navy canvas (#1E3A8A) with centered crisp white ellipse.",
    "Page 3 (theme-cream): warm cream canvas (#FEF3C7 via ThemeScope bg1) with amber rectangle (#D97706).",
  ],
  nativeCheckpoints: [
    "Verify Slide 1 omits <p:bg> while Slides 2 and 3 emit <p:bg><p:bgPr><a:solidFill> in PowerPoint/WPS.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 480,
        height: 270,
        theme: { colors: { surfaceBg: "#FFFFFF", cardAccent: "#3B82F6" } },
      },
      createElement(
        Slide,
        { id: "default-white" },
        createElement(Shape, {
          id: "center-rect",
          x: 190,
          y: 85,
          width: 100,
          height: 100,
          fill: "#0F172A",
        }),
      ),
      createElement(
        Slide,
        { id: "solid-navy", background: { color: "#1E3A8A", alpha: 1 } },
        createElement(Shape, {
          id: "center-oval",
          geometry: "ellipse",
          x: 160,
          y: 75,
          width: 160,
          height: 120,
          fill: "#FFFFFF",
        }),
      ),
      createElement(
        Slide,
        {
          id: "theme-cream",
          background: "surfaceBg",
          theme: { colors: { surfaceBg: "#FEF3C7", cardAccent: "#D97706" } },
        },
        createElement(Shape, {
          id: "amber-card",
          x: 120,
          y: 75,
          width: 240,
          height: 120,
          fill: "cardAccent",
        }),
      ),
    );
  },
  document() {
    return deck({
      width: 480,
      height: 270,
      theme: { colors: { surfaceBg: "#FFFFFF", cardAccent: "#3B82F6" } },
    }).append(
      slide({ id: "default-white" }).append(
        shape({
          id: "center-rect",
          x: 190,
          y: 85,
          width: 100,
          height: 100,
          fill: "#0F172A",
        }),
      ),
      slide({
        id: "solid-navy",
        background: { color: "#1E3A8A", alpha: 1 },
      }).append(
        shape({
          id: "center-oval",
          geometry: "ellipse",
          x: 160,
          y: 75,
          width: 160,
          height: 120,
          fill: "#FFFFFF",
        }),
      ),
      slide({
        id: "theme-cream",
        background: "surfaceBg",
        theme: { colors: { surfaceBg: "#FEF3C7", cardAccent: "#D97706" } },
      }).append(
        shape({
          id: "amber-card",
          x: 120,
          y: 75,
          width: 240,
          height: 120,
          fill: "cardAccent",
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.equal(validation.slides[0]?.background, undefined, author);
    assert.equal(validation.slides[1]?.background?.type, "solid", author);
    assert.equal(validation.slides[2]?.background?.type, "solid", author);

    const expectedCorners: readonly (readonly [
      number,
      number,
      number,
      number,
    ])[] = [
      [255, 255, 255, 255],
      [0x1e, 0x3a, 0x8a, 255],
      [0xfe, 0xf3, 0xc7, 255],
    ];
    for (let i = 0; i < 3; i++) {
      const page = preview.pages[i]!;
      assert.deepEqual(page.diagnostics, [], author);
      const png = PNG.sync.read(Buffer.from(page.png!));
      assert.deepEqual([...png.data.subarray(0, 4)], [...expectedCorners[i]!]);
    }
  },
};

export const slideBackgroundRejectionScenario: CapabilityScenario = {
  id: "slide-background-rejections",
  capability: "slide-background",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying gradient and image backgrounds emit located per-page preview errors without blocking valid PPTX or valid pages.",
  visualCheckpoints: [
    "Page 1 & Page 2: preview rejected cleanly (`png: undefined`) with located `/background` errors.",
    "Page 3 (solid-ok): dark emerald canvas (#065F46) with centered white rectangle.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 retains gradient fill and Page 2 retains embedded background image.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(Slide, {
        id: "gradient-bg",
        background: linearGradient(45, "#0F172A", "#38BDF8"),
      }),
      createElement(Slide, {
        id: "image-bg",
        background: { image: solidSamplePng() },
      }),
      createElement(
        Slide,
        { id: "solid-ok", background: "#065F46" },
        createElement(Shape, {
          id: "white-card",
          x: 140,
          y: 75,
          width: 200,
          height: 120,
          fill: "#FFFFFF",
        }),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({
        id: "gradient-bg",
        background: linearGradient(45, "#0F172A", "#38BDF8"),
      }),
      slide({
        id: "image-bg",
        background: { image: solidSamplePng() },
      }),
      slide({ id: "solid-ok", background: "#065F46" }).append(
        shape({
          id: "white-card",
          x: 140,
          y: 75,
          width: 200,
          height: 120,
          fill: "#FFFFFF",
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(validation.slides[0]?.background?.type, "gradient", author);
    assert.equal(validation.slides[1]?.background?.type, "image", author);
    assert.equal(validation.slides[2]?.background?.type, "solid", author);

    assert.equal(preview.pages[0]?.png, undefined, author);
    assert.equal(
      preview.pages[0]?.diagnostics[0]?.path,
      "deck/slide[gradient-bg]/background",
      author,
    );
    assert.equal(preview.pages[1]?.png, undefined, author);
    assert.equal(
      preview.pages[1]?.diagnostics[0]?.path,
      "deck/slide[image-bg]/background",
      author,
    );
    assert.ok(preview.pages[2]?.png !== undefined, author);
    assert.deepEqual(preview.pages[2]?.diagnostics, [], author);
  },
};
