import assert from "node:assert/strict";
import { createRequire } from "node:module";

import { deck, shape, slide } from "@react-pptx/capabilities/document";
import { linearGradient, radialGradient } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Deck, Slide } from "@react-pptx/react";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

const require = createRequire(import.meta.url);
const { encode: encodeJpeg } = require("jpeg-js") as typeof import("jpeg-js");

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

function solidSampleJpeg(): Uint8Array {
  const image = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < image.data.length; i += 4) {
    image.data[i] = 0xc2;
    image.data[i + 1] = 0x41;
    image.data[i + 2] = 0x0c;
    image.data[i + 3] = 255;
  }
  return new Uint8Array(encodeJpeg(image, 100).data);
}

// Valid 1x1 teal (#0D9488) GIF89a.
const sampleGifBytes = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0x80, 0, 0, 0x0d, 0x94, 0x88,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0x00, 0x02, 0x02, 0x44, 0x01,
  0x00, 0x3b,
]);

function createCorruptPng(validPng: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(validPng);
  const idat = Buffer.from(bytes).indexOf(Buffer.from("IDAT"));
  bytes[idat + 6] = (bytes[idat + 6] ?? 0) ^ 0xff;
  bytes[idat + 7] = (bytes[idat + 7] ?? 0) ^ 0xff;
  return bytes;
}

function createCorruptJpeg(validJpeg: Uint8Array): Uint8Array {
  const bytes = new Uint8Array(validJpeg);
  const sos = Buffer.from(bytes).indexOf(Buffer.from([0xff, 0xda]));
  const sosLen = ((bytes[sos + 2] ?? 0) << 8) + (bytes[sos + 3] ?? 0);
  const scanStart = sos + 2 + sosLen;
  for (let i = scanStart; i < bytes.length - 4; i++) {
    bytes[i] = 0x55;
  }
  return bytes;
}

const corruptGifBytes = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 2, 0, 0x80, 0, 0, 0xff, 0xff, 0xff,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 2, 0, 2, 0, 0x00, 0x02, 0x02, 0xbc, 0x00,
  0x00, 0x3b,
]);

const pngBgBytes = solidSamplePng();
const jpegBgBytes = solidSampleJpeg();
const corruptPngBytes = createCorruptPng(pngBgBytes);
const corruptJpegBytes = createCorruptJpeg(jpegBgBytes);

export const slideBackgroundPositiveScenario: CapabilityScenario = {
  id: "slide-background-variants",
  capability: "slide-background",
  description:
    "Dual-author slide-background scenario covering default white, solid hex, local ThemeScope token, linear/radial gradients with alpha, and PNG/JPEG/GIF image backgrounds.",
  visualCheckpoints: [
    "Page 1 (default-white): clean white canvas with centered dark slate square (#0F172A).",
    "Page 2 (solid-navy): deep navy canvas (#1E3A8A) with centered crisp white ellipse.",
    "Page 3 (theme-cream): warm cream canvas (#FEF3C7 via ThemeScope) with amber rectangle (#D97706).",
    "Page 4 (linear-gradient-bg): horizontal 90deg linear gradient from local ThemeScope surfaceBg (#0F172A) to #38BDF8.",
    "Page 5 (radial-gradient-bg): centered radial gradient from #FEF3C7 at center to #1E3A8A at corners.",
    "Page 6 (png-image-bg): deep blue PNG image background (#224488).",
    "Page 7 (jpeg-image-bg): warm rust JPEG image background (#C2410C).",
    "Page 8 (gif-image-bg): teal GIF89a image background (#0D9488).",
  ],
  nativeCheckpoints: [
    "Verify Slide 1 omits <p:bg>, Slides 2-3 emit <a:solidFill>, Slides 4-5 emit <a:gradFill>, and Slides 6-8 emit <a:blipFill> with PNG/JPEG/GIF media.",
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
      createElement(
        Slide,
        {
          id: "linear-gradient-bg",
          theme: { colors: { surfaceBg: "#0F172A" } },
          background: linearGradient(
            90,
            { color: "surfaceBg", position: 0 },
            { color: { color: "#38BDF8", alpha: 0.85 }, position: 1 },
          ),
        },
        createElement(Shape, {
          id: "grad-badge",
          x: 190,
          y: 105,
          width: 100,
          height: 60,
          fill: "#FFFFFF",
        }),
      ),
      createElement(Slide, {
        id: "radial-gradient-bg",
        background: radialGradient(
          { color: "#FEF3C7", position: 0 },
          { color: "#1E3A8A", position: 1 },
        ),
      }),
      createElement(Slide, {
        id: "png-image-bg",
        background: { image: pngBgBytes },
      }),
      createElement(Slide, {
        id: "jpeg-image-bg",
        background: { image: jpegBgBytes },
      }),
      createElement(Slide, {
        id: "gif-image-bg",
        background: { image: sampleGifBytes },
      }),
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
      slide({
        id: "linear-gradient-bg",
        theme: { colors: { surfaceBg: "#0F172A" } },
        background: linearGradient(
          90,
          { color: "surfaceBg", position: 0 },
          { color: { color: "#38BDF8", alpha: 0.85 }, position: 1 },
        ),
      }).append(
        shape({
          id: "grad-badge",
          x: 190,
          y: 105,
          width: 100,
          height: 60,
          fill: "#FFFFFF",
        }),
      ),
      slide({
        id: "radial-gradient-bg",
        background: radialGradient(
          { color: "#FEF3C7", position: 0 },
          { color: "#1E3A8A", position: 1 },
        ),
      }),
      slide({ id: "png-image-bg", background: { image: pngBgBytes } }),
      slide({ id: "jpeg-image-bg", background: { image: jpegBgBytes } }),
      slide({ id: "gif-image-bg", background: { image: sampleGifBytes } }),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 8, author);
    assert.equal(validation.slides[0]?.background, undefined, author);
    assert.equal(validation.slides[1]?.background?.type, "solid", author);
    assert.equal(validation.slides[2]?.background?.type, "solid", author);
    assert.equal(validation.slides[3]?.background?.type, "gradient", author);
    assert.equal(validation.slides[4]?.background?.type, "gradient", author);
    assert.equal(validation.slides[5]?.background?.type, "image", author);
    assert.equal(validation.slides[6]?.background?.type, "image", author);
    assert.equal(validation.slides[7]?.background?.type, "image", author);

    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, [], `${author}: ${page.id}`);
      assert.ok(page.png !== undefined, `${author}: ${page.id}`);
    }

    const px = (pageIdx: number, x: number, y: number) => {
      const png = PNG.sync.read(Buffer.from(preview.pages[pageIdx]!.png!));
      const idx = (y * png.width + x) * 4;
      return [...png.data.subarray(idx, idx + 4)];
    };

    assert.deepEqual(px(0, 0, 0), [255, 255, 255, 255]);
    assert.deepEqual(px(1, 0, 0), [0x1e, 0x3a, 0x8a, 255]);
    assert.deepEqual(px(2, 0, 0), [0xfe, 0xf3, 0xc7, 255]);

    // Linear gradient: left edge is dark slate, right edge is sky blue.
    expect(preview.pages[3]?.svg).toContain("<linearGradient");
    const gradLeft = px(3, 2, 135);
    const gradRight = px(3, 477, 135);
    assert.ok(gradLeft[2]! < 100 && gradRight[2]! > 200, author);

    // Radial gradient: center (240, 135) is warm cream, corner (2, 2) is deep navy.
    expect(preview.pages[4]?.svg).toContain("<radialGradient");
    const radCenter = px(4, 240, 135);
    const radCorner = px(4, 2, 2);
    assert.ok(radCenter[0]! > 230 && radCorner[0]! < 50, author);

    // Image backgrounds (PNG, JPEG, GIF)
    assert.deepEqual(px(5, 240, 135), [0x22, 0x44, 0x88, 255]);
    const jpegCenter = px(6, 240, 135);
    assert.ok(
      jpegCenter[0]! > 170 && jpegCenter[1]! < 90 && jpegCenter[2]! < 40,
      author,
    );
    assert.deepEqual(px(7, 240, 135), [0x0d, 0x94, 0x88, 255]);
  },
};

export const slideBackgroundRejectionScenario: CapabilityScenario = {
  id: "slide-background-rejections",
  capability: "slide-background",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying corrupt PNG, JPEG, and GIF background image streams emit located preview-corrupt-image errors without blocking valid PPTX or sibling pages.",
  visualCheckpoints: [
    "Pages 1-3 (corrupt PNG, JPEG, GIF backgrounds): preview rejected cleanly (`png: undefined`) with located `/background` errors.",
    "Page 4 (solid-ok): dark emerald canvas (#065F46) with centered white rectangle.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Pages 1-3 retain embedded background image parts while Page 4 retains solid emerald background.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(Slide, {
        id: "corrupt-png-bg",
        background: { image: corruptPngBytes },
      }),
      createElement(Slide, {
        id: "corrupt-jpeg-bg",
        background: { image: corruptJpegBytes },
      }),
      createElement(Slide, {
        id: "corrupt-gif-bg",
        background: { image: corruptGifBytes },
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
        id: "corrupt-png-bg",
        background: { image: corruptPngBytes },
      }),
      slide({
        id: "corrupt-jpeg-bg",
        background: { image: corruptJpegBytes },
      }),
      slide({
        id: "corrupt-gif-bg",
        background: { image: corruptGifBytes },
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
    assert.equal(validation.slides[0]?.background?.type, "image", author);
    assert.equal(validation.slides[1]?.background?.type, "image", author);
    assert.equal(validation.slides[2]?.background?.type, "image", author);
    assert.equal(validation.slides[3]?.background?.type, "solid", author);

    for (const [idx, id] of [
      [0, "corrupt-png-bg"],
      [1, "corrupt-jpeg-bg"],
      [2, "corrupt-gif-bg"],
    ] as const) {
      const page = preview.pages[idx]!;
      assert.equal(page.png, undefined, `${author}:${id}`);
      assert.equal(page.svg, undefined, `${author}:${id}`);
      assert.equal(
        page.diagnostics[0]?.path,
        `deck/slide[${id}]/background`,
        `${author}:${id}`,
      );
      assert.match(page.diagnostics[0]?.message ?? "", /decode/i);
    }
    assert.ok(preview.pages[3]?.png !== undefined, author);
    assert.deepEqual(preview.pages[3]?.diagnostics, [], author);
  },
};
