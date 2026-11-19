import assert from "node:assert/strict";
import { createRequire } from "node:module";

import { deck, image, shape, slide } from "@react-pptx/capabilities/document";
import { Image } from "@react-pptx/capabilities/image-elements";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Deck, Slide } from "@react-pptx/core";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

const require = createRequire(import.meta.url);
const { encode: encodeJpeg } = require("jpeg-js") as typeof import("jpeg-js");

function createThreeZoneBitmap(): PNG {
  // 120x60 bitmap: left third (x < 40) red, top-right blue, bottom-right gold.
  const source = new PNG({ width: 120, height: 60 });
  for (let y = 0; y < source.height; y++) {
    for (let x = 0; x < source.width; x++) {
      source.data.set(
        x < 40
          ? [220, 45, 45, 255]
          : y < 30
            ? [30, 130, 220, 255]
            : [245, 190, 30, 255],
        (y * source.width + x) * 4,
      );
    }
  }
  return source;
}

const sourceBitmap = createThreeZoneBitmap();
const samplePngBytes = new Uint8Array(PNG.sync.write(sourceBitmap));
const sampleJpegBytes = new Uint8Array(encodeJpeg(sourceBitmap, 100).data);

const sampleGifBytes = new Uint8Array(
  Buffer.from(
    "47494638396101000100800000ffffff0000002c00000000010001000002024401003b",
    "hex",
  ),
);

function createCorruptPng(): Uint8Array {
  const bytes = new Uint8Array(samplePngBytes);
  const idat = Buffer.from(bytes).indexOf(Buffer.from("IDAT"));
  bytes[idat + 6] = (bytes[idat + 6] ?? 0) ^ 0xff;
  bytes[idat + 7] = (bytes[idat + 7] ?? 0) ^ 0xff;
  return bytes;
}

function createCorruptJpeg(): Uint8Array {
  const bytes = new Uint8Array(sampleJpegBytes);
  const sos = Buffer.from(bytes).indexOf(Buffer.from([0xff, 0xda]));
  const sosLen = ((bytes[sos + 2] ?? 0) << 8) + (bytes[sos + 3] ?? 0);
  const scanStart = sos + 2 + sosLen;
  for (let i = scanStart; i < bytes.length - 4; i++) {
    bytes[i] = 0x55;
  }
  return bytes;
}

const corruptPngBytes = createCorruptPng();
const corruptJpegBytes = createCorruptJpeg();

export const imagePositiveScenario: CapabilityScenario = {
  id: "image-placement-and-transforms",
  capability: "rp:image",
  description:
    "Dual-author Image scenario covering PNG & JPEG placement (contain, cover, fill, explicit crop), rotate, flipH/flipV, opacity, and media deduplication.",
  visualCheckpoints: [
    "Page 1 (png-placement): light slate canvas (#E8EDF2); left image (contain) is letterboxed vertically into 160x80 centered in its 160x200 slot showing red/blue/gold; middle image (cover) fills 160x200 showing centered vertical band with thin red strip on left and blue/gold on right; right image (explicit crop left: 1/3) hides red completely and displays blue top / gold bottom.",
    "Page 2 (jpeg-transforms): white canvas; top-left stretched JPEG (fill); top-middle 30deg rotated JPEG; bottom-left flipH+flipV JPEG (gold top-left, blue bottom-left, red right); right side 0.6 opacity JPEG overlapping dark slate backdrop (#0F172A).",
    "Page 3 (shared-media-reuse): two instances of the same PNG (full and top-cropped) sharing a single deduplicated media part.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all <p:pic> objects preserve native <a:srcRect>, <a:xfrm rot/flipH/flipV>, <a:alphaModFix>, and deduplicated ppt/media/image1.png + image2.jpeg relationships.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "png-placement", background: "#E8EDF2" },
        createElement(Image, {
          id: "contain",
          src: samplePngBytes,
          fit: "contain",
          x: 30,
          y: 60,
          width: 160,
          height: 200,
        }),
        createElement(Image, {
          id: "cover",
          src: samplePngBytes,
          fit: "cover",
          x: 240,
          y: 60,
          width: 160,
          height: 200,
        }),
        createElement(Image, {
          id: "explicit-crop",
          src: samplePngBytes,
          crop: { left: 1 / 3 },
          x: 450,
          y: 60,
          width: 160,
          height: 200,
        }),
      ),
      createElement(
        Slide,
        { id: "jpeg-transforms", background: "#FFFFFF" },
        createElement(Image, {
          id: "jpeg-fill",
          src: sampleJpegBytes,
          fit: "fill",
          x: 30,
          y: 40,
          width: 150,
          height: 90,
        }),
        createElement(Image, {
          id: "jpeg-rotated",
          src: sampleJpegBytes,
          x: 220,
          y: 55,
          width: 120,
          height: 60,
          style: { rotate: 30 },
        }),
        createElement(Image, {
          id: "jpeg-flipped",
          src: sampleJpegBytes,
          x: 30,
          y: 190,
          width: 140,
          height: 70,
          flipH: true,
          flipV: true,
        }),
        createElement(Shape, {
          id: "backdrop",
          x: 390,
          y: 110,
          width: 180,
          height: 140,
          fill: "#0F172A",
          style: { zIndex: 1 },
        }),
        createElement(Image, {
          id: "jpeg-opacity",
          src: sampleJpegBytes,
          x: 410,
          y: 125,
          width: 140,
          height: 110,
          style: { opacity: 0.6, zIndex: 2 },
        }),
      ),
      createElement(
        Slide,
        { id: "shared-media-reuse", background: "#F8FAFC" },
        createElement(Image, {
          id: "shared-full",
          src: samplePngBytes,
          x: 60,
          y: 110,
          width: 180,
          height: 90,
        }),
        createElement(Image, {
          id: "shared-top-cropped",
          src: samplePngBytes,
          crop: { top: 0.5 },
          x: 340,
          y: 110,
          width: 180,
          height: 90,
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "png-placement", background: "#E8EDF2" }).append(
        image({
          id: "contain",
          src: samplePngBytes,
          fit: "contain",
          x: 30,
          y: 60,
          width: 160,
          height: 200,
        }),
        image({
          id: "cover",
          src: samplePngBytes,
          fit: "cover",
          x: 240,
          y: 60,
          width: 160,
          height: 200,
        }),
        image({
          id: "explicit-crop",
          src: samplePngBytes,
          crop: { left: 1 / 3 },
          x: 450,
          y: 60,
          width: 160,
          height: 200,
        }),
      ),
      slide({ id: "jpeg-transforms", background: "#FFFFFF" }).append(
        image({
          id: "jpeg-fill",
          src: sampleJpegBytes,
          fit: "fill",
          x: 30,
          y: 40,
          width: 150,
          height: 90,
        }),
        image({
          id: "jpeg-rotated",
          src: sampleJpegBytes,
          x: 220,
          y: 55,
          width: 120,
          height: 60,
          style: { rotate: 30 },
        }),
        image({
          id: "jpeg-flipped",
          src: sampleJpegBytes,
          x: 30,
          y: 190,
          width: 140,
          height: 70,
          flipH: true,
          flipV: true,
        }),
        shape({
          id: "backdrop",
          x: 390,
          y: 110,
          width: 180,
          height: 140,
          fill: "#0F172A",
          style: { zIndex: 1 },
        }),
        image({
          id: "jpeg-opacity",
          src: sampleJpegBytes,
          x: 410,
          y: 125,
          width: 140,
          height: 110,
          style: { opacity: 0.6, zIndex: 2 },
        }),
      ),
      slide({ id: "shared-media-reuse", background: "#F8FAFC" }).append(
        image({
          id: "shared-full",
          src: samplePngBytes,
          x: 60,
          y: 110,
          width: 180,
          height: 90,
        }),
        image({
          id: "shared-top-cropped",
          src: samplePngBytes,
          crop: { top: 0.5 },
          x: 340,
          y: 110,
          width: 180,
          height: 90,
        }),
      ),
    );
  },
  assert({ author, pptx, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.equal(validation.slides.length, 3, author);

    // Media deduplication: across all 3 slides and 9 image objects, only 1 PNG + 1 JPEG part exist.
    const zip = unzipSync(pptx);
    const mediaParts = Object.keys(zip)
      .filter((entry) => entry.startsWith("ppt/media/"))
      .sort();
    assert.equal(mediaParts.length, 2, author);
    assert.ok(
      mediaParts.some((name) => name.endsWith(".png")),
      author,
    );
    assert.ok(
      mediaParts.some((name) => name.endsWith(".jpeg")),
      author,
    );

    const p1 = preview.pages[0]!;
    assert.deepEqual(p1.diagnostics, [], author);
    const containNode = p1.nodes.find((n) => n.authorId === "contain")!;
    assert.deepEqual(
      containNode.box,
      { x: 30, y: 120, width: 160, height: 80 },
      author,
    );
    const p1Pixels = PNG.sync.read(Buffer.from(p1.png!));
    // Letterbox top above contain (110, 80) matches slide background #E8EDF2.
    const letterboxOffset = (80 * p1Pixels.width + 110) * 4;
    assert.deepEqual(
      [...p1Pixels.data.subarray(letterboxOffset, letterboxOffset + 4)],
      [0xe8, 0xed, 0xf2, 255],
      author,
    );
    // Explicit crop left: 1/3 removes red: top half (530, 100) is blue, bottom half (530, 210) is gold.
    const cropTopOffset = (100 * p1Pixels.width + 530) * 4;
    assert.ok((p1Pixels.data[cropTopOffset + 2] ?? 0) > 200, author);
    assert.ok((p1Pixels.data[cropTopOffset] ?? 255) < 50, author);
    const cropBottomOffset = (210 * p1Pixels.width + 530) * 4;
    assert.ok((p1Pixels.data[cropBottomOffset] ?? 0) > 220, author);
    assert.ok((p1Pixels.data[cropBottomOffset + 1] ?? 0) > 170, author);
    assert.ok((p1Pixels.data[cropBottomOffset + 2] ?? 255) < 50, author);

    // Native XML preserves crop and placement on Slide 1.
    const slide1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
    assert.match(slide1Xml, /<a:srcRect l="30000" r="30000"\/>/u, author);
    assert.match(slide1Xml, /<a:srcRect l="33333"\/>/u, author);

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    const rotatedNode = p2.nodes.find((n) => n.authorId === "jpeg-rotated")!;
    const rad = (30 * Math.PI) / 180;
    expect(rotatedNode.screenBox?.width).toBeCloseTo(
      120 * Math.cos(rad) + 60 * Math.sin(rad),
      1,
    );
    expect(rotatedNode.screenBox?.height).toBeCloseTo(
      120 * Math.sin(rad) + 60 * Math.cos(rad),
      1,
    );
    const p2Pixels = PNG.sync.read(Buffer.from(p2.png!));
    // Flipped (flipH + flipV) image at (30..170, 190..260): top-left (50, 205) must be gold!
    const flippedTopLeft = (205 * p2Pixels.width + 50) * 4;
    assert.ok((p2Pixels.data[flippedTopLeft] ?? 0) > 220, author);
    assert.ok((p2Pixels.data[flippedTopLeft + 1] ?? 0) > 170, author);
    assert.ok((p2Pixels.data[flippedTopLeft + 2] ?? 255) < 50, author);
    // Opacity 0.6 over dark #0F172A backdrop: blue channel at top-right (520, 145) is darkened (~138 vs 220).
    const blendedOffset = (145 * p2Pixels.width + 520) * 4;
    assert.ok(
      (p2Pixels.data[blendedOffset + 2] ?? 0) > 120 &&
        (p2Pixels.data[blendedOffset + 2] ?? 255) < 180,
      author,
    );

    const p3 = preview.pages[2]!;
    assert.deepEqual(p3.diagnostics, [], author);
    const p3Pixels = PNG.sync.read(Buffer.from(p3.png!));
    // Top-cropped shared-top-cropped at (340..520, 110..200): right side (480, 130) must be gold (not blue).
    const sharedRightTop = (130 * p3Pixels.width + 480) * 4;
    assert.ok((p3Pixels.data[sharedRightTop] ?? 0) > 220, author);
    assert.ok((p3Pixels.data[sharedRightTop + 2] ?? 255) < 50, author);
  },
};

export const imageRejectionScenario: CapabilityScenario = {
  id: "image-preview-rejections",
  capability: "rp:image",
  description:
    "Dual-author negative scenario verifying GIF, non-rect mask, border/shadow styles, and corrupt PNG/JPEG streams emit located errors while PPTX and sibling pages survive.",
  visualCheckpoints: [
    "Page 1 (unsupported-image-preview): no PNG emitted (preview rejected cleanly with 6 located errors).",
    "Page 2 (valid-sibling-image): clean white canvas with centered three-zone PNG image.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 retains native GIF, ellipse-masked PNG, bordered/shadowed PNGs, and embedded picture relationships.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "unsupported-image-preview" },
        createElement(Image, {
          id: "gif-img",
          src: sampleGifBytes,
          x: 20,
          y: 20,
          width: 40,
          height: 40,
        }),
        createElement(Image, {
          id: "shaped-img",
          src: samplePngBytes,
          geometry: "ellipse",
          x: 80,
          y: 20,
          width: 60,
          height: 40,
        }),
        createElement(Image, {
          id: "bordered-img",
          src: samplePngBytes,
          x: 160,
          y: 20,
          width: 60,
          height: 40,
          style: { borderColor: "#1E293B", borderWidth: 3 },
        }),
        createElement(Image, {
          id: "shadowed-img",
          src: samplePngBytes,
          x: 240,
          y: 20,
          width: 60,
          height: 40,
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        createElement(Image, {
          id: "corrupt-png-img",
          src: corruptPngBytes,
          x: 20,
          y: 90,
          width: 60,
          height: 40,
        }),
        createElement(Image, {
          id: "corrupt-jpeg-img",
          src: corruptJpegBytes,
          x: 100,
          y: 90,
          width: 60,
          height: 40,
        }),
      ),
      createElement(
        Slide,
        { id: "valid-sibling-image", background: "#FFFFFF" },
        createElement(Image, {
          id: "valid-img",
          src: samplePngBytes,
          x: 150,
          y: 90,
          width: 180,
          height: 90,
        }),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "unsupported-image-preview" }).append(
        image({
          id: "gif-img",
          src: sampleGifBytes,
          x: 20,
          y: 20,
          width: 40,
          height: 40,
        }),
        image({
          id: "shaped-img",
          src: samplePngBytes,
          geometry: "ellipse",
          x: 80,
          y: 20,
          width: 60,
          height: 40,
        }),
        image({
          id: "bordered-img",
          src: samplePngBytes,
          x: 160,
          y: 20,
          width: 60,
          height: 40,
          style: { borderColor: "#1E293B", borderWidth: 3 },
        }),
        image({
          id: "shadowed-img",
          src: samplePngBytes,
          x: 240,
          y: 20,
          width: 60,
          height: 40,
          style: { boxShadow: { blur: 4, distance: 2 } },
        }),
        image({
          id: "corrupt-png-img",
          src: corruptPngBytes,
          x: 20,
          y: 90,
          width: 60,
          height: 40,
        }),
        image({
          id: "corrupt-jpeg-img",
          src: corruptJpegBytes,
          x: 100,
          y: 90,
          width: 60,
          height: 40,
        }),
      ),
      slide({ id: "valid-sibling-image", background: "#FFFFFF" }).append(
        image({
          id: "valid-img",
          src: samplePngBytes,
          x: 150,
          y: 90,
          width: 180,
          height: 90,
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(validation.slides[0]?.objects.length, 6, author);

    const p1 = preview.pages[0]!;
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.diagnostics.length, 6, author);
    for (const id of [
      "gif-img",
      "shaped-img",
      "bordered-img",
      "shadowed-img",
      "corrupt-png-img",
      "corrupt-jpeg-img",
    ]) {
      assert.ok(
        p1.diagnostics.some(
          (d) =>
            d.severity === "error" &&
            d.page === 1 &&
            d.path.includes(`image[${id}]`),
        ),
        `${author}: missing diagnostic for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    assert.ok(p2.png !== undefined, author);
    const p2Pixels = PNG.sync.read(Buffer.from(p2.png));
    const redLeft = (135 * p2Pixels.width + 180) * 4;
    assert.ok((p2Pixels.data[redLeft] ?? 0) > 200, author);
    assert.ok((p2Pixels.data[redLeft + 2] ?? 255) < 60, author);
  },
};
