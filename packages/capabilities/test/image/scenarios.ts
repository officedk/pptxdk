import assert from "node:assert/strict";
import { createRequire } from "node:module";

import { deck, image, shape, slide } from "@react-pptx/capabilities/document";
import { Image } from "@react-pptx/capabilities/image-elements";
import { path, preset } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Deck, Slide } from "@react-pptx/react";
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
const corruptGifBytes = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 2, 0, 0x80, 0, 0, 0xff, 0xff, 0xff,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 2, 0, 2, 0, 0x00, 0x02, 0x02, 0xbc, 0x00,
  0x00, 0x3b,
]);

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
        createElement(Image, {
          id: "gif-sample",
          src: sampleGifBytes,
          x: 260,
          y: 230,
          width: 60,
          height: 60,
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
        image({
          id: "gif-sample",
          src: sampleGifBytes,
          x: 260,
          y: 230,
          width: 60,
          height: 60,
        }),
      ),
    );
  },
  assert({ author, pptx, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.equal(validation.slides.length, 3, author);

    // Media deduplication: across all 3 slides and 10 image objects, only 1 PNG + 1 JPEG + 1 GIF part exist.
    const zip = unzipSync(pptx);
    const mediaParts = Object.keys(zip)
      .filter((entry) => entry.startsWith("ppt/media/"))
      .sort();
    assert.equal(mediaParts.length, 3, author);
    assert.ok(
      mediaParts.some((name) => name.endsWith(".png")),
      author,
    );
    assert.ok(
      mediaParts.some((name) => name.endsWith(".jpeg")),
      author,
    );
    assert.ok(
      mediaParts.some((name) => name.endsWith(".gif")),
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

const triangleMaskGeometry = path("M0 0 L10 0 L5 10 Z", {
  viewBox: [10, 10],
});

export const imageMasksAndEffectsScenario: CapabilityScenario = {
  id: "image-masks-borders-and-effects",
  capability: "rp:image",
  description:
    "Dual-author Image scenario covering preset & custom geometry masks, solid/dashed/dotted borders, boxShadow, softEdge, opacity+border, and flip+rotate composition across PNG, JPEG, and GIF.",
  visualCheckpoints: [
    "Page 1 (geometry-masks-and-borders): light #F8FAFC canvas; ellipse PNG with dashed border, adjusted roundRect JPEG with dotted border, star5 GIF with solid border, smileyFace PNG (fill-none mouth arc excluded from clip), and custom triangle path PNG.",
    "Page 2 (effects-and-transforms): white canvas; shadowed & flipped+rotated rightArrow PNG, feathered softEdge diamond JPEG, and semi-transparent bordered PNG over dark backdrop.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify native <p:pic><p:spPr> preserves <a:prstGeom>/<a:custGeom>, <a:prstDash>, <a:outerShdw rotWithShape='0'>, <a:softEdge>, <a:alphaModFix>, and <a:xfrm flipH/rot>.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "geometry-masks-and-borders", background: "#F8FAFC" },
        createElement(Image, {
          id: "ellipse-dashed",
          src: samplePngBytes,
          geometry: "ellipse",
          x: 24,
          y: 30,
          width: 130,
          height: 90,
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "dashed",
          },
        }),
        createElement(Image, {
          id: "roundrect-dotted",
          src: sampleJpegBytes,
          geometry: preset("roundRect", { adj: 25000 }),
          x: 180,
          y: 30,
          width: 130,
          height: 90,
          style: {
            borderColor: "#DC2626",
            borderWidth: 2,
            borderStyle: "dotted",
          },
        }),
        createElement(Image, {
          id: "star-gif",
          src: sampleGifBytes,
          geometry: preset("star5", { adj: 22000 }),
          x: 340,
          y: 30,
          width: 110,
          height: 110,
          style: {
            borderColor: "#0F766E",
            borderWidth: 2,
            borderStyle: "solid",
          },
        }),
        createElement(Image, {
          id: "smiley-png",
          src: samplePngBytes,
          geometry: "smileyFace",
          x: 480,
          y: 30,
          width: 110,
          height: 110,
          style: {
            borderColor: "#0F172A",
            borderWidth: 2,
          },
        }),
        createElement(Image, {
          id: "custom-triangle",
          src: samplePngBytes,
          geometry: triangleMaskGeometry,
          x: 24,
          y: 170,
          width: 140,
          height: 120,
          style: {
            borderColor: "#1D4ED8",
            borderWidth: 2,
          },
        }),
      ),
      createElement(
        Slide,
        { id: "effects-and-transforms", background: "#FFFFFF" },
        createElement(Image, {
          id: "arrow-shadow-flip-rot",
          src: samplePngBytes,
          geometry: "rightArrow",
          x: 40,
          y: 50,
          width: 160,
          height: 100,
          flipH: true,
          style: {
            rotate: 20,
            borderColor: "#0F172A",
            borderWidth: 2,
            boxShadow: { color: "#000000", blur: 6, distance: 6, angle: 45 },
          },
        }),
        createElement(Image, {
          id: "soft-diamond",
          src: sampleJpegBytes,
          geometry: "diamond",
          x: 240,
          y: 50,
          width: 140,
          height: 110,
          style: { softEdge: 10 },
        }),
        createElement(Shape, {
          id: "opacity-backdrop",
          x: 410,
          y: 45,
          width: 190,
          height: 140,
          fill: "#0F172A",
          style: { zIndex: 1 },
        }),
        createElement(Image, {
          id: "opacity-bordered",
          src: samplePngBytes,
          geometry: preset("roundRect", { adj: 15000 }),
          x: 430,
          y: 65,
          width: 150,
          height: 100,
          style: {
            opacity: 0.5,
            borderColor: "#FFFFFF",
            borderWidth: 4,
            zIndex: 2,
          },
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "geometry-masks-and-borders", background: "#F8FAFC" }).append(
        image({
          id: "ellipse-dashed",
          src: samplePngBytes,
          geometry: "ellipse",
          x: 24,
          y: 30,
          width: 130,
          height: 90,
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "dashed",
          },
        }),
        image({
          id: "roundrect-dotted",
          src: sampleJpegBytes,
          geometry: preset("roundRect", { adj: 25000 }),
          x: 180,
          y: 30,
          width: 130,
          height: 90,
          style: {
            borderColor: "#DC2626",
            borderWidth: 2,
            borderStyle: "dotted",
          },
        }),
        image({
          id: "star-gif",
          src: sampleGifBytes,
          geometry: preset("star5", { adj: 22000 }),
          x: 340,
          y: 30,
          width: 110,
          height: 110,
          style: {
            borderColor: "#0F766E",
            borderWidth: 2,
            borderStyle: "solid",
          },
        }),
        image({
          id: "smiley-png",
          src: samplePngBytes,
          geometry: "smileyFace",
          x: 480,
          y: 30,
          width: 110,
          height: 110,
          style: {
            borderColor: "#0F172A",
            borderWidth: 2,
          },
        }),
        image({
          id: "custom-triangle",
          src: samplePngBytes,
          geometry: triangleMaskGeometry,
          x: 24,
          y: 170,
          width: 140,
          height: 120,
          style: {
            borderColor: "#1D4ED8",
            borderWidth: 2,
          },
        }),
      ),
      slide({ id: "effects-and-transforms", background: "#FFFFFF" }).append(
        image({
          id: "arrow-shadow-flip-rot",
          src: samplePngBytes,
          geometry: "rightArrow",
          x: 40,
          y: 50,
          width: 160,
          height: 100,
          flipH: true,
          style: {
            rotate: 20,
            borderColor: "#0F172A",
            borderWidth: 2,
            boxShadow: { color: "#000000", blur: 6, distance: 6, angle: 45 },
          },
        }),
        image({
          id: "soft-diamond",
          src: sampleJpegBytes,
          geometry: "diamond",
          x: 240,
          y: 50,
          width: 140,
          height: 110,
          style: { softEdge: 10 },
        }),
        shape({
          id: "opacity-backdrop",
          x: 410,
          y: 45,
          width: 190,
          height: 140,
          fill: "#0F172A",
          style: { zIndex: 1 },
        }),
        image({
          id: "opacity-bordered",
          src: samplePngBytes,
          geometry: preset("roundRect", { adj: 15000 }),
          x: 430,
          y: 65,
          width: 150,
          height: 100,
          style: {
            opacity: 0.5,
            borderColor: "#FFFFFF",
            borderWidth: 4,
            zIndex: 2,
          },
        }),
      ),
    );
  },
  assert({ author, pptx, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.equal(validation.slides.length, 2, author);

    const p1 = preview.pages[0]!;
    assert.deepEqual(p1.diagnostics, [], author);
    expect(p1.svg).toContain("<clipPath");
    expect(p1.svg).toContain("stroke-dasharray=");
    const p1Pixels = PNG.sync.read(Buffer.from(p1.png!));
    // Top-left corner (26, 32) outside ellipse mask stays #F8FAFC background.
    const outsideEllipse = (32 * p1Pixels.width + 26) * 4;
    assert.deepEqual(
      [...p1Pixels.data.subarray(outsideEllipse, outsideEllipse + 4)],
      [0xf8, 0xfa, 0xfc, 255],
      author,
    );
    // Bottom-left corner (26, 286) outside custom triangle stays #F8FAFC background.
    const outsideTriangle = (286 * p1Pixels.width + 26) * 4;
    assert.deepEqual(
      [...p1Pixels.data.subarray(outsideTriangle, outsideTriangle + 4)],
      [0xf8, 0xfa, 0xfc, 255],
      author,
    );
    // Top-center (94, 190) inside custom triangle is red/blue image pixel, not background.
    const insideTriangle = (190 * p1Pixels.width + 94) * 4;
    assert.ok(
      (p1Pixels.data[insideTriangle] ?? 0) > 180 ||
        (p1Pixels.data[insideTriangle + 2] ?? 0) > 180,
      author,
    );

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    expect(p2.svg).toContain("<feDropShadow");
    expect(p2.svg).toContain("<mask");
    // Border stroke-opacity must be 0.5 (not double-multiplied to 0.25).
    expect(p2.svg).toContain('stroke-opacity="0.5"');
    expect(p2.svg).not.toContain('stroke-opacity="0.25"');

    const zip = unzipSync(pptx);
    const slide1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
    assert.match(slide1Xml, /<a:prstGeom prst="ellipse"/u, author);
    assert.match(slide1Xml, /<a:prstDash val="dash"\/>/u, author);
    assert.match(slide1Xml, /<a:prstDash val="sysDot"\/>/u, author);
    assert.match(slide1Xml, /<a:custGeom>/u, author);
    const slide2Xml = strFromU8(zip["ppt/slides/slide2.xml"]!);
    assert.match(slide2Xml, /<a:outerShdw[^>]*rotWithShape="0"/u, author);
    assert.match(slide2Xml, /<a:softEdge /u, author);
  },
};

export const imageRejectionScenario: CapabilityScenario = {
  id: "image-preview-rejections",
  capability: "rp:image",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying only corrupt PNG/JPEG/GIF pixel streams emit located preview-corrupt-image errors while PPTX and shaped/bordered/shadowed sibling pages survive.",
  visualCheckpoints: [
    "Page 1 (corrupt-image-preview): no PNG emitted (preview rejected cleanly with 3 located preview-corrupt-image errors).",
    "Page 2 (valid-sibling-image): clean white canvas with ellipse-masked, bordered, shadowed three-zone PNG image.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify Page 1 retains native media parts and Page 2 retains ellipse mask, border, and shadow.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "corrupt-image-preview" },
        createElement(Image, {
          id: "corrupt-gif-img",
          src: corruptGifBytes,
          x: 20,
          y: 20,
          width: 40,
          height: 40,
        }),
        createElement(Image, {
          id: "corrupt-png-img",
          src: corruptPngBytes,
          x: 80,
          y: 20,
          width: 60,
          height: 40,
        }),
        createElement(Image, {
          id: "corrupt-jpeg-img",
          src: corruptJpegBytes,
          x: 160,
          y: 20,
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
          geometry: "ellipse",
          x: 150,
          y: 90,
          width: 180,
          height: 90,
          style: {
            borderColor: "#1E293B",
            borderWidth: 2,
            boxShadow: { blur: 4, distance: 3 },
          },
        }),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "corrupt-image-preview" }).append(
        image({
          id: "corrupt-gif-img",
          src: corruptGifBytes,
          x: 20,
          y: 20,
          width: 40,
          height: 40,
        }),
        image({
          id: "corrupt-png-img",
          src: corruptPngBytes,
          x: 80,
          y: 20,
          width: 60,
          height: 40,
        }),
        image({
          id: "corrupt-jpeg-img",
          src: corruptJpegBytes,
          x: 160,
          y: 20,
          width: 60,
          height: 40,
        }),
      ),
      slide({ id: "valid-sibling-image", background: "#FFFFFF" }).append(
        image({
          id: "valid-img",
          src: samplePngBytes,
          geometry: "ellipse",
          x: 150,
          y: 90,
          width: 180,
          height: 90,
          style: {
            borderColor: "#1E293B",
            borderWidth: 2,
            boxShadow: { blur: 4, distance: 3 },
          },
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(validation.slides[0]?.objects.length, 3, author);

    const p1 = preview.pages[0]!;
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.diagnostics.length, 3, author);
    for (const id of [
      "corrupt-gif-img",
      "corrupt-png-img",
      "corrupt-jpeg-img",
    ]) {
      assert.ok(
        p1.diagnostics.some(
          (d) =>
            d.severity === "error" &&
            d.page === 1 &&
            d.path.includes(`image[${id}]`) &&
            /decode/i.test(d.message),
        ),
        `${author}: missing corrupt-image diagnostic for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    assert.ok(p2.png !== undefined, author);
    const p2Pixels = PNG.sync.read(Buffer.from(p2.png));
    const redLeft = (135 * p2Pixels.width + 185) * 4;
    assert.ok((p2Pixels.data[redLeft] ?? 0) > 200, author);
    assert.ok((p2Pixels.data[redLeft + 2] ?? 255) < 60, author);
  },
};
