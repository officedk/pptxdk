import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

import { arrowTypes } from "@pptxdk/capabilities";
import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  font,
  shape,
  slide,
  text,
} from "@pptxdk/capabilities/document";
import {
  linearGradient,
  path,
  radialGradient,
} from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";

const require = createRequire(import.meta.url);
const { encode: encodeJpeg } = require("jpeg-js") as typeof import("jpeg-js");

function createBoxFillPng(): Uint8Array {
  const img = new PNG({ width: 20, height: 10 });
  for (let y = 0; y < 10; y += 1) {
    for (let x = 0; x < 20; x += 1) {
      img.data.set(
        x < 10 ? [0x0ea, 0x58, 0x0c, 255] : [0x02, 0x84, 0xc7, 255],
        (y * 20 + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(img));
}

function createCorruptJpegFrom(pngSource: PNG): Uint8Array {
  const jpeg = new Uint8Array(encodeJpeg(pngSource, 90).data);
  const sos = Buffer.from(jpeg).indexOf(Buffer.from([0xff, 0xda]));
  const sosLen = ((jpeg[sos + 2] ?? 0) << 8) + (jpeg[sos + 3] ?? 0);
  for (let i = sos + 2 + sosLen; i < jpeg.length - 4; i += 1) {
    jpeg[i] = 0x55;
  }
  return jpeg;
}

const twoToneSource = new PNG({ width: 20, height: 10 });
for (let y = 0; y < 10; y += 1) {
  for (let x = 0; x < 20; x += 1) {
    twoToneSource.data.set(
      x < 10 ? [255, 0, 0, 255] : [0, 0, 255, 255],
      (y * 20 + x) * 4,
    );
  }
}

const sampleFillPng = createBoxFillPng();
const sampleFillGif = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0x80, 0, 0, 0x10, 0xb9, 0x81,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0x00, 0x02, 0x02, 0x44, 0x01,
  0x00, 0x3b,
]);

const corruptPngBytes = (() => {
  const bytes = new Uint8Array(PNG.sync.write(twoToneSource));
  const idat = Buffer.from(bytes).indexOf(Buffer.from("IDAT"));
  bytes[idat + 6] = (bytes[idat + 6] ?? 0) ^ 0xff;
  bytes[idat + 7] = (bytes[idat + 7] ?? 0) ^ 0xff;
  return bytes;
})();
const corruptJpegBytes = createCorruptJpegFrom(twoToneSource);
const corruptGifBytes = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 2, 0, 0x80, 0, 0, 0xff, 0xff, 0xff,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 2, 0, 2, 0, 0x00, 0x02, 0x02, 0xbc, 0x00,
  0x00, 0x3b,
]);

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
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

export const boxPositiveScenario: CapabilityScenario = {
  id: "box-containers-and-nesting",
  capability: "box",
  description:
    "Dual-author Box scenario covering transparent Flex/Grid layout containers, visible filled/bordered/rotated/linked Boxes, nested Shape & wrapped Text, and native theme / slide themeOverride.",
  visualCheckpoints: [
    "Page 1 (box-nesting-and-styles): light slate #F8FAFC canvas; transparent outer Flex container holds two side-by-side visible cards.",
    "Page 2 (box-theme-and-wrapping): white canvas; green-tinted Grid card resolved from native scheme slots and container style; hollow bordered Box below.",
    "Page 3 (box-visual-effects): linear-gradient card (via Slide themeOverride) with borderRadius & dashed border, radial-gradient card, PNG & GIF image-filled rounded Boxes, rotated screen-space boxShadow card, and feathered softEdge card.",
    "Page 4 (box-arrow-types): Box grid traversing all 6 ArrowType values across straight lines, curved open paths, dashed strokes, and flipped/rotated segments.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify transparent Boxes emit zero shapes while visible Boxes emit native <p:sp> shapes with card-accent carrying rot and <a:hlinkClick> to box-theme-and-wrapping.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        theme: {
          colors: {
            accent1: "#0F766E",
            lt2: "#ECFDF5",
            accent2: "#991B1B",
          },
          fonts: { major: jostFont, minor: notoFont },
        },
      },
      createElement(
        Slide,
        { id: "box-nesting-and-styles", background: "#F8FAFC" },
        createElement(
          Box,
          {
            id: "shell",
            style: {
              position: "absolute",
              left: 28,
              top: 28,
              width: 584,
              display: "flex",
              flexDirection: "column",
              gap: 16,
              padding: 12,
            },
          },
          createElement(
            Box,
            {
              id: "cards-row",
              style: {
                display: "flex",
                flexDirection: "row",
                gap: 24,
                alignItems: "flex-start",
              },
            },
            createElement(
              Box,
              {
                id: "card-primary",
                style: {
                  width: 260,
                  padding: 14,
                  display: "flex",
                  flexDirection: "column",
                  gap: 10,
                  background: "#EFF6FF",
                  border: { color: "#2563EB", width: 2, style: "solid" },
                },
              },
              createElement(Shape, {
                id: "primary-badge",
                style: { width: 56, height: 28, background: "#1D4ED8" },
              }),
              createElement(
                Text,
                {
                  id: "primary-copy",
                  style: {
                    fontFamily: notoFont,
                    fontSize: 16,
                    lineHeight: 24,
                    color: "#0F172A",
                  },
                },
                "容器布局 Shared Box Layout with automatic wrapping",
              ),
            ),
            createElement(
              Box,
              {
                id: "card-accent",
                link: { id: "box-theme-and-wrapping" },
                style: {
                  width: 260,
                  padding: 14,
                  display: "flex",
                  flexDirection: "column",
                  gap: 10,
                  background: "#FEF3C7",
                  border: { color: "#D97706", width: 2, style: "solid" },
                  opacity: 0.9,
                  rotate: 3,
                },
              },
              createElement(Shape, {
                id: "accent-badge",
                style: { width: 56, height: 28, background: "#0F766E" },
              }),
              createElement(
                Text,
                {
                  id: "accent-copy",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 16,
                    lineHeight: 24,
                    color: "#78350F",
                  },
                },
                "Click to inspect theme scope & grid layout inside nested Box",
              ),
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "box-theme-and-wrapping", background: "#FFFFFF" },
        createElement(
          Box,
          {
            id: "themed-section",
            style: {
              position: "absolute",
              left: 32,
              top: 28,
              width: 576,
              display: "flex",
              flexDirection: "column",
              gap: 18,
              fontFamily: notoFont,
              fontSize: 15,
              color: "accent1",
            },
          },
          createElement(
            Box,
            {
              id: "grid-panel",
              style: {
                display: "grid",
                gridTemplateColumns: "120px 1fr",
                gap: 16,
                padding: 16,
                background: "lt2",
                border: { color: "#10B981", width: 2, style: "solid" },
              },
            },
            createElement(Shape, {
              id: "grid-swatch",
              style: { width: 120, height: 72, background: "accent1" },
            }),
            createElement(
              Text,
              {
                id: "grid-text",
                style: {
                  fontFamily: "major",
                  fontSize: 17,
                  lineHeight: 25,
                },
              },
              "Container style and native scheme slots drive lt2 fill, accent1 color, and major font inside nested Grid Box.",
            ),
          ),
          createElement(
            Box,
            {
              id: "hollow-frame",
              style: {
                width: 320,
                padding: 12,
                border: { color: "#64748B", width: 2 },
              },
            },
            createElement(
              Box,
              {
                id: "hollow-inner",
                style: { display: "flex", flexDirection: "row", gap: 12 },
              },
              createElement(Shape, {
                id: "coral-chip",
                style: { width: 90, height: 36, background: "#F97316" },
              }),
              createElement(
                Text,
                {
                  id: "hollow-caption",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 15,
                    color: "#334155",
                  },
                },
                "Border-only outer Box",
              ),
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        {
          id: "box-visual-effects",
          background: "#FFFFFF",
          themeOverride: { colors: { lt2: "#1D4ED8", accent1: "#9333EA" } },
        },
        createElement(
          Box,
          {
            id: "effects-grid",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 592,
              display: "grid",
              gridTemplateColumns: "180px 180px 180px",
              gap: 24,
            },
          },
          createElement(Box, {
            id: "linear-rounded-dashed",
            style: {
              height: 90,
              background: linearGradient(90, "lt2", "accent1"),
              border: { color: "#0F172A", width: 2, style: "dashed" },
              borderRadius: 14,
            },
          }),
          createElement(Box, {
            id: "radial-dotted",
            style: {
              height: 90,
              background: radialGradient("#FEF3C7", "#1E3A8A"),
              border: { color: "#F59E0B", width: 2, style: "dotted" },
              borderRadius: 10,
            },
          }),
          createElement(Box, {
            id: "image-png-fill",
            style: {
              height: 90,
              background: { image: sampleFillPng },
              border: { color: "#0F172A", width: 2 },
              borderRadius: 12,
              opacity: 0.9,
            },
          }),
          createElement(Box, {
            id: "image-gif-fill",
            style: {
              height: 90,
              background: { image: sampleFillGif },
              borderRadius: 8,
            },
          }),
          createElement(Box, {
            id: "rotated-shadow",
            style: {
              height: 90,
              background: "#E2E8F0",
              border: { color: "#334155", width: 1.5 },
              borderRadius: 8,
              rotate: 15,
              boxShadow: {
                color: { color: "#0F172A", alpha: 0.45 },
                blur: 6,
                distance: 8,
                angle: 90,
              },
            },
          }),
          createElement(Box, {
            id: "soft-edge-box",
            style: {
              height: 90,
              background: "#0EA5E9",
              softEdge: 10,
            },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "box-arrow-types", background: "#F8FAFC" },
        createElement(
          Box,
          {
            id: "arrow-gallery",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 592,
              display: "flex",
              flexDirection: "column",
              gap: 16,
              padding: 12,
              background: "#FFFFFF",
              border: { color: "#CBD5E1", width: 1 },
              borderRadius: 10,
            },
          },
          ...arrowTypes.map((arrowType, idx) =>
            createElement(
              Box,
              {
                key: arrowType,
                id: `arrow-row-${arrowType}`,
                style: {
                  display: "flex",
                  flexDirection: "row",
                  gap: 24,
                  alignItems: "center",
                },
              },
              createElement(Shape, {
                id: `line-${arrowType}`,
                geometry: "line",
                style: {
                  width: 240,
                  height: 2,
                  flipH: idx % 2 === 1,
                  border: {
                    color: "#0F172A",
                    width: 3,
                    style: arrowType === "arrow" ? "dashed" : "solid",
                    startArrow: arrowType,
                    endArrow: arrowType,
                  },
                },
              }),
              createElement(Shape, {
                id: `curve-${arrowType}`,
                geometry: path("M 0 18 Q 90 -8 180 18", {
                  viewBox: [180, 22],
                  fill: false,
                }),
                style: {
                  width: 180,
                  height: 22,
                  border: {
                    color: "#2563EB",
                    width: 2.5,
                    startArrow: arrowType,
                    endArrow: arrowType,
                  },
                },
              }),
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
      theme: {
        colors: {
          accent1: "#0F766E",
          lt2: "#ECFDF5",
          accent2: "#991B1B",
        },
        fonts: { major: jostFont, minor: notoFont },
      },
    }).append(
      slide({ id: "box-nesting-and-styles", background: "#F8FAFC" }).append(
        box({
          id: "shell",
          style: {
            position: "absolute",
            left: 28,
            top: 28,
            width: 584,
            display: "flex",
            flexDirection: "column",
            gap: 16,
            padding: 12,
          },
        }).append(
          box({
            id: "cards-row",
            style: {
              display: "flex",
              flexDirection: "row",
              gap: 24,
              alignItems: "flex-start",
            },
          }).append(
            box({
              id: "card-primary",
              style: {
                width: 260,
                padding: 14,
                display: "flex",
                flexDirection: "column",
                gap: 10,
                background: "#EFF6FF",
                border: { color: "#2563EB", width: 2, style: "solid" },
              },
            }).append(
              shape({
                id: "primary-badge",
                style: { width: 56, height: 28, background: "#1D4ED8" },
              }),
              text("容器布局 Shared Box Layout with automatic wrapping", {
                id: "primary-copy",
                style: {
                  fontFamily: notoFont,
                  fontSize: 16,
                  lineHeight: 24,
                  color: "#0F172A",
                },
              }),
            ),
            box({
              id: "card-accent",
              link: { id: "box-theme-and-wrapping" },
              style: {
                width: 260,
                padding: 14,
                display: "flex",
                flexDirection: "column",
                gap: 10,
                background: "#FEF3C7",
                border: { color: "#D97706", width: 2, style: "solid" },
                opacity: 0.9,
                rotate: 3,
              },
            }).append(
              shape({
                id: "accent-badge",
                style: { width: 56, height: 28, background: "#0F766E" },
              }),
              text(
                "Click to inspect theme scope & grid layout inside nested Box",
                {
                  id: "accent-copy",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 16,
                    lineHeight: 24,
                    color: "#78350F",
                  },
                },
              ),
            ),
          ),
        ),
      ),
      slide({ id: "box-theme-and-wrapping", background: "#FFFFFF" }).append(
        box({
          id: "themed-section",
          style: {
            position: "absolute",
            left: 32,
            top: 28,
            width: 576,
            display: "flex",
            flexDirection: "column",
            gap: 18,
            fontFamily: notoFont,
            fontSize: 15,
            color: "accent1",
          },
        }).append(
          box({
            id: "grid-panel",
            style: {
              display: "grid",
              gridTemplateColumns: "120px 1fr",
              gap: 16,
              padding: 16,
              background: "lt2",
              border: { color: "#10B981", width: 2, style: "solid" },
            },
          }).append(
            shape({
              id: "grid-swatch",
              style: { width: 120, height: 72, background: "accent1" },
            }),
            text(
              "Container style and native scheme slots drive lt2 fill, accent1 color, and major font inside nested Grid Box.",
              {
                id: "grid-text",
                style: {
                  fontFamily: "major",
                  fontSize: 17,
                  lineHeight: 25,
                },
              },
            ),
          ),
          box({
            id: "hollow-frame",
            style: {
              width: 320,
              padding: 12,
              border: { color: "#64748B", width: 2 },
            },
          }).append(
            box({
              id: "hollow-inner",
              style: { display: "flex", flexDirection: "row", gap: 12 },
            }).append(
              shape({
                id: "coral-chip",
                style: { width: 90, height: 36, background: "#F97316" },
              }),
              text("Border-only outer Box", {
                id: "hollow-caption",
                style: {
                  fontFamily: jostFont,
                  fontSize: 15,
                  color: "#334155",
                },
              }),
            ),
          ),
        ),
      ),
      slide({
        id: "box-visual-effects",
        background: "#FFFFFF",
        themeOverride: { colors: { lt2: "#1D4ED8", accent1: "#9333EA" } },
      }).append(
        box({
          id: "effects-grid",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 592,
            display: "grid",
            gridTemplateColumns: "180px 180px 180px",
            gap: 24,
          },
        }).append(
          box({
            id: "linear-rounded-dashed",
            style: {
              height: 90,
              background: linearGradient(90, "lt2", "accent1"),
              border: { color: "#0F172A", width: 2, style: "dashed" },
              borderRadius: 14,
            },
          }),
          box({
            id: "radial-dotted",
            style: {
              height: 90,
              background: radialGradient("#FEF3C7", "#1E3A8A"),
              border: { color: "#F59E0B", width: 2, style: "dotted" },
              borderRadius: 10,
            },
          }),
          box({
            id: "image-png-fill",
            style: {
              height: 90,
              background: { image: sampleFillPng },
              border: { color: "#0F172A", width: 2 },
              borderRadius: 12,
              opacity: 0.9,
            },
          }),
          box({
            id: "image-gif-fill",
            style: {
              height: 90,
              background: { image: sampleFillGif },
              borderRadius: 8,
            },
          }),
          box({
            id: "rotated-shadow",
            style: {
              height: 90,
              background: "#E2E8F0",
              border: { color: "#334155", width: 1.5 },
              borderRadius: 8,
              rotate: 15,
              boxShadow: {
                color: { color: "#0F172A", alpha: 0.45 },
                blur: 6,
                distance: 8,
                angle: 90,
              },
            },
          }),
          box({
            id: "soft-edge-box",
            style: {
              height: 90,
              background: "#0EA5E9",
              softEdge: 10,
            },
          }),
        ),
      ),
      slide({ id: "box-arrow-types", background: "#F8FAFC" }).append(
        box({
          id: "arrow-gallery",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 592,
            display: "flex",
            flexDirection: "column",
            gap: 16,
            padding: 12,
            background: "#FFFFFF",
            border: { color: "#CBD5E1", width: 1 },
            borderRadius: 10,
          },
        }).append(
          ...arrowTypes.map((arrowType, idx) =>
            box({
              id: `arrow-row-${arrowType}`,
              style: {
                display: "flex",
                flexDirection: "row",
                gap: 24,
                alignItems: "center",
              },
            }).append(
              shape({
                id: `line-${arrowType}`,
                geometry: "line",
                style: {
                  width: 240,
                  height: 2,
                  flipH: idx % 2 === 1,
                  border: {
                    color: "#0F172A",
                    width: 3,
                    style: arrowType === "arrow" ? "dashed" : "solid",
                    startArrow: arrowType,
                    endArrow: arrowType,
                  },
                },
              }),
              shape({
                id: `curve-${arrowType}`,
                geometry: path("M 0 18 Q 90 -8 180 18", {
                  viewBox: [180, 22],
                  fill: false,
                }),
                style: {
                  width: 180,
                  height: 22,
                  border: {
                    color: "#2563EB",
                    width: 2.5,
                    startArrow: arrowType,
                    endArrow: arrowType,
                  },
                },
              }),
            ),
          ),
        ),
      ),
    );
  },
  assert({ resolution, pptx, preview, validation }) {
    assert.equal(preview.pages.length, 4);
    const page1 = preview.pages[0]!;
    const page2 = preview.pages[1]!;
    const page3 = preview.pages[2]!;
    const page4 = preview.pages[3]!;
    for (const page of [page1, page2, page3, page4]) {
      assert.deepEqual(page.errors, []);
      assert.ok(page.png !== undefined);
    }

    // Layout region vs native object identity distinction on Page 1.
    assert.equal(
      page1.nodes.some((n) => n.authorId === "shell"),
      false,
    );
    assert.equal(
      page1.nodes.some((n) => n.authorId === "cards-row"),
      false,
    );
    assert.equal(page1.nodes.length, 6);
    const shellObserved = resolution
      .slide("box-nesting-and-styles")
      .get("shell");
    assert.ok(shellObserved.box !== undefined);
    assert.equal(shellObserved.objects.length, 6);
    const accentCard = page1.nodes.find((n) => n.authorId === "card-accent")!;
    assert.equal(accentCard.link, "slide:box-theme-and-wrapping");
    assert.ok((accentCard.bounds?.height ?? 0) > (accentCard.box?.height ?? 0));

    // PPTX independent verification: link and rotation written on visible Box only.
    assert.equal(validation.slides[0]?.objects.length, 6);
    const pptxLinkedBox = validation.slides[0]?.objects.find(
      (o) => o.link !== undefined,
    );
    assert.deepEqual(pptxLinkedBox?.link, {
      slideId: "box-theme-and-wrapping",
    });
    const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
    assert.match(slide1Xml, /rot="180000"/u);

    // Page 2: Scheme colors and container styles resolved in both SVG and PPTX, hollow frame has fill="none".
    expect(page2.svg).toContain('fill="#ECFDF5"');
    expect(page2.svg).toContain('fill="#0F766E"');
    expect(page2.svg).toContain('fill="none"');
    expect(page2.svg).toContain('stroke="#64748B"');
    const wrappedText = page2.nodes.find((n) => n.authorId === "grid-text")!;
    assert.ok((wrappedText.box?.height ?? 0) > 45);

    // Page 3: shared visual effects rendered in SVG, pixels, and emitted to PPTX.
    expect(page3.svg).toContain("<linearGradient");
    expect(page3.svg).toContain("<radialGradient");
    expect(page3.svg).toContain("<pattern");
    expect(page3.svg).toContain("<feDropShadow");
    expect(page3.svg).toContain("<feGaussianBlur");
    expect(page3.svg).toContain("stroke-dasharray");
    const slide3Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide3.xml"]!);
    expect(slide3Xml).toContain("<a:blipFill");
    expect(slide3Xml).toContain("<a:outerShdw");
    expect(slide3Xml).toContain("<a:softEdge");

    const p3Png = PNG.sync.read(Buffer.from(page3.png!));
    const px3 = (x: number, y: number) => {
      const i = (y * p3Png.width + x) * 4;
      return [...p3Png.data.subarray(i, i + 4)];
    };
    // borderRadius=14 clips top-left corner of linear-rounded-dashed (24..204, 24..114).
    assert.deepEqual(px3(25, 25), [255, 255, 255, 255]);
    const gradLeft = px3(36, 69);
    const gradRight = px3(192, 69);
    assert.ok(gradLeft[2]! > gradLeft[0]! && gradRight[0]! > 100);
    // radial-dotted (228..408, 24..114): warm center vs navy outer area.
    assert.ok(px3(318, 69)[0]! > 220 && px3(236, 32)[2]! > 100);
    // image-png-fill (432..612, 24..114): rounded corner clipped, left orange, right blue.
    assert.deepEqual(px3(433, 25), [255, 255, 255, 255]);
    assert.ok(px3(470, 69)[0]! > 200 && px3(570, 69)[2]! > 180);
    // image-gif-fill (24..204, 138..228): teal #10B981.
    assert.deepEqual(px3(114, 183), [0x10, 0xb9, 0x81, 255]);
    // soft-edge-box (432..612, 138..228): solid center vs feathered inner edge.
    assert.deepEqual(px3(522, 183), [0x0e, 0xa5, 0xe9, 255]);
    assert.ok(px3(434, 183)[0]! > 100);

    // Page 4: all 6 ArrowType values present in both PPTX and preview SVG + rendered pixels.
    const slide4Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide4.xml"]!);
    const p4Png = PNG.sync.read(Buffer.from(page4.png!));
    for (const arrowType of arrowTypes) {
      expect(slide4Xml).toContain(`<a:headEnd type="${arrowType}"/>`);
      expect(slide4Xml).toContain(`<a:tailEnd type="${arrowType}"/>`);
      const lineNode = page4.nodes.find(
        (n) => n.authorId === `line-${arrowType}`,
      );
      const curveNode = page4.nodes.find(
        (n) => n.authorId === `curve-${arrowType}`,
      );
      assert.ok(lineNode !== undefined, `missing line-${arrowType}`);
      assert.ok(curveNode !== undefined, `missing curve-${arrowType}`);
      const cx = Math.round(
        (lineNode.box?.x ?? 0) + (lineNode.box?.width ?? 0) / 2,
      );
      const cy = Math.round(
        (lineNode.box?.y ?? 0) + (lineNode.box?.height ?? 0) / 2,
      );
      let hasInk = false;
      let hasGap = false;
      for (let dx = -12; dx <= 12; dx += 1) {
        const idx = (cy * p4Png.width + cx + dx) * 4;
        const r = p4Png.data[idx] ?? 255;
        if (r < 200) hasInk = true;
        if (r === 255) hasGap = true;
      }
      assert.ok(hasInk, `expected rendered stroke at line-${arrowType}`);
      if (arrowType === "arrow") {
        assert.ok(hasGap, "expected dashed gap on dashed line-arrow");
      }
    }
  },
};

export const boxRejectionScenario: CapabilityScenario = {
  id: "box-unsupported-styles-and-rejections",
  capability: "box",
  kind: "rejection",
  description:
    "Dual-author Box negative preview scenario verifying corrupt PNG, JPEG, and GIF backgroundImage streams report preview-corrupt-image while keeping valid sibling pages and PPTX.",
  visualCheckpoints: [
    "Pages 1-3 omit PNG/SVG and carry located preview-corrupt-image error diagnostics on corrupt PNG, JPEG, and GIF image-filled Boxes.",
    "Page 4 (box-valid-sibling): renders a white canvas with an emerald #ECFDF5 / #10B981 bordered Box containing a dark #065F46 Shape.",
  ],
  nativeCheckpoints: [
    "Exported PPTX passes validation and retains embedded background image parts on Pages 1-3 even when static preview rejects corrupt pixel streams.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 320, height: 180 },
      createElement(
        Slide,
        { id: "box-corrupt-png-reject" },
        createElement(Box, {
          id: "corrupt-png-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptPngBytes },
          },
        }),
      ),
      createElement(
        Slide,
        { id: "box-corrupt-jpeg-reject" },
        createElement(Box, {
          id: "corrupt-jpeg-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptJpegBytes },
          },
        }),
      ),
      createElement(
        Slide,
        { id: "box-corrupt-gif-reject" },
        createElement(Box, {
          id: "corrupt-gif-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptGifBytes },
          },
        }),
      ),
      createElement(
        Slide,
        { id: "box-valid-sibling", background: "#FFFFFF" },
        createElement(
          Box,
          {
            id: "valid-box",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 240,
              height: 100,
              padding: 16,
              background: "#ECFDF5",
              border: { color: "#10B981", width: 2 },
            },
          },
          createElement(Shape, {
            id: "ok-dot",
            style: { width: 60, height: 60, background: "#065F46" },
          }),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 320, height: 180 }).append(
      slide({ id: "box-corrupt-png-reject" }).append(
        box({
          id: "corrupt-png-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptPngBytes },
          },
        }),
      ),
      slide({ id: "box-corrupt-jpeg-reject" }).append(
        box({
          id: "corrupt-jpeg-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptJpegBytes },
          },
        }),
      ),
      slide({ id: "box-corrupt-gif-reject" }).append(
        box({
          id: "corrupt-gif-box",
          style: {
            width: 140,
            height: 70,
            background: { image: corruptGifBytes },
          },
        }),
      ),
      slide({ id: "box-valid-sibling", background: "#FFFFFF" }).append(
        box({
          id: "valid-box",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 240,
            height: 100,
            padding: 16,
            background: "#ECFDF5",
            border: { color: "#10B981", width: 2 },
          },
        }).append(
          shape({
            id: "ok-dot",
            style: { width: 60, height: 60, background: "#065F46" },
          }),
        ),
      ),
    );
  },
  assert({ preview }) {
    assert.equal(preview.pages.length, 4);
    for (const [idx, boxId] of [
      [0, "corrupt-png-box"],
      [1, "corrupt-jpeg-box"],
      [2, "corrupt-gif-box"],
    ] as const) {
      const page = preview.pages[idx]!;
      assert.equal(page.png, undefined);
      assert.equal(page.svg, undefined);
      assert.equal(page.errors[0]?.code, "preview-corrupt-image");
      assert.match(page.errors[0]?.path ?? "", new RegExp(`box\\[${boxId}\\]`));
      assert.match(page.errors[0]?.message ?? "", /decode/i);
    }
    assert.ok(preview.pages[3]?.png !== undefined);
    assert.deepEqual(preview.pages[3]?.errors, []);
  },
};
