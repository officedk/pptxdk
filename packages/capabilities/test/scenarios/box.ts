import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Deck, Slide } from "@react-pptx/core";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { expect } from "vitest";

import { Box } from "../../src/box-elements.ts";
import { box, deck, font, shape, slide, text } from "../../src/document.ts";
import { Shape } from "../../src/shape-elements.ts";
import { Text } from "../../src/text-elements.ts";
import { linearGradient } from "../../src/visual.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
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
        "../../../react-pptx/test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

export const boxPositiveScenario: CapabilityScenario = {
  id: "box-containers-and-nesting",
  capability: "rp:box",
  description:
    "Dual-author Box scenario covering transparent Flex/Grid layout containers, visible filled/bordered/rotated/linked Boxes, nested Shape & wrapped Text, and scoped ThemeScope overrides.",
  visualCheckpoints: [
    "Page 1 (box-nesting-and-styles): light slate #F8FAFC canvas; transparent outer Flex container holds two side-by-side visible cards. Left card (#EFF6FF fill, 2px #2563EB border) shows a blue #1D4ED8 Shape badge above wrapped CJK/Latin text '容器布局 Shared Box Layout'. Right card (#FEF3C7 fill, 2px #D97706 border, 0.9 opacity, 3deg rotation) shows a teal #0F766E Shape badge above wrapped text 'Click to inspect theme scope & grid'.",
    "Page 2 (box-theme-and-wrapping): white canvas; green-tinted Grid card (#ECFDF5 fill, 2px #10B981 border resolved from local ThemeScope) places a #0F766E Shape in column 1 and wrapped multi-line Jost text in column 2; below it, a hollow #64748B bordered Box frames a nested coral #F97316 Shape without filling the outer box.",
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
          colors: { accentInk: "#991B1B", surface: "#FEE2E2" },
          fonts: { headingFont: notoFont },
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
                  backgroundColor: "#EFF6FF",
                  borderColor: "#2563EB",
                  borderWidth: 2,
                  borderStyle: "solid",
                },
              },
              createElement(Shape, {
                id: "primary-badge",
                width: 56,
                height: 28,
                fill: "#1D4ED8",
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
                  backgroundColor: "#FEF3C7",
                  borderColor: "#D97706",
                  borderWidth: 2,
                  borderStyle: "solid",
                  opacity: 0.9,
                  rotate: 3,
                },
              },
              createElement(Shape, {
                id: "accent-badge",
                width: 56,
                height: 28,
                fill: "#0F766E",
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
            theme: {
              colors: { accentInk: "#0F766E", surface: "#ECFDF5" },
              fonts: { headingFont: jostFont },
            },
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
              color: "accentInk",
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
                backgroundColor: "surface",
                borderColor: "#10B981",
                borderWidth: 2,
                borderStyle: "solid",
              },
            },
            createElement(Shape, {
              id: "grid-swatch",
              width: 120,
              height: 72,
              fill: "accentInk",
            }),
            createElement(
              Text,
              {
                id: "grid-text",
                style: {
                  fontFamily: "headingFont",
                  fontSize: 17,
                  lineHeight: 25,
                },
              },
              "ThemeScope overrides surface fill, accentInk color, and headingFont inside nested Grid Box.",
            ),
          ),
          createElement(
            Box,
            {
              id: "hollow-frame",
              style: {
                width: 320,
                padding: 12,
                borderColor: "#64748B",
                borderWidth: 2,
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
                width: 90,
                height: 36,
                fill: "#F97316",
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
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      theme: {
        colors: { accentInk: "#991B1B", surface: "#FEE2E2" },
        fonts: { headingFont: notoFont },
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
                backgroundColor: "#EFF6FF",
                borderColor: "#2563EB",
                borderWidth: 2,
                borderStyle: "solid",
              },
            }).append(
              shape({
                id: "primary-badge",
                width: 56,
                height: 28,
                fill: "#1D4ED8",
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
                backgroundColor: "#FEF3C7",
                borderColor: "#D97706",
                borderWidth: 2,
                borderStyle: "solid",
                opacity: 0.9,
                rotate: 3,
              },
            }).append(
              shape({
                id: "accent-badge",
                width: 56,
                height: 28,
                fill: "#0F766E",
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
          theme: {
            colors: { accentInk: "#0F766E", surface: "#ECFDF5" },
            fonts: { headingFont: jostFont },
          },
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
            color: "accentInk",
          },
        }).append(
          box({
            id: "grid-panel",
            style: {
              display: "grid",
              gridTemplateColumns: "120px 1fr",
              gap: 16,
              padding: 16,
              backgroundColor: "surface",
              borderColor: "#10B981",
              borderWidth: 2,
              borderStyle: "solid",
            },
          }).append(
            shape({
              id: "grid-swatch",
              width: 120,
              height: 72,
              fill: "accentInk",
            }),
            text(
              "ThemeScope overrides surface fill, accentInk color, and headingFont inside nested Grid Box.",
              {
                id: "grid-text",
                style: {
                  fontFamily: "headingFont",
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
              borderColor: "#64748B",
              borderWidth: 2,
            },
          }).append(
            box({
              id: "hollow-inner",
              style: { display: "flex", flexDirection: "row", gap: 12 },
            }).append(
              shape({
                id: "coral-chip",
                width: 90,
                height: 36,
                fill: "#F97316",
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
    );
  },
  assert({ resolution, pptx, preview, validation }) {
    assert.equal(preview.pages.length, 2);
    const page1 = preview.pages[0]!;
    const page2 = preview.pages[1]!;
    assert.deepEqual(page1.diagnostics, []);
    assert.deepEqual(page2.diagnostics, []);
    assert.ok(page1.png !== undefined);
    assert.ok(page2.png !== undefined);

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
    const shellObserved = resolution.query
      .getSlide("box-nesting-and-styles")
      .get("shell")
      .observe();
    assert.ok(shellObserved.box !== undefined);
    assert.equal(shellObserved.objects.length, 6);
    const accentCard = page1.nodes.find((n) => n.authorId === "card-accent")!;
    assert.equal(accentCard.link, "slide:box-theme-and-wrapping");
    assert.ok(
      (accentCard.screenBox?.height ?? 0) > (accentCard.box?.height ?? 0),
    );

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

    // Page 2: ThemeScope tokens resolved in both SVG and PPTX, hollow frame has fill="none".
    expect(page2.svg).toContain('fill="#ECFDF5"');
    expect(page2.svg).toContain('fill="#0F766E"');
    expect(page2.svg).toContain('fill="none"');
    expect(page2.svg).toContain('stroke="#64748B"');
    const wrappedText = page2.nodes.find((n) => n.authorId === "grid-text")!;
    assert.ok((wrappedText.box?.height ?? 0) > 45);
  },
};

export const boxRejectionScenario: CapabilityScenario = {
  id: "box-unsupported-styles-and-rejections",
  capability: "rp:box",
  description:
    "Dual-author Box negative preview scenario verifying rejection of gradient fills, borderRadius, and dashed borders while keeping valid sibling pages and PPTX.",
  visualCheckpoints: [
    "Pages 1-3 omit PNG/SVG and carry located error diagnostics on gradient, borderRadius, and dashed-border Boxes.",
    "Page 4 (box-valid-sibling): renders a white canvas with an emerald #ECFDF5 / #10B981 bordered Box containing a dark #065F46 Shape.",
  ],
  nativeCheckpoints: [
    "Exported PPTX passes validation and retains native gradient/rounded/dashed Box shapes even when static preview rejects those pages.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 320, height: 180 },
      createElement(
        Slide,
        { id: "box-gradient-reject" },
        createElement(Box, {
          id: "grad-box",
          style: {
            width: 140,
            height: 70,
            backgroundImage: linearGradient(90, "#1D4ED8", "#9333EA"),
          },
        }),
      ),
      createElement(
        Slide,
        { id: "box-radius-reject" },
        createElement(Box, {
          id: "radius-box",
          style: {
            width: 140,
            height: 70,
            backgroundColor: "#E2E8F0",
            borderRadius: 12,
          },
        }),
      ),
      createElement(
        Slide,
        { id: "box-dashed-reject" },
        createElement(Box, {
          id: "dashed-box",
          style: {
            width: 140,
            height: 70,
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
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
              backgroundColor: "#ECFDF5",
              borderColor: "#10B981",
              borderWidth: 2,
            },
          },
          createElement(Shape, {
            id: "ok-dot",
            width: 60,
            height: 60,
            fill: "#065F46",
          }),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 320, height: 180 }).append(
      slide({ id: "box-gradient-reject" }).append(
        box({
          id: "grad-box",
          style: {
            width: 140,
            height: 70,
            backgroundImage: linearGradient(90, "#1D4ED8", "#9333EA"),
          },
        }),
      ),
      slide({ id: "box-radius-reject" }).append(
        box({
          id: "radius-box",
          style: {
            width: 140,
            height: 70,
            backgroundColor: "#E2E8F0",
            borderRadius: 12,
          },
        }),
      ),
      slide({ id: "box-dashed-reject" }).append(
        box({
          id: "dashed-box",
          style: {
            width: 140,
            height: 70,
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
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
            backgroundColor: "#ECFDF5",
            borderColor: "#10B981",
            borderWidth: 2,
          },
        }).append(
          shape({
            id: "ok-dot",
            width: 60,
            height: 60,
            fill: "#065F46",
          }),
        ),
      ),
    );
  },
  assert({ preview }) {
    assert.equal(preview.pages.length, 4);
    assert.equal(preview.pages[0]?.png, undefined);
    assert.match(
      preview.pages[0]?.diagnostics[0]?.path ?? "",
      /box\[grad-box\]/,
    );
    assert.equal(preview.pages[1]?.png, undefined);
    assert.match(
      preview.pages[1]?.diagnostics[0]?.path ?? "",
      /box\[radius-box\]/,
    );
    assert.equal(preview.pages[2]?.png, undefined);
    assert.match(
      preview.pages[2]?.diagnostics[0]?.path ?? "",
      /box\[dashed-box\]/,
    );
    assert.ok(preview.pages[3]?.png !== undefined);
    assert.deepEqual(preview.pages[3]?.diagnostics, []);
  },
};
