import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  blockArc,
  Box,
  BuildError,
  build,
  Deck,
  Group,
  Image,
  linearGradient,
  path,
  preset,
  Shape,
  Slide,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const imageBytes = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

function deck(children: readonly ReturnType<typeof createElement>[]) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { slideId: "geometry-visual" }, ...children),
  );
}

describe("几何、对象视觉和文字视觉", () => {
  it("公开构建并独立读取预设、SVG、自定义几何及适用对象视觉", async () => {
    const result = await build(
      deck([
        createElement(Shape, {
          key: "preset",
          testId: "preset-callout",
          geometry: preset("wedgeRoundRectCallout", {
            adj1: 28_000,
            adj2: 42_000,
          }),
          x: 10,
          y: 10,
          width: 160,
          height: 100,
          flipH: true,
          style: {
            backgroundColor: { color: "accent1", alpha: 0.7 },
            borderColor: "accent2",
            borderWidth: 2,
            borderStyle: "dashed",
            rotate: -15,
            boxShadow: {
              color: { color: "dk1", alpha: 0.4 },
              blur: 3,
              distance: 2,
              angle: 125,
            },
            softEdge: 1,
          },
        }),
        createElement(Shape, {
          key: "arc",
          testId: "block-arc",
          geometry: blockArc({
            startAngle: 20,
            endAngle: 260,
            thickness: 0.75,
          }),
          x: 190,
          y: 10,
          width: 100,
          height: 100,
          style: {
            backgroundImage: linearGradient(45, "accent1", "accent4"),
            opacity: 0.8,
          },
        }),
        createElement(Shape, {
          key: "svg",
          testId: "svg-path",
          geometry: path(
            [
              "M 8 8 H 92 V 92 H 8 Z",
              "M 30 30 C 40 10 60 10 70 30 S 60 55 50 50 Q 40 45 35 55 T 30 70 A 10 8 20 0 1 30 30 Z",
            ],
            { viewBox: [100, 100], fillRule: "evenodd" },
          ),
          x: 310,
          y: 10,
          width: 100,
          height: 100,
          style: { backgroundColor: "accent3" },
        }),
        createElement(Image, {
          key: "image",
          testId: "styled-image",
          src: imageBytes,
          x: 430,
          y: 10,
          width: 120,
          height: 80,
          geometry: "cloud",
          flipV: true,
          style: {
            borderColor: "accent2",
            borderWidth: 1,
            opacity: 0.5,
            softEdge: 2,
          },
        }),
        createElement(Text, {
          key: "text",
          testId: "styled-text",
          x: 570,
          y: 10,
          width: 160,
          height: 60,
          style: {
            fontFamily: "Jost",
            letterSpacing: 1.25,
            textTransform: "uppercase",
            textDecoration: "underline",
            textShadow: { color: "accent5", blur: 2 },
            backgroundColor: "lt2",
            borderRadius: 8,
          },
          children: "ßeta",
        }),
        createElement(Group, {
          key: "group",
          testId: "rotated-group",
          style: { rotate: 12 },
          children: createElement(Shape, {
            testId: "group-member",
            x: 760,
            y: 10,
            width: 80,
            height: 60,
            style: { backgroundColor: "accent6" },
          }),
        }),
        createElement(Box, {
          key: "box-visible",
          testId: "visible-box",
          style: {
            width: 70,
            height: 50,
            backgroundColor: "accent2",
            borderRadius: 10,
          },
        }),
        createElement(Box, {
          key: "box-transparent",
          testId: "transparent-box",
          style: { width: 70, height: 50 },
        }),
      ]),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:prstGeom prst="wedgeRoundRectCallout">/u);
    assert.match(slideXml, /<a:gd name="adj1" fmla="val 28000"\/>/u);
    assert.match(slideXml, /<a:custGeom>/u);
    assert.match(slideXml, /<a:cubicBezTo>/u);
    assert.match(slideXml, /<a:quadBezTo>/u);
    assert.match(slideXml, /<a:gradFill/u);
    assert.match(slideXml, /<a:outerShdw/u);
    assert.match(slideXml, /<a:softEdge/u);
    assert.match(slideXml, /<a:alphaModFix amt="50000"\/>/u);
    assert.match(slideXml, /cap="all"/u);
    assert.match(slideXml, /u="sng"/u);
    assert.match(slideXml, /spc="94"/u);
    assert.match(slideXml, /<a:effectLst><a:outerShdw/u);
    assert.match(slideXml, /rot="720000"/u);
    assert.match(slideXml, /flipH="1"/u);
    assert.match(slideXml, /flipV="1"/u);
    assert.ok(
      result.report.diagnostics.some(
        (item) => item.code === "geometry-adjustment-range",
      ),
    );
    const observed = result.report.slides[0]!.objects;
    assert.ok(
      observed.some(
        (item) => item.testId === "visible-box" && item.kind === "shape",
      ),
    );
    assert.ok(!observed.some((item) => item.testId === "transparent-box"));
    assert.ok(
      file.slides[0]!.objects.some((item) => item.geometry === "custom"),
    );
    assert.ok(
      file.slides[0]!.objects.some(
        (item) => item.geometry === "wedgeRoundRectCallout",
      ),
    );
  });

  it("定位拒绝非法几何、渐变、文本间距和不可表达组合", async () => {
    assert.throws(
      () => path("M 0 0 L 2 ?", { viewBox: [10, 10] }),
      (error: unknown) =>
        error instanceof Error && error.message.includes("SVG 路径"),
    );
    const invalidCases: readonly [ReturnType<typeof createElement>, string][] =
      [
        [
          createElement(Shape, {
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            geometry: "unknownPreset",
          }),
          "unsupported-geometry",
        ],
        [
          createElement(Shape, {
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            style: {
              backgroundImage: linearGradient(
                0,
                { color: "accent1", position: 0 },
                { color: "accent2", position: 1.1 },
              ),
            },
          }),
          "invalid-gradient",
        ],
        [
          createElement(Text, {
            x: 0,
            y: 0,
            width: 80,
            height: 30,
            style: { letterSpacing: 6000 },
            children: "too wide",
          }),
          "invalid-letter-spacing",
        ],
        [
          createElement(Shape, {
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            geometry: path("M0 0 L10 10", { viewBox: [10, 10] }),
            style: { borderRadius: 2 },
          }),
          "unsupported-style-combination",
        ],
      ];

    for (const [element, code] of invalidCases) {
      await assert.rejects(
        build(deck([element])),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((item) => item.code === code),
      );
    }
  });
});
