import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import { font } from "@pptxdk/capabilities/font";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Image } from "@pptxdk/capabilities/image-elements";
import {
  blockArc,
  linearGradient,
  path,
  preset,
  type Geometry,
} from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import { build } from "../../../test/test-helpers.ts";

const imageBytes = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/wide-20000x1.png", import.meta.url),
  ),
);
const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-400-Book.ttf",
          import.meta.url,
        ),
      ),
    ),
  },
]);

function deck(children: readonly ReturnType<typeof createElement>[]) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "geometry-visual" }, ...children),
  );
}

describe("geometry, object visuals, and text visuals", () => {
  it("publicly builds and independently reads presets, SVG, custom geometries, and applicable object visuals", async () => {
    const result = await build(
      deck([
        createElement(Shape, {
          key: "preset",
          id: "preset-callout",
          geometry: preset("wedgeRoundRectCallout", {
            adj1: 28_000,
            adj2: 42_000,
            adj3: 16_667,
          }),
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 160,
            height: 100,
            flipH: true,
            background: { color: "accent1", alpha: 0.7 },
            border: { color: "accent2", width: 2, style: "dashed" },
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
          id: "block-arc",
          geometry: blockArc({
            startAngle: 20,
            endAngle: 260,
            thickness: 0.75,
          }),
          style: {
            position: "absolute",
            left: 190,
            top: 10,
            width: 100,
            height: 100,
            background: linearGradient(45, "accent1", "accent4"),
            opacity: 0.8,
          },
        }),
        createElement(Shape, {
          key: "svg",
          id: "svg-path",
          geometry: path(
            [
              "M 8 8 H 92 V 92 H 8 Z",
              "M 30 30 C 40 10 60 10 70 30 S 60 55 50 50 Q 40 45 35 55 T 30 70 A 10 8 20 0 1 30 30 Z",
            ],
            { viewBox: [100, 100], fillRule: "evenodd" },
          ),
          style: {
            position: "absolute",
            left: 310,
            top: 10,
            width: 100,
            height: 100,
            background: "accent3",
          },
        }),
        createElement(Image, {
          key: "image",
          id: "styled-image",
          src: imageBytes,
          geometry: "cloud",
          style: {
            position: "absolute",
            left: 430,
            top: 10,
            width: 120,
            height: 80,
            flipV: true,
            border: { color: "accent2", width: 1 },
            opacity: 0.5,
            softEdge: 2,
          },
        }),
        createElement(Text, {
          key: "text",
          id: "styled-text",
          style: {
            position: "absolute",
            left: 570,
            top: 10,
            width: 160,
            height: 60,
            fontFamily: "Jost",
            letterSpacing: 1.25,
            textTransform: "uppercase",
            textDecoration: "underline",
            textShadow: { color: "accent5", blur: 2 },
            background: "lt2",
            borderRadius: 8,
          },
          children: "ßeta",
        }),
        createElement(Group, {
          key: "group",
          id: "rotated-group",
          style: { rotate: 12 },
          children: createElement(Shape, {
            id: "group-member",
            style: {
              position: "absolute",
              left: 760,
              top: 10,
              width: 80,
              height: 60,
              background: "accent6",
            },
          }),
        }),
        createElement(Box, {
          key: "box-visible",
          id: "visible-box",
          style: {
            width: 70,
            height: 50,
            background: "accent2",
            borderRadius: 10,
            rotate: 15,
          },
        }),
        createElement(Box, {
          key: "box-transparent",
          id: "transparent-box",
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
    assert.match(
      slideXml,
      /<a:prstGeom prst="blockArc"><a:avLst><a:gd name="adj1" fmla="val 1200000"\/><a:gd name="adj2" fmla="val 15600000"\/><a:gd name="adj3" fmla="val 50000"\/><\/a:avLst><\/a:prstGeom>/u,
    );
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
    assert.match(slideXml, /rot="900000"/u);
    assert.match(slideXml, /flipH="1"/u);
    assert.match(slideXml, /flipV="1"/u);
    const observed = result.slide("geometry-visual").objects;
    assert.ok(
      observed.some(
        (item) => item.authorId === "visible-box" && item.kind === "shape",
      ),
    );
    assert.ok(!observed.some((item) => item.authorId === "transparent-box"));
    assert.ok(
      file.slides[0]!.objects.some((item) => item.geometry === "custom"),
    );
    assert.ok(
      file.slides[0]!.objects.some(
        (item) => item.geometry === "wedgeRoundRectCallout",
      ),
    );
  });

  it("allows transparent Box elements without native size to carry only layout semantics", async () => {
    const result = await build(
      deck([
        createElement(Box, { key: "empty", id: "empty-box" }),
        createElement(Box, {
          key: "zero-height",
          id: "zero-height-box",
          style: { width: 100, height: 0, rotate: 15 },
        }),
      ]),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.ok(
      !result.report.slides[0]?.objects.some(
        ({ authorId: testId }) =>
          testId === "empty-box" || testId === "zero-height-box",
      ),
    );
  });

  it("rejects invalid rotation declarations on transparent Box elements according to writable units", async () => {
    for (const rotate of [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      "bad",
    ] as const) {
      await assert.rejects(
        build(
          deck([
            createElement(Box, {
              id: "badrotate",
              style: {
                width: 100,
                height: 20,
                rotate: rotate as unknown as number,
              },
            }),
          ]),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (item) =>
              item.code === "invalid-style" &&
              item.path.includes("badrotate") &&
              item.message.includes("rotate"),
          ),
      );
    }
  });

  it("derives corner radius adjustment values for reused presets without mutating author inputs", async () => {
    const geometry = preset("roundRect", Object.freeze({}));
    const result = await build(
      deck([
        createElement(Shape, {
          key: "radius-10",
          geometry,
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
            borderRadius: 10,
          },
        }),
        createElement(Shape, {
          key: "radius-20",
          geometry,
          style: {
            position: "absolute",
            left: 120,
            top: 0,
            width: 100,
            height: 100,
            borderRadius: 20,
          },
        }),
      ]),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /fmla="val 10000"/u);
    assert.match(slideXml, /fmla="val 20000"/u);
    assert.deepEqual(geometry.adjust, {});
  });

  it("measures auto height using character bounds after actual kerning quantization and case transformation", async () => {
    const result = await build(
      deck([
        createElement(Text, {
          key: "spacing-decimal",
          id: "spacing-decimal",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 300.42,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            letterSpacing: 1.25,
          },
          children: "AAAAAAAAAAAAAAAAAAAA",
        }),
        createElement(Text, {
          key: "spacing-equivalent",
          id: "spacing-equivalent",
          style: {
            position: "absolute",
            left: 0,
            top: 60,
            width: 300.42,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            letterSpacing: 94 / 75,
          },
          children: "AAAAAAAAAAAAAAAAAAAA",
        }),
        createElement(Text, {
          key: "uppercase-expansion",
          id: "uppercase-expansion",
          style: {
            position: "absolute",
            left: 0,
            top: 120,
            width: 16.44,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            textTransform: "uppercase",
          },
          children: "ß",
        }),
        createElement(Text, {
          key: "uppercase-literal",
          id: "uppercase-literal",
          style: {
            position: "absolute",
            left: 0,
            top: 240,
            width: 16.44,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            textTransform: "uppercase",
          },
          children: "SS",
        }),
      ]),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideLoc = result.slide("geometry-visual");
    const height = (testId: string) => slideLoc.get(testId).box?.height;

    assert.equal(height("spacing-decimal"), height("spacing-equivalent"));
    assert.equal(height("uppercase-expansion"), height("uppercase-literal"));
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.equal((slideXml.match(/spc="94"/gu) ?? []).length, 2);
    assert.equal((slideXml.match(/cap="all"/gu) ?? []).length, 2);
  });

  it("rejects numbers following close and non-0/1 arc flags with exact location", () => {
    const child = spawnSync(
      process.execPath,
      [
        "--conditions=source",
        "--input-type=module",
        "--eval",
        [
          'import { path } from "@pptxdk/capabilities/shape";',
          'for (const data of ["M0 0Z1", "M0 0A10 10 0 2 3 20 20", "M0 0A10 10 0 1.0 0 20 20"]) {',
          "  try { path(data, { viewBox: [30, 30] }); process.exitCode = 1; break; }",
          '  catch (error) { if (error?.code !== "invalid-geometry") { process.exitCode = 2; break; } }',
          "}",
          'if (process.exitCode === undefined) path("M0 0Z m5 5 l1 1", { viewBox: [30, 30] });',
          'if (process.exitCode === undefined) process.stdout.write("ok");',
        ].join("\n"),
      ],
      { timeout: 2_000 },
    );

    assert.equal(
      child.status,
      0,
      `invalid path probe failed: ${child.error?.message ?? child.stderr.toString()}`,
    );
    assert.equal(child.stdout.toString(), "ok");
  });

  it("rejects unwritable widths according to final DrawingML stroke EMU bounds", async () => {
    const valid = await build(
      deck([
        createElement(Shape, {
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 30,
            height: 30,
            border: { color: "accent1", width: 2112 },
          },
        }),
      ]),
    );
    const file = await validatePptx(valid.pptx);
    assert.deepEqual(file.errors, []);
    assert.match(
      strFromU8(unzipSync(valid.pptx)["ppt/slides/slide1.xml"]!),
      /<a:ln w="20116800">/u,
    );

    await assert.rejects(
      build(
        deck([
          createElement(Shape, {
            id: "line-too-wide",
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 30,
              height: 30,
              border: { color: "accent1", width: 3000 },
            },
          }),
        ]),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-style" && path.includes("shape[line-too-wide]"),
        ),
    );
  });

  it("rejects custom geometry path space that quantizes to zero", async () => {
    await assert.rejects(
      build(
        deck([
          createElement(Shape, {
            id: "tiny-view-box",
            geometry: path("M0 0 L.0001 .0001", {
              viewBox: [0.0001, 0.0001],
            }),
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 100,
              height: 100,
              border: { color: "accent1" },
            },
          }),
        ]),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path: diagnosticPath }) =>
            code === "invalid-prop" &&
            diagnosticPath.includes("shape[tiny-view-box]"),
        ),
    );

    const valid = await build(
      deck([
        createElement(Shape, {
          id: "one-unit-view-box",
          geometry: path("M0 0 L.001 .001", { viewBox: [0.001, 0.001] }),
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 100,
            border: { color: "accent1" },
          },
        }),
      ]),
    );
    assert.deepEqual((await validatePptx(valid.pptx)).errors, []);
    assert.match(
      strFromU8(unzipSync(valid.pptx)["ppt/slides/slide1.xml"]!),
      /<a:path w="1" h="1">/u,
    );
  });

  it("preserves normal line breaks and shares measurement and native wrap between nowrap Text and Shape", async () => {
    const content = "alpha beta gamma delta";
    const result = await build(
      deck([
        createElement(Text, {
          key: "normal-wrap",
          id: "normal-wrap",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 30,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            textWrap: "normal",
          },
          children: content,
        }),
        createElement(Text, {
          key: "no-wrap",
          id: "no-wrap",
          style: {
            position: "absolute",
            left: 80,
            top: 0,
            width: 30,
            height: "auto",
            fontFamily: jost,
            fontSize: 20,
            textWrap: "nowrap",
          },
          children: content,
        }),
        createElement(Box, {
          key: "normal-min-content",
          id: "normal-min-content",
          style: {
            display: "flex",
            flexDirection: "row",
            width: 160,
            height: 40,
            alignItems: "start",
          },
          children: [
            createElement(Text, {
              key: "normal-flex-text",
              id: "normal-flex-text",
              style: { fontFamily: jost, fontSize: 20, textWrap: "normal" },
              children: content,
            }),
            createElement(Shape, {
              key: "normal-flex-marker",
              id: "normal-flex-marker",
              style: { width: 20, height: 20 },
            }),
          ],
        }),
        createElement(Box, {
          key: "nowrap-min-content",
          id: "nowrap-min-content",
          style: {
            display: "flex",
            flexDirection: "row",
            width: 160,
            height: 40,
            alignItems: "start",
          },
          children: [
            createElement(Text, {
              key: "nowrap-flex-text",
              id: "nowrap-flex-text",
              style: { fontFamily: jost, fontSize: 20, textWrap: "nowrap" },
              children: content,
            }),
            createElement(Shape, {
              key: "nowrap-flex-marker",
              id: "nowrap-flex-marker",
              style: { width: 20, height: 20, minWidth: 1 },
            }),
          ],
        }),
        createElement(
          Shape,
          {
            key: "shape-no-wrap",
            id: "shape-no-wrap",
            style: {
              position: "absolute",
              left: 300,
              top: 0,
              width: 100,
              height: 50,
              fontFamily: jost,
              fontSize: 20,
              textWrap: "nowrap",
            },
          },
          content,
        ),
      ]),
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    const slideLoc = result.slide("geometry-visual");
    assert.ok(
      (slideLoc.get("normal-wrap").box?.height ?? 0) >
        (slideLoc.get("no-wrap").box?.height ?? 0),
    );
    assert.ok(
      (slideLoc.get("nowrap-min-content").get("nowrap-flex-text").box?.width ??
        0) >
        (slideLoc.get("normal-min-content").get("normal-flex-text").box
          ?.width ?? 0),
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:bodyPr wrap="square"/u);
    assert.equal((slideXml.match(/<a:bodyPr wrap="none"/gu) ?? []).length, 3);
  });

  it("rejects invalid geometry, gradient, text spacing, and unrepresentable combinations with exact location", async () => {
    assert.throws(
      () => path("M 0 0 L 2 ?", { viewBox: [10, 10] }),
      (error: unknown) =>
        error instanceof Error && error.message.includes("SVG path"),
    );
    const invalidCases: readonly [ReturnType<typeof createElement>, string][] =
      [
        [
          createElement(Shape, {
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 10,
              height: 10,
            },
            geometry: "unknownPreset" as Geometry,
          }),
          "invalid-prop",
        ],
        [
          createElement(Shape, {
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 10,
              height: 10,
              background: linearGradient(
                0,
                { color: "accent1", position: 0 },
                { color: "accent2", position: 1.1 },
              ),
            },
          }),
          "invalid-prop",
        ],
        [
          createElement(Text, {
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 80,
              height: 30,
              letterSpacing: 6000,
            },
            children: "too wide",
          }),
          "invalid-style",
        ],
        [
          createElement(Shape, {
            geometry: path("M0 0 L10 10", { viewBox: [10, 10] }),
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 10,
              height: 10,
              borderRadius: 2,
            },
          }),
          "invalid-style",
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

  it("writes valid DrawingML headEnd and tailEnd line arrows on preset and custom path shapes", async () => {
    const result = await build(
      deck([
        createElement(Shape, {
          id: "arrow-line",
          geometry: preset("line"),
          style: {
            position: "absolute",
            left: 10,
            top: 20,
            width: 100,
            height: 1,
            border: {
              color: "accent1",
              width: 2,
              startArrow: "oval",
              endArrow: "triangle",
            },
          },
        }),
        createElement(Shape, {
          id: "arrow-path",
          geometry: path("M 0 0 L 150 80", {
            viewBox: [150, 80],
            fill: false,
          }),
          style: {
            position: "absolute",
            left: 20,
            top: 40,
            width: 150,
            height: 80,
            border: {
              color: "accent2",
              width: 1.5,
              style: "dashed",
              startArrow: "stealth",
              endArrow: "diamond",
            },
          },
        }),
        createElement(Shape, {
          id: "arrow-none-arrow",
          geometry: preset("line"),
          style: {
            position: "absolute",
            left: 30,
            top: 60,
            width: 80,
            height: 1,
            border: {
              color: "accent3",
              width: 1,
              startArrow: "none",
              endArrow: "arrow",
            },
          },
        }),
      ]),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const parts = unzipSync(result.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);

    // Verify headEnd and tailEnd elements exist
    assert.ok(slideXml.includes('<a:headEnd type="oval"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="triangle"/>'));
    assert.ok(slideXml.includes('<a:headEnd type="stealth"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="diamond"/>'));
    assert.ok(slideXml.includes('<a:headEnd type="none"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="arrow"/>'));

    // Verify DrawingML schema child element order: solidFill -> prstDash -> headEnd -> tailEnd
    const dashedLineMatch = slideXml.match(
      /<a:ln [^>]*>(.*?)<a:prstDash val="dash"\/>(.*?)<a:headEnd type="stealth"\/>(.*?)<a:tailEnd type="diamond"\/>.*?<\/a:ln>/s,
    );
    assert.ok(
      dashedLineMatch,
      "Dashed line elements must appear in DrawingML schema sequence",
    );
    assert.ok(dashedLineMatch[1]?.includes("<a:solidFill>"));
  });

  it("rejects invalid line arrow types on Shape style", async () => {
    for (const invalidType of ["circle", "bar", "unknown", 123] as const) {
      await assert.rejects(
        build(
          deck([
            createElement(Shape, {
              id: "bad-start-arrow",
              geometry: preset("line"),
              style: {
                position: "absolute",
                left: 10,
                top: 20,
                width: 100,
                height: 1,
                border: {
                  color: "accent1",
                  width: 2,
                  startArrow: invalidType as any,
                },
              },
            }),
          ]),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diag) => diag.code === "invalid-style"),
      );

      await assert.rejects(
        build(
          deck([
            createElement(Shape, {
              id: "bad-end-arrow",
              geometry: preset("line"),
              style: {
                position: "absolute",
                left: 10,
                top: 20,
                width: 100,
                height: 1,
                border: {
                  color: "accent1",
                  width: 2,
                  endArrow: invalidType as any,
                },
              },
            }),
          ]),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diag) => diag.code === "invalid-style"),
      );
    }
  });

  it("validates line end type attributes during validatePptx", async () => {
    const valid = await build(
      deck([
        createElement(Shape, {
          id: "arrow-sample",
          geometry: preset("line"),
          style: {
            position: "absolute",
            left: 10,
            top: 20,
            width: 100,
            height: 1,
            border: {
              color: "accent1",
              width: 2,
              endArrow: "triangle",
            },
          },
        }),
      ]),
    );

    const parts = unzipSync(valid.pptx);
    const originalSlideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    // Corrupt the tailEnd type attribute
    const corruptedSlideXml = originalSlideXml.replace(
      'type="triangle"',
      'type="curly"',
    );
    parts["ppt/slides/slide1.xml"] = new TextEncoder().encode(
      corruptedSlideXml,
    );

    const { zipSync } = await import("fflate");
    const corruptedPptx = zipSync(parts);
    const file = await validatePptx(corruptedPptx);
    assert.ok(
      file.errors.some((err) => err.code === "invalid-line-end-type"),
      "Expected invalid-line-end-type error on corrupted line end type",
    );
  });

  it("rejects startArrow and endArrow on Box and Text with invalid-style", async () => {
    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="box-start-arrow">
            <Box
              style={{
                border: {
                  color: "accent1",
                  // @ts-expect-error -- testing unsupported arrow on Box border
                  startArrow: "triangle",
                },
              }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-style" &&
            diag.message.includes("startArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="box-end-arrow">
            <Box
              style={{
                border: {
                  color: "accent1",
                  // @ts-expect-error -- testing unsupported arrow on Box border
                  endArrow: "triangle",
                },
              }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-style" && diag.message.includes("endArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="text-start-arrow">
            <Text
              style={{
                border: {
                  color: "accent1",
                  // @ts-expect-error -- testing unsupported arrow on Text border
                  startArrow: "triangle",
                },
              }}
            >
              Hello
            </Text>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-style" &&
            diag.message.includes("startArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="text-end-arrow">
            <Text
              style={{
                border: {
                  color: "accent1",
                  // @ts-expect-error -- testing unsupported arrow on Text border
                  endArrow: "triangle",
                },
              }}
            >
              Hello
            </Text>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-style" && diag.message.includes("endArrow"),
        ),
    );
  });
});
