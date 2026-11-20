import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";

import { Box } from "@react-pptx/capabilities/box-elements";
import { font } from "@react-pptx/capabilities/font";
import { Group } from "@react-pptx/capabilities/group-elements";
import { Image } from "@react-pptx/capabilities/image-elements";
import {
  blockArc,
  linearGradient,
  path,
  preset,
} from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { Text } from "@react-pptx/capabilities/text-elements";
import { BuildError } from "@react-pptx/core";
import { Deck, Slide } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import { build } from "../test-helpers.ts";

const imageBytes = new Uint8Array(
  readFileSync(new URL("../assets/wide-20000x1.png", import.meta.url)),
);
const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
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
          id: "block-arc",
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
          id: "svg-path",
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
          id: "styled-image",
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
          id: "styled-text",
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
          id: "rotated-group",
          style: { rotate: 12 },
          children: createElement(Shape, {
            id: "group-member",
            x: 760,
            y: 10,
            width: 80,
            height: 60,
            style: { backgroundColor: "accent6" },
          }),
        }),
        createElement(Box, {
          key: "box-visible",
          id: "visible-box",
          style: {
            width: 70,
            height: 50,
            backgroundColor: "accent2",
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
    assert.ok(
      result.report.diagnostics.some(
        (item) => item.code === "geometry-adjustment-range",
      ),
    );
    const observed = result.report.slides[0]!.objects;
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
              item.code === "invalid-visual-style" &&
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
          x: 0,
          y: 0,
          width: 100,
          height: 100,
          style: { borderRadius: 10 },
        }),
        createElement(Shape, {
          key: "radius-20",
          geometry,
          x: 120,
          y: 0,
          width: 100,
          height: 100,
          style: { borderRadius: 20 },
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
          x: 0,
          y: 0,
          width: 300.42,
          height: "auto",
          style: { fontFamily: jost, fontSize: 20, letterSpacing: 1.25 },
          children: "AAAAAAAAAAAAAAAAAAAA",
        }),
        createElement(Text, {
          key: "spacing-equivalent",
          id: "spacing-equivalent",
          x: 0,
          y: 60,
          width: 300.42,
          height: "auto",
          style: {
            fontFamily: jost,
            fontSize: 20,
            letterSpacing: 94 / 75,
          },
          children: "AAAAAAAAAAAAAAAAAAAA",
        }),
        createElement(Text, {
          key: "uppercase-expansion",
          id: "uppercase-expansion",
          x: 0,
          y: 120,
          width: 16.44,
          height: "auto",
          style: { fontFamily: jost, fontSize: 20, textTransform: "uppercase" },
          children: "ß",
        }),
        createElement(Text, {
          key: "uppercase-literal",
          id: "uppercase-literal",
          x: 0,
          y: 240,
          width: 16.44,
          height: "auto",
          style: { fontFamily: jost, fontSize: 20, textTransform: "uppercase" },
          children: "SS",
        }),
      ]),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objects = result.report.slides[0]?.objects ?? [];
    const height = (testId: string) =>
      objects.find((object) => object.authorId === testId)?.box?.height;

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
          'import { path } from "@react-pptx/capabilities/shape";',
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
          x: 0,
          y: 0,
          width: 30,
          height: 30,
          style: { borderColor: "accent1", borderWidth: 2112 },
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
            x: 0,
            y: 0,
            width: 30,
            height: 30,
            style: { borderColor: "accent1", borderWidth: 3000 },
          }),
        ]),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-visual-style" &&
            path.includes("shape[line-too-wide]"),
        ),
    );
  });

  it("rejects custom geometry path space that quantizes to zero", async () => {
    await assert.rejects(
      build(
        deck([
          createElement(Shape, {
            id: "tiny-view-box",
            x: 0,
            y: 0,
            width: 100,
            height: 100,
            geometry: path("M0 0 L.0001 .0001", {
              viewBox: [0.0001, 0.0001],
            }),
            style: { borderColor: "accent1" },
          }),
        ]),
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path: diagnosticPath }) =>
            code === "invalid-geometry" &&
            diagnosticPath.includes("shape[tiny-view-box]"),
        ),
    );

    const valid = await build(
      deck([
        createElement(Shape, {
          id: "one-unit-view-box",
          x: 0,
          y: 0,
          width: 100,
          height: 100,
          geometry: path("M0 0 L.001 .001", { viewBox: [0.001, 0.001] }),
          style: { borderColor: "accent1" },
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
          x: 0,
          y: 0,
          width: 30,
          height: "auto",
          style: { fontFamily: jost, fontSize: 20, textWrap: "normal" },
          children: content,
        }),
        createElement(Text, {
          key: "no-wrap",
          id: "no-wrap",
          x: 80,
          y: 0,
          width: 30,
          height: "auto",
          style: { fontFamily: jost, fontSize: 20, textWrap: "nowrap" },
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
              width: 20,
              height: 20,
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
              width: 20,
              height: 20,
              style: { minWidth: 1 },
            }),
          ],
        }),
        createElement(Shape, {
          key: "shape-no-wrap",
          id: "shape-no-wrap",
          x: 300,
          y: 0,
          width: 100,
          height: 50,
          text: content,
          style: { fontFamily: jost, fontSize: 20, textWrap: "nowrap" },
        }),
      ]),
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    const objects = result.report.slides[0]?.objects ?? [];
    const object = (testId: string) =>
      objects.find((candidate) => candidate.authorId === testId);
    assert.ok(
      (object("normal-wrap")?.box?.height ?? 0) >
        (object("no-wrap")?.box?.height ?? 0),
    );
    assert.ok(
      (object("nowrap-flex-text")?.box?.width ?? 0) >
        (object("normal-flex-text")?.box?.width ?? 0),
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

  it("writes valid DrawingML headEnd and tailEnd line arrows on preset and custom path shapes", async () => {
    const result = await build(
      deck([
        createElement(Shape, {
          id: "arrow-line",
          x: 10,
          y: 20,
          width: 100,
          height: 1,
          geometry: preset("line"),
          style: {
            borderColor: "accent1",
            borderWidth: 2,
            startArrow: "oval",
            endArrow: "triangle",
          },
        }),
        createElement(Shape, {
          id: "arrow-path",
          x: 20,
          y: 40,
          width: 150,
          height: 80,
          geometry: path("M 0 0 L 150 80", {
            viewBox: [150, 80],
            fill: false,
          }),
          style: {
            borderColor: "accent2",
            borderWidth: 1.5,
            borderStyle: "dashed",
            startArrow: "stealth",
            endArrow: "diamond",
          },
        }),
        createElement(Shape, {
          id: "arrow-none-arrow",
          x: 30,
          y: 60,
          width: 80,
          height: 1,
          geometry: preset("line"),
          style: {
            borderColor: "accent3",
            borderWidth: 1,
            startArrow: "none",
            endArrow: "arrow",
          },
        }),
      ]),
    );

    assert.deepEqual(result.report.diagnostics, []);
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
              x: 10,
              y: 20,
              width: 100,
              height: 1,
              geometry: preset("line"),
              style: {
                borderColor: "accent1",
                borderWidth: 2,
                startArrow: invalidType as any,
              },
            }),
          ]),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (diag) =>
              diag.code === "invalid-visual-style" &&
              diag.message.includes("startArrow"),
          ),
      );

      await assert.rejects(
        build(
          deck([
            createElement(Shape, {
              id: "bad-end-arrow",
              x: 10,
              y: 20,
              width: 100,
              height: 1,
              geometry: preset("line"),
              style: {
                borderColor: "accent1",
                borderWidth: 2,
                endArrow: invalidType as any,
              },
            }),
          ]),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (diag) =>
              diag.code === "invalid-visual-style" &&
              diag.message.includes("endArrow"),
          ),
      );
    }
  });

  it("validates line end type attributes during validatePptx", async () => {
    const valid = await build(
      deck([
        createElement(Shape, {
          id: "arrow-sample",
          x: 10,
          y: 20,
          width: 100,
          height: 1,
          geometry: preset("line"),
          style: {
            borderColor: "accent1",
            borderWidth: 2,
            endArrow: "triangle",
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
            {/* @ts-expect-error -- testing unknown style on Box */}
            <Box style={{ startArrow: "triangle" }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            (diag.code === "invalid-style" || diag.code === "unknown-style") &&
            diag.message.includes("startArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="box-end-arrow">
            {/* @ts-expect-error -- testing unknown style on Box */}
            <Box style={{ endArrow: "triangle" }} />
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            (diag.code === "invalid-style" || diag.code === "unknown-style") &&
            diag.message.includes("endArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="text-start-arrow">
            {/* @ts-expect-error -- testing unknown style on Text */}
            <Text style={{ startArrow: "triangle" }}>Hello</Text>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            (diag.code === "invalid-style" || diag.code === "unknown-style") &&
            diag.message.includes("startArrow"),
        ),
    );

    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="text-end-arrow">
            {/* @ts-expect-error -- testing unknown style on Text */}
            <Text style={{ endArrow: "triangle" }}>Hello</Text>
          </Slide>
        </Deck>,
      ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            (diag.code === "invalid-style" || diag.code === "unknown-style") &&
            diag.message.includes("endArrow"),
        ),
    );
  });
});
