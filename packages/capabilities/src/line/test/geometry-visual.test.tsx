import assert from "node:assert/strict";

import { Line } from "@pptxdk/capabilities/line-elements";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import { build } from "../../../test/test-helpers.ts";

describe("Line geometry, strokes, and arrows", () => {
  it("writes valid DrawingML headEnd and tailEnd line arrows in schema element sequence", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Line, {
            id: "arrow-line",
            stroke: {
              color: "accent1",
              width: 2,
              startArrow: "oval",
              endArrow: "triangle",
            },
            style: {
              position: "absolute",
              left: 10,
              top: 20,
              width: 100,
              height: 0,
            },
          }),
          createElement(Line, {
            id: "arrow-curved",
            routing: "curved",
            stroke: {
              color: "accent2",
              width: 1.5,
              style: "dashed",
              startArrow: "stealth",
              endArrow: "diamond",
            },
            style: {
              position: "absolute",
              left: 20,
              top: 40,
              width: 150,
              height: 80,
            },
          }),
          createElement(Line, {
            id: "arrow-none-arrow",
            stroke: {
              color: "accent3",
              width: 1,
              startArrow: "none",
              endArrow: "arrow",
            },
            style: {
              position: "absolute",
              left: 30,
              top: 60,
              width: 80,
              height: 0,
            },
          }),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const parts = unzipSync(result.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.ok(slideXml.includes('<a:headEnd type="oval"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="triangle"/>'));
    assert.ok(slideXml.includes('<a:headEnd type="stealth"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="diamond"/>'));
    assert.ok(slideXml.includes('<a:headEnd type="none"/>'));
    assert.ok(slideXml.includes('<a:tailEnd type="arrow"/>'));

    const dashedLineMatch = slideXml.match(
      /<a:ln [^>]*>(.*?)<a:prstDash val="dash"\/>(.*?)<a:headEnd type="stealth"\/>(.*?)<a:tailEnd type="diamond"\/>.*?<\/a:ln>/s,
    );
    assert.ok(dashedLineMatch);
    assert.ok(dashedLineMatch[1]?.includes("<a:solidFill>"));
  });

  it("rejects invalid arrow types on Line stroke and validates line end attributes via validatePptx", async () => {
    for (const invalidType of ["circle", "bar", "unknown", 123] as const) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            { width: 960, height: 540 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Line, {
                id: "bad-start-arrow",
                stroke: {
                  color: "accent1",
                  width: 2,
                  startArrow: invalidType as never,
                },
                style: {
                  position: "absolute",
                  left: 10,
                  top: 20,
                  width: 100,
                  height: 0,
                },
              }),
            ),
          ),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diag) => diag.code === "invalid-style"),
      );
      await assert.rejects(
        build(
          createElement(
            Deck,
            { width: 960, height: 540 },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Line, {
                id: "bad-end-arrow",
                stroke: {
                  color: "accent1",
                  width: 2,
                  endArrow: invalidType as never,
                },
                style: {
                  position: "absolute",
                  left: 10,
                  top: 20,
                  width: 100,
                  height: 0,
                },
              }),
            ),
          ),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diag) => diag.code === "invalid-style"),
      );
    }

    const valid = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(Line, {
            id: "arrow-sample",
            stroke: { color: "accent1", width: 2, endArrow: "triangle" },
            style: {
              position: "absolute",
              left: 10,
              top: 20,
              width: 100,
              height: 0,
            },
          }),
        ),
      ),
    );
    const parts = unzipSync(valid.pptx);
    const corruptedSlideXml = strFromU8(
      parts["ppt/slides/slide1.xml"]!,
    ).replace('type="triangle"', 'type="curly"');
    parts["ppt/slides/slide1.xml"] = new TextEncoder().encode(
      corruptedSlideXml,
    );
    const file = await validatePptx(zipSync(parts));
    assert.ok(
      file.errors.some((err) => err.code === "invalid-line-end-type"),
      "Expected invalid-line-end-type error on corrupted line end type",
    );
  });
});
