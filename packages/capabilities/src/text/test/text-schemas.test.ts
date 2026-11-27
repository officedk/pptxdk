import assert from "node:assert/strict";

import {
  slideBackgroundCapability,
  textCapability,
} from "@pptxdk/capabilities";
import { font } from "@pptxdk/capabilities/font";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  paragraphStyleSchema,
  parseParagraphStyle,
  parseRunStyle,
  parseTextInsets,
  runStyleSchema,
  textListSchema,
} from "../../shared/text-schemas.ts";

describe("Text & font local Zod schemas", () => {
  it("validates textList, paragraphStyle, runStyle, and insets schemas cleanly", () => {
    const bullet = textListSchema.parse({
      type: "bullet",
      character: "•",
      level: 2,
    });
    assert.deepEqual(bullet, { type: "bullet", character: "•", level: 2 });

    const numbered = textListSchema.parse({
      type: "number",
      style: "arabicPeriod",
      startAt: 3,
      level: 1,
    });
    assert.deepEqual(numbered, {
      type: "number",
      style: "arabicPeriod",
      startAt: 3,
      level: 1,
    });

    const pStyle = parseParagraphStyle({
      align: "center",
      lineHeight: "24pt",
      spacingAfter: 8,
      fontSize: "32pt",
      fontFamily: "minor",
      list: { type: "bullet", level: 1 },
    });
    assert.equal(pStyle.fontSize, "32pt");
    assert.equal(pStyle.fontFamily, "minor");

    const rStyle = parseRunStyle({
      fontSize: 32,
      fontFamily: "major",
      letterSpacing: 1.5,
      textDecoration: "underline",
      textTransform: "uppercase",
      lang: "en-US",
    });
    assert.equal(rStyle.fontSize, 32);
    assert.equal(rStyle.fontFamily, "major");

    assert.deepEqual(parseTextInsets({ padding: [4, 8] }), {
      top: 4,
      right: 8,
      bottom: 4,
      left: 8,
    });
    assert.deepEqual(parseTextInsets({ padding: 10, paddingLeft: 20 }), {
      top: 10,
      right: 10,
      bottom: 10,
      left: 20,
    });

    assert.equal(
      paragraphStyleSchema.safeParse({ verticalAlign: "top" }).success,
      false,
    );
    assert.equal(runStyleSchema.safeParse({ align: "left" }).success, false);
  });

  it("maps Zod schema violations onto exact CapabilityError codes in both schemas and textCapability", async () => {
    assert.throws(
      () => parseParagraphStyle({ list: { type: "bullet", character: "ab" } }),
      (error: unknown) =>
        error instanceof Error &&
        "code" in error &&
        error.code === "invalid-list",
    );
    assert.throws(
      () => parseParagraphStyle({ list: { type: "number", startAt: 0 } }),
      (error: unknown) =>
        error instanceof Error &&
        "code" in error &&
        error.code === "invalid-list",
    );
    assert.throws(
      () => parseRunStyle({ align: "center" }),
      (error: unknown) =>
        error instanceof Error &&
        "code" in error &&
        error.code === "unknown-style-property",
    );
    assert.throws(
      () => parseTextInsets({ padding: -2 }),
      (error: unknown) =>
        error instanceof Error &&
        "code" in error &&
        error.code === "invalid-text-insets",
    );

    const resolve = createResolver([slideBackgroundCapability, textCapability]);
    const badTextDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "bad",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 200,
                height: 40,
              },
            },
            createElement(
              P,
              null,
              createElement(
                Span,
                // @ts-expect-error -- align is invalid on Span.style
                { style: { align: "center" } },
                "bad span",
              ),
            ),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(badTextDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "unknown-prop"),
    );

    const badSlideNumDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "bad-num",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 200,
                height: 40,
              },
            },
            createElement(SlideNumber, null, "child"),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(badSlideNumDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-prop"),
    );

    assert.throws(
      // @ts-expect-error -- substitute must be boolean
      () => font("Jost*", [], { substitute: "yes" }),
      TypeError,
    );
  });
});
