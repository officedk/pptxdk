import assert from "node:assert/strict";

import { font } from "@pptxdk/capabilities/font";
import { slideBackgroundCapability } from "@pptxdk/capabilities/slide-background";
import { textCapability } from "@pptxdk/capabilities/text";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

const resolveText = createResolver([slideBackgroundCapability, textCapability]);

describe("Text & font public contract schemas", () => {
  it("resolves valid textList, paragraphStyle, runStyle, and insets across public JSX and resolver boundaries", async () => {
    const validDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "ok",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 240,
                height: 140,
                padding: 10,
                paddingLeft: 20,
              },
            },
            createElement(
              P,
              {
                style: {
                  align: "center",
                  lineHeight: "24pt",
                  spacingAfter: 8,
                  fontSize: "32pt",
                  fontFamily: "minor",
                  list: { type: "bullet", character: "•", level: 2 },
                },
              },
              createElement(
                Span,
                {
                  style: {
                    fontSize: 32,
                    fontFamily: "major",
                    letterSpacing: 1.5,
                    textDecoration: "underline",
                    textTransform: "uppercase",
                    lang: "en-US",
                  },
                },
                "Run",
              ),
            ),
            createElement(
              P,
              {
                style: {
                  list: {
                    type: "number",
                    style: "arabicPeriod",
                    startAt: 3,
                    level: 1,
                  },
                },
              },
              "Page ",
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );

    const resolution = resolveText(validDoc);
    assert.equal(resolution.slides.length, 1);
  });

  it("maps invalid list, run/paragraph style, insets, SlideNumber children, and font options onto typed errors", async () => {
    const invalidCases = [
      {
        label: "multi-char bullet",
        code: "invalid-style",
        el: createElement(
          P,
          { style: { list: { type: "bullet", character: "ab" } } },
          "bad",
        ),
      },
      {
        label: "zero number startAt",
        code: "invalid-style",
        el: createElement(
          P,
          { style: { list: { type: "number", startAt: 0 } } },
          "bad",
        ),
      },
      {
        label: "verticalAlign on Paragraph",
        code: "unknown-prop",
        el: createElement(
          P,
          { style: { verticalAlign: "top" } as never },
          "bad",
        ),
      },
      {
        label: "align on Span",
        code: "unknown-prop",
        el: createElement(
          P,
          null,
          createElement(
            Span,
            // @ts-expect-error -- align is invalid on Span.style
            { style: { align: "center" } },
            "bad span",
          ),
        ),
      },
      {
        label: "children on SlideNumber",
        code: "invalid-prop",
        el: createElement(SlideNumber, null, "child"),
      },
    ] as const;

    for (const item of invalidCases) {
      const doc = await expandDocument(
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
              item.el,
            ),
          ),
        ),
      );
      assert.throws(
        () => resolveText(doc),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === item.code),
        `${item.label} should fail with ${item.code}`,
      );
    }

    const badPaddingDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "bad-pad",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 200,
                height: 40,
                padding: -2,
              },
            },
            "bad padding",
          ),
        ),
      ),
    );
    assert.throws(
      () => resolveText(badPaddingDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-style"),
    );

    assert.throws(
      // @ts-expect-error -- substitute must be boolean
      () => font("Jost*", [], { substitute: "yes" }),
      TypeError,
    );
  });
});
