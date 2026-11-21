import assert from "node:assert/strict";

import { P, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import { build } from "../test-helpers.ts";

function deck(text: ReturnType<typeof createElement>) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "cover" }, text),
  );
}

describe("DrawingML text field ranges", () => {
  it("rejects invalid values according to font size serialization ranges and accepts boundaries", async () => {
    for (const [fontSize, code] of [
      [0.1, "invalid-font-size"],
      ["6000pt", "invalid-font-size"],
    ] as const) {
      await assert.rejects(
        build(
          deck(
            createElement(Text, {
              x: 0,
              y: 0,
              width: 120,
              height: 60,
              style: { fontSize },
              children: "字体大小边界",
            }),
          ),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diagnostic) => diagnostic.code === code),
      );
    }

    for (const [fontSize, points] of [
      ["1pt", 1],
      ["4000pt", 4000],
    ] as const) {
      const result = await build(
        deck(
          createElement(Text, {
            x: 0,
            y: 0,
            width: 120,
            height: 60,
            style: { fontSize },
            children: "合法字号边界",
          }),
        ),
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(
        file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        points,
      );
    }
  });

  it("rejects out-of-range values according to spacing serialization ranges and preserves valid unit conversions", async () => {
    for (const style of [
      { lineHeight: "1000000px" },
      { lineHeight: "2112.01px" },
      { spacingAfter: 2112.01 },
    ] as const) {
      await assert.rejects(
        build(
          deck(
            createElement(
              Text,
              {
                x: 0,
                y: 0,
                width: 120,
                height: 60,
              },
              createElement(P, { style }, "间距边界"),
            ),
          ),
        ),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (diagnostic) =>
              diagnostic.code === "invalid-text-length" ||
              diagnostic.code === "invalid-paragraph-spacing",
          ),
      );
    }

    const result = await build(
      deck(
        createElement(
          Text,
          {
            x: 0,
            y: 0,
            width: 120,
            height: 120,
          },
          createElement(
            P,
            { style: { lineHeight: "2112px", spacingAfter: 2112 } },
            "合法最大间距",
          ),
          createElement(P, { style: { lineHeight: "1pt" } }, "一磅行距"),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects[0]?.text?.paragraphs.map(
        ({ lineSpacingPoints, spacingAfterPoints }) => ({
          lineSpacingPoints,
          spacingAfterPoints,
        }),
      ),
      [
        { lineSpacingPoints: 1584, spacingAfterPoints: 1584 },
        { lineSpacingPoints: 1, spacingAfterPoints: undefined },
      ],
    );
  });

  it("reports actually out-of-range font sizes and paragraph spacings during standalone file reading", async () => {
    const result = await build(
      deck(
        createElement(
          Text,
          {
            x: 0,
            y: 0,
            width: 120,
            height: 60,
            style: { lineHeight: "24px" },
          },
          createElement(P, { style: { spacingAfter: 20 } }, "原生文件检查"),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.match(slide, /sz="1800"/u);
    assert.match(slide, /<a:spcAft><a:spcPts val="1500"\/>/u);
    parts["ppt/slides/slide1.xml"] = strToU8(
      slide
        .replace('sz="1800"', 'sz="99"')
        .replace('val="1800"', 'val="-1"')
        .replace('val="1500"', 'val="158401"'),
    );

    const file = await validatePptx(zipSync(parts));
    assert.deepEqual(
      file.errors
        .filter(({ code }) =>
          ["invalid-font-size", "invalid-text-spacing"].includes(code),
        )
        .map(({ code }) => code)
        .sort(),
      ["invalid-font-size", "invalid-text-spacing", "invalid-text-spacing"],
    );
  });
});
