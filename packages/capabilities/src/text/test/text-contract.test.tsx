import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import {
  deck,
  font as impFont,
  paragraph,
  slide,
  slideNumber,
  text,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { textCapability } from "@pptxdk/capabilities/text";
import { P, SlideNumber, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, expandDocument, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);
const impJostFont = impFont("Jost*", [{ weight: 400, data: jostBook }]);

describe("textCapability resolved contract", () => {
  it("resolves text metrics, rotation, and overset textMetrics and paints glyph outlines without calling emit", async () => {
    const resolve = createResolver([slideBackgroundCapability, textCapability]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "wrapped",
              name: "WrappedTitle",
              style: {
                position: "absolute",
                left: 20,
                top: 30,
                width: 110,
                height: 40,
                fontFamily: jostFont,
                fontSize: 18,
                color: "#1E293B",
                rotate: 15,
              },
            },
            createElement(
              P,
              null,
              "Page ",
              createElement(SlideNumber, null),
              " wraps across multiple lines inside fixed box",
            ),
          ),
        ),
      ),
    );

    const resolution = resolve(doc);
    const locator = resolution.slide("cover").get("wrapped");
    assert.equal(locator.kind, "text");
    assert.equal(
      locator.text,
      "Page 1 wraps across multiple lines inside fixed box",
    );
    assert.deepEqual(locator.box, { x: 20, y: 30, width: 110, height: 40 });
    assert.equal(locator.textMetrics?.status, "measured");
    assert.equal(locator.textMetrics.fits, false);
    assert.ok(
      locator.textMetrics.content.height > locator.textMetrics.available.height,
    );
    assert.ok(locator.textMetrics.content.lineCount > 1);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    const page = painted.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg?.includes("<path"));
    assert.ok(page.svg?.includes("rotate(15"));

    const paintedNode = page.nodes.find((node) => node.authorId === "wrapped");
    assert.ok(paintedNode !== undefined);
    assert.ok((paintedNode.bounds?.width ?? 0) > 110);
    assert.ok((paintedNode.bounds?.height ?? 0) > 40);
  });

  it("defers font embedding and missing-glyph embedding checks to emit using resolve-time resource claims", async () => {
    const resolve = createResolver([slideBackgroundCapability, textCapability]);
    const validDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "title",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 120,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 18,
              },
            },
            createElement(
              P,
              null,
              "Slide ",
              createElement(SlideNumber, null),
              " long text that wraps softly",
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "second" },
          createElement(
            Text,
            {
              id: "page-two",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 120,
                height: 40,
                fontFamily: jostFont,
                fontSize: 18,
              },
            },
            createElement(P, null, "Page ", createElement(SlideNumber, null)),
          ),
        ),
      ),
    );

    // Resolve without embedFonts; emit with subset embedding must embed .fntdata from resolve claims.
    const validResolution = resolve(validDoc);
    const emittedSubset = emit(validResolution, { embedFonts: "subset" });
    const validation = await validatePptx(emittedSubset.pptx);
    assert.deepEqual(validation.errors, []);
    const zip = unzipSync(emittedSubset.pptx);
    assert.ok(
      Object.keys(zip).some(
        (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
      ),
    );
    const slide1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
    const slide2Xml = strFromU8(zip["ppt/slides/slide2.xml"]!);
    assert.ok(slide1Xml.includes('type="slidenum"'));
    assert.ok(slide2Xml.includes('type="slidenum"'));
    assert.equal(slide1Xml.includes("<a:br"), false);

    // Paint and non-embedded emit also work independently from the same resolution.
    const emittedPlain = emit(validResolution);
    assert.equal(
      Object.keys(unzipSync(emittedPlain.pptx)).some((entry) =>
        entry.startsWith("ppt/fonts/"),
      ),
      false,
    );
    const painted = paint(validResolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.deepEqual(painted.pages[1]?.errors, []);

    // Fixed-box missing glyph: resolve succeeds with unmeasured textMetrics & FontUsage.missingGlyphs; paint reports preview-missing-glyph; emit with embedFonts fails.
    const missingGlyphDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "cjk",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 120,
                height: 40,
                fontFamily: jostFont,
                fontSize: 18,
              },
            },
            "中文",
          ),
        ),
      ),
    );
    const missingResolution = resolve(missingGlyphDoc);
    const cjkMetrics = missingResolution.slide("cover").get("cjk").textMetrics;
    assert.equal(cjkMetrics?.status, "unmeasured");
    assert.equal(cjkMetrics.reason, "missing-glyphs");
    assert.deepEqual(cjkMetrics.fonts[0]?.missingGlyphs, ["中", "文"]);
    assert.deepEqual(missingResolution.report.fonts[0]?.missingGlyphs, [
      "中",
      "文",
    ]);
    const missingPainted = paint(missingResolution);
    assert.equal(missingPainted.pages[0]?.svg, undefined);
    assert.equal(missingPainted.pages[0]?.png, undefined);
    assert.equal(
      missingPainted.pages[0]?.errors[0]?.code,
      "preview-missing-glyph",
    );
    assert.throws(
      () => emit(missingResolution, { embedFonts: "full" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );

    // Content-sized missing glyph throws BuildError(font-missing-glyph) directly during resolve.
    const autoHeightMissingDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "auto-missing" },
          createElement(
            Text,
            {
              id: "auto-cjk",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 120,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 18,
              },
            },
            "中文",
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(autoHeightMissingDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );
  });

  it("validates XML 1.0 characters and text shadows during resolve before emit", () => {
    const resolve = createResolver([slideBackgroundCapability, textCapability]);

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Valid content", {
                id: "bad-name",
                name: "Bad\u0000Name",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                },
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-prop" &&
            path === "deck/slide[cover]/text[bad-name]",
        ),
    );

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Bad\u0000Text", {
                id: "bad-text",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                },
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-prop"),
    );

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Hello", {
                id: "bad-font",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                  fontFamily: "Bad\u0000Font",
                },
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-prop"),
    );

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Hello", {
                id: "bad-shadow",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                  textShadow: { color: "#000000", blur: -1 },
                },
              }),
            ),
          ),
        ),
      (error: unknown) => error instanceof BuildError,
    );
  });

  it("keeps resolve closures isolated from subsequent editable document mutations and emit failures", async () => {
    const resolve = createResolver([slideBackgroundCapability, textCapability]);
    const textNode = text([paragraph(["Snapshot ", slideNumber()])], {
      id: "mutable",
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 180,
        height: "auto",
        fontFamily: impJostFont,
        fontSize: 18,
      },
    });
    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(textNode),
    );

    const resolution = resolve(doc);
    // Mutate the editable document after resolve.
    textNode.content("Mutated text");

    const painted = paint(resolution);
    const emitted = emit(resolution);
    assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    assert.equal(resolution.slide("cover").get("mutable").text, "Snapshot 1");
    const slideXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes("Snapshot "));
    assert.equal(slideXml.includes("Mutated text"), false);
    assert.ok(painted.pages[0]?.png !== undefined);

    // Fixed-height text without font bytes resolves and emits cleanly, while report.fonts & textMetrics reflect missing resource and paint reports preview-missing-font.
    const noFontDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        text("System fallback", {
          id: "no-font",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 160,
            height: 50,
            fontFamily: "MissingFamily",
            fontSize: 16,
          },
        }),
      ),
    );
    const noFontResolution = resolve(noFontDoc);
    assert.ok(
      noFontResolution.report.fonts.some(
        (usage) =>
          usage.family === "MissingFamily" && usage.hasResource === false,
      ),
    );
    const noFontMetrics = noFontResolution
      .slide("cover")
      .get("no-font").textMetrics;
    assert.equal(noFontMetrics?.status, "unmeasured");
    assert.equal(noFontMetrics.reason, "missing-font-resource");
    const noFontPaint = paint(noFontResolution);
    assert.equal(noFontPaint.pages[0]?.svg, undefined);
    assert.equal(noFontPaint.pages[0]?.png, undefined);
    assert.equal(noFontPaint.pages[0]?.errors[0]?.code, "preview-missing-font");
    const noFontEmit = emit(noFontResolution);
    assert.deepEqual((await validatePptx(noFontEmit.pptx)).errors, []);
  });

  it("distinguishes unsupported host declarations from invalid non-declaration text children", async () => {
    const resolve = createResolver([slideBackgroundCapability, textCapability]);

    const shapeInTextDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "bad-text",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 200,
                height: 60,
              },
            },
            createElement(Shape, {
              style: {
                position: "absolute",
                left: 0,
                top: 0,
                width: 40,
                height: 20,
              },
            }),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(shapeInTextDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-structure"),
    );

    const shapeInParagraphDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "bad-p",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 200,
                height: 60,
              },
            },
            createElement(
              P,
              null,
              createElement(Shape, {
                style: {
                  position: "absolute",
                  left: 0,
                  top: 0,
                  width: 40,
                  height: 20,
                },
              }),
            ),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(shapeInParagraphDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-structure"),
    );

    const malformedChild = new CommonNode("bogus", "", {});
    const badText = text("", {
      id: "malformed",
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 200,
        height: 60,
      },
    });
    badText.append(malformedChild);
    const malformedDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(badText),
    );
    assert.throws(
      () => resolve(malformedDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-structure"),
    );
  });
});
