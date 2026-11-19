import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { expandDocument } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { paint } from "@react-pptx/core/preview";
import { emit } from "@react-pptx/emit";
import {
  BuildError,
  Deck,
  P,
  Slide,
  SlideNumber,
  Text,
  font,
  slideBackgroundCapability,
  textCapability,
} from "@react-pptx/react-pptx";
import {
  deck,
  font as impFont,
  paragraph,
  slide,
  slideNumber,
  text,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBook = readFileSync(
  new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);
const impJostFont = impFont("Jost*", [{ weight: 400, data: jostBook }]);

describe("textCapability resolved contract", () => {
  it("resolves text metrics, rotation, and overset warnings and paints glyph outlines without calling emit", async () => {
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
              x: 20,
              y: 30,
              width: 110,
              height: 40,
              style: {
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
    const slideReport = resolution.report.slides[0];
    assert.ok(slideReport !== undefined);
    const observed = slideReport.objects.find(
      (obj) => obj.authorId === "wrapped",
    );
    assert.ok(observed !== undefined);
    assert.equal(observed.kind, "text");
    assert.equal(observed.name, "WrappedTitle");
    assert.equal(
      observed.text,
      "Page 1 wraps across multiple lines inside fixed box",
    );
    assert.deepEqual(observed.box, { x: 20, y: 30, width: 110, height: 40 });
    assert.ok(
      resolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" &&
          diag.code === "overset" &&
          diag.path === "deck/slide[cover]/text[wrapped]",
      ),
    );

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    const page = painted.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg?.includes("<path"));
    assert.ok(page.svg?.includes("rotate(15"));

    const paintedNode = page.nodes.find((node) => node.authorId === "wrapped");
    assert.ok(paintedNode !== undefined);
    assert.ok((paintedNode.screenBox?.width ?? 0) > 110);
    assert.ok((paintedNode.screenBox?.height ?? 0) > 40);
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
              x: 20,
              y: 20,
              width: 120,
              height: "auto",
              style: { fontFamily: jostFont, fontSize: 18 },
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
              x: 20,
              y: 20,
              width: 120,
              height: 40,
              style: { fontFamily: jostFont, fontSize: 18 },
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
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.deepEqual(painted.pages[1]?.diagnostics, []);

    // Missing glyph: resolve without embedFonts succeeds with warning, emit with embedFonts fails.
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
              x: 20,
              y: 20,
              width: 120,
              height: 40,
              style: { fontFamily: jostFont, fontSize: 18 },
            },
            "中文",
          ),
        ),
      ),
    );
    const missingResolution = resolve(missingGlyphDoc);
    assert.ok(
      missingResolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" && diag.code === "font-missing-glyph",
      ),
    );
    assert.throws(
      () => emit(missingResolution, { embedFonts: "full" }),
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
                x: 10,
                y: 10,
                width: 100,
                height: 40,
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-xml-character" &&
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
                x: 10,
                y: 10,
                width: 100,
                height: 40,
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-xml-character"),
    );

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Hello", {
                id: "bad-font",
                x: 10,
                y: 10,
                width: 100,
                height: 40,
                style: { fontFamily: "Bad\u0000Font" },
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-xml-character"),
    );

    assert.throws(
      () =>
        resolve(
          deck({ width: 480, height: 270 }).append(
            slide({ id: "cover" }).append(
              text("Hello", {
                id: "bad-shadow",
                x: 10,
                y: 10,
                width: 100,
                height: 40,
                style: {
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
      x: 20,
      y: 20,
      width: 180,
      height: "auto",
      style: { fontFamily: impJostFont, fontSize: 18 },
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
    assert.equal(
      resolution.query.getSlide("cover").get("mutable").observe().text,
      "Snapshot 1",
    );
    const slideXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes("Snapshot "));
    assert.equal(slideXml.includes("Mutated text"), false);
    assert.ok(painted.pages[0]?.png !== undefined);

    // Fixed-height text without font bytes resolves and emits cleanly, while paint reports error without affecting emit.
    const noFontDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "cover" }).append(
        text("System fallback", {
          id: "no-font",
          x: 20,
          y: 20,
          width: 160,
          height: 50,
          style: { fontFamily: "MissingFamily", fontSize: 16 },
        }),
      ),
    );
    const noFontResolution = resolve(noFontDoc);
    assert.ok(
      noFontResolution.report.diagnostics.some(
        (diag) => diag.code === "font-measurement-unavailable",
      ),
    );
    const noFontPaint = paint(noFontResolution);
    assert.equal(noFontPaint.pages[0]?.png, undefined);
    assert.ok((noFontPaint.pages[0]?.diagnostics.length ?? 0) > 0);
    const noFontEmit = emit(noFontResolution);
    assert.deepEqual((await validatePptx(noFontEmit.pptx)).errors, []);
  });
});
