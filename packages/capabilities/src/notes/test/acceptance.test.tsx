import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import { boxCapability } from "@pptxdk/capabilities/box";
import {
  box,
  deck,
  font,
  notes,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { notesCapability } from "@pptxdk/capabilities/notes";
import { Notes } from "@pptxdk/capabilities/notes-elements";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import {
  BuildError,
  type CapabilityModule,
  type Diagnostic,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import type { CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement, type ComponentType, type ReactNode } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import {
  notesPositiveScenario,
  notesPreviewIsolationScenario,
} from "./scenarios.ts";

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);

function previewDocWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      resolution,
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

function previewJsxWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (element: ReactNode) => {
    const resolution = resolveWith(await expandDocument(element));
    return {
      resolution,
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Notes capability acceptance & bound slide contract", () => {
  it("runs Notes dual-author positive and preview isolation scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-notes-accept-"));
    try {
      for (const scenario of [
        notesPositiveScenario,
        notesPreviewIsolationScenario,
      ]) {
        const { runs } = await runCapabilityScenario(scenario, {
          outputDir: outDir,
        });
        assert.equal(runs.length, 2);
        assert.deepEqual(
          runs.map((r) => r.author),
          ["jsx", "document"],
        );
      }
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 15_000);

  it("binds notesCapability via BoundSlideCapabilityModule, trims textCapability cleanly, and enforces rejections", async () => {
    assert.equal("bindSlide" in notesCapability, true);
    assert.equal("measureLayout" in notesCapability, false);

    // Notes-only slide previews cleanly as an empty canvas without preview errors or leaked notes text.
    const notesOnlyPreview = previewDocWith([
      notesCapability,
      slideBackgroundCapability,
    ]);
    const notesOnlyResult = await notesOnlyPreview(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "notes-only" }).append(notes("Speaker only")),
      ),
    );
    assert.deepEqual(notesOnlyResult.pages[0]?.diagnostics, []);
    assert.ok(notesOnlyResult.pages[0]?.png !== undefined);
    const notesOnlyPng = PNG.sync.read(
      Buffer.from(notesOnlyResult.pages[0].png),
    );
    for (let i = 0; i < notesOnlyPng.data.length; i += 4) {
      assert.deepEqual(
        [...notesOnlyPng.data.subarray(i, i + 4)],
        [255, 255, 255, 255],
      );
    }

    const trimmedModules = [
      shapeCapability,
      notesCapability,
      slideBackgroundCapability,
    ] as const;
    const previewDocTrimmed = previewDocWith(trimmedModules);
    const previewJsxTrimmed = previewJsxWith(trimmedModules);

    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "trimmed-text" }).append(
        shape({
          id: "card",
          x: 40,
          y: 40,
          width: 200,
          height: 100,
          fill: "#1D4ED8",
        }),
        notes(
          [
            paragraph([
              span("Trimmed text check ", { style: { fontWeight: "bold" } }),
              slideNumber(),
            ]),
          ],
          { id: "speaker" },
        ),
      ),
    );
    const docResult = await previewDocTrimmed(doc);
    assert.deepEqual(docResult.pages[0]?.diagnostics, []);
    assert.ok(docResult.pages[0]?.png !== undefined);
    assert.deepEqual(docResult.report.slides[0]?.capabilities, {
      notes: { text: "Trimmed text check 1", paragraphCount: 1 },
    });

    const jsxResult = await previewJsxTrimmed(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "trimmed-text" },
          createElement(Shape, {
            id: "card",
            x: 40,
            y: 40,
            width: 200,
            height: 100,
            fill: "#1D4ED8",
          }),
          createElement(
            Notes,
            { id: "speaker" },
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontWeight: "bold" } },
                "Trimmed text check ",
              ),
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    assert.deepEqual(jsxResult.report.slides[0]?.capabilities, {
      notes: { text: "Trimmed text check 1", paragraphCount: 1 },
    });

    async function expectBuildErrorCode(
      run: () => Promise<unknown>,
    ): Promise<readonly Diagnostic[]> {
      try {
        await run();
      } catch (error) {
        assert.ok(error instanceof BuildError);
        return error.diagnostics;
      }
      assert.fail("Expected BuildError");
    }

    const withBoxModules = [...trimmedModules, boxCapability] as const;
    const previewDocWithBox = previewDocWith(withBoxModules);
    const previewJsxWithBox = previewJsxWith(withBoxModules);
    const previewWithoutNotes = previewDocWith([
      shapeCapability,
      boxCapability,
    ]);
    const previewJsxWithoutNotes = previewJsxWith([
      shapeCapability,
      boxCapability,
    ]);

    const dupHiddenDiags = await expectBuildErrorCode(() =>
      previewDocWithBox(
        deck().append(
          slide({ id: "dup" }).append(
            box({ style: { display: "none" } }).append(notes("hidden")),
            notes("visible"),
          ),
        ),
      ),
    );
    assert.deepEqual(
      dupHiddenDiags.map(({ code }) => code),
      ["duplicate-notes"],
    );

    const dupJsxDiags = await expectBuildErrorCode(() =>
      previewJsxWithBox(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "dup-jsx" },
            createElement(Notes, null, createElement(P, null, "one")),
            createElement(Notes, null, createElement(P, null, "two")),
          ),
        ),
      ),
    );
    assert.deepEqual(
      dupJsxDiags.map(({ code }) => code),
      ["duplicate-notes"],
    );

    const badChildDiags = await expectBuildErrorCode(() =>
      previewDocTrimmed(
        deck().append(
          slide({ id: "bad-child" }).append(notes(text("illegal"))),
        ),
      ),
    );
    assert.equal(badChildDiags[0]?.code, "unsupported-text-child");

    const unknownPropDiags = await expectBuildErrorCode(() =>
      previewJsxTrimmed(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "unknown-prop" },
            createElement(
              Notes as unknown as ComponentType<Record<string, unknown>>,
              { layout: { x: 1 } },
              createElement(P, null, "bad"),
            ),
          ),
        ),
      ),
    );
    assert.equal(unknownPropDiags[0]?.code, "unknown-parameter");

    const missingHiddenDocDiags = await expectBuildErrorCode(() =>
      previewWithoutNotes(
        deck().append(
          slide({ id: "missing-hidden" }).append(
            box({ style: { display: "none" } }).append(notes("hidden")),
          ),
        ),
      ),
    );
    assert.deepEqual(
      missingHiddenDocDiags.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "missing-capability",
          path: "deck/slide[missing-hidden]/box[1]/notes[1]",
        },
      ],
    );

    const missingJsxDiags = await expectBuildErrorCode(() =>
      previewJsxWithoutNotes(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "missing-jsx" },
            createElement(Notes, null, createElement(P, null, "trimmed")),
          ),
        ),
      ),
    );
    assert.equal(missingJsxDiags[0]?.code, "missing-capability");

    const dupIdDiags = await expectBuildErrorCode(() =>
      previewDocTrimmed(
        deck().append(
          slide({ id: "dup-id" }).append(
            shape({ id: "same", x: 10, y: 10, width: 40, height: 20 }),
            notes("body", { id: "same" }),
          ),
        ),
      ),
    );
    assert.equal(dupIdDiags[0]?.code, "duplicate-author-id");
  });

  it("migrates notesCapability to BoundSlideCapability with resolve-time claims, paint without notes leakage, and emit-only shared notes parts and font embedding", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      notesCapability,
    ]);

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "n1", background: "#FFFFFF" },
          createElement(
            Notes,
            { style: { fontFamily: jostFont, fontSize: 12 } },
            createElement(
              P,
              null,
              "Speaker slide ",
              createElement(SlideNumber, null),
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "n2", background: "#FFFFFF" },
          createElement(
            Notes,
            { style: { fontFamily: jostFont, fontSize: 12 } },
            createElement(P, null, "Second note"),
          ),
        ),
        createElement(Slide, { id: "n3", background: "#FFFFFF" }),
      ),
    );

    const resolution = resolve(doc);
    assert.deepEqual(resolution.report.slides[0]?.capabilities?.["notes"], {
      text: "Speaker slide 1",
      paragraphCount: 1,
    });
    assert.deepEqual(resolution.report.slides[1]?.capabilities?.["notes"], {
      text: "Second note",
      paragraphCount: 1,
    });
    assert.equal(
      resolution.report.slides[2]?.capabilities?.["notes"],
      undefined,
    );

    // Paint succeeds cleanly without emitting notes parts or leaking notes text into canvas SVG.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 3);
    for (const page of painted.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
    }
    assert.equal(painted.pages[0]?.svg?.includes("Speaker slide"), false);

    // Emit registers shared notesMaster, notesTheme, per-slide notesSlides, field IDs, and embedded fonts.
    const emitted1 = emit(resolution, { embedFonts: "subset" });
    const emitted2 = emit(resolution, { embedFonts: "subset" });
    assert.deepEqual((await validatePptx(emitted1.pptx)).errors, []);
    assert.deepEqual((await validatePptx(emitted2.pptx)).errors, []);

    const entries1 = unzipSync(emitted1.pptx);
    const entries2 = unzipSync(emitted2.pptx);
    const notesMasterCount = Object.keys(entries1).filter((name) =>
      name.startsWith("ppt/notesMasters/notesMaster"),
    ).length;
    assert.equal(notesMasterCount, 1);
    assert.deepEqual(
      Object.keys(entries1).toSorted(),
      Object.keys(entries2).toSorted(),
    );

    const notesSlidePaths = Object.keys(entries1)
      .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(name))
      .toSorted();
    assert.equal(notesSlidePaths.length, 2);
    const notes1Xml = strFromU8(entries1[notesSlidePaths[0]!]!);
    assert.match(notes1Xml, /<a:fld id="\{[0-9A-F-]+\}" type="slidenum"/u);
    assert.ok(
      Object.keys(entries1).some((name) => name.startsWith("ppt/fonts/")),
    );

    // Notes with missing glyph resolves and paints, but fails only when emit requests font embedding.
    const missingGlyphDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "missing-glyph", background: "#FFFFFF" },
          createElement(
            Notes,
            { style: { fontFamily: jostFont } },
            createElement(P, null, "Notes 中文"),
          ),
        ),
      ),
    );
    const missingGlyphResolution = resolve(missingGlyphDoc);
    assert.deepEqual(paint(missingGlyphResolution).pages[0]?.diagnostics, []);
    assert.deepEqual(
      (await validatePptx(emit(missingGlyphResolution).pptx)).errors,
      [],
    );
    assert.throws(
      () => emit(missingGlyphResolution, { embedFonts: "subset" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "font-missing-glyph"),
    );

    // Duplicate Notes and invalid XML 1.0 characters fail during resolve.
    const dupDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "dup-notes" },
          createElement(Notes, null, createElement(P, null, "First")),
          createElement(Notes, null, createElement(P, null, "Second")),
        ),
      ),
    );
    assert.throws(
      () => resolve(dupDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "duplicate-notes"),
    );

    const badXmlDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "bad-xml-notes" },
          createElement(Notes, null, createElement(P, null, "Bad\u0000note")),
        ),
      ),
    );
    assert.throws(
      () => resolve(badXmlDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "invalid-xml-character"),
    );
  });
});
