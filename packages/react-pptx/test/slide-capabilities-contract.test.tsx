import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { expandDocument } from "@react-pptx/core";
import { createResolver, emit } from "@react-pptx/core/build";
import { paint } from "@react-pptx/core/preview";
import {
  BuildError,
  Deck,
  Group,
  Image,
  MotionScope,
  Notes,
  P,
  Shape,
  Slide,
  SlideNumber,
  Step,
  Text,
  after,
  createMorphScope,
  createMotionScopeHandle,
  fade,
  font,
  groupCapability,
  imageCapability,
  linearGradient,
  motionCapability,
  notesCapability,
  observeMotion,
  shapeCapability,
  slideBackgroundCapability,
  spin,
  targets,
  textCapability,
  transitionCapability,
  transitions,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it, vi } from "vitest";

const jostBook = readFileSync(
  new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("slide capabilities resolved contract", () => {
  it("migrates slideBackgroundCapability to ResolvedBoundSlideCapability with paint-only solid/default rendering, preview error on image/gradient, and emit-only resource registration", async () => {
    const bindSpy = vi.spyOn(slideBackgroundCapability, "bindSlide");
    try {
      const resolve = createResolver([slideBackgroundCapability]);
      const doc = await expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(Slide, { id: "default-bg" }),
          createElement(Slide, { id: "solid-bg", background: "#1E293B" }),
          createElement(Slide, {
            id: "gradient-bg",
            background: linearGradient(
              90,
              { color: "#0F172A", position: 0 },
              { color: "#38BDF8", position: 1 },
            ),
          }),
          createElement(Slide, {
            id: "image-bg",
            background: { image: png1x1 },
          }),
        ),
      );

      const resolution = resolve(doc);
      assert.equal(bindSpy.mock.calls.length, 4);
      for (const callResult of bindSpy.mock.results) {
        const bound = callResult.value as Record<string, unknown>;
        assert.equal(typeof bound["resolveSlide"], "function");
        assert.equal("compileSlide" in bound, false);
        assert.equal("previewSlide" in bound, false);
      }

      // Paint directly from Resolution without running emit().
      const painted = paint(resolution);
      assert.equal(painted.pages.length, 4);

      assert.deepEqual(painted.pages[0]?.diagnostics, []);
      assert.ok(painted.pages[0]?.png !== undefined);
      assert.match(painted.pages[0]?.svg ?? "", /fill="white"/u);

      assert.deepEqual(painted.pages[1]?.diagnostics, []);
      assert.ok(painted.pages[1]?.png !== undefined);
      assert.match(painted.pages[1]?.svg ?? "", /fill="#1E293B"/u);

      // Gradient and image backgrounds report preview error diagnostics and omit PNG,
      // while emit() on the same Resolution succeeds.
      for (const pageIndex of [2, 3]) {
        const page = painted.pages[pageIndex]!;
        assert.equal(page.png, undefined);
        assert.ok(
          page.diagnostics.some(
            (diag) =>
              diag.severity === "error" &&
              diag.message.includes("image or gradient backgrounds"),
          ),
        );
      }

      const emitted1 = emit(resolution);
      const emitted2 = emit(resolution);
      assert.deepEqual((await validatePptx(emitted1.pptx)).errors, []);
      assert.deepEqual((await validatePptx(emitted2.pptx)).errors, []);

      const entries1 = unzipSync(emitted1.pptx);
      const entries2 = unzipSync(emitted2.pptx);
      const media1 = Object.keys(entries1).filter((name) =>
        name.startsWith("ppt/media/"),
      );
      const media2 = Object.keys(entries2).filter((name) =>
        name.startsWith("ppt/media/"),
      );
      assert.equal(media1.length, 1);
      assert.deepEqual(media2, media1);

      const slide4Xml = strFromU8(entries1["ppt/slides/slide4.xml"]!);
      assert.match(slide4Xml, /<p:bg><p:bgPr><a:blipFill>/u);
      const slide4Rels = strFromU8(
        entries1["ppt/slides/_rels/slide4.xml.rels"]!,
      );
      assert.match(slide4Rels, /relationships\/image/u);

      // Invalid background fails during resolve.
      const invalidDoc = await expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(Slide, {
            id: "bad-bg",
            background: { image: png1x1, extra: true } as never,
          }),
        ),
      );
      assert.throws(
        () => resolve(invalidDoc),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (diag) =>
              diag.code === "invalid-slide-background" &&
              diag.path === "deck/slide[bad-bg]/background",
          ),
      );
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("migrates transitionCapability to ResolvedBoundSlideCapability with Morph continuity, paint without diagnostics, and emit target checks", async () => {
    const bindSpy = vi.spyOn(transitionCapability, "bindSlide");
    try {
      const resolve = createResolver([
        slideBackgroundCapability,
        shapeCapability,
        imageCapability,
        transitionCapability,
      ]);
      const morph = createMorphScope();
      const hero = morph.handle("hero");

      const doc = await expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            {
              id: "s1",
              background: "#FFFFFF",
              advanceAfter: 1200,
              transition: transitions.fade({ duration: 500 }),
            },
            createElement(Shape, {
              id: "hero-a",
              morphKey: hero,
              x: 20,
              y: 20,
              width: 80,
              height: 40,
              fill: "#2563EB",
            }),
          ),
          createElement(
            Slide,
            {
              id: "s2",
              background: "#FFFFFF",
              transition: transitions.morph({
                duration: 900,
                fallback: transitions.fade(),
              }),
            },
            createElement(Shape, {
              id: "hero-b",
              morphKey: hero,
              x: 120,
              y: 40,
              width: 100,
              height: 60,
              fill: "#1D4ED8",
            }),
          ),
        ),
      );

      const resolution = resolve(doc);
      assert.equal(bindSpy.mock.calls.length, 2);
      const firstBound = bindSpy.mock.results[0]?.value as Record<
        string,
        unknown
      >;
      const secondBound = bindSpy.mock.results[1]?.value as Record<
        string,
        unknown
      >;
      assert.equal(typeof firstBound["resolveSlide"], "function");
      assert.equal("compileSlide" in firstBound, false);
      assert.equal("previewSlide" in firstBound, false);
      assert.deepEqual(secondBound["continuity"], { matchSourceKinds: true });

      assert.deepEqual(resolution.report.extensions, ["p14", "p159"]);
      assert.deepEqual(
        resolution.report.slides[1]?.capabilities?.["transition"],
        {
          kind: "morph",
          by: "object",
          speed: "slow",
          durationMs: 900,
          fallback: { kind: "fade", speed: "slow" },
        },
      );

      // Paint succeeds on the same Resolution even when baselineOnly target rejects emit.
      const painted = paint(resolution);
      assert.equal(painted.pages.length, 2);
      assert.deepEqual(painted.pages[0]?.diagnostics, []);
      assert.deepEqual(painted.pages[1]?.diagnostics, []);
      assert.ok(painted.pages[1]?.png !== undefined);

      assert.throws(
        () => emit(resolution, { target: targets.baselineOnly }),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((diag) => diag.code === "forbidden-mechanism"),
      );

      const emitted = emit(resolution, { target: targets.transitional });
      const file = await validatePptx(emitted.pptx, {
        allowedExtensions: ["p14", "p159"],
      });
      assert.deepEqual(file.errors, []);
      assert.equal(file.slides[0]?.advanceAfterMs, 1200);
      assert.equal(file.slides[1]?.transition?.kind, "morph");

      // Cross-slide Morph kind mismatch still fails during resolve.
      const mismatchDoc = await expandDocument(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "m1" },
            createElement(Shape, {
              id: "shape-src",
              morphKey: hero,
              x: 20,
              y: 20,
              width: 80,
              height: 40,
            }),
          ),
          createElement(
            Slide,
            {
              id: "m2",
              transition: transitions.morph({ fallback: transitions.fade() }),
            },
            createElement(Image, {
              id: "img-dst",
              src: png1x1,
              morphKey: hero,
              x: 20,
              y: 20,
              width: 80,
              height: 40,
            }),
          ),
        ),
      );
      assert.throws(
        () => resolve(mismatchDoc),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some(
            (diag) => diag.code === "morph-object-kind-mismatch",
          ),
      );
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("migrates motionCapability to ResolvedBoundSlideCapability with resolve-time WeakMap query binding, nativeProjections, paint without diagnostics, and emit identity preservation", async () => {
    const boundHandles: Record<string, unknown>[] = [];
    const trackedMotion: typeof motionCapability = {
      ...motionCapability,
      bindSlide(input, context) {
        const handle = motionCapability.bindSlide(input, context);
        boundHandles.push(handle);
        return handle;
      },
    };

    const resolve = createResolver([
      slideBackgroundCapability,
      shapeCapability,
      groupCapability,
      textCapability,
      trackedMotion,
    ]);
    const intro = createMotionScopeHandle();
    const captionScope = createMotionScopeHandle();
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "motion-slide", background: "#FFFFFF" },
          createElement(
            MotionScope,
            { scope: intro, mode: "sequence" },
            createElement(
              Group,
              {
                id: "spin-group",
                emphasis: spin({ duration: 500 }),
              },
              createElement(Shape, {
                id: "group-member",
                x: 20,
                y: 20,
                width: 60,
                height: 40,
                fill: "#2563EB",
                enter: fade({ duration: 200 }),
              }),
            ),
          ),
          createElement(
            MotionScope,
            { scope: captionScope, startAt: after(intro.end, 100) },
            createElement(
              Text,
              {
                id: "caption",
                x: 20,
                y: 120,
                width: 200,
                height: 60,
                style: { fontFamily: jostFont, fontSize: 16 },
                enter: fade({ duration: 300 }),
              },
              createElement(P, null, "Delayed caption"),
            ),
          ),
          createElement(
            Step,
            null,
            createElement(Shape, {
              id: "click-box",
              x: 260,
              y: 40,
              width: 80,
              height: 60,
              fill: "#10B981",
              enter: fade({ duration: 250 }),
            }),
          ),
        ),
      ),
    );

    const resolution = resolve(doc);
    assert.equal(boundHandles.length, 1);
    assert.equal(typeof boundHandles[0]?.["resolveSlide"], "function");
    assert.equal("compileSlide" in boundHandles[0]!, false);
    assert.equal("previewSlide" in boundHandles[0]!, false);

    // observeMotion works immediately on Resolution before emit().
    const resolvedSlide = resolution.report.slides[0]!;
    const resolvedMotion = observeMotion(resolvedSlide);
    assert.ok(resolvedMotion !== undefined);
    assert.equal(resolvedMotion.steps.length, 2);
    assert.equal(resolvedMotion.steps[0]?.trigger, "auto");
    assert.equal(resolvedMotion.steps[1]?.trigger, "click");

    const motionObs = resolvedSlide.capabilities?.["motion"] as {
      readonly nativeProjections?: readonly {
        readonly logicalTargetId: number;
        readonly nativeEffects: readonly { readonly timingId: number }[];
      }[];
    };
    assert.equal(motionObs.nativeProjections?.length, 1);

    // Paint succeeds without diagnostics from the same Resolution.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);

    // Emit preserves slide observation identity so observeMotion works identically on BuildReport.
    const emitted = emit(resolution);
    const emittedMotion = observeMotion(emitted.report.slides[0]!);
    assert.equal(emittedMotion, resolvedMotion);
    assert.equal(
      emitted.report.slides[0]?.capabilities?.["motion"],
      resolvedSlide.capabilities?.["motion"],
    );

    const file = await validatePptx(emitted.pptx);
    assert.deepEqual(file.errors, []);
    const fileTimingIds = new Set(
      file.slides[0]?.animations
        ?.flatMap((step) => step.effects)
        .map((effect) => effect.timingId) ?? [],
    );
    for (const nativeEffect of motionObs.nativeProjections[0]!.nativeEffects) {
      assert.ok(fileTimingIds.has(nativeEffect.timingId));
    }

    // Invalid motion timing or missing target fails during resolve.
    const missingCueHandle = createMotionScopeHandle();
    const badScope = createMotionScopeHandle();
    const invalidDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "bad-motion" },
          createElement(
            MotionScope,
            { scope: badScope, startAt: after(missingCueHandle.end, 100) },
            createElement(Shape, {
              id: "orphan",
              x: 10,
              y: 10,
              width: 50,
              height: 50,
              enter: fade({ duration: 200 }),
            }),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolve(invalidDoc),
      (error: unknown) =>
        error instanceof BuildError && error.diagnostics.length > 0,
    );
  });

  it("migrates notesCapability to ResolvedBoundSlideCapability with resolve-time claims, paint without notes leakage, and emit-only shared notes parts and font embedding", async () => {
    const boundHandles: Record<string, unknown>[] = [];
    const trackedNotes: typeof notesCapability = {
      ...notesCapability,
      bindSlide(input, context) {
        const handle = notesCapability.bindSlide(input, context);
        boundHandles.push(handle);
        return handle;
      },
    };

    const resolve = createResolver([slideBackgroundCapability, trackedNotes]);

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
    assert.equal(boundHandles.length, 3);
    for (const bound of boundHandles) {
      assert.equal(typeof bound["resolveSlide"], "function");
      assert.equal("compileSlide" in bound, false);
      assert.equal("previewSlide" in bound, false);
    }
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
