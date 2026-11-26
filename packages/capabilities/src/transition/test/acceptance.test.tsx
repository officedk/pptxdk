import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import { imageCapability } from "@pptxdk/capabilities/image";
import { Image } from "@pptxdk/capabilities/image-elements";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import {
  transitionCapability,
  transitions,
} from "@pptxdk/capabilities/transition";
import { BuildError, createMorphScope, targets } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import {
  runTransitionRejectionScenarios,
  transitionBaselineVariantsScenario,
  transitionPositiveScenario,
} from "./scenarios.ts";

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("Transition capability acceptance & bound slide contract", () => {
  it("runs Transition and Morph dual-author scenarios through shared runner and verifies rejections", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-trans-"));
    try {
      for (const scenario of [
        transitionPositiveScenario,
        transitionBaselineVariantsScenario,
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
      await runTransitionRejectionScenarios();
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 15_000);

  it("migrates transitionCapability to BoundSlideCapability with Morph continuity, paint without diagnostics, and emit target checks", async () => {
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
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 80,
              height: 40,
              background: "#2563EB",
            },
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
            style: {
              position: "absolute",
              left: 120,
              top: 40,
              width: 100,
              height: 60,
              background: "#1D4ED8",
            },
          }),
        ),
      ),
    );

    const resolution = resolve(doc);

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
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.deepEqual(painted.pages[1]?.errors, []);
    assert.ok(painted.pages[1]?.png !== undefined);

    assert.throws(
      () => emit(resolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "target-violation"),
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
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 80,
              height: 40,
            },
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
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 80,
              height: 40,
            },
          }),
        ),
      ),
    );
    assert.throws(
      () => resolve(mismatchDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-continuity" &&
            diag.reason === "kind-mismatch",
        ),
    );
  });
});
