import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@react-pptx/capabilities";
import { imageCapability } from "@react-pptx/capabilities/image";
import { Image } from "@react-pptx/capabilities/image-elements";
import { shapeCapability } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import {
  transitionCapability,
  transitions,
} from "@react-pptx/capabilities/transition";
import { BuildError, createMorphScope, targets } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  runTransitionRejectionScenarios,
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
      const runs = await runCapabilityScenario(transitionPositiveScenario, {
        outputDir: outDir,
      });
      assert.equal(runs.length, 2);
      assert.deepEqual(
        runs.map((r) => r.author),
        ["jsx", "document"],
      );
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
  });
});
