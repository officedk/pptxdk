import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import { createMotionScopeHandle, font } from "@pptxdk/capabilities/document";
import { groupCapability } from "@pptxdk/capabilities/group";
import { Group } from "@pptxdk/capabilities/group-elements";
import {
  after,
  fade,
  motionCapability,
  motionSchedule,
  spin,
} from "@pptxdk/capabilities/motion";
import { MotionScope, Step } from "@pptxdk/capabilities/motion-elements";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { textCapability } from "@pptxdk/capabilities/text";
import { P, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError, createMorphScope } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { createElement, type ComponentType } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import {
  motionPositiveScenario,
  verifyMotionCollaborationAndRejections,
} from "./scenarios.ts";

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);

describe("Motion capability acceptance & bound slide contract", () => {
  it("runs Motion dual-author choreography scenario, static canvas preview, Group/Chart collaboration, and rejections", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-motion-accept-"));
    try {
      const { runs } = await runCapabilityScenario(motionPositiveScenario, {
        outputDir: outDir,
      });
      assert.equal(runs.length, 2);
      assert.deepEqual(
        runs.map((r) => r.author),
        ["jsx", "document"],
      );
      await verifyMotionCollaborationAndRejections();
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 15_000);

  it("migrates motionCapability to BoundSlideCapability with resolve-time WeakMap query binding, nativeProjections, paint without diagnostics, and emit identity preservation", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      shapeCapability,
      groupCapability,
      textCapability,
      motionCapability,
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
              { id: "spin-group", emphasis: spin({ duration: 500 }) },
              createElement(Shape, {
                id: "group-member",
                enter: fade({ duration: 200 }),
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 60,
                  height: 40,
                  background: "#2563EB",
                },
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
                enter: fade({ duration: 300 }),
                style: {
                  position: "absolute",
                  left: 20,
                  top: 120,
                  width: 200,
                  height: 60,
                  fontFamily: jostFont,
                  fontSize: 16,
                },
              },
              createElement(P, null, "Delayed caption"),
            ),
          ),
          createElement(
            Step,
            null,
            createElement(Shape, {
              id: "click-box",
              enter: fade({ duration: 250 }),
              style: {
                position: "absolute",
                left: 260,
                top: 40,
                width: 80,
                height: 60,
                background: "#10B981",
              },
            }),
          ),
        ),
      ),
    );

    const resolution = resolve(doc);

    // motionSchedule works immediately on Resolution before emit().
    const resolvedSlide = resolution.report.slides[0]!;
    const resolvedMotion = motionSchedule(resolvedSlide);
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
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png !== undefined);

    // Emit preserves slide observation identity so motionSchedule works identically on BuildReport.
    const emitted = emit(resolution);
    const emittedMotion = motionSchedule(emitted.report.slides[0]!);
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
              enter: fade({ duration: 200 }),
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 50,
                height: 50,
              },
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

  it("delegates transparent declaration prop validation (including morphKey) to compileDeclaration without core prop readbacks", async () => {
    const morphKey = createMorphScope().handle("step-key");
    const resolve = createResolver([motionCapability, shapeCapability]);
    const doc = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          Slide,
          { id: "transparent-own-validation" },
          createElement(
            Step as unknown as ComponentType<Record<string, unknown>>,
            { n: 1, morphKey },
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
      () => resolve(doc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.map(({ code }) => code).join(",") === "unknown-prop",
    );
  });
});
