import type { PaintContext, PaintOutput } from "@pptxdk/core/capability";
import { defineCanvasMechanism, defineShowMechanism } from "@pptxdk/import";
import { describe, it } from "vitest";

import {
  runProductionMechanismVerifications,
  textBoxRoundtripScenario,
  type ExhaustiveMechanismVerificationMap,
} from "./mechanism-verification.ts";

// Compile-time directional protocol checks (never executed at runtime).
function compileDirectionalChecks(): void {
  // @ts-expect-error -- canvas mechanism missing emit direction
  defineCanvasMechanism({
    key: "bad-canvas-no-emit",
    role: "canvas",
    paint: (_model: unknown, _ctx: PaintContext): PaintOutput => ({ svg: "" }),
    import: {
      target: "shape-tree",
      localName: "sp",
      read: () => ({}),
    },
  });

  // @ts-expect-error -- canvas mechanism missing import direction
  defineCanvasMechanism({
    key: "bad-canvas-no-import",
    role: "canvas",
    emit: () => ({ xml: "" }),
    paint: (_model: unknown, _ctx: PaintContext): PaintOutput => ({ svg: "" }),
  });

  // @ts-expect-error -- canvas mechanism missing paint direction
  defineCanvasMechanism({
    key: "bad-canvas-no-paint",
    role: "canvas",
    emit: () => ({ xml: "" }),
    import: {
      target: "shape-tree",
      localName: "sp",
      read: () => ({}),
    },
  });

  // @ts-expect-error -- show mechanism missing emit direction
  defineShowMechanism({
    key: "bad-show-no-emit",
    role: "show",
    import: {
      target: "slide",
      slot: "transition",
      read: () => ({}),
    },
  });

  // @ts-expect-error -- show mechanism missing import direction
  defineShowMechanism({
    key: "bad-show-no-import",
    role: "show",
    emit: () => ({ xml: "" }),
  });

  defineShowMechanism({
    key: "bad-show-with-paint",
    role: "show",
    emit: () => ({ xml: "" }),
    // @ts-expect-error -- show mechanism forbids paint direction
    paint: () => ({ svg: "" }),
    import: {
      target: "slide",
      slot: "transition",
      read: () => ({}),
    },
  });

  const sampleCanvas = defineCanvasMechanism({
    key: "sample-canvas",
    role: "canvas",
    emit: () => ({ xml: "" }),
    paint: () => ({ svg: "" }),
    import: {
      target: "shape-tree",
      localName: "sp",
      read: () => ({}),
    },
  });

  const sampleShow = defineShowMechanism({
    key: "sample-show",
    role: "show",
    emit: () => ({ xml: "" }),
    import: {
      target: "slide",
      slot: "transition",
      read: () => ({}),
    },
  });

  type SampleMechanisms = typeof sampleCanvas | typeof sampleShow;

  // @ts-expect-error -- verification map missing sample-show key
  const missingKeyMap: ExhaustiveMechanismVerificationMap<SampleMechanisms> = {
    "sample-canvas": {
      scenario: textBoxRoundtripScenario,
      assertNative: () => {},
      assertPaint: () => {},
    },
  };

  const missingPaintAssertMap: ExhaustiveMechanismVerificationMap<SampleMechanisms> =
    {
      // @ts-expect-error -- canvas verification entry requires assertPaint
      "sample-canvas": {
        scenario: textBoxRoundtripScenario,
        assertNative: () => {},
      },
      "sample-show": {
        scenario: textBoxRoundtripScenario,
        assertNative: () => {},
      },
    };

  const invalidShowPaintMap: ExhaustiveMechanismVerificationMap<SampleMechanisms> =
    {
      "sample-canvas": {
        scenario: textBoxRoundtripScenario,
        assertNative: () => {},
        assertPaint: () => {},
      },
      "sample-show": {
        scenario: textBoxRoundtripScenario,
        assertNative: () => {},
        // @ts-expect-error -- show verification entry forbids assertPaint
        assertPaint: () => {},
      },
    };

  const missingNativeAssertMap: ExhaustiveMechanismVerificationMap<SampleMechanisms> =
    {
      "sample-canvas": {
        scenario: textBoxRoundtripScenario,
        assertNative: () => {},
        assertPaint: () => {},
      },
      // @ts-expect-error -- show verification entry requires assertNative
      "sample-show": {
        scenario: textBoxRoundtripScenario,
      },
    };

  void missingKeyMap;
  void missingPaintAssertMap;
  void invalidShowPaintMap;
  void missingNativeAssertMap;
}
void compileDirectionalChecks;

describe("Exhaustive production mechanism verification", () => {
  it("runs dual-author build, import roundtrip, and native assertions for every wired mechanism", async () => {
    await runProductionMechanismVerifications();
  });
});
