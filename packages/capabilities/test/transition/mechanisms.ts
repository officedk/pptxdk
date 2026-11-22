import assert from "node:assert/strict";

import {
  deck,
  defineSlideLayout,
  shape,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { transitions } from "@pptxdk/capabilities/transition";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../mechanism-verification.ts";

const stripsLayout = defineSlideLayout({
  id: "strips-layout",
  name: "Strips Layout",
  type: "blank",
  transition: transitions.strips({ direction: "rightDown" }),
});

const pushLayout = defineSlideLayout({
  id: "push-layout",
  name: "Push Layout",
  type: "cust",
  transition: transitions.push({ from: "bottom" }),
});

export const transitionRoundtripScenario: CapabilityScenario = {
  id: "mechanism-transition-roundtrip",
  capability: "transition",
  description:
    "Minimal dual-author Transitional scenario covering fade, strips, push, flip, morph, advanceAfter, and three-layer master/layout/slide ownership through import roundtrip.",
  visualCheckpoints: [
    "Pages 1-3 render clean colored shapes on white canvas with zero diagnostics regardless of transitions.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX preserve master fade+advTm, layout strips and push, and slide morph (with push fallback) and flip (with strips fallback).",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        {
          id: "brand-master",
          name: "Brand Master",
          transition: transitions.fade({ duration: 500 }),
          advanceAfter: 3000,
        },
        createElement(SlideLayout, { layout: stripsLayout }),
        createElement(SlideLayout, { layout: pushLayout }),
      ),
      createElement(
        Slide,
        { id: "slide-inherited-strips", layout: stripsLayout },
        createElement(Shape, {
          id: "card-1",
          x: 40,
          y: 40,
          width: 140,
          height: 80,
          fill: "#2563EB",
        }),
      ),
      createElement(
        Slide,
        {
          id: "slide-morph",
          layout: pushLayout,
          transition: transitions.morph({
            duration: 1200,
            by: "word",
            fallback: transitions.push({ from: "left" }),
          }),
          advanceAfter: 2500,
        },
        createElement(Shape, {
          id: "card-2",
          x: 200,
          y: 100,
          width: 160,
          height: 90,
          fill: "#059669",
        }),
      ),
      createElement(
        Slide,
        {
          id: "slide-flip",
          layout: stripsLayout,
          transition: transitions.flip({
            direction: "left",
            fallback: transitions.strips({ direction: "leftUp" }),
          }),
        },
        createElement(Shape, {
          id: "card-3",
          x: 360,
          y: 140,
          width: 160,
          height: 90,
          fill: "#D97706",
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "brand-master",
        name: "Brand Master",
        transition: transitions.fade({ duration: 500 }),
        advanceAfter: 3000,
      }).append(slideLayout(stripsLayout), slideLayout(pushLayout)),
      slide({ id: "slide-inherited-strips", layout: stripsLayout }).append(
        shape({
          id: "card-1",
          x: 40,
          y: 40,
          width: 140,
          height: 80,
          fill: "#2563EB",
        }),
      ),
      slide({
        id: "slide-morph",
        layout: pushLayout,
        transition: transitions.morph({
          duration: 1200,
          by: "word",
          fallback: transitions.push({ from: "left" }),
        }),
        advanceAfter: 2500,
      }).append(
        shape({
          id: "card-2",
          x: 200,
          y: 100,
          width: 160,
          height: 90,
          fill: "#059669",
        }),
      ),
      slide({
        id: "slide-flip",
        layout: stripsLayout,
        transition: transitions.flip({
          direction: "left",
          fallback: transitions.strips({ direction: "leftUp" }),
        }),
      }).append(
        shape({
          id: "card-3",
          x: 360,
          y: 140,
          width: 160,
          height: 90,
          fill: "#D97706",
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.masters?.length, 1);
    assert.equal(validation.slides.length, 3);
  },
};

export const transitionMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  | "transition-fade"
  | "transition-strips"
  | "transition-push"
  | "transition-flip"
  | "transition-morph"
> = {
  "transition-fade": {
    scenario: transitionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.masters?.[0]?.advanceAfterMs, 3000);
      assert.deepEqual(validation.masters?.[0]?.transition, {
        kind: "fade",
        speed: "fast",
        durationMs: 500,
        fallback: { kind: "fade", speed: "fast" },
      });
    },
  },
  "transition-strips": {
    scenario: transitionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.deepEqual(validation.masters?.[0]?.layouts[0]?.transition, {
        kind: "strips",
        speed: "slow",
        direction: "rd",
      });
      assert.equal(validation.slides[0]?.transition, undefined);
    },
  },
  "transition-push": {
    scenario: transitionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.deepEqual(validation.masters?.[0]?.layouts[1]?.transition, {
        kind: "push",
        speed: "slow",
        direction: "d",
      });
    },
  },
  "transition-flip": {
    scenario: transitionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.deepEqual(validation.slides[2]?.transition, {
        kind: "flip",
        speed: "slow",
        direction: "l",
        fallback: { kind: "strips", speed: "slow", direction: "lu" },
      });
    },
  },
  "transition-morph": {
    scenario: transitionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides[1]?.advanceAfterMs, 2500);
      assert.deepEqual(validation.slides[1]?.transition, {
        kind: "morph",
        speed: "slow",
        durationMs: 1200,
        by: "word",
        fallback: { kind: "push", speed: "slow", direction: "l" },
      });
    },
  },
};
