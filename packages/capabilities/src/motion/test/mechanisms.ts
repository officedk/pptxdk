import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  createMotionScopeHandle,
  deck,
  defineSlideLayout,
  font,
  group as documentGroup,
  motionScope,
  paragraph,
  shape,
  slide,
  slideLayout,
  slideMaster,
  step,
  text,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import {
  after,
  breathe,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
import { MotionScope, Step } from "@pptxdk/capabilities/motion-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, Text } from "@pptxdk/capabilities/text-elements";
import { transitions } from "@pptxdk/capabilities/transition";
import { createMorphScope } from "@pptxdk/core";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  ProductionMechanismKey,
} from "../../../test/mechanism-verification.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const motionMorphScope = createMorphScope();
const morphCardKey = motionMorphScope.handle("morph-card");

const motionMasterLayout = defineSlideLayout({
  id: "motion-layout",
  name: "Motion Layout",
  type: "title",
  showMasterPlaceholderAnimations: true,
  placeholders: {
    title: { type: "title", x: 32, y: 20, width: 320, height: 48 },
  },
  motion: {
    steps: [
      {
        trigger: "auto",
        effects: [
          {
            target: { placeholder: "title" },
            effect: fade({ duration: 250 }),
          },
        ],
      },
    ],
  },
});

export const motionRoundtripScenario: CapabilityScenario = {
  id: "mechanism-motion-roundtrip",
  capability: "motion",
  description:
    "Dual-author Transitional scenario covering all 11 motion presets, Step click groups, MotionScope sequence/stagger/after(cue), word/paragraph/chart granularity, Group spin, !!morph continuity, and master/layout timing ownership through import roundtrip.",
  visualCheckpoints: [
    "Static canvas pages render all shapes, text, group, and chart with zero diagnostics regardless of animations.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain all 11 animation presets, click steps, after(cue) startAfter dependencies, word/paragraph/chart builds, master/layout timing, and !!morph continuity.",
  ],
  jsx() {
    const scopeA = createMotionScopeHandle();
    const scopeB = createMotionScopeHandle();
    return createElement(
      Deck,
      { width: 640, height: 360, style: { fontFamily: jostFont } },
      createElement(
        SlideMaster,
        {
          id: "motion-master",
          name: "Motion Master",
          placeholders: {
            ftr: { type: "ftr", x: 32, y: 316, width: 200, height: 24 },
          },
          motion: {
            steps: [
              {
                trigger: "click",
                effects: [
                  {
                    target: { placeholder: "ftr" },
                    effect: wipe({ from: "left", duration: 200 }),
                  },
                ],
              },
            ],
          },
        },
        createElement(SlideLayout, { layout: motionMasterLayout }),
      ),
      createElement(
        Slide,
        {
          id: "motion-slide-1",
          layout: motionMasterLayout,
          showMasterPlaceholderAnimations: true,
        },
        createElement(
          MotionScope,
          { scope: scopeA },
          createElement(Shape, {
            id: "fade-shape",
            x: 24,
            y: 24,
            width: 80,
            height: 44,
            fill: "#2563EB",
            enter: fade({ duration: 200, delay: 20 }),
            emphasis: fillColor({
              from: "#2563EB",
              to: "#10B981",
              duration: 300,
            }),
          }),
          createElement(Shape, {
            id: "wipe-shape",
            x: 116,
            y: 24,
            width: 80,
            height: 44,
            fill: "#0284C7",
            enter: wipe({ from: "right", duration: 220 }),
          }),
          createElement(Shape, {
            id: "zoom-shape",
            x: 208,
            y: 24,
            width: 80,
            height: 44,
            fill: "#059669",
            enter: zoom({ duration: 240 }),
          }),
          createElement(Shape, {
            id: "fly-shape",
            x: 300,
            y: 24,
            width: 80,
            height: 44,
            fill: "#D97706",
            enter: fly({ from: "bottom", duration: 260 }),
          }),
          createElement(
            Text,
            {
              id: "words-text",
              x: 24,
              y: 148,
              width: 220,
              height: 48,
              style: { fontFamily: jostFont, fontSize: 14 },
              enter: fade({ by: "word", intervalMs: 50, duration: 180 }),
            },
            "Motion word iterate",
          ),
        ),
        createElement(
          MotionScope,
          { scope: scopeB, startAt: after(scopeA.end, 40) },
          createElement(Shape, {
            id: "float-up-shape",
            x: 24,
            y: 84,
            width: 80,
            height: 44,
            fill: "#7C3AED",
            enter: floatUp({ duration: 220 }),
          }),
          createElement(Shape, {
            id: "float-down-shape",
            x: 116,
            y: 84,
            width: 80,
            height: 44,
            fill: "#9333EA",
            enter: floatDown({ duration: 220 }),
          }),
          createElement(Shape, {
            id: "spin-in-shape",
            x: 208,
            y: 84,
            width: 80,
            height: 44,
            fill: "#DB2777",
            enter: spinIn({ duration: 280 }),
          }),
          createElement(Shape, {
            id: "stretch-shape",
            x: 300,
            y: 84,
            width: 80,
            height: 44,
            fill: "#164E3D",
            enter: stretch({ duration: 240 }),
            emphasis: repeat(
              breathe({ period: 800, from: "#164E3D", to: "#4ADE80" }),
              2,
            ),
          }),
        ),
        createElement(
          Step,
          { n: 1 },
          createElement(
            Text,
            {
              id: "para-text",
              x: 24,
              y: 206,
              width: 220,
              height: 60,
              style: { fontFamily: jostFont, fontSize: 14 },
              enter: fade({ by: "paragraph", duration: 200 }),
            },
            createElement(P, null, "Para one"),
            createElement(P, null, "Para two"),
          ),
          createElement(
            Group,
            {
              id: "spin-group",
              emphasis: spin({ by: 90, duration: 300 }),
            },
            createElement(Shape, {
              id: "group-member",
              x: 410,
              y: 24,
              width: 80,
              height: 44,
              fill: "#F59E0B",
              enter: fade({ duration: 150 }),
            }),
          ),
        ),
        createElement(
          Step,
          { n: 2 },
          createElement(Chart, {
            id: "motion-chart",
            x: 280,
            y: 148,
            width: 240,
            height: 150,
            enter: fade({ by: "series", duration: 140 }),
            axes: {
              cat: { kind: "category" },
              val: { kind: "value" },
            },
            groups: [
              {
                kind: "column",
                axes: { category: "cat", value: "val" },
                data: { kind: "shared-categories", categories: ["A", "B"] },
                series: [
                  { name: "S1", values: [10, 20] },
                  { name: "S2", values: [15, 25] },
                ],
              },
            ],
          }),
          createElement(Shape, {
            id: "morph-card-1",
            morphKey: morphCardKey,
            x: 530,
            y: 24,
            width: 80,
            height: 44,
            fill: "#3B82F6",
          }),
        ),
      ),
      createElement(
        Slide,
        {
          id: "motion-slide-2",
          layout: motionMasterLayout,
          transition: transitions.morph({
            duration: 600,
            by: "object",
            fallback: transitions.fade(),
          }),
        },
        createElement(Shape, {
          id: "morph-card-2",
          morphKey: morphCardKey,
          x: 440,
          y: 180,
          width: 140,
          height: 80,
          fill: "#3B82F6",
        }),
      ),
    );
  },
  document() {
    const scopeA = createMotionScopeHandle();
    const scopeB = createMotionScopeHandle();
    return deck({
      width: 640,
      height: 360,
      style: { fontFamily: jostFont },
    }).append(
      slideMaster({
        id: "motion-master",
        name: "Motion Master",
        placeholders: {
          ftr: { type: "ftr", x: 32, y: 316, width: 200, height: 24 },
        },
        motion: {
          steps: [
            {
              trigger: "click",
              effects: [
                {
                  target: { placeholder: "ftr" },
                  effect: wipe({ from: "left", duration: 200 }),
                },
              ],
            },
          ],
        },
      }).append(slideLayout(motionMasterLayout)),
      slide({
        id: "motion-slide-1",
        layout: motionMasterLayout,
        showMasterPlaceholderAnimations: true,
      }).append(
        motionScope(scopeA).append(
          shape({
            id: "fade-shape",
            x: 24,
            y: 24,
            width: 80,
            height: 44,
            fill: "#2563EB",
            enter: fade({ duration: 200, delay: 20 }),
            emphasis: fillColor({
              from: "#2563EB",
              to: "#10B981",
              duration: 300,
            }),
          }),
          shape({
            id: "wipe-shape",
            x: 116,
            y: 24,
            width: 80,
            height: 44,
            fill: "#0284C7",
            enter: wipe({ from: "right", duration: 220 }),
          }),
          shape({
            id: "zoom-shape",
            x: 208,
            y: 24,
            width: 80,
            height: 44,
            fill: "#059669",
            enter: zoom({ duration: 240 }),
          }),
          shape({
            id: "fly-shape",
            x: 300,
            y: 24,
            width: 80,
            height: 44,
            fill: "#D97706",
            enter: fly({ from: "bottom", duration: 260 }),
          }),
          text("Motion word iterate", {
            id: "words-text",
            x: 24,
            y: 148,
            width: 220,
            height: 48,
            style: { fontFamily: jostFont, fontSize: 14 },
            enter: fade({ by: "word", intervalMs: 50, duration: 180 }),
          }),
        ),
        motionScope(scopeB, { startAt: after(scopeA.end, 40) }).append(
          shape({
            id: "float-up-shape",
            x: 24,
            y: 84,
            width: 80,
            height: 44,
            fill: "#7C3AED",
            enter: floatUp({ duration: 220 }),
          }),
          shape({
            id: "float-down-shape",
            x: 116,
            y: 84,
            width: 80,
            height: 44,
            fill: "#9333EA",
            enter: floatDown({ duration: 220 }),
          }),
          shape({
            id: "spin-in-shape",
            x: 208,
            y: 84,
            width: 80,
            height: 44,
            fill: "#DB2777",
            enter: spinIn({ duration: 280 }),
          }),
          shape({
            id: "stretch-shape",
            x: 300,
            y: 84,
            width: 80,
            height: 44,
            fill: "#164E3D",
            enter: stretch({ duration: 240 }),
            emphasis: repeat(
              breathe({ period: 800, from: "#164E3D", to: "#4ADE80" }),
              2,
            ),
          }),
        ),
        step({ n: 1 }).append(
          text([paragraph("Para one"), paragraph("Para two")], {
            id: "para-text",
            x: 24,
            y: 206,
            width: 220,
            height: 60,
            style: { fontFamily: jostFont, fontSize: 14 },
            enter: fade({ by: "paragraph", duration: 200 }),
          }),
          documentGroup({
            id: "spin-group",
            emphasis: spin({ by: 90, duration: 300 }),
          }).append(
            shape({
              id: "group-member",
              x: 410,
              y: 24,
              width: 80,
              height: 44,
              fill: "#F59E0B",
              enter: fade({ duration: 150 }),
            }),
          ),
        ),
        step({ n: 2 }).append(
          documentChart({
            id: "motion-chart",
            x: 280,
            y: 148,
            width: 240,
            height: 150,
            enter: fade({ by: "series", duration: 140 }),
            axes: {
              cat: { kind: "category" },
              val: { kind: "value" },
            },
            groups: [
              {
                kind: "column",
                axes: { category: "cat", value: "val" },
                data: { kind: "shared-categories", categories: ["A", "B"] },
                series: [
                  { name: "S1", values: [10, 20] },
                  { name: "S2", values: [15, 25] },
                ],
              },
            ],
          }),
          shape({
            id: "morph-card-1",
            morphKey: morphCardKey,
            x: 530,
            y: 24,
            width: 80,
            height: 44,
            fill: "#3B82F6",
          }),
        ),
      ),
      slide({
        id: "motion-slide-2",
        layout: motionMasterLayout,
        transition: transitions.morph({
          duration: 600,
          by: "object",
          fallback: transitions.fade(),
        }),
      }).append(
        shape({
          id: "morph-card-2",
          morphKey: morphCardKey,
          x: 440,
          y: 180,
          width: 140,
          height: 80,
          fill: "#3B82F6",
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 2);
  },
};

type MotionMechanismKey = Extract<ProductionMechanismKey, `motion-${string}`>;

function allEffectsOfSlide(
  validation: Parameters<
    ExhaustiveMechanismVerificationMap["motion-fade"]["assertNative"]
  >[0]["validation"],
) {
  return validation.slides[0]?.animations?.flatMap((s) => s.effects) ?? [];
}

export const motionMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  MotionMechanismKey
> = {
  "motion-fade": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some((e) => e.preset === "fade"),
        `${stage} missing fade`,
      );
    },
  },
  "motion-wipe": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some(
          (e) => e.preset === "wipe" && e.from === "right",
        ),
        `${stage} missing wipe`,
      );
    },
  },
  "motion-zoom": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some((e) => e.preset === "zoom"),
        `${stage} missing zoom`,
      );
    },
  },
  "motion-fly": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some(
          (e) => e.preset === "fly" && e.from === "bottom",
        ),
        `${stage} missing fly`,
      );
    },
  },
  "motion-float": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const effects = allEffectsOfSlide(validation);
      assert.ok(effects.some((e) => e.preset === "floatUp"));
      assert.ok(effects.some((e) => e.preset === "floatDown"));
    },
  },
  "motion-spin-in": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some((e) => e.preset === "spinIn"),
      );
    },
  },
  "motion-stretch": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some((e) => e.preset === "stretch"),
      );
    },
  },
  "motion-spin": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some(
          (e) => e.preset === "spin" && e.by === 90,
        ),
      );
    },
  },
  "motion-breathe": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some(
          (e) => e.preset === "breathe" && e.repeatCount === 2,
        ),
      );
    },
  },
  "motion-fill-color": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      assert.ok(
        allEffectsOfSlide(validation).some(
          (e) => e.preset === "fillColor" && e.toColor === "#10B981",
        ),
      );
    },
  },
  "motion-choreography": {
    scenario: motionRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const s1 = validation.slides[0]!;
      assert.equal(s1.showMasterPlaceholderAnimations, true);
      assert.equal(s1.animations?.length, 3);
      assert.equal(s1.animations?.[0]?.trigger, "auto");
      assert.equal(s1.animations?.[1]?.trigger, "click");
      assert.equal(s1.animations?.[2]?.trigger, "click");
      const effects = allEffectsOfSlide(validation);
      assert.ok(effects.some((e) => e.startAfter !== undefined));
      assert.ok(effects.some((e) => e.by === "word" && e.intervalMs === 50));
      assert.ok(effects.some((e) => e.targetRange?.type === "paragraph"));
      assert.ok(effects.some((e) => e.targetRange?.type === "chart"));
      assert.equal(
        validation.masters?.[0]?.animations?.[0]?.effects[0]?.preset,
        "wipe",
      );
      assert.equal(
        validation.masters?.[0]?.layouts[0]?.animations?.[0]?.effects[0]
          ?.preset,
        "fade",
      );
      assert.equal(
        validation.slides[0]?.objects.at(-1)?.name,
        validation.slides[1]?.objects[0]?.name,
      );
    },
  },
};
