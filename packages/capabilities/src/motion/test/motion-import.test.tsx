import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  deck,
  defineSlideLayout,
  font,
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
  appear,
  breathe,
  collapse,
  createMotionScopeHandle,
  disappear,
  fade,
  fadeOut,
  fillColor,
  floatDown,
  floatOut,
  floatUp,
  fly,
  flyOut,
  motionCapability,
  onStep,
  repeat,
  spin,
  spinIn,
  spinOut,
  stretch,
  wipe,
  wipeOut,
  zoom,
  zoomOut,
} from "@pptxdk/capabilities/motion";
import {
  breatheEffectSchema,
  breatheOptionsSchema,
  easingSchema,
  effectOptionsSchema,
  effectPresetSchema,
  effectSchema,
  emphasisEffectDeclarationSchema,
  emphasisEffectSchema,
  entranceEffectDeclarationSchema,
  entranceEffectSchema,
  exitEffectDeclarationSchema,
  exitEffectSchema,
  fillColorOptionsSchema,
  motionDirectionSchema,
  motionGranularitySchema,
  motionHexColorSchema,
  motionMomentOffsetSchema,
  motionMomentSchema,
  motionPropsSchema,
  motionScopeAuthorPropsSchema,
  motionScopeHandleSchema,
  motionScopeModeSchema,
  motionStepAuthorPropsSchema,
  motionTargetRangeSchema,
  motionTimingSchema,
  ownerMotionSchema,
  repeatCountSchema,
  repeatedEffectSchema,
  spinOptionsSchema,
  steppedEffectSchema,
} from "@pptxdk/capabilities/motion";
import { MotionScope, Step } from "@pptxdk/capabilities/motion-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, Text } from "@pptxdk/capabilities/text-elements";
import { transitions } from "@pptxdk/capabilities/transition";
import { CommonContainer, CommonNode, createMorphScope } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  morphKeySchema,
  morphPropsSchema,
  parseNativeMorphToken,
} from "../../shared/morph-api.ts";

describe("Motion Zod v4 schemas & Morph handle identity", () => {
  it("validates all object effect presets, triggers, durations, delays, easing, repeat counts, and target ranges", () => {
    assert.equal(motionDirectionSchema.parse("left"), "left");
    assert.equal(effectPresetSchema.parse("fillColor"), "fillColor");
    assert.equal(motionScopeModeSchema.parse("stagger"), "stagger");
    assert.equal(
      motionGranularitySchema.parse("seriesElements"),
      "seriesElements",
    );
    assert.equal(repeatCountSchema.parse("forever"), "forever");
    assert.equal(repeatCountSchema.parse(3), 3);
    assert.equal(repeatCountSchema.safeParse(0).success, false);

    assert.deepEqual(easingSchema.parse({ accelerate: 0.3, decelerate: 0.7 }), {
      accelerate: 0.3,
      decelerate: 0.7,
    });
    assert.equal(
      easingSchema.safeParse({ accelerate: 0.6, decelerate: 0.5 }).success,
      false,
    );

    assert.deepEqual(
      motionTargetRangeSchema.parse({ type: "paragraph", start: 0, end: 2 }),
      { type: "paragraph", start: 0, end: 2 },
    );
    assert.deepEqual(
      motionTargetRangeSchema.parse({
        type: "chart",
        step: "ptInSeries",
        seriesIndex: 1,
        categoryIndex: 2,
      }),
      {
        type: "chart",
        step: "ptInSeries",
        seriesIndex: 1,
        categoryIndex: 2,
      },
    );

    const entrances = [
      appear({ delay: 25 }),
      fade({ duration: 250, delay: 50 }),
      wipe({ from: "top", by: "paragraph" }),
      zoom({ easing: { accelerate: 0.2 } }),
      fly({ from: "right", by: "word", intervalMs: 80 }),
      floatUp({ by: "letter", intervalMs: 60 }),
      floatDown({ duration: 400 }),
      spinIn({ duration: 500 }),
      stretch({ duration: 300 }),
    ] as const;

    for (const eff of entrances) {
      assert.deepEqual(entranceEffectSchema.parse(eff), eff);
      assert.deepEqual(entranceEffectDeclarationSchema.parse(eff), eff);
      assert.deepEqual(effectSchema.parse(eff), eff);
      assert.equal(exitEffectSchema.safeParse(eff).success, false);
    }

    const exits = [
      disappear({ delay: 10 }),
      fadeOut({ duration: 250, delay: 40 }),
      wipeOut({ from: "left" }),
      flyOut({ from: "right" }),
      zoomOut({ easing: { decelerate: 0.3 } }),
      floatOut({ duration: 400 }),
      spinOut({ duration: 500 }),
      collapse({ duration: 300 }),
    ] as const;

    for (const eff of exits) {
      assert.deepEqual(exitEffectSchema.parse(eff), eff);
      assert.deepEqual(exitEffectDeclarationSchema.parse(eff), eff);
      assert.deepEqual(effectSchema.parse(eff), eff);
      assert.equal(entranceEffectSchema.safeParse(eff).success, false);
      assert.equal(emphasisEffectSchema.safeParse(eff).success, false);
    }

    const spinEff = spin({ by: -90, duration: 400 });
    const breatheEff = breathe({
      period: 800,
      from: "#164E3D",
      to: "#4ADE80",
    });
    const fillEff = fillColor({
      from: "#EEEEEE",
      to: "#16324F",
      duration: 600,
    });
    assert.deepEqual(emphasisEffectSchema.parse(spinEff), spinEff);
    assert.deepEqual(breatheEffectSchema.parse(breatheEff), breatheEff);
    assert.deepEqual(emphasisEffectSchema.parse(fillEff), fillEff);
    assert.equal(motionHexColorSchema.parse("#164e3d"), "#164E3D");

    const repeatedSpin = repeat(spinEff, 2);
    const repeatedBreathe = repeat(breatheEff, "forever");
    assert.deepEqual(repeatedEffectSchema.parse(repeatedSpin), repeatedSpin);
    assert.deepEqual(
      emphasisEffectDeclarationSchema.parse(repeatedBreathe),
      repeatedBreathe,
    );

    assert.equal(
      effectOptionsSchema.safeParse({ duration: 500 }).success,
      true,
    );
    assert.equal(spinOptionsSchema.safeParse({ from: 180 }).success, true);
    assert.equal(
      fillColorOptionsSchema.safeParse({ from: "#000000", to: "#FFFFFF" })
        .success,
      true,
    );
    assert.equal(
      breatheOptionsSchema.safeParse({
        period: 800,
        from: "#164E3D",
        to: "#4ADE80",
      }).success,
      true,
    );
  });

  it("preserves runtime handle identity for MotionScopeHandle, MotionTiming, and MorphKey", () => {
    const scope = createMotionScopeHandle();
    const cue = scope.cue("ready");
    const delayed = after(cue, 120);

    assert.equal(motionScopeHandleSchema.parse(scope), scope);
    assert.deepEqual(motionMomentSchema.parse(cue), cue);
    assert.deepEqual(motionMomentOffsetSchema.parse(delayed), delayed);
    assert.deepEqual(motionTimingSchema.parse(delayed), delayed);
    assert.equal(
      motionScopeHandleSchema.safeParse({ type: "other" }).success,
      false,
    );

    assert.equal(
      motionScopeAuthorPropsSchema.parse({
        scope,
        mode: "stagger",
        each: 100,
        startAt: delayed,
        publishStart: scope.cue("begin"),
        publishEnd: cue,
      }).scope,
      scope,
    );
    assert.equal(motionStepAuthorPropsSchema.parse({ n: 2 }).n, 2);
    assert.equal(
      motionPropsSchema.safeParse({
        enter: [fade(), fly({ from: "left" })],
        emphasis: spin(),
        motion: true,
      }).success,
      true,
    );

    assert.equal(
      ownerMotionSchema.safeParse({
        steps: [
          {
            trigger: "click",
            effects: [
              {
                target: { placeholder: "title" },
                effect: fade({ duration: 300 }),
              },
            ],
          },
        ],
      }).success,
      true,
    );

    const morph = createMorphScope();
    const hero = morph.handle("hero");
    assert.equal(morphKeySchema.parse(hero), hero);
    assert.equal(morphPropsSchema.parse({ morphKey: hero }).morphKey, hero);
    assert.equal(morphKeySchema.safeParse({ label: "hero" }).success, false);
    assert.equal(parseNativeMorphToken("!!hero"), "hero");
    assert.equal(parseNativeMorphToken("!!"), undefined);
    assert.equal(parseNativeMorphToken("Hero Card"), undefined);
  });
});

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);

const resolveAll = createResolver(documentCapabilities);
const importAll = createImporter(documentCapabilities);

function rewritePartXml(
  pptx: Uint8Array,
  partPath: string,
  transform: (xml: string) => string,
): Uint8Array {
  const parts = unzipSync(pptx);
  const raw = parts[partPath];
  assert.ok(raw !== undefined, `Missing part ${partPath}`);
  const next = transform(strFromU8(raw));
  return zipSync({ ...parts, [partPath]: strToU8(next) });
}

function removePart(pptx: Uint8Array, partPath: string): Uint8Array {
  const parts = unzipSync(pptx);
  delete parts[partPath];
  return zipSync(parts);
}

describe("Motion mechanisms, roundtrip target identity, Morph continuity, ownership, and partial recovery", () => {
  it("registers typed show mechanisms on motionCapability with role 'show' and no paint", () => {
    const keys = motionCapability.mechanisms?.map((m) => m.key) ?? [];
    assert.deepEqual(keys, [
      "motion-appear",
      "motion-fade",
      "motion-wipe",
      "motion-zoom",
      "motion-fly",
      "motion-float",
      "motion-spin-in",
      "motion-stretch",
      "motion-spin",
      "motion-breathe",
      "motion-fill-color",
      "motion-exit",
      "motion-choreography",
    ]);
    for (const mechanism of motionCapability.mechanisms ?? []) {
      assert.equal(mechanism.role, "show");
      assert.equal(typeof mechanism.emit, "function");
      const importSpec = mechanism.import as {
        readonly target: string;
        readonly slot: string;
      };
      assert.equal(importSpec.target, "slide");
      assert.equal(importSpec.slot, "timing");
      assert.equal("paint" in mechanism, false);
    }
  });

  it("imports <p:timing> onto live CommonNodes and preserves object, paragraph/word/letter, and chart series/point targets after shifting spids", async () => {
    const intro = createMotionScopeHandle();
    const afterIntro = createMotionScopeHandle();
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 720, height: 405 },
        createElement(
          Slide,
          { id: "motion-roundtrip" },
          createElement(
            MotionScope,
            { scope: intro },
            createElement(Shape, {
              id: "hero-shape",
              enter: fade({
                duration: 250,
                delay: 40,
                easing: { accelerate: 0.25 },
              }),
              emphasis: fillColor({
                from: "#2563EB",
                to: "#10B981",
                duration: 400,
              }),
              style: {
                position: "absolute",
                left: 30,
                top: 24,
                width: 120,
                height: 60,
                background: "#2563EB",
              },
            }),
            createElement(
              Text,
              {
                id: "words-text",
                enter: fly({
                  from: "right",
                  by: "word",
                  intervalMs: 60,
                  duration: 200,
                }),
                style: {
                  position: "absolute",
                  left: 170,
                  top: 24,
                  width: 220,
                  height: 60,
                  fontFamily: jostFont,
                  fontSize: 16,
                },
              },
              "Hello motion words",
            ),
          ),
          createElement(
            MotionScope,
            { scope: afterIntro, startAt: after(intro.end, 80) },
            createElement(Shape, {
              id: "breather",
              enter: spinIn({ duration: 300 }),
              emphasis: repeat(
                breathe({ period: 800, from: "#164E3D", to: "#4ADE80" }),
                2,
              ),
              style: {
                position: "absolute",
                left: 30,
                top: 110,
                width: 100,
                height: 50,
                background: "#164E3D",
              },
            }),
          ),
          createElement(
            Step,
            { n: 1 },
            createElement(
              Text,
              {
                id: "para-text",
                enter: fade({ by: "paragraph", duration: 250 }),
                style: {
                  position: "absolute",
                  left: 170,
                  top: 110,
                  width: 240,
                  height: 90,
                  fontFamily: jostFont,
                  fontSize: 14,
                },
              },
              createElement(P, null, "First paragraph"),
              createElement(P, null, "Second paragraph"),
            ),
            createElement(
              Group,
              { id: "spin-group", emphasis: spin({ by: 90, duration: 400 }) },
              createElement(Shape, {
                id: "group-box",
                enter: zoom({ duration: 220 }),
                style: {
                  position: "absolute",
                  left: 430,
                  top: 30,
                  width: 80,
                  height: 60,
                  background: "#F59E0B",
                },
              }),
            ),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(Chart, {
              id: "sales-chart",
              enter: fade({ by: "seriesElements", duration: 120 }),
              axes: {
                cat: { kind: "category" },
                val: { kind: "value" },
              },
              groups: [
                {
                  kind: "column",
                  axes: { category: "cat", value: "val" },
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2"],
                  },
                  series: [
                    { name: "Rev", values: [10, 20] },
                    { name: "Cost", values: [5, 8] },
                  ],
                },
              ],
              style: {
                position: "absolute",
                left: 420,
                top: 120,
                width: 260,
                height: 220,
              },
            }),
          ),
        ),
      ),
    );

    const emitted1 = emit(resolveAll(doc));
    const v1 = await validatePptx(emitted1.pptx);
    assert.deepEqual(v1.errors, []);

    const imported = importAll(emitted1.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    // Post-import edits: insert a leading static shape before existing children so EVERY emitted spid shifts +1,
    // add a 3rd paragraph to para-text, and add a 3rd category to sales-chart!
    const importedSlides = imported.document.children.filter(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide",
    );
    const importedSlide = importedSlides[0]!;
    const insertedSpacer = shape({
      id: "inserted-leading-spacer",
      style: {
        position: "absolute",
        left: 5,
        top: 5,
        width: 20,
        height: 20,
        background: "#CBD5E1",
      },
    });
    const firstChild = importedSlide.children[0]!;
    assert.ok(firstChild instanceof CommonNode);
    insertedSpacer.moveTo(importedSlide, firstChild);

    // Find imported live text & chart nodes and extend them to verify dynamic re-targeting
    const allNodes: CommonNode[] = [];
    const visit = (nodes: readonly unknown[]): void => {
      for (const n of nodes) {
        if (n instanceof CommonNode) {
          allNodes.push(n);
          visit(n.children);
        }
      }
    };
    visit(importedSlide.children);

    const extractText = (value: unknown): string => {
      if (typeof value === "string") return value;
      if (
        typeof value === "object" &&
        value !== null &&
        "text" in value &&
        typeof (value as { readonly text?: unknown }).text === "string"
      ) {
        return (value as { readonly text: string }).text;
      }
      if (value instanceof CommonNode) {
        return value.children.map(extractText).join("");
      }
      return "";
    };
    const paraTextNode = allNodes.find(
      (n): n is CommonContainer =>
        n instanceof CommonContainer &&
        n.kind === "text" &&
        extractText(n).includes("First paragraph"),
    );
    assert.ok(paraTextNode !== undefined);
    paraTextNode.append(paragraph("Third added paragraph"));

    const chartNode = allNodes.find((n) => n.kind === "chart");
    assert.ok(chartNode !== undefined);
    chartNode.set({
      groups: [
        {
          kind: "column",
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q2", "Q3"],
          },
          series: [
            { name: "Rev", values: [10, 20, 30] },
            { name: "Cost", values: [5, 8, 11] },
          ],
        },
      ],
    });

    const emitted2 = emit(resolveAll(imported.document));
    const v2 = await validatePptx(emitted2.pptx);
    assert.deepEqual(v2.errors, []);

    const s2 = v2.slides[0]!;
    // First object is the newly inserted spacer -> unanimated
    const spacerSpid = s2.objects[0]!.id;
    const allAnim2 = s2.animations?.flatMap((step) => step.effects) ?? [];
    assert.equal(
      allAnim2.some((e) => e.targetId === spacerSpid),
      false,
    );

    // Hero shape shifted spid yet still has fade + fillColor!
    const heroObj2 = s2.objects[1]!;
    const heroAnims = allAnim2.filter((e) => e.targetId === heroObj2.id);
    assert.equal(
      heroAnims.some((e) => e.preset === "fade" && e.delayMs === 40),
      true,
    );
    assert.equal(
      heroAnims.some(
        (e) => e.preset === "fillColor" && e.fromColor === "#2563EB",
      ),
      true,
    );

    // Paragraph text now covers all 3 paragraphs (indices 0, 1, 2)!
    const paraObj2 = s2.objects.find((o) =>
      o.text?.paragraphs.some((p) =>
        p.runs.some((r) => r.text.includes("Third added paragraph")),
      ),
    );
    assert.ok(paraObj2 !== undefined);
    const paraAnims = allAnim2.filter((e) => e.targetId === paraObj2.id);
    assert.equal(paraAnims.length, 3);
    assert.deepEqual(
      paraAnims.map((e) => e.targetRange),
      [
        { type: "paragraph", start: 0, end: 0 },
        { type: "paragraph", start: 1, end: 1 },
        { type: "paragraph", start: 2, end: 2 },
      ],
    );

    // Chart now targets updated 2 series x 3 categories = 6 points (+ gridLegend)!
    const chartObj2 = s2.objects.find((o) => o.kind === "chart");
    assert.ok(chartObj2 !== undefined);
    const chartAnims = allAnim2.filter((e) => e.targetId === chartObj2.id);
    assert.equal(
      chartAnims.filter(
        (e) =>
          e.targetRange?.type === "chart" &&
          e.targetRange.step === "ptInSeries",
      ).length,
      6,
    );

    // Word iterate & startAfter preserved
    const wordsObj2 = s2.objects[2]!;
    const wordsAnim = allAnim2.find((e) => e.targetId === wordsObj2.id);
    assert.equal(wordsAnim?.by, "word");
    assert.equal(wordsAnim?.intervalMs, 60);
    const breatherObj2 = s2.objects[3]!;
    const breatherEnter = allAnim2.find(
      (e) => e.targetId === breatherObj2.id && e.preset === "spinIn",
    );
    assert.ok(breatherEnter?.startAfter !== undefined);
  });

  it("reconstructs cross-slide Morph continuity ONLY from !!... tokens and never guesses from plain display names or geometry", async () => {
    const morph = createMorphScope();
    const tokenKey = morph.handle("card-hero");

    const sourceDeck = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1" }).append(
        shape({
          id: "morph-1",
          morphKey: tokenKey,
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 120,
            height: 80,
            background: "#2563EB",
          },
        }),
        shape({
          id: "plain-1",
          name: "Shared Display Name",
          style: {
            position: "absolute",
            left: 240,
            top: 40,
            width: 120,
            height: 80,
            background: "#10B981",
          },
        }),
      ),
      slide({
        id: "s2",
        transition: transitions.morph({
          duration: 600,
          by: "object",
          fallback: transitions.fade(),
        }),
      }).append(
        shape({
          id: "morph-2",
          morphKey: tokenKey,
          style: {
            position: "absolute",
            left: 80,
            top: 120,
            width: 200,
            height: 110,
            background: "#2563EB",
          },
        }),
        shape({
          id: "plain-2",
          name: "Shared Display Name",
          style: {
            position: "absolute",
            left: 240,
            top: 40,
            width: 120,
            height: 80,
            background: "#10B981",
          },
        }),
      ),
    );

    const emitted1 = emit(resolveAll(sourceDeck));
    const imported = importAll(emitted1.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedSlides = imported.document.children.filter(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide",
    );
    const slide1 = importedSlides[0]!;
    const slide2 = importedSlides[1]!;
    const morphObj1 = slide1.children[0] as CommonNode;
    const plainObj1 = slide1.children[1] as CommonNode;
    const morphObj2 = slide2.children[0] as CommonNode;
    const plainObj2 = slide2.children[1] as CommonNode;

    assert.ok(morphObj1.props["morphKey"] !== undefined);
    assert.equal(morphObj1.props["morphKey"], morphObj2.props["morphKey"]);
    assert.equal(morphObj1.props["name"], undefined);
    assert.equal(morphObj2.props["name"], undefined);

    // Plain display names & matching geometry MUST NOT become morphKey
    assert.equal(plainObj1.props["morphKey"], undefined);
    assert.equal(plainObj2.props["morphKey"], undefined);
    assert.equal(plainObj1.props["name"], "Shared Display Name");

    const emitted2 = emit(resolveAll(imported.document));
    const v2 = await validatePptx(emitted2.pptx);
    assert.deepEqual(v2.errors, []);
    assert.equal(
      v2.slides[0]!.objects[0]!.name,
      v2.slides[1]!.objects[0]!.name,
    );
    assert.equal(v2.slides[0]!.objects[0]!.name.startsWith("!!"), true);
    assert.equal(v2.slides[0]!.objects[1]!.name, "Shared Display Name");
  });

  it("retains SlideMaster, SlideLayout, and Slide <p:timing> ownership plus showMasterPhAnim (0/1)", async () => {
    const animLayout = defineSlideLayout({
      id: "anim-layout",
      name: "Animated Layout",
      type: "title",
      showMasterPlaceholderAnimations: false,
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 400,
            height: 60,
          },
        },
      },
      motion: {
        steps: [
          {
            trigger: "click",
            effects: [
              {
                target: { placeholder: "title" },
                effect: fade({ duration: 300 }),
              },
            ],
          },
        ],
      },
    });

    const masterDeck = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "master-1",
        name: "Brand Master",
        placeholders: {
          ftr: {
            type: "ftr",
            style: {
              position: "absolute",
              left: 40,
              top: 310,
              width: 200,
              height: 24,
            },
          },
        },
        motion: {
          steps: [
            {
              trigger: "auto",
              effects: [
                {
                  target: { placeholder: "ftr" },
                  effect: wipe({ from: "left", duration: 250 }),
                },
              ],
            },
          ],
        },
      }).append(slideLayout(animLayout)),
      slide({
        id: "slide-ph-anim",
        layout: animLayout,
        showMasterPlaceholderAnimations: true,
      }).append(
        shape({
          id: "slide-card",
          enter: zoom({ duration: 200 }),
          style: {
            position: "absolute",
            left: 60,
            top: 120,
            width: 140,
            height: 80,
            background: "#2563EB",
          },
        }),
      ),
    );

    const emitted1 = emit(resolveAll(masterDeck));
    const v1 = await validatePptx(emitted1.pptx);
    assert.deepEqual(v1.errors, []);
    assert.equal(v1.masters?.[0]?.animations?.[0]?.effects[0]?.preset, "wipe");
    assert.equal(
      v1.masters?.[0]?.layouts[0]?.showMasterPlaceholderAnimations,
      false,
    );
    assert.equal(
      v1.masters?.[0]?.layouts[0]?.animations?.[0]?.effects[0]?.preset,
      "fade",
    );
    assert.equal(v1.slides[0]?.showMasterPlaceholderAnimations, true);
    assert.equal(v1.slides[0]?.animations?.[0]?.effects[0]?.preset, "zoom");

    const imported = importAll(emitted1.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const emitted2 = emit(resolveAll(imported.document));
    const v2 = await validatePptx(emitted2.pptx);
    assert.deepEqual(v2.errors, []);
    assert.equal(v2.masters?.[0]?.animations?.[0]?.effects[0]?.preset, "wipe");
    assert.equal(
      v2.masters?.[0]?.layouts[0]?.showMasterPlaceholderAnimations,
      false,
    );
    assert.equal(
      v2.masters?.[0]?.layouts[0]?.animations?.[0]?.effects[0]?.preset,
      "fade",
    );
    assert.equal(v2.slides[0]?.showMasterPlaceholderAnimations, true);
    assert.equal(v2.slides[0]?.animations?.[0]?.effects[0]?.preset, "zoom");
  });

  it("emits localized diagnostics with cascaded dependencies for missing targets or unsupported <p:animMotion>, while preserving independent animations", async () => {
    const scopeA = createMotionScopeHandle();
    const scopeDependent = createMotionScopeHandle();
    const png1x1 = new Uint8Array(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
        "base64",
      ),
    );

    const d = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "partial-recovery" },
          createElement(
            MotionScope,
            { scope: scopeA },
            createElement(
              Text,
              {
                id: "target-missing",
                enter: fade({ by: "word", intervalMs: 50, duration: 200 }),
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 80,
                  height: 60,
                  fontFamily: jostFont,
                  fontSize: 14,
                },
              },
              "Missing target text",
            ),
          ),
          createElement(
            MotionScope,
            { scope: scopeDependent, startAt: after(scopeA.end, 50) },
            createElement(Shape, {
              id: "dependent-shape",
              enter: fade({ duration: 200 }),
              style: {
                position: "absolute",
                left: 120,
                top: 20,
                width: 80,
                height: 60,
                background: "#F59E0B",
              },
            }),
          ),
          createElement(Shape, {
            id: "anim-motion-shape",
            enter: fade({ duration: 250 }),
            style: {
              position: "absolute",
              left: 220,
              top: 20,
              width: 80,
              height: 60,
              background: "#EF4444",
            },
          }),
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "independent-shape",
              enter: wipe({ from: "left", duration: 300 }),
              style: {
                position: "absolute",
                left: 340,
                top: 20,
                width: 100,
                height: 60,
                background: "#10B981",
              },
            }),
          ),
        ),
      ),
    );

    const base = emit(resolveAll(d)).pptx;
    // Inject two failure modes in ppt/slides/slide1.xml:
    // 1) Make the first animated effect point to non-existent spid="999"
    // 2) Replace the 3rd effect's childTnLst behavior with unsupported <p:animMotion>
    const corrupted = rewritePartXml(base, "ppt/slides/slide1.xml", (xml) => {
      let replacedFirst = false;
      const withMissingSpid = xml.replace(/spid="2"/gu, (match) => {
        if (!replacedFirst) {
          replacedFirst = true;
          return 'spid="999"';
        }
        return match;
      });
      // Replace the <p:animEffect .../> on spid="4" with <p:animMotion path="M 0 0 L 0.25 0"/>
      const next = withMissingSpid.replace(
        /<p:animEffect\b[^>]*><p:cBhvr><p:cTn id="(\d+)"[^>]*><\/p:cTn><p:tgtEl><p:spTgt spid="4"><\/p:spTgt><\/p:tgtEl><\/p:cBhvr><\/p:animEffect>/u,
        '<p:animMotion origin="layout" path="M 0 0 L 0.25 0 "><p:cBhvr><p:cTn id="$1" dur="250" fill="hold"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:animMotion>',
      );
      assert.notEqual(next, withMissingSpid);
      return next;
    });

    const imported = importAll(corrupted);
    assert.ok(
      imported.report.diagnostics.length >= 2,
      `Expected >=2 diagnostics, got: ${JSON.stringify(imported.report.diagnostics)}`,
    );
    const missingDiag = imported.report.diagnostics.find(
      (diag) =>
        diag.code === "unsupported-mechanism" && diag.location.objectId === 999,
    );
    assert.ok(
      missingDiag !== undefined && missingDiag.code === "unsupported-mechanism",
    );
    assert.equal(missingDiag.disposition, "omitted");
    assert.equal(missingDiag.location.part, "ppt/slides/slide1.xml");
    assert.equal(missingDiag.location.slideIndex, 0);
    assert.ok((missingDiag.cascaded?.length ?? 0) >= 1);

    const motionPathDiag = imported.report.diagnostics.find(
      (diag) =>
        diag.code === "unsupported-mechanism" &&
        diag.localName === "animMotion",
    );
    assert.ok(motionPathDiag !== undefined);
    assert.equal(motionPathDiag.disposition, "omitted");
    assert.equal(motionPathDiag.location.part, "ppt/slides/slide1.xml");
    assert.equal(motionPathDiag.location.slideIndex, 0);

    // Re-emitting the recovered deck preserves all 4 healthy canvas objects and the independent Wipe animation on shape id 5!
    const reemitted = emit(resolveAll(imported.document));
    const vr = await validatePptx(reemitted.pptx);
    assert.deepEqual(vr.errors, []);
    assert.equal(vr.slides.length, 1);
    assert.equal(vr.slides[0]?.objects.length, 4);
    const allAnimR = vr.slides[0]?.animations?.flatMap((s) => s.effects) ?? [];
    assert.equal(
      allAnimR.some((e) => e.preset === "wipe" && e.from === "left"),
      true,
    );

    // Also verify partial recovery when a broken image (<p:pic> missing its media part) is omitted during import
    const imgDeck = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "broken-image-timing" },
          createElement(
            "image" as never,
            {
              id: "broken-pic",
              src: png1x1,
              style: {
                position: "absolute",
                left: 40,
                top: 40,
                width: 80,
                height: 80,
              },
              enter: fade({ duration: 200 }),
            } as never,
          ),
          createElement(Shape, {
            id: "survivor-shape",
            enter: zoom({ duration: 250 }),
            style: {
              position: "absolute",
              left: 160,
              top: 40,
              width: 100,
              height: 80,
              background: "#10B981",
            },
          }),
        ),
      ),
    );
    const baseImgPptx = emit(resolveAll(imgDeck)).pptx;
    const mediaPart = Object.keys(unzipSync(baseImgPptx)).find((p) =>
      p.startsWith("ppt/media/"),
    )!;
    const brokenImgPptx = removePart(baseImgPptx, mediaPart);
    const importedBroken = importAll(brokenImgPptx);
    const brokenDiag = importedBroken.report.diagnostics.find(
      (diag) =>
        diag.code === "unsupported-mechanism" &&
        diag.mechanismHint === "omitted-target:corrupt",
    );
    assert.ok(brokenDiag !== undefined);
    assert.equal(brokenDiag.disposition, "omitted");
    assert.equal(brokenDiag.location.part, "ppt/slides/slide1.xml");
    assert.equal(brokenDiag.location.slideIndex, 0);
    const reemittedBroken = emit(resolveAll(importedBroken.document));
    const vb = await validatePptx(reemittedBroken.pptx);
    assert.deepEqual(vb.errors, []);
    assert.equal(vb.slides.length, 1);
    assert.equal(vb.slides[0]?.objects.length, 1);
    assert.equal(
      vb.slides[0]?.animations?.[0]?.effects.some((e) => e.preset === "zoom"),
      true,
    );
  });

  it("preserves visual spTree z-order when a referenced MotionScope spans non-contiguous objects", async () => {
    const scopeS = createMotionScopeHandle();
    const afterS = createMotionScopeHandle();
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360, style: { fontFamily: jostFont } },
        createElement(
          Slide,
          { id: "non-contiguous-z" },
          createElement(
            MotionScope,
            { scope: scopeS },
            createElement(
              Text,
              {
                id: "back-animated",
                enter: fade({ by: "word", intervalMs: 50, duration: 200 }),
                style: {
                  position: "absolute",
                  left: 40,
                  top: 40,
                  width: 140,
                  height: 80,
                  fontFamily: jostFont,
                  fontSize: 14,
                  zIndex: 1,
                },
              },
              "First word group",
            ),
            createElement(Shape, {
              id: "front-animated",
              enter: zoom({ duration: 200 }),
              style: {
                position: "absolute",
                left: 100,
                top: 60,
                width: 120,
                height: 80,
                zIndex: 3,
                background: "#2563EB",
              },
            }),
          ),
          createElement(Shape, {
            id: "middle-unanimated",
            style: {
              position: "absolute",
              left: 70,
              top: 50,
              width: 120,
              height: 80,
              zIndex: 2,
              background: "#E2E8F0",
            },
          }),
          createElement(
            MotionScope,
            { scope: afterS, startAt: after(scopeS.end, 30) },
            createElement(Shape, {
              id: "after-shape",
              enter: wipe({ from: "left", duration: 200 }),
              style: {
                position: "absolute",
                left: 260,
                top: 60,
                width: 100,
                height: 60,
                zIndex: 4,
                background: "#10B981",
              },
            }),
          ),
        ),
      ),
    );

    const p1 = emit(resolveAll(doc)).pptx;
    const imported = importAll(p1);
    assert.deepEqual(imported.report.diagnostics, []);
    const p2 = emit(resolveAll(imported.document)).pptx;
    const v2 = await validatePptx(p2);
    assert.deepEqual(v2.errors, []);
    const s2 = v2.slides[0]!;
    // spTree stacking order preserved: Text (x=40) -> Slate Shape (x=70) -> Blue Shape (x=100) -> Green Shape (x=260)
    assert.equal(s2.objects[0]?.text?.text, "First word group");
    assert.deepEqual(
      s2.objects.map((o) => o.boxEmu.x),
      [40 * 9525, 70 * 9525, 100 * 9525, 260 * 9525],
    );
    const anims = s2.animations?.flatMap((st) => st.effects) ?? [];
    const afterEffect = anims.find((e) => e.targetId === s2.objects[3]!.id);
    assert.ok(afterEffect?.startAfter !== undefined);
  });

  it("imports exit animations, appear, multi-emphasis arrays, and cross-step same-target effects folded back to onStep(n, ...)", async () => {
    const steppedFade = onStep(1, fade({ duration: 200 }));
    const steppedSpin = onStep(2, repeat(spin({ duration: 300 }), 2));
    const steppedFadeOut = onStep(3, fadeOut({ duration: 250 }));
    assert.equal(steppedEffectSchema.safeParse(steppedFade).success, true);
    assert.equal(
      entranceEffectDeclarationSchema.safeParse(steppedFade).success,
      true,
    );
    assert.equal(
      exitEffectDeclarationSchema.safeParse(steppedFade).success,
      false,
    );
    assert.equal(
      emphasisEffectDeclarationSchema.safeParse(steppedSpin).success,
      true,
    );
    assert.equal(
      exitEffectDeclarationSchema.safeParse(steppedFadeOut).success,
      true,
    );
    assert.equal(
      steppedEffectSchema.safeParse(onStep(0, fade())).success,
      false,
    );
    assert.equal(
      steppedEffectSchema.safeParse(onStep(1.5, fade())).success,
      false,
    );

    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "exit-onstep-import" }).append(
        step({ n: 1 }).append(
          shape({
            id: "lifecycle-card",
            enter: appear(),
            emphasis: [
              spin({ duration: 300 }),
              fillColor({ from: "#2563EB", to: "#10B981", duration: 200 }),
            ],
            exit: onStep(3, fadeOut({ duration: 250 })),
            style: {
              position: "absolute",
              left: 30,
              top: 30,
              width: 120,
              height: 60,
              background: "#2563EB",
            },
          }),
        ),
        step({ n: 2 }).append(
          shape({
            id: "step2-card",
            enter: zoom({ duration: 200 }),
            exit: collapse({ duration: 180 }),
            style: {
              position: "absolute",
              left: 180,
              top: 30,
              width: 120,
              height: 60,
              background: "#F59E0B",
            },
          }),
        ),
      ),
    );

    const built1 = emit(resolveAll(doc));
    const pptxBytes = new Uint8Array(built1.pptx);
    const imported = importAll(pptxBytes);
    pptxBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const built2 = emit(resolveAll(imported.document));
    const validated2 = await validatePptx(built2.pptx);
    assert.deepEqual(validated2.errors, []);

    const s2 = validated2.slides[0]!;
    assert.equal(s2.animations?.length, 3);
    assert.equal(s2.animations?.[0]?.effects.length, 3); // appear + spin + fillColor in step 1
    assert.deepEqual(
      s2.animations?.[1]?.effects.map((e) => e.preset),
      ["zoom", "collapse"],
    ); // zoom + collapse in step 2
    assert.equal(s2.animations?.[2]?.effects.length, 1); // fadeOut on lifecycle-card in step 3
    assert.equal(s2.animations?.[2]?.effects[0]?.kind, "exit");
    assert.equal(s2.animations?.[2]?.effects[0]?.preset, "fadeOut");
    assert.equal(
      s2.animations?.[2]?.effects[0]?.targetId,
      s2.animations?.[0]?.effects[0]?.targetId,
    );
  });

  it("roundtrips leaf startAt and publishEnd cues without diagnostics", async () => {
    const cues = createMotionScopeHandle();
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "leaf-cue-roundtrip" },
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "first",
              enter: fade({ duration: 200 }),
              publishEnd: cues.cue("ready"),
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 80,
                height: 40,
                background: "#2563EB",
              },
            }),
            createElement(Shape, {
              id: "second",
              enter: fade({ duration: 150 }),
              startAt: after(cues.cue("ready"), 45),
              style: {
                position: "absolute",
                left: 120,
                top: 20,
                width: 80,
                height: 40,
                background: "#10B981",
              },
            }),
          ),
        ),
      ),
    );

    const built1 = emit(resolveAll(doc));
    const pptxBytes = new Uint8Array(built1.pptx);
    const imported = importAll(pptxBytes);
    pptxBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const built2 = emit(resolveAll(imported.document));
    const validated2 = await validatePptx(built2.pptx);
    assert.deepEqual(validated2.errors, []);
    const effs = validated2.slides[0]?.animations?.[0]?.effects ?? [];
    assert.equal(effs.length, 2);
    assert.equal(effs[0]?.delayMs, 0);
    assert.equal(effs[1]?.delayMs, 245);
  });

  it("folds single-paragraph <p:pRg> timelines back onto <P> nodes inside Text and Shape without breaking by:'paragraph'", async () => {
    const doc = deck({ width: 720, height: 405 }).append(
      slide({ id: "p-import-slide" }).append(
        text(
          [
            paragraph("Line 0", {
              enter: fade({ duration: 200 }),
              exit: onStep(3, fadeOut({ duration: 150 })),
            }),
            step({ n: 2 }).append(
              paragraph("Line 1", {
                enter: wipe({ from: "left", duration: 160 }),
              }),
            ),
          ],
          {
            id: "per-p-text",
            style: {
              position: "absolute",
              left: 30,
              top: 30,
              width: 280,
              height: 140,
            },
          },
        ),
        shape(
          [
            step({ n: 1 }).append(paragraph("Shape P0", { enter: appear() })),
            step({ n: 2 }).append(
              paragraph("Shape P1", { enter: fade({ duration: 180 }) }),
            ),
          ],
          {
            id: "per-p-shape",
            style: {
              position: "absolute",
              left: 340,
              top: 30,
              width: 240,
              height: 120,
              background: "#1E293B",
            },
          },
        ),
        step({ n: 1 }).append(
          text([paragraph("All A"), paragraph("All B")], {
            id: "whole-by-p-text",
            enter: fade({ by: "paragraph", duration: 200 }),
            style: {
              position: "absolute",
              left: 30,
              top: 200,
              width: 280,
              height: 90,
            },
          }),
        ),
      ),
    );

    const built1 = emit(resolveAll(doc));
    const pptxBytes = new Uint8Array(built1.pptx);
    const imported = importAll(pptxBytes);
    pptxBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const built2 = emit(resolveAll(imported.document));
    const validated2 = await validatePptx(built2.pptx);
    assert.deepEqual(validated2.errors, []);

    const slide2 = validated2.slides[0]!;
    // Auto (Line 0 enter) + Step 1 (Shape P0 + whole-by-p-text x2) + Step 2 (Line 1 + Shape P1) + Step 3 (Line 0 exit)
    assert.equal(slide2.animations?.length, 4);
    assert.equal(slide2.animations?.[0]?.trigger, "auto");
    assert.deepEqual(slide2.animations?.[0]?.effects[0]?.targetRange, {
      type: "paragraph",
      start: 0,
      end: 0,
    });
    assert.equal(slide2.animations?.[3]?.effects[0]?.kind, "exit");
    assert.deepEqual(slide2.animations?.[3]?.effects[0]?.targetRange, {
      type: "paragraph",
      start: 0,
      end: 0,
    });
  });
});
