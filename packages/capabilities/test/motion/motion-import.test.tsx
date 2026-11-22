import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  deck,
  defineSlideLayout,
  font,
  shape,
  slide,
  slideLayout,
  slideMaster,
  paragraph,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import {
  after,
  breathe,
  createMotionScopeHandle,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  motionCapability,
  repeat,
  spin,
  spinIn,
  stretch,
  wipe,
  zoom,
} from "@pptxdk/capabilities/motion";
import {
  BreatheEffectSchema,
  BreatheOptionsSchema,
  EasingSchema,
  EffectOptionsSchema,
  EffectPresetSchema,
  EffectSchema,
  EmphasisEffectDeclarationSchema,
  EmphasisEffectSchema,
  EntranceEffectDeclarationSchema,
  EntranceEffectSchema,
  FillColorOptionsSchema,
  MotionDirectionSchema,
  MotionGranularitySchema,
  MotionHexColorSchema,
  MotionMomentOffsetSchema,
  MotionMomentSchema,
  MotionPropsSchema,
  MotionScopeAuthorPropsSchema,
  MotionScopeHandleSchema,
  MotionScopeModeSchema,
  MotionStepAuthorPropsSchema,
  MotionTargetRangeSchema,
  MotionTimingSchema,
  OwnerMotionSchema,
  RepeatCountSchema,
  RepeatedEffectSchema,
  SpinOptionsSchema,
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
  MorphKeySchema,
  MorphPropsSchema,
  parseNativeMorphToken,
} from "../../src/shared/morph-api.ts";

describe("Motion Zod v4 schemas & Morph handle identity", () => {
  it("validates all object effect presets, triggers, durations, delays, easing, repeat counts, and target ranges", () => {
    assert.equal(MotionDirectionSchema.parse("left"), "left");
    assert.equal(EffectPresetSchema.parse("fillColor"), "fillColor");
    assert.equal(MotionScopeModeSchema.parse("stagger"), "stagger");
    assert.equal(
      MotionGranularitySchema.parse("seriesElements"),
      "seriesElements",
    );
    assert.equal(RepeatCountSchema.parse("forever"), "forever");
    assert.equal(RepeatCountSchema.parse(3), 3);
    assert.equal(RepeatCountSchema.safeParse(0).success, false);

    assert.deepEqual(EasingSchema.parse({ accelerate: 0.3, decelerate: 0.7 }), {
      accelerate: 0.3,
      decelerate: 0.7,
    });
    assert.equal(
      EasingSchema.safeParse({ accelerate: 0.6, decelerate: 0.5 }).success,
      false,
    );

    assert.deepEqual(
      MotionTargetRangeSchema.parse({ type: "paragraph", start: 0, end: 2 }),
      { type: "paragraph", start: 0, end: 2 },
    );
    assert.deepEqual(
      MotionTargetRangeSchema.parse({
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
      assert.deepEqual(EntranceEffectSchema.parse(eff), eff);
      assert.deepEqual(EntranceEffectDeclarationSchema.parse(eff), eff);
      assert.deepEqual(EffectSchema.parse(eff), eff);
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
    assert.deepEqual(EmphasisEffectSchema.parse(spinEff), spinEff);
    assert.deepEqual(BreatheEffectSchema.parse(breatheEff), breatheEff);
    assert.deepEqual(EmphasisEffectSchema.parse(fillEff), fillEff);
    assert.equal(MotionHexColorSchema.parse("#164e3d"), "#164E3D");

    const repeatedSpin = repeat(spinEff, 2);
    const repeatedBreathe = repeat(breatheEff, "forever");
    assert.deepEqual(RepeatedEffectSchema.parse(repeatedSpin), repeatedSpin);
    assert.deepEqual(
      EmphasisEffectDeclarationSchema.parse(repeatedBreathe),
      repeatedBreathe,
    );

    assert.equal(
      EffectOptionsSchema.safeParse({ duration: 500 }).success,
      true,
    );
    assert.equal(SpinOptionsSchema.safeParse({ from: 180 }).success, true);
    assert.equal(
      FillColorOptionsSchema.safeParse({ from: "#000000", to: "#FFFFFF" })
        .success,
      true,
    );
    assert.equal(
      BreatheOptionsSchema.safeParse({
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

    assert.equal(MotionScopeHandleSchema.parse(scope), scope);
    assert.deepEqual(MotionMomentSchema.parse(cue), cue);
    assert.deepEqual(MotionMomentOffsetSchema.parse(delayed), delayed);
    assert.deepEqual(MotionTimingSchema.parse(delayed), delayed);
    assert.equal(
      MotionScopeHandleSchema.safeParse({ type: "other" }).success,
      false,
    );

    assert.equal(
      MotionScopeAuthorPropsSchema.parse({
        scope,
        mode: "stagger",
        each: 100,
        startAt: delayed,
        publishStart: scope.cue("begin"),
        publishEnd: cue,
      }).scope,
      scope,
    );
    assert.equal(MotionStepAuthorPropsSchema.parse({ n: 2 }).n, 2);
    assert.equal(
      MotionPropsSchema.safeParse({
        enter: [fade(), fly({ from: "left" })],
        emphasis: spin(),
        motion: true,
      }).success,
      true,
    );

    assert.equal(
      OwnerMotionSchema.safeParse({
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
    assert.equal(MorphKeySchema.parse(hero), hero);
    assert.equal(MorphPropsSchema.parse({ morphKey: hero }).morphKey, hero);
    assert.equal(MorphKeySchema.safeParse({ label: "hero" }).success, false);
    assert.equal(parseNativeMorphToken("!!hero"), "hero");
    assert.equal(parseNativeMorphToken("!!"), undefined);
    assert.equal(parseNativeMorphToken("Hero Card"), undefined);
  });
});

const jostBook = readFileSync(
  new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
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
      "motion-choreography",
    ]);
    for (const mechanism of motionCapability.mechanisms ?? []) {
      assert.equal(mechanism.role, "show");
      assert.equal(typeof mechanism.emit, "function");
      assert.equal(mechanism.import.target, "slide");
      assert.equal(mechanism.import.slot, "timing");
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
              x: 30,
              y: 24,
              width: 120,
              height: 60,
              fill: "#2563EB",
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
            }),
            createElement(
              Text,
              {
                id: "words-text",
                x: 170,
                y: 24,
                width: 220,
                height: 60,
                style: { fontFamily: jostFont, fontSize: 16 },
                enter: fly({
                  from: "right",
                  by: "word",
                  intervalMs: 60,
                  duration: 200,
                }),
              },
              "Hello motion words",
            ),
          ),
          createElement(
            MotionScope,
            { scope: afterIntro, startAt: after(intro.end, 80) },
            createElement(Shape, {
              id: "breather",
              x: 30,
              y: 110,
              width: 100,
              height: 50,
              fill: "#164E3D",
              enter: spinIn({ duration: 300 }),
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
                x: 170,
                y: 110,
                width: 240,
                height: 90,
                style: { fontFamily: jostFont, fontSize: 14 },
                enter: fade({ by: "paragraph", duration: 250 }),
              },
              createElement(P, null, "First paragraph"),
              createElement(P, null, "Second paragraph"),
            ),
            createElement(
              Group,
              {
                id: "spin-group",
                emphasis: spin({ by: 90, duration: 400 }),
              },
              createElement(Shape, {
                id: "group-box",
                x: 430,
                y: 30,
                width: 80,
                height: 60,
                fill: "#F59E0B",
                enter: zoom({ duration: 220 }),
              }),
            ),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(Chart, {
              id: "sales-chart",
              x: 420,
              y: 120,
              width: 260,
              height: 220,
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
      x: 5,
      y: 5,
      width: 20,
      height: 20,
      fill: "#CBD5E1",
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
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "#2563EB",
        }),
        shape({
          id: "plain-1",
          name: "Shared Display Name",
          x: 240,
          y: 40,
          width: 120,
          height: 80,
          fill: "#10B981",
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
          x: 80,
          y: 120,
          width: 200,
          height: 110,
          fill: "#2563EB",
        }),
        shape({
          id: "plain-2",
          name: "Shared Display Name",
          x: 240,
          y: 40,
          width: 120,
          height: 80,
          fill: "#10B981",
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
        title: { type: "title", x: 40, y: 40, width: 400, height: 60 },
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
          ftr: { type: "ftr", x: 40, y: 310, width: 200, height: 24 },
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
          x: 60,
          y: 120,
          width: 140,
          height: 80,
          fill: "#2563EB",
          enter: zoom({ duration: 200 }),
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
                x: 20,
                y: 20,
                width: 80,
                height: 60,
                style: { fontFamily: jostFont, fontSize: 14 },
                enter: fade({ by: "word", intervalMs: 50, duration: 200 }),
              },
              "Missing target text",
            ),
          ),
          createElement(
            MotionScope,
            { scope: scopeDependent, startAt: after(scopeA.end, 50) },
            createElement(Shape, {
              id: "dependent-shape",
              x: 120,
              y: 20,
              width: 80,
              height: 60,
              fill: "#F59E0B",
              enter: fade({ duration: 200 }),
            }),
          ),
          createElement(Shape, {
            id: "anim-motion-shape",
            x: 220,
            y: 20,
            width: 80,
            height: 60,
            fill: "#EF4444",
            enter: fade({ duration: 250 }),
          }),
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "independent-shape",
              x: 340,
              y: 20,
              width: 100,
              height: 60,
              fill: "#10B981",
              enter: wipe({ from: "left", duration: 300 }),
            }),
          ),
        ),
      ),
    );
    void png1x1;

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
    void removePart;

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
    assert.ok((missingDiag.cascaded?.length ?? 0) >= 1);

    const motionPathDiag = imported.report.diagnostics.find(
      (diag) =>
        diag.code === "unsupported-mechanism" &&
        diag.localName === "animMotion",
    );
    assert.ok(motionPathDiag !== undefined);

    // Re-emitting the recovered deck preserves the independent Wipe animation on shape id 5!
    const reemitted = emit(resolveAll(imported.document));
    const vr = await validatePptx(reemitted.pptx);
    assert.deepEqual(vr.errors, []);
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
              x: 40,
              y: 40,
              width: 80,
              height: 80,
              enter: fade({ duration: 200 }),
            } as never,
          ),
          createElement(Shape, {
            id: "survivor-shape",
            x: 160,
            y: 40,
            width: 100,
            height: 80,
            fill: "#10B981",
            enter: zoom({ duration: 250 }),
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
    assert.ok(
      importedBroken.report.diagnostics.some(
        (diag) =>
          diag.code === "unsupported-mechanism" &&
          diag.mechanismHint === "omitted-target:corrupt",
      ),
    );
    const reemittedBroken = emit(resolveAll(importedBroken.document));
    const vb = await validatePptx(reemittedBroken.pptx);
    assert.deepEqual(vb.errors, []);
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
                x: 40,
                y: 40,
                width: 140,
                height: 80,
                style: { fontFamily: jostFont, fontSize: 14, zIndex: 1 },
                enter: fade({ by: "word", intervalMs: 50, duration: 200 }),
              },
              "First word group",
            ),
            createElement(Shape, {
              id: "front-animated",
              x: 100,
              y: 60,
              width: 120,
              height: 80,
              fill: "#2563EB",
              style: { zIndex: 3 },
              enter: zoom({ duration: 200 }),
            }),
          ),
          createElement(Shape, {
            id: "middle-unanimated",
            x: 70,
            y: 50,
            width: 120,
            height: 80,
            fill: "#E2E8F0",
            style: { zIndex: 2 },
          }),
          createElement(
            MotionScope,
            { scope: afterS, startAt: after(scopeS.end, 30) },
            createElement(Shape, {
              id: "after-shape",
              x: 260,
              y: 60,
              width: 100,
              height: 60,
              fill: "#10B981",
              style: { zIndex: 4 },
              enter: wipe({ from: "left", duration: 200 }),
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
});
