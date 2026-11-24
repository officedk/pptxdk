import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdir, mkdtemp, rename, rm, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import {
  BuildError,
  CodegenError,
  build,
  codegen,
  importPptx,
  type CodegenResult,
} from "pptxdk";
import { previewFromSource, type PreviewResult } from "pptxdk/node";
import {
  Box,
  Chart,
  Deck,
  Group,
  MotionScope,
  P,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Step,
  Text,
  after,
  breathe,
  defineSlideLayout,
  expandDocument,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  font,
  repeat,
  spin,
  spinIn,
  stretch,
  useMotionScope,
  wipe,
  zoom,
} from "pptxdk/react";
import {
  box,
  chart,
  createMotionScopeHandle,
  deck,
  group,
  motionScope,
  paragraph,
  shape,
  slide,
  slideLayout,
  slideMaster,
  step,
  text,
} from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
} from "./helpers/codegen-acceptance.ts";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));

const jostBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBytes }]);

function samplePngPixel(
  png: PNG,
  pxX: number,
  pxY: number,
): readonly [number, number, number, number] {
  const idx = (Math.round(pxY) * png.width + Math.round(pxX)) * 4;
  return [
    png.data[idx]!,
    png.data[idx + 1]!,
    png.data[idx + 2]!,
    png.data[idx + 3]!,
  ];
}

async function previewGeneratedSourceRelocated(
  result: CodegenResult,
): Promise<PreviewResult> {
  const tempRoot = await mkdtemp(
    join(tmpdir(), "pptxdk-codegen-motion-preview-"),
  );
  try {
    const consumerRoot = join(tempRoot, "consumer");
    const nodeModulesDir = join(consumerRoot, "node_modules");
    await mkdir(nodeModulesDir, { recursive: true });
    await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
    await symlink(
      join(packageRoot, "node_modules/react"),
      join(nodeModulesDir, "react"),
      "dir",
    );

    const originDir = join(tempRoot, "origin");
    await writeCodegenFiles(originDir, result.files);
    const relocatedDir = join(consumerRoot, "relocated");
    await rename(originDir, relocatedDir);

    return await previewFromSource(join(relocatedDir, result.entry));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

const animatedLayout = defineSlideLayout({
  id: "motion-layout",
  name: "Motion Layout",
  type: "title",
  showMasterPlaceholderAnimations: false,
  placeholders: {
    title: { type: "title", x: 32, y: 16, width: 320, height: 36 },
  },
  motion: {
    steps: [
      {
        trigger: "click",
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

function createMotionAcceptanceVanillaDeck() {
  const leadScope = createMotionScopeHandle();
  const followerScope = createMotionScopeHandle();
  const staggerScope = createMotionScopeHandle();
  const afterStaggerScope = createMotionScopeHandle();
  const wordSourceScope = createMotionScopeHandle();
  const afterWordScope = createMotionScopeHandle();
  const readyCue = leadScope.cue("ready");
  const staggerDoneCue = staggerScope.cue("done");
  const wordDoneCue = wordSourceScope.cue("wordDone");

  return deck({ width: 640, height: 360 }).append(
    slideMaster({
      id: "motion-master",
      name: "Motion Master",
      placeholders: {
        ftr: { type: "ftr", x: 32, y: 316, width: 200, height: 24 },
      },
      motion: {
        steps: [
          {
            trigger: "auto",
            effects: [
              {
                target: { placeholder: "ftr" },
                effect: wipe({ from: "left", duration: 220 }),
              },
            ],
          },
        ],
      },
    }).append(slideLayout(animatedLayout)),
    slide({
      id: "choreo-steps",
      layout: animatedLayout,
      showMasterPlaceholderAnimations: true,
      background: "#F8FAFC",
    }).append(
      step({ n: 1 }).append(
        motionScope(leadScope, {
          mode: "sequence",
          publishEnd: readyCue,
        }).append(
          shape({
            id: "hero-badge",
            x: 32,
            y: 64,
            width: 110,
            height: 60,
            fill: "#2563EB",
            enter: fade({ duration: 240 }),
            emphasis: fillColor({
              from: "#2563EB",
              to: "#10B981",
              duration: 300,
            }),
          }),
          text("Sequence Title", {
            id: "hero-title",
            x: 160,
            y: 64,
            width: 190,
            height: 60,
            enter: fly({ from: "left", duration: 220 }),
            style: {
              fontFamily: jostFont,
              fontSize: 15,
              color: "#0F172A",
              backgroundColor: "#EFF6FF",
              padding: 8,
            },
          }),
        ),
        motionScope(followerScope, {
          startAt: after(readyCue, 80),
        }).append(
          shape({
            id: "follower-card",
            x: 368,
            y: 64,
            width: 140,
            height: 60,
            fill: "#10B981",
            enter: wipe({ from: "top", duration: 200 }),
          }),
        ),
      ),
      step({ n: 2 }).append(
        text([paragraph("Paragraph one"), paragraph("Paragraph two")], {
          id: "para-bullets",
          x: 32,
          y: 144,
          width: 240,
          height: 80,
          enter: floatUp({ by: "paragraph", duration: 180 }),
          style: {
            fontFamily: jostFont,
            fontSize: 14,
            color: "#0F172A",
            backgroundColor: "#FFFFFF",
            padding: 8,
          },
        }),
        group({
          id: "spin-group",
          emphasis: spin({ by: 90, duration: 360 }),
        }).append(
          shape({
            id: "group-chip",
            x: 296,
            y: 144,
            width: 90,
            height: 60,
            fill: "#F59E0B",
            enter: zoom({ duration: 190 }),
          }),
        ),
      ),
      step().append(
        shape({
          id: "auto-step-shape",
          x: 410,
          y: 144,
          width: 120,
          height: 60,
          fill: "#8B5CF6",
          enter: [floatDown({ duration: 160 }), stretch({ duration: 180 })],
        }),
        chart({
          id: "animated-chart",
          x: 32,
          y: 234,
          width: 240,
          height: 100,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [
                { name: "Rev", values: [10, 20] },
                { name: "Cost", values: [4, 9] },
              ],
            },
          ],
          enter: fade({ by: "series", duration: 180 }),
          style: { fontFamily: jostFont },
        }),
      ),
    ),
    slide({
      id: "presets-and-auto",
      background: "#FFFFFF",
    }).append(
      motionScope(staggerScope, {
        mode: "stagger",
        each: 100,
        delay: 30,
        publishEnd: staggerDoneCue,
      }).append(
        shape({
          id: "stagger-a",
          x: 32,
          y: 40,
          width: 120,
          height: 64,
          fill: "#0EA5E9",
          enter: spinIn({ duration: 210 }),
        }),
        shape({
          id: "stagger-b",
          x: 172,
          y: 40,
          width: 120,
          height: 64,
          fill: "#14B8A6",
          enter: fade({ duration: 200 }),
        }),
      ),
      motionScope(afterStaggerScope, {
        startAt: after(staggerDoneCue, 50),
      }).append(
        motionScope(wordSourceScope, { publishEnd: wordDoneCue }).append(
          text("Word by word", {
            id: "words-text",
            x: 312,
            y: 40,
            width: 160,
            height: 64,
            enter: wipe({
              from: "right",
              by: "word",
              intervalMs: 60,
              duration: 200,
            }),
            style: { fontFamily: jostFont, fontSize: 15, color: "#0F172A" },
          }),
        ),
      ),
      motionScope(afterWordScope, {
        startAt: after(wordDoneCue, 50),
      }).append(
        shape({
          id: "after-word-card",
          x: 492,
          y: 40,
          width: 100,
          height: 64,
          fill: "#9333EA",
          enter: fade({ duration: 150 }),
        }),
      ),
      text("ABC", {
        id: "letter-text",
        x: 536,
        y: 140,
        width: 72,
        height: 48,
        enter: fade({ by: "letter", intervalMs: 40, duration: 120 }),
        style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
      }),
      box({
        id: "box-defaults",
        enter: floatUp({ duration: 180 }),
        style: {
          position: "absolute",
          left: 32,
          top: 140,
          width: 300,
          display: "flex",
          flexDirection: "row",
          gap: 16,
        },
      }).append(
        shape({
          id: "box-inherited",
          width: 120,
          height: 64,
          fill: "#DC2626",
        }),
        shape({
          id: "box-muted",
          width: 120,
          height: 64,
          fill: "#64748B",
          motion: false,
          enter: fade(),
        }),
      ),
      shape({
        id: "breathe-lamp",
        x: 360,
        y: 140,
        width: 160,
        height: 64,
        fill: "#164E3D",
        emphasis: repeat(
          breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
          "forever",
        ),
      }),
    ),
  );
}

async function createMotionAcceptanceReactDeck() {
  function MotionAcceptanceDeck() {
    const leadScope = useMotionScope();
    const followerScope = useMotionScope();
    const staggerScope = useMotionScope();
    const afterStaggerScope = useMotionScope();
    const wordSourceScope = useMotionScope();
    const afterWordScope = useMotionScope();
    const readyCue = leadScope.cue("ready");
    const staggerDoneCue = staggerScope.cue("done");
    const wordDoneCue = wordSourceScope.cue("wordDone");

    return createElement(
      Deck,
      { width: 640, height: 360 },
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
                trigger: "auto",
                effects: [
                  {
                    target: { placeholder: "ftr" },
                    effect: wipe({ from: "left", duration: 220 }),
                  },
                ],
              },
            ],
          },
        },
        createElement(SlideLayout, { layout: animatedLayout }),
      ),
      createElement(
        Slide,
        {
          id: "choreo-steps",
          layout: animatedLayout,
          showMasterPlaceholderAnimations: true,
          background: "#F8FAFC",
        },
        createElement(
          Step,
          { n: 1 },
          createElement(
            MotionScope,
            { scope: leadScope, mode: "sequence", publishEnd: readyCue },
            createElement(Shape, {
              id: "hero-badge",
              x: 32,
              y: 64,
              width: 110,
              height: 60,
              fill: "#2563EB",
              enter: fade({ duration: 240 }),
              emphasis: fillColor({
                from: "#2563EB",
                to: "#10B981",
                duration: 300,
              }),
            }),
            createElement(
              Text,
              {
                id: "hero-title",
                x: 160,
                y: 64,
                width: 190,
                height: 60,
                enter: fly({ from: "left", duration: 220 }),
                style: {
                  fontFamily: jostFont,
                  fontSize: 15,
                  color: "#0F172A",
                  backgroundColor: "#EFF6FF",
                  padding: 8,
                },
              },
              "Sequence Title",
            ),
          ),
          createElement(
            MotionScope,
            { scope: followerScope, startAt: after(readyCue, 80) },
            createElement(Shape, {
              id: "follower-card",
              x: 368,
              y: 64,
              width: 140,
              height: 60,
              fill: "#10B981",
              enter: wipe({ from: "top", duration: 200 }),
            }),
          ),
        ),
        createElement(
          Step,
          { n: 2 },
          createElement(
            Text,
            {
              id: "para-bullets",
              x: 32,
              y: 144,
              width: 240,
              height: 80,
              enter: floatUp({ by: "paragraph", duration: 180 }),
              style: {
                fontFamily: jostFont,
                fontSize: 14,
                color: "#0F172A",
                backgroundColor: "#FFFFFF",
                padding: 8,
              },
            },
            createElement(P, null, "Paragraph one"),
            createElement(P, null, "Paragraph two"),
          ),
          createElement(
            Group,
            {
              id: "spin-group",
              emphasis: spin({ by: 90, duration: 360 }),
            },
            createElement(Shape, {
              id: "group-chip",
              x: 296,
              y: 144,
              width: 90,
              height: 60,
              fill: "#F59E0B",
              enter: zoom({ duration: 190 }),
            }),
          ),
        ),
        createElement(
          Step,
          null,
          createElement(Shape, {
            id: "auto-step-shape",
            x: 410,
            y: 144,
            width: 120,
            height: 60,
            fill: "#8B5CF6",
            enter: [floatDown({ duration: 160 }), stretch({ duration: 180 })],
          }),
          createElement(Chart, {
            id: "animated-chart",
            x: 32,
            y: 234,
            width: 240,
            height: 100,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [
                  { name: "Rev", values: [10, 20] },
                  { name: "Cost", values: [4, 9] },
                ],
              },
            ],
            enter: fade({ by: "series", duration: 180 }),
            style: { fontFamily: jostFont },
          }),
        ),
      ),
      createElement(
        Slide,
        { id: "presets-and-auto", background: "#FFFFFF" },
        createElement(
          MotionScope,
          {
            scope: staggerScope,
            mode: "stagger",
            each: 100,
            delay: 30,
            publishEnd: staggerDoneCue,
          },
          createElement(Shape, {
            id: "stagger-a",
            x: 32,
            y: 40,
            width: 120,
            height: 64,
            fill: "#0EA5E9",
            enter: spinIn({ duration: 210 }),
          }),
          createElement(Shape, {
            id: "stagger-b",
            x: 172,
            y: 40,
            width: 120,
            height: 64,
            fill: "#14B8A6",
            enter: fade({ duration: 200 }),
          }),
        ),
        createElement(
          MotionScope,
          { scope: afterStaggerScope, startAt: after(staggerDoneCue, 50) },
          createElement(
            MotionScope,
            { scope: wordSourceScope, publishEnd: wordDoneCue },
            createElement(
              Text,
              {
                id: "words-text",
                x: 312,
                y: 40,
                width: 160,
                height: 64,
                enter: wipe({
                  from: "right",
                  by: "word",
                  intervalMs: 60,
                  duration: 200,
                }),
                style: { fontFamily: jostFont, fontSize: 15, color: "#0F172A" },
              },
              "Word by word",
            ),
          ),
        ),
        createElement(
          MotionScope,
          { scope: afterWordScope, startAt: after(wordDoneCue, 50) },
          createElement(Shape, {
            id: "after-word-card",
            x: 492,
            y: 40,
            width: 100,
            height: 64,
            fill: "#9333EA",
            enter: fade({ duration: 150 }),
          }),
        ),
        createElement(
          Text,
          {
            id: "letter-text",
            x: 536,
            y: 140,
            width: 72,
            height: 48,
            enter: fade({ by: "letter", intervalMs: 40, duration: 120 }),
            style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
          },
          "ABC",
        ),
        createElement(
          Box,
          {
            id: "box-defaults",
            enter: floatUp({ duration: 180 }),
            style: {
              position: "absolute",
              left: 32,
              top: 140,
              width: 300,
              display: "flex",
              flexDirection: "row",
              gap: 16,
            },
          },
          createElement(Shape, {
            id: "box-inherited",
            width: 120,
            height: 64,
            fill: "#DC2626",
          }),
          createElement(Shape, {
            id: "box-muted",
            width: 120,
            height: 64,
            fill: "#64748B",
            motion: false,
            enter: fade(),
          }),
        ),
        createElement(Shape, {
          id: "breathe-lamp",
          x: 360,
          y: 140,
          width: 160,
          height: 64,
          fill: "#164E3D",
          emphasis: repeat(
            breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
            "forever",
          ),
        }),
      ),
    );
  }

  return expandDocument(createElement(MotionAcceptanceDeck));
}

describe("pptxdk motion choreography codegen acceptance", () => {
  it("preserves effects, click steps, scopes, cues, granularity, OwnerMotion, and static canvas across direct and imported vanilla/react roundtrips", async () => {
    const vanillaDoc = createMotionAcceptanceVanillaDeck();
    const reactDoc = await createMotionAcceptanceReactDeck();

    const initialBuilt = await build(vanillaDoc, { embedFonts: "full" });
    assert.deepEqual(
      initialBuilt.report.diagnostics.filter((d) => d.severity === "error"),
      [],
    );
    const bytesCopy = new Uint8Array(initialBuilt.pptx);
    const imported = importPptx(bytesCopy);
    bytesCopy.fill(0);
    assert.deepEqual(
      imported.report.diagnostics.filter(
        (diag) => !(diag.code === "unconsumed-xml" && diag.localName === "bg"),
      ),
      [],
    );

    const sources = [
      { label: "direct-vanilla", doc: vanillaDoc },
      { label: "direct-react", doc: reactDoc },
      { label: "imported-pptx", doc: imported.document },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc }) => [
      {
        label: `${label}:vanilla`,
        result: codegen(doc, { target: "vanilla" }),
      },
      { label: `${label}:react`, result: codegen(doc, { target: "react" }) },
    ]);

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      const painted = await previewGeneratedSourceRelocated(item.result);

      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );
      assert.equal(file.slides.length, 2, `${item.label} slide count`);
      assert.equal(painted.pages.length, 2, `${item.label} preview page count`);
      for (const [pageIdx, page] of painted.pages.entries()) {
        assert.deepEqual(
          page.diagnostics,
          [],
          `${item.label} preview page ${pageIdx + 1} diagnostics`,
        );
        assert.ok(page.png instanceof Uint8Array);
      }

      // Master & Layout OwnerMotion + showMasterPlaceholderAnimations
      assert.equal(
        file.masters?.[0]?.animations?.[0]?.trigger,
        "auto",
        `${item.label} master trigger`,
      );
      assert.equal(
        file.masters?.[0]?.animations?.[0]?.effects[0]?.preset,
        "wipe",
        `${item.label} master wipe preset`,
      );
      assert.equal(
        file.masters?.[0]?.layouts[0]?.showMasterPlaceholderAnimations,
        false,
        `${item.label} layout showMasterPhAnim`,
      );
      assert.equal(
        file.masters?.[0]?.layouts[0]?.animations?.[0]?.trigger,
        "click",
        `${item.label} layout trigger`,
      );
      assert.equal(
        file.masters?.[0]?.layouts[0]?.animations?.[0]?.effects[0]?.preset,
        "fade",
        `${item.label} layout fade preset`,
      );

      // Slide 1 timing tree: 3 click steps, cross-object startAfter, paragraph granularity, Group spin
      const s1 = file.slides[0]!;
      assert.equal(
        s1.showMasterPlaceholderAnimations,
        true,
        `${item.label} s1 showMasterPhAnim`,
      );
      assert.equal(s1.animations?.length, 3, `${item.label} s1 step count`);
      assert.deepEqual(
        s1.animations?.map((st) => st.trigger),
        ["click", "click", "click"],
        `${item.label} s1 click step triggers`,
      );

      const s1Effects = s1.animations.flatMap((st) => st.effects);
      const step1Effects = s1.animations[0]!.effects;
      assert.ok(
        step1Effects.some((e) => e.preset === "fade" && e.durationMs === 240),
        `${item.label} s1 fade`,
      );
      assert.ok(
        step1Effects.some(
          (e) =>
            e.preset === "fillColor" &&
            e.fromColor === "#2563EB" &&
            e.toColor === "#10B981",
        ),
        `${item.label} s1 fillColor`,
      );
      assert.ok(
        step1Effects.some((e) => e.preset === "fly" && e.from === "left"),
        `${item.label} s1 fly`,
      );
      const followerWipe = step1Effects.find(
        (e) => e.preset === "wipe" && e.from === "top",
      );
      assert.ok(followerWipe !== undefined, `${item.label} s1 follower wipe`);
      assert.equal(
        followerWipe.delayMs,
        600,
        `${item.label} s1 follower delayMs matches sequence end + 80ms`,
      );

      const step2Effects = s1.animations[1]!.effects;
      const paraFloatUps = step2Effects.filter(
        (e) => e.preset === "floatUp" && e.targetRange?.type === "paragraph",
      );
      assert.deepEqual(
        paraFloatUps.map((e) => e.targetRange),
        [
          { type: "paragraph", start: 0, end: 0 },
          { type: "paragraph", start: 1, end: 1 },
        ],
        `${item.label} s1 paragraph granularity`,
      );
      assert.ok(
        step2Effects.some((e) => e.preset === "zoom"),
        `${item.label} s1 group child zoom`,
      );
      assert.ok(
        step2Effects.some((e) => e.preset === "spin" && e.by === 90),
        `${item.label} s1 group spin`,
      );

      const step3Effects = s1.animations[2]!.effects;
      assert.ok(
        step3Effects.some((e) => e.preset === "floatDown"),
        `${item.label} s1 step 3 floatDown`,
      );
      assert.ok(
        step3Effects.some((e) => e.preset === "stretch"),
        `${item.label} s1 step 3 stretch`,
      );
      assert.ok(
        step3Effects.some(
          (e) =>
            e.targetRange?.type === "chart" && e.targetRange.step === "series",
        ),
        `${item.label} s1 step 3 chart series granularity`,
      );
      assert.equal(s1Effects.length > 0, true);

      // Slide 2 timing tree: auto stagger + dependent word wipe + Box default floatUp (excluding muted) + forever breathe
      const s2 = file.slides[1]!;
      assert.equal(s2.animations?.length, 1, `${item.label} s2 auto step`);
      assert.equal(s2.animations?.[0]?.trigger, "auto");
      const s2Effects = s2.animations[0].effects;
      assert.ok(
        s2Effects.some((e) => e.preset === "spinIn"),
        `${item.label} s2 spinIn`,
      );
      const wordWipe = s2Effects.find(
        (e) => e.preset === "wipe" && e.by === "word",
      );
      assert.ok(wordWipe !== undefined, `${item.label} s2 word wipe`);
      assert.equal(wordWipe.intervalMs, 60);
      assert.equal(
        wordWipe.delayMs,
        380,
        `${item.label} s2 word wipe starts at stagger end (30 + 100 + 200) + 50ms`,
      );
      const afterWordFade = s2Effects.find(
        (e) => e.preset === "fade" && e.durationMs === 150,
      );
      assert.ok(
        afterWordFade !== undefined,
        `${item.label} s2 after-word fade`,
      );
      assert.ok(
        afterWordFade.startAfter !== undefined,
        `${item.label} s2 after-word startAfter dependency`,
      );
      assert.equal(
        afterWordFade.startAfter.offsetMs,
        item.label.startsWith("imported-pptx:") ? 100 : 50,
      );
      const letterFade = s2Effects.find(
        (e) => e.preset === "fade" && e.by === "letter",
      );
      assert.ok(letterFade !== undefined, `${item.label} s2 letter fade`);
      assert.equal(letterFade.intervalMs, 40);
      assert.ok(
        s2Effects.some(
          (e) => e.preset === "breathe" && e.repeatCount === "forever",
        ),
        `${item.label} s2 forever breathe`,
      );
      // Only box-inherited got floatUp (box-muted had motion={false})
      assert.equal(
        s2Effects.filter((e) => e.preset === "floatUp").length,
        1,
        `${item.label} s2 box-muted excluded from animation`,
      );

      // Verify actual PPTX <p:timing> and <p:bldLst> XML
      const parts = unzipSync(built.pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      const slide2Xml = strFromU8(parts["ppt/slides/slide2.xml"]!);
      assert.match(slide1Xml, /<p:timing>/u);
      assert.match(slide1Xml, /<p:bldP\b[^>]*build="p"/u);
      assert.match(slide1Xml, /<a:bldChart bld="series"/u);
      assert.match(slide2Xml, /repeatCount="indefinite"/u);

      // Static edit canvas PNG keeps enter-animated objects visibly rendered (no animation hiding)
      const png1 = PNG.sync.read(Buffer.from(painted.pages[0]!.png!));
      const png2 = PNG.sync.read(Buffer.from(painted.pages[1]!.png!));
      assert.deepEqual(
        samplePngPixel(png1, 70, 94),
        [0x25, 0x63, 0xeb, 255],
        `${item.label} s1 hero-badge static pixel`,
      );
      assert.deepEqual(
        samplePngPixel(png1, 420, 94),
        [0x10, 0xb9, 0x81, 255],
        `${item.label} s1 follower-card static pixel`,
      );
      assert.deepEqual(
        samplePngPixel(png2, 80, 72),
        [0x0e, 0xa5, 0xe9, 255],
        `${item.label} s2 stagger-a static pixel`,
      );
      assert.deepEqual(
        samplePngPixel(png2, 420, 172),
        [0x16, 0x4e, 0x3d, 255],
        `${item.label} s2 breathe-lamp static pixel`,
      );
    }
  }, 90_000);

  it("recomputes cross-object dependent start times when editing predecessor duration in direct-authored vanilla and react source while preserving independent click groups", async () => {
    const vanillaDoc = createMotionAcceptanceVanillaDeck();
    const reactDoc = await createMotionAcceptanceReactDeck();

    const directCases = [
      {
        label: "direct-edit:vanilla",
        result: codegen(vanillaDoc, { target: "vanilla" }),
      },
      {
        label: "direct-edit:react",
        result: codegen(reactDoc, { target: "react" }),
      },
    ] as const;

    const baselineBuilt = await Promise.all(
      directCases.map((item) => rebuildGeneratedSourceRelocated(item.result)),
    );

    const mutatedCases = directCases.map(({ label, result }) => {
      const source = result.files[result.entry] as string;
      assert.match(source, /\.cue\("ready"\)/u, `${label} cue preserved`);
      assert.match(
        source,
        /after\([^,]+?\.cue\("ready"\), 80\)/u,
        `${label} after(cue, 80) preserved`,
      );
      // Increase predecessor hero-title fly duration from 220ms -> 620ms (+400ms)
      const updated = source.replace(
        /fly\(\{ from: "left", duration: 220 \}\)/u,
        'fly({ from: "left", duration: 620 })',
      );
      assert.notEqual(updated, source, `${label} predecessor duration mutated`);
      return {
        label,
        result: {
          ...result,
          files: {
            ...result.files,
            [result.entry]: updated,
          },
        },
      };
    });

    await verifyGeneratedSourcesTypecheck(mutatedCases);

    for (const [idx, item] of mutatedCases.entries()) {
      const beforeMotion = baselineBuilt[idx]!.built.report.slides[0]
        ?.capabilities?.["motion"] as {
        readonly steps: readonly {
          readonly trigger: string;
          readonly effects: readonly {
            readonly target: { readonly authorId?: string };
            readonly startMs?: number;
          }[];
        }[];
      };
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(file.errors, [], `${item.label} validation errors`);
      const afterMotion = built.report.slides[0]?.capabilities?.["motion"] as {
        readonly steps: readonly {
          readonly trigger: string;
          readonly effects: readonly {
            readonly target: { readonly authorId?: string };
            readonly startMs?: number;
          }[];
        }[];
      };

      const beforeFollowerStart = beforeMotion.steps[0]?.effects.find(
        (e) => e.target.authorId === "follower-card",
      )?.startMs;
      const afterFollowerStart = afterMotion.steps[0]?.effects.find(
        (e) => e.target.authorId === "follower-card",
      )?.startMs;
      assert.equal(beforeFollowerStart, 300 + 220 + 80);
      assert.equal(afterFollowerStart, 300 + 620 + 80);

      // Independent click steps 2 and 3 remain separate click trigger groups starting at 0ms.
      assert.deepEqual(
        file.slides[0]?.animations?.map((st) => st.trigger),
        ["click", "click", "click"],
        `${item.label} click groups preserved`,
      );
      assert.equal(
        afterMotion.steps[1]?.effects[0]?.startMs,
        0,
        `${item.label} step 2 stays independent`,
      );
      assert.equal(
        afterMotion.steps[2]?.effects[0]?.startMs,
        0,
        `${item.label} step 3 stays independent`,
      );
    }
  }, 60_000);

  it("preserves missing-cue and cross-step-cue diagnostics with exact document paths through codegen + buildFromSource, and rejects unexpressible motion via CodegenError", async () => {
    const missingSource = createMotionScopeHandle();
    const missingFollower = createMotionScopeHandle();
    const missingCue = missingSource.cue("never-mounted");
    const missingCueDeck = deck({ width: 320, height: 180 }).append(
      slide({ id: "missing" }).append(
        motionScope(missingFollower, { startAt: after(missingCue, 20) }).append(
          shape({
            id: "orphan",
            x: 10,
            y: 10,
            width: 40,
            height: 20,
            enter: fade(),
          }),
        ),
      ),
    );

    const crossA = createMotionScopeHandle();
    const crossB = createMotionScopeHandle();
    const crossCue = crossA.cue("done");
    const crossStepDeck = deck({ width: 320, height: 180 }).append(
      slide({ id: "cross-step" }).append(
        step({ n: 1 }).append(
          motionScope(crossA, { publishEnd: crossCue }).append(
            shape({
              id: "s1",
              x: 10,
              y: 10,
              width: 40,
              height: 20,
              enter: fade(),
            }),
          ),
        ),
        step({ n: 2 }).append(
          motionScope(crossB, { startAt: after(crossCue, 10) }).append(
            shape({
              id: "s2",
              x: 10,
              y: 40,
              width: 40,
              height: 20,
              enter: fade(),
            }),
          ),
        ),
      ),
    );

    const rejectionCases = [
      {
        label: "missing-cue:vanilla",
        result: codegen(missingCueDeck, { target: "vanilla" }),
        expectedCodes: ["missing-cue", "missing-scope"],
        expectedPath: "deck/slide[missing]/motion-scope[1]",
      },
      {
        label: "missing-cue:react",
        result: codegen(missingCueDeck, { target: "react" }),
        expectedCodes: ["missing-cue", "missing-scope"],
        expectedPath: "deck/slide[missing]/motion-scope[1]",
      },
      {
        label: "cross-step:vanilla",
        result: codegen(crossStepDeck, { target: "vanilla" }),
        expectedCodes: ["cross-step-motion"],
        expectedPath: "deck/slide[cross-step]/motion-step[2]/motion-scope[1]",
      },
      {
        label: "cross-step:react",
        result: codegen(crossStepDeck, { target: "react" }),
        expectedCodes: ["cross-step-motion"],
        expectedPath: "deck/slide[cross-step]/motion-step[2]/motion-scope[1]",
      },
    ] as const;

    await verifyGeneratedSourcesTypecheck(
      rejectionCases.map(({ label, result }) => ({ label, result })),
    );

    for (const item of rejectionCases) {
      await assert.rejects(
        () =>
          rebuildGeneratedSourceRelocated(item.result, {
            verifyVanillaIsolation: false,
          }),
        (error: unknown) => {
          assert.ok(
            error instanceof BuildError,
            `${item.label}: expected BuildError`,
          );
          const diag = error.diagnostics.find((d) =>
            (item.expectedCodes as readonly string[]).includes(d.code),
          );
          assert.ok(
            diag !== undefined,
            `${item.label}: unexpected diagnostics ${JSON.stringify(error.diagnostics)}`,
          );
          assert.equal(
            diag.path,
            item.expectedPath,
            `${item.label}: diagnostic path`,
          );
          return true;
        },
      );
    }

    const unexpressibleDeck = deck({ width: 320, height: 180 }).append(
      slide({ id: "bad-codegen" }).append(
        shape({
          id: "bad-enter",
          x: 10,
          y: 10,
          width: 40,
          height: 20,
          enter: spin() as unknown as ReturnType<typeof fade>,
        }),
      ),
    );
    assert.throws(
      () => codegen(unexpressibleDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[bad-codegen]/shape[bad-enter]@enter",
    );
  }, 60_000);
});
