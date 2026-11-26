import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  after,
  breathe,
  chart,
  createMotionScopeHandle,
  deck,
  defineSlideLayout,
  fade,
  fillColor,
  floatDown,
  floatUp,
  fly,
  motionScope,
  repeat,
  shape,
  slide,
  slideLayout,
  slideMaster,
  spin,
  spinIn,
  step,
  stretch,
  text,
  wipe,
  zoom,
} from "../../document.ts";
import { MotionScope, Step } from "../motion-elements.ts";

describe("motion capability codegen", () => {
  it("generates vanilla and react source for motion-step (filtering stepIdentity) and motion-scope choreography", () => {
    const lead = createMotionScopeHandle();
    const afterLeadEnd = createMotionScopeHandle();
    const stagger = createMotionScopeHandle();
    const follower = createMotionScopeHandle();
    const leadReady = lead.cue("ready");
    const staggerBegin = stagger.cue("begin");
    const staggerDone = stagger.cue("done");

    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "choreo" }).append(
        step().append(
          motionScope(lead, { mode: "sequence", publishEnd: leadReady }).append(
            shape({
              id: "a",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 80,
                height: 40,
              },
            }),
          ),
          motionScope(afterLeadEnd, { startAt: lead.end }).append(
            shape({
              id: "a-end",
              style: {
                position: "absolute",
                left: 10,
                top: 60,
                width: 80,
                height: 40,
              },
            }),
          ),
        ),
        step({ n: 2 }).append(
          motionScope(stagger, {
            mode: "stagger",
            each: 120,
            delay: 40,
            startAt: after(leadReady, 80),
            publishStart: staggerBegin,
            publishEnd: staggerDone,
            motion: true,
          }).append(
            shape({
              id: "b",
              style: {
                position: "absolute",
                left: 100,
                top: 10,
                width: 80,
                height: 40,
              },
            }),
            shape({
              id: "c",
              style: {
                position: "absolute",
                left: 190,
                top: 10,
                width: 80,
                height: 40,
              },
            }),
          ),
          motionScope(follower, {
            startAt: stagger.start,
            motion: false,
          }).append(
            shape({
              id: "d",
              style: {
                position: "absolute",
                left: 280,
                top: 10,
                width: 80,
                height: 40,
              },
            }),
          ),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(
      vanillaSource,
      /import \{[\s\S]*\bafter\b[\s\S]*\bcreateMotionScopeHandle\b[\s\S]*\bmotionScope\b[\s\S]*\bstep\b[\s\S]*\} from "@officedk\/pptxdk\/vanilla";/u,
    );
    assert.doesNotMatch(vanillaSource, /stepIdentity/u);
    assert.match(
      vanillaSource,
      /const motionScopeHandle = createMotionScopeHandle\(\);/u,
    );
    assert.match(vanillaSource, /step\(\)\.append\(/u);
    assert.match(vanillaSource, /step\(\{ n: 2 \}\)\.append\(/u);
    assert.match(
      vanillaSource,
      /motionScope\(\s*motionScopeHandle,\s*\{ mode: "sequence", publishEnd: motionScopeHandle\.cue\("ready"\) \},/u,
    );
    assert.match(
      vanillaSource,
      /startAt: after\(motionScopeHandle\.cue\("ready"\), 80\)/u,
    );
    assert.match(
      vanillaSource,
      /publishStart: motionScopeHandle3\.cue\("begin"\)/u,
    );
    assert.match(
      vanillaSource,
      /publishEnd: motionScopeHandle3\.cue\("done"\)/u,
    );
    assert.match(vanillaSource, /startAt: motionScopeHandle\.end/u);
    assert.match(vanillaSource, /startAt: motionScopeHandle3\.start/u);
    assert.match(vanillaSource, /motion: false/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(
      reactSource,
      /import \{[\s\S]*\bMotionScope\b[\s\S]*\bStep\b[\s\S]*\bafter\b[\s\S]*\buseMotionScope\b[\s\S]*\} from "@officedk\/pptxdk\/react";/u,
    );
    assert.doesNotMatch(reactSource, /stepIdentity/u);
    assert.match(reactSource, /const motionScopeHandle = useMotionScope\(\);/u);
    assert.match(reactSource, /<Step>\s*<MotionScope/u);
    assert.match(reactSource, /<Step n=\{2\}>/u);
    assert.match(
      reactSource,
      /<MotionScope\s+scope=\{motionScopeHandle\}\s+mode="sequence"\s+publishEnd=\{motionScopeHandle\.cue\("ready"\)\}/u,
    );
    assert.match(reactSource, /startAt=\{motionScopeHandle\.end\}/u);
    assert.match(
      reactSource,
      /startAt=\{after\(motionScopeHandle\.cue\("ready"\), 80\)\}/u,
    );
    assert.match(reactSource, /startAt=\{motionScopeHandle3\.start\}/u);
  });

  it("rejects malformed motion-step and motion-scope declarations with CodegenError and precise paths", async () => {
    const generate = createCodegen(documentCapabilities);
    const scope = createMotionScopeHandle();

    const badStepNumberDeck = deck().append(
      slide({ id: "s1" }).append(
        step({ n: 0 }).append(
          shape({
            id: "a",
            style: {
              position: "absolute",
              left: 0,
              top: 0,
              width: 20,
              height: 20,
            },
          }),
        ),
      ),
    );
    assert.throws(
      () => generate(badStepNumberDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-step[1]@n",
    );

    const badStepPropDeck = deck().append(
      slide({ id: "s1" }).append(
        step({ extra: true } as unknown as Parameters<typeof step>[0]),
      ),
    );
    assert.throws(
      () => generate(badStepPropDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-step[1]@extra",
    );

    const badStepTextDeck = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          Slide,
          { id: "s1" },
          createElement(Step, { n: 1 }, "stray step text"),
        ),
      ),
    );
    assert.throws(
      () => generate(badStepTextDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-step[1]",
    );

    const masterScopeDeck = deck().append(
      slideMaster({ id: "m1" }).append(motionScope(scope)),
    );
    assert.throws(
      () => generate(masterScopeDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide-master[m1]/motion-scope[1]",
    );

    const badScopeEachDeck = deck().append(
      slide({ id: "s1" }).append(
        motionScope(scope, { mode: "stagger", each: Number.NaN }),
      ),
    );
    assert.throws(
      () => generate(badScopeEachDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-scope[1]@each",
    );

    const badScopeStartAtDeck = deck().append(
      slide({ id: "s1" }).append(
        motionScope(scope, {
          startAt: "invalid-timing" as unknown as ReturnType<typeof after>,
        }),
      ),
    );
    assert.throws(
      () => generate(badScopeStartAtDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-scope[1]@startAt",
    );

    const badScopeTextDeck = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          Slide,
          { id: "s1" },
          createElement(MotionScope, { scope }, "stray scope text"),
        ),
      ),
    );
    assert.throws(
      () => generate(badScopeTextDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/motion-scope[1]",
    );
  });

  it("serializes all entrance/emphasis presets, repeat wrappers, entrance arrays, motion flags, and chart/text granularities", () => {
    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "effects" }).append(
        shape({
          id: "s-fade",
          enter: fade(),
          emphasis: spin(),
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 40,
            height: 20,
          },
        }),
        shape({
          id: "s-wipe-fill",
          enter: wipe({ from: "left", duration: 300, delay: 40 }),
          emphasis: fillColor({
            from: "#2563EB",
            to: "#10B981",
            duration: 450,
            easing: { accelerate: 0.25, decelerate: 0.5 },
          }),
          style: {
            position: "absolute",
            left: 60,
            top: 10,
            width: 40,
            height: 20,
          },
        }),
        shape({
          id: "s-zoom-breathe",
          enter: [
            zoom({ duration: 220 }),
            fly({ from: "top", duration: 260 }),
            floatUp(),
            floatDown({ delay: 20 }),
            spinIn({ duration: 320 }),
            stretch({ duration: 280 }),
          ],
          emphasis: repeat(
            breathe({
              period: 1200,
              from: "#164E3D",
              to: "#4ADE80",
              delay: 50,
            }),
            "forever",
          ),
          style: {
            position: "absolute",
            left: 110,
            top: 10,
            width: 40,
            height: 20,
          },
        }),
        shape({
          id: "s-spin-repeat",
          enter: repeat(fade({ duration: 150 }), 2),
          emphasis: repeat(spin({ by: -90, duration: 400 }), 2),
          motion: false,
          style: {
            position: "absolute",
            left: 160,
            top: 10,
            width: 40,
            height: 20,
          },
        }),
        text("Word by word", {
          id: "t-words",
          enter: fade({ by: "word", intervalMs: 60, duration: 200 }),
          motion: true,
          style: {
            position: "absolute",
            left: 10,
            top: 60,
            width: 180,
            height: 40,
          },
        }),
      ),
    );

    const generate = createCodegen(documentCapabilities);
    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /enter: fade\(\)/u);
    assert.match(vanillaSource, /emphasis: spin\(\)/u);
    assert.match(
      vanillaSource,
      /enter: wipe\(\{ from: "left", duration: 300, delay: 40 \}\)/u,
    );
    assert.match(
      vanillaSource,
      /emphasis: fillColor\(\{\s*from: "#2563EB",\s*to: "#10B981",\s*duration: 450,\s*easing: \{ accelerate: 0\.25, decelerate: 0\.5 \},\s*\}\)/u,
    );
    assert.match(vanillaSource, /zoom\(\{ duration: 220 \}\)/u);
    assert.match(vanillaSource, /fly\(\{ from: "top", duration: 260 \}\)/u);
    assert.match(vanillaSource, /floatUp\(\)/u);
    assert.match(vanillaSource, /floatDown\(\{ delay: 20 \}\)/u);
    assert.match(vanillaSource, /spinIn\(\{ duration: 320 \}\)/u);
    assert.match(vanillaSource, /stretch\(\{ duration: 280 \}\)/u);
    assert.match(
      vanillaSource,
      /emphasis: repeat\(\s*breathe\(\{ period: 1200, from: "#164E3D", to: "#4ADE80", delay: 50 \}\),\s*"forever",\s*\)/u,
    );
    assert.doesNotMatch(vanillaSource, /breathe\(\{[^}]*duration:/u);
    assert.match(
      vanillaSource,
      /enter: repeat\(fade\(\{ duration: 150 \}\), 2\)/u,
    );
    assert.match(
      vanillaSource,
      /emphasis: repeat\(spin\(\{ by: -90, duration: 400 \}\), 2\)/u,
    );
    assert.match(vanillaSource, /motion: false/u);
    assert.match(vanillaSource, /motion: true/u);
    assert.match(
      vanillaSource,
      /enter: fade\(\{ duration: 200, by: "word", intervalMs: 60 \}\)/u,
    );

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /enter=\{fade\(\)\}/u);
    assert.match(reactSource, /emphasis=\{spin\(\)\}/u);
    assert.match(reactSource, /motion=\{false\}/u);
    assert.match(
      reactSource,
      /enter=\{repeat\(fade\(\{ duration: 150 \}\), 2\)\}/u,
    );
    assert.match(
      reactSource,
      /emphasis=\{repeat\(spin\(\{ by: -90, duration: 400 \}\), 2\)\}/u,
    );

    const chartDoc = deck().append(
      slide({ id: "chart-motion" }).append(
        chart({
          id: "c1",
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S1", values: [1, 2] }],
            },
          ],
          enter: fade({ by: "seriesElements", duration: 150 }),
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 300,
            height: 200,
          },
        }),
      ),
    );
    const chartVanilla = generate(chartDoc, { target: "vanilla" });
    assert.match(
      chartVanilla.files[chartVanilla.entry] as string,
      /enter: fade\(\{ duration: 150, by: "seriesElements" \}\)/u,
    );
  });

  it("rejects invalid enter, emphasis, motion, and repeat declarations with CodegenError and exact paths", () => {
    const generate = createCodegen(documentCapabilities);

    const wrongKindEnter = deck().append(
      slide({ id: "s1" }).append(
        shape({
          id: "bad-enter",
          enter: spin() as unknown as ReturnType<typeof fade>,
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 20,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(wrongKindEnter, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/shape[bad-enter]@enter",
    );

    const wrongKindEmphasis = deck().append(
      slide({ id: "s1" }).append(
        shape({
          id: "bad-emphasis",
          emphasis: fade() as unknown as ReturnType<typeof spin>,
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 20,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(wrongKindEmphasis, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/shape[bad-emphasis]@emphasis",
    );

    const nanDurationArray = deck().append(
      slide({ id: "s1" }).append(
        shape({
          id: "bad-arr",
          enter: [fade(), wipe({ duration: Number.NaN })],
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 20,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(nanDurationArray, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/shape[bad-arr]@enter[1].duration",
    );

    const badRepeatCount = deck().append(
      slide({ id: "s1" }).append(
        shape({
          id: "bad-rep",
          emphasis: repeat(spin(), 0),
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 20,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(badRepeatCount, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/shape[bad-rep]@emphasis.count",
    );

    const badMotionFlag = deck().append(
      slide({ id: "s1" }).append(
        shape({
          id: "bad-flag",
          motion: "off" as unknown as boolean,
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 20,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(badMotionFlag, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/shape[bad-flag]@motion",
    );
  });

  it("serializes OwnerMotion on SlideMaster and SlideLayout with showMasterPlaceholderAnimations, and rejects invalid OwnerMotion or Slide.motion", () => {
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
            left: 32,
            top: 24,
            width: 480,
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
                range: { type: "paragraph", start: 0, end: 0 },
              },
            ],
          },
        ],
      },
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "master-1",
        name: "Motion Master",
        placeholders: {
          ftr: {
            type: "ftr",
            style: {
              position: "absolute",
              left: 32,
              top: 490,
              width: 220,
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
                {
                  target: { spid: 7 },
                  effect: repeat(spin({ by: 90, duration: 400 }), 2),
                },
              ],
            },
          ],
        },
      }).append(slideLayout(animLayout)),
      slide({
        id: "s1",
        layout: animLayout,
        showMasterPlaceholderAnimations: true,
      }).append(text("Slide title", { id: "t1", placeholder: "title" })),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /showMasterPlaceholderAnimations: false/u);
    assert.match(vanillaSource, /showMasterPlaceholderAnimations: true/u);
    assert.match(vanillaSource, /target: \{ placeholder: "title" \}/u);
    assert.match(vanillaSource, /effect: fade\(\{ duration: 300 \}\)/u);
    assert.match(
      vanillaSource,
      /range: \{ type: "paragraph", start: 0, end: 0 \}/u,
    );
    assert.match(vanillaSource, /target: \{ placeholder: "ftr" \}/u);
    assert.match(
      vanillaSource,
      /effect: wipe\(\{ from: "left", duration: 250 \}\)/u,
    );
    assert.match(vanillaSource, /target: \{ spid: 7 \}/u);
    assert.match(
      vanillaSource,
      /effect: repeat\(spin\(\{ by: 90, duration: 400 \}\), 2\)/u,
    );

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /showMasterPlaceholderAnimations: false/u);
    assert.match(
      reactSource,
      /<Slide id="s1" layout=\{master1AnimLayout\} showMasterPlaceholderAnimations>/u,
    );
    assert.match(reactSource, /<SlideMaster[\s\S]*?motion=\{\{/u);

    const slideWithOwnerMotion = deck().append(
      slide({
        id: "bad-slide",
        motion: { steps: [] },
      } as unknown as Parameters<typeof slide>[0]),
    );
    assert.throws(
      () => generate(slideWithOwnerMotion, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-slide-property" &&
        error.diagnostics[0]?.path === "deck/slide[bad-slide]@motion",
    );

    const badMasterMotion = deck().append(
      slideMaster({
        id: "bad-master",
        motion: {
          steps: [
            {
              trigger: "auto",
              effects: [
                {
                  target: { placeholder: "ftr" },
                  effect: fade({ duration: Number.NaN }),
                },
              ],
            },
          ],
        },
      }),
    );
    assert.throws(
      () => generate(badMasterMotion, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide-master[bad-master]@motion.steps[0].effects[0].effect.duration",
    );
  });

  it("roundtrips imported chart granularity animations and preserves deterministic output", () => {
    const chartDoc = deck({ width: 640, height: 360 }).append(
      slide({ id: "chart-by" }).append(
        step({ n: 1 }).append(
          chart({
            id: "c-series",
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [
                  { name: "Rev", values: [10, 20] },
                  { name: "Cost", values: [5, 8] },
                ],
              },
            ],
            enter: fade({ by: "series", duration: 180 }),
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 280,
              height: 150,
            },
          }),
          chart({
            id: "c-cat-el",
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [
                  { name: "Rev", values: [10, 20] },
                  { name: "Cost", values: [5, 8] },
                ],
              },
            ],
            enter: wipe({
              from: "left",
              by: "categoryElements",
              duration: 160,
            }),
            style: {
              position: "absolute",
              left: 320,
              top: 20,
              width: 280,
              height: 150,
            },
          }),
        ),
      ),
    );

    const resolveAll = createResolver(documentCapabilities);
    const importAll = createImporter(documentCapabilities);
    const built = emit(resolveAll(chartDoc));
    const pptxCopy = new Uint8Array(built.pptx);
    const imported = importAll(pptxCopy);
    pptxCopy.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const generate = createCodegen(documentCapabilities);
    const vanilla1 = generate(imported.document, {
      target: "vanilla",
    });
    const vanilla2 = generate(imported.document, {
      target: "vanilla",
    });
    assert.equal(
      vanilla1.files[vanilla1.entry],
      vanilla2.files[vanilla2.entry],
    );

    const vanillaSource = vanilla1.files[vanilla1.entry] as string;
    assert.match(vanillaSource, /by: "series"/u);
    assert.match(vanillaSource, /by: "categoryElements"/u);
    assert.match(vanillaSource, /step\(\{ n: 1 \}\)/u);

    const react = generate(imported.document, {
      target: "react",
    });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /by: "series"/u);
    assert.match(reactSource, /by: "categoryElements"/u);
    assert.match(reactSource, /<Step n=\{1\}>/u);
  });
});
