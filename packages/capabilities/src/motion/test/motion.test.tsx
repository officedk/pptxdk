import assert from "node:assert/strict";

import {
  after,
  Box,
  Chart,
  breathe,
  fade,
  fly,
  floatDown,
  floatUp,
  Group,
  Image,
  MotionScope,
  P,
  Shape,
  shapeCapability,
  spin,
  spinIn,
  stretch,
  Step,
  Text,
  motionCapability,
  observeMotion,
  repeat,
  transitionCapability,
  transitions,
  type MotionObservation,
  useMotionScope,
  wipe,
  zoom,
} from "@pptxdk/capabilities";
import {
  deck as documentDeck,
  paragraph as documentParagraph,
  shape as documentShape,
  slide as documentSlide,
  step as documentStep,
} from "@pptxdk/capabilities/document";
import {
  fade as documentFade,
  fly as documentFly,
} from "@pptxdk/capabilities/motion";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, Fragment, type ReactNode } from "react";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

function rewriteSlideXml(
  pptx: Uint8Array,
  rewrite: (xml: string) => string,
): Uint8Array {
  const path = "ppt/slides/slide1.xml";
  const parts = unzipSync(pptx);
  const xml = strFromU8(parts[path]!);
  const next = rewrite(xml);
  assert.notEqual(next, xml);
  return zipSync({ ...parts, [path]: strToU8(next) });
}

function animationEffectBlocks(xml: string): {
  readonly block: string;
  readonly timingId: number;
  readonly presetId: number;
}[] {
  return [
    ...xml.matchAll(
      /<p:par><p:cTn id="(\d+)"[^>]*presetID="(\d+)"[^>]*>[\s\S]*?<\/p:par>/gu,
    ),
  ].map(([block, timingId, presetId]) => ({
    block,
    timingId: Number(timingId),
    presetId: Number(presetId),
  }));
}

function renumberTimingNodeIds(xml: string, offset: number): string {
  return xml
    .replace(
      /(<p:cTn\b[^>]*\bid=")(\d+)(")/gu,
      (_match, prefix: string, id: string, suffix: string) =>
        `${prefix}${Number(id) + offset}${suffix}`,
    )
    .replace(
      /(<p:tn\b[^>]*\bval=")(\d+)(")/gu,
      (_match, prefix: string, id: string, suffix: string) =>
        `${prefix}${Number(id) + offset}${suffix}`,
    );
}

function RepeatedSpinDeck() {
  const body = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "finite-repeat" },
      createElement(
        MotionScope,
        { scope: body },
        createElement(Shape, {
          id: "three-turns",
          x: 20,
          y: 20,
          width: 40,
          height: 40,
          fill: "#164E3D",
          emphasis: repeat(spin({ duration: 100, from: 0 }), 3),
        }),
      ),
    ),
  );
}

function TimeProjectionDeck({
  outerDelay,
  innerDelay,
  duration,
  repeatCount,
}: {
  readonly outerDelay: number;
  readonly innerDelay: number;
  readonly duration: number;
  readonly repeatCount: number;
}) {
  const outer = useMotionScope();
  const inner = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { id: "time-projection" },
      createElement(
        MotionScope,
        { scope: outer, delay: outerDelay },
        createElement(
          MotionScope,
          { scope: inner, delay: innerDelay },
          createElement(Shape, {
            id: "timed-repeat",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
            emphasis: repeat(spin({ duration }), repeatCount),
          }),
        ),
      ),
    ),
  );
}

function RepeatedAndBreathingDeck() {
  const choreography = useMotionScope();
  const body = useMotionScope();
  const lamp = useMotionScope();
  const afterBody = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "breathing-scope" },
      createElement(
        MotionScope,
        { scope: choreography },
        createElement(
          MotionScope,
          { scope: body },
          createElement(Shape, {
            id: "finite-body",
            name: "Finite body - spin three turns",
            x: 20,
            y: 20,
            width: 40,
            height: 40,
            fill: "#164E3D",
            emphasis: repeat(spin({ duration: 1200 }), 3),
          }),
          createElement(Shape, {
            id: "finite-breathe",
            name: "Finite body - two breathe cycles",
            x: 20,
            y: 70,
            width: 20,
            height: 20,
            fill: "#FFFFFF",
            emphasis: repeat(
              breathe({ period: 2000, from: "#164E3D", to: "#4ADE80" }),
              2,
            ),
          }),
        ),
        createElement(
          MotionScope,
          { scope: lamp },
          createElement(
            Group,
            {
              id: "native-lamp-group",
              name: "Native group - two spins",
              emphasis: repeat(spin({ duration: 1200 }), 2),
            },
            createElement(Shape, {
              id: "infinite-lamp",
              name: "Group member - breathe until slide exit",
              x: 80,
              y: 20,
              width: 40,
              height: 40,
              fill: "#164E3D",
              emphasis: repeat(
                breathe({
                  period: 1600,
                  from: "#164E3D",
                  to: "#4ADE80",
                }),
                "forever",
              ),
            }),
          ),
        ),
      ),
      createElement(
        MotionScope,
        { scope: afterBody, startAt: body.end },
        createElement(Shape, {
          id: "after-body",
          x: 140,
          y: 20,
          width: 40,
          height: 40,
          fill: "#4ADE80",
          enter: fade({ duration: 100 }),
        }),
      ),
      createElement(
        Step,
        { n: 1 },
        createElement(Shape, {
          id: "first-click",
          x: 20,
          y: 80,
          width: 40,
          height: 20,
          fill: "#3B82F6",
          enter: fade({ duration: 80 }),
        }),
      ),
      createElement(
        Step,
        { n: 2 },
        createElement(Shape, {
          id: "second-click",
          x: 80,
          y: 80,
          width: 40,
          height: 20,
          fill: "#F97316",
          enter: fade({ duration: 80 }),
        }),
      ),
    ),
    createElement(
      Slide,
      { id: "after-breathing-slide" },
      createElement(Shape, {
        id: "next-page",
        x: 20,
        y: 20,
        width: 40,
        height: 40,
        fill: "#4ADE80",
      }),
    ),
  );
}

function InfiniteSequenceTail({
  blockAfter,
}: {
  readonly blockAfter: boolean;
}) {
  const sequence = useMotionScope();
  const lamp = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: blockAfter ? "blocked-infinite-tail" : "infinite-tail" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        createElement(Shape, {
          id: "before-lamp",
          x: 0,
          y: 0,
          width: 20,
          height: 20,
          enter: fade({ duration: 100 }),
        }),
        createElement(
          MotionScope,
          { scope: lamp },
          createElement(Shape, {
            id: "tail-lamp",
            x: 30,
            y: 0,
            width: 20,
            height: 20,
            fill: "#164E3D",
            emphasis: repeat(
              breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
              "forever",
            ),
          }),
        ),
        blockAfter
          ? createElement(Shape, {
              id: "blocked-follower",
              x: 60,
              y: 0,
              width: 20,
              height: 20,
              enter: fade({ duration: 100 }),
            })
          : null,
      ),
    ),
  );
}

describe("object motion", () => {
  it("enters choreography completion and actual file by total count for finite repeat", async () => {
    const result = await build(createElement(RepeatedSpinDeck));
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const effect = motion.steps[0]?.effects[0];
    const scope = motion.scopes?.[0];

    assert.deepEqual(
      [
        effect?.preset,
        effect?.durationMs,
        effect?.repeatCount,
        effect?.startMs,
        effect?.endMs,
        effect?.completion,
      ],
      ["spin", 100, 3, 0, 300, undefined],
    );
    assert.deepEqual(
      [scope?.completion, scope?.startMs, scope?.endMs],
      ["known", 0, 300],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileEffect = file.slides[0]?.animations?.[0]?.effects[0];
    assert.deepEqual(
      [fileEffect?.preset, fileEffect?.durationMs, fileEffect?.repeatCount],
      ["spin", 100, 3],
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const timing = strFromU8(slidePart);
    assert.match(timing, /<p:cTn[^>]*presetID="8"[^>]*repeatCount="3000"/u);
  });

  it("keeps continuous breathe without natural completion and waits only for finite body scopes", async () => {
    const result = await build(createElement(RepeatedAndBreathingDeck));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const effects = motion.steps.flatMap((step) => step.effects);
    const bodyEffect = effects.find(
      ({ target }) => target.authorId === "finite-body",
    );
    const finiteBreathe = effects.find(
      ({ target }) => target.authorId === "finite-breathe",
    );
    const lampEffect = effects.find(
      ({ target }) => target.authorId === "infinite-lamp",
    );
    const groupEffect = effects.find(
      ({ target }) => target.authorId === "native-lamp-group",
    );
    const follower = effects.find(
      ({ target }) => target.authorId === "after-body",
    );
    assert.deepEqual(
      [
        bodyEffect?.preset,
        bodyEffect?.repeatCount,
        bodyEffect?.startMs,
        bodyEffect?.endMs,
      ],
      ["spin", 3, 0, 3600],
    );
    assert.deepEqual(
      [
        finiteBreathe?.preset,
        finiteBreathe?.durationMs,
        finiteBreathe?.repeatCount,
        finiteBreathe?.endMs,
        finiteBreathe?.fromColor,
        finiteBreathe?.toColor,
      ],
      ["breathe", 2000, 2, 4000, "#164E3D", "#4ADE80"],
    );
    assert.deepEqual(
      [
        lampEffect?.preset,
        lampEffect?.durationMs,
        lampEffect?.repeatCount,
        lampEffect?.fromColor,
        lampEffect?.toColor,
        lampEffect?.startMs,
        lampEffect?.endMs,
        lampEffect?.completion,
      ],
      [
        "breathe",
        1600,
        "forever",
        "#164E3D",
        "#4ADE80",
        0,
        undefined,
        "no-natural-completion",
      ],
    );
    assert.equal("endMs" in (lampEffect ?? {}), false);
    assert.deepEqual(
      [groupEffect?.preset, groupEffect?.repeatCount, groupEffect?.endMs],
      ["spin", 2, 2400],
    );
    assert.notEqual(groupEffect?.target.id, lampEffect?.target.id);
    assert.deepEqual([follower?.startMs, follower?.endMs], [4000, 4100]);
    const bodyScope = motion.scopes?.find((scope) =>
      scope.members.some(
        (member) =>
          member.kind === "object" && member.id === bodyEffect?.target.id,
      ),
    );
    const lampScope = motion.scopes?.find((scope) =>
      scope.members.some(
        (member) =>
          member.kind === "object" && member.id === lampEffect?.target.id,
      ),
    );
    const choreography = motion.scopes?.find(
      (scope) =>
        scope.members.some(
          (member) => member.kind === "scope" && member.id === bodyScope?.id,
        ) &&
        scope.members.some(
          (member) => member.kind === "scope" && member.id === lampScope?.id,
        ),
    );
    assert.deepEqual(
      [bodyScope?.completion, bodyScope?.startMs, bodyScope?.endMs],
      ["known", 0, 4000],
    );
    assert.deepEqual(
      [lampScope?.completion, lampScope?.startMs, lampScope?.endMs],
      ["no-natural-completion", 0, undefined],
    );
    assert.equal("endMs" in (lampScope ?? {}), false);
    assert.deepEqual(
      [choreography?.completion, choreography?.startMs, choreography?.endMs],
      ["no-natural-completion", 0, undefined],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const bodyObject = slide?.objects.find(
      ({ authorId: testId }) => testId === "finite-body",
    );
    const finiteBreatheObject = slide?.objects.find(
      ({ authorId: testId }) => testId === "finite-breathe",
    );
    const lampObject = slide?.objects
      .flatMap((object) => [object, ...(object.members ?? [])])
      .find(({ authorId: testId }) => testId === "infinite-lamp");
    const groupObject = slide?.objects.find(
      ({ authorId: testId }) => testId === "native-lamp-group",
    );
    const fileEffects = file.slides[0]?.animations?.flatMap(
      (step) => step.effects,
    );
    const fileBody = fileEffects?.find(
      ({ targetId }) => targetId === bodyObject?.id,
    );
    const fileFiniteBreathe = fileEffects?.find(
      ({ targetId }) => targetId === finiteBreatheObject?.id,
    );
    const fileLamp = fileEffects?.find(
      ({ targetId, preset }) =>
        targetId === lampObject?.id && preset === "breathe",
    );
    const groupProjection = motion.nativeProjections?.find(
      ({ logicalTargetId, preset }) =>
        logicalTargetId === groupObject?.id && preset === "spin",
    );
    const fileGroupEffects =
      fileEffects?.filter(({ timingId }) =>
        groupProjection?.nativeEffects.some(
          (native) => native.timingId === timingId,
        ),
      ) ?? [];
    assert.deepEqual(
      [fileBody?.durationMs, fileBody?.repeatCount, fileBody?.completion],
      [1200, 3, undefined],
    );
    assert.deepEqual(
      [
        fileFiniteBreathe?.preset,
        fileFiniteBreathe?.durationMs,
        fileFiniteBreathe?.repeatCount,
        fileFiniteBreathe?.fromColor,
        fileFiniteBreathe?.toColor,
      ],
      ["breathe", 2000, 2, "#164E3D", "#4ADE80"],
    );
    assert.deepEqual(
      [
        fileLamp?.preset,
        fileLamp?.durationMs,
        fileLamp?.repeatCount,
        fileLamp?.fromColor,
        fileLamp?.toColor,
        fileLamp?.completion,
      ],
      [
        "breathe",
        1600,
        "forever",
        "#164E3D",
        "#4ADE80",
        "no-natural-completion",
      ],
    );
    assert.deepEqual(
      fileGroupEffects.map(({ preset, repeatCount, durationMs }) => [
        preset,
        repeatCount,
        durationMs,
      ]),
      groupObject?.members?.map(() => ["spin", 2, 1200]),
    );
    assert.deepEqual(
      fileGroupEffects.map(({ targetId }) => targetId),
      groupObject?.members?.map(({ id }) => id),
    );
    assert.ok(
      fileGroupEffects.some(({ targetId }) => targetId === fileLamp?.targetId),
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const timing = strFromU8(slidePart);
    const finiteShapeStart = timing.indexOf(
      `<p:cNvPr id="${finiteBreatheObject?.id}" name="Finite body - two breathe cycles"`,
    );
    const finiteShapeFill = timing.indexOf(
      '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>',
      finiteShapeStart,
    );
    const nextShape = timing.indexOf("<p:sp>", finiteShapeStart + 1);
    assert.ok(
      finiteShapeStart >= 0 &&
        finiteShapeFill > finiteShapeStart &&
        (nextShape < 0 || finiteShapeFill < nextShape),
    );
    assert.match(
      timing,
      /<p:cTn[^>]*presetID="0"[^>]*presetClass="emph"[^>]*repeatCount="indefinite"[^>]*autoRev="1"/u,
    );
    assert.match(
      timing,
      /<p:animClr clrSpc="rgb">[\s\S]*?<p:cTn[^>]*dur="800"[^>]*>[\s\S]*?<p:attrName>fillcolor<\/p:attrName>[\s\S]*?<p:from><a:srgbClr val="164E3D"><\/a:srgbClr><\/p:from>[\s\S]*?<p:to><a:srgbClr val="4ADE80"><\/a:srgbClr><\/p:to><\/p:animClr>/u,
    );
    assert.match(
      timing,
      /<p:cTn[^>]*presetID="0"[^>]*repeatCount="2000"[^>]*autoRev="1"/u,
    );
    assert.match(
      timing,
      /<p:cTn[^>]*presetID="0"[^>]*repeatCount="2000"[^>]*autoRev="1"[^>]*>[\s\S]*?<p:animClr clrSpc="rgb"><p:cBhvr><p:cTn[^>]*dur="1000"/u,
    );
  });

  it("rejects waiting for file effects without natural completion in standalone reader", async () => {
    const result = await build(createElement(RepeatedAndBreathingDeck));
    const parts = unzipSync(result.pptx);
    const slidePath = "ppt/slides/slide1.xml";
    const slidePart = parts[slidePath];
    assert.ok(slidePart);
    const timing = strFromU8(slidePart);
    const infiniteSourceId = timing.match(
      /<p:cTn id="(\d+)"[^>]*presetID="0"[^>]*repeatCount="indefinite"/u,
    )?.[1];
    const finiteTargetId = timing.match(
      /<p:cTn id="(\d+)"[^>]*presetID="8"/u,
    )?.[1];
    assert.ok(infiniteSourceId);
    assert.ok(finiteTargetId);
    const condition = new RegExp(
      `(<p:cTn id="${finiteTargetId}"[^>]*><p:stCondLst>)[\\s\\S]*?(</p:stCondLst>)`,
      "u",
    );
    const invalidTiming = timing.replace(
      condition,
      `$1<p:cond evt="end" delay="0"><p:tn val="${infiniteSourceId}"></p:tn></p:cond>$2`,
    );
    assert.notEqual(invalidTiming, timing);

    const invalidFile = await validatePptx(
      zipSync({ ...parts, [slidePath]: strToU8(invalidTiming) }),
    );
    assert.ok(
      invalidFile.errors.some(({ code }) => code === "no-natural-completion"),
    );
  });

  it("restricts ranges by actual timing fields and preserves precisely representable long repeats", async () => {
    const result = await build(
      createElement(TimeProjectionDeck, {
        outerDelay: 4_294_967_295,
        innerDelay: 0,
        duration: 2_500_000_000,
        repeatCount: 4,
      }),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const effect = motion.steps[0]?.effects[0];
    assert.deepEqual(
      [effect?.startMs, effect?.endMs, effect?.durationMs, effect?.repeatCount],
      [4_294_967_295, 14_294_967_295, 2_500_000_000, 4],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const timing = strFromU8(slidePart);
    assert.match(timing, /<p:cond delay="4294967295"><\/p:cond>/u);
    assert.match(
      timing,
      /<p:cTn[^>]*presetID="8"[^>]*repeatCount="4000"[^>]*>[\s\S]*?<p:animRot from="21600000" to="0"><p:cBhvr><p:cTn[^>]*dur="2500000000"/u,
    );

    const maxPeriod = 8_589_934_590;
    const breatheResult = await build(
      createElement(
        Deck,
        { width: 120, height: 100 },
        createElement(
          Slide,
          { id: "maximum-breathe-period" },
          createElement(Shape, {
            id: "maximum-breathe-period",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
            fill: "#164E3D",
            emphasis: breathe({
              period: maxPeriod,
              from: "#164E3D",
              to: "#4ADE80",
            }),
          }),
        ),
      ),
    );
    const breatheMotion = breatheResult.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const breatheEffect = breatheMotion.steps[0]?.effects[0];
    assert.deepEqual(
      [breatheEffect?.durationMs, breatheEffect?.endMs],
      [maxPeriod, maxPeriod],
    );
    const breatheFile = await validatePptx(breatheResult.pptx);
    assert.deepEqual(breatheFile.errors, []);
    const breatheSlide = unzipSync(breatheResult.pptx)["ppt/slides/slide1.xml"];
    assert.ok(breatheSlide);
    assert.match(
      strFromU8(breatheSlide),
      /<p:animClr clrSpc="rgb"><p:cBhvr><p:cTn[^>]*dur="4294967295"/u,
    );
  });

  it("preserves individually encodable nested delays and rejects known cues losing integer precision", async () => {
    const cases = [
      {
        outerDelay: 4_294_967_295,
        innerDelay: 1,
        duration: 1000,
        repeatCount: 2,
      },
      {
        outerDelay: Number.MAX_SAFE_INTEGER,
        innerDelay: 0,
        duration: 2000,
        repeatCount: 2,
      },
    ] as const;

    const [nestedDelays, unsafeTotal] = cases;
    assert.ok(nestedDelays);
    assert.ok(unsafeTotal);
    const result = await build(createElement(TimeProjectionDeck, nestedDelays));
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const effect = motion.steps[0]?.effects[0];
    assert.deepEqual(
      [effect?.startMs, effect?.endMs],
      [4_294_967_296, 4_294_969_296],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const timing = strFromU8(slidePart);
    assert.match(timing, /<p:cond delay="4294967295"><\/p:cond>/u);
    assert.match(timing, /<p:cond delay="1"><\/p:cond>/u);

    await assert.rejects(
      build(createElement(TimeProjectionDeck, unsafeTotal)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-motion-time");
        return true;
      },
    );
  });

  it("permits final sequence items to loop continuously and rejects sequence successors blocked by them", async () => {
    const result = await build(
      createElement(InfiniteSequenceTail, { blockAfter: false }),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs, completion }) => [
        target.authorId,
        startMs,
        endMs,
        completion,
      ]),
      [
        ["before-lamp", 0, 100, undefined],
        ["tail-lamp", 100, undefined, "no-natural-completion"],
      ],
    );
    assert.equal(motion.scopes?.[0]?.completion, "no-natural-completion");
    assert.equal("endMs" in (motion.scopes?.[0] ?? {}), false);

    await assert.rejects(
      build(createElement(InfiniteSequenceTail, { blockAfter: true })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "no-natural-completion");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /shape\[blocked-follower\]/u,
        );
        return true;
      },
    );
  });

  it("rejects invalid repeat counts, effects, and breathe parameters for finite repeat", async () => {
    const breatheWithGranularity = {
      period: 2000,
      from: "#164E3D",
      to: "#4ADE80",
      by: "word" as const,
    };
    const breatheWithUnsupportedDuration = {
      period: 2000,
      from: "#164E3D",
      to: "#4ADE80",
      duration: 1000,
    };
    const cases = [
      {
        id: "repeat-fade",
        props: {
          enter: repeat(fade({ duration: 100 }), 2),
        },
        code: "unsupported-repeat",
      },
      {
        id: "repeat-zero",
        props: { emphasis: repeat(spin(), 0) },
        code: "invalid-repeat-count",
      },
      {
        id: "repeat-missing-count",
        props: { emphasis: repeat(spin(), undefined as never) },
        code: "invalid-repeat-count",
      },
      {
        id: "repeat-count-limit",
        props: { emphasis: repeat(spin(), 4_294_968) },
        code: "invalid-repeat-count",
      },
      {
        id: "repeat-duration-limit",
        props: {
          emphasis: repeat(spin({ duration: 4_294_967_296 }), 1),
        },
        code: "invalid-repeat-duration",
      },
      {
        id: "repeat-forever-spin",
        props: { emphasis: repeat(spin(), "forever") },
        code: "unsupported-repeat",
      },
      {
        id: "repeat-granularity",
        props: { emphasis: repeat(spin({ by: "letter" }), 2) },
        code: "unsupported-repeat-combination",
      },
      {
        id: "repeat-breathe-granularity",
        props: {
          fill: "#164E3D",
          emphasis: repeat(breathe(breatheWithGranularity), 2),
        },
        code: "unsupported-repeat-combination",
      },
      {
        id: "breathe-unsupported-duration",
        props: {
          fill: "#164E3D",
          emphasis: breathe(breatheWithUnsupportedDuration),
        },
        code: "unknown-parameter",
      },
      {
        id: "breathe-null-easing",
        props: {
          fill: "#164E3D",
          emphasis: breathe({
            period: 2000,
            from: "#164E3D",
            to: "#4ADE80",
            easing: null,
          } as never),
        },
        code: "invalid-easing",
      },
      {
        id: "breathe-false-easing",
        props: {
          fill: "#164E3D",
          emphasis: breathe({
            period: 2000,
            from: "#164E3D",
            to: "#4ADE80",
            easing: false,
          } as never),
        },
        code: "invalid-easing",
      },
      {
        id: "odd-breathe-period",
        props: {
          emphasis: breathe({
            period: 1601,
            from: "#164E3D",
            to: "#4ADE80",
          }),
        },
        code: "invalid-breathe-period",
      },
      {
        id: "breathe-period-limit",
        props: {
          emphasis: breathe({
            period: 8_589_934_592,
            from: "#164E3D",
            to: "#4ADE80",
          }),
        },
        code: "invalid-breathe-period",
      },
      {
        id: "dark-breathe-endpoint",
        props: {
          fill: "#FFFFFF",
          emphasis: breathe({
            period: 1600,
            from: "#FFFFFF",
            to: "#000000",
          }),
        },
        code: "invalid-breathe-colors",
      },
      {
        id: "unfilled-breathe-target",
        props: {
          emphasis: breathe({
            period: 1600,
            from: "#164E3D",
            to: "#4ADE80",
          }),
        },
        code: "invalid-breathe-target",
      },
    ] as const;

    for (const { id: testId, props, code } of cases) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { id: `invalid-${testId}` },
              createElement(Shape, {
                id: testId,
                x: 0,
                y: 0,
                width: 20,
                height: 20,
                ...props,
              } as never),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          assert.match(
            error.diagnostics[0]?.path ?? "",
            new RegExp(`shape\\[${testId}\\]`, "u"),
          );
          return true;
        },
      );
    }
  });

  it("binds breathe color only to filled Shape members and not to native Group containers", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          {},
          createElement(
            Slide,
            { id: "group-breathe-target" },
            createElement(
              Group,
              {
                id: "group-breathe",
                emphasis: breathe({
                  period: 1600,
                  from: "#164E3D",
                  to: "#4ADE80",
                }),
              },
              createElement(Shape, {
                id: "group-breathe-child",
                x: 20,
                y: 20,
                width: 20,
                height: 20,
                fill: "#164E3D",
              }),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-breathe-target");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /group\[group-breathe\]/u,
        );
        return true;
      },
    );
  });

  it("preserves paragraph build without altering effect order when paragraph effects coexist with whole effects", async () => {
    const cases: readonly {
      readonly id: string;
      readonly paragraphs: readonly string[];
      readonly paragraphEffectFirst: boolean;
    }[] = [
      {
        id: "mixed-one-paragraph",
        paragraphs: ["一段文字😀"],
        paragraphEffectFirst: true,
      },
      {
        id: "mixed-two-paragraphs",
        paragraphs: ["第一段😀", "第二段"],
        paragraphEffectFirst: false,
      },
    ];

    for (const { id: testId, paragraphs, paragraphEffectFirst } of cases) {
      const paragraphEffect = fade({ by: "paragraph", duration: 100 });
      const wholeEffect = fly({ duration: 1000 });
      const enter = paragraphEffectFirst
        ? [paragraphEffect, wholeEffect]
        : [wholeEffect, paragraphEffect];
      const children =
        paragraphs.length === 1
          ? [paragraphs[0]]
          : paragraphs.map((paragraph, index) =>
              createElement(P, { key: `${testId}-${index}` }, paragraph),
            );
      const result = await build(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: testId },
            createElement(
              Text,
              {
                id: testId,
                x: 0,
                y: 0,
                width: 200,
                height: 100,
                enter,
              },
              ...children,
            ),
          ),
        ),
      );
      const target = result.report.slides[0]?.objects[0];
      assert.ok(target);
      const motion = result.report.slides[0]?.capabilities?.[
        "motion"
      ] as MotionObservation;
      assert.deepEqual(
        motion.steps[0]?.effects.map(
          ({ target: effectTarget, preset, durationMs, by }) => [
            effectTarget.authorId,
            preset,
            durationMs,
            by,
            effectTarget.range,
          ],
        ),
        paragraphEffectFirst
          ? [
              [
                testId,
                "fade",
                100,
                "paragraph",
                { type: "paragraph", start: 0, end: 0 },
              ],
              [testId, "fly", 1000, undefined, undefined],
            ]
          : [
              [testId, "fly", 1000, undefined, undefined],
              [
                testId,
                "fade",
                100,
                "paragraph",
                { type: "paragraph", start: 0, end: 0 },
              ],
              [
                testId,
                "fade",
                100,
                "paragraph",
                { type: "paragraph", start: 1, end: 1 },
              ],
            ],
      );

      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(
        file.slides[0]?.objects[0]?.text?.paragraphs.length,
        paragraphs.length,
      );
      assert.deepEqual(
        file.slides[0]?.animations?.[0]?.effects.map(
          ({ preset, durationMs, targetRange, nativeBuild }) => [
            preset,
            durationMs,
            targetRange,
            nativeBuild,
          ],
        ),
        paragraphEffectFirst
          ? [
              [
                "fade",
                100,
                { type: "paragraph", start: 0, end: 0 },
                "paragraph",
              ],
              ["fly", 1000, undefined, "paragraph"],
            ]
          : [
              ["fly", 1000, undefined, "paragraph"],
              [
                "fade",
                100,
                { type: "paragraph", start: 0, end: 0 },
                "paragraph",
              ],
              [
                "fade",
                100,
                { type: "paragraph", start: 1, end: 1 },
                "paragraph",
              ],
            ],
      );

      const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
      assert.ok(slidePart);
      assert.ok(
        strFromU8(slidePart).includes(
          `<p:bldP spid="${target.id}" grpId="0" build="p"></p:bldP>`,
        ),
      );
    }
  });

  it("generates text effects paragraph by paragraph according to native paragraph ranges and reads back from same file", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "paragraph-motion" },
          createElement(
            Text,
            {
              id: "paragraphs",
              x: 10,
              y: 10,
              width: 240,
              height: 100,
              enter: fade({ by: "paragraph", duration: 300 }),
            },
            createElement(P, null, "First"),
            createElement(P, null, "Second"),
          ),
        ),
      ),
    );

    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.range,
        startMs,
        endMs,
      ]),
      [
        [{ type: "paragraph", start: 0, end: 0 }, 0, 300],
        [{ type: "paragraph", start: 1, end: 1 }, 300, 600],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ targetId, targetRange, nativeBuild, durationMs, delayMs }) => [
          targetId,
          targetRange,
          nativeBuild,
          durationMs,
          delayMs,
        ],
      ),
      [
        [
          file.slides[0]?.objects[0]?.id,
          { type: "paragraph", start: 0, end: 0 },
          "paragraph",
          300,
          0,
        ],
        [
          file.slides[0]?.objects[0]?.id,
          { type: "paragraph", start: 1, end: 1 },
          "paragraph",
          300,
          300,
        ],
      ],
    );
  });

  it("preserves word-by-word unknown completion and starts successors with a single native end event", async () => {
    function LetterAndFollower() {
      const letters = useMotionScope();
      const follower = useMotionScope();
      return createElement(
        Fragment,
        null,
        createElement(
          MotionScope,
          { scope: letters },
          createElement(
            Text,
            {
              id: "letters",
              name: "letters",
              x: 10,
              y: 10,
              width: 240,
              height: 60,
              enter: fade({
                by: "letter",
                duration: 1000,
                intervalMs: 500,
              }),
            },
            "LATE EARLY",
          ),
        ),
        createElement(Text, {
          id: "words",
          name: "words",
          x: 10,
          y: 70,
          width: 240,
          height: 40,
          enter: fade({ by: "word", duration: 600 }),
          children: "First second",
        }),
        createElement(
          MotionScope,
          { scope: follower, startAt: letters.end },
          createElement(Shape, {
            id: "after-letters",
            name: "after-letters",
            x: 10,
            y: 90,
            style: { width: 60, height: 30 },
            fill: "#123456",
            enter: fade({ duration: 200 }),
          }),
        ),
      );
    }

    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "letter-end" },
          createElement(LetterAndFollower),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const effects = motion.steps[0]?.effects ?? [];
    const source = effects.find(({ target }) => target.authorId === "letters");
    const follower = effects.find(
      ({ target }) => target.authorId === "after-letters",
    );
    const wordEffect = effects.find(
      ({ target }) => target.authorId === "words",
    );
    assert.deepEqual(
      [
        source?.by,
        source?.intervalMs,
        source?.durationMs,
        source?.startMs,
        source?.endMs,
        source?.completion,
      ],
      ["letter", 500, 1000, 0, undefined, "unknown"],
    );
    const sourceScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) => member.kind === "object" && member.id === source?.target.id,
      ),
    );
    assert.deepEqual(follower?.startAfter, {
      scopeId: sourceScope?.id,
      point: "end",
      offsetMs: 0,
    });
    assert.equal(follower?.startMs, undefined);
    assert.equal(follower?.endMs, undefined);
    assert.equal(follower?.completion, "unknown");
    assert.deepEqual(
      [
        wordEffect?.by,
        wordEffect?.intervalMs,
        wordEffect?.durationMs,
        wordEffect?.startMs,
        wordEffect?.endMs,
        wordEffect?.completion,
      ],
      ["word", 60, 600, 0, undefined, "unknown"],
    );
    assert.deepEqual(
      [sourceScope?.completion, sourceScope?.startMs, sourceScope?.endMs],
      ["unknown", 0, undefined],
    );
    const followerScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "object" && member.id === follower?.target.id,
      ),
    );
    assert.deepEqual(
      [
        followerScope?.completion,
        followerScope?.startMs,
        followerScope?.endMs,
        followerScope?.startAfter &&
          "scopeId" in followerScope.startAfter &&
          followerScope.startAfter.scopeId,
      ],
      ["unknown", undefined, undefined, sourceScope?.id],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ name }) => name),
      ["letters", "words", "after-letters"],
    );
    const fileEffects = file.slides[0]?.animations?.[0]?.effects ?? [];
    const sourceFileEffect = fileEffects.find(
      ({ targetId }) => targetId === source?.target.id,
    );
    const followerFileEffect = fileEffects.find(
      ({ targetId }) => targetId === follower?.target.id,
    );
    const wordFileEffect = fileEffects.find(
      ({ targetId }) => targetId === wordEffect?.target.id,
    );
    assert.deepEqual(
      [
        sourceFileEffect?.by,
        sourceFileEffect?.intervalMs,
        sourceFileEffect?.durationMs,
        sourceFileEffect?.completion,
      ],
      ["letter", 500, 1000, "unknown"],
    );
    assert.equal(followerFileEffect?.startAfter?.source, "container");
    assert.equal(followerFileEffect?.startAfter?.point, "end");
    assert.equal(followerFileEffect?.startAfter?.offsetMs, 0);
    assert.deepEqual(
      [
        wordFileEffect?.by,
        wordFileEffect?.intervalMs,
        wordFileEffect?.durationMs,
        wordFileEffect?.completion,
      ],
      ["word", 60, 600, "unknown"],
    );
    const sourceTimingId = followerFileEffect?.startAfter?.timingId;
    const sourceContainer = file.slides[0]?.animationContainers?.find(
      ({ timingId }) => timingId === sourceTimingId,
    );
    assert.equal(sourceContainer?.endSync, "all");
    assert.ok(
      sourceContainer?.children.some(
        ({ kind, timingId }) =>
          kind === "effect" && timingId === sourceFileEffect?.timingId,
      ),
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    assert.ok(
      strFromU8(slidePart).includes(
        `<p:cond evt="end" delay="0"><p:tn val="${sourceTimingId}"></p:tn></p:cond>`,
      ),
    );
  });

  it("expands chart effects according to native series and category ranges and preserves chart objects", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [2, 3] },
        { name: "South", values: [4, 5] },
      ],
    };
    const result = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "chart-motion" },
          createElement(Chart, {
            id: "series-chart",
            name: "series-chart",
            x: 10,
            y: 10,
            width: 210,
            height: 200,
            groups: [
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: data.categories,
                },
                series: data.series,
              },
            ],
            enter: wipe({ by: "series", duration: 400, from: "left" }),
          }),
          createElement(Chart, {
            id: "category-chart",
            name: "category-chart",
            x: 240,
            y: 10,
            width: 210,
            height: 200,
            groups: [
              {
                kind: "line",
                data: {
                  kind: "shared-categories",
                  categories: data.categories,
                },
                series: data.series,
              },
            ],
            enter: fade({ by: "category", duration: 300 }),
          }),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, by, startMs, endMs }) => [
        target.authorId,
        by,
        target.range,
        startMs,
        endMs,
      ]),
      [
        [
          "series-chart",
          "series",
          {
            type: "chart",
            step: "gridLegend",
            seriesIndex: -3,
            categoryIndex: -3,
          },
          0,
          400,
        ],
        [
          "category-chart",
          "category",
          {
            type: "chart",
            step: "gridLegend",
            seriesIndex: -3,
            categoryIndex: -3,
          },
          0,
          300,
        ],
        [
          "category-chart",
          "category",
          {
            type: "chart",
            step: "category",
            seriesIndex: -4,
            categoryIndex: 0,
          },
          300,
          600,
        ],
        [
          "series-chart",
          "series",
          { type: "chart", step: "series", seriesIndex: 0, categoryIndex: -4 },
          400,
          800,
        ],
        [
          "category-chart",
          "category",
          {
            type: "chart",
            step: "category",
            seriesIndex: -4,
            categoryIndex: 1,
          },
          600,
          900,
        ],
        [
          "series-chart",
          "series",
          { type: "chart", step: "series", seriesIndex: 1, categoryIndex: -4 },
          800,
          1200,
        ],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ kind, name }) => [kind, name]),
      [
        ["chart", "series-chart"],
        ["chart", "category-chart"],
      ],
    );
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ targetId, targetRange, nativeBuild, durationMs, delayMs }) => [
          targetId,
          targetRange,
          nativeBuild,
          durationMs,
          delayMs,
        ],
      ),
      [
        [
          file.slides[0]?.objects[0]?.id,
          {
            type: "chart",
            step: "gridLegend",
            seriesIndex: -3,
            categoryIndex: -3,
          },
          "series",
          400,
          0,
        ],
        [
          file.slides[0]?.objects[1]?.id,
          {
            type: "chart",
            step: "gridLegend",
            seriesIndex: -3,
            categoryIndex: -3,
          },
          "category",
          300,
          0,
        ],
        [
          file.slides[0]?.objects[1]?.id,
          {
            type: "chart",
            step: "category",
            seriesIndex: -4,
            categoryIndex: 0,
          },
          "category",
          300,
          300,
        ],
        [
          file.slides[0]?.objects[0]?.id,
          { type: "chart", step: "series", seriesIndex: 0, categoryIndex: -4 },
          "series",
          400,
          400,
        ],
        [
          file.slides[0]?.objects[1]?.id,
          {
            type: "chart",
            step: "category",
            seriesIndex: -4,
            categoryIndex: 1,
          },
          "category",
          300,
          600,
        ],
        [
          file.slides[0]?.objects[0]?.id,
          { type: "chart", step: "series", seriesIndex: 1, categoryIndex: -4 },
          "series",
          400,
          800,
        ],
      ],
    );
  });

  it("expands all four chart orders from unequal XY point slots without matching coordinates", async () => {
    const groups = [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        series: [
          { name: "Long", x: [3, null, 1], y: [5, 2, null] },
          { name: "Short", x: [2, 4], y: [6, 8] },
        ],
      },
    ];
    const orders = [
      "series",
      "category",
      "seriesElements",
      "categoryElements",
    ] as const;
    const chart = (by: (typeof orders)[number], x: number) =>
      createElement(Chart, {
        id: `xy-${by}`,
        name: `xy-${by}`,
        x,
        y: 10,
        width: 200,
        height: 200,
        groups,
        enter: fade({ by, duration: 100 }),
      });
    const result = await build(
      createElement(
        Deck,
        { width: 900, height: 260 },
        createElement(
          Slide,
          { id: "xy-chart-orders" },
          chart("series", 10),
          chart("category", 230),
          chart("seriesElements", 450),
          chart("categoryElements", 670),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 2 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 2 },
      ],
    } satisfies Record<
      (typeof orders)[number],
      readonly {
        readonly step: string;
        readonly seriesIndex: number;
        readonly categoryIndex: number;
      }[]
    >;
    for (const by of orders) {
      assert.deepEqual(
        motion.steps[0]?.effects
          .filter(({ target }) => target.authorId === `xy-${by}`)
          .map(({ target, by: observedBy, startMs, endMs }) => ({
            by: observedBy,
            range: target.range,
            startMs,
            endMs,
          })),
        expected[by].map((range, index) => ({
          by,
          range: { type: "chart", ...range },
          startMs: index * 100,
          endMs: (index + 1) * 100,
        })),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const objects = file.slides[0]?.objects ?? [];
    const nativeBuild = {
      series: "series",
      category: "category",
      seriesElements: "seriesEl",
      categoryElements: "categoryEl",
    } as const;
    for (const by of orders) {
      const object = objects.find(({ name }) => name === `xy-${by}`);
      assert.ok(object !== undefined);
      assert.deepEqual(
        file.slides[0]?.animations?.[0]?.effects
          .filter(({ targetId }) => targetId === object.id)
          .map(({ targetRange, nativeBuild: observedBuild }) => ({
            nativeBuild: observedBuild,
            range: targetRange,
          })),
        expected[by].map((range) => ({
          nativeBuild: nativeBuild[by],
          range: { type: "chart", ...range },
        })),
        by,
      );
    }
  });

  it("preserves duplicate category positions in all four chart orders", async () => {
    const groups = [
      {
        kind: "column" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q1", "Q3"],
        },
        series: [
          { name: "North", values: [1, 2, 3], color: "#1D4ED8" },
          { name: "South", values: [4, 5, 6], color: "#B45309" },
        ],
      },
    ];
    const orders = [
      "series",
      "category",
      "seriesElements",
      "categoryElements",
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 900, height: 260 },
        createElement(
          Slide,
          { id: "duplicate-category-orders" },
          ...orders.map((by, index) =>
            createElement(Chart, {
              id: `duplicate-${by}`,
              name: `duplicate-${by}`,
              x: 10 + index * 220,
              y: 10,
              width: 200,
              height: 200,
              groups,
              enter: fade({ by, duration: 100 }),
            }),
          ),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects
        .filter(({ target }) => target.authorId === "duplicate-seriesElements")
        .map(({ target }) => target.range),
      [
        {
          type: "chart",
          step: "gridLegend",
          seriesIndex: -3,
          categoryIndex: -3,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 0,
          categoryIndex: 0,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 0,
          categoryIndex: 1,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 0,
          categoryIndex: 2,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 1,
          categoryIndex: 0,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 1,
          categoryIndex: 1,
        },
        {
          type: "chart",
          step: "ptInSeries",
          seriesIndex: 1,
          categoryIndex: 2,
        },
      ],
    );
    assert.deepEqual(
      motion.steps[0]?.effects
        .filter(
          ({ target }) => target.authorId === "duplicate-categoryElements",
        )
        .map(({ target }) => target.range),
      [
        {
          type: "chart",
          step: "gridLegend",
          seriesIndex: -3,
          categoryIndex: -3,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 0,
          categoryIndex: 0,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 1,
          categoryIndex: 0,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 0,
          categoryIndex: 1,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 1,
          categoryIndex: 1,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 0,
          categoryIndex: 2,
        },
        {
          type: "chart",
          step: "ptInCategory",
          seriesIndex: 1,
          categoryIndex: 2,
        },
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });

  it("indexes same-kind chart groups by formal chart-wide series order", async () => {
    const categories = ["A", "B", "C"];
    const groups = [
      {
        kind: "column" as const,
        data: { kind: "shared-categories" as const, categories },
        series: [
          { name: "Group 1A", values: [1, 2, 3], color: "#1D4ED8" },
          { name: "Group 1B", values: [4, 5, 6], color: "#C2410C" },
        ],
      },
      {
        kind: "column" as const,
        data: { kind: "shared-categories" as const, categories },
        series: [
          { name: "Group 2A", values: [7, 8, 9], color: "#0F766E" },
          { name: "Group 2B", values: [10, 11, 12], color: "#7C3AED" },
        ],
      },
    ];
    const result = await build(
      createElement(
        Deck,
        { width: 600, height: 300 },
        createElement(
          Slide,
          { id: "multi-group-orders" },
          createElement(Chart, {
            id: "multi-group-series-elements",
            name: "multi-group-series-elements",
            x: 10,
            y: 10,
            width: 280,
            height: 280,
            groups,
            enter: fade({ by: "seriesElements", duration: 100 }),
          }),
          createElement(Chart, {
            id: "multi-group-category-elements",
            name: "multi-group-category-elements",
            x: 310,
            y: 10,
            width: 280,
            height: 280,
            groups,
            enter: fade({ by: "categoryElements", duration: 100 }),
          }),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const seriesElements = motion.steps[0]?.effects
      .filter(({ target }) => target.authorId === "multi-group-series-elements")
      .map(({ target }) => target.range);
    assert.deepEqual(seriesElements, [
      {
        type: "chart",
        step: "gridLegend",
        seriesIndex: -3,
        categoryIndex: -3,
      },
      ...[0, 1, 2, 3].flatMap((seriesIndex) =>
        [0, 1, 2].map((categoryIndex) => ({
          type: "chart",
          step: "ptInSeries",
          seriesIndex,
          categoryIndex,
        })),
      ),
    ]);
    const categoryElements = motion.steps[0]?.effects
      .filter(
        ({ target }) => target.authorId === "multi-group-category-elements",
      )
      .map(({ target }) => target.range);
    assert.deepEqual(categoryElements, [
      {
        type: "chart",
        step: "gridLegend",
        seriesIndex: -3,
        categoryIndex: -3,
      },
      ...[0, 1, 2].flatMap((categoryIndex) =>
        [0, 1, 2, 3].map((seriesIndex) => ({
          type: "chart",
          step: "ptInCategory",
          seriesIndex,
          categoryIndex,
        })),
      ),
    ]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });

  it("runs all four chart orders over the six category chart kinds through public build", async () => {
    const orders = [
      "series",
      "category",
      "seriesElements",
      "categoryElements",
    ] as const;
    const kinds = ["column", "bar", "line", "area", "pie", "doughnut"] as const;
    const categories = ["Q1", "Q2", "Q3"];
    const data = { kind: "shared-categories" as const, categories };
    const twoSeries = [
      { name: "North", values: [1, 2, 3] },
      { name: "South", values: [4, 5, 6] },
    ];
    const singleSeries = [{ name: "Only", values: [1, 2, 3] }];
    const groupFor = (kind: (typeof kinds)[number]) => {
      switch (kind) {
        case "column":
          return { kind, data, series: twoSeries } as const;
        case "bar":
          return { kind, data, series: twoSeries } as const;
        case "line":
          return { kind, data, series: twoSeries } as const;
        case "area":
          return { kind, data, series: twoSeries } as const;
        case "pie":
          return { kind, data, series: singleSeries } as const;
        case "doughnut":
          return { kind, data, series: singleSeries } as const;
      }
    };
    const result = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        ...kinds.flatMap((kind) =>
          orders.map((by, orderIndex) =>
            createElement(
              Slide,
              {
                id: `${kind}-${by}`,
              },
              createElement(Chart, {
                id: `${kind}-${by}`,
                name: `${kind}-${by}`,
                x: 10,
                y: 10,
                width: 460,
                height: 250,
                groups: [groupFor(kind)],
                enter: fade({ by, duration: 100 + orderIndex }),
              }),
            ),
          ),
        ),
      ),
    );
    for (const kind of kinds) {
      for (const by of orders) {
        const slide = result.report.slides.find(
          ({ id }) => id === `${kind}-${by}`,
        );
        assert.ok(slide !== undefined, `${kind}/${by}`);
        const motion = slide.capabilities?.["motion"] as
          | MotionObservation
          | undefined;
        const ranges =
          motion?.steps[0]?.effects.map(({ target }) => target.range) ?? [];
        const singleValue = kind === "pie" || kind === "doughnut";
        const expectedDataRanges =
          by === "series"
            ? singleValue
              ? [{ step: "series", seriesIndex: 0, categoryIndex: -4 }]
              : [
                  { step: "series", seriesIndex: 0, categoryIndex: -4 },
                  { step: "series", seriesIndex: 1, categoryIndex: -4 },
                ]
            : by === "category"
              ? categories.map((_, categoryIndex) => ({
                  step: "category",
                  seriesIndex: -4,
                  categoryIndex,
                }))
              : by === "seriesElements"
                ? (singleValue ? [0] : [0, 1]).flatMap((seriesIndex) =>
                    categories.map((_, categoryIndex) => ({
                      step: "ptInSeries",
                      seriesIndex,
                      categoryIndex,
                    })),
                  )
                : categories.flatMap((_, categoryIndex) =>
                    (singleValue ? [0] : [0, 1]).map((seriesIndex) => ({
                      step: "ptInCategory",
                      seriesIndex,
                      categoryIndex,
                    })),
                  );
        assert.deepEqual(
          ranges,
          [
            {
              type: "chart",
              step: "gridLegend",
              seriesIndex: -3,
              categoryIndex: -3,
            },
            ...expectedDataRanges.map((range) => ({ type: "chart", ...range })),
          ],
          `${kind}/${by}`,
        );
      }
    }
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });

  it("validates point-level chart build steps, series identities, bounds, and groups", async () => {
    const groups = [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        series: [
          { name: "Short", x: [1, 2], y: [3, 4] },
          { name: "Long", x: [5, 6, 7], y: [8, 9, 10] },
        ],
      },
    ];
    const result = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "point-chart-validation" },
          createElement(Chart, {
            id: "point-chart",
            name: "point-chart",
            x: 10,
            y: 10,
            width: 460,
            height: 250,
            groups,
            enter: fade({ by: "seriesElements", duration: 100 }),
          }),
        ),
      ),
    );
    const source = await validatePptx(result.pptx);
    assert.deepEqual(source.errors, []);
    const object = source.slides[0]?.objects.find(
      ({ kind }) => kind === "chart",
    );
    assert.ok(object !== undefined);
    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = strFromU8(unzipSync(result.pptx)[slidePath]!);
    const buildElement = slideXml.match(
      /<p:bldGraphic\b[\s\S]*?<\/p:bldGraphic>/u,
    )?.[0];
    assert.ok(buildElement !== undefined);
    const groupMatch = buildElement.match(/grpId="(\d+)"/u);
    assert.ok(groupMatch !== null);
    const groupId = Number(groupMatch[1]);
    const expectMutationError = async (
      label: string,
      mutate: (xml: string) => string,
      code: string,
      message: RegExp,
    ): Promise<void> => {
      const read = await validatePptx(rewriteSlideXml(result.pptx, mutate));
      assert.ok(
        read.errors.some(
          (error) => error.code === code && message.test(error.message),
        ),
        `${label}: ${JSON.stringify(read.errors)}`,
      );
    };

    await expectMutationError(
      "nonexistent global series",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="0" categoryIdx="0" bldStep="ptInSeries"',
          'seriesIdx="9" categoryIdx="0" bldStep="ptInSeries"',
        ),
      "invalid-animation-target-range",
      /series that does not exist/u,
    );
    await expectMutationError(
      "short series point beyond its own count",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="1" categoryIdx="2" bldStep="ptInSeries"',
          'seriesIdx="0" categoryIdx="2" bldStep="ptInSeries"',
        ),
      "invalid-animation-target-range",
      /exceeds points in native chart series/u,
    );
    await expectMutationError(
      "point sentinel",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="0" categoryIdx="0" bldStep="ptInSeries"',
          'seriesIdx="0" categoryIdx="-4" bldStep="ptInSeries"',
        ),
      "invalid-animation-target-range",
      /non-negative series and category indices/u,
    );
    await expectMutationError(
      "point step mismatched with build",
      (xml) =>
        xml.replace(
          buildElement,
          buildElement.replace('bld="seriesEl"', 'bld="categoryEl"'),
        ),
      "invalid-animation-target-range",
      /does not match native chart build mode/u,
    );
    await expectMutationError(
      "point step paired with category batch build",
      (xml) =>
        xml.replace(
          buildElement,
          buildElement.replace('bld="seriesEl"', 'bld="category"'),
        ),
      "invalid-animation-target-range",
      /does not match native chart build mode/u,
    );
    await expectMutationError(
      "point gridLegend sentinel",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="-3" categoryIdx="-3" bldStep="gridLegend"',
          'seriesIdx="-4" categoryIdx="-3" bldStep="gridLegend"',
        ),
      "invalid-animation-target-range",
      /native -3 sentinel/u,
    );
    await expectMutationError(
      "point build group",
      (xml) =>
        xml.replace(
          buildElement,
          buildElement.replace(`grpId="${groupId}"`, 'grpId="777"'),
        ),
      "invalid-chart-build-group",
      /actual chart animation effect/u,
    );
  });

  it("rejects point granularity on non-charts and mixed granularity on one chart", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "non-chart-points" },
            createElement(Shape, {
              id: "not-a-chart",
              x: 10,
              y: 10,
              width: 100,
              height: 50,
              enter: fade({ by: "seriesElements" }),
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-motion-granularity");
        return true;
      },
    );
    await assert.rejects(
      build(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: "mixed-points" },
            createElement(Chart, {
              id: "mixed-chart",
              x: 10,
              y: 10,
              width: 460,
              height: 250,
              groups: [
                {
                  kind: "column",
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2"],
                  },
                  series: [{ name: "Only", values: [1, 2] }],
                },
              ],
              enter: [fade({ by: "series" }), fade({ by: "categoryElements" })],
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-motion-granularity");
        return true;
      },
    );
  });

  it("completes every point range before a sequenced successor starts", async () => {
    function PointSequenceDeck() {
      const chartScope = useMotionScope();
      const followerScope = useMotionScope();
      return createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "point-sequence" },
          createElement(
            MotionScope,
            { scope: chartScope },
            createElement(Chart, {
              id: "point-sequence-chart",
              x: 10,
              y: 10,
              width: 460,
              height: 200,
              groups: [
                {
                  kind: "column",
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2", "Q3"],
                  },
                  series: [
                    { name: "North", values: [1, 2, 3] },
                    { name: "South", values: [4, 5, 6] },
                  ],
                },
              ],
              enter: fade({ by: "categoryElements", duration: 100 }),
            }),
          ),
          createElement(
            MotionScope,
            {
              scope: followerScope,
              startAt: after(chartScope.end, 25),
            },
            createElement(Shape, {
              id: "point-sequence-follower",
              x: 10,
              y: 220,
              width: 40,
              height: 30,
              enter: fade({ duration: 50 }),
            }),
          ),
        ),
      );
    }
    const result = await build(createElement(PointSequenceDeck));
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const chartEffects = motion.steps[0]?.effects.filter(
      ({ target }) => target.authorId === "point-sequence-chart",
    );
    assert.deepEqual(
      chartEffects?.map(({ startMs, endMs }) => [startMs, endMs]),
      [
        [0, 100],
        [100, 200],
        [200, 300],
        [300, 400],
        [400, 500],
        [500, 600],
        [600, 700],
      ],
    );
    assert.deepEqual(
      motion.steps[0]?.effects
        .filter(({ target }) => target.authorId === "point-sequence-follower")
        .map(({ startMs, endMs }) => [startMs, endMs]),
      [[725, 775]],
    );
  });

  it("associates chart builds with actual chart effects by target and group after native timing renumbering", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [10, 20] },
        { name: "South", values: [12, 18] },
      ],
    };
    const chart = (testId: string, x: number) =>
      createElement(Chart, {
        id: testId,
        name: testId,
        x,
        y: 10,
        width: 210,
        height: 200,
        groups: [
          {
            kind: "column" as const,
            data: {
              kind: "shared-categories" as const,
              categories: data.categories,
            },
            series: data.series,
          },
        ],
        enter: fade({ by: "series", duration: 400 }),
      });
    const result = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "chart-build-groups" },
          chart("first-chart", 10),
          chart("second-chart", 240),
        ),
      ),
    );

    const source = await validatePptx(result.pptx);
    assert.deepEqual(source.errors, []);
    assert.equal(source.slides[0]?.objects.length, 2);

    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = strFromU8(unzipSync(result.pptx)[slidePath]!);
    const declared = [
      ...slideXml.matchAll(/<p:bldGraphic spid="(\d+)" grpId="(\d+)">/gu),
    ].map((match) => ({
      targetId: Number(match[1]),
      groupId: Number(match[2]),
    }));
    assert.equal(declared.length, 2);
    const [first, second] = declared;
    assert.ok(first !== undefined && second !== undefined);
    assert.notEqual(first.targetId, second.targetId);
    assert.notEqual(first.groupId, second.groupId);

    const cases: readonly [string, Uint8Array][] = [
      [
        "renumbered timing nodes",
        rewriteSlideXml(result.pptx, (xml) => renumberTimingNodeIds(xml, 100)),
      ],
      [
        "shared group number across charts",
        rewriteSlideXml(result.pptx, (xml) =>
          xml.replaceAll(
            `grpId="${second.groupId}"`,
            `grpId="${first.groupId}"`,
          ),
        ),
      ],
      [
        "zero group number",
        rewriteSlideXml(result.pptx, (xml) =>
          xml.replaceAll(`grpId="${first.groupId}"`, 'grpId="0"'),
        ),
      ],
    ];
    for (const [label, pptx] of cases) {
      const read = await validatePptx(pptx);
      assert.deepEqual(read.errors, [], label);
      const effects = read.slides[0]?.animations?.[0]?.effects ?? [];
      assert.equal(effects.length, 6, label);
      assert.deepEqual(
        effects.map(({ nativeBuild }) => nativeBuild),
        Array.from({ length: 6 }, () => "series"),
        label,
      );
    }
  });

  it("rejects invalid chart build group associations and ranges", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [10, 20] },
        { name: "South", values: [12, 18] },
      ],
    };
    const chart = (testId: string, x: number) =>
      createElement(Chart, {
        id: testId,
        name: testId,
        x,
        y: 10,
        width: 210,
        height: 200,
        groups: [
          {
            kind: "column" as const,
            data: {
              kind: "shared-categories" as const,
              categories: data.categories,
            },
            series: data.series,
          },
        ],
        enter: fade({ by: "series", duration: 400 }),
      });
    const result = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "chart-build-negative" },
          createElement(Shape, {
            id: "not-a-chart",
            x: 240,
            y: 220,
            width: 40,
            height: 40,
          }),
          chart("first-chart", 10),
          chart("second-chart", 260),
        ),
      ),
    );

    const source = await validatePptx(result.pptx);
    assert.deepEqual(source.errors, []);
    const objects = source.slides[0]?.objects ?? [];
    const charts = objects.filter(({ kind }) => kind === "chart");
    const shape = objects.find(({ kind }) => kind === "shape");
    assert.equal(charts.length, 2);
    assert.ok(shape !== undefined);
    const [firstChart, secondChart] = charts;
    assert.ok(firstChart !== undefined && secondChart !== undefined);

    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = strFromU8(unzipSync(result.pptx)[slidePath]!);
    const buildElements = [
      ...slideXml.matchAll(/<p:bldGraphic\b[\s\S]*?<\/p:bldGraphic>/gu),
    ].map((match) => match[0]);
    assert.equal(buildElements.length, 2);
    const firstBuildElement = buildElements.find((element) =>
      element.includes(`spid="${firstChart.id}"`),
    );
    const secondBuildElement = buildElements.find((element) =>
      element.includes(`spid="${secondChart.id}"`),
    );
    assert.ok(firstBuildElement !== undefined);
    assert.ok(secondBuildElement !== undefined);
    const firstGroupMatch = firstBuildElement.match(/grpId="(\d+)"/u);
    const secondGroupMatch = secondBuildElement.match(/grpId="(\d+)"/u);
    assert.ok(firstGroupMatch !== null && secondGroupMatch !== null);
    const firstGroup = Number(firstGroupMatch[1]);
    const secondGroup = Number(secondGroupMatch[1]);
    assert.notEqual(firstGroup, secondGroup);
    const firstBuild = (spid: number, grpId: number): string =>
      `<p:bldGraphic spid="${spid}" grpId="${grpId}">`;

    const expectMutationError = async (
      label: string,
      mutate: (xml: string) => string,
      code: string,
      message: RegExp,
      absent: readonly string[] = [],
    ): Promise<void> => {
      const read = await validatePptx(rewriteSlideXml(result.pptx, mutate));
      assert.ok(
        read.errors.some(
          (error) => error.code === code && message.test(error.message),
        ),
        `${label}: ${JSON.stringify(read.errors)}`,
      );
      for (const absentCode of absent) {
        assert.equal(
          read.errors.some((error) => error.code === absentCode),
          false,
          `${label}: ${JSON.stringify(read.errors)}`,
        );
      }
    };

    await expectMutationError(
      "nonexistent target",
      (xml) =>
        xml.replace(
          firstBuildElement,
          firstBuildElement.replace(`spid="${firstChart.id}"`, 'spid="9999"'),
        ),
      "invalid-chart-build-group",
      /does not exist/u,
    );
    await expectMutationError(
      "non-chart target",
      (xml) =>
        xml.replace(
          firstBuildElement,
          firstBuildElement.replace(
            `spid="${firstChart.id}"`,
            `spid="${shape.id}"`,
          ),
        ),
      "invalid-chart-build-group",
      /not a chart object/u,
    );
    await expectMutationError(
      "unused group",
      (xml) =>
        xml.replace(
          firstBuildElement,
          firstBuildElement.replace(`grpId="${firstGroup}"`, 'grpId="777"'),
        ),
      "invalid-chart-build-group",
      /actual chart animation effect/u,
      ["invalid-animation-target-range"],
    );
    await expectMutationError(
      "effect uses another chart group",
      (xml) =>
        xml
          .replaceAll(`grpId="${firstGroup}"`, `grpId="${secondGroup}"`)
          .replace(
            firstBuild(firstChart.id, secondGroup),
            firstBuild(firstChart.id, firstGroup),
          ),
      "invalid-chart-build-group",
      /build declaration of the same chart group/u,
      ["invalid-animation-target-range"],
    );
    await expectMutationError(
      "missing declaration",
      (xml) => xml.replace(firstBuildElement, ""),
      "invalid-chart-build-group",
      /build declaration of the same chart group/u,
      ["invalid-animation-target-range"],
    );
    await expectMutationError(
      "duplicate pair",
      (xml) =>
        xml.replace(firstBuildElement, firstBuildElement + firstBuildElement),
      "duplicate-chart-build-group",
      /Duplicate build group/u,
    );
    await expectMutationError(
      "series range overflow",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="0" categoryIdx="-4"',
          'seriesIdx="3" categoryIdx="-4"',
        ),
      "invalid-animation-target-range",
      /exceeds series/u,
      ["invalid-chart-build-group"],
    );
    await expectMutationError(
      "build mode mismatch",
      (xml) =>
        xml.replace(
          firstBuildElement,
          firstBuildElement.replace('bld="series"', 'bld="category"'),
        ),
      "invalid-animation-target-range",
      /does not match native chart build mode/u,
      ["invalid-chart-build-group"],
    );
    await expectMutationError(
      "gridLegend sentinel",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="-3" categoryIdx="-3"',
          'seriesIdx="-2" categoryIdx="-3"',
        ),
      "invalid-animation-target-range",
      /native -3 sentinel/u,
      ["invalid-chart-build-group"],
    );
  });

  it("generates fade from public build declaration and reads back targets and timing from the same PPTX", async () => {
    const effect = fade({
      duration: 800,
      delay: 120,
      easing: { decelerate: 0.6 },
    });
    assert.deepEqual(effect, {
      type: "effect",
      kind: "entrance",
      preset: "fade",
      duration: 800,
      delay: 120,
      easing: { decelerate: 0.6 },
    });
    assert.equal("id" in effect, false);

    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "motion" },
          createElement(Shape, {
            id: "moving",
            name: "moving",
            x: 12,
            y: 18,
            width: 60,
            height: 40,
            enter: effect,
          }),
        ),
      ),
    );
    const target = result.report.slides[0]?.objects[0];
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(motion.steps, [
      {
        trigger: "auto",
        effects: [
          {
            target: { id: target?.id, name: "moving", authorId: "moving" },
            kind: "entrance",
            preset: "fade",
            durationMs: 800,
            delayMs: 120,
            easing: { decelerate: 0.6 },
            startMs: 120,
            endMs: 920,
          },
        ],
      },
    ]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.animations, [
      {
        trigger: "auto",
        effects: [
          {
            targetId: target?.id,
            timingId: 3,
            kind: "entrance",
            preset: "fade",
            durationMs: 800,
            delayMs: 120,
            easing: { decelerate: 0.6 },
          },
        ],
      },
    ]);

    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const slideXml = strFromU8(slidePart);
    assert.ok(slideXml.includes(`<p:spTgt spid="${target?.id}">`));
    assert.ok(
      slideXml.indexOf("<p:transition") < slideXml.indexOf("<p:timing>") ||
        !slideXml.includes("<p:transition"),
    );
  });

  it("writes readable target references for basic native object capabilities", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "native-targets" },
          createElement(Shape, {
            id: "shape-target",
            name: "shape-target",
            x: 0,
            y: 0,
            width: 16,
            height: 16,
            enter: fade(),
          }),
          createElement(Text, {
            id: "text-target",
            name: "text-target",
            x: 20,
            y: 0,
            width: 80,
            height: 24,
            emphasis: spin({ from: -90 }),
            children: "Text target",
          }),
          createElement(Image, {
            id: "image-target",
            name: "image-target",
            src: png,
            x: 110,
            y: 0,
            width: 16,
            height: 16,
            enter: wipe({ from: "right" }),
          }),
          createElement(
            Group,
            { id: "group-target", name: "group-target", enter: zoom() },
            createElement(Shape, {
              id: "group-member",
              x: 140,
              y: 0,
              width: 16,
              height: 16,
            }),
          ),
        ),
      ),
    );
    const observation = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      observation.steps[0]?.effects.map(({ target, preset }) => [
        target.authorId,
        preset,
      ]),
      [
        ["shape-target", "fade"],
        ["text-target", "spin"],
        ["image-target", "wipe"],
        ["group-target", "zoom"],
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, preset }) => [
        file.slides[0]?.objects
          .flatMap((object) => [object, ...(object.members ?? [])])
          .find(({ id }) => id === targetId)?.name,
        preset,
      ]),
      [
        ["shape-target", "fade"],
        ["text-target", "spin"],
        ["image-target", "wipe"],
        ["group-target", "zoom"],
      ],
    );
  });

  it("groups click steps by declaration identity, not merging auto-numbering with explicit numbering", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "steps" },
          createElement(Shape, {
            id: "automatic-entry",
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            enter: fade(),
          }),
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "explicit-first",
              x: 12,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 200, delay: 50 }),
            }),
          ),
          createElement(
            Step,
            {},
            createElement(Shape, {
              id: "generated-first",
              x: 24,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 300, delay: 10 }),
            }),
          ),
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "explicit-second",
              x: 36,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 400, delay: 100 }),
            }),
          ),
          createElement(
            Step,
            {},
            createElement(Shape, {
              id: "generated-second",
              x: 48,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 120, delay: 25 }),
            }),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(Shape, {
              id: "disabled",
              x: 60,
              y: 0,
              width: 10,
              height: 10,
              motion: false,
              enter: floatDown(),
            }),
            createElement(Shape, {
              id: "static",
              x: 72,
              y: 0,
              width: 10,
              height: 10,
            }),
            createElement(
              Group,
              { id: "disabled-group", motion: false },
              createElement(Shape, {
                id: "disabled-descendant",
                x: 84,
                y: 0,
                width: 10,
                height: 10,
                enter: fade(),
              }),
            ),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(Shape, {
              id: "explicit-after-ignored",
              x: 96,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 300, delay: 40 }),
            }),
          ),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps.map((step) => [
        step.trigger,
        step.number,
        step.effects.map(({ target, durationMs, delayMs, startMs, endMs }) => [
          target.authorId,
          durationMs,
          delayMs,
          startMs,
          endMs,
        ]),
      ]),
      [
        ["auto", undefined, [["automatic-entry", 500, 0, 0, 500]]],
        [
          "click",
          1,
          [
            ["explicit-first", 200, 50, 50, 250],
            ["explicit-second", 400, 100, 100, 500],
          ],
        ],
        ["click", undefined, [["generated-first", 300, 10, 10, 310]]],
        ["click", undefined, [["generated-second", 120, 25, 25, 145]]],
        ["click", 2, [["explicit-after-ignored", 300, 40, 40, 340]]],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.map((step) => [
        step.trigger,
        step.effects.map(({ targetId, durationMs, delayMs }) => [
          result.report.slides[0]?.objects.find(({ id }) => id === targetId)
            ?.authorId,
          durationMs,
          delayMs,
        ]),
      ]),
      [
        ["auto", [["automatic-entry", 500, 0]]],
        [
          "click",
          [
            ["explicit-first", 200, 50],
            ["explicit-second", 400, 100],
          ],
        ],
        ["click", [["generated-first", 300, 10]]],
        ["click", [["generated-second", 120, 25]]],
        ["click", [["explicit-after-ignored", 300, 40]]],
      ],
    );
  });

  it("writes eight common basic prototype effects and applicable direction, rotation, and easing parameters", async () => {
    const effects = [
      ["fade", fade({ easing: { accelerate: 0.2, decelerate: 0.3 } }), "enter"],
      ["wipe", wipe({ from: "left" }), "enter"],
      ["zoom", zoom(), "enter"],
      ["fly", fly({ from: "top" }), "enter"],
      ["floatUp", floatUp(), "enter"],
      ["spinIn", spinIn(), "enter"],
      ["stretch", stretch(), "enter"],
      ["spin", spin({ from: 225, easing: { decelerate: 1 } }), "emphasis"],
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "presets" },
          ...effects.map(([testId, effect, prop], index) =>
            createElement(Shape, {
              key: testId,
              id: testId,
              x: index * 20,
              y: 0,
              width: 16,
              height: 16,
              [prop]: effect,
            }),
          ),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, preset, from, easing }) => [
        target.authorId,
        preset,
        from,
        easing,
      ]),
      [
        ["fade", "fade", undefined, { accelerate: 0.2, decelerate: 0.3 }],
        ["wipe", "wipe", "left", undefined],
        ["zoom", "zoom", undefined, undefined],
        ["fly", "fly", "top", undefined],
        ["floatUp", "floatUp", undefined, undefined],
        ["spinIn", "spinIn", undefined, undefined],
        ["stretch", "stretch", undefined, undefined],
        ["spin", "spin", 225, { decelerate: 1 }],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ preset, from, easing }) => [preset, from, easing],
      ),
      [
        ["fade", undefined, { accelerate: 0.2, decelerate: 0.3 }],
        ["wipe", "left", undefined],
        ["zoom", undefined, undefined],
        ["fly", "top", undefined],
        ["floatUp", undefined, undefined],
        ["spinIn", undefined, undefined],
        ["stretch", undefined, undefined],
        ["spin", 225, { decelerate: 1 }],
      ],
    );
  });

  it("fades in floatDown from above the object and preserves downward displacement in the file", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "float-down" },
          createElement(Shape, {
            id: "down-default",
            x: 20,
            y: 20,
            width: 24,
            height: 24,
            fill: "#164E3D",
            enter: floatDown(),
          }),
          createElement(Shape, {
            id: "down-custom",
            x: 60,
            y: 20,
            width: 24,
            height: 24,
            fill: "#164E3D",
            enter: floatDown({
              duration: 250,
              delay: 40,
              easing: { decelerate: 0.5 },
            }),
          }),
          createElement(Shape, {
            id: "up-regression",
            x: 100,
            y: 20,
            width: 24,
            height: 24,
            fill: "#164E3D",
            enter: floatUp(),
          }),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(
        ({ target, preset, durationMs, delayMs }) => [
          target.authorId,
          preset,
          durationMs,
          delayMs,
        ],
      ),
      [
        ["down-default", "floatDown", 500, 0],
        ["up-regression", "floatUp", 500, 0],
        ["down-custom", "floatDown", 250, 40],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileEffects = file.slides[0]?.animations?.[0]?.effects;
    assert.deepEqual(
      fileEffects?.map(({ preset, durationMs, delayMs, easing }) => [
        preset,
        durationMs,
        delayMs,
        easing,
      ]),
      [
        ["floatDown", 500, 0, undefined],
        ["floatUp", 500, 0, undefined],
        ["floatDown", 250, 40, { decelerate: 0.5 }],
      ],
    );
    const downTargetId = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "down-default",
    )?.id;
    const upTargetId = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "up-regression",
    )?.id;
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      slideXml,
      /presetID="47" presetClass="entr" presetSubtype="0"/u,
    );
    const positionAnimation = (targetId: number | undefined, axis: string) =>
      [...slideXml.matchAll(/<p:anim\b[\s\S]*?<\/p:anim>/gu)].find(
        ([behavior]) =>
          behavior?.includes(`<p:spTgt spid="${targetId}">`) &&
          behavior.includes(`<p:attrName>${axis}</p:attrName>`),
      )?.[0];
    const downX = positionAnimation(downTargetId, "ppt_x");
    const downY = positionAnimation(downTargetId, "ppt_y");
    assert.ok(downX);
    assert.ok(downY);
    assert.deepEqual(
      [...downX.matchAll(/<p:strVal val="([^"]+)"><\/p:strVal>/gu)].map(
        ([, value]) => value,
      ),
      ["#ppt_x", "#ppt_x"],
    );
    assert.deepEqual(
      [...downY.matchAll(/<p:strVal val="([^"]+)"><\/p:strVal>/gu)].map(
        ([, value]) => value,
      ),
      ["#ppt_y-.1", "#ppt_y"],
    );
    const downNativeBehaviors = [
      ...slideXml.matchAll(/<p:(set|animEffect|anim)\b[\s\S]*?<\/p:\1>/gu),
    ]
      .map(([behavior]) => behavior ?? "")
      .filter((behavior) =>
        behavior.includes(`<p:spTgt spid="${downTargetId}">`),
      );
    assert.deepEqual(
      downNativeBehaviors.map((behavior) => /^<p:(\w+)/u.exec(behavior)?.[1]),
      ["set", "animEffect", "anim", "anim"],
    );
    assert.match(downNativeBehaviors[0] ?? "", /style\.visibility/u);
    assert.match(downNativeBehaviors[0] ?? "", /val="visible"/u);
    assert.match(downNativeBehaviors[1] ?? "", /filter="fade"/u);
    const upY = positionAnimation(upTargetId, "ppt_y");
    assert.ok(upY);
    assert.match(upY, /<p:strVal val="#ppt_y\+\.1"><\/p:strVal>/u);
  });

  it("fades downward floatDown for groups and whole Text with background together", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "float-down-targets" },
          createElement(
            Group,
            {
              id: "group-down",
              enter: floatDown({ duration: 250, delay: 35 }),
            },
            createElement(Shape, {
              id: "group-member",
              x: 20,
              y: 20,
              width: 24,
              height: 24,
              fill: "#164E3D",
            }),
          ),
          createElement(
            Text,
            {
              id: "whole-down",
              x: 80,
              y: 20,
              width: 180,
              height: 100,
              style: {
                backgroundColor: "#38BDF8",
                borderColor: "#1D4ED8",
                borderWidth: 2,
              },
              enter: floatDown({ duration: 250 }),
            },
            createElement(P, { key: "first" }, "First paragraph"),
            createElement(P, { key: "second" }, "Second paragraph"),
          ),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const reportObjects = result.report.slides[0]?.objects ?? [];
    const groupId = reportObjects.find(
      ({ authorId: testId }) => testId === "group-down",
    )?.id;
    const textId = reportObjects.find(
      ({ authorId: testId }) => testId === "whole-down",
    )?.id;
    assert.ok(groupId !== undefined);
    assert.ok(textId !== undefined);

    const fileEffects =
      file.slides[0]?.animations?.flatMap(({ effects }) => effects) ?? [];
    const groupEffect = fileEffects.find(
      ({ targetId }) => targetId === groupId,
    );
    assert.deepEqual(
      [groupEffect?.preset, groupEffect?.durationMs, groupEffect?.delayMs],
      ["floatDown", 250, 35],
    );

    const textEffect = fileEffects.find(({ targetId }) => targetId === textId);
    assert.deepEqual(
      [
        textEffect?.preset,
        textEffect?.nativeBuild,
        textEffect?.nativeBuildGroupId,
        textEffect?.nativeBuildAnimBg,
        textEffect?.durationMs,
      ],
      ["floatDown", "paragraph", 1, true, 250],
    );
    const textPositionYs = textEffect?.nativeBehaviors?.filter(
      ({ behavior, targetType, attributeNames }) =>
        behavior === "anim" &&
        targetType === "paragraph" &&
        attributeNames[0] === "ppt_y",
    );
    assert.deepEqual(
      textPositionYs?.map(({ targetRange, from, to, durationMs }) => [
        targetRange,
        from,
        to,
        durationMs,
      ]),
      [
        [{ type: "paragraph", start: 0, end: 0 }, "-0.1", "0", 250],
        [{ type: "paragraph", start: 1, end: 1 }, "-0.1", "0", 250],
      ],
    );
    const textBackgroundY = textEffect?.nativeBehaviors?.find(
      ({ behavior, targetType, attributeNames }) =>
        behavior === "anim" &&
        targetType === "background" &&
        attributeNames[0] === "ppt_y",
    );
    assert.deepEqual(
      [textBackgroundY?.from, textBackgroundY?.to, textBackgroundY?.durationMs],
      ["-0.1", "0", 250],
    );
    const textPositionX = textEffect?.nativeBehaviors?.find(
      ({ behavior, targetType, attributeNames }) =>
        behavior === "anim" &&
        targetType === "paragraph" &&
        attributeNames[0] === "ppt_x",
    );
    assert.deepEqual([textPositionX?.from, textPositionX?.to], ["0", "0"]);
    assert.ok(
      textEffect?.nativeBehaviors?.some(
        ({ behavior, targetType, filter }) =>
          behavior === "animEffect" &&
          targetType === "background" &&
          filter === "fade",
      ),
    );

    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const groupY = [...slideXml.matchAll(/<p:anim\b[\s\S]*?<\/p:anim>/gu)].find(
      ([behavior]) =>
        behavior?.includes(`<p:spTgt spid="${groupId}">`) &&
        behavior.includes("<p:attrName>ppt_y</p:attrName>"),
    )?.[0];
    assert.ok(groupY);
    assert.deepEqual(
      [...groupY.matchAll(/<p:strVal val="([^"]+)"><\/p:strVal>/gu)].map(
        ([, value]) => value,
      ),
      ["#ppt_y-.1", "#ppt_y"],
    );
  });

  it("expands whole Fly on the same native Text into shared paragraph displacement and reads back background role", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { id: "whole-text-projection" },
          createElement(Text, {
            id: "mixed-text",
            x: 800,
            y: 500,
            width: 400,
            height: 160,
            style: {
              backgroundColor: "#38BDF8",
              borderColor: "#1D4ED8",
              borderWidth: 4,
            },
            enter: [
              fly({ from: "left", duration: 4000 }),
              fade({ by: "paragraph", duration: 700, delay: 1500 }),
            ],
            children: [
              createElement(
                P,
                { key: "long" },
                "A deliberately long first paragraph",
              ),
              createElement(P, { key: "short" }, "SHORT"),
            ],
          }),
        ),
      ),
    );
    const textObject = result.report.slides[0]?.objects[0];
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, preset, by, endMs }) => [
        target.authorId,
        preset,
        by,
        endMs,
      ]),
      [
        ["mixed-text", "fly", undefined, 4000],
        ["mixed-text", "fade", "paragraph", 2200],
        ["mixed-text", "fade", "paragraph", 2900],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const readback = file.slides[0]?.animations?.flatMap(
      (step) => step.effects,
    );
    const flyEffect = readback?.find(({ preset }) => preset === "fly");
    const shortFade = readback?.find(
      ({ preset, delayMs }) => preset === "fade" && delayMs === 2200,
    );
    assert.equal(flyEffect?.targetId, textObject?.id);
    assert.deepEqual(
      [
        flyEffect?.nativeBuild,
        flyEffect?.nativeBuildGroupId,
        flyEffect?.nativeBuildAnimBg,
      ],
      ["paragraph", 1, true],
    );
    assert.deepEqual(
      flyEffect?.nativeBehaviors?.map(
        ({ behavior, targetType, targetRange, attributeNames, durationMs }) => [
          behavior,
          targetType,
          targetRange,
          attributeNames,
          durationMs,
        ],
      ),
      [
        [
          "set",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["style.visibility"],
          1,
        ],
        [
          "set",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["style.visibility"],
          1,
        ],
        ["set", "background", undefined, ["style.visibility"], 1],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["ppt_x"],
          4000,
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["ppt_y"],
          4000,
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["ppt_x"],
          4000,
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["ppt_y"],
          4000,
        ],
        ["anim", "background", undefined, ["ppt_x"], 4000],
        ["anim", "background", undefined, ["ppt_y"], 4000],
      ],
    );
    assert.equal(shortFade?.delayMs, 2200);
    assert.deepEqual(
      [shortFade?.nativeBuild, shortFade?.nativeBuildGroupId],
      ["paragraph", 0],
    );
    assert.equal(file.slides[0]?.objects.length, 1);

    const slidePath = "ppt/slides/slide1.xml";
    const slideXml = strFromU8(unzipSync(result.pptx)[slidePath]!);
    const invalidBuildGroup = slideXml.replace(
      `<p:bldP spid="${textObject?.id}" grpId="1" build="p" animBg="1"></p:bldP>`,
      `<p:bldP spid="${textObject?.id}" grpId="2" build="p" animBg="1"></p:bldP>`,
    );
    assert.notEqual(invalidBuildGroup, slideXml);
    const wrongGroupFile = await validatePptx(
      zipSync({
        ...unzipSync(result.pptx),
        [slidePath]: strToU8(invalidBuildGroup),
      }),
    );
    assert.ok(
      wrongGroupFile.errors.some(
        ({ code }) =>
          code === "inconsistent-animation-target-range" ||
          code === "invalid-animation-target-range",
      ),
    );

    const missingParagraphRole = slideXml.replace(
      "<p:attrName>ppt_y</p:attrName>",
      "<p:attrName>ppt_unexpected</p:attrName>",
    );
    assert.notEqual(missingParagraphRole, slideXml);
    const missingRoleFile = await validatePptx(
      zipSync({
        ...unzipSync(result.pptx),
        [slidePath]: strToU8(missingParagraphRole),
      }),
    );
    assert.ok(
      missingRoleFile.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const missingBackgroundRole = slideXml.replace(
      `<p:bldP spid="${textObject?.id}" grpId="1" build="p" animBg="1"></p:bldP>`,
      `<p:bldP spid="${textObject?.id}" grpId="1" build="p" animBg="0"></p:bldP>`,
    );
    assert.notEqual(missingBackgroundRole, slideXml);
    const missingBackgroundFile = await validatePptx(
      zipSync({
        ...unzipSync(result.pptx),
        [slidePath]: strToU8(missingBackgroundRole),
      }),
    );
    assert.ok(
      missingBackgroundFile.errors.some(
        ({ code }) => code === "inconsistent-animation-background-role",
      ),
    );

    const invalidBehaviorDuration = slideXml.replace(
      /(<p:anim\b[^>]*><p:cBhvr\b[^>]*><p:cTn id="\d+" dur=")4000("[^>]*>)/u,
      (_match, prefix: string, suffix: string) => `${prefix}3999${suffix}`,
    );
    assert.notEqual(invalidBehaviorDuration, slideXml);
    const inconsistentDurationFile = await validatePptx(
      zipSync({
        ...unzipSync(result.pptx),
        [slidePath]: strToU8(invalidBehaviorDuration),
      }),
    );
    assert.ok(
      inconsistentDurationFile.errors.some(
        ({ code }) => code === "inconsistent-animation-behavior-duration",
      ),
    );
  });

  it("preserves background role of Fly in shared text build group of whole array", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { id: "whole-text-background-array" },
          createElement(Text, {
            id: "filled-whole-array",
            x: 800,
            y: 500,
            width: 400,
            height: 160,
            style: {
              backgroundColor: "#38BDF8",
              borderColor: "#1D4ED8",
              borderWidth: 4,
            },
            enter: [fly({ from: "left", duration: 4000 }), fade()],
            children: [
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.flatMap(({ effects }) =>
        effects.map(({ preset, nativeBuild, nativeBuildAnimBg }) => [
          preset,
          nativeBuild,
          nativeBuildAnimBg,
        ]),
      ),
      [
        ["fly", "paragraph", true],
        ["fade", "paragraph", true],
      ],
    );
  });

  it("projects whole Text filter and SpinIn size to actual background", async () => {
    const scenarios = [
      {
        preset: "fade",
        effect: fade({ duration: 4000 }),
        id: "fade-background",
        roles: ["animEffect:fade"],
      },
      {
        preset: "wipe",
        effect: wipe({ from: "left", duration: 4000 }),
        id: "wipe-background",
        roles: ["animEffect:wipe(left)"],
      },
      {
        preset: "zoom",
        effect: zoom({ duration: 4000 }),
        id: "zoom-background",
        roles: [
          "anim:ppt_x",
          "anim:ppt_y",
          "anim:ppt_w",
          "anim:ppt_h",
          "animEffect:fade",
        ],
      },
      {
        preset: "floatUp",
        effect: floatUp({ duration: 4000 }),
        id: "float-up-background",
        roles: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
      },
      {
        preset: "spinIn",
        effect: spinIn({ duration: 4000 }),
        id: "spin-in-background",
        roles: [
          "anim:ppt_x",
          "anim:ppt_y",
          "anim:ppt_w",
          "anim:ppt_h",
          "animRot:r",
          "animEffect:fade",
        ],
      },
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        ...scenarios.map(({ preset, effect, id: slideId }) =>
          createElement(
            Slide,
            { key: preset, id: slideId },
            createElement(Text, {
              id: `${preset}-text`,
              x: 800,
              y: 500,
              width: 400,
              height: 160,
              style: {
                backgroundColor: "#38BDF8",
                borderColor: "#1D4ED8",
                borderWidth: 4,
              },
              enter: effect,
              children: [
                createElement(P, { key: "long" }, "A long first paragraph"),
                createElement(P, { key: "short" }, "SHORT"),
              ],
            }),
          ),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) => {
        const effect = slide.animations?.[0]?.effects[0];
        const backgroundBehaviors = effect?.nativeBehaviors?.filter(
          ({ behavior, targetType }) =>
            targetType === "background" && behavior !== "set",
        );
        return [
          effect?.preset,
          effect?.nativeBuildAnimBg,
          backgroundBehaviors?.map(({ behavior, attributeNames, filter }) =>
            behavior === "animEffect"
              ? `${behavior}:${filter}`
              : `${behavior}:${attributeNames[0]}`,
          ),
        ];
      }),
      scenarios.map(({ preset, roles }) => [preset, true, roles]),
    );

    const removeBackgroundBehavior = (
      xml: string,
      targetId: number,
      behavior: "anim" | "animEffect",
      marker: string,
    ) => {
      const target = `<p:spTgt spid="${targetId}"><p:bg></p:bg></p:spTgt>`;
      const selected = [
        ...xml.matchAll(/<p:(anim|animEffect|animRot)\b[\s\S]*?<\/p:\1>/gu),
      ].find(
        ([nativeBehavior]) =>
          nativeBehavior?.includes(target) && nativeBehavior.includes(marker),
      )?.[0];
      assert.ok(
        selected !== undefined && selected.startsWith(`<p:${behavior}`),
      );
      return xml.replace(selected, "");
    };
    const originalParts = unzipSync(result.pptx);
    const spinInSlide = 4;
    const spinInObjectId = result.report.slides[spinInSlide]?.objects[0]?.id;
    assert.ok(spinInObjectId !== undefined);
    const spinInPath = `ppt/slides/slide${spinInSlide + 1}.xml`;
    const spinInXml = strFromU8(originalParts[spinInPath]!);
    const missingSpinInWidth = removeBackgroundBehavior(
      spinInXml,
      spinInObjectId,
      "anim",
      "<p:attrName>ppt_w</p:attrName>",
    );
    const invalidSpinInFile = await validatePptx(
      zipSync({ ...originalParts, [spinInPath]: strToU8(missingSpinInWidth) }),
    );
    assert.ok(
      invalidSpinInFile.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const fadeObjectId = result.report.slides[0]?.objects[0]?.id;
    assert.ok(fadeObjectId !== undefined);
    const fadePath = "ppt/slides/slide1.xml";
    const fadeXml = strFromU8(originalParts[fadePath]!);
    const missingFadeFilter = removeBackgroundBehavior(
      fadeXml,
      fadeObjectId,
      "animEffect",
      'filter="fade"',
    );
    const invalidFadeFile = await validatePptx(
      zipSync({ ...originalParts, [fadePath]: strToU8(missingFadeFilter) }),
    );
    assert.ok(
      invalidFadeFile.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
  });

  it("expands whole Stretch from common text center and reads back paragraph and background displacement", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { id: "whole-stretch-common-center" },
          createElement(Text, {
            id: "stretched-text",
            x: 800,
            y: 500,
            width: 400,
            height: 160,
            style: {
              backgroundColor: "#38BDF8",
              borderColor: "#1D4ED8",
              borderWidth: 4,
            },
            enter: stretch({ duration: 2000 }),
            children: [
              createElement(
                P,
                { key: "long" },
                "A deliberately long first paragraph",
              ),
              createElement(P, { key: "short" }, "SHORT"),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effect = file.slides[0]?.animations?.[0]?.effects[0];
    assert.equal(effect?.preset, "stretch");
    assert.deepEqual(
      effect?.nativeBehaviors
        ?.filter(({ attributeNames }) => attributeNames.includes("ppt_x"))
        .map(({ targetType, targetRange, from, to, additive }) => [
          targetType,
          targetRange,
          from,
          to,
          additive,
        ]),
      [
        [
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          "0.78125-#ppt_x",
          "0",
          "sum",
        ],
        [
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          "0.78125-#ppt_x",
          "0",
          "sum",
        ],
        ["background", undefined, "0.78125-#ppt_x", "0", "sum"],
      ],
    );

    const textId = result.report.slides[0]?.objects[0]?.id;
    assert.ok(textId !== undefined);
    const path = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const xml = strFromU8(parts[path]!);
    const target = `<p:spTgt spid="${textId}"><p:txEl><p:pRg st="0" end="0"></p:pRg></p:txEl></p:spTgt>`;
    const missingOffset = [
      ...xml.matchAll(/<p:anim\b[\s\S]*?<\/p:anim>/gu),
    ].find(
      ([nativeBehavior]) =>
        nativeBehavior?.includes(target) &&
        nativeBehavior.includes("<p:attrName>ppt_x</p:attrName>"),
    )?.[0];
    assert.ok(missingOffset !== undefined);
    const invalidFile = await validatePptx(
      zipSync({
        ...parts,
        [path]: strToU8(xml.replace(missingOffset, "")),
      }),
    );
    assert.ok(
      invalidFile.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
  });

  it("writes a whole-object entrance for a shape whose own text has a visible background", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "whole-shape-entrance" },
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              id: "whole-shape-entry",
              x: 120,
              y: 140,
              width: 360,
              height: 96,
              fill: "#0070C0",
              text: "Whole shape entrance",
              textStyle: { fontSize: 20, color: "#FFFFFF" },
              enter: fade({ duration: 300 }),
            }),
          ),
        ),
      ),
    );
    const shape = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "whole-shape-entry",
    );
    assert.ok(shape !== undefined);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, preset }) => [
        target.authorId,
        preset,
      ]),
      [["whole-shape-entry", "fade"]],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effect = file.slides[0]?.animations?.[0]?.effects[0];
    assert.deepEqual(
      [
        effect?.targetId,
        effect?.targetRange,
        effect?.preset,
        effect?.kind,
        effect?.durationMs,
      ],
      [shape.id, undefined, "fade", "entrance", 300],
    );
    // Native whole-shape Fade carries the background inside the object target instead of a separate p:bg branch.
    assert.deepEqual(
      [
        effect?.nativeBuild,
        effect?.nativeBuildGroupId,
        effect?.nativeBuildAnimBg,
      ],
      [undefined, 0, true],
    );

    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const wholeTarget = `<p:spTgt spid="${shape.id}"></p:spTgt>`;
    const wholeBehaviors = [
      ...slideXml.matchAll(/<p:(set|animEffect)\b[^>]*>[\s\S]*?<\/p:\1>/gu),
    ]
      .map(([behavior]) => behavior ?? "")
      .filter((behavior) => behavior.includes(wholeTarget));
    assert.deepEqual(
      wholeBehaviors.map((behavior) =>
        behavior.startsWith("<p:set") ? "set" : "animEffect",
      ),
      ["set", "animEffect"],
    );
    assert.doesNotMatch(
      slideXml,
      new RegExp(`<p:spTgt spid="${shape.id}"><p:txEl>`, "u"),
    );
    assert.doesNotMatch(
      slideXml,
      new RegExp(`<p:spTgt spid="${shape.id}"><p:bg>`, "u"),
    );
    assert.match(
      slideXml,
      new RegExp(
        `<p:bldP spid="${shape.id}" grpId="0" animBg="1"></p:bldP>`,
        "u",
      ),
    );
  });

  it("keeps whole Fade and Wipe without geometry companions when paragraph effects exist", async () => {
    for (const [id, whole] of [
      ["mixed-fade", fade({ duration: 4000 })],
      ["mixed-wipe", wipe({ duration: 4000 })],
    ] as const) {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id },
            createElement(
              Step,
              { n: 1 },
              createElement(
                Shape,
                {
                  id: `${id}-shape`,
                  x: 100,
                  y: 100,
                  width: 400,
                  height: 160,
                  fill: "#0070C0",
                  textStyle: { fontSize: 20, color: "#FFFFFF" },
                  enter: [
                    whole,
                    fade({ by: "paragraph", duration: 700, delay: 1500 }),
                  ],
                },
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ),
            ),
          ),
        ),
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], id);
      const effects =
        file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
      assert.equal(effects.length, 3, id);
      assert.equal(
        effects.filter(
          ({ preset }) => preset === (id === "mixed-fade" ? "fade" : "wipe"),
        ).length,
        id === "mixed-fade" ? 3 : 1,
        id,
      );
      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      assert.doesNotMatch(slideXml, /<p:childTnLst><\/p:childTnLst>/u);
      const wholeBlock = animationEffectBlocks(slideXml).find(
        ({ block, presetId }) =>
          (presetId === 10 || presetId === 22) &&
          block.includes('nodeType="clickEffect"'),
      );
      assert.ok(wholeBlock !== undefined);
      assert.doesNotMatch(wholeBlock.block, /<p:pRg/u);
    }
  });

  it("splits a mixed whole Shape Fly into one synchronized geometry companion per paragraph", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "mixed-fly" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "mixed-fly-shape",
                x: 100,
                y: 100,
                width: 400,
                height: 160,
                fill: "#0070C0",
                textStyle: { fontSize: 20, color: "#FFFFFF" },
                enter: [
                  fly({ duration: 4000 }),
                  fade({ by: "paragraph", duration: 700, delay: 1500 }),
                ],
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    const shapeId = result.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "mixed-fly-shape",
    )?.id;
    assert.ok(shapeId !== undefined);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, preset, by, endMs }) => [
        target.authorId,
        preset,
        by,
        endMs,
      ]),
      [
        ["mixed-fly-shape", "fly", undefined, 4000],
        ["mixed-fly-shape", "fade", "paragraph", 2200],
        ["mixed-fly-shape", "fade", "paragraph", 2900],
      ],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects =
      file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
    const flyEffects = effects.filter(({ preset }) => preset === "fly");
    assert.equal(flyEffects.length, 3);
    const whole = flyEffects.find(
      ({ targetRange }) => targetRange === undefined,
    );
    assert.ok(whole !== undefined);
    assert.deepEqual(
      [
        whole.nativeBuild,
        whole.nativeBuildGroupId,
        whole.nativeBuildAnimBg,
        whole.delayMs,
        whole.durationMs,
      ],
      ["paragraph", 0, true, 0, 4000],
    );
    const companions = flyEffects
      .filter(({ targetRange }) => targetRange?.type === "paragraph")
      .sort(
        (left, right) =>
          (left.targetRange?.type === "paragraph"
            ? left.targetRange.start
            : 0) -
          (right.targetRange?.type === "paragraph"
            ? right.targetRange.start
            : 0),
      );
    assert.deepEqual(
      companions.map(
        ({
          targetRange,
          delayMs,
          durationMs,
          nativeBuild,
          nativeBuildGroupId,
          nativeBuildAnimBg,
        }) => [
          targetRange,
          delayMs,
          durationMs,
          nativeBuild,
          nativeBuildGroupId,
          nativeBuildAnimBg,
        ],
      ),
      [
        [
          { type: "paragraph", start: 0, end: 0 },
          0,
          4000,
          "paragraph",
          0,
          true,
        ],
        [
          { type: "paragraph", start: 1, end: 1 },
          0,
          4000,
          "paragraph",
          0,
          true,
        ],
      ],
    );

    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const blocks = animationEffectBlocks(slideXml);
    const flyBlocks = blocks.filter(({ presetId }) => presetId === 2);
    assert.equal(flyBlocks.length, 3);
    const wholeBlock = flyBlocks.find(({ block }) =>
      block.includes('nodeType="clickEffect"'),
    );
    assert.ok(wholeBlock !== undefined);
    const companionBlocks = flyBlocks
      .filter(({ block }) => block.includes('nodeType="withEffect"'))
      .sort(
        (left, right) =>
          Number(/<p:pRg st="(\d+)"/u.exec(left.block)?.[1]) -
          Number(/<p:pRg st="(\d+)"/u.exec(right.block)?.[1]),
      );
    assert.equal(companionBlocks.length, 2);
    const fadeBlocks = blocks.filter(({ presetId }) => presetId === 10);
    assert.equal(fadeBlocks.length, 2);
    const order = [
      wholeBlock.block,
      fadeBlocks[0]?.block ?? "",
      fadeBlocks[1]?.block ?? "",
      companionBlocks[0]?.block ?? "",
      companionBlocks[1]?.block ?? "",
    ].map((block) => slideXml.indexOf(block));
    assert.deepEqual(
      order,
      [...order].sort((left, right) => left - right),
    );

    const attributes = (block: string): string[] =>
      [...block.matchAll(/<p:attrName>([^<]+)<\/p:attrName>/gu)].map(
        ([, name]) => name ?? "",
      );
    assert.deepEqual(attributes(wholeBlock.block), [
      "style.visibility",
      "ppt_x",
      "ppt_y",
    ]);
    assert.doesNotMatch(wholeBlock.block, /<p:pRg/u);
    assert.deepEqual(
      companionBlocks.map(({ block }) => ({
        range: /<p:pRg st="(\d+)" end="(\d+)">/u.exec(block)?.slice(1),
        attributes: attributes(block),
        set: block.includes("<p:set>"),
      })),
      [
        { range: ["0", "0"], attributes: ["ppt_x", "ppt_y"], set: false },
        { range: ["1", "1"], attributes: ["ppt_x", "ppt_y"], set: false },
      ],
    );
    assert.match(
      companionBlocks[0]?.block ?? "",
      /#ppt_y\+0\.814814814814815/u,
    );
    assert.doesNotMatch(slideXml, /<p:bg>/u);

    const removeOne = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(companionBlocks[0]?.block ?? "", ""),
    );
    assert.ok(
      (await validatePptx(removeOne)).errors.some(
        ({ code }) => code === "inconsistent-animation-target-range",
      ),
    );
    const removeAll = rewriteSlideXml(result.pptx, (xml) =>
      xml
        .replace(companionBlocks[0]?.block ?? "", "")
        .replace(companionBlocks[1]?.block ?? "", ""),
    );
    const removeAllErrors = (await validatePptx(removeAll)).errors;
    assert.equal(
      removeAllErrors.filter(
        ({ code }) => code === "inconsistent-animation-target-range",
      ).length,
      2,
    );
    const wrongDelay = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companionBlocks[0]?.block ?? "",
        (companionBlocks[0]?.block ?? "").replace(
          '<p:cond delay="0">',
          '<p:cond delay="50">',
        ),
      ),
    );
    assert.ok(
      (await validatePptx(wrongDelay)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
    const wrongRepeat = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companionBlocks[0]?.block ?? "",
        (companionBlocks[0]?.block ?? "").replace(
          '<p:cTn id="',
          '<p:cTn repeatCount="2000" id="',
        ),
      ),
    );
    assert.ok(
      (await validatePptx(wrongRepeat)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
    const wrongAxis = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companionBlocks[0]?.block ?? "",
        (companionBlocks[0]?.block ?? "").replace(
          "<p:attrName>ppt_y</p:attrName>",
          "<p:attrName>ppt_x</p:attrName>",
        ),
      ),
    );
    assert.ok(
      (await validatePptx(wrongAxis)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
    const backgroundPrimary = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        wholeBlock.block,
        wholeBlock.block.replaceAll(
          `<p:spTgt spid="${shapeId}"></p:spTgt>`,
          `<p:spTgt spid="${shapeId}"><p:bg></p:bg></p:spTgt>`,
        ),
      ),
    );
    assert.deepEqual((await validatePptx(backgroundPrimary)).errors, []);
    const mixedPrimaryRoles = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        wholeBlock.block,
        wholeBlock.block.replace(
          `<p:spTgt spid="${shapeId}"></p:spTgt>`,
          `<p:spTgt spid="${shapeId}"><p:bg></p:bg></p:spTgt>`,
        ),
      ),
    );
    assert.ok(
      (await validatePptx(mixedPrimaryRoles)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
    const wrongWholeAxis = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        wholeBlock.block,
        wholeBlock.block.replace(
          "<p:attrName>ppt_y</p:attrName>",
          "<p:attrName>ppt_x</p:attrName>",
        ),
      ),
    );
    assert.ok(
      (await validatePptx(wrongWholeAxis)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
  });

  it("keeps whole Stretch and SpinIn roles on the object while companions carry paragraph geometry", async () => {
    const cases = [
      {
        id: "mixed-stretch",
        effect: stretch({ duration: 4000 }),
        preset: "stretch",
        wholeAttributes: ["style.visibility", "ppt_w", "ppt_h"],
        companionAttributes: ["ppt_x", "ppt_w", "ppt_h"],
      },
      {
        id: "mixed-spin-in",
        effect: spinIn({ duration: 4000 }),
        preset: "spinIn",
        wholeAttributes: [
          "style.visibility",
          "ppt_w",
          "ppt_h",
          "style.rotation",
        ],
        companionAttributes: [
          "ppt_w",
          "ppt_h",
          "style.rotation",
          "ppt_x",
          "ppt_y",
        ],
      },
    ];
    for (const item of cases) {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: item.id },
            createElement(
              Step,
              { n: 1 },
              createElement(
                Shape,
                {
                  id: `${item.id}-shape`,
                  x: 100,
                  y: 100,
                  width: 400,
                  height: 160,
                  fill: "#0070C0",
                  textStyle: { fontSize: 20, color: "#FFFFFF" },
                  enter: [
                    item.effect,
                    fade({ by: "paragraph", duration: 700, delay: 1500 }),
                  ],
                },
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ),
            ),
          ),
        ),
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], item.id);
      const effects =
        file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
      const geometryEffects = effects.filter(
        ({ preset }) => preset === item.preset,
      );
      assert.equal(geometryEffects.length, 3, item.id);
      assert.equal(
        geometryEffects.filter(({ targetRange }) => targetRange === undefined)
          .length,
        1,
        item.id,
      );
      assert.deepEqual(
        geometryEffects
          .filter(({ targetRange }) => targetRange?.type === "paragraph")
          .map(({ targetRange }) => targetRange),
        [
          { type: "paragraph", start: 0, end: 0 },
          { type: "paragraph", start: 1, end: 1 },
        ],
        item.id,
      );
      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      const blocks = animationEffectBlocks(slideXml);
      const geometryBlocks = blocks.filter(
        ({ presetId }) => presetId === (item.preset === "stretch" ? 17 : 49),
      );
      assert.equal(geometryBlocks.length, 3, item.id);
      const wholeBlock = geometryBlocks.find(({ block }) =>
        block.includes('nodeType="clickEffect"'),
      );
      assert.ok(wholeBlock !== undefined);
      const companionBlocks = geometryBlocks.filter(({ block }) =>
        block.includes('nodeType="withEffect"'),
      );
      assert.equal(companionBlocks.length, 2, item.id);
      const attributes = (block: string): string[] =>
        [...block.matchAll(/<p:attrName>([^<]+)<\/p:attrName>/gu)].map(
          ([, name]) => name ?? "",
        );
      assert.deepEqual(attributes(wholeBlock.block), item.wholeAttributes);
      for (const companion of companionBlocks) {
        assert.deepEqual(attributes(companion.block), item.companionAttributes);
        assert.doesNotMatch(companion.block, /<p:set>/u);
      }
      if (item.preset === "spinIn") {
        const spinCompanion = companionBlocks[0]?.block ?? "";
        assert.equal(
          (spinCompanion.match(/additive="base"/gu) ?? []).length,
          2,
        );
        assert.doesNotMatch(spinCompanion, /-#ppt_x|-#ppt_y/u);
        assert.match(spinCompanion, /fmla="[^"]*\$\*\(/u);
      }
    }
  });

  it("rejects sum, delta, or identity coordinates on mixed SpinIn companions", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "mixed-spin-negative" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "mixed-spin-negative-shape",
                x: 100,
                y: 100,
                width: 400,
                height: 160,
                fill: "#0070C0",
                textStyle: { fontSize: 20, color: "#FFFFFF" },
                enter: [
                  spinIn({ duration: 1000 }),
                  fade({ by: "paragraph", duration: 700, delay: 300 }),
                ],
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const companion = animationEffectBlocks(slideXml).find(
      ({ block, presetId }) =>
        presetId === 49 && block.includes('nodeType="withEffect"'),
    )?.block;
    assert.ok(companion !== undefined);
    const sumCoordinates = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companion,
        companion.replace(
          '<p:cBhvr additive="base">',
          '<p:cBhvr additive="sum">',
        ),
      ),
    );
    assert.ok(
      (await validatePptx(sumCoordinates)).errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );
    const deltaCoordinates = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companion,
        companion.replace(
          /<p:tav tm="0" fmla="([^"]*)">/u,
          '<p:tav tm="0" fmla="$1-#ppt_x">',
        ),
      ),
    );
    assert.ok(
      (await validatePptx(deltaCoordinates)).errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );
    const identityCoordinates = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companion,
        companion.replace(
          /<p:tav tm="0" fmla="[^"]*">/u,
          '<p:tav tm="0" fmla="#ppt_x">',
        ),
      ),
    );
    assert.ok(
      (await validatePptx(identityCoordinates)).errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );
  });

  it("matches companions to each non-overlapping whole effect instance", async () => {
    const cases = [
      {
        id: "sequential-fly-instances",
        enter: [
          fly({ duration: 1000 }),
          fly({ duration: 2000, delay: 1200 }),
          fade({ by: "paragraph", duration: 700, delay: 300 }),
        ],
      },
      {
        id: "same-duration-fly-instances",
        enter: [
          fly({ duration: 1000 }),
          fly({ duration: 1000, delay: 1200 }),
          fade({ by: "paragraph", duration: 700, delay: 300 }),
        ],
      },
      {
        id: "same-parameter-fly-instances",
        enter: [
          fly({ duration: 1000 }),
          fly({ duration: 1000 }),
          fade({ by: "paragraph", duration: 700, delay: 300 }),
        ],
      },
      {
        id: "two-spin-in-instances",
        enter: [
          spinIn({ duration: 1000 }),
          spinIn({ duration: 1000, delay: 1200 }),
          fade({ by: "paragraph", duration: 700, delay: 300 }),
        ],
      },
      {
        id: "fly-and-spin-in-instances",
        enter: [
          fly({ duration: 1000 }),
          spinIn({ duration: 1000, delay: 1200 }),
          fade({ by: "paragraph", duration: 700, delay: 300 }),
        ],
      },
    ] as const;
    for (const item of cases) {
      const result = await build(
        createElement(
          Deck,
          { width: 960, height: 540 },
          createElement(
            Slide,
            { id: item.id },
            createElement(
              Step,
              { n: 1 },
              createElement(
                Shape,
                {
                  id: `${item.id}-shape`,
                  x: 100,
                  y: 100,
                  width: 400,
                  height: 160,
                  fill: "#0070C0",
                  textStyle: { fontSize: 20, color: "#FFFFFF" },
                  enter: item.enter,
                },
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ),
            ),
          ),
        ),
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], item.id);
      const motion = result.report.slides[0]?.capabilities?.[
        "motion"
      ] as MotionObservation;
      assert.equal(motion.steps[0]?.effects.length, 4, item.id);
    }

    const documentResult = await buildDocument(
      documentDeck({ width: 960, height: 540 }).append(
        documentSlide({ id: "document-same-fly-instances" }).append(
          documentStep({ n: 1 }).append(
            documentShape({
              id: "document-same-fly-instances-shape",
              x: 100,
              y: 100,
              width: 400,
              height: 160,
              fill: "#0070C0",
              textStyle: { fontSize: 20, color: "#FFFFFF" },
              enter: [
                documentFly({ duration: 1000 }),
                documentFly({ duration: 1000 }),
                documentFade({ by: "paragraph", duration: 700, delay: 300 }),
              ],
            }).content(
              documentParagraph("First paragraph"),
              documentParagraph("Second paragraph"),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual((await validatePptx(documentResult.pptx)).errors, []);
  });

  it("rejects missing or extra companions in same-parameter whole effect instances", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "same-parameter-companion-negatives" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "same-parameter-companion-shape",
                x: 100,
                y: 100,
                width: 400,
                height: 160,
                fill: "#0070C0",
                textStyle: { fontSize: 20, color: "#FFFFFF" },
                enter: [
                  fly({ duration: 1000 }),
                  fly({ duration: 1000 }),
                  fade({ by: "paragraph", duration: 700, delay: 300 }),
                ],
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const companions = animationEffectBlocks(slideXml)
      .filter(
        ({ block, presetId }) =>
          presetId === 2 &&
          block.includes('nodeType="withEffect"') &&
          block.includes("<p:pRg"),
      )
      .map(({ block }) => block);
    assert.equal(companions.length, 4);
    const missingOne = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(companions[0] ?? "", ""),
    );
    assert.ok(
      (await validatePptx(missingOne)).errors.some(
        ({ code }) =>
          code === "inconsistent-animation-target-range" ||
          code === "invalid-animation-projection-target",
      ),
    );
    const extraCompanion = (companions[0] ?? "").replace(
      /\bid="(\d+)"/gu,
      (_match, id: string) => `id="${Number(id) + 10000}"`,
    );
    const extraOne = rewriteSlideXml(result.pptx, (xml) =>
      xml.replace(
        companions[0] ?? "",
        `${companions[0] ?? ""}${extraCompanion}`,
      ),
    );
    assert.ok(
      (await validatePptx(extraOne)).errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
  });

  it("keeps cued delayed mixed projection companions synchronized with the whole completion", async () => {
    function CuedMixedProjectionDeck() {
      const source = useMotionScope();
      const mixed = useMotionScope();
      const follower = useMotionScope();
      return createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "mixed-scheduling" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              MotionScope,
              { scope: source, publishEnd: source.cue("ready") },
              createElement(Shape, {
                id: "schedule-source",
                x: 20,
                y: 20,
                width: 80,
                height: 40,
                fill: "#336699",
                enter: fade({ duration: 100 }),
              }),
            ),
            createElement(
              MotionScope,
              {
                scope: mixed,
                startAt: after(source.cue("ready"), 120),
                publishEnd: mixed.cue("done"),
              },
              createElement(
                Shape,
                {
                  id: "schedule-mixed",
                  x: 100,
                  y: 100,
                  width: 400,
                  height: 160,
                  fill: "#0070C0",
                  textStyle: { fontSize: 20, color: "#FFFFFF" },
                  enter: [
                    fly({ duration: 4000, delay: 100 }),
                    fade({ by: "paragraph", duration: 700, delay: 1500 }),
                  ],
                  emphasis: repeat(spin({ duration: 500 }), 2),
                },
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ),
            ),
            createElement(
              MotionScope,
              { scope: follower, startAt: after(mixed.cue("done"), 20) },
              createElement(Shape, {
                id: "schedule-follower",
                x: 20,
                y: 80,
                width: 80,
                height: 40,
                fill: "#993366",
                enter: fade({ duration: 100 }),
              }),
            ),
          ),
        ),
      );
    }

    const result = await build(createElement(CuedMixedProjectionDeck));
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects =
      file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
    const flyEffects = effects.filter(({ preset }) => preset === "fly");
    assert.equal(flyEffects.length, 3);
    assert.deepEqual(
      [...new Set(flyEffects.map(({ delayMs }) => delayMs))],
      [320],
    );
    assert.ok(flyEffects.every(({ durationMs }) => durationMs === 4000));
    const spinEffect = effects.find(({ preset }) => preset === "spin");
    assert.deepEqual([spinEffect?.repeatCount, spinEffect?.delayMs], [2, 220]);
    const followerObject = result.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "schedule-follower",
    );
    const followerEffect = effects.find(
      ({ targetId }) => targetId === followerObject?.id,
    );
    const wholeFly = flyEffects.find(
      ({ targetRange }) => targetRange === undefined,
    );
    assert.ok(wholeFly !== undefined);
    assert.equal(
      followerEffect?.delayMs,
      wholeFly.delayMs + wholeFly.durationMs + 20,
    );
  });

  it("keeps explicit paragraph granularity on a shape with its own text and background", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "shape-paragraph-entrance" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "paragraph-shape-entry",
                x: 120,
                y: 140,
                width: 360,
                height: 140,
                fill: "#0070C0",
                textStyle: { fontSize: 20, color: "#FFFFFF" },
                enter: fade({ by: "paragraph", duration: 300 }),
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    const shape = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "paragraph-shape-entry",
    );
    assert.ok(shape !== undefined);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target }) => target.range),
      [
        { type: "paragraph", start: 0, end: 0 },
        { type: "paragraph", start: 1, end: 1 },
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ targetRange, nativeBuild }) => [targetRange, nativeBuild],
      ),
      [
        [{ type: "paragraph", start: 0, end: 0 }, "paragraph"],
        [{ type: "paragraph", start: 1, end: 1 }, "paragraph"],
      ],
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      slideXml,
      new RegExp(
        `<p:spTgt spid="${shape.id}"><p:txEl><p:pRg st="0" end="0">`,
        "u",
      ),
    );
    assert.match(
      slideXml,
      new RegExp(`<p:bldP spid="${shape.id}" grpId="0" build="p"`, "u"),
    );
    assert.doesNotMatch(
      slideXml,
      new RegExp(`<p:spTgt spid="${shape.id}"></p:spTgt>`, "u"),
    );
  });

  it("keeps multi-paragraph shape text and effect arrays on whole-object targets", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "whole-shape-array" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "whole-shape-array-entry",
                x: 120,
                y: 140,
                width: 360,
                height: 140,
                fill: "#F97316",
                textStyle: { fontSize: 20, color: "#111827" },
                enter: [
                  fade({ duration: 200 }),
                  fly({ from: "left", duration: 300 }),
                ],
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    const shape = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "whole-shape-array-entry",
    );
    assert.ok(shape !== undefined);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, preset }) => [
        target.range,
        preset,
      ]),
      [
        [undefined, "fade"],
        [undefined, "fly"],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ targetRange, preset }) => [targetRange, preset],
      ),
      [
        [undefined, "fade"],
        [undefined, "fly"],
      ],
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.doesNotMatch(
      slideXml,
      new RegExp(`<p:spTgt spid="${shape.id}"><p:txEl>`, "u"),
    );
    assert.doesNotMatch(
      slideXml,
      new RegExp(`<p:spTgt spid="${shape.id}"><p:bg>`, "u"),
    );
    assert.equal(
      slideXml.match(
        new RegExp(
          `<p:bldP spid="${shape.id}" grpId="0" animBg="1"></p:bldP>`,
          "gu",
        ),
      )?.length,
      1,
    );
  });

  it("checks every build target of a whole-object shape before accepting animBg", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "whole-shape-background-role" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Shape,
              {
                id: "whole-shape-background-role-entry",
                x: 120,
                y: 140,
                width: 360,
                height: 140,
                fill: "#0070C0",
                textStyle: { fontSize: 20, color: "#FFFFFF" },
                enter: fade({ duration: 300 }),
              },
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ),
          ),
        ),
      ),
    );
    const shape = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "whole-shape-background-role-entry",
    );
    assert.ok(shape !== undefined);
    const initialFile = await validatePptx(result.pptx);
    assert.deepEqual(initialFile.errors, []);
    assert.deepEqual(
      [
        initialFile.slides[0]?.animations?.[0]?.effects[0]?.nativeBuild,
        initialFile.slides[0]?.animations?.[0]?.effects[0]?.nativeBuildAnimBg,
      ],
      [undefined, true],
    );

    const path = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const xml = strFromU8(parts[path]!);
    const shapeTarget = `<p:spTgt spid="${shape.id}"></p:spTgt>`;
    const backgroundTarget = `<p:spTgt spid="${shape.id}"><p:bg></p:bg></p:spTgt>`;
    const behavior = (tag: "set" | "animEffect"): string => {
      const found = xml.match(
        new RegExp(`<p:${tag}\\b[\\s\\S]*?</p:${tag}>`, "u"),
      )?.[0];
      assert.ok(found !== undefined);
      return found;
    };
    const setTargetToBackground = (tag: "set" | "animEffect"): string =>
      xml.replace(
        behavior(tag),
        behavior(tag).replace(shapeTarget, backgroundTarget),
      );
    const buildWithAnimBg = (source: string, animBg: "0" | "1"): string =>
      source.replace(
        `<p:bldP spid="${shape.id}" grpId="0" animBg="1"></p:bldP>`,
        `<p:bldP spid="${shape.id}" grpId="0" animBg="${animBg}"></p:bldP>`,
      );
    const assertBackgroundRoleRejected = async (
      slideXml: string,
    ): Promise<void> => {
      assert.notEqual(slideXml, xml);
      const file = await validatePptx(
        zipSync({ ...parts, [path]: strToU8(slideXml) }),
      );
      assert.ok(
        file.errors.some(
          ({ code }) => code === "inconsistent-animation-background-role",
        ),
      );
    };

    // Either behavior may carry the real background branch; the other whole target must not
    // exempt the group from the animBg check.
    await assertBackgroundRoleRejected(
      buildWithAnimBg(setTargetToBackground("set"), "0"),
    );
    await assertBackgroundRoleRejected(
      buildWithAnimBg(setTargetToBackground("animEffect"), "0"),
    );

    // Native whole-object targets without a background branch may still declare animBg=false.
    const wholeWithoutAnimBg = buildWithAnimBg(xml, "0");
    assert.notEqual(wholeWithoutAnimBg, xml);
    const wholeFile = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(wholeWithoutAnimBg) }),
    );
    assert.deepEqual(wholeFile.errors, []);
    assert.equal(
      wholeFile.slides[0]?.animations?.[0]?.effects[0]?.nativeBuildAnimBg,
      false,
    );
  });

  it("keeps split background and ordinary build readers valid beside whole-object roles", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "split-background" },
          createElement(
            Step,
            { n: 1 },
            createElement(Text, {
              id: "split-background-text",
              x: 100,
              y: 100,
              width: 400,
              height: 160,
              style: {
                backgroundColor: "#FEF3C7",
                borderColor: "#1D4ED8",
                borderWidth: 2,
              },
              enter: fade({ duration: 300 }),
              children: [
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ],
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "paragraph-modes" },
          createElement(
            Step,
            { n: 1 },
            createElement(Text, {
              id: "word-text",
              x: 100,
              y: 100,
              width: 400,
              height: 120,
              enter: fade({ by: "word" }),
              children: [createElement(P, { key: "one" }, "Word one two")],
            }),
            createElement(Text, {
              id: "letter-text",
              x: 100,
              y: 300,
              width: 400,
              height: 120,
              enter: fade({ by: "letter" }),
              children: [createElement(P, { key: "one" }, "Letter")],
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "group-and-array" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Group,
              {
                id: "group-entry",
                emphasis: spin({ duration: 1000 }),
              },
              createElement(Shape, {
                id: "group-shape",
                x: 100,
                y: 100,
                width: 80,
                height: 80,
                fill: "#F97316",
              }),
              createElement(Text, {
                id: "group-text",
                x: 200,
                y: 100,
                width: 200,
                height: 80,
                enter: fade(),
                children: [createElement(P, { key: "one" }, "Group text")],
              }),
            ),
            createElement(
              Shape,
              {
                id: "array-shape",
                x: 100,
                y: 300,
                width: 300,
                height: 120,
                fill: "#0070C0",
                textStyle: { fontSize: 20 },
                enter: [fade(), fly({ from: "left" })],
              },
              createElement(P, { key: "one" }, "Array text"),
            ),
          ),
        ),
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const splitText = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "split-background-text",
    );
    assert.ok(splitText !== undefined);
    const splitEffect = file.slides[0]?.animations?.[0]?.effects[0];
    assert.deepEqual(
      [
        splitEffect?.targetId,
        splitEffect?.nativeBuild,
        splitEffect?.nativeBuildGroupId,
        splitEffect?.nativeBuildAnimBg,
      ],
      [splitText.id, "paragraph", 1, true],
    );

    function CuedDeck() {
      const source = useMotionScope();
      const follower = useMotionScope();
      return createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "cue" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              MotionScope,
              { scope: source, publishEnd: source.cue("ready") },
              createElement(Shape, {
                id: "cue-source",
                x: 20,
                y: 20,
                width: 80,
                height: 40,
                fill: "#336699",
                enter: fade(),
              }),
            ),
            createElement(
              MotionScope,
              { scope: follower, startAt: after(source.cue("ready"), 20) },
              createElement(Shape, {
                id: "cue-follower",
                x: 20,
                y: 80,
                width: 80,
                height: 40,
                fill: "#993366",
                enter: fade(),
              }),
            ),
          ),
        ),
      );
    }

    const cueResult = await build(createElement(CuedDeck));
    assert.deepEqual((await validatePptx(cueResult.pptx)).errors, []);

    const path = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const xml = strFromU8(parts[path]!);
    const wrongAnimBg = xml.replace(
      `<p:bldP spid="${splitText.id}" grpId="1" build="p" animBg="1"></p:bldP>`,
      `<p:bldP spid="${splitText.id}" grpId="1" build="p" animBg="0"></p:bldP>`,
    );
    assert.notEqual(wrongAnimBg, xml);
    const wrongFile = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(wrongAnimBg) }),
    );
    assert.ok(
      wrongFile.errors.some(
        ({ code }) => code === "inconsistent-animation-background-role",
      ),
    );
  });

  it("initializes whole Text entrance by paragraphs and background while preserving by-P deferred build", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "future-whole-text" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Group,
              {
                id: "first-click-group",
                emphasis: repeat(spin({ duration: 1600 }), 2),
              },
              createElement(Shape, {
                id: "first-click-member",
                x: 320,
                y: 230,
                width: 82,
                height: 82,
                fill: "#F97316",
              }),
            ),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(Text, {
              id: "future-text",
              x: 180,
              y: 405,
              width: 600,
              height: 55,
              style: {
                backgroundColor: "#FEF3C7",
                borderColor: "#1D4ED8",
                borderWidth: 2,
              },
              enter: fade({ duration: 500 }),
              children: [
                createElement(P, { key: "first" }, "Visible on click two"),
                createElement(
                  P,
                  { key: "second" },
                  "The second paragraph follows",
                ),
              ],
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "paragraph-delay-control" },
          createElement(Text, {
            id: "paragraph-delay-text",
            x: 130,
            y: 150,
            width: 700,
            height: 290,
            style: {
              backgroundColor: "#DBEAFE",
              borderColor: "#2563EB",
              borderWidth: 2,
            },
            enter: [
              fly({ duration: 4000, from: "left" }),
              fade({ by: "paragraph", duration: 700, delay: 1500 }),
            ],
            children: [
              createElement(P, { key: "long" }, "First paragraph"),
              createElement(P, { key: "short" }, "Second paragraph"),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const futureTextId = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "future-text",
    )?.id;
    assert.ok(futureTextId !== undefined);
    const futureXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const futureFade = file.slides[0]?.animations?.[1]?.effects.find(
      ({ targetId, preset }) => targetId === futureTextId && preset === "fade",
    );
    assert.deepEqual(
      futureFade?.nativeBehaviors
        ?.filter(({ behavior }) => behavior === "set")
        .map(
          ({ targetType, targetRange, to }) =>
            `${targetType}:${targetRange?.type === "paragraph" ? `${targetRange.start}-${targetRange.end}` : ""}:${to}`,
        )
        .sort(),
      ["background::visible", "paragraph:0-0:visible", "paragraph:1-1:visible"],
    );
    const mainSequenceStart = futureXml.indexOf('nodeType="mainSeq"');
    const firstClickCondition = futureXml.indexOf(
      'delay="indefinite"',
      mainSequenceStart,
    );
    assert.ok(
      mainSequenceStart >= 0 &&
        firstClickCondition >= 0 &&
        !futureXml
          .slice(mainSequenceStart, firstClickCondition)
          .includes('val="hidden"'),
    );
    assert.deepEqual(
      [
        futureFade?.nativeBuild,
        futureFade?.nativeBuildGroupId,
        futureFade?.nativeBuildAnimBg,
      ],
      ["paragraph", 1, true],
    );

    const slidePath = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const visibilitySets = [...futureXml.matchAll(/<p:set>[\s\S]*?<\/p:set>/gu)]
      .map(([nativeSet]) => nativeSet ?? "")
      .filter((nativeSet) =>
        nativeSet.includes(`<p:spTgt spid="${futureTextId}"`),
      );
    const firstParagraphSetter = visibilitySets.find((nativeSet) =>
      nativeSet.includes('<p:pRg st="0" end="0">'),
    );
    assert.ok(firstParagraphSetter !== undefined);
    const missingParagraphRole = await validatePptx(
      zipSync({
        ...parts,
        [slidePath]: strToU8(futureXml.replace(firstParagraphSetter, "")),
      }),
    );
    assert.ok(
      missingParagraphRole.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const backgroundSetter = visibilitySets.find((nativeSet) =>
      nativeSet.includes("<p:bg></p:bg>"),
    );
    assert.ok(backgroundSetter !== undefined);
    const missingBackgroundRole = await validatePptx(
      zipSync({
        ...parts,
        [slidePath]: strToU8(futureXml.replace(backgroundSetter, "")),
      }),
    );
    assert.ok(
      missingBackgroundRole.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const duplicatedWholeSetter = firstParagraphSetter
      .replace(/id="\d+"/u, 'id="99999"')
      .replace(
        new RegExp(
          `<p:spTgt spid="${futureTextId}"><p:txEl><p:pRg st="0" end="0"><\\/p:pRg><\\/p:txEl><\\/p:spTgt>`,
          "u",
        ),
        `<p:spTgt spid="${futureTextId}"></p:spTgt>`,
      );
    const wholeRoleMixedWithParagraphRoles = await validatePptx(
      zipSync({
        ...parts,
        [slidePath]: strToU8(
          futureXml.replace(
            firstParagraphSetter,
            `${firstParagraphSetter}${duplicatedWholeSetter}`,
          ),
        ),
      }),
    );
    assert.ok(
      wholeRoleMixedWithParagraphRoles.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const paragraphBehaviors = [
      ...futureXml.matchAll(/<p:(set|animEffect)\b[^>]*>[\s\S]*?<\/p:\1>/gu),
    ]
      .map(([nativeBehavior]) => nativeBehavior ?? "")
      .filter((nativeBehavior) =>
        nativeBehavior.includes(`<p:spTgt spid="${futureTextId}"><p:txEl>`),
      );
    assert.equal(paragraphBehaviors.length, 4);
    const backgroundOnlyXml = paragraphBehaviors.reduce(
      (xml, nativeBehavior) => xml.replace(nativeBehavior, ""),
      futureXml,
    );
    const missingAllParagraphRoles = await validatePptx(
      zipSync({
        ...parts,
        [slidePath]: strToU8(backgroundOnlyXml),
      }),
    );
    assert.ok(
      missingAllParagraphRoles.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const backgroundOnlyBuildXml = backgroundOnlyXml.replace(
      `<p:bldP spid="${futureTextId}" grpId="1" build="p" animBg="1"></p:bldP>`,
      `<p:bldP spid="${futureTextId}" grpId="1" build="allAtOnce" animBg="1"></p:bldP>`,
    );
    assert.notEqual(backgroundOnlyBuildXml, backgroundOnlyXml);
    const legalBackgroundOnlyFile = await validatePptx(
      zipSync({
        ...parts,
        [slidePath]: strToU8(backgroundOnlyBuildXml),
      }),
    );
    assert.deepEqual(legalBackgroundOnlyFile.errors, []);
    assert.deepEqual(
      legalBackgroundOnlyFile.slides[0]?.animations?.[1]?.effects[0]?.nativeBehaviors?.map(
        ({ behavior, targetType }) => [behavior, targetType],
      ),
      [
        ["set", "background"],
        ["animEffect", "background"],
      ],
    );

    const delayedTextId = result.report.slides[1]?.objects[0]?.id;
    assert.ok(delayedTextId !== undefined);
    const delayedXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide2.xml"]!,
    );
    assert.match(
      delayedXml,
      new RegExp(`<p:bldP spid="${delayedTextId}" grpId="0" build="p"`, "u"),
    );
    assert.doesNotMatch(
      delayedXml,
      new RegExp(
        `<p:set>[\\s\\S]*?<p:spTgt spid="${delayedTextId}"></p:spTgt>[\\s\\S]*?<p:strVal val="hidden"`,
        "u",
      ),
    );
    assert.equal(
      file.slides[1]?.animations?.[0]?.effects.find(
        ({ targetId, preset, targetRange }) =>
          targetId === delayedTextId &&
          preset === "fade" &&
          targetRange?.type === "paragraph" &&
          targetRange.start === 1,
      )?.delayMs,
      2200,
    );
  });

  it("preserves paragraph and background roles when Group Spin is projected to Text members", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "group-text-spin-projection" },
          createElement(
            Group,
            {
              id: "text-member-group",
              emphasis: repeat(spin({ duration: 4000, from: 360 }), 1),
            },
            createElement(Text, {
              id: "animated-text-member",
              x: 130,
              y: 150,
              width: 400,
              height: 180,
              style: {
                backgroundColor: "#FEF3C7",
                borderColor: "#1D4ED8",
                borderWidth: 3,
              },
              enter: fade({ duration: 4000 }),
              children: [
                createElement(
                  P,
                  { key: "long" },
                  "A longer paragraph remains inside the same native text object.",
                ),
                createElement(P, { key: "short" }, "A short second paragraph"),
              ],
            }),
            createElement(Shape, {
              id: "static-orange-member",
              x: 600,
              y: 220,
              width: 55,
              height: 55,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );
    const slide = result.report.slides[0];
    const group = slide?.objects.find(
      ({ authorId: testId }) => testId === "text-member-group",
    );
    const text = group?.members?.find(
      ({ authorId: testId }) => testId === "animated-text-member",
    );
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const projection = motion.nativeProjections?.find(
      ({ logicalTargetId, preset }) =>
        logicalTargetId === group?.id && preset === "spin",
    );
    assert.deepEqual(
      projection?.nativeEffects.map(({ targetId }) => targetId),
      group?.members?.map(({ id }) => id),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileGroup = file.slides[0]?.objects.find(
      ({ name }) => name === group?.name,
    );
    assert.deepEqual(
      fileGroup?.members?.map(({ id }) => id),
      group?.members?.map(({ id }) => id),
    );
    const effects = file.slides[0]?.animations?.flatMap((step) => step.effects);
    const projectedTextSpin = effects?.find(
      ({ preset, targetId }) => preset === "spin" && targetId === text?.id,
    );
    assert.deepEqual(
      projectedTextSpin?.nativeBehaviors?.map(
        ({ behavior, targetType, targetRange, attributeNames }) => [
          behavior,
          targetType,
          targetRange,
          attributeNames,
        ],
      ),
      [
        [
          "animRot",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["r"],
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["ppt_x"],
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 0, end: 0 },
          ["ppt_y"],
        ],
        [
          "animRot",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["r"],
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["ppt_x"],
        ],
        [
          "anim",
          "paragraph",
          { type: "paragraph", start: 1, end: 1 },
          ["ppt_y"],
        ],
        ["animRot", "background", undefined, ["r"]],
        ["anim", "background", undefined, ["ppt_x"]],
        ["anim", "background", undefined, ["ppt_y"]],
      ],
    );
    assert.deepEqual(
      [
        projectedTextSpin?.nativeBuild,
        projectedTextSpin?.nativeBuildGroupId,
        projectedTextSpin?.nativeBuildAnimBg,
      ],
      ["paragraph", 1, true],
    );

    const textId = text?.id;
    assert.ok(textId !== undefined);
    const path = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const xml = strFromU8(parts[path]!);
    const firstParagraphX = [
      ...xml.matchAll(/<p:anim\b[\s\S]*?<\/p:anim>/gu),
    ].find(
      ([nativeBehavior]) =>
        nativeBehavior?.includes(
          `<p:spTgt spid="${textId}"><p:txEl><p:pRg st="0" end="0"></p:pRg></p:txEl></p:spTgt>`,
        ) && nativeBehavior.includes("<p:attrName>ppt_x</p:attrName>"),
    )?.[0];
    assert.ok(firstParagraphX !== undefined);
    const invalidFile = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(xml.replace(firstParagraphX, "")) }),
    );
    assert.ok(
      invalidFile.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );
  });

  it("shares text build responsibilities between Group Text paragraph animations and projection", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "group-text-paragraph-build" },
          createElement(
            Group,
            {
              id: "paragraph-build-group",
              emphasis: spin({ duration: 4000, from: 360 }),
            },
            createElement(Text, {
              id: "paragraph-build-text",
              x: 100,
              y: 100,
              width: 330,
              height: 160,
              style: {
                backgroundColor: "#FEF3C7",
                borderColor: "#1D4ED8",
                borderWidth: 2,
              },
              enter: fade({ by: "paragraph", duration: 700, delay: 1500 }),
              children: [
                createElement(P, { key: "long" }, "A longer paragraph."),
                createElement(P, { key: "short" }, "A short second paragraph."),
              ],
            }),
            createElement(Shape, {
              id: "paragraph-build-shape",
              x: 550,
              y: 240,
              width: 70,
              height: 70,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );

    const group = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "paragraph-build-group",
    );
    const text = group?.members?.find(
      ({ authorId: testId }) => testId === "paragraph-build-text",
    );
    assert.ok(text !== undefined);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects = file.slides[0]?.animations?.flatMap((step) => step.effects);
    const projectedSpin = effects?.find(
      ({ preset, targetId }) => preset === "spin" && targetId === text.id,
    );
    assert.deepEqual(
      [
        projectedSpin?.nativeBuild,
        projectedSpin?.nativeBuildGroupId,
        projectedSpin?.nativeBuildAnimBg,
      ],
      ["allAtOnce", 1, true],
    );

    const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(
      xml,
      new RegExp(
        `<p:bldP spid="${text.id}" grpId="0" build="p" animBg="0">`,
        "u",
      ),
    );
    assert.match(
      xml,
      new RegExp(
        `<p:bldP spid="${text.id}" grpId="1" build="allAtOnce" animBg="1">`,
        "u",
      ),
    );
  });

  it("requires no background role for build group of Group Text word-by-word projection", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "group-text-word-build" },
          createElement(
            Group,
            { id: "word-build-group", emphasis: spin({ duration: 4000 }) },
            createElement(Text, {
              id: "word-build-text",
              x: 100,
              y: 100,
              width: 330,
              height: 160,
              enter: fade({ by: "word", duration: 700, delay: 1500 }),
              children: [
                createElement(P, { key: "first" }, "A first paragraph."),
                createElement(P, { key: "second" }, "A second paragraph."),
              ],
            }),
            createElement(Shape, {
              id: "word-build-shape",
              x: 550,
              y: 240,
              width: 70,
              height: 70,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );
    const group = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "word-build-group",
    );
    const text = group?.members?.find(
      ({ authorId: testId }) => testId === "word-build-text",
    );
    assert.ok(text !== undefined);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects = file.slides[0]?.animations?.flatMap((step) => step.effects);
    const wordFade = effects?.find(
      ({ preset, targetId }) => preset === "fade" && targetId === text.id,
    );
    const projectedSpin = effects?.find(
      ({ preset, targetId }) => preset === "spin" && targetId === text.id,
    );
    assert.equal(wordFade?.nativeBuildGroupId, 0);
    assert.deepEqual(
      [projectedSpin?.nativeBuild, projectedSpin?.nativeBuildGroupId],
      ["allAtOnce", 1],
    );

    const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(
      xml,
      new RegExp(
        `<p:bldP spid="${text.id}" grpId="1" build="allAtOnce"></p:bldP>`,
        "u",
      ),
    );
    assert.match(xml, /<p:iterate type="wd">/u);
  });

  it("shares paragraph build groups across whole effect arrays on the same Group Text target", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "group-text-whole-array-build" },
          createElement(
            Group,
            { id: "whole-array-group", emphasis: spin({ duration: 4000 }) },
            createElement(Text, {
              id: "whole-array-text",
              x: 100,
              y: 100,
              width: 330,
              height: 160,
              enter: fade({ duration: 700 }),
              emphasis: spin({ duration: 900 }),
              children: [
                createElement(P, { key: "first" }, "A first paragraph."),
                createElement(P, { key: "second" }, "A second paragraph."),
              ],
            }),
            createElement(Shape, {
              id: "whole-array-shape",
              x: 550,
              y: 240,
              width: 70,
              height: 70,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );
    const group = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "whole-array-group",
    );
    const text = group?.members?.find(
      ({ authorId: testId }) => testId === "whole-array-text",
    );
    assert.ok(text !== undefined);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
    assert.match(
      xml,
      new RegExp(`<p:bldP spid="${text.id}" grpId="1" build="p">`, "u"),
    );
    assert.doesNotMatch(
      xml,
      new RegExp(`<p:bldP spid="${text.id}" grpId="0"`, "u"),
    );
  });

  it("still triggers parent Spin projection in automatic steps for Group member click motions", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "group-cross-step-projection" },
          createElement(
            Group,
            {
              id: "cross-step-group",
              emphasis: repeat(spin({ duration: 8000, from: 360 }), 2),
            },
            createElement(Shape, {
              id: "cross-step-orange",
              x: 465,
              y: 250,
              width: 58,
              height: 58,
              fill: "#F97316",
            }),
            createElement(
              Step,
              { n: 1 },
              createElement(Shape, {
                id: "cross-step-green",
                x: 320,
                y: 230,
                width: 82,
                height: 82,
                fill: "#16A34A",
                emphasis: repeat(
                  breathe({
                    period: 1600,
                    from: "#16A34A",
                    to: "#86EFAC",
                  }),
                  "forever",
                ),
              }),
            ),
          ),
        ),
      ),
    );
    const slide = result.report.slides[0];
    const group = slide?.objects.find(
      ({ authorId: testId }) => testId === "cross-step-group",
    );
    const green = group?.members?.find(
      ({ authorId: testId }) => testId === "cross-step-green",
    );
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const steps = motion.steps;
    assert.equal(steps[0]?.trigger, "auto");
    assert.equal(steps[1]?.trigger, "click");
    const projection = motion.nativeProjections?.find(
      ({ logicalTargetId, preset }) =>
        logicalTargetId === group?.id && preset === "spin",
    );
    const memberIds = group?.members?.map(({ id }) => id) ?? [];
    assert.deepEqual(
      projection?.nativeEffects.map(({ targetId }) => targetId),
      memberIds,
    );
    assert.ok(
      steps[1]?.effects.some(
        ({ preset, target }) => preset === "breathe" && target.id === green?.id,
      ),
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const animations = file.slides[0]?.animations ?? [];
    assert.deepEqual(
      animations.map(({ trigger, effects }) => [
        trigger,
        effects.map(({ preset, targetId }) => [preset, targetId]),
      ]),
      [
        [
          "auto",
          [
            ["spin", memberIds[0]],
            ["spin", memberIds[1]],
          ],
        ],
        ["click", [["breathe", green?.id]]],
      ],
    );
  });

  it("preserves native Group and projects finite whole Spin to all real members", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "group-spin-projection" },
          createElement(
            Group,
            {
              id: "spinning-group",
              name: "spinning-group",
              emphasis: repeat(spin({ duration: 1600, from: 360 }), 2),
            },
            createElement(Shape, {
              id: "breathing-member",
              x: 40,
              y: 40,
              width: 40,
              height: 40,
              fill: "#164E3D",
              emphasis: repeat(
                breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
                "forever",
              ),
            }),
            createElement(Shape, {
              id: "static-member",
              x: 100,
              y: 50,
              width: 30,
              height: 30,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );
    const slide = result.report.slides[0];
    const group = slide?.objects.find(
      ({ authorId: testId }) => testId === "spinning-group",
    );
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const logicalSpin = motion.steps[0]?.effects.find(
      ({ preset, target }) => preset === "spin" && target.id === group?.id,
    );
    assert.deepEqual(
      [logicalSpin?.durationMs, logicalSpin?.repeatCount, logicalSpin?.endMs],
      [1600, 2, 3200],
    );

    const projection = motion.nativeProjections?.find(
      ({ logicalTargetId, preset }) =>
        logicalTargetId === group?.id && preset === "spin",
    );
    const members = group?.members ?? [];
    assert.deepEqual(
      projection?.nativeEffects.map(({ targetId }) => targetId),
      members.map(({ id }) => id),
    );
    assert.equal(
      new Set(projection?.nativeEffects.map(({ timingId }) => timingId)).size,
      2,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileObjects = file.slides[0]?.objects ?? [];
    const fileGroup = fileObjects.find(({ name }) => name === "spinning-group");
    const fileMembers = fileGroup?.members ?? [];
    assert.deepEqual(
      fileMembers.map(({ id }) => id),
      members.map(({ id }) => id),
    );
    const fileEffects =
      file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
    const projectedSpins = fileEffects.filter(
      ({ preset }) => preset === "spin",
    );
    const memberBreathe = fileEffects.find(
      ({ preset, targetId }) =>
        preset === "breathe" &&
        targetId ===
          members.find(({ authorId: testId }) => testId === "breathing-member")
            ?.id,
    );
    assert.deepEqual(
      projectedSpins.map(({ targetId, durationMs, repeatCount }) => [
        targetId,
        durationMs,
        repeatCount,
      ]),
      members.map(({ id }) => [id, 1600, 2]),
    );
    assert.deepEqual(
      projectedSpins.map(({ timingId }) => timingId),
      projection?.nativeEffects.map(({ timingId }) => timingId),
    );
    assert.deepEqual(
      projectedSpins[0]?.nativeBehaviors?.map(
        ({ behavior, attributeNames, additive, durationMs }) => [
          behavior,
          attributeNames,
          additive,
          durationMs,
        ],
      ),
      [
        ["animRot", ["r"], "sum", 1600],
        ["anim", ["ppt_x"], "sum", 1600],
        ["anim", ["ppt_y"], "sum", 1600],
      ],
    );
    assert.deepEqual(
      [memberBreathe?.repeatCount, memberBreathe?.completion],
      ["forever", "no-natural-completion"],
    );
    assert.equal(
      projectedSpins.some(({ targetId }) => targetId === group?.id),
      false,
    );
  });

  it("preserves existing Fly and SpinIn entrance arrays and emits common center offset for paragraphs", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { id: "text-motion-array" },
          createElement(Text, {
            id: "array-text",
            x: 800,
            y: 500,
            width: 400,
            height: 160,
            enter: [
              fly({ from: "left", duration: 4000 }),
              spinIn({ duration: 4000 }),
              fade({ by: "paragraph", duration: 700, delay: 1500 }),
            ],
            children: [
              createElement(P, { key: "wide" }, "A long first paragraph"),
              createElement(P, { key: "short" }, "SHORT"),
            ],
          }),
        ),
      ),
    );
    const textObject = result.report.slides[0]?.objects[0];
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ preset }) => preset),
      ["fly", "spinIn", "fade", "fade"],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileEffects =
      file.slides[0]?.animations?.flatMap((step) => step.effects) ?? [];
    const flyEffect = fileEffects.find(({ preset }) => preset === "fly");
    const spinEffect = fileEffects.find(({ preset }) => preset === "spinIn");
    const delayedFade = fileEffects.find(
      ({ preset, delayMs }) => preset === "fade" && delayMs === 2200,
    );
    assert.deepEqual(
      [flyEffect?.targetId, spinEffect?.targetId, delayedFade?.targetId],
      [textObject?.id, textObject?.id, textObject?.id],
    );
    assert.deepEqual(
      [flyEffect?.nativeBuild, spinEffect?.nativeBuild],
      ["paragraph", "paragraph"],
    );
    const spinPositions = spinEffect?.nativeBehaviors?.filter(
      ({ attributeNames }) =>
        attributeNames.some((name) => name === "ppt_x" || name === "ppt_y"),
    );
    assert.deepEqual(
      spinPositions?.map(
        ({
          attributeNames,
          additive,
          from,
          to,
          calculationMode,
          keyframes,
        }) => [
          attributeNames,
          additive,
          from,
          to,
          calculationMode,
          keyframes?.map(({ time, value, formula }) => [
            time,
            value,
            formula?.includes("$") ?? false,
          ]),
        ],
      ),
      [
        [
          ["ppt_x"],
          "sum",
          "0",
          "1",
          "lin",
          [
            [0, "0", true],
            [100000, "1", false],
          ],
        ],
        [
          ["ppt_y"],
          "sum",
          "0",
          "1",
          "lin",
          [
            [0, "0", true],
            [100000, "1", false],
          ],
        ],
        [
          ["ppt_x"],
          "sum",
          "0",
          "1",
          "lin",
          [
            [0, "0", true],
            [100000, "1", false],
          ],
        ],
        [
          ["ppt_y"],
          "sum",
          "0",
          "1",
          "lin",
          [
            [0, "0", true],
            [100000, "1", false],
          ],
        ],
      ],
    );
    assert.equal(
      fileEffects.filter(
        ({ targetId, preset }) =>
          targetId === textObject?.id &&
          (preset === "fly" || preset === "spinIn"),
      ).length,
      2,
    );
  });

  it("adds native paragraph build layer for pure whole Text entrance array while preserving logical whole target", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { id: "whole-text-motion-array" },
          createElement(Text, {
            id: "whole-array-text",
            x: 800,
            y: 500,
            width: 400,
            height: 160,
            enter: [
              fly({ from: "left", duration: 4000 }),
              spinIn({ duration: 4000 }),
            ],
            children: [
              createElement(
                P,
                { key: "long" },
                "A deliberately long first paragraph",
              ),
              createElement(P, { key: "short" }, "A shorter indented item"),
            ],
          }),
        ),
      ),
    );
    const logicalMotion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      logicalMotion.steps[0]?.effects.map(({ preset, target, by }) => [
        preset,
        target.range,
        by,
      ]),
      [
        ["fly", undefined, undefined],
        ["spinIn", undefined, undefined],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const effects = file.slides[0]?.animations?.[0]?.effects ?? [];
    assert.deepEqual(
      effects.map(({ preset, nativeBuild, nativeBuildGroupId }) => [
        preset,
        nativeBuild,
        nativeBuildGroupId,
      ]),
      [
        ["fly", "paragraph", 1],
        ["spinIn", "paragraph", 1],
      ],
    );
    for (const effect of effects) {
      const paragraphIndexes = new Set(
        effect.nativeBehaviors?.flatMap(({ targetType, targetRange }) =>
          targetType === "paragraph" && targetRange?.type === "paragraph"
            ? [targetRange.start]
            : [],
        ),
      );
      assert.deepEqual(
        [...paragraphIndexes].sort((left, right) => left - right),
        [0, 1],
      );
    }
  });

  it("writes eight existing whole Text presets paragraph by paragraph and independently reads back all native ranges", async () => {
    const scenarios = [
      ["fade", fade(), "enter"],
      ["wipe", wipe({ from: "left" }), "enter"],
      ["zoom", zoom(), "enter"],
      ["fly", fly({ from: "right" }), "enter"],
      ["floatUp", floatUp(), "enter"],
      ["spinIn", spinIn(), "enter"],
      ["stretch", stretch(), "enter"],
      ["spin", spin({ from: 225 }), "emphasis"],
    ] as const;
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        ...scenarios.map(([preset, effect, property]) =>
          createElement(
            Slide,
            { key: preset, id: `whole-${preset}` },
            createElement(Text, {
              id: `${preset}-text`,
              x: 800,
              y: 500,
              width: 400,
              height: 160,
              [property]: effect,
              children: [
                createElement(
                  P,
                  { key: "long" },
                  "A deliberately long paragraph",
                ),
                createElement(P, { key: "short" }, "SHORT"),
              ],
            }),
          ),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide, index) => {
        const expected = scenarios[index];
        const textObject = result.report.slides[index]?.objects[0];
        const effect = slide.animations?.[0]?.effects[0];
        const behaviors = effect?.nativeBehaviors ?? [];
        const paragraphIndexes = new Set(
          behaviors.flatMap(({ targetType, targetRange }) =>
            targetType === "paragraph" && targetRange?.type === "paragraph"
              ? [targetRange.start]
              : [],
          ),
        );
        assert.deepEqual(
          [
            effect?.targetId,
            effect?.preset,
            effect?.nativeBuild,
            effect?.nativeBuildGroupId,
          ],
          [
            textObject?.id,
            expected?.[0],
            expected?.[2] === "enter" ? "paragraph" : "allAtOnce",
            1,
          ],
        );
        assert.ok(
          behaviors.every(
            ({ targetType }) =>
              targetType === "paragraph" || targetType === "background",
          ),
        );
        return [
          effect?.preset,
          [...paragraphIndexes].sort((left, right) => left - right),
        ];
      }),
      scenarios.map(([preset]) => [preset, [0, 1]]),
    );
  });

  it("ends whole SpinIn at original text box rotation angle", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "rotated-whole-spin-in" },
          createElement(Text, {
            id: "rotated-text",
            x: 20,
            y: 20,
            width: 200,
            height: 100,
            style: { rotate: 12 },
            enter: spinIn({ duration: 1000 }),
            children: [
              createElement(P, { key: "first" }, "First paragraph"),
              createElement(P, { key: "second" }, "Second paragraph"),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const spinInEffect = file.slides[0]?.animations?.[0]?.effects[0];
    assert.deepEqual(
      spinInEffect?.nativeBehaviors
        ?.filter(({ attributeNames }) =>
          attributeNames.includes("style.rotation"),
        )
        .map(({ from, to, targetRange }) => [from, to, targetRange]),
      [
        ["360", "0", { type: "paragraph", start: 0, end: 0 }],
        ["360", "0", { type: "paragraph", start: 1, end: 1 }],
      ],
    );
  });

  it("reads decimal-layout SpinIn and Spin using actual EMU geometry", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 853, height: 617 },
        ...(["fractional-spin-in", "fractional-spin"] as const).map((slideId) =>
          createElement(
            Slide,
            { key: slideId, id: slideId },
            createElement(Text, {
              id: slideId,
              x: 71.12345,
              y: 83.12345,
              width: 317.12345,
              height: 127.12345,
              ...(slideId === "fractional-spin-in"
                ? { enter: spinIn({ duration: 1234 }) }
                : { emphasis: spin({ duration: 1234, from: 137.25 }) }),
              children: [
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ],
            }),
          ),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const firstSlideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const secondSlideXml = strFromU8(parts["ppt/slides/slide2.xml"]!);
    assert.ok(firstSlideXml.includes('<a:off x="677451" y="791751"/>'));
    assert.ok(firstSlideXml.includes('<a:ext cx="3020601" cy="1210851"/>'));
    assert.ok(secondSlideXml.includes('<a:off x="677451" y="791751"/>'));
    assert.ok(secondSlideXml.includes('<a:ext cx="3020601" cy="1210851"/>'));

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.ok(
      file.slides[0]?.animations?.some(({ effects }) =>
        effects.some(({ preset }) => preset === "spinIn"),
      ),
    );
    assert.ok(
      file.slides[1]?.animations?.some(({ effects }) =>
        effects.some(({ preset }) => preset === "spin"),
      ),
    );

    const spinInEffect = file.slides[0]?.animations
      ?.flatMap(({ effects }) => effects)
      .find(({ preset }) => preset === "spinIn");
    const positionFormula = spinInEffect?.nativeBehaviors?.find(
      ({ attributeNames }) => attributeNames[0] === "ppt_x",
    )?.keyframes?.[0]?.formula;
    assert.ok(positionFormula !== undefined);
    const centerX = /^(-?(?:\d+(?:\.\d*)?|\.\d+))\+\$\*/u.exec(
      positionFormula,
    )?.[1];
    assert.ok(centerX !== undefined);
    const presentationXml = strFromU8(parts["ppt/presentation.xml"]!);
    const pageWidthEmu = Number(
      /<p:sldSz cx="(\d+)"/u.exec(presentationXml)?.[1],
    );
    assert.ok(Number.isSafeInteger(pageWidthEmu) && pageWidthEmu > 0);
    const wrongCenterX = String(Number(centerX) + 1 / pageWidthEmu);
    const wrongCenterSlide = firstSlideXml.replace(
      /fmla="([^"]+)"/gu,
      (attribute, formula: string) =>
        formula.includes("cos(") && formula.includes("sin(")
          ? `fmla="${formula.replaceAll(centerX, wrongCenterX)}"`
          : attribute,
    );
    assert.notEqual(wrongCenterSlide, firstSlideXml);
    const wrongCenterReadback = await validatePptx(
      zipSync({
        ...parts,
        "ppt/slides/slide1.xml": strToU8(wrongCenterSlide),
      }),
    );
    assert.ok(
      wrongCenterReadback.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
      "must reject when common center deviates from actual geometry by 1 EMU",
    );
  });

  it("reads back common center when coordinates, dimensions, and slide dimensions cross EMU rounding boundaries", async () => {
    const cases = [
      {
        id: "rounding-down",
        x: (1_000_000 + 0.25) / 9525,
        y: (2_000_000 + 0.75) / 9525,
        width: (3_000_000 + 0.25) / 9525,
        height: (1_000_000 + 0.75) / 9525,
      },
      {
        id: "rounding-up",
        x: (1_500_000 + 0.75) / 9525,
        y: (2_500_000 + 0.25) / 9525,
        width: (2_000_000 + 0.75) / 9525,
        height: (1_500_000 + 0.25) / 9525,
      },
    ] as const;
    const result = await build(
      createElement(
        Deck,
        {
          width: (8_000_000 + 0.499) / 9525,
          height: (5_000_000 + 0.501) / 9525,
        },
        createElement(
          Slide,
          { id: "emu-rounding-boundaries" },
          ...cases.map(({ id: testId, ...box }) =>
            createElement(Text, {
              key: testId,
              id: testId,
              ...box,
              emphasis: spin({ duration: 1234, from: 137.25 }),
              children: createElement(P, {}, testId),
            }),
          ),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const presentationXml = strFromU8(parts["ppt/presentation.xml"]!);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.ok(presentationXml.includes('<p:sldSz cx="8000000" cy="5000001"/>'));
    assert.ok(slideXml.includes('<a:off x="1000000" y="2000001"/>'));
    assert.ok(slideXml.includes('<a:ext cx="3000000" cy="1000001"/>'));
    assert.ok(slideXml.includes('<a:off x="1500001" y="2500000"/>'));
    assert.ok(slideXml.includes('<a:ext cx="2000001" cy="1500000"/>'));

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.animations?.flatMap(({ effects }) => effects).length,
      2,
    );
  });

  it("reads back full formulas and numeric progress for common-center SpinIn and Group Spin", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "whole-fly-spin-in" },
          createElement(Text, {
            id: "whole-spin-in-text",
            x: 180,
            y: 140,
            width: 420,
            height: 160,
            style: { rotate: 12, backgroundColor: "#FEF3C7" },
            enter: [
              fly({ from: "left", duration: 4000 }),
              spinIn({ duration: 4000 }),
            ],
            children: [
              createElement(P, { key: "long" }, "A long first paragraph"),
              createElement(P, { key: "short" }, "SHORT"),
            ],
          }),
        ),
        createElement(
          Slide,
          { id: "group-spin-members" },
          createElement(
            Group,
            {
              id: "formula-spin-group",
              name: "formula-spin-group",
              emphasis: spin({ duration: 4000, from: 225 }),
            },
            createElement(Text, {
              id: "formula-spin-text-member",
              x: 110,
              y: 130,
              width: 260,
              height: 120,
              style: { backgroundColor: "#DBEAFE" },
              enter: fade({ duration: 4000 }),
              children: [
                createElement(P, { key: "long" }, "A long member paragraph"),
                createElement(P, { key: "short" }, "SHORT"),
              ],
            }),
            createElement(Shape, {
              id: "formula-spin-shape-member",
              name: "formula-spin-shape-member",
              x: 470,
              y: 230,
              width: 64,
              height: 64,
              fill: "#F97316",
            }),
          ),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.objects[0]?.rotationDegrees, 12);

    const firstSlideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const spinInEffect = file.slides[0]?.animations?.[0]?.effects.find(
      ({ preset }) => preset === "spinIn",
    );
    const spinInPositions = spinInEffect?.nativeBehaviors?.filter(
      ({ attributeNames }) =>
        attributeNames.includes("ppt_x") || attributeNames.includes("ppt_y"),
    );
    assert.equal(spinInPositions?.length, 6);
    assert.deepEqual(
      spinInPositions?.map(({ attributeNames, additive, from, to }) => [
        attributeNames,
        additive,
        from,
        to,
      ]),
      [
        [["ppt_x"], "sum", "0", "1"],
        [["ppt_y"], "sum", "0", "1"],
        [["ppt_x"], "sum", "0", "1"],
        [["ppt_y"], "sum", "0", "1"],
        [["ppt_x"], "sum", "0", "1"],
        [["ppt_y"], "sum", "0", "1"],
      ],
    );
    for (const behavior of spinInPositions ?? []) {
      assert.equal(behavior.calculationMode, "lin");
      assert.deepEqual(
        behavior.keyframes?.map(({ time, value }) => [time, value]),
        [
          [0, "0"],
          [100000, "1"],
        ],
      );
      const formula = behavior.keyframes?.[0]?.formula;
      assert.ok(formula?.includes("$"));
      assert.ok(formula?.includes("#ppt_x"));
      assert.ok(formula?.includes("#ppt_y"));
      assert.ok(formula?.includes("cos(") && formula.includes("sin("));
      assert.equal(
        formula,
        behavior.attributeNames[0] === "ppt_x"
          ? "0.40625+$*((#ppt_x-0.40625)*cos((360*pi/180)*(1-$))-(#ppt_y-0.407407407407407)*0.5625*sin((360*pi/180)*(1-$)))-#ppt_x"
          : "0.407407407407407+$*((#ppt_x-0.40625)*1.77777777777778*sin((360*pi/180)*(1-$))+(#ppt_y-0.407407407407407)*cos((360*pi/180)*(1-$)))-#ppt_y",
      );
      assert.ok(firstSlideXml.includes(`fmla="${formula}"`));
    }

    const group = file.slides[1]?.objects.find(
      ({ name }) => name === "formula-spin-group",
    );
    const memberSpins = file.slides[1]?.animations
      ?.flatMap(({ effects }) => effects)
      .filter(({ preset }) => preset === "spin");
    assert.deepEqual(
      memberSpins?.map(({ targetId }) => targetId),
      group?.members?.map(({ id }) => id),
    );
    assert.ok(
      memberSpins?.every((effect) =>
        effect.nativeBehaviors
          ?.filter(
            ({ attributeNames }) =>
              attributeNames.includes("ppt_x") ||
              attributeNames.includes("ppt_y"),
          )
          .every(
            (behavior) =>
              behavior.additive === "sum" &&
              behavior.calculationMode === "lin" &&
              behavior.keyframes?.length === 2 &&
              behavior.keyframes[0]?.time === 0 &&
              behavior.keyframes[0]?.value === "0" &&
              behavior.keyframes[1]?.time === 100000 &&
              behavior.keyframes[1]?.value === "1" &&
              behavior.keyframes[0]?.formula?.includes("(225*pi/180)") === true,
          ),
      ),
    );
    assert.ok(
      memberSpins?.every((effect) =>
        effect.nativeBehaviors?.some(
          ({ behavior, from, to, additive }) =>
            behavior === "animRot" &&
            from === "13500000" &&
            to === "0" &&
            additive === "sum",
        ),
      ),
    );

    const textId = result.report.slides[0]?.objects[0]?.id;
    assert.ok(textId !== undefined);
    const geometryXml = [
      ...firstSlideXml.matchAll(/<p:anim\b[\s\S]*?<\/p:anim>/gu),
    ]
      .map(([behavior]) => behavior ?? "")
      .find(
        (behavior) =>
          behavior.includes("fmla=") &&
          behavior.includes(
            `<p:spTgt spid="${textId}"><p:txEl><p:pRg st="0" end="0"></p:pRg></p:txEl></p:spTgt>`,
          ) &&
          behavior.includes("<p:attrName>ppt_x</p:attrName>"),
      );
    assert.ok(geometryXml !== undefined);
    const formula = spinInPositions?.[0]?.keyframes?.[0]?.formula;
    assert.ok(formula !== undefined);
    const slidePath = "ppt/slides/slide1.xml";
    const parts = unzipSync(result.pptx);
    const validateSlide = (xml: string) =>
      validatePptx(zipSync({ ...parts, [slidePath]: strToU8(xml) }));

    const missingFormulaXml = geometryXml.replace(/\s+fmla="[^"]+"/u, "");
    assert.notEqual(missingFormulaXml, geometryXml);
    const missingFormula = await validateSlide(
      firstSlideXml.replace(geometryXml, missingFormulaXml),
    );
    assert.ok(
      missingFormula.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );

    const wrongFormula = formula.replace("0.40625", "0.5");
    assert.notEqual(wrongFormula, formula);
    const changedFormula = await validateSlide(
      firstSlideXml.replace(
        geometryXml,
        geometryXml.replace(`fmla="${formula}"`, `fmla="${wrongFormula}"`),
      ),
    );
    assert.ok(
      changedFormula.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );

    const wrongSum = await validateSlide(
      firstSlideXml.replace(
        geometryXml,
        geometryXml.replace('additive="sum"', 'additive="base"'),
      ),
    );
    assert.ok(
      wrongSum.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
    );

    const missingPositionRole = await validateSlide(
      firstSlideXml.replace(geometryXml, ""),
    );
    assert.ok(
      missingPositionRole.errors.some(
        ({ code }) => code === "invalid-animation-projection-target",
      ),
    );

    const groupSlidePath = "ppt/slides/slide2.xml";
    const groupSlideXml = strFromU8(parts[groupSlidePath]!);
    const validateGroupSlide = (xml: string) =>
      validatePptx(zipSync({ ...parts, [groupSlidePath]: strToU8(xml) }));
    const rewriteFormulas = (
      xml: string,
      rewrite: (formula: string) => string,
      targetId?: number,
    ) =>
      xml.replace(/<p:anim\b[\s\S]*?<\/p:anim>/gu, (behavior) => {
        if (
          targetId !== undefined &&
          !behavior.includes(`<p:spTgt spid="${targetId}">`)
        ) {
          return behavior;
        }
        return behavior.replace(
          /fmla="([^"]+)"/u,
          (_match, formula: string) => `fmla="${rewrite(formula)}"`,
        );
      });
    const invalidGeometryInputs = [
      [
        "position function identically zero",
        () =>
          validateSlide(
            rewriteFormulas(firstSlideXml, (formula) => `0*(${formula})`),
          ),
      ],
      [
        "sine and cosine swapped",
        () =>
          validateSlide(
            rewriteFormulas(firstSlideXml, (formula) =>
              formula
                .replaceAll("cos(", "__cos__(")
                .replaceAll("sin(", "cos(")
                .replaceAll("__cos__(", "sin("),
            ),
          ),
      ],
      [
        "wrong slide aspect ratio",
        () =>
          validateSlide(
            rewriteFormulas(firstSlideXml, (formula) =>
              formula
                .replaceAll("0.5625", "1")
                .replaceAll("1.77777777777778", "1"),
            ),
          ),
      ],
      [
        "wrong whole text center",
        () =>
          validateSlide(
            rewriteFormulas(firstSlideXml, (formula) =>
              formula
                .replaceAll("0.40625", "0.9")
                .replaceAll("0.407407407407407", "0.1"),
            ),
          ),
      ],
    ] as const;
    for (const [name, readback] of invalidGeometryInputs) {
      const invalid = await readback();
      assert.ok(
        invalid.errors.some(
          ({ code }) => code === "invalid-animation-progress-formula",
        ),
        `${name} in actual animation XML must be rejected`,
      );
    }

    const groupShapeId = group?.members?.find(
      ({ name }) => name === "formula-spin-shape-member",
    )?.id;
    assert.ok(groupShapeId !== undefined);
    const wrongMemberAngle = await validateGroupSlide(
      rewriteFormulas(
        groupSlideXml,
        (formula) => formula.replaceAll("225*pi/180", "30*pi/180"),
        groupShapeId,
      ),
    );
    assert.ok(
      wrongMemberAngle.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
      "member position angle must match actual animRot angle",
    );

    const wrongMemberCenter = await validateGroupSlide(
      rewriteFormulas(
        groupSlideXml,
        (formula) => {
          const centerX = formula.match(
            /#ppt_x-(-?(?:\d+(?:\.\d*)?|\.\d+))/u,
          )?.[1];
          const centerY = formula.match(
            /#ppt_y-(-?(?:\d+(?:\.\d*)?|\.\d+))/u,
          )?.[1];
          assert.ok(centerX !== undefined && centerY !== undefined);
          return formula.replaceAll(centerX, "0.7").replaceAll(centerY, "0.2");
        },
        groupShapeId,
      ),
    );
    assert.ok(
      wrongMemberCenter.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
      "member position center must match actual group geometry",
    );

    const zeroFunctionReadback = await validateSlide(
      rewriteFormulas(firstSlideXml, (formula) => `0*(${formula})`),
    );
    const preservedSpinIn = zeroFunctionReadback.slides[0]?.animations
      ?.flatMap(({ effects }) => effects)
      .find(({ preset }) => preset === "spinIn");
    assert.ok(
      preservedSpinIn?.nativeBehaviors?.some(
        ({ keyframes }) => keyframes?.[0]?.formula?.startsWith("0*(") === true,
      ),
      "unsupported formulas must still be preserved in readback results for diagnostics",
    );
  });

  it("preserves object initial state for relative Spin and reads back positive and negative rotation angles", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { id: "relative-spin" },
          createElement(Shape, {
            id: "positive-spin-shape",
            x: 80,
            y: 60,
            width: 64,
            height: 64,
            style: { rotate: 17 },
            fill: "#F97316",
            emphasis: spin({ duration: 1200, by: 120 }),
          }),
          createElement(
            Text,
            {
              id: "legacy-spin-text",
              x: 80,
              y: 150,
              width: 120,
              height: 40,
              emphasis: spin({
                duration: 1200,
                from: 90,
                by: "letter",
                intervalMs: 50,
              }),
            },
            "two words",
          ),
          createElement(
            Group,
            {
              id: "negative-spin-group",
              emphasis: spin({ duration: 1600, by: -80 }),
            },
            createElement(Text, {
              id: "negative-spin-text",
              x: 220,
              y: 130,
              width: 240,
              height: 120,
              style: { rotate: 12, backgroundColor: "#FEF3C7" },
              enter: fade({ duration: 100 }),
              children: [
                createElement(P, { key: "first" }, "First paragraph"),
                createElement(P, { key: "second" }, "Second paragraph"),
              ],
            }),
            createElement(Shape, {
              id: "negative-spin-member",
              x: 520,
              y: 230,
              width: 64,
              height: 64,
              fill: "#1D4ED8",
            }),
          ),
        ),
      ),
    );
    const slide = result.query.getSlide("relative-spin");
    const queryEffects = observeMotion(slide.observation)?.steps.flatMap(
      ({ effects }) => effects,
    );
    const positiveQueryEffect = queryEffects?.find(
      ({ target }) => target.authorId === "positive-spin-shape",
    );
    const groupQueryEffect = queryEffects?.find(
      ({ target }) => target.authorId === "negative-spin-group",
    );
    const legacyQueryEffect = queryEffects?.find(
      ({ target }) => target.authorId === "legacy-spin-text",
    );
    assert.equal(positiveQueryEffect?.from, undefined);
    assert.equal(positiveQueryEffect?.by, 120);
    assert.equal(groupQueryEffect?.from, undefined);
    assert.equal(groupQueryEffect?.by, -80);
    assert.equal(legacyQueryEffect?.from, 90);
    assert.equal(legacyQueryEffect?.by, "letter");
    assert.equal(legacyQueryEffect?.intervalMs, 50);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const fileSlide = file.slides[0];
    const positiveShapeObservation = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "positive-spin-shape",
    );
    const positiveShape = fileSlide?.objects.find(
      ({ id }) => id === positiveShapeObservation?.id,
    );
    const spinEffects = fileSlide?.animations?.flatMap(
      ({ effects }) => effects,
    );
    const positiveSpin = spinEffects?.find(
      ({ targetId }) => targetId === positiveShape?.id,
    );
    assert.ok(positiveShapeObservation);
    assert.ok(positiveShape);
    assert.equal(positiveShape?.rotationDegrees, 17);
    assert.equal(positiveSpin?.from, undefined);
    assert.equal(positiveSpin?.by, 120);
    const positiveRotation = positiveSpin?.nativeBehaviors?.find(
      ({ behavior }) => behavior === "animRot",
    );
    assert.ok(positiveRotation);
    assert.deepEqual(
      [
        positiveRotation.by,
        positiveRotation.from,
        positiveRotation.to,
        positiveRotation.additive,
      ],
      ["7200000", undefined, undefined, "sum"],
    );

    const legacyTextObservation = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "legacy-spin-text",
    );
    const legacySpin = spinEffects?.find(
      ({ targetId }) => targetId === legacyTextObservation?.id,
    );
    assert.equal(legacySpin?.from, 90);
    assert.equal(legacySpin?.by, "letter");
    assert.equal(legacySpin?.intervalMs, 50);

    const negativeGroupObservation = result.report.slides[0]?.objects.find(
      ({ authorId: testId }) => testId === "negative-spin-group",
    );
    const negativeGroup = fileSlide?.objects.find(
      ({ id }) => id === negativeGroupObservation?.id,
    );
    const negativeTextObservation = negativeGroupObservation?.members?.find(
      ({ authorId: testId }) => testId === "negative-spin-text",
    );
    const negativeText = negativeGroup?.members?.find(
      ({ id }) => id === negativeTextObservation?.id,
    );
    assert.equal(negativeText?.rotationDegrees, 12);
    const negativeSpins = spinEffects?.filter(
      ({ preset, targetId }) =>
        preset === "spin" &&
        negativeGroup?.members?.some(({ id }) => id === targetId),
    );
    assert.ok(negativeSpins && negativeSpins.length >= 2);
    assert.ok(
      negativeSpins.every(
        (effect) =>
          effect.from === undefined &&
          effect.by === -80 &&
          effect.nativeBehaviors?.some(
            ({ behavior, by, additive }) =>
              behavior === "animRot" && by === "-4800000" && additive === "sum",
          ) === true,
      ),
    );
    assert.ok(
      negativeSpins.some((effect) =>
        effect.nativeBehaviors?.some(
          ({ attributeNames, keyframes }) =>
            attributeNames.includes("ppt_x") &&
            keyframes?.[0]?.formula?.includes("(-80*pi/180)*$") === true,
        ),
      ),
    );
    const groupMotion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const groupProjection = groupMotion.nativeProjections?.find(
      ({ logicalTargetId, preset }) =>
        logicalTargetId === negativeGroup?.id && preset === "spin",
    );
    assert.deepEqual(
      groupProjection?.nativeEffects.map(({ targetId }) => targetId),
      negativeGroup?.members?.map(({ id }) => id),
    );

    const positionFormula = negativeSpins[0]?.nativeBehaviors?.find(
      ({ attributeNames }) => attributeNames.includes("ppt_x"),
    )?.keyframes?.[0]?.formula;
    assert.ok(positionFormula !== undefined);
    const mixedProgressFormula = positionFormula.replace(
      /(sin\(\([^)]*\*pi\/180\)\*)\$(\))/u,
      "$1(1-$)$2",
    );
    assert.notEqual(mixedProgressFormula, positionFormula);
    const relativeParts = unzipSync(result.pptx);
    const relativeSlidePath = "ppt/slides/slide1.xml";
    const relativeSlideXml = strFromU8(relativeParts[relativeSlidePath]!);
    assert.ok(relativeSlideXml.includes(`fmla="${positionFormula}"`));
    const mixedProgressReadback = await validatePptx(
      zipSync({
        ...relativeParts,
        [relativeSlidePath]: strToU8(
          relativeSlideXml.replace(
            `fmla="${positionFormula}"`,
            `fmla="${mixedProgressFormula}"`,
          ),
        ),
      }),
    );
    assert.ok(
      mixedProgressReadback.errors.some(
        ({ code }) => code === "invalid-animation-progress-formula",
      ),
      "the same coordinate formula for relative Spin cannot mix increasing and decreasing progress",
    );
  });

  it("rejects invalid effect parameters, missing motion capabilities, and non-existent file targets", async () => {
    for (const [effect, code] of [
      [fade({ duration: 0 }), "invalid-duration"],
      [fade({ duration: null } as never), "invalid-duration"],
      [fade({ delay: -1 }), "invalid-delay"],
      [fade({ unexpected: true } as never), "unknown-parameter"],
      [fade({ easing: null } as never), "invalid-easing"],
      [fade({ by: "letter", intervalMs: 0 }), "invalid-motion-interval"],
      [
        fade({ easing: { accelerate: 0.8, decelerate: 0.4 } }),
        "invalid-easing",
      ],
      [wipe({ from: null } as never), "invalid-direction"],
      [floatDown({ from: "top" } as never), "invalid-motion"],
      [fade({ by: "paragraph" }), "invalid-motion-granularity"],
    ] as const) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { id: "invalid" },
              createElement(Shape, {
                id: "invalid-effect",
                x: 0,
                y: 0,
                width: 16,
                height: 16,
                enter: effect,
              }),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          assert.match(
            error.diagnostics[0]?.path ?? "",
            /shape\[invalid-effect\]/u,
          );
          return true;
        },
      );
    }

    await assert.rejects(
      async () =>
        emit(
          createResolver([shapeCapability, transitionCapability])(
            await expandDocument(
              createElement(
                Deck,
                {},
                createElement(
                  Slide,
                  { id: "trimmed" },
                  createElement(Shape, {
                    id: "needs-motion",
                    x: 0,
                    y: 0,
                    width: 16,
                    height: 16,
                    enter: fade(),
                  }),
                ),
              ),
            ),
          ),
        ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /shape\[needs-motion\]/u,
        );
        assert.match(error.diagnostics[0]?.message ?? "", /motion/u);
        return true;
      },
    );

    const result = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { id: "target" },
          createElement(Shape, {
            id: "target-object",
            x: 0,
            y: 0,
            width: 16,
            height: 16,
            enter: fade(),
          }),
        ),
      ),
    );
    const targetId = result.report.slides[0]?.objects[0]?.id;
    const parts = unzipSync(result.pptx);
    const slidePath = "ppt/slides/slide1.xml";
    const slide = parts[slidePath];
    assert.ok(slide);
    const malformedSlide = strFromU8(slide).replaceAll(
      `<p:spTgt spid="${targetId}">`,
      '<p:spTgt spid="9999">',
    );
    assert.notEqual(malformedSlide, strFromU8(slide));
    parts[slidePath] = strToU8(malformedSlide);
    const invalidFile = await validatePptx(zipSync(parts));
    assert.ok(
      invalidFile.errors.some(
        ({ code }) => code === "missing-animation-target",
      ),
      JSON.stringify(invalidFile.errors),
    );

    const mismatchedPresetClass = strFromU8(slide).replace(
      'presetClass="entr"',
      'presetClass="emph"',
    );
    assert.notEqual(mismatchedPresetClass, strFromU8(slide));
    parts[slidePath] = strToU8(mismatchedPresetClass);
    const invalidKindFile = await validatePptx(zipSync(parts));
    assert.ok(
      invalidKindFile.errors.some(
        ({ code }) => code === "invalid-animation-kind",
      ),
      JSON.stringify(invalidKindFile.errors),
    );
  });

  it("delegates motion parameters to formal validation in wrapper components and preserves disabled semantics", async () => {
    const animatedShape = (testId: string) =>
      createElement(Shape, {
        id: testId,
        x: 0,
        y: 0,
        width: 20,
        height: 20,
        enter: fade({ duration: 100 }),
      });
    const scopeDeck = (props: Record<string, unknown>) => {
      function ScopedMotion() {
        const scope = useMotionScope();
        return createElement(
          MotionScope,
          { scope, ...props } as never,
          animatedShape("scope-child"),
        );
      }
      return createElement(
        Deck,
        {},
        createElement(
          Slide,
          { id: "wrapper-motion-parameters" },
          createElement(ScopedMotion),
        ),
      );
    };
    const slideDeck = (child: ReactNode) =>
      createElement(
        Deck,
        {},
        createElement(Slide, { id: "wrapper-motion-parameters" }, child),
      );
    const invalidInputs = [
      {
        element: scopeDeck({ motion: "false" }),
        code: "invalid-motion",
        path: /motion-scope/u,
      },
      {
        element: slideDeck(
          createElement(
            Box,
            {
              motion: "false" as never,
              style: { width: 100, height: 80 },
            },
            animatedShape("box-child"),
          ),
        ),
        code: "invalid-motion",
        path: /box/u,
      },
      {
        element: scopeDeck({ delayMs: 800 }),
        code: "unknown-parameter",
        path: /motion-scope/u,
      },
      {
        element: slideDeck(
          createElement(
            Step,
            { index: 2 } as never,
            animatedShape("step-child"),
          ),
        ),
        code: "unknown-parameter",
        path: /motion-step/u,
      },
    ] as const;
    for (const { element, code, path } of invalidInputs) {
      await assert.rejects(build(element), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(
          error.diagnostics[0]?.code,
          code,
          JSON.stringify(error.diagnostics),
        );
        assert.match(error.diagnostics[0]?.path ?? "", path);
        return true;
      });
    }

    const disabledScope = await build(scopeDeck({ motion: false }));
    const disabledScopeMotion = disabledScope.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation | undefined;
    assert.ok(disabledScopeMotion);
    assert.equal(
      disabledScopeMotion.steps.flatMap(({ effects }) => effects).length,
      0,
    );
    assert.deepEqual(
      (await validatePptx(disabledScope.pptx)).slides[0]?.animations,
      undefined,
    );

    const disabledBox = await build(
      slideDeck(
        createElement(
          Box,
          {
            motion: false,
            style: {
              width: 100,
              height: 80,
              backgroundColor: "#123456",
            },
          },
          animatedShape("disabled-box-child"),
        ),
      ),
    );
    assert.deepEqual(
      (await validatePptx(disabledBox.pptx)).slides[0]?.animations,
      undefined,
    );
  });

  it("rejects declarations exceeding native angle ranges or mismatching effect presets", async () => {
    const cases: readonly {
      readonly id: string;
      readonly props: Record<string, unknown>;
      readonly code: string;
    }[] = [
      {
        id: "rotation-over-int32",
        props: { emphasis: spin({ from: 40_000 }) },
        code: "invalid-rotation",
      },
      {
        id: "rotation-too-large",
        props: { emphasis: spin({ from: 1e20 }) },
        code: "invalid-rotation",
      },
      {
        id: "rotation-null-from",
        props: { emphasis: spin({ from: null } as never) },
        code: "invalid-rotation",
      },
      {
        id: "rotation-by-over-int32",
        props: { emphasis: spin({ by: 40_000 }) },
        code: "invalid-rotation",
      },
      {
        id: "rotation-from-and-by",
        props: { emphasis: spin({ from: 90, by: -90 } as never) },
        code: "invalid-motion",
      },
      {
        id: "spin-in-enter",
        props: { enter: { ...spin(), kind: "entrance" } },
        code: "invalid-motion",
      },
      {
        id: "fade-as-emphasis",
        props: { emphasis: { ...fade(), kind: "emphasis" } },
        code: "invalid-motion",
      },
    ];
    for (const { id: testId, props, code } of cases) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { id: "invalid-preset" },
              createElement(Shape, {
                id: testId,
                x: 0,
                y: 0,
                width: 16,
                height: 16,
                ...props,
              } as never),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          assert.match(
            error.diagnostics[0]?.path ?? "",
            new RegExp(`shape\\[${testId}\\]`, "u"),
          );
          return true;
        },
      );
    }
  });

  it("writes transitions and animations into fixed slide slots regardless of module registration order", async () => {
    const resolveReversed = createResolver([
      motionCapability,
      transitionCapability,
      shapeCapability,
    ]);
    const result = emit(
      resolveReversed(
        await expandDocument(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { id: "slot-order", transition: transitions.fade() },
              createElement(Shape, {
                id: "timed",
                x: 0,
                y: 0,
                width: 16,
                height: 16,
                enter: fade(),
              }),
            ),
          ),
        ),
      ),
    );
    const slidePart = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    assert.ok(slidePart);
    const slideXml = strFromU8(slidePart);
    assert.ok(
      slideXml.indexOf("<p:transition") < slideXml.indexOf("<p:timing>"),
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });
});
