import assert from "node:assert/strict";

import {
  build,
  BuildError,
  Chart,
  createBuilder,
  Deck,
  breathe,
  fade,
  fly,
  floatUp,
  Group,
  Image,
  MotionScope,
  P,
  Shape,
  shapeCapability,
  Slide,
  spin,
  spinIn,
  stretch,
  Step,
  Text,
  motionCapability,
  repeat,
  transitionCapability,
  transitions,
  type MotionObservation,
  useMotionScope,
  wipe,
  zoom,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement, Fragment } from "react";
import { describe, it } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

function RepeatedSpinDeck() {
  const body = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "finite-repeat" },
      createElement(
        MotionScope,
        { scope: body },
        createElement(Shape, {
          testId: "three-turns",
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
      { slideId: "time-projection" },
      createElement(
        MotionScope,
        { scope: outer, delay: outerDelay },
        createElement(
          MotionScope,
          { scope: inner, delay: innerDelay },
          createElement(Shape, {
            testId: "timed-repeat",
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
      { slideId: "breathing-scope" },
      createElement(
        MotionScope,
        { scope: choreography },
        createElement(
          MotionScope,
          { scope: body },
          createElement(Shape, {
            testId: "finite-body",
            name: "Finite body - spin three turns",
            x: 20,
            y: 20,
            width: 40,
            height: 40,
            fill: "#164E3D",
            emphasis: repeat(spin({ duration: 1200 }), 3),
          }),
          createElement(Shape, {
            testId: "finite-breathe",
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
              testId: "native-lamp-group",
              name: "Native group - two spins",
              emphasis: repeat(spin({ duration: 1200 }), 2),
            },
            createElement(Shape, {
              testId: "infinite-lamp",
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
          testId: "after-body",
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
          testId: "first-click",
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
          testId: "second-click",
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
      { slideId: "after-breathing-slide" },
      createElement(Shape, {
        testId: "next-page",
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
      { slideId: blockAfter ? "blocked-infinite-tail" : "infinite-tail" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        createElement(Shape, {
          testId: "before-lamp",
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
            testId: "tail-lamp",
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
              testId: "blocked-follower",
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

describe("对象动画", () => {
  it("有限重复按总轮数进入编排完成与实际文件", async () => {
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

  it("连续呼吸保持无自然完成且只等待有限正文作用域", async () => {
    const result = await build(createElement(RepeatedAndBreathingDeck));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const effects = motion.steps.flatMap((step) => step.effects);
    const bodyEffect = effects.find(
      ({ target }) => target.testId === "finite-body",
    );
    const finiteBreathe = effects.find(
      ({ target }) => target.testId === "finite-breathe",
    );
    const lampEffect = effects.find(
      ({ target }) => target.testId === "infinite-lamp",
    );
    const groupEffect = effects.find(
      ({ target }) => target.testId === "native-lamp-group",
    );
    const follower = effects.find(
      ({ target }) => target.testId === "after-body",
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
      ({ testId }) => testId === "finite-body",
    );
    const finiteBreatheObject = slide?.objects.find(
      ({ testId }) => testId === "finite-breathe",
    );
    const lampObject = slide?.objects
      .flatMap((object) => [object, ...(object.members ?? [])])
      .find(({ testId }) => testId === "infinite-lamp");
    const groupObject = slide?.objects.find(
      ({ testId }) => testId === "native-lamp-group",
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

  it("独立读取器拒绝等待没有自然完成的文件效果", async () => {
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

  it("按实际计时字段限制范围并保留可精确表达的长重复", async () => {
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
          { slideId: "maximum-breathe-period" },
          createElement(Shape, {
            testId: "maximum-breathe-period",
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

  it("保留可分别编码的嵌套延迟并拒绝失去整数精度的已知时刻", async () => {
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

  it("允许顺序末项持续循环并拒绝被它阻挡的顺序后继", async () => {
    const result = await build(
      createElement(InfiniteSequenceTail, { blockAfter: false }),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs, completion }) => [
        target.testId,
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

  it("拒绝有限重复的非法次数、效果与呼吸参数", async () => {
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
        testId: "repeat-fade",
        props: {
          enter: repeat(fade({ duration: 100 }), 2),
        },
        code: "unsupported-repeat",
      },
      {
        testId: "repeat-zero",
        props: { emphasis: repeat(spin(), 0) },
        code: "invalid-repeat-count",
      },
      {
        testId: "repeat-missing-count",
        props: { emphasis: repeat(spin(), undefined as never) },
        code: "invalid-repeat-count",
      },
      {
        testId: "repeat-count-limit",
        props: { emphasis: repeat(spin(), 4_294_968) },
        code: "invalid-repeat-count",
      },
      {
        testId: "repeat-duration-limit",
        props: {
          emphasis: repeat(spin({ duration: 4_294_967_296 }), 1),
        },
        code: "invalid-repeat-duration",
      },
      {
        testId: "repeat-forever-spin",
        props: { emphasis: repeat(spin(), "forever") },
        code: "unsupported-repeat",
      },
      {
        testId: "repeat-granularity",
        props: { emphasis: repeat(spin({ by: "letter" }), 2) },
        code: "unsupported-repeat-combination",
      },
      {
        testId: "repeat-breathe-granularity",
        props: {
          fill: "#164E3D",
          emphasis: repeat(breathe(breatheWithGranularity), 2),
        },
        code: "unsupported-repeat-combination",
      },
      {
        testId: "breathe-unsupported-duration",
        props: {
          fill: "#164E3D",
          emphasis: breathe(breatheWithUnsupportedDuration),
        },
        code: "unknown-parameter",
      },
      {
        testId: "breathe-null-easing",
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
        testId: "breathe-false-easing",
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
        testId: "odd-breathe-period",
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
        testId: "breathe-period-limit",
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
        testId: "dark-breathe-endpoint",
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
        testId: "unfilled-breathe-target",
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

    for (const { testId, props, code } of cases) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { slideId: `invalid-${testId}` },
              createElement(Shape, {
                testId,
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

  it("呼吸颜色只绑定有填充的 Shape 成员，不绑定原生 Group 容器", async () => {
    await assert.rejects(
      build(
        createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "group-breathe-target" },
            createElement(
              Group,
              {
                testId: "group-breathe",
                emphasis: breathe({
                  period: 1600,
                  from: "#164E3D",
                  to: "#4ADE80",
                }),
              },
              createElement(Shape, {
                testId: "group-breathe-child",
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

  it("段落效果与整体效果并存时保留段落构建且不改变效果顺序", async () => {
    const cases: readonly {
      readonly testId: string;
      readonly paragraphs: readonly string[];
      readonly paragraphEffectFirst: boolean;
    }[] = [
      {
        testId: "mixed-one-paragraph",
        paragraphs: ["一段文字😀"],
        paragraphEffectFirst: true,
      },
      {
        testId: "mixed-two-paragraphs",
        paragraphs: ["第一段😀", "第二段"],
        paragraphEffectFirst: false,
      },
    ];

    for (const { testId, paragraphs, paragraphEffectFirst } of cases) {
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
            { slideId: testId },
            createElement(
              Text,
              {
                testId,
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
            effectTarget.testId,
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
              ["fly", 1000, undefined, "allAtOnce"],
            ]
          : [
              ["fly", 1000, undefined, "allAtOnce"],
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

  it("按原生段落范围逐段生成文字效果并从同次文件读回", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "paragraph-motion" },
          createElement(
            Text,
            {
              testId: "paragraphs",
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

  it("保留逐字未知完成并用单个原生 end 事件启动后继", async () => {
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
              testId: "letters",
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
          testId: "words",
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
            testId: "after-letters",
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
          { slideId: "letter-end" },
          createElement(LetterAndFollower),
        ),
      ),
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const effects = motion.steps[0]?.effects ?? [];
    const source = effects.find(({ target }) => target.testId === "letters");
    const follower = effects.find(
      ({ target }) => target.testId === "after-letters",
    );
    const wordEffect = effects.find(({ target }) => target.testId === "words");
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

  it("按原生系列和分类范围展开图表效果并保留图表对象", async () => {
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
          { slideId: "chart-motion" },
          createElement(Chart, {
            testId: "series-chart",
            x: 10,
            y: 10,
            width: 210,
            height: 200,
            kind: "column",
            data,
            enter: wipe({ by: "series", duration: 400, from: "left" }),
          }),
          createElement(Chart, {
            testId: "category-chart",
            x: 240,
            y: 10,
            width: 210,
            height: 200,
            kind: "line",
            data,
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
        target.testId,
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

  it("从公开构建声明生成淡入，并从同次 PPTX 读回目标和时间", async () => {
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
          { slideId: "motion" },
          createElement(Shape, {
            testId: "moving",
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
            target: { id: target?.id, name: "moving", testId: "moving" },
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

  it("为基础原生对象能力写入可读回的目标引用", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "native-targets" },
          createElement(Shape, {
            testId: "shape-target",
            x: 0,
            y: 0,
            width: 16,
            height: 16,
            enter: fade(),
          }),
          createElement(Text, {
            testId: "text-target",
            x: 20,
            y: 0,
            width: 80,
            height: 24,
            emphasis: spin({ from: -90 }),
            children: "Text target",
          }),
          createElement(Image, {
            testId: "image-target",
            src: png,
            x: 110,
            y: 0,
            width: 16,
            height: 16,
            enter: wipe({ from: "right" }),
          }),
          createElement(
            Group,
            { testId: "group-target", enter: zoom() },
            createElement(Shape, {
              testId: "group-member",
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
        target.testId,
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

  it("点击步按声明身份分组，自动编号不与显式编号合并", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "steps" },
          createElement(Shape, {
            testId: "automatic-entry",
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
              testId: "explicit-first",
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
              testId: "generated-first",
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
              testId: "explicit-second",
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
              testId: "generated-second",
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
              testId: "disabled",
              x: 60,
              y: 0,
              width: 10,
              height: 10,
              motion: false,
              enter: fade(),
            }),
            createElement(Shape, {
              testId: "static",
              x: 72,
              y: 0,
              width: 10,
              height: 10,
            }),
            createElement(
              Group,
              { testId: "disabled-group", motion: false },
              createElement(Shape, {
                testId: "disabled-descendant",
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
              testId: "explicit-after-ignored",
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
          target.testId,
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
            ?.testId,
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

  it("写入原型常见的八种基础效果及适用的方向、旋转和 easing 参数", async () => {
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
          { slideId: "presets" },
          ...effects.map(([testId, effect, prop], index) =>
            createElement(Shape, {
              key: testId,
              testId,
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
        target.testId,
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

  it("将同一原生 Text 的 whole Fly 展开为共同段落位移并读回背景角色", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { slideId: "whole-text-projection" },
          createElement(Text, {
            testId: "mixed-text",
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
        target.testId,
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
      ["allAtOnce", 1, true],
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
        ["set", "object", undefined, ["style.visibility"], 1],
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
      `<p:bldP spid="${textObject?.id}" grpId="1" build="allAtOnce" animBg="1"></p:bldP>`,
      `<p:bldP spid="${textObject?.id}" grpId="2" build="allAtOnce" animBg="1"></p:bldP>`,
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
      `<p:bldP spid="${textObject?.id}" grpId="1" build="allAtOnce" animBg="1"></p:bldP>`,
      `<p:bldP spid="${textObject?.id}" grpId="1" build="allAtOnce" animBg="0"></p:bldP>`,
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

  it("在 whole 数组的共享文字构建组中保留 Fly 的背景角色", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { slideId: "whole-text-background-array" },
          createElement(Text, {
            testId: "filled-whole-array",
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

  it("将 whole Text 滤镜和 SpinIn 尺寸投影到真实背景", async () => {
    const scenarios = [
      {
        preset: "fade",
        effect: fade({ duration: 4000 }),
        slideId: "fade-background",
        roles: ["animEffect:fade"],
      },
      {
        preset: "wipe",
        effect: wipe({ from: "left", duration: 4000 }),
        slideId: "wipe-background",
        roles: ["animEffect:wipe(left)"],
      },
      {
        preset: "zoom",
        effect: zoom({ duration: 4000 }),
        slideId: "zoom-background",
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
        slideId: "float-up-background",
        roles: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
      },
      {
        preset: "spinIn",
        effect: spinIn({ duration: 4000 }),
        slideId: "spin-in-background",
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
        ...scenarios.map(({ preset, effect, slideId }) =>
          createElement(
            Slide,
            { key: preset, slideId },
            createElement(Text, {
              testId: `${preset}-text`,
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
          ({ targetType }) => targetType === "background",
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
          nativeBehavior?.includes(target) === true &&
          nativeBehavior.includes(marker),
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
        ({ code }) => code === "inconsistent-animation-background-role",
      ),
    );
  });

  it("whole Stretch 从共同文字中心展开并读回段落和背景位移", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { slideId: "whole-stretch-common-center" },
          createElement(Text, {
            testId: "stretched-text",
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
        nativeBehavior?.includes(target) === true &&
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

  it("whole Text 的未来点击入口先隐藏，同时保留 by-P 延迟构建", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "future-whole-text" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              Group,
              {
                testId: "first-click-group",
                emphasis: repeat(spin({ duration: 1600 }), 2),
              },
              createElement(Shape, {
                testId: "first-click-member",
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
              testId: "future-text",
              x: 180,
              y: 405,
              width: 600,
              height: 55,
              enter: fade({ duration: 500 }),
              children: [
                createElement(P, { key: "line" }, "Visible on click two"),
              ],
            }),
          ),
        ),
        createElement(
          Slide,
          { slideId: "paragraph-delay-control" },
          createElement(Text, {
            testId: "paragraph-delay-text",
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
      ({ testId }) => testId === "future-text",
    )?.id;
    assert.ok(futureTextId !== undefined);
    const futureXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const visibilitySets = [...futureXml.matchAll(/<p:set>[\s\S]*?<\/p:set>/gu)]
      .map(([nativeSet]) => nativeSet ?? "")
      .filter((nativeSet) =>
        nativeSet.includes(`<p:spTgt spid="${futureTextId}"></p:spTgt>`),
      )
      .map((nativeSet) => nativeSet.match(/<p:strVal val="([^"]+)"/u)?.[1]);
    assert.deepEqual(visibilitySets, ["hidden", "visible"]);
    const mainSequenceStart = futureXml.indexOf('nodeType="mainSeq"');
    const firstClickCondition = futureXml.indexOf(
      'delay="indefinite"',
      mainSequenceStart,
    );
    assert.ok(
      mainSequenceStart >= 0 &&
        firstClickCondition >= 0 &&
        futureXml.indexOf('val="hidden"', mainSequenceStart) <
          firstClickCondition,
    );
    const futureFade = file.slides[0]?.animations?.[1]?.effects.find(
      ({ targetId, preset }) => targetId === futureTextId && preset === "fade",
    );
    assert.deepEqual(
      [futureFade?.nativeBuild, futureFade?.nativeBuildGroupId],
      ["allAtOnce", 1],
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

  it("Group Spin 投影到 Text 成员时保留段落和背景角色", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "group-text-spin-projection" },
          createElement(
            Group,
            {
              testId: "text-member-group",
              emphasis: repeat(spin({ duration: 4000, from: 360 }), 1),
            },
            createElement(Text, {
              testId: "animated-text-member",
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
              testId: "static-orange-member",
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
      ({ testId }) => testId === "text-member-group",
    );
    const text = group?.members?.find(
      ({ testId }) => testId === "animated-text-member",
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
      ["allAtOnce", 1, true],
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
        ) === true && nativeBehavior.includes("<p:attrName>ppt_x</p:attrName>"),
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

  it("Group Text 段落动画与投影共享文字构建职责", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "group-text-paragraph-build" },
          createElement(
            Group,
            {
              testId: "paragraph-build-group",
              emphasis: spin({ duration: 4000, from: 360 }),
            },
            createElement(Text, {
              testId: "paragraph-build-text",
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
              testId: "paragraph-build-shape",
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
      ({ testId }) => testId === "paragraph-build-group",
    );
    const text = group?.members?.find(
      ({ testId }) => testId === "paragraph-build-text",
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

  it("Group Text 逐词投影的构建组无需背景角色", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "group-text-word-build" },
          createElement(
            Group,
            { testId: "word-build-group", emphasis: spin({ duration: 4000 }) },
            createElement(Text, {
              testId: "word-build-text",
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
              testId: "word-build-shape",
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
      ({ testId }) => testId === "word-build-group",
    );
    const text = group?.members?.find(
      ({ testId }) => testId === "word-build-text",
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

  it("Group Text 同目标的 whole 效果数组共享段落构建组", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "group-text-whole-array-build" },
          createElement(
            Group,
            { testId: "whole-array-group", emphasis: spin({ duration: 4000 }) },
            createElement(Text, {
              testId: "whole-array-text",
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
              testId: "whole-array-shape",
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
      ({ testId }) => testId === "whole-array-group",
    );
    const text = group?.members?.find(
      ({ testId }) => testId === "whole-array-text",
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

  it("Group 成员的点击动效仍触发自动步骤中的父 Spin 投影", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Slide,
          { slideId: "group-cross-step-projection" },
          createElement(
            Group,
            {
              testId: "cross-step-group",
              emphasis: repeat(spin({ duration: 8000, from: 360 }), 2),
            },
            createElement(Shape, {
              testId: "cross-step-orange",
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
                testId: "cross-step-green",
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
      ({ testId }) => testId === "cross-step-group",
    );
    const green = group?.members?.find(
      ({ testId }) => testId === "cross-step-green",
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

  it("保留原生 Group 并将有限整体 Spin 投影到全部真实成员", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "group-spin-projection" },
          createElement(
            Group,
            {
              testId: "spinning-group",
              emphasis: repeat(spin({ duration: 1600, from: 360 }), 2),
            },
            createElement(Shape, {
              testId: "breathing-member",
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
              testId: "static-member",
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
      ({ testId }) => testId === "spinning-group",
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
          members.find(({ testId }) => testId === "breathing-member")?.id,
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

  it("保留原有 Fly 与 SpinIn entrance 数组并为段落输出共同中心偏移", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { slideId: "text-motion-array" },
          createElement(Text, {
            testId: "array-text",
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
      ["allAtOnce", "allAtOnce"],
    );
    const spinPositions = spinEffect?.nativeBehaviors?.filter(
      ({ attributeNames }) =>
        attributeNames.some((name) => name === "ppt_x" || name === "ppt_y"),
    );
    assert.deepEqual(
      spinPositions?.map(({ attributeNames, additive, from, to }) => [
        attributeNames,
        additive,
        from?.includes("cos(") && from.includes("sin("),
        to,
      ]),
      [
        [["ppt_x"], "sum", true, "0"],
        [["ppt_y"], "sum", true, "0"],
        [["ppt_x"], "sum", true, "0"],
        [["ppt_y"], "sum", true, "0"],
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

  it("为纯 whole Text entrance 数组增加原生段落构建层但保留逻辑 whole target", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 1280, height: 900 },
        createElement(
          Slide,
          { slideId: "whole-text-motion-array" },
          createElement(Text, {
            testId: "whole-array-text",
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

  it("将八种既有 whole Text 预设逐段写出并独立读回全部原生范围", async () => {
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
            { key: preset, slideId: `whole-${preset}` },
            createElement(Text, {
              testId: `${preset}-text`,
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
          [textObject?.id, expected?.[0], "allAtOnce", 1],
        );
        assert.ok(
          behaviors.every(
            ({ targetType, behavior }) =>
              targetType === "paragraph" ||
              (targetType === "object" && behavior === "set"),
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

  it("whole SpinIn 结束于文字框原有旋转角度", async () => {
    const result = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "rotated-whole-spin-in" },
          createElement(Text, {
            testId: "rotated-text",
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
        ["372", "12", { type: "paragraph", start: 0, end: 0 }],
        ["372", "12", { type: "paragraph", start: 1, end: 1 }],
      ],
    );
  });

  it("拒绝非法效果参数、缺失的 motion 能力和不存在的文件目标", async () => {
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
      [fade({ by: "paragraph" }), "invalid-motion-granularity"],
    ] as const) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { slideId: "invalid" },
              createElement(Shape, {
                testId: "invalid-effect",
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
      createBuilder([shapeCapability, transitionCapability])(
        createElement(
          Deck,
          {},
          createElement(
            Slide,
            { slideId: "trimmed" },
            createElement(Shape, {
              testId: "needs-motion",
              x: 0,
              y: 0,
              width: 16,
              height: 16,
              enter: fade(),
            }),
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
          { slideId: "target" },
          createElement(Shape, {
            testId: "target-object",
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

  it("拒绝超出原生角度范围及与效果预设不匹配的声明", async () => {
    const cases: readonly {
      readonly testId: string;
      readonly props: Record<string, unknown>;
      readonly code: string;
    }[] = [
      {
        testId: "rotation-over-int32",
        props: { emphasis: spin({ from: 40_000 }) },
        code: "invalid-rotation",
      },
      {
        testId: "rotation-too-large",
        props: { emphasis: spin({ from: 1e20 }) },
        code: "invalid-rotation",
      },
      {
        testId: "spin-in-enter",
        props: { enter: { ...spin(), kind: "entrance" } },
        code: "invalid-motion",
      },
      {
        testId: "fade-as-emphasis",
        props: { emphasis: { ...fade(), kind: "emphasis" } },
        code: "invalid-motion",
      },
    ];
    for (const { testId, props, code } of cases) {
      await assert.rejects(
        build(
          createElement(
            Deck,
            {},
            createElement(
              Slide,
              { slideId: "invalid-preset" },
              createElement(Shape, {
                testId,
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

  it("按固定页面槽位写入转场和动画，与模块注册顺序无关", async () => {
    const buildReversed = createBuilder([
      motionCapability,
      transitionCapability,
      shapeCapability,
    ]);
    const result = await buildReversed(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "slot-order", transition: transitions.fade() },
          createElement(Shape, {
            testId: "timed",
            x: 0,
            y: 0,
            width: 16,
            height: 16,
            enter: fade(),
          }),
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
