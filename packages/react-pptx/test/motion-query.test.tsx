import assert from "node:assert/strict";

import {
  after,
  breathe,
  build,
  BuildQueryError,
  Deck,
  fade,
  MotionScope,
  P,
  repeat,
  Shape,
  Slide,
  Step,
  Text,
  useMotionScope,
  observeMotion,
} from "@react-pptx/react-pptx";
import { createElement, Fragment } from "react";
import { describe, it } from "vitest";

function KnownRelationsDeck() {
  const source = useMotionScope();
  const paragraphs = useMotionScope();
  const touching = useMotionScope();
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "motion-relations" },
      createElement(
        Fragment,
        null,
        createElement(
          MotionScope,
          { scope: source, publishEnd: source.cue("ready") },
          createElement(Shape, {
            testId: "source",
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            fill: "#123456",
            enter: fade({ duration: 100 }),
          }),
        ),
        createElement(
          MotionScope,
          { scope: paragraphs, mode: "stagger", each: 25 },
          createElement(
            Text,
            {
              testId: "paragraph-parts",
              x: 0,
              y: 20,
              width: 100,
              height: 30,
              enter: fade({ by: "paragraph", duration: 100 }),
            },
            createElement(P, null, "first"),
            createElement(P, null, "second"),
          ),
        ),
        createElement(
          MotionScope,
          { scope: touching, startAt: source.end },
          createElement(Shape, {
            testId: "touching",
            x: 40,
            y: 0,
            width: 10,
            height: 10,
            fill: "#abcdef",
            enter: fade({ duration: 10 }),
          }),
        ),
        createElement(
          MotionScope,
          {
            scope: follower,
            startAt: after(source.cue("ready"), 20),
          },
          createElement(Shape, {
            testId: "follower",
            x: 20,
            y: 0,
            width: 10,
            height: 10,
            fill: "#654321",
            enter: [fade({ duration: 50 }), fade({ duration: 30 })],
          }),
        ),
      ),
    ),
  );
}

function MixedIntervalsDeck() {
  const words = useMotionScope();
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "motion-coverage" },
      createElement(
        Fragment,
        null,
        createElement(
          MotionScope,
          { scope: words },
          createElement(
            Text,
            {
              testId: "unknown-words",
              x: 0,
              y: 0,
              width: 100,
              height: 20,
              enter: fade({ by: "word", duration: 100 }),
            },
            "first second",
          ),
        ),
        createElement(
          MotionScope,
          { scope: follower, startAt: words.end },
          createElement(Shape, {
            testId: "unknown-start-finite-end",
            x: 40,
            y: 0,
            width: 10,
            height: 10,
            fill: "#123456",
            enter: fade({ duration: 80 }),
          }),
        ),
        createElement(Shape, {
          testId: "known-finite",
          x: 0,
          y: 30,
          width: 10,
          height: 10,
          fill: "#123456",
          enter: fade({ duration: 100 }),
        }),
        createElement(Shape, {
          testId: "no-natural-end",
          x: 20,
          y: 30,
          width: 10,
          height: 10,
          fill: "#654321",
          emphasis: repeat(
            breathe({ period: 200, from: "#123456", to: "#654321" }),
            "forever",
          ),
        }),
      ),
    ),
  );
}

function ClickStepsDeck() {
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "motion-steps" },
      createElement(
        Step,
        { n: 1 },
        createElement(Shape, {
          testId: "first-click",
          x: 0,
          y: 0,
          width: 10,
          height: 10,
          fill: "#123456",
          enter: fade({ duration: 100 }),
        }),
      ),
      createElement(
        Step,
        { n: 2 },
        createElement(Shape, {
          testId: "second-click",
          x: 20,
          y: 0,
          width: 10,
          height: 10,
          fill: "#654321",
          enter: fade({ duration: 100 }),
        }),
      ),
    ),
  );
}

describe("observeMotion", () => {
  it("保留效果实例、具名时刻声明并查询精确关系", async () => {
    const result = await build(createElement(KnownRelationsDeck));
    const slide = result.query.slide("motion-relations");
    const motion = observeMotion(slide.observation);

    assert.equal(motion?.evidence, "generation");
    assert.equal(motion?.steps.length, 1);
    const step = motion?.steps[0];
    assert.ok(step);
    const source = step.effects.find(
      ({ target }) => target.testId === "source",
    );
    const followers = step.effects.filter(
      ({ target }) => target.testId === "follower",
    );
    const touching = step.effects.find(
      ({ target }) => target.testId === "touching",
    );
    assert.ok(source);
    assert.ok(touching);
    assert.equal(followers.length, 2);
    assert.deepEqual(
      followers.map(({ effectIndex }) => effectIndex),
      [0, 1],
    );
    const paragraphParts = step.effects.filter(
      ({ target }) => target.testId === "paragraph-parts",
    );
    assert.deepEqual(
      followers.map(({ start }) => start.state),
      ["known", "known"],
    );
    assert.deepEqual(
      followers.map(({ end }) => end.state),
      ["known", "known"],
    );
    assert.ok(followers[0]?.ref.effectIndex === 0);
    assert.deepEqual(
      paragraphParts.map(({ partIndex, ref }) => [partIndex, ref.partIndex]),
      [
        [0, 0],
        [1, 1],
      ],
    );
    const paragraphScope = step.scopes.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "effect" &&
          member.ref.targetId === paragraphParts[0]?.ref.targetId,
      ),
    );
    assert.ok(paragraphScope);
    assert.ok(paragraphParts[0]?.scopes.includes(paragraphScope.ref));
    assert.equal(
      step.difference(source.end.ref, followers[0]!.start.ref).state,
      "known",
    );
    assert.deepEqual(step.difference(source.end.ref, followers[0]!.start.ref), {
      state: "known",
      ms: 20,
    });
    assert.deepEqual(step.before(source.end.ref, followers[0]!.start.ref), {
      state: "known",
      value: true,
    });
    assert.deepEqual(step.overlap(source.ref, followers[0]!.ref), {
      state: "known",
      value: false,
    });
    assert.deepEqual(step.overlap(source.ref, touching.ref), {
      state: "known",
      value: false,
    });
    assert.deepEqual(step.intervalsCovering(100, [source.ref, touching.ref]), {
      complete: true,
      items: [
        { effect: source.ref, result: { state: "known", value: false } },
        { effect: touching.ref, result: { state: "known", value: true } },
      ],
    });
    assert.equal(step.cues.length, 1);
    assert.equal(step.cues[0]?.name, "ready");
    assert.ok(
      step.dependencies.some(
        (dependency) =>
          dependency.kind === "cue-use" && dependency.offsetMs === 20,
      ),
    );
    assert.ok(Object.isFrozen(motion));
    assert.ok(Object.isFrozen(step.effects));
    assert.ok(Object.isFrozen(paragraphScope.members));
    assert.ok(Object.isFrozen(paragraphScope.members[0]));
  });

  it("覆盖结果保留未知完成和无自然完成项目", async () => {
    const result = await build(createElement(MixedIntervalsDeck));
    const slide = result.query.slide("motion-coverage");
    const motion = observeMotion(slide.observation);
    const step = motion?.steps[0];
    assert.ok(step);
    assert.deepEqual(step.effects.map(({ end }) => end.state).sort(), [
      "known",
      "no-natural-completion",
      "unknown",
      "unknown",
    ]);

    const covering = step.intervalsCovering(50);
    assert.equal(covering.items.length, 4);
    assert.equal(covering.complete, false);
    assert.equal(
      covering.items.filter(({ result: item }) => item.state === "unknown")
        .length,
      2,
    );
    assert.equal(
      covering.items.filter(
        ({ result: item }) => item.state === "no-natural-completion",
      ).length,
      1,
    );
    const finite = step.effects.find(
      ({ target }) => target.testId === "known-finite",
    );
    assert.ok(finite);
    const wordEffect = step.effects.find(
      ({ target }) => target.testId === "unknown-words",
    );
    const follower = step.effects.find(
      ({ target }) => target.testId === "unknown-start-finite-end",
    );
    assert.ok(wordEffect);
    assert.ok(follower);
    assert.deepEqual(step.difference(follower.start.ref, follower.end.ref), {
      state: "known",
      ms: 80,
    });
    assert.deepEqual(step.before(wordEffect.start.ref, wordEffect.end.ref), {
      state: "known",
      value: true,
    });
    assert.deepEqual(step.intervalsCovering(100, [finite.ref]), {
      complete: true,
      items: [{ effect: finite.ref, result: { state: "known", value: false } }],
    });
  });

  it("跨步返回时钟不同；跨构建引用抛查询错误", async () => {
    const first = await build(createElement(ClickStepsDeck));
    const second = await build(createElement(ClickStepsDeck));
    const firstMotion = observeMotion(
      first.query.slide("motion-steps").observation,
    );
    const secondMotion = observeMotion(
      second.query.slide("motion-steps").observation,
    );
    assert.equal(firstMotion?.steps.length, 2);
    assert.equal(secondMotion?.steps.length, 2);
    const firstStep = firstMotion?.steps[0];
    const secondStep = firstMotion?.steps[1];
    assert.ok(firstStep);
    assert.ok(secondStep);
    const firstMoment = firstStep.effects[0]?.start.ref;
    const secondMoment = secondStep.effects[0]?.start.ref;
    const foreignMoment = secondMotion?.steps[0]?.effects[0]?.start.ref;
    assert.ok(firstMoment);
    assert.ok(secondMoment);
    assert.ok(foreignMoment);
    assert.deepEqual(firstStep.difference(firstMoment, secondMoment), {
      state: "unknown",
      reason: "different-step",
    });
    const secondEffect = secondStep.effects[0]?.ref;
    assert.ok(secondEffect);
    assert.deepEqual(firstStep.intervalsCovering(50, [secondEffect]), {
      complete: false,
      items: [
        {
          effect: secondEffect,
          result: { state: "unknown", reason: "different-step" },
        },
      ],
    });
    assert.throws(
      () => firstStep.difference(firstMoment, foreignMoment),
      (error: unknown) =>
        error instanceof BuildQueryError && error.code === "foreign-reference",
    );
    assert.equal(
      observeMotion(first.query.slide("motion-steps").observation),
      firstMotion,
    );
    assert.throws(
      () =>
        observeMotion({
          ...first.query.slide("motion-steps").observation,
          capabilities: {
            ...first.query.slide("motion-steps").observation.capabilities,
            motion: {
              ...(first.query.slide("motion-steps").observation.capabilities?.[
                "motion"
              ] as object),
            },
          },
        }),
      (error: unknown) =>
        error instanceof BuildQueryError &&
        error.code === "unbound-observation",
    );
  });
});
