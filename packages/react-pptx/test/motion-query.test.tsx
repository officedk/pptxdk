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
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
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

function DeclarationClosureDeck() {
  const top = useMotionScope();
  const parent = useMotionScope();
  const empty = useMotionScope();
  const dependent = useMotionScope();
  const deepOne = useMotionScope();
  const deepTwo = useMotionScope();
  const deepThree = useMotionScope();
  const deepLeaf = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "motion-declaration-closure" },
      createElement(
        MotionScope,
        { scope: top, delay: 10 },
        createElement(
          MotionScope,
          { scope: parent, delay: 200 },
          createElement(MotionScope, {
            scope: empty,
            delay: 40,
            publishEnd: empty.cue("ready"),
          }),
          createElement(MotionScope, {
            scope: dependent,
            startAt: after(empty.cue("ready"), 10),
          }),
          createElement(
            Step,
            { n: 1 },
            createElement(Shape, {
              testId: "closure-click",
              x: 0,
              y: 0,
              width: 10,
              height: 10,
              enter: fade({ duration: 100 }),
            }),
          ),
          createElement(
            MotionScope,
            { scope: deepOne, delay: 31 },
            createElement(
              Step,
              { n: 1 },
              createElement(Shape, {
                testId: "closure-deep-one-click",
                x: 15,
                y: 0,
                width: 10,
                height: 10,
                enter: fade({ duration: 100 }),
              }),
            ),
            createElement(
              MotionScope,
              { scope: deepTwo, delay: 32 },
              createElement(
                Step,
                { n: 1 },
                createElement(Shape, {
                  testId: "closure-deep-two-click",
                  x: 30,
                  y: 0,
                  width: 10,
                  height: 10,
                  enter: fade({ duration: 100 }),
                }),
              ),
              createElement(
                MotionScope,
                { scope: deepThree, delay: 33 },
                createElement(
                  Step,
                  { n: 1 },
                  createElement(Shape, {
                    testId: "closure-deep-three-click",
                    x: 45,
                    y: 0,
                    width: 10,
                    height: 10,
                    enter: fade({ duration: 100 }),
                  }),
                ),
                createElement(MotionScope, {
                  scope: deepLeaf,
                  delay: 34,
                }),
              ),
            ),
          ),
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

function EmptyIntervalDeck({
  sourceDuration,
}: {
  readonly sourceDuration?: number;
}) {
  const source = useMotionScope();
  const empty = useMotionScope();
  const finite = useMotionScope();
  const overlapping = useMotionScope();
  const infinite = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "empty-interval" },
      sourceDuration === undefined
        ? createElement(
            MotionScope,
            { scope: source },
            createElement(
              Text,
              {
                testId: "unknown-source",
                x: 0,
                y: 0,
                width: 100,
                height: 20,
                enter: fade({ by: "word", duration: 100 }),
              },
              "first second",
            ),
          )
        : createElement(
            MotionScope,
            { scope: source },
            createElement(Shape, {
              testId: "known-source",
              x: 0,
              y: 0,
              width: 10,
              height: 10,
              fill: "#123456",
              enter: fade({ duration: sourceDuration }),
            }),
          ),
      createElement(MotionScope, {
        scope: empty,
        startAt: after(source.end, 50),
      }),
      createElement(
        MotionScope,
        { scope: finite, startAt: source.end },
        createElement(Shape, {
          testId: "finite-interval",
          x: 20,
          y: 0,
          width: 10,
          height: 10,
          fill: "#654321",
          enter: fade({ duration: 100 }),
        }),
      ),
      createElement(
        MotionScope,
        { scope: infinite, startAt: source.end },
        createElement(Shape, {
          testId: "unbounded-interval",
          x: 40,
          y: 0,
          width: 10,
          height: 10,
          fill: "#abcdef",
          emphasis: repeat(
            breathe({ period: 200, from: "#123456", to: "#654321" }),
            "forever",
          ),
        }),
      ),
      createElement(
        MotionScope,
        { scope: overlapping, startAt: after(source.end, 50) },
        createElement(Shape, {
          testId: "overlapping-interval",
          x: 60,
          y: 0,
          width: 10,
          height: 10,
          fill: "#fedcba",
          enter: fade({ duration: 40 }),
        }),
      ),
    ),
  );
}

function DeclaredScopeClosureDeck() {
  const parent = useMotionScope();
  const empty = useMotionScope();
  const dependent = useMotionScope();
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "scope-closure" },
      createElement(
        MotionScope,
        { scope: parent, delay: 200 },
        createElement(MotionScope, { scope: empty, delay: 40 }),
        createElement(MotionScope, {
          scope: dependent,
          startAt: after(empty.end, 10),
        }),
        createElement(
          Step,
          { n: 1 },
          createElement(Shape, {
            testId: "click-member",
            x: 0,
            y: 0,
            width: 10,
            height: 10,
            fill: "#123456",
            enter: fade({ duration: 100 }),
          }),
        ),
      ),
    ),
  );
}

function WordMotionDeck() {
  return createElement(
    Deck,
    { width: 120, height: 100 },
    createElement(
      Slide,
      { slideId: "word-projection" },
      createElement(
        Text,
        {
          testId: "word-text",
          x: 0,
          y: 0,
          width: 100,
          height: 20,
          enter: fade({
            by: "word",
            duration: 100,
            easing: { accelerate: 0.2 },
          }),
        },
        "first second",
      ),
    ),
  );
}

describe("observeMotion", () => {
  it("把已证明的空半开区间判为不重叠", async () => {
    for (const sourceDuration of [100, 800, 3000]) {
      const result = await build(
        createElement(EmptyIntervalDeck, { sourceDuration }),
      );
      const step = observeMotion(
        result.query.slide("empty-interval").observation,
      )?.steps[0];
      assert.ok(step);
      const empty = step.scopes.find(({ members }) => members.length === 0);
      const finite = step.effects.find(
        ({ target }) => target.testId === "finite-interval",
      );
      const overlapping = step.effects.find(
        ({ target }) => target.testId === "overlapping-interval",
      );
      assert.ok(empty);
      assert.ok(finite);
      assert.ok(overlapping);
      assert.deepEqual(step.difference(empty.start.ref, empty.end.ref), {
        state: "known",
        ms: 0,
      });
      assert.deepEqual(step.overlap(empty.ref, finite.ref), {
        state: "known",
        value: false,
      });
      assert.deepEqual(step.overlap(finite.ref, overlapping.ref), {
        state: "known",
        value: true,
      });
    }

    const unknownResult = await build(createElement(EmptyIntervalDeck));
    const unknownStep = observeMotion(
      unknownResult.query.slide("empty-interval").observation,
    )?.steps[0];
    assert.ok(unknownStep);
    const empty = unknownStep.scopes.find(
      ({ members }) => members.length === 0,
    );
    const finite = unknownStep.effects.find(
      ({ target }) => target.testId === "finite-interval",
    );
    const overlapping = unknownStep.effects.find(
      ({ target }) => target.testId === "overlapping-interval",
    );
    const unbounded = unknownStep.effects.find(
      ({ target }) => target.testId === "unbounded-interval",
    );
    assert.ok(empty);
    assert.ok(finite);
    assert.ok(overlapping);
    assert.ok(unbounded);
    assert.deepEqual(unknownStep.difference(empty.start.ref, empty.end.ref), {
      state: "known",
      ms: 0,
    });
    assert.deepEqual(unknownStep.overlap(empty.ref, finite.ref), {
      state: "known",
      value: false,
    });
    assert.deepEqual(unknownStep.overlap(finite.ref, overlapping.ref), {
      state: "known",
      value: true,
    });
    assert.deepEqual(unknownStep.overlap(empty.ref, unbounded.ref), {
      state: "known",
      value: false,
    });
  });

  it("按声明闭包保留多层作用域关系且不扩展活动时间", async () => {
    const result = await build(createElement(DeclarationClosureDeck));
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const motion = observeMotion(
      result.query.slide("motion-declaration-closure").observation,
    );
    const automatic = motion?.steps.find(({ trigger }) => trigger === "auto");
    const click = motion?.steps.find(
      ({ trigger, number }) => trigger === "click" && number === 1,
    );
    assert.ok(automatic);
    assert.ok(click);
    assert.equal(automatic.scopes.length, 8);
    assert.equal(automatic.effects.length, 0);

    const automaticByDelay = new Map(
      automatic.scopes.map((scope) => [scope.delayMs, scope] as const),
    );
    const top = automaticByDelay.get(10);
    const parent = automaticByDelay.get(200);
    const empty = automaticByDelay.get(40);
    const dependent = automaticByDelay.get(0);
    const deepOne = automaticByDelay.get(31);
    const deepTwo = automaticByDelay.get(32);
    const deepThree = automaticByDelay.get(33);
    const deepLeaf = automaticByDelay.get(34);
    const closure = [
      top,
      parent,
      empty,
      dependent,
      deepOne,
      deepTwo,
      deepThree,
      deepLeaf,
    ];
    assert.ok(closure.every((scope) => scope !== undefined));
    assert.deepEqual(
      [
        top,
        parent,
        empty,
        dependent,
        deepOne,
        deepTwo,
        deepThree,
        deepLeaf,
      ].map((scope) => scope?.delayMs),
      [10, 200, 40, 0, 31, 32, 33, 34],
    );

    const relationships = [
      [top!, parent!],
      [parent!, empty!],
      [parent!, dependent!],
      [parent!, deepOne!],
      [deepOne!, deepTwo!],
      [deepTwo!, deepThree!],
      [deepThree!, deepLeaf!],
    ] as const;
    for (const [ancestor, child] of relationships) {
      assert.equal(child.parent, ancestor.ref);
      assert.ok(
        ancestor.members.some(
          (member) => member.kind === "scope" && member.ref === child.ref,
        ),
      );
    }

    assert.equal(top!.step, automatic.ref);
    assert.equal(parent!.start.state, "known");
    assert.equal(parent!.start.ms, 210);
    assert.equal(parent!.end.state, "known");
    assert.equal(parent!.end.ms, 210);
    assert.equal(empty!.end.state, "known");
    assert.equal(empty!.end.ms, 250);
    assert.equal(dependent!.start.state, "known");
    assert.equal(dependent!.start.ms, 260);
    assert.ok(
      automatic.cues.some(
        ({ name, publishedBy, point }) =>
          name === "ready" && publishedBy === empty!.ref && point === "end",
      ),
    );
    assert.ok(
      automatic.dependencies.some(
        (dependency) =>
          dependency.kind === "cue-publish" &&
          dependency.scope === empty!.ref &&
          dependency.point === "end",
      ),
    );
    assert.ok(
      automatic.dependencies.some(
        (dependency) =>
          dependency.kind === "cue-use" &&
          dependency.scope === dependent!.ref &&
          dependency.cue.scopeId === empty!.ref.id &&
          dependency.offsetMs === 10,
      ),
    );

    const clickTop = click.scopes.find(({ delayMs }) => delayMs === 10);
    const clickParent = click.scopes.find(({ delayMs }) => delayMs === 200);
    const clickDeepOne = click.scopes.find(({ delayMs }) => delayMs === 31);
    const clickEffect = click.effects.find(
      ({ target }) => target.testId === "closure-click",
    );
    assert.ok(clickTop);
    assert.ok(clickParent);
    assert.ok(clickDeepOne);
    assert.ok(clickEffect);
    assert.equal(clickTop.step, click.ref);
    assert.equal(clickTop.parent, undefined);
    assert.equal(clickParent.parent, clickTop.ref);
    assert.equal(clickDeepOne.parent, clickParent.ref);
    assert.equal(clickParent.members.length, 2);
    assert.ok(
      clickParent.members.some(
        (member) => member.kind === "scope" && member.ref === clickDeepOne.ref,
      ),
    );
    assert.ok(
      clickParent.members.some(
        (member) => member.kind === "effect" && member.ref === clickEffect.ref,
      ),
    );
    assert.equal(clickEffect.start.state, "known");
    assert.equal(clickEffect.start.ms, 210);
    assert.equal(clickEffect.end.state, "known");
    assert.equal(clickEffect.end.ms, 310);
    assert.notEqual(top!.ref, clickTop.ref);
    assert.deepEqual(automatic.difference(top!.start.ref, clickTop.start.ref), {
      state: "unknown",
      reason: "different-step",
    });
    assert.equal(
      click.scopes.some(({ delayMs }) => delayMs === 40 || delayMs === 0),
      false,
    );
  });

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

  it("保留跨触发步声明闭包并区分结构父子与本步排程", async () => {
    const result = await build(createElement(DeclaredScopeClosureDeck));
    const motion = observeMotion(
      result.query.slide("scope-closure").observation,
    );
    const automatic = motion?.steps.find(({ trigger }) => trigger === "auto");
    const click = motion?.steps.find(
      ({ trigger, number }) => trigger === "click" && number === 1,
    );
    assert.ok(automatic);
    assert.ok(click);
    const automaticParent = automatic.scopes.find(
      ({ delayMs, parent }) => delayMs === 200 && parent === undefined,
    );
    const empty = automatic.scopes.find(
      ({ delayMs, members, parent }) =>
        delayMs === 40 && members.length === 0 && parent !== undefined,
    );
    const dependent = automatic.scopes.find(
      ({ delayMs, members, parent }) =>
        delayMs === 0 && members.length === 0 && parent !== undefined,
    );
    const clickParent = click.scopes.find(
      ({ delayMs, parent }) => delayMs === 200 && parent === undefined,
    );
    assert.ok(automaticParent);
    assert.ok(empty);
    assert.ok(dependent);
    assert.ok(clickParent);
    assert.equal(empty.parent, automaticParent.ref);
    assert.equal(dependent.parent, automaticParent.ref);
    assert.deepEqual(
      automaticParent.members.map(({ ref }) => ref),
      [empty.ref, dependent.ref],
    );
    assert.equal(automaticParent.step, automatic.ref);
    assert.equal(clickParent.step, click.ref);
    assert.notEqual(automaticParent.ref, clickParent.ref);
    assert.deepEqual(
      clickParent.members.map(({ kind }) => kind),
      ["effect"],
    );
    assert.equal(automatic.effects.length, 0);
    assert.equal(click.effects.length, 1);
    assert.deepEqual(
      [automaticParent.start, empty.start, dependent.start].map(
        ({ state, ...time }) => [state, "ms" in time ? time.ms : undefined],
      ),
      [
        ["known", 200],
        ["known", 240],
        ["known", 250],
      ],
    );
    const reportMotion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const reportParent = reportMotion.scopes?.find(
      ({ id }) => id === automaticParent.ref.id,
    );
    assert.deepEqual(
      reportParent?.members.map(({ kind }) => kind),
      ["object"],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(
      file.slides[0]?.animations?.map(({ trigger, effects }) => [
        trigger,
        effects.map(
          ({ targetId }) =>
            result.report.slides[0]?.objects.find(({ id }) => id === targetId)
              ?.testId,
        ),
      ]),
      [["click", ["click-member"]]],
    );
    assert.ok(
      automatic.dependencies.some(
        (dependency) =>
          dependency.kind === "scope-start" &&
          dependency.scope === dependent.ref &&
          dependency.offsetMs === 10,
      ),
    );
  });

  it("只公开效果观察字段并冻结其嵌套值", async () => {
    const result = await build(createElement(WordMotionDeck));
    const step = observeMotion(
      result.query.slide("word-projection").observation,
    )?.steps[0];
    assert.ok(step);
    const effect = step.effects.find(
      ({ target }) => target.testId === "word-text",
    );
    assert.ok(effect);
    assert.equal(effect.by, "word");
    assert.equal(effect.durationMs, 100);
    assert.equal(effect.intervalMs, 10);
    for (const privateField of [
      "stepKey",
      "order",
      "targetTextBody",
      "targetFilled",
      "iterate",
    ]) {
      assert.equal(Object.hasOwn(effect, privateField), false);
    }
    assert.ok(Object.isFrozen(effect));
    assert.ok(Object.isFrozen(effect.target));
    assert.ok(effect.easing);
    assert.ok(Object.isFrozen(effect.easing));
    assert.throws(() => {
      (effect.target as { testId?: string }).testId = "changed";
    }, TypeError);
    assert.throws(() => {
      (effect.easing as { accelerate?: number }).accelerate = 0.8;
    }, TypeError);
    assert.equal(effect.target.testId, "word-text");
    assert.equal(effect.easing.accelerate, 0.2);
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
