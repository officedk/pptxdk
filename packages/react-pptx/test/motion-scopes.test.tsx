import assert from "node:assert/strict";

import {
  after,
  Box,
  build,
  BuildError,
  Deck,
  fade,
  Group,
  MotionScope,
  Shape,
  Slide,
  Step,
  useMotionScope,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { createElement, Fragment } from "react";
import { describe, it } from "vitest";

function movingShape(testId: string, duration: number) {
  return createElement(Shape, {
    testId,
    x: 0,
    y: 0,
    width: 20,
    height: 20,
    enter: fade({ duration }),
  });
}

function StaggeredSource() {
  const source = useMotionScope();
  const namedCompletion = source.cue("finished");
  const directFollower = useMotionScope();
  const namedFollower = useMotionScope();

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "scope-order" },
      createElement(
        MotionScope,
        { scope: directFollower, startAt: after(source.end, 50) },
        movingShape("after-end", 300),
      ),
      createElement(
        MotionScope,
        { scope: namedFollower, startAt: after(namedCompletion, 25) },
        movingShape("after-named-end", 300),
      ),
      createElement(
        MotionScope,
        {
          scope: source,
          mode: "stagger",
          each: 100,
          publishEnd: namedCompletion,
        },
        movingShape("long-first", 1000),
        movingShape("medium-second", 800),
        movingShape("short-last", 50),
      ),
    ),
  );
}

function EmptyScopeDeck({
  includeSource,
  sourceMode = "empty",
}: {
  readonly includeSource: boolean;
  readonly sourceMode?: "empty" | "static" | "disabled";
}) {
  const empty = useMotionScope();
  const emptyStart = empty.cue("empty-start");
  const emptyEnd = empty.cue("empty-end");
  const fromStart = useMotionScope();
  const fromEnd = useMotionScope();
  const fromNamedStart = useMotionScope();
  const fromNamedEnd = useMotionScope();
  const sourceChildren =
    sourceMode === "empty"
      ? null
      : sourceMode === "static"
        ? createElement(Shape, {
            testId: "static-source",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
          })
        : movingShape("disabled-source", 100);

  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-scope" },
      includeSource
        ? createElement(
            MotionScope,
            {
              scope: empty,
              publishStart: emptyStart,
              publishEnd: emptyEnd,
              motion: sourceMode !== "disabled",
            },
            sourceChildren,
          )
        : null,
      createElement(
        MotionScope,
        { scope: fromStart, startAt: after(empty.start, 50) },
        movingShape("from-start", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromEnd, startAt: after(empty.end, 80) },
        movingShape("from-end", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromNamedStart, startAt: after(emptyStart, 30) },
        movingShape("from-named-start", 100),
      ),
      createElement(
        MotionScope,
        { scope: fromNamedEnd, startAt: after(emptyEnd, 40) },
        movingShape("from-named-end", 100),
      ),
    ),
  );
}

function CyclicScopes() {
  const first = useMotionScope();
  const second = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "cycle" },
      createElement(
        MotionScope,
        { scope: first, startAt: second.end },
        movingShape("first", 100),
      ),
      createElement(
        MotionScope,
        { scope: second, startAt: first.end },
        movingShape("second", 100),
      ),
    ),
  );
}

function CrossStepScopes() {
  const first = useMotionScope();
  const second = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "cross-step" },
      createElement(
        Step,
        { n: 1 },
        createElement(
          MotionScope,
          { scope: first },
          movingShape("first-step", 100),
        ),
      ),
      createElement(
        Step,
        { n: 2 },
        createElement(
          MotionScope,
          { scope: second, startAt: first.end },
          movingShape("second-step", 100),
        ),
      ),
    ),
  );
}

function SequencedScopes() {
  const sequence = useMotionScope();
  const parallel = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "sequence" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        movingShape("first", 200),
        createElement(
          MotionScope,
          { scope: parallel },
          movingShape("parallel-short", 100),
          movingShape("parallel-long", 700),
        ),
        movingShape("last", 100),
      ),
    ),
  );
}

function EmptySequenceMember() {
  const sequence = useMotionScope();
  const empty = useMotionScope();
  const emptyCompletion = empty.cue("done");
  const follower = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-sequence-member" },
      createElement(
        MotionScope,
        { scope: sequence, mode: "sequence" },
        movingShape("before-empty", 200),
        createElement(
          MotionScope,
          { scope: empty, publishEnd: emptyCompletion },
          null,
        ),
        movingShape("after-empty", 100),
      ),
      createElement(
        MotionScope,
        { scope: follower, startAt: after(emptyCompletion, 10) },
        movingShape("after-empty-completion", 100),
      ),
    ),
  );
}

function EmptyStaggerMember() {
  const stagger = useMotionScope();
  const empty = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "empty-stagger-member" },
      createElement(
        MotionScope,
        { scope: stagger, mode: "stagger", each: 100 },
        movingShape("stagger-before-empty", 100),
        createElement(MotionScope, { scope: empty }, null),
        movingShape("stagger-after-empty", 100),
      ),
    ),
  );
}

function SelfStartedMember() {
  const group = useMotionScope();
  const leader = useMotionScope();
  const later = useMotionScope();
  const afterGroup = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "self-started-member" },
      createElement(
        MotionScope,
        { scope: group },
        createElement(
          MotionScope,
          { scope: leader },
          movingShape("leader", 100),
        ),
        createElement(
          MotionScope,
          { scope: later, startAt: after(leader.end, 20) },
          movingShape("self-started", 50),
        ),
      ),
      createElement(
        MotionScope,
        { scope: afterGroup, startAt: after(group.end, 10) },
        movingShape("after-group", 50),
      ),
    ),
  );
}

function BoxAndDisabledMembers() {
  const stagger = useMotionScope();
  return createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { slideId: "box-motion" },
      createElement(
        Box,
        { enter: fade({ delay: 800 }) },
        createElement(Shape, {
          testId: "box-default",
          x: 0,
          y: 0,
          width: 20,
          height: 20,
        }),
        createElement(Shape, {
          testId: "own-entry",
          x: 24,
          y: 0,
          width: 20,
          height: 20,
          enter: fade({ duration: 200 }),
        }),
      ),
      createElement(
        MotionScope,
        { scope: stagger, mode: "stagger", each: 100 },
        movingShape("stagger-first", 100),
        createElement(Shape, {
          testId: "static",
          x: 48,
          y: 0,
          width: 20,
          height: 20,
        }),
        createElement(
          Box,
          { motion: false },
          movingShape("disabled-descendant", 100),
        ),
        movingShape("stagger-second", 100),
      ),
    ),
  );
}

function ReusedNamedScope() {
  const source = useMotionScope();
  const done = source.cue("done");
  const follower = useMotionScope();
  return createElement(
    Fragment,
    null,
    createElement(
      MotionScope,
      { scope: source, publishEnd: done },
      movingShape("source", 100),
    ),
    createElement(
      MotionScope,
      { scope: follower, startAt: after(done, 10) },
      movingShape("follower", 100),
    ),
  );
}

function OrdinaryComposition() {
  return createElement(
    Fragment,
    null,
    movingShape("function-first", 100),
    createElement(
      Group,
      { testId: "native-group" },
      movingShape("group-first", 100),
      movingShape("group-second", 100),
    ),
  );
}

describe("组件编排作用域", () => {
  it("等待交错作用域的全部成员，并允许前向引用完成与命名时刻", async () => {
    const result = await build(createElement(StaggeredSource));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    const times = motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
      target.testId,
      startMs,
      endMs,
    ]);

    assert.deepEqual(times, [
      ["long-first", 0, 1000],
      ["medium-second", 100, 900],
      ["short-last", 200, 250],
      ["after-named-end", 1025, 1325],
      ["after-end", 1050, 1350],
    ]);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["long-first", 0],
        ["medium-second", 100],
        ["short-last", 200],
        ["after-named-end", 1025],
        ["after-end", 1050],
      ],
    );
  });

  it("保留空作用域及其边界时刻，并把整段条件移除报告为缺失引用", async () => {
    for (const sourceMode of ["empty", "static", "disabled"] as const) {
      const result = await build(
        createElement(EmptyScopeDeck, { includeSource: true, sourceMode }),
      );
      const slide = result.report.slides[0];
      const motion = slide?.capabilities?.["motion"] as MotionObservation;
      assert.deepEqual(motion.scopes?.[0], {
        id: "scope-1",
        mode: "parallel",
        delayMs: 0,
        startMs: 0,
        endMs: 0,
        completion: "known",
        members: [],
      });
      assert.deepEqual(
        motion.steps[0]?.effects.map(({ target, startMs }) => [
          target.testId,
          startMs,
        ]),
        [
          ["from-named-start", 30],
          ["from-named-end", 40],
          ["from-start", 50],
          ["from-end", 80],
        ],
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(slide?.objects.length, sourceMode === "empty" ? 4 : 5);
      assert.equal(
        file.slides[0]?.objects.length,
        sourceMode === "empty" ? 4 : 5,
      );
    }

    await assert.rejects(
      build(createElement(EmptyScopeDeck, { includeSource: false })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-cue");
        return true;
      },
    );
  });

  it("拒绝依赖环和明确跨点击步的作用域起止关系", async () => {
    await assert.rejects(
      build(createElement(CyclicScopes)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "motion-cycle");
        return true;
      },
    );

    await assert.rejects(
      build(createElement(CrossStepScopes)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "cross-step-motion");
        return true;
      },
    );
  });

  it("顺序作用域等待嵌套并行作用域的全部成员", async () => {
    const result = await build(createElement(SequencedScopes));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["first", 0, 200],
        ["parallel-short", 200, 300],
        ["parallel-long", 200, 900],
        ["last", 900, 1000],
      ],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["first", 0],
        ["parallel-short", 200],
        ["parallel-long", 200],
        ["last", 900],
      ],
    );
  });

  it("空子作用域保留顺序边界，但不占交错位置", async () => {
    const sequenceResult = await build(createElement(EmptySequenceMember));
    const sequenceMotion = sequenceResult.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      sequenceMotion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["before-empty", 0, 200],
        ["after-empty", 200, 300],
        ["after-empty-completion", 210, 310],
      ],
    );
    assert.deepEqual(
      sequenceMotion.scopes?.find(({ id }) => id === "scope-2"),
      {
        id: "scope-2",
        parentId: "scope-1",
        mode: "parallel",
        delayMs: 0,
        startMs: 200,
        endMs: 200,
        completion: "known",
        members: [],
      },
    );
    assert.deepEqual((await validatePptx(sequenceResult.pptx)).errors, []);

    const staggerResult = await build(createElement(EmptyStaggerMember));
    const staggerMotion = staggerResult.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      staggerMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["stagger-before-empty", 0],
        ["stagger-after-empty", 100],
      ],
    );
    assert.deepEqual((await validatePptx(staggerResult.pptx)).errors, []);
  });

  it("自行指定起点的并行成员仍计入父作用域整体完成", async () => {
    const result = await build(createElement(SelfStartedMember));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, endMs }) => [
        target.testId,
        startMs,
        endMs,
      ]),
      [
        ["leader", 0, 100],
        ["self-started", 120, 170],
        ["after-group", 180, 230],
      ],
    );
    assert.equal(motion.scopes?.[0]?.endMs, 170);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
  });

  it("Box 默认入场只补给没有自带入场的对象，静态和禁用对象不占交错位置", async () => {
    const result = await build(createElement(BoxAndDisabledMembers));
    const slide = result.report.slides[0];
    const motion = slide?.capabilities?.["motion"] as MotionObservation;
    assert.deepEqual(
      motion.steps[0]?.effects.map(({ target, startMs, delayMs }) => [
        target.testId,
        startMs,
        delayMs,
      ]),
      [
        ["own-entry", 0, 0],
        ["stagger-first", 0, 0],
        ["stagger-second", 100, 0],
        ["box-default", 800, 800],
      ],
    );
    const stagger = motion.scopes?.find(({ mode }) => mode === "stagger");
    assert.deepEqual(
      stagger?.members.map((member) =>
        member.kind === "object"
          ? slide?.objects.find(({ id }) => id === member.id)?.testId
          : member.kind,
      ),
      ["stagger-first", "stagger-second"],
    );
    assert.equal(
      slide?.objects.some(({ testId }) => testId === "disabled-descendant"),
      true,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(({ targetId, delayMs }) => [
        slide?.objects.find(({ id }) => id === targetId)?.testId,
        delayMs,
      ]),
      [
        ["own-entry", 0],
        ["stagger-first", 0],
        ["stagger-second", 100],
        ["box-default", 800],
      ],
    );
  });

  it("同名组件实例互相隔离，普通函数、Fragment 和原生 Group 不自动成作用域", async () => {
    const reused = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "reused" },
          createElement(ReusedNamedScope),
          createElement(ReusedNamedScope),
        ),
      ),
    );
    const reusedMotion = reused.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.deepEqual(
      reusedMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["source", 0],
        ["source", 0],
        ["follower", 110],
        ["follower", 110],
      ],
    );
    assert.equal(reusedMotion.scopes?.length, 4);

    const ordinary = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { slideId: "ordinary" },
          createElement(OrdinaryComposition),
        ),
      ),
    );
    const ordinaryMotion = ordinary.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.equal(ordinaryMotion.scopes, undefined);
    assert.deepEqual(
      ordinaryMotion.steps[0]?.effects.map(({ target, startMs }) => [
        target.testId,
        startMs,
      ]),
      [
        ["function-first", 0],
        ["group-first", 0],
        ["group-second", 0],
      ],
    );
  });
});
