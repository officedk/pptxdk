import assert from "node:assert/strict";

import * as jsx from "@react-pptx/react-pptx";
import {
  after,
  box,
  build,
  BuildError,
  createMorphScope,
  createMotionScopeHandle,
  deck,
  fade,
  motionScope,
  observeMotion,
  shape,
  slide,
  step,
  transitions,
  type MotionObservation,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

function motionOf(
  result: Awaited<ReturnType<typeof build>>,
  index = 0,
): MotionObservation {
  return result.report.slides[index]?.capabilities?.[
    "motion"
  ] as MotionObservation;
}

function effectMap(motion: MotionObservation | undefined) {
  return new Map(
    motion?.steps
      .flatMap((entry) => entry.effects)
      .map((effect) => [effect.target.testId, effect]) ?? [],
  );
}

describe("plain authoring choreography", () => {
  it("derives scopes, click steps, and published cues from the current document", async () => {
    const source = createMotionScopeHandle();
    const follower = createMotionScopeHandle();
    const ready = source.cue("ready");
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "pure-choreography" }).append(
        step({ n: 1 }).append(
          motionScope(source, { publishEnd: ready }).append(
            shape({
              testId: "source",
              x: 10,
              y: 10,
              width: 40,
              height: 20,
              enter: fade({ duration: 100 }),
            }),
          ),
          motionScope(follower, {
            startAt: after(ready, 25),
            delay: 5,
          }).append(
            shape({
              testId: "follower",
              x: 10,
              y: 40,
              width: 40,
              height: 20,
              enter: fade({ duration: 50 }),
            }),
          ),
        ),
        step().append(
          shape({
            testId: "independent-step",
            x: 10,
            y: 70,
            width: 40,
            height: 20,
            enter: fade({ duration: 30 }),
          }),
        ),
      ),
    );

    const result = await build(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const motion = motionOf(result);
    assert.deepEqual(
      motion.steps.map((entry) => [
        entry.trigger,
        entry.number,
        entry.effects.map(({ target, startMs, endMs }) => [
          target.testId,
          startMs,
          endMs,
        ]),
      ]),
      [
        [
          "click",
          1,
          [
            ["source", 0, 100],
            ["follower", 130, 180],
          ],
        ],
        ["click", undefined, [["independent-step", 0, 30]]],
      ],
    );
    assert.deepEqual(
      motion.scopes?.map(({ mode, members, startMs, endMs, completion }) => [
        mode,
        members.length,
        startMs,
        endMs,
        completion,
      ]),
      [
        ["parallel", 1, 0, 100, "known"],
        ["parallel", 1, 130, 180, "known"],
      ],
    );

    const query = observeMotion(result.report.slides[0]!);
    assert.ok(query !== undefined);
    assert.equal(query.steps.length, 2);
    assert.equal(query.steps[0]?.scopes[0]?.members.length, 1);
    assert.deepEqual(
      query.steps[0]?.cues.map(({ name, point }) => [name, point]),
      [["ready", "end"]],
    );
  });

  it("derives Box default entrance and disabled ancestor chains from the current tree", async () => {
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "derived-defaults" }).append(
        box({
          testId: "outer",
          enter: fade({ duration: 200, delay: 800 }),
          style: {
            width: 200,
            height: 120,
            backgroundColor: "#123456",
          },
        }).append(
          box({
            testId: "inner-background",
            style: {
              width: 120,
              height: 60,
              backgroundColor: "#654321",
            },
          }).append(
            shape({
              testId: "inherited-child",
              x: 4,
              y: 4,
              width: 20,
              height: 20,
            }),
          ),
          shape({
            testId: "own-entry-child",
            x: 40,
            y: 4,
            width: 20,
            height: 20,
            enter: fade({ duration: 100 }),
          }),
        ),
        box({
          testId: "disabled",
          motion: false,
          enter: fade({ duration: 300 }),
          style: {
            width: 100,
            height: 60,
            backgroundColor: "#445566",
          },
        }).append(
          shape({
            testId: "disabled-child",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
          shape({
            testId: "disabled-reenabled",
            motion: true,
            x: 28,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
        ),
        box({
          testId: "transparent-default",
          enter: fade({ duration: 70 }),
          style: { width: 100, height: 40 },
        }).append(
          shape({
            testId: "transparent-child",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
          }),
        ),
      ),
    );

    const result = await build(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const effects = effectMap(motionOf(result));
    assert.deepEqual(
      [
        "outer",
        "inner-background",
        "inherited-child",
        "own-entry-child",
        "transparent-child",
      ].map((testId) => {
        const effect = effects.get(testId);
        return [testId, effect?.startMs, effect?.delayMs, effect?.durationMs];
      }),
      [
        ["outer", 800, 800, 200],
        ["inner-background", 800, 800, 200],
        ["inherited-child", 800, 800, 200],
        ["own-entry-child", 0, 0, 100],
        ["transparent-child", 0, 0, 70],
      ],
    );
    assert.equal(effects.has("disabled"), false);
    assert.equal(effects.has("disabled-child"), false);
    assert.equal(effects.has("disabled-reenabled"), false);
    assert.equal(effects.has("transparent-default"), false);
  });

  it("re-derives scope and click step after an explicit same-document move", async () => {
    const firstScope = createMotionScopeHandle();
    const secondScope = createMotionScopeHandle();
    const moved = shape({
      testId: "moved",
      x: 10,
      y: 10,
      width: 40,
      height: 20,
      enter: fade({ duration: 100 }),
    });
    const firstTarget = motionScope(firstScope).append(moved);
    const secondTarget = motionScope(secondScope);
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "move-scope" }).append(
        step({ n: 1 }).append(firstTarget),
        step({ n: 2 }).append(secondTarget),
      ),
    );

    const before = await build(document);
    assert.deepEqual(
      motionOf(before).steps.map((entry) => [
        entry.number,
        entry.effects.map(({ target }) => target.testId),
      ]),
      [
        [1, ["moved"]],
        [2, []],
      ],
    );

    moved.moveTo(secondTarget);
    const afterMove = await build(document);
    assert.deepEqual(
      motionOf(afterMove).steps.map((entry) => [
        entry.number,
        entry.effects.map(({ target }) => target.testId),
      ]),
      [
        [1, []],
        [2, ["moved"]],
      ],
    );
    assert.equal(
      motionOf(afterMove).scopes?.find(({ members }) => members.length > 0)?.id,
      motionOf(afterMove).scopes?.[1]?.id,
    );
    // The completed result keeps the timeline produced by its own build.
    assert.deepEqual(
      motionOf(before).steps.map((entry) => entry.effects.length),
      [1, 0],
    );
    assert.equal(moved.type, "rp:shape");
  });

  it("re-derives scope parentage when an explicit move nests one scope in another", async () => {
    const outer = createMotionScopeHandle();
    const inner = createMotionScopeHandle();
    const outerScope = motionScope(outer);
    const innerScope = motionScope(inner).append(
      shape({
        testId: "nested",
        x: 10,
        y: 10,
        width: 40,
        height: 20,
        enter: fade(),
      }),
    );
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "move-scope-parent" }).append(outerScope, innerScope),
    );

    const before = await build(document);
    assert.equal(
      motionOf(before).scopes?.find(({ id }) => id === "scope-2")?.parentId,
      undefined,
    );

    innerScope.moveTo(outerScope);
    const after = await build(document);
    const scopes = motionOf(after).scopes;
    const outerScopeId = scopes?.[0]?.id;
    assert.equal(
      scopes?.find(({ id }) => id === "scope-2")?.parentId,
      outerScopeId,
    );
    assert.equal(
      scopes?.find(({ id }) => id === outerScopeId)?.members[0]?.kind,
      "scope",
    );
    assert.equal(
      motionOf(before).scopes?.find(({ id }) => id === "scope-2")?.parentId,
      undefined,
    );
  });

  it("reports a broken scope reference after the referenced node is removed", async () => {
    const source = createMotionScopeHandle();
    const follower = createMotionScopeHandle();
    const ready = source.cue("ready");
    const sourceScope = motionScope(source, { publishEnd: ready }).append(
      shape({
        testId: "source",
        x: 10,
        y: 10,
        width: 40,
        height: 20,
        enter: fade(),
      }),
    );
    const followerScope = motionScope(follower, {
      startAt: after(ready, 10),
    }).append(
      shape({
        testId: "follower",
        x: 10,
        y: 40,
        width: 40,
        height: 20,
        enter: fade(),
      }),
    );
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "broken-reference" }).append(sourceScope, followerScope),
    );
    assert.deepEqual(
      (await validatePptx((await build(document)).pptx)).errors,
      [],
    );

    sourceScope.remove();
    await assert.rejects(build(document), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.ok(
        error.diagnostics.some(({ code }) =>
          ["missing-scope", "missing-cue"].includes(code),
        ),
        JSON.stringify(error.diagnostics),
      );
      return true;
    });
  });

  it("keeps explicit choreography and continuity associations when values are snapshotted", async () => {
    const scope = createMotionScopeHandle();
    const cue = scope.cue("done");
    const morph = createMorphScope();
    const continuity = morph.handle("card");
    const timing = {
      type: "motion-offset" as const,
      moment: { ...cue },
      offsetMs: 40,
    };
    const first = shape({
      testId: "first",
      x: 10,
      y: 10,
      width: 40,
      height: 20,
      enter: fade(),
    });
    const document = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "associations" }).append(
        motionScope(scope, { publishEnd: cue }).append(first),
        motionScope(createMotionScopeHandle(), { startAt: timing }).append(
          shape({
            testId: "second",
            x: 10,
            y: 40,
            width: 40,
            height: 20,
            enter: fade({ duration: 50 }),
          }),
        ),
      ),
      slide({ slideId: "associations-next" }).append(
        box({
          morphKey: continuity,
          style: { width: 80, height: 40, backgroundColor: "#123456" },
        }),
      ),
    );
    timing.offsetMs = 999;
    const result = await build(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const effects = [...effectMap(motionOf(result)).values()];
    assert.equal(
      effects.find(({ target }) => target.testId === "second")?.startMs,
      540,
    );
    assert.equal(first.props["morphKey"], undefined);
    assert.equal(
      result.report.slides[1]?.objects[0]?.name?.startsWith(
        "!!react-pptx-morph-",
      ),
      true,
    );
  });

  it("matches real JSX expansion for the same content, timeline, and diagnostics", async () => {
    const pureSource = createMotionScopeHandle();
    const pureFollower = createMotionScopeHandle();
    const pureCue = pureSource.cue("ready");
    const pureDocument = deck({ width: 320, height: 180 }).append(
      slide({ slideId: "equivalent" }).append(
        step({ n: 1 }).append(
          box({
            testId: "card",
            enter: fade({ duration: 200, delay: 50 }),
            style: {
              width: 200,
              height: 100,
              backgroundColor: "#336699",
            },
          }).append(
            shape({
              testId: "card-child",
              x: 8,
              y: 8,
              width: 40,
              height: 20,
            }),
          ),
          motionScope(pureSource, { publishEnd: pureCue }).append(
            shape({
              testId: "source",
              x: 10,
              y: 20,
              width: 40,
              height: 20,
              enter: fade({ duration: 100 }),
            }),
          ),
          motionScope(pureFollower, {
            startAt: after(pureCue, 20),
          }).append(
            shape({
              testId: "follower",
              x: 10,
              y: 50,
              width: 40,
              height: 20,
              enter: fade({ duration: 40 }),
            }),
          ),
        ),
        step().append(
          shape({
            testId: "independent",
            x: 10,
            y: 80,
            width: 40,
            height: 20,
            enter: fade({ duration: 30 }),
          }),
        ),
      ),
    );

    function EquivalentDeck() {
      const source = jsx.useMotionScope();
      const follower = jsx.useMotionScope();
      return createElement(
        jsx.Deck,
        { width: 320, height: 180 },
        createElement(
          jsx.Slide,
          { slideId: "equivalent" },
          createElement(
            jsx.Step,
            { n: 1 },
            createElement(
              jsx.Box,
              {
                testId: "card",
                enter: fade({ duration: 200, delay: 50 }),
                style: {
                  width: 200,
                  height: 100,
                  backgroundColor: "#336699",
                },
              },
              createElement(jsx.Shape, {
                testId: "card-child",
                x: 8,
                y: 8,
                width: 40,
                height: 20,
              }),
            ),
            createElement(
              jsx.MotionScope,
              { scope: source, publishEnd: source.cue("ready") },
              createElement(jsx.Shape, {
                testId: "source",
                x: 10,
                y: 20,
                width: 40,
                height: 20,
                enter: fade({ duration: 100 }),
              }),
            ),
            createElement(
              jsx.MotionScope,
              { scope: follower, startAt: after(source.cue("ready"), 20) },
              createElement(jsx.Shape, {
                testId: "follower",
                x: 10,
                y: 50,
                width: 40,
                height: 20,
                enter: fade({ duration: 40 }),
              }),
            ),
          ),
          createElement(
            jsx.Step,
            {},
            createElement(jsx.Shape, {
              testId: "independent",
              x: 10,
              y: 80,
              width: 40,
              height: 20,
              enter: fade({ duration: 30 }),
            }),
          ),
        ),
      );
    }

    const [pureResult, jsxResult] = await Promise.all([
      build(pureDocument),
      jsx.build(createElement(EquivalentDeck)),
    ]);
    assert.deepEqual(
      pureResult.report.slides[0]?.objects.map(({ testId, kind, box }) => [
        testId,
        kind,
        box,
      ]),
      jsxResult.report.slides[0]?.objects.map(({ testId, kind, box }) => [
        testId,
        kind,
        box,
      ]),
    );
    const normalize = (motion: MotionObservation) => ({
      steps: motion.steps.map((entry) => [
        entry.trigger,
        entry.number,
        entry.effects.map(({ target, preset, startMs, endMs, durationMs }) => [
          target.testId,
          preset,
          startMs,
          endMs,
          durationMs,
        ]),
      ]),
      scopes: motion.scopes?.map(
        ({ mode, delayMs, startMs, endMs, completion, members }) => [
          mode,
          delayMs,
          startMs,
          endMs,
          completion,
          members.map((member) => member.kind),
        ],
      ),
    });
    assert.deepEqual(
      normalize(motionOf(pureResult)),
      normalize(motionOf(jsxResult)),
    );
    assert.deepEqual(pureResult.report.diagnostics, []);
    assert.deepEqual(jsxResult.report.diagnostics, []);
    assert.deepEqual((await validatePptx(jsxResult.pptx)).errors, []);
  });

  it("builds transitions, navigation, and Morph through the plain entry", async () => {
    const scope = createMorphScope();
    const hero = scope.handle("hero");
    const style = { width: 100, height: 50, backgroundColor: "#336699" };
    const document = deck({ width: 320, height: 180 }).append(
      slide({
        slideId: "first",
        transition: transitions.morph({
          duration: 1000,
          by: "object",
          fallback: transitions.fade(),
        }),
      }).append(
        box({ morphKey: hero, style }),
        shape({
          testId: "next",
          link: { slideId: "second" },
          x: 10,
          y: 70,
          width: 40,
          height: 20,
          fill: "#112233",
        }),
      ),
      slide({ slideId: "second" }).append(box({ morphKey: hero, style })),
    );

    const result = await build(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.equal(result.report.slides[0]?.objects[1]?.link, "slide:second");
    const names = result.report.slides
      .flatMap(({ objects }) =>
        objects.filter(({ kind }) => kind === "shape").map(({ name }) => name),
      )
      .filter((name) => name.startsWith("!!react-pptx-morph-"));
    assert.equal(names.length, 2);
    assert.equal(names[0], names[1]);
    const { unzipSync } = await import("fflate");
    const xml = new TextDecoder().decode(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"],
    );
    assert.match(xml, /p159:morph/u);
    assert.match(xml, /mc:AlternateContent/u);
    assert.match(xml, /hlinksldjump/u);
  });
});
