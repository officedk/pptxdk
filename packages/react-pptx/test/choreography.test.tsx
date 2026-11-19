import assert from "node:assert/strict";

import * as jsx from "@react-pptx/react-pptx";
import { createResolver } from "@react-pptx/react-pptx/build";
import {
  after,
  box,
  build,
  BuildError,
  createMorphScope,
  createMotionScopeHandle,
  deck,
  documentCapabilities,
  emit,
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
import { strFromU8, unzipSync } from "fflate";
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
      .map((effect) => [effect.target.authorId, effect]) ?? [],
  );
}

function hasTimingXml(result: { readonly pptx: Uint8Array }): boolean {
  const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
  return xml.includes("<p:timing");
}

describe("plain authoring choreography", () => {
  it("derives scopes, click steps, and published cues from the current document", async () => {
    const source = createMotionScopeHandle();
    const follower = createMotionScopeHandle();
    const ready = source.cue("ready");
    const document = deck({ width: 320, height: 180 }).append(
      slide({ id: "pure-choreography" }).append(
        step({ n: 1 }).append(
          motionScope(source, { publishEnd: ready }).append(
            shape({
              id: "source",
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
              id: "follower",
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
            id: "independent-step",
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
          target.authorId,
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
      slide({ id: "derived-defaults" }).append(
        box({
          id: "outer",
          enter: fade({ duration: 200, delay: 800 }),
          style: {
            width: 200,
            height: 120,
            backgroundColor: "#123456",
          },
        }).append(
          box({
            id: "inner-background",
            style: {
              width: 120,
              height: 60,
              backgroundColor: "#654321",
            },
          }).append(
            shape({
              id: "inherited-child",
              x: 4,
              y: 4,
              width: 20,
              height: 20,
            }),
          ),
          shape({
            id: "own-entry-child",
            x: 40,
            y: 4,
            width: 20,
            height: 20,
            enter: fade({ duration: 100 }),
          }),
        ),
        box({
          id: "disabled",
          motion: false,
          enter: fade({ duration: 300 }),
          style: {
            width: 100,
            height: 60,
            backgroundColor: "#445566",
          },
        }).append(
          shape({
            id: "disabled-child",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
          shape({
            id: "disabled-reenabled",
            motion: true,
            x: 28,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
        ),
        box({
          id: "transparent-default",
          enter: fade({ duration: 70 }),
          style: { width: 100, height: 40 },
        }).append(
          shape({
            id: "transparent-child",
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

  it("checks trimmed capabilities against the structure-resolved motion", async () => {
    const manualModules = [jsx.boxCapability, jsx.shapeCapability] as const;
    const filteredResolve = createResolver(
      documentCapabilities.filter(
        (capability) => capability.declaration !== "motion",
      ),
    );
    const manualFullResolve = createResolver([
      ...manualModules,
      jsx.motionCapability,
    ]);
    const manualCroppedResolve = createResolver(manualModules);
    const filtered = async (doc: Parameters<typeof build>[0]) =>
      emit(filteredResolve(doc));
    const fullBuilds = [
      ["default-full", build],
      [
        "manual-full",
        async (doc: Parameters<typeof build>[0]) =>
          emit(manualFullResolve(doc)),
      ],
    ] as const;
    const croppedBuilds = [
      ["default-filtered", filtered],
      [
        "manual-cropped",
        async (doc: Parameters<typeof build>[0]) =>
          emit(manualCroppedResolve(doc)),
      ],
    ] as const;
    const style = { width: 100, height: 60 };
    const scenarios = [
      {
        name: "inherited-default",
        content: box({ enter: fade({ duration: 70 }), style }).append(
          shape({
            id: "inherited",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
          }),
        ),
        fullTiming: true,
        cropped: "missing",
      },
      {
        name: "disabled-explicit",
        content: box({ motion: false, enter: fade(), style }).append(
          shape({
            id: "disabled",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
        ),
        fullTiming: false,
        cropped: "silent",
      },
      {
        name: "plain",
        content: box({ style }).append(
          shape({
            id: "plain",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
          }),
        ),
        fullTiming: false,
        cropped: "silent",
      },
      {
        name: "local-override",
        content: box({ enter: fade({ duration: 70 }), style }).append(
          shape({
            id: "override",
            x: 4,
            y: 4,
            width: 20,
            height: 20,
            enter: fade({ duration: 20 }),
          }),
        ),
        fullTiming: true,
        cropped: "missing",
      },
      {
        name: "disabled-reenabled",
        content: box({ motion: false, enter: fade(), style }).append(
          shape({
            id: "reenabled",
            motion: true,
            x: 4,
            y: 4,
            width: 20,
            height: 20,
            enter: fade(),
          }),
        ),
        fullTiming: false,
        cropped: "silent",
      },
    ] as const;

    for (const scenario of scenarios) {
      const document = deck({ width: 320, height: 180 }).append(
        slide({ id: scenario.name }).append(scenario.content),
      );
      for (const [mode, run] of fullBuilds) {
        const result = await run(document);
        assert.equal(
          hasTimingXml(result),
          scenario.fullTiming,
          `${scenario.name} ${mode}`,
        );
        assert.deepEqual(
          (await validatePptx(result.pptx)).errors,
          [],
          `${scenario.name} ${mode}`,
        );
      }

      for (const [mode, run] of croppedBuilds) {
        if (scenario.cropped === "missing") {
          await assert.rejects(run(document), (error: unknown) => {
            assert.ok(error instanceof BuildError);
            assert.ok(
              error.diagnostics.some(
                ({ code, path }) =>
                  code === "missing-capability" &&
                  path.includes("shape[") &&
                  path.includes(
                    scenario.name === "local-override"
                      ? "override"
                      : "inherited",
                  ),
              ),
              `${scenario.name} ${mode}`,
            );
            return true;
          });
        } else {
          const result = await run(document);
          assert.equal(hasTimingXml(result), false, `${scenario.name} ${mode}`);
          assert.deepEqual(
            (await validatePptx(result.pptx)).errors,
            [],
            `${scenario.name} ${mode}`,
          );
        }
      }
    }
  });

  it("checks trimmed JSX capabilities against the structure-resolved motion", async () => {
    const manualModules = [jsx.boxCapability, jsx.shapeCapability] as const;
    const filteredResolve = createResolver(
      documentCapabilities.filter(
        (capability) => capability.declaration !== "motion",
      ),
    );
    const manualFullResolve = createResolver([
      ...manualModules,
      jsx.motionCapability,
    ]);
    const manualCroppedResolve = createResolver(manualModules);
    const filtered = async (doc: Parameters<typeof build>[0]) =>
      emit(filteredResolve(doc));
    const manualFull = async (
      element: Parameters<typeof jsx.expandDocument>[0],
    ) => emit(manualFullResolve(await jsx.expandDocument(element)));
    const manualCropped = async (
      element: Parameters<typeof jsx.expandDocument>[0],
    ) => emit(manualCroppedResolve(await jsx.expandDocument(element)));
    const style = { width: 100, height: 60 };
    const cases = [
      {
        name: "jsx-inherited",
        element: createElement(
          jsx.Deck,
          { width: 320, height: 180 },
          createElement(
            jsx.Slide,
            { id: "jsx-inherited" },
            createElement(
              jsx.Box,
              { enter: fade({ duration: 70 }), style },
              createElement(jsx.Shape, {
                id: "jsx-inherited",
                x: 4,
                y: 4,
                width: 20,
                height: 20,
              }),
            ),
          ),
        ),
        fullTiming: true,
        cropped: "missing",
      },
      {
        name: "jsx-disabled",
        element: createElement(
          jsx.Deck,
          { width: 320, height: 180 },
          createElement(
            jsx.Slide,
            { id: "jsx-disabled" },
            createElement(
              jsx.Box,
              { motion: false, enter: fade(), style },
              createElement(jsx.Shape, {
                id: "jsx-disabled",
                x: 4,
                y: 4,
                width: 20,
                height: 20,
                enter: fade(),
              }),
            ),
          ),
        ),
        fullTiming: false,
        cropped: "silent",
      },
    ] as const;

    for (const scenario of cases) {
      const expanded = await jsx.expandDocument(scenario.element);
      const fullBuilds = [
        ["default-full", () => build(expanded)],
        ["manual-full", () => manualFull(scenario.element)],
      ] as const;
      const croppedBuilds = [
        ["default-filtered", () => filtered(expanded)],
        ["manual-cropped", () => manualCropped(scenario.element)],
      ] as const;
      for (const [mode, run] of fullBuilds) {
        const result = await run();
        assert.equal(
          hasTimingXml(result),
          scenario.fullTiming,
          `${scenario.name} ${mode}`,
        );
        assert.deepEqual(
          (await validatePptx(result.pptx)).errors,
          [],
          `${scenario.name} ${mode}`,
        );
      }
      for (const [mode, run] of croppedBuilds) {
        if (scenario.cropped === "missing") {
          await assert.rejects(run, (error: unknown) => {
            assert.ok(error instanceof BuildError);
            assert.ok(
              error.diagnostics.some(
                ({ code, path }) =>
                  code === "missing-capability" &&
                  path.includes(`shape[${scenario.name}]`),
              ),
              `${scenario.name} ${mode}`,
            );
            return true;
          });
        } else {
          const result = await run();
          assert.equal(hasTimingXml(result), false, `${scenario.name} ${mode}`);
          assert.deepEqual(
            (await validatePptx(result.pptx)).errors,
            [],
            `${scenario.name} ${mode}`,
          );
        }
      }
    }
  });

  it("re-derives scope and click step after an explicit same-document move", async () => {
    const firstScope = createMotionScopeHandle();
    const secondScope = createMotionScopeHandle();
    const moved = shape({
      id: "moved",
      x: 10,
      y: 10,
      width: 40,
      height: 20,
      enter: fade({ duration: 100 }),
    });
    const firstTarget = motionScope(firstScope).append(moved);
    const secondTarget = motionScope(secondScope);
    const document = deck({ width: 320, height: 180 }).append(
      slide({ id: "move-scope" }).append(
        step({ n: 1 }).append(firstTarget),
        step({ n: 2 }).append(secondTarget),
      ),
    );

    const before = await build(document);
    assert.deepEqual(
      motionOf(before).steps.map((entry) => [
        entry.number,
        entry.effects.map(({ target }) => target.authorId),
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
        entry.effects.map(({ target }) => target.authorId),
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
        id: "nested",
        x: 10,
        y: 10,
        width: 40,
        height: 20,
        enter: fade(),
      }),
    );
    const document = deck({ width: 320, height: 180 }).append(
      slide({ id: "move-scope-parent" }).append(outerScope, innerScope),
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
        id: "source",
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
        id: "follower",
        x: 10,
        y: 40,
        width: 40,
        height: 20,
        enter: fade(),
      }),
    );
    const document = deck({ width: 320, height: 180 }).append(
      slide({ id: "broken-reference" }).append(sourceScope, followerScope),
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
      id: "first",
      x: 10,
      y: 10,
      width: 40,
      height: 20,
      enter: fade(),
    });
    const document = deck({ width: 320, height: 180 }).append(
      slide({ id: "associations" }).append(
        motionScope(scope, { publishEnd: cue }).append(first),
        motionScope(createMotionScopeHandle(), { startAt: timing }).append(
          shape({
            id: "second",
            x: 10,
            y: 40,
            width: 40,
            height: 20,
            enter: fade({ duration: 50 }),
          }),
        ),
      ),
      slide({ id: "associations-next" }).append(
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
      effects.find(({ target }) => target.authorId === "second")?.startMs,
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
      slide({ id: "equivalent" }).append(
        step({ n: 1 }).append(
          box({
            id: "card",
            enter: fade({ duration: 200, delay: 50 }),
            style: {
              width: 200,
              height: 100,
              backgroundColor: "#336699",
            },
          }).append(
            shape({
              id: "card-child",
              x: 8,
              y: 8,
              width: 40,
              height: 20,
            }),
          ),
          motionScope(pureSource, { publishEnd: pureCue }).append(
            shape({
              id: "source",
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
              id: "follower",
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
            id: "independent",
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
          { id: "equivalent" },
          createElement(
            jsx.Step,
            { n: 1 },
            createElement(
              jsx.Box,
              {
                id: "card",
                enter: fade({ duration: 200, delay: 50 }),
                style: {
                  width: 200,
                  height: 100,
                  backgroundColor: "#336699",
                },
              },
              createElement(jsx.Shape, {
                id: "card-child",
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
                id: "source",
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
                id: "follower",
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
              id: "independent",
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
      pureResult.report.slides[0]?.objects.map(
        ({ authorId: testId, kind, box }) => [testId, kind, box],
      ),
      jsxResult.report.slides[0]?.objects.map(
        ({ authorId: testId, kind, box }) => [testId, kind, box],
      ),
    );
    const normalize = (motion: MotionObservation) => ({
      steps: motion.steps.map((entry) => [
        entry.trigger,
        entry.number,
        entry.effects.map(({ target, preset, startMs, endMs, durationMs }) => [
          target.authorId,
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
        id: "first",
        transition: transitions.morph({
          duration: 1000,
          by: "object",
          fallback: transitions.fade(),
        }),
      }).append(
        box({ morphKey: hero, style }),
        shape({
          id: "next",
          link: { id: "second" },
          x: 10,
          y: 70,
          width: 40,
          height: 20,
          fill: "#112233",
        }),
      ),
      slide({ id: "second" }).append(box({ morphKey: hero, style })),
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
