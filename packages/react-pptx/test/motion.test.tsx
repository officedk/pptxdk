import assert from "node:assert/strict";

import {
  build,
  BuildError,
  createBuilder,
  Deck,
  fade,
  fly,
  floatUp,
  Group,
  Image,
  Shape,
  shapeCapability,
  Slide,
  spin,
  spinIn,
  stretch,
  Step,
  Text,
  motionCapability,
  transitionCapability,
  transitions,
  type MotionObservation,
  wipe,
  zoom,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

describe("对象动画", () => {
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

  it("拒绝非法效果参数、缺失的 motion 能力和不存在的文件目标", async () => {
    for (const [effect, code] of [
      [fade({ duration: 0 }), "invalid-duration"],
      [fade({ duration: null } as never), "invalid-duration"],
      [fade({ delay: -1 }), "invalid-delay"],
      [fade({ unexpected: true } as never), "unknown-parameter"],
      [fade({ easing: null } as never), "invalid-easing"],
      [
        fade({ easing: { accelerate: 0.8, decelerate: 0.4 } }),
        "invalid-easing",
      ],
      [wipe({ from: null } as never), "invalid-direction"],
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
