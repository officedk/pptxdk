import assert from "node:assert/strict";

import {
  build,
  BuildError,
  Chart,
  createBuilder,
  Deck,
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

describe("对象动画", () => {
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
              ["fly", 1000, undefined, undefined],
            ]
          : [
              ["fly", 1000, undefined, undefined],
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
    assert.deepEqual(follower?.startAfter, {
      targetId: source?.target.id,
      effectIndex: 0,
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
    const sourceScope = motion.scopes?.find(({ members }) =>
      members.some(
        (member) => member.kind === "object" && member.id === source?.target.id,
      ),
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
        followerScope?.startAfter?.targetId,
      ],
      ["unknown", undefined, undefined, source?.target.id],
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
    assert.equal(followerFileEffect?.startAfter?.targetId, source?.target.id);
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
