import assert from "node:assert/strict";

import {
  build,
  Deck,
  Slide,
  Shape,
  Step,
  fillColor,
  BuildError,
  repeat,
  Text,
  MotionScope,
  useMotionScope,
  fade,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { unzipSync, strFromU8, strToU8, zipSync } from "fflate";
import { it } from "vitest";

it("填充色强调按点击步变色并保持终态，允许从亮色变为暗色", async () => {
  const result = await build(
    <Deck width={320} height={180}>
      <Slide slideId="color">
        <Step n={1}>
          <Shape
            x={20}
            y={20}
            width={80}
            height={40}
            fill="#EEEEEE"
            emphasis={fillColor({
              from: "#eeeeee",
              to: "#16324f",
              duration: 600,
            })}
          />
        </Step>
      </Slide>
    </Deck>,
  );
  const file = await validatePptx(result.pptx);
  assert.deepEqual(file.errors, []);
  const steps = file.slides[0]?.animations;
  assert.equal(steps?.[0]?.trigger, "click");
  const effect = steps?.[0]?.effects[0];
  assert.deepEqual(
    [effect?.preset, effect?.durationMs, effect?.fromColor, effect?.toColor],
    ["fillColor", 600, "#EEEEEE", "#16324F"],
  );
  const xml = strFromU8(unzipSync(result.pptx)["ppt/slides/slide1.xml"]!);
  assert.match(xml, /<p:animClr/);
  assert.match(xml, /dur="600" fill="hold"/);
  assert.doesNotMatch(xml, /autoRev=/);
  assert.match(xml, /<p:from><a:srgbClr val="EEEEEE"/);
  assert.match(xml, /<p:to><a:srgbClr val="16324F"/);
});

it("拒绝非法颜色、重复与没有实色填充的目标", async () => {
  for (const [content, code] of [
    [
      <Shape
        x={0}
        y={0}
        width={40}
        height={40}
        fill="#FFFFFF"
        emphasis={fillColor({ from: "red", to: "#16324F" })}
      />,
      "invalid-fill-color",
    ],
    [
      <Shape
        x={0}
        y={0}
        width={40}
        height={40}
        fill="#FFFFFF"
        emphasis={repeat(fillColor({ from: "#FFFFFF", to: "#16324F" }), 2)}
      />,
      "unsupported-repeat",
    ],
    [
      <Shape
        x={0}
        y={0}
        width={40}
        height={40}
        emphasis={fillColor({ from: "#FFFFFF", to: "#16324F" })}
      />,
      "invalid-fill-color-target",
    ],
    [
      <Text
        x={0}
        y={0}
        width={120}
        height={40}
        emphasis={fillColor({ from: "#FFFFFF", to: "#16324F" })}
      >
        文字
      </Text>,
      "invalid-fill-color-target",
    ],
  ] as const) {
    await assert.rejects(
      build(
        <Deck width={320} height={180}>
          <Slide slideId="invalid">{content}</Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, code);
        return true;
      },
    );
  }
});

it("多个变色组件的完成时刻能连接后续内容，重排后仍各自独立", async () => {
  function Region({ name, n }: { name: string; n: number }) {
    const region = useMotionScope();
    const caption = useMotionScope();
    return (
      <Step n={n}>
        <MotionScope scope={region}>
          <Shape
            testId={name}
            x={20}
            y={n * 40}
            width={50}
            height={30}
            fill="#FFFFFF"
            emphasis={fillColor({
              from: "#FFFFFF",
              to: "#138A83",
              duration: 400,
            })}
          />
        </MotionScope>
        <MotionScope scope={caption} startAt={region.end}>
          <Shape
            testId={`${name}-caption`}
            x={80}
            y={n * 40}
            width={100}
            height={30}
            fill="#16324F"
            enter={fade({ duration: 200 })}
          />
        </MotionScope>
      </Step>
    );
  }
  for (const names of [
    ["east", "west"],
    ["west", "east"],
  ]) {
    const result = await build(
      <Deck width={320} height={180}>
        <Slide slideId="regions">
          {names.map((name, index) => (
            <Region key={name} name={name} n={index + 1} />
          ))}
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const steps = file.slides[0]?.animations;
    assert.equal(steps?.length, 2);
    for (const [index, step] of (steps ?? []).entries()) {
      assert.equal(step.trigger, "click");
      assert.deepEqual(
        step.effects.map(({ preset, durationMs, delayMs }) => [
          preset,
          durationMs,
          delayMs,
        ]),
        [
          ["fillColor", 400, 0],
          ["fade", 200, 400],
        ],
      );
      const target = file.slides[0]?.objects.find(
        ({ id }) => id === step.effects[0]?.targetId,
      );
      assert.equal(target?.name, names[index]);
    }
  }
});

it("独立检查拒绝丢失终态保持的颜色动画", async () => {
  const result = await build(
    <Deck width={320} height={180}>
      <Slide slideId="hold">
        <Shape
          x={0}
          y={0}
          width={40}
          height={40}
          fill="#FFFFFF"
          emphasis={fillColor({ from: "#FFFFFF", to: "#138A83" })}
        />
      </Slide>
    </Deck>,
  );
  const parts = unzipSync(result.pptx);
  parts["ppt/slides/slide1.xml"] = strToU8(
    strFromU8(parts["ppt/slides/slide1.xml"]!).replaceAll(
      'fill="hold"',
      'fill="remove"',
    ),
  );
  const file = await validatePptx(zipSync(parts));
  assert.ok(
    file.errors.some(({ code }) => code === "invalid-fill-color-animation"),
  );
});
