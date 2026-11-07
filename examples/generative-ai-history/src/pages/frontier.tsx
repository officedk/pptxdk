import { P, Shape, Text, path, preset } from "@react-pptx/react-pptx";
import { Fragment } from "react";

import { Circle, OutlineBox, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, transparent } from "../theme.ts";

export function FrontierPage() {
  return (
    <SlideFrame
      slideId="frontier"
      chapter="前沿｜视频、实时交互与推理"
      title="2024｜生成的边界继续外移"
      number={10}
      source="[18] OpenAI, 2024-02-15 · [16] OpenAI, 2024-05-13 · [17] OpenAI, 2024-09-12"
    >
      <Rule x={416} y={174} width={1} color={colors.line} thickness={446} />
      <Rule x={802} y={174} width={1} color={colors.line} thickness={446} />
      {[0, 1, 2, 3].map((index) => (
        <Fragment key={index}>
          <OutlineBox
            x={64 + index * 78}
            y={202}
            width={66}
            height={76}
            color={colors.cobalt}
          />
          <Circle
            x={76 + index * 78 + index * 6}
            y={238 - index * 5}
            diameter={18}
            fill={index === 3 ? colors.sun : colors.cyan}
          />
          <Shape
            geometry={path("M 0 42 L 22 20 L 34 36 L 48 14 L 66 42", {
              viewBox: [66, 76],
              fill: false,
            })}
            x={64 + index * 78}
            y={202}
            width={66}
            height={76}
            fill={transparent}
            style={{ borderColor: colors.ink, borderWidth: 1 }}
          />
        </Fragment>
      ))}
      <Text
        x={64}
        y={312}
        width={320}
        height={26}
        style={{ fontSize: 14, color: colors.cobalt }}
      >
        02.15 · Sora 技术预览
      </Text>
      <Text
        x={64}
        y={354}
        width={320}
        height={44}
        style={{ fontFamily: "display", fontSize: 29, color: colors.ink }}
      >
        生成动态场景
      </Text>
      <Text
        x={64}
        y={414}
        width={310}
        height={110}
        style={{ fontSize: 17, lineHeight: 28, color: colors.muted }}
      >
        OpenAI 展示文本生成视频模型 Sora，并把视频生成模型描述为世界模拟器。
      </Text>
      <Text
        x={64}
        y={554}
        width={310}
        height={60}
        style={{ fontSize: 14, lineHeight: 22, color: colors.cobalt }}
      >
        <P>输出不再止于静态画面，</P>
        <P>还要保持时间与空间的一致性。</P>
      </Text>
      <Shape
        geometry={preset("ellipse")}
        x={506}
        y={196}
        width={132}
        height={96}
        fill={transparent}
        style={{ borderColor: colors.cobalt, borderWidth: 2 }}
      />
      <Circle x={552} y={222} diameter={42} fill={colors.sun} />
      {[0, 1, 2, 3, 4, 5, 6].map((index) => (
        <Shape
          key={index}
          x={674 + index * 12}
          y={238 - (index % 3) * 12}
          width={5}
          height={22 + (index % 3) * 24}
          fill={colors.cyan}
        />
      ))}
      <Text
        x={450}
        y={312}
        width={320}
        height={26}
        style={{ fontSize: 14, color: colors.cobalt }}
      >
        05.13 · GPT-4o
      </Text>
      <Text
        x={450}
        y={354}
        width={320}
        height={44}
        style={{ fontFamily: "display", fontSize: 29, color: colors.ink }}
      >
        打通语音与视觉
      </Text>
      <Text
        x={450}
        y={414}
        width={310}
        height={110}
        style={{ fontSize: 17, lineHeight: 28, color: colors.muted }}
      >
        同一模型端到端处理文本、音频、图像与视频输入，并生成文本、音频或图像；各模态分阶段开放。
      </Text>
      <Text
        x={450}
        y={554}
        width={310}
        height={60}
        style={{ fontSize: 14, lineHeight: 22, color: colors.cobalt }}
      >
        <P>多模态从多次转换，</P>
        <P>走向更直接的实时交互。</P>
      </Text>
      <Shape
        geometry={path(
          "M 0 100 L 60 100 L 120 20 M 60 100 L 120 60 M 60 100 L 120 100 M 60 100 L 120 140 M 60 100 L 120 180",
          { viewBox: [140, 200], fill: false },
        )}
        x={870}
        y={190}
        width={160}
        height={118}
        fill={transparent}
        style={{ borderColor: colors.cobalt, borderWidth: 2 }}
      />
      <Circle x={1040} y={190} diameter={16} fill={colors.cyan} />
      <Circle x={1040} y={224} diameter={16} fill={colors.cyan} />
      <Circle x={1040} y={258} diameter={16} fill={colors.sun} />
      <Circle x={1040} y={292} diameter={16} fill={colors.cyan} />
      <Text
        x={836}
        y={354}
        width={380}
        height={44}
        style={{ fontFamily: "display", fontSize: 29, color: colors.ink }}
      >
        把计算留给推理
      </Text>
      <Text
        x={836}
        y={312}
        width={380}
        height={26}
        style={{ fontSize: 14, color: colors.cobalt }}
      >
        09.12 · OpenAI o1
      </Text>
      <Text
        x={836}
        y={414}
        width={350}
        height={110}
        style={{ fontSize: 17, lineHeight: 28, color: colors.muted }}
      >
        o1 通过强化学习训练更长的内部推理过程，在回答前投入更多计算。
      </Text>
      <Text
        x={836}
        y={554}
        width={350}
        height={60}
        style={{ fontSize: 14, lineHeight: 22, color: colors.cobalt }}
      >
        <P>扩展的对象从参数与数据，</P>
        <P>进一步延伸到推理时间。</P>
      </Text>
    </SlideFrame>
  );
}
