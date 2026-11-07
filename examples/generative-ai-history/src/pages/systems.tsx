import { Shape, Text, path, preset } from "@react-pptx/react-pptx";
import { Fragment } from "react";

import { OutlineBox } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, transparent } from "../theme.ts";

const nodes = [
  {
    x: 128,
    y: 216,
    title: "检索",
    body: "引入模型之外的资料",
    center: [256, 260],
  },
  {
    x: 128,
    y: 448,
    title: "代码执行",
    body: "把计算交给确定性工具",
    center: [256, 492],
  },
  {
    x: 952,
    y: 216,
    title: "图像理解",
    body: "让视觉进入同一上下文",
    center: [1024, 260],
  },
  {
    x: 952,
    y: 448,
    title: "外部工具",
    body: "把回答转化为实际动作",
    center: [1024, 492],
  },
] as const;

export function SystemsPage() {
  return (
    <SlideFrame
      slideId="systems"
      chapter="系统｜模型与外部环境协同"
      title="2023｜从单个模型，到协同系统"
      number={9}
      source="[14] Yao et al., 2022 / ICLR 2023 · [15] OpenAI, GPT-4 Technical Report, 2023"
    >
      {nodes.map((node) => (
        <Shape
          key={node.title}
          geometry={path(
            `M 576 190 L ${node.center[0] - 64} ${node.center[1] - 170}`,
            {
              viewBox: [1152, 400],
              fill: false,
            },
          )}
          x={64}
          y={170}
          width={1152}
          height={400}
          fill={transparent}
          style={{
            borderColor: colors.cyan,
            borderWidth: 2,
            borderStyle: "dashed",
          }}
        />
      ))}
      <Shape
        geometry={preset("ellipse")}
        x={548}
        y={270}
        width={184}
        height={184}
        fill={colors.indigo}
      />
      <Shape
        geometry={preset("ellipse")}
        x={566}
        y={288}
        width={148}
        height={148}
        fill={transparent}
        style={{ borderColor: colors.sun, borderWidth: 2 }}
      />
      <Text
        x={566}
        y={318}
        width={148}
        height={48}
        style={{
          fontFamily: "display",
          fontSize: 29,
          color: colors.paperBright,
          align: "center",
        }}
      >
        语言模型
      </Text>
      <Text
        x={566}
        y={370}
        width={148}
        height={30}
        style={{ fontSize: 14, color: colors.cyan, align: "center" }}
      >
        规划 · 生成 · 选择
      </Text>
      {nodes.map((node, index) => (
        <Fragment key={index}>
          <OutlineBox
            x={node.x}
            y={node.y}
            width={200}
            height={88}
            color={index === 3 ? colors.sun : colors.cobalt}
          />
          <Text
            x={node.x + 16}
            y={node.y + 12}
            width={168}
            height={30}
            style={{ fontSize: 20, color: colors.ink, align: "center" }}
          >
            {node.title}
          </Text>
          <Text
            x={node.x + 12}
            y={node.y + 48}
            width={176}
            height={26}
            style={{ fontSize: 13, color: colors.muted, align: "center" }}
          >
            {node.body}
          </Text>
        </Fragment>
      ))}
      <Text
        x={390}
        y={174}
        width={500}
        height={38}
        style={{
          fontFamily: "display",
          fontSize: 26,
          color: colors.cobalt,
          align: "center",
        }}
      >
        生成的单位，从一段内容变成一次任务
      </Text>
      <Text
        x={64}
        y={592}
        width={552}
        height={62}
        style={{ fontSize: 14, lineHeight: 22, color: colors.muted }}
      >
        ReAct 把推理轨迹与行动交替组织，让模型能够围绕外部结果继续规划。
      </Text>
      <Text
        x={664}
        y={592}
        width={552}
        height={62}
        style={{ fontSize: 14, lineHeight: 22, color: colors.muted }}
      >
        GPT-4 技术报告进一步描述了图像与文本输入、文本输出的多模态能力。
      </Text>
    </SlideFrame>
  );
}
