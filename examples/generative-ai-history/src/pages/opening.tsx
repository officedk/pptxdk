import { P, Slide, Text } from "@react-pptx/react-pptx";

import { SamplingField } from "../components/sampling-field.tsx";
import { colors } from "../theme.ts";

export function OpeningPage() {
  return (
    <Slide id="opening" background={colors.indigo}>
      <Text
        x={72}
        y={64}
        width={520}
        height={32}
        style={{ fontSize: 16, color: colors.cyan, letterSpacing: 2 }}
      >
        1950—2024 · 技术、媒介与人的意图
      </Text>
      <Text
        x={68}
        y={154}
        width={650}
        height={206}
        style={{
          fontFamily: "display",
          fontSize: 76,
          lineHeight: 98,
          color: colors.paperBright,
        }}
      >
        <P>机器如何</P>
        <P>学会创造</P>
      </Text>
      <Text
        x={72}
        y={390}
        width={540}
        height={48}
        style={{ fontSize: 27, color: colors.sun }}
      >
        生成式人工智能简史
      </Text>
      <Text
        x={72}
        y={464}
        width={550}
        height={68}
        style={{ fontSize: 19, lineHeight: 30, color: colors.mist }}
      >
        <P>从图灵之问出发，追踪分布学习、规模扩展与交互方式</P>
        <P>如何共同塑造会写、会画、会推理的模型。</P>
      </Text>
      <SamplingField x={738} y={118} size={452} dark />
      <Text
        x={850}
        y={590}
        width={250}
        height={26}
        style={{ fontSize: 14, color: colors.cyan, align: "center" }}
      >
        噪声 → 结构 → 生成
      </Text>
      <Text
        x={72}
        y={656}
        width={600}
        height={24}
        style={{ fontSize: 12, color: colors.muted }}
      >
        一份以原始论文与官方发布为依据的视觉年表
      </Text>
    </Slide>
  );
}
