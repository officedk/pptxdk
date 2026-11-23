import { P, Slide, Text } from "pptxdk/react";

import { SamplingField } from "../components/sampling-field.tsx";
import { coverLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

export function OpeningPage() {
  return (
    <Slide
      id="opening"
      layout={coverLayout}
      showMasterShapes={false}
      placeholders={{
        title: [<P key="first">机器如何</P>, <P key="second">学会创造</P>],
        subtitle: "生成式人工智能简史",
      }}
    >
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
