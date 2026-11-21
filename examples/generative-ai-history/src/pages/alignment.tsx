import { Text } from "pptxdk/react";
import { Fragment } from "react";

import { Arrow, OutlineBox, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const steps = [
  { number: "1", title: "预训练", body: "从大规模文本中\n学习预测下一个词" },
  { number: "2", title: "示范微调", body: "用人工示范回答\n教模型遵循指令" },
  { number: "3", title: "偏好建模", body: "比较候选回答\n训练奖励模型" },
  { number: "4", title: "强化学习", body: "依据奖励信号\n优化回答策略" },
] as const;

export function AlignmentPage() {
  return (
    <SlideFrame
      id="alignment"
      chapter="对齐｜InstructGPT 与 ChatGPT"
      title="2022｜从会补全，到听得懂指令"
      number={8}
      source="[12] Ouyang et al., 2022 · [13] OpenAI, 2022-11-30"
    >
      <Text
        x={64}
        y={158}
        width={1148}
        height={38}
        style={{ fontSize: 18, color: colors.muted }}
      >
        人类反馈强化学习（RLHF）把“哪种回答更有用”加入训练流程。
      </Text>
      {steps.map((step, index) => {
        const x = 64 + index * 288;
        return (
          <Fragment key={index}>
            <OutlineBox
              x={x}
              y={224}
              width={240}
              height={132}
              color={index === 3 ? colors.sun : colors.cobalt}
            />
            <Text
              x={x + 18}
              y={240}
              width={34}
              height={36}
              style={{
                fontFamily: "display",
                fontSize: 28,
                color: index === 3 ? colors.sun : colors.cobalt,
              }}
            >
              {step.number}
            </Text>
            <Text
              x={x + 58}
              y={242}
              width={160}
              height={34}
              style={{ fontSize: 22, color: colors.ink }}
            >
              {step.title}
            </Text>
            <Text
              x={x + 20}
              y={292}
              width={200}
              height={54}
              style={{
                fontSize: 15,
                lineHeight: 24,
                color: colors.muted,
                align: "center",
              }}
            >
              {step.body}
            </Text>
            {index < steps.length - 1 && (
              <Arrow x={x + 250} y={280} width={28} color={colors.cyan} />
            )}
          </Fragment>
        );
      })}
      <Rule x={64} y={398} width={1152} color={colors.line} />
      <Text
        x={64}
        y={430}
        width={242}
        height={32}
        style={{ fontSize: 16, color: colors.cobalt }}
      >
        语言模型原有目标
      </Text>
      <Text
        x={316}
        y={420}
        width={896}
        height={52}
        style={{ fontFamily: "display", fontSize: 28, color: colors.ink }}
      >
        给定上下文，续写统计上合适的文本
      </Text>
      <Rule x={64} y={486} width={1152} color={colors.mist} />
      <Text
        x={64}
        y={520}
        width={242}
        height={32}
        style={{ fontSize: 16, color: colors.cobalt }}
      >
        对齐新增的训练信号
      </Text>
      <Text
        x={316}
        y={510}
        width={896}
        height={52}
        style={{ fontFamily: "body", fontSize: 26, color: colors.ink }}
      >
        依据人类偏好，选择更有帮助、更符合意图的回答
      </Text>
      <Text
        x={64}
        y={606}
        width={1148}
        height={34}
        style={{ fontSize: 15, color: colors.muted, align: "center" }}
      >
        1 月 27 日：OpenAI 发布 InstructGPT　·　11 月 30 日：ChatGPT
        研究预览推出
      </Text>
    </SlideFrame>
  );
}
