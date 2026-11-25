import { Slide, Text } from "@officedk/pptxdk/react";
import { Fragment } from "react";

import { OutlineBox, Rule } from "../components/primitives.tsx";
import { contentLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

const questions = [
  {
    number: "01",
    title: "可靠性",
    claim: "让流畅回答经得起核验",
    body: "评估、引用与工具校验，决定生成结果能否进入高风险流程。",
  },
  {
    number: "02",
    title: "权利",
    claim: "让训练与创作建立可追溯的关系",
    body: "数据授权、署名和收益分配，仍需要技术与制度协作。",
  },
  {
    number: "03",
    title: "可达性",
    claim: "让更多人真正用得上这些能力",
    body: "成本、能耗、语言和地域差异，会影响技术的实际受益者。",
  },
] as const;

export function QuestionsPage() {
  return (
    <Slide
      id="questions"
      layout={contentLayout}
      placeholders={{
        chapter: "开放问题｜历史仍在书写",
        title: "下一章，由约束共同写成",
        source: "本页为基于前述历史的观察与开放问题",
      }}
    >
      <Text
        x={64}
        y={158}
        width={1148}
        height={38}
        style={{ fontSize: 18, color: colors.muted }}
      >
        生成能力继续扩展时，三个问题会与技术进步并行。
      </Text>
      {questions.map((question, index) => {
        const y = 226 + index * 116;
        return (
          <Fragment key={index}>
            <Text
              x={64}
              y={y}
              width={76}
              height={58}
              style={{
                fontFamily: "display",
                fontSize: 43,
                color: index === 2 ? colors.sun : colors.cobalt,
              }}
            >
              {question.number}
            </Text>
            <Text
              x={164}
              y={y + 2}
              width={134}
              height={42}
              style={{ fontFamily: "display", fontSize: 29, color: colors.ink }}
            >
              {question.title}
            </Text>
            <Text
              x={330}
              y={y + 2}
              width={410}
              height={38}
              style={{ fontSize: 21, color: colors.cobalt }}
            >
              {question.claim}
            </Text>
            <Text
              x={776}
              y={y}
              width={440}
              height={74}
              style={{ fontSize: 16, lineHeight: 26, color: colors.muted }}
            >
              {question.body}
            </Text>
            <Rule x={64} y={y + 88} width={1152} color={colors.line} />
          </Fragment>
        );
      })}
      <OutlineBox
        x={64}
        y={586}
        width={1152}
        height={64}
        color={colors.cobalt}
      />
      <Text
        x={90}
        y={596}
        width={1100}
        height={42}
        style={{
          fontFamily: "display",
          fontSize: 26,
          color: colors.ink,
          align: "center",
        }}
      >
        技术史的下一次转折，也会是一场关于使用方式的抉择。
      </Text>
    </Slide>
  );
}
