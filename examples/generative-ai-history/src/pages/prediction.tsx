import { Text } from "pptxdk/react";
import { Fragment } from "react";

import { OutlineBox, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors } from "../theme.ts";

const candidates = [
  { word: "创造", value: 0.72, width: 376, color: colors.sun },
  { word: "计算", value: 0.17, width: 172, color: colors.cobalt },
  { word: "模仿", value: 0.06, width: 96, color: colors.cyan },
] as const;

export function PredictionPage() {
  return (
    <SlideFrame
      id="prediction"
      chapter="起点｜把创造改写为概率问题"
      title="先学会预测，再谈创造"
      number={3}
      source="[1] Turing, 1950 · [2] Weizenbaum, 1966 · 候选概率仅为原理示意"
    >
      <Text
        x={64}
        y={176}
        width={120}
        height={66}
        style={{ fontFamily: "display", fontSize: 54, color: colors.cobalt }}
      >
        1950
      </Text>
      <Text
        x={190}
        y={184}
        width={340}
        height={34}
        style={{ fontSize: 23, color: colors.ink }}
      >
        图灵：用行为重述智能
      </Text>
      <Text
        x={190}
        y={230}
        width={340}
        height={94}
        style={{ fontSize: 17, lineHeight: 27, color: colors.muted }}
      >
        《计算机器与智能》提出“模仿游戏”，把“机器能思考吗”转化为对话中能否分辨机器的问题。
      </Text>
      <Rule x={64} y={350} width={454} color={colors.line} />
      <Text
        x={64}
        y={388}
        width={120}
        height={66}
        style={{ fontFamily: "display", fontSize: 54, color: colors.cobalt }}
      >
        1966
      </Text>
      <Text
        x={190}
        y={396}
        width={340}
        height={34}
        style={{ fontSize: 23, color: colors.ink }}
      >
        ELIZA：规则制造对话
      </Text>
      <Text
        x={190}
        y={442}
        width={340}
        height={94}
        style={{ fontSize: 17, lineHeight: 27, color: colors.muted }}
      >
        模式匹配与预设脚本已能模拟谈话。它提醒后来者：流畅的语言行为，不等于真正理解。
      </Text>
      <OutlineBox
        x={584}
        y={176}
        width={632}
        height={402}
        color={colors.cobalt}
      />
      <Text
        x={620}
        y={202}
        width={540}
        height={28}
        style={{ fontSize: 16, color: colors.cobalt }}
      >
        统计语言模型面对的问题
      </Text>
      <Text
        x={620}
        y={252}
        width={540}
        height={62}
        style={{ fontFamily: "display", fontSize: 37, color: colors.ink }}
      >
        “机器如何学会 ______”
      </Text>
      <Text
        x={620}
        y={326}
        width={540}
        height={28}
        style={{ fontSize: 16, color: colors.muted }}
      >
        p（下一个词｜已经出现的词）
      </Text>
      {candidates.map((candidate, index) => {
        const y = 382 + index * 58;
        return (
          <Fragment key={index}>
            <Text
              x={620}
              y={y}
              width={64}
              height={30}
              style={{ fontSize: 18, color: colors.ink }}
            >
              {candidate.word}
            </Text>
            <Rule
              x={696}
              y={y + 11}
              width={400}
              color={colors.mist}
              thickness={10}
            />
            <Rule
              x={696}
              y={y + 11}
              width={candidate.width}
              color={candidate.color}
              thickness={10}
            />
            <Text
              x={1110}
              y={y}
              width={68}
              height={30}
              style={{ fontSize: 16, color: colors.ink, align: "right" }}
            >
              {candidate.value.toFixed(2)}
            </Text>
          </Fragment>
        );
      })}
      <Text
        x={620}
        y={548}
        width={540}
        height={22}
        style={{ fontSize: 12, color: colors.muted }}
      >
        仅列前三个候选；数值为示意，不对应某个真实模型。
      </Text>
      <Text
        x={584}
        y={608}
        width={632}
        height={38}
        style={{
          fontFamily: "display",
          fontSize: 25,
          color: colors.cobalt,
          align: "center",
        }}
      >
        生成的第一课：估计下一个符号的分布。
      </Text>
    </SlideFrame>
  );
}
