import { Chart, Shape, Text } from "@officedk/pptxdk/react";

import { MetricCard, PanelCard, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, fonts } from "../theme.ts";

const benchmarkCategories = [
  "MMLU-Pro",
  "GPQA-Diamond",
  "MATH-500",
  "AIME 2024",
  "Codeforces (%ile)",
  "LiveCodeBench",
] as const;

const r1Scores = [84.0, 71.5, 97.3, 79.8, 96.3, 65.9] as const;
const o1Scores = [84.6, 75.7, 96.4, 79.2, 96.6, 63.4] as const;
const v3Scores = [75.9, 59.1, 90.2, 39.2, 58.7, 40.5] as const;
const sonnetScores = [78.0, 65.0, 78.3, 16.0, 20.3, 38.9] as const;

const highlightRows = [
  {
    domain: "奥数与高阶定理推理",
    metric: "MATH-500: 97.3% · AIME 2024: 79.8%",
    note: "在两项高区分度数学竞赛基准上正面超越 OpenAI o1-1217 (96.4% / 79.2%)",
    tone: "copper" as const,
  },
  {
    domain: "算法竞赛与工程修复",
    metric: "Codeforces: 96.3% 百分位 · SWE-bench: 49.2%",
    note: "算法竞赛等级分达 2,029 Elo，真实软件工程补丁修复率超过 o1-1217 (48.9%)",
    tone: "compute" as const,
  },
  {
    domain: "长上下文与开放式指令",
    metric: "Arena-Hard: 92.3% (R1) / 85.5% (V3)",
    note: "AlpacaEval 2.0 胜率达 87.6% (R1) 与 70.0% (V3)，兼具深度推理与自然表达",
    tone: "comm" as const,
  },
] as const;

/** Slide 13: Frontier benchmark comparison across DeepSeek-R1, DeepSeek-V3, OpenAI o1-1217, and Claude-3.5-Sonnet. */
export function FrontierBenchmarksPage() {
  return (
    <SlideFrame
      id="frontier-benchmarks"
      kicker="04 · 前沿基准全面对齐 / DeepSeek-R1 · DeepSeek-V3 · OpenAI o1-1217 · Claude-3.5-Sonnet"
      title="多维前沿基准评测：开源权重全面对齐第一梯队闭源推理旗舰"
      takeaway="MATH-500 达 97.3%、Codeforces 超越 96.3% 人类选手，V3 领跑非推理开源模型"
      takeawayTone="copper"
      source="数据来源：DeepSeek-V3 Technical Report Table 6 · DeepSeek-R1 Technical Report Table 4"
    >
      {/* Left main: Multi-series grouped column chart */}
      <PanelCard
        x={48}
        y={108}
        width={784}
        height={552}
        title="核心知识、高阶数学与算法编程基准全面对比 (%)"
        subtitle="DeepSeek-R1 (赤铜) 与 OpenAI o1-1217 (冷铅灰) 处于同一性能包络线，DeepSeek-V3 (深钴蓝) 领跑通用基座"
        rightBadge="6 大权威基准"
        tone="copper"
      >
        <Chart
          id="frontier-benchmark-chart"
          legend="top"
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: {
                fontFamily: fonts.body,
                fontSize: 10.5,
                color: colors.ink,
              },
            },
            value: {
              kind: "value",
              min: 0,
              max: 105,
              majorUnit: 25,
              numberFormat: '0"%"',
              gridlines: { color: colors.line, width: 1 },
              line: false,
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
                color: colors.muted,
              },
            },
          }}
          groups={[
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: benchmarkCategories,
              },
              series: [
                {
                  name: "DeepSeek-R1 (开源推理)",
                  values: r1Scores,
                  color: colors.copper,
                },
                {
                  name: "OpenAI o1-1217 (闭源推理)",
                  values: o1Scores,
                  color: colors.baseline,
                },
                {
                  name: "DeepSeek-V3 (开源 MoE 基座)",
                  values: v3Scores,
                  color: colors.compute,
                },
                {
                  name: "Claude-3.5-Sonnet-1022",
                  values: sonnetScores,
                  color: colors.comm,
                },
              ],
              dataLabels: false,
            },
          ]}
          style={{
            position: "absolute",
            left: 64,
            top: 170,
            width: 752,
            height: 416,
            fontFamily: fonts.body,
            fontSize: 10.5,
            color: colors.ink,
          }}
        />
        <Rule x={66} y={596} width={748} />
        <Text
          style={{
            position: "absolute",
            left: 66,
            top: 606,
            width: 748,
            height: 42,
            fontSize: 11,
            lineHeight: 16,
            color: colors.muted,
          }}
        >
          注：AIME 2024、MATH-500、GPQA-Diamond 与 LiveCodeBench
          (2024.08–2025.01) 均为 Pass@1 口径；Codeforces 为人类参赛者百分位排名
          (Percentile)。
        </Text>
      </PanelCard>

      {/* Right top: 2 Hero KPI Cards */}
      <MetricCard
        x={848}
        y={108}
        width={184}
        height={92}
        label="DeepSeek-R1 MATH-500"
        value="97.3%"
        caption="超越 OpenAI o1-1217 (96.4%)"
        tone="copper"
      />
      <MetricCard
        x={1048}
        y={108}
        width={184}
        height={92}
        label="Codeforces 算法评级"
        value="96.3%ile"
        caption="2,029 Elo · 国际特级大师线"
        tone="compute"
      />

      {/* Right bottom: Detailed capability breakdown panel */}
      <PanelCard
        x={848}
        y={214}
        width={384}
        height={446}
        title="关键能力象限优势拆解"
        subtitle="从基座预训练效率向后训练思维链推理复利传导"
        rightBadge="全栈协同收益"
        tone="compute"
      >
        {highlightRows.map((item, idx) => {
          const boxY = 278 + idx * 120;
          return (
            <Shape
              key={item.domain}
              style={{
                position: "absolute",
                left: 866,
                top: boxY,
                width: 348,
                height: 106,
                background: colors.paper,
                border: { color: colors.line, width: 1 },
              }}
            />
          );
        })}
        {highlightRows.map((item, idx) => {
          const boxY = 278 + idx * 120;
          const accent =
            item.tone === "copper"
              ? colors.copper
              : item.tone === "comm"
                ? colors.comm
                : colors.compute;
          return (
            <Shape
              key={`bar-${idx}`}
              style={{
                position: "absolute",
                left: 866,
                top: boxY,
                width: 4,
                height: 106,
                background: accent,
              }}
            />
          );
        })}
        {highlightRows.map((item, idx) => {
          const boxY = 278 + idx * 120;
          return (
            <Text
              key={`dom-${idx}`}
              style={{
                position: "absolute",
                left: 882,
                top: boxY + 10,
                width: 318,
                height: 20,
                fontSize: 12.5,
                color: colors.steel,
                textWrap: "nowrap",
              }}
            >
              {item.domain}
            </Text>
          );
        })}
        {highlightRows.map((item, idx) => {
          const boxY = 278 + idx * 120;
          const textColor =
            item.tone === "copper"
              ? colors.copperText
              : item.tone === "comm"
                ? colors.commText
                : colors.compute;
          return (
            <Text
              key={`met-${idx}`}
              style={{
                position: "absolute",
                left: 882,
                top: boxY + 34,
                width: 318,
                height: 20,
                fontSize: 11.5,
                color: textColor,
                textWrap: "nowrap",
              }}
            >
              {item.metric}
            </Text>
          );
        })}
        {highlightRows.map((item, idx) => {
          const boxY = 278 + idx * 120;
          return (
            <Text
              key={`note-${idx}`}
              style={{
                position: "absolute",
                left: 882,
                top: boxY + 58,
                width: 318,
                height: 40,
                fontSize: 10.5,
                lineHeight: 15,
                color: colors.muted,
              }}
            >
              {item.note}
            </Text>
          );
        })}
      </PanelCard>
    </SlideFrame>
  );
}
