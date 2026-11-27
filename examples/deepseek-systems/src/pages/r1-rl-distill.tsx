import {
  Chart,
  fade,
  MotionScope,
  Shape,
  Step,
  Text,
  useMotionScope,
} from "@officedk/pptxdk/react";

import { PanelCard, SemanticTag } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, fonts } from "../theme.ts";

const r1PipelineStages = [
  {
    step: "Stage 1 · 冷启动微调",
    tag: "数千条长 CoT 样本",
    tone: "compute" as const,
    desc: "在 DeepSeek-V3-Base 上使用精选长思维链样本微调，克服 R1-Zero 可读性差与多语言混杂问题，建立规范反思格式。",
  },
  {
    step: "Stage 2 · 推理强化学习",
    tag: "GRPO + 语言一致性奖励",
    tone: "copper" as const,
    desc: "在数学、代码、科学逻辑推理任务上执行大规模 GRPO 强化学习，引入语言一致性奖励信号直至收敛。",
  },
  {
    step: "Stage 3 · 拒绝采样与 SFT",
    tag: "600K 推理 + 200K 通用",
    tone: "comm" as const,
    desc: "由 RL 检查点执行拒绝采样并融合通用写作、问答与自我认知数据（共 80 万条），重新微调 V3-Base 模型。",
  },
  {
    step: "Stage 4 · 全场景强化学习",
    tag: "推理规则 + 人类偏好对齐",
    tone: "compute" as const,
    desc: "结合基于规则的推理奖励与面向有用性 (Helpfulness) 及安全性 (Harmlessness) 的偏好奖励，产出最终 DeepSeek-R1。",
  },
] as const;

const distillCategories = [
  "GPT-4o-0513",
  "OpenAI o1-mini",
  "QwQ-32B-Preview",
  "R1-Distill-Qwen-7B",
  "R1-Distill-Qwen-14B",
  "R1-Distill-Qwen-32B",
  "R1-Distill-Llama-70B",
] as const;

const aime2024Scores = [9.3, 63.6, 50.0, 55.5, 69.7, 72.6, 70.0] as const;
const math500Scores = [74.6, 90.0, 90.6, 92.8, 93.9, 94.3, 94.5] as const;

/** Slide 12: DeepSeek-R1 GRPO reinforcement learning pipeline and distillation transfer. */
export function R1RlDistillPage() {
  const stageScope = useMotionScope();

  return (
    <SlideFrame
      id="r1-rl-distill"
      kicker="04 · 强化学习推理涌现与蒸馏迁移 / DeepSeek-R1-Zero · GRPO · Distillation"
      title="GRPO 强化学习与四阶段演化：无需价值网络激发自发反思与小模型蒸馏"
      takeaway="R1-Distill-32B 在 AIME 2024 达 72.6%，全面超越 OpenAI o1-mini (63.6%)"
      takeawayTone="copper"
      source="数据来源：DeepSeek-R1 Technical Report (arXiv:2501.12948) · DeepSeek-R1 开源仓库评测表"
    >
      {/* Top Left: GRPO Algorithm without Critic Model */}
      <PanelCard
        x={48}
        y={108}
        width={384}
        height={244}
        title="GRPO 组相对策略优化机制"
        subtitle="省去与策略模型同规模的 Critic 价值网络，训练显存与计算成本减半"
        rightBadge="无 Critic 架构"
        tone="copper"
      >
        <Shape
          style={{
            position: "absolute",
            left: 66,
            top: 170,
            width: 348,
            height: 76,
            background: colors.paper,
            border: { color: colors.copper, width: 1 },
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 80,
            top: 178,
            width: 320,
            height: 20,
            fontSize: 12,
            color: colors.copperText,
            textWrap: "nowrap",
          }}
        >
          组内采样相对优势估计 (Group Relative Advantage)：
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 80,
            top: 202,
            width: 320,
            height: 24,
            fontFamily: fonts.metric,
            fontSize: 14,
            fontWeight: 700,
            color: colors.ink,
            textWrap: "nowrap",
          }}
        >
          A_i = (r_i - mean(r_1..G)) / std(r_1..G)
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 80,
            top: 224,
            width: 320,
            height: 18,
            fontSize: 10.5,
            color: colors.muted,
            textWrap: "nowrap",
          }}
        >
          对每个问题采样一组输出 {`{o_1, ..., o_G}`}，以规则奖励均值作基线
        </Text>

        <Text
          style={{
            position: "absolute",
            left: 66,
            top: 256,
            width: 348,
            height: 84,
            fontSize: 11,
            lineHeight: 16,
            color: colors.steel,
          }}
        >
          • DeepSeek-R1-Zero 纯强化学习验证：跳过监督微调直接对 V3-Base 施以
          GRPO，AIME 2024 准确率从 15.6% 跃升至 71.0%（多数投票达
          86.7%），自发涌现长思维链、回溯验证与“顿悟时刻 (Aha Moment)”。
        </Text>
      </PanelCard>

      {/* Top Right: 4-Stage DeepSeek-R1 Pipeline with Step + MotionScope */}
      <PanelCard
        x={448}
        y={108}
        width={784}
        height={244}
        title="DeepSeek-R1 四阶段训练流水线"
        subtitle="冷启动数据锚定可读格式，两轮大规模强化学习结合拒绝采样实现推理与通用能力兼备"
        rightBadge="点击步进展示四阶段"
        tone="compute"
      >
        <Step>
          <MotionScope scope={stageScope} mode="stagger" each={120}>
            {r1PipelineStages.map((st, idx) => {
              const cardX = 466 + idx * 188;
              const accent =
                st.tone === "copper"
                  ? colors.copper
                  : st.tone === "comm"
                    ? colors.comm
                    : colors.compute;
              return (
                <Shape
                  key={st.step}
                  id={`r1-stage-card-${idx}`}
                  enter={fade({ duration: 340 })}
                  style={{
                    position: "absolute",
                    left: cardX,
                    top: 170,
                    width: 178,
                    height: 166,
                    background: colors.paper,
                    border: { color: accent, width: 1 },
                  }}
                />
              );
            })}
            {r1PipelineStages.map((st, idx) => {
              const cardX = 466 + idx * 188;
              const accent =
                st.tone === "copper"
                  ? colors.copper
                  : st.tone === "comm"
                    ? colors.comm
                    : colors.compute;
              return (
                <Shape
                  key={`top-${st.step}`}
                  enter={fade({ duration: 340 })}
                  style={{
                    position: "absolute",
                    left: cardX,
                    top: 170,
                    width: 178,
                    height: 4,
                    background: accent,
                  }}
                />
              );
            })}
            {r1PipelineStages.map((st, idx) => {
              const cardX = 466 + idx * 188;
              return (
                <Text
                  key={`title-${st.step}`}
                  enter={fade({ duration: 340 })}
                  style={{
                    position: "absolute",
                    left: cardX + 10,
                    top: 182,
                    width: 158,
                    height: 20,
                    fontSize: 12,
                    color: colors.ink,
                    textWrap: "nowrap",
                  }}
                >
                  {st.step}
                </Text>
              );
            })}
            {r1PipelineStages.map((st, idx) => {
              const cardX = 466 + idx * 188;
              return (
                <SemanticTag
                  key={`tag-${st.step}`}
                  x={cardX + 10}
                  y={206}
                  width={158}
                  height={22}
                  text={st.tag}
                  tone={st.tone}
                />
              );
            })}
            {r1PipelineStages.map((st, idx) => {
              const cardX = 466 + idx * 188;
              return (
                <Text
                  key={`desc-${st.step}`}
                  enter={fade({ duration: 340 })}
                  style={{
                    position: "absolute",
                    left: cardX + 10,
                    top: 236,
                    width: 158,
                    height: 92,
                    fontSize: 10.5,
                    lineHeight: 15,
                    color: colors.muted,
                  }}
                >
                  {st.desc}
                </Text>
              );
            })}
          </MotionScope>
        </Step>
      </PanelCard>

      {/* Bottom: Distillation Benchmark Grouped Column Chart + Insight */}
      <PanelCard
        x={48}
        y={366}
        width={1184}
        height={294}
        title="DeepSeek-R1 推理轨迹蒸馏至稠密小模型 (Qwen / Llama) 评测对比 (%)"
        subtitle="直接利用 R1 产出的 80 万高质量推理样本对开源稠密模型做 SFT，显著超越在小模型上直接运行大规模 RL"
        rightBadge="蒸馏迁移能力跃升"
        tone="copper"
      >
        <Chart
          id="r1-distill-chart"
          legend="top"
          axes={{
            category: {
              kind: "category",
              line: { color: colors.steel, width: 1 },
              style: {
                fontFamily: fonts.body,
                fontSize: 10,
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
                categories: distillCategories,
              },
              series: [
                {
                  name: "AIME 2024 Pass@1 (%)",
                  values: aime2024Scores,
                  color: colors.copper,
                },
                {
                  name: "MATH-500 Pass@1 (%)",
                  values: math500Scores,
                  color: colors.compute,
                },
              ],
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: "0.0",
                style: {
                  fontFamily: fonts.body,
                  fontSize: 9.5,
                  color: colors.ink,
                },
              },
            },
          ]}
          style={{
            position: "absolute",
            left: 64,
            top: 426,
            width: 836,
            height: 220,
            fontFamily: fonts.body,
            fontSize: 10,
            color: colors.ink,
          }}
        />

        {/* Right callout card inside bottom panel */}
        <Shape
          style={{
            position: "absolute",
            left: 916,
            top: 430,
            width: 296,
            height: 212,
            background: colors.paper,
            border: { color: colors.line, width: 1 },
          }}
        />
        <Shape
          style={{
            position: "absolute",
            left: 916,
            top: 430,
            width: 4,
            height: 212,
            background: colors.copper,
          }}
        />
        <Text
          style={{
            position: "absolute",
            left: 934,
            top: 442,
            width: 264,
            height: 22,
            fontSize: 13,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          核心发现：大模型探索 vs 小模型蒸馏
        </Text>
        <Text
          style={{
            position: "absolute",
            left: 934,
            top: 470,
            width: 264,
            height: 160,
            fontSize: 11,
            lineHeight: 16.5,
            color: colors.muted,
          }}
        >
          1. 7B 超越 32B 预览版：R1-Distill-Qwen-7B 在 AIME 2024 取得
          55.5%，超越 QwQ-32B-Preview (50.0%)。{"\n"}
          2. 14B/32B 全面反超 o1-mini：Distill-32B 在 AIME 达 72.6%、MATH-500 达
          94.3%。{"\n"}
          3. 蒸馏优于小模型从零 RL：对 Qwen-32B-Base 直接做逾万步 RL 仅达 47.0%
          (AIME)，远低于从 R1 蒸馏的 72.6%。
        </Text>
      </PanelCard>
    </SlideFrame>
  );
}
