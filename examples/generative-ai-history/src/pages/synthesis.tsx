import { Slide, Text } from "pptxdk/react";
import { Fragment } from "react";

import { Arrow, Rule } from "../components/primitives.tsx";
import { summaryLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

const rows = [
  {
    name: "语言模型",
    formula: "p（下一个词｜上下文）",
    note: "把序列补全变成可学习的分布",
  },
  {
    name: "VAE",
    formula: "潜变量 z → 解码样本",
    note: "在结构化的潜空间中采样",
  },
  {
    name: "GAN",
    formula: "生成器 ↔ 判别器",
    note: "用对抗目标逼近真实数据分布",
  },
  { name: "扩散模型", formula: "加噪 ↔ 去噪", note: "学习逐步逆转破坏过程" },
  {
    name: "人类反馈",
    formula: "候选回答 → 偏好选择",
    note: "让输出分布更接近人的意图",
  },
] as const;

export function SynthesisPage() {
  return (
    <Slide
      id="synthesis"
      layout={summaryLayout}
      placeholders={{
        chapter: "收束｜技术不同，问题相通",
        title: "七十四年，一条未变的主线",
        source:
          "概念归纳：各方法的训练目标与采样机制并不相同，此处强调共同的分布视角",
      }}
    >
      <Text
        x={64}
        y={154}
        width={1148}
        height={42}
        style={{
          fontFamily: "display",
          fontSize: 29,
          color: colors.paperBright,
          align: "center",
        }}
      >
        从数据中学习结构，再用条件与反馈引导生成
      </Text>
      {[
        { x: 300, label: "数据" },
        { x: 516, label: "分布" },
        { x: 732, label: "采样" },
        { x: 948, label: "新样本" },
      ].map((item, index) => (
        <Fragment key={index}>
          <Text
            x={item.x}
            y={214}
            width={120}
            height={44}
            style={{
              fontSize: 22,
              color: index === 3 ? colors.sun : colors.cyan,
              align: "center",
            }}
          >
            {item.label}
          </Text>
          {index < 3 && (
            <Arrow x={item.x + 138} y={226} width={58} color={colors.muted} />
          )}
        </Fragment>
      ))}
      <Rule x={64} y={286} width={1152} color={colors.cobalt} />
      {rows.map((row, index) => {
        const y = 306 + index * 61;
        return (
          <Fragment key={index}>
            <Text
              x={64}
              y={y}
              width={190}
              height={34}
              style={{
                fontFamily: "display",
                fontSize: 24,
                color: colors.paperBright,
              }}
            >
              {row.name}
            </Text>
            <Text
              x={286}
              y={y + 3}
              width={340}
              height={32}
              style={{ fontSize: 17, color: colors.sun }}
            >
              {row.formula}
            </Text>
            <Text
              x={670}
              y={y + 4}
              width={542}
              height={32}
              style={{ fontSize: 16, color: colors.mist }}
            >
              {row.note}
            </Text>
            {index < rows.length - 1 && (
              <Rule
                x={64}
                y={y + 48}
                width={1152}
                color={{ color: colors.cobalt, alpha: 0.45 }}
              />
            )}
          </Fragment>
        );
      })}
      <Text
        x={64}
        y={626}
        width={1152}
        height={30}
        style={{ fontSize: 16, color: colors.cyan, align: "center" }}
      >
        媒介一直在变：从符号序列，到像素、声音与行动轨迹。
      </Text>
    </Slide>
  );
}
