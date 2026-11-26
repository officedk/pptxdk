import { Slide, Text } from "@officedk/pptxdk/react";
import { Fragment } from "react";

import { Circle, Rule } from "../components/primitives.tsx";
import { contentLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

const moments = [
  {
    year: "1950",
    title: "图灵之问",
    detail: "用可观察行为重述智能",
    above: true,
  },
  {
    year: "1966",
    title: "ELIZA",
    detail: "规则模拟自然语言对话",
    above: false,
  },
  {
    year: "2013–14",
    title: "VAE / GAN",
    detail: "生成模型迎来新架构",
    above: true,
  },
  {
    year: "2017",
    title: "Transformer",
    detail: "注意力成为序列主干",
    above: false,
  },
  {
    year: "2020",
    title: "GPT-3 / DDPM",
    detail: "规模与扩散同时推进",
    above: true,
  },
  {
    year: "2022",
    title: "ChatGPT",
    detail: "对齐与对话走向大众",
    above: false,
  },
  {
    year: "2024",
    title: "多模态与推理",
    detail: "实时交互，延长思考",
    above: true,
  },
] as const;

const eras = [
  { x: 64, width: 250, label: "规则与概率", years: "1950—2012" },
  { x: 326, width: 278, label: "生成架构", years: "2013—2017" },
  { x: 616, width: 300, label: "规模与对齐", years: "2018—2022" },
  { x: 928, width: 288, label: "系统与多模态", years: "2023—2024" },
] as const;

export function TimelinePage() {
  return (
    <Slide
      id="timeline"
      layout={contentLayout}
      placeholders={{
        chapter: "导览｜七十四年的四次转向",
        title: "这不是一次突然的降临",
        source: "关键节点对应原始论文与官方发布；完整出处见文末参考资料",
      }}
    >
      <Rule x={96} y={384} width={1080} color={colors.cobalt} thickness={2} />
      {moments.map((moment, index) => {
        const x = 92 + index * 178;
        const yearY = moment.above ? 208 : 434;
        const titleY = moment.above ? 260 : 486;
        const detailY = moment.above ? 298 : 522;
        return (
          <Fragment key={index}>
            <Rule
              x={x + 7}
              y={moment.above ? 322 : 392}
              width={1}
              color={colors.line}
              thickness={moment.above ? 50 : 34}
            />
            <Circle
              x={x}
              y={377}
              diameter={16}
              fill={index === 6 ? colors.sun : colors.paper}
              stroke={index === 6 ? colors.sun : colors.cobalt}
              strokeWidth={2}
            />
            <Text
              style={{
                position: "absolute",
                left: Math.min(x - 26, 1052),
                top: yearY,
                width: 166,
                height: 48,
                fontFamily: "display",
                fontSize: moment.year.length > 4 ? 30 : 36,
                color: colors.ink,
                textWrap: "nowrap",
              }}
            >
              {moment.year}
            </Text>
            <Text
              style={{
                position: "absolute",
                left: Math.min(x - 26, 1052),
                top: titleY,
                width: 164,
                height: 30,
                fontSize: 18,
                color: colors.cobalt,
                textWrap: "nowrap",
              }}
            >
              {moment.title}
            </Text>
            <Text
              style={{
                position: "absolute",
                left: Math.min(x - 26, 1052),
                top: detailY,
                width: 162,
                height: 54,
                fontSize: 14,
                lineHeight: 21,
                color: colors.muted,
              }}
            >
              {moment.detail}
            </Text>
          </Fragment>
        );
      })}
      {eras.map((era, index) => (
        <Fragment key={index}>
          <Rule
            x={era.x}
            y={602}
            width={era.width}
            color={index === 3 ? colors.sun : colors.cyan}
            thickness={4}
          />
          <Text
            style={{
              position: "absolute",
              left: era.x,
              top: 616,
              width: era.width,
              height: 24,
              fontSize: 16,
              color: colors.ink,
            }}
          >
            {era.label}
          </Text>
          <Text
            style={{
              position: "absolute",
              left: era.x,
              top: 642,
              width: era.width,
              height: 20,
              fontSize: 12,
              color: colors.muted,
            }}
          >
            {era.years}
          </Text>
        </Fragment>
      ))}
    </Slide>
  );
}
