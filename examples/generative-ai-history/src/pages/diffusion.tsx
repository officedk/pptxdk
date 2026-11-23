import { Slide, path, preset, Shape, Text } from "pptxdk/react";
import { Fragment } from "react";

import { OutlineBox, Rule } from "../components/primitives.tsx";
import { contentLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

function random(index: number, salt: number) {
  const value = Math.sin(index * 47.13 + salt * 19.79) * 15731.743;
  return value - Math.floor(value);
}

function Landscape({
  x,
  stage,
}: {
  readonly x: number;
  readonly stage: number;
}) {
  const clarity = stage / 4;
  return (
    <Fragment key={x}>
      <Shape x={x} y={220} width={190} height={206} fill={colors.paperBright} />
      {stage > 0 && (
        <Fragment key={x}>
          <Shape
            geometry={preset("ellipse")}
            x={x + 126}
            y={252}
            width={30}
            height={30}
            fill={{ color: colors.sun, alpha: clarity }}
          />
          <Shape
            geometry={path(
              "M 0 150 L 46 80 L 78 120 L 116 52 L 190 150 L 190 206 L 0 206 Z",
              {
                viewBox: [190, 206],
              },
            )}
            x={x}
            y={220}
            width={190}
            height={206}
            fill={{ color: colors.cobalt, alpha: clarity }}
          />
          <Shape
            geometry={path(
              "M 0 168 L 58 108 L 98 150 L 142 96 L 190 142 L 190 206 L 0 206 Z",
              {
                viewBox: [190, 206],
              },
            )}
            x={x}
            y={220}
            width={190}
            height={206}
            fill={{ color: colors.ink, alpha: clarity * 0.85 }}
          />
        </Fragment>
      )}
      {Array.from({ length: 64 }, (_, index) => {
        const size = 4 + random(index, 3) * 10;
        return (
          <Shape
            key={index}
            x={x + random(index, 1) * (190 - size)}
            y={220 + random(index, 2) * (206 - size)}
            width={size}
            height={size}
            fill={{
              color:
                index % 7 === 0
                  ? colors.sun
                  : index % 3 === 0
                    ? colors.cyan
                    : colors.cobalt,
              alpha: (1 - clarity) * (0.25 + random(index, 4) * 0.7),
            }}
          />
        );
      })}
      <OutlineBox
        x={x}
        y={220}
        width={190}
        height={206}
        color={stage === 4 ? colors.sun : colors.line}
      />
    </Fragment>
  );
}

const milestones = [
  {
    year: "2020",
    title: "DDPM",
    body: "用概率模型学习逐步逆转加噪过程，扩散生成质量显著提升。",
    source: "[9]",
  },
  {
    year: "2021",
    title: "CLIP",
    body: "从自然语言监督中学习视觉表示，把文字与图像联系起来。",
    source: "[10]",
  },
  {
    year: "2022",
    title: "Latent Diffusion",
    body: "把扩散过程放进潜空间，降低高分辨率图像合成的计算负担。",
    source: "[11]",
  },
] as const;

export function DiffusionPage() {
  return (
    <Slide
      id="diffusion"
      layout={contentLayout}
      placeholders={{
        chapter: "媒介｜文字成为图像的条件",
        title: "2020–2022｜图像从噪声中显影",
        source:
          "[9] Ho et al., 2020 · [10] Radford et al., 2021 · [11] Rombach et al., 2022 · 图像为原理示意",
      }}
    >
      <Text
        x={64}
        y={158}
        width={1148}
        height={38}
        style={{ fontSize: 18, color: colors.muted }}
      >
        从随机噪声出发，模型反复估计并移除噪声，让满足文字条件的结构逐步浮现。
      </Text>
      {[0, 1, 2, 3, 4].map((stage) => {
        const x = 64 + stage * 232;
        return (
          <Fragment key={x}>
            <Landscape x={x} stage={stage} />
            <Text
              x={x}
              y={440}
              width={190}
              height={26}
              style={{ fontSize: 14, color: colors.cobalt, align: "center" }}
            >
              {
                ["纯噪声", "结构浮现", "轮廓形成", "细节收敛", "生成图像"][
                  stage
                ]
              }
            </Text>
            <Text
              x={x}
              y={468}
              width={190}
              height={22}
              style={{ fontSize: 12, color: colors.muted, align: "center" }}
            >
              {["t = T", "t = ¾T", "t = ½T", "t = ¼T", "t = 0"][stage]}
            </Text>
          </Fragment>
        );
      })}
      <Rule x={64} y={510} width={1152} color={colors.line} />
      {milestones.map((milestone, index) => {
        const x = 64 + index * 390;
        return (
          <Fragment key={x}>
            <Text
              x={x}
              y={534}
              width={86}
              height={42}
              style={{
                fontFamily: "display",
                fontSize: 30,
                color: colors.cobalt,
              }}
            >
              {milestone.year}
            </Text>
            <Text
              x={x + 92}
              y={538}
              width={250}
              height={30}
              style={{ fontSize: 19, color: colors.ink }}
            >
              {milestone.title} {milestone.source}
            </Text>
            <Text
              x={x + 92}
              y={576}
              width={282}
              height={66}
              style={{ fontSize: 14, lineHeight: 22, color: colors.muted }}
            >
              {milestone.body}
            </Text>
          </Fragment>
        );
      })}
    </Slide>
  );
}
