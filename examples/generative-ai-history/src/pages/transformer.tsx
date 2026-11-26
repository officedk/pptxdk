import { Slide, path, Shape, Text } from "@officedk/pptxdk/react";
import { Fragment } from "react";

import { OutlineBox } from "../components/primitives.tsx";
import { contentLayout } from "../templates/layouts.ts";
import { colors, transparent } from "../theme.ts";

const tokens = ["生", "成", "式", "人", "工", "智", "能"];
const weights = [0.05, 0.08, 0.12, 0.09, 0.16, 0.36, 0.14];

export function TransformerPage() {
  const queryX = 490;
  const queryY = 238;
  return (
    <Slide
      id="transformer"
      layout={contentLayout}
      placeholders={{
        chapter: "架构｜Attention Is All You Need",
        title: "2017｜注意力重写了序列建模",
        source: "[5] Vaswani et al., 2017 · 注意力连线粗细仅为原理示意",
      }}
    >
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 164,
          width: 780,
          height: 40,
          fontSize: 18,
          color: colors.muted,
        }}
      >
        以“智”为查询：模型按相关性读取序列中的其他位置。
      </Text>
      {tokens.map((token, index) => {
        const x = 70 + index * 108;
        const targetX = x + 29;
        return (
          <Shape
            key={token}
            geometry={path(`M ${queryX - 32} 120 L ${targetX - 64} 303`, {
              viewBox: [830, 330],
              fill: false,
            })}
            style={{
              position: "absolute",
              left: 64,
              top: 174,
              width: 830,
              height: 330,
              background: transparent,
              border: {
                color: index === 5 ? colors.sun : colors.cobalt,
                width: 0.6 + weights[index]! * 8,
              },
              opacity: index === 5 ? 0.95 : 0.42,
            }}
          />
        );
      })}
      <Shape
        style={{
          position: "absolute",
          left: queryX,
          top: queryY,
          width: 64,
          height: 64,
          background: colors.sun,
          rotate: 45,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: queryX,
          top: queryY + 4,
          width: 64,
          height: 56,
          fontFamily: "display",
          fontSize: 31,
          color: colors.indigo,
          align: "center",
        }}
      >
        智
      </Text>
      <Text
        style={{
          position: "absolute",
          left: queryX - 20,
          top: queryY - 40,
          width: 104,
          height: 28,
          fontSize: 14,
          color: colors.cobalt,
          align: "center",
        }}
      >
        当前查询
      </Text>
      {tokens.map((token, index) => {
        const x = 70 + index * 108;
        return (
          <Fragment key={index}>
            <OutlineBox
              x={x}
              y={448}
              width={58}
              height={58}
              color={index === 5 ? colors.sun : colors.cobalt}
            />
            <Text
              style={{
                position: "absolute",
                left: x,
                top: 455,
                width: 58,
                height: 44,
                fontFamily: "display",
                fontSize: 28,
                color: colors.ink,
                align: "center",
              }}
            >
              {token}
            </Text>
            <Text
              style={{
                position: "absolute",
                left: x - 8,
                top: 516,
                width: 74,
                height: 24,
                fontSize: 13,
                color: colors.muted,
                align: "center",
              }}
            >
              {weights[index]!.toFixed(2)}
            </Text>
          </Fragment>
        );
      })}
      <Text
        style={{
          position: "absolute",
          left: 916,
          top: 190,
          width: 294,
          height: 44,
          fontFamily: "display",
          fontSize: 27,
          color: colors.cobalt,
        }}
      >
        关键变化
      </Text>
      {[
        ["并行", "摆脱必须逐词递归的计算顺序"],
        ["远程依赖", "让任意位置直接交换信息"],
        ["可堆叠", "构成后续大模型的基本模块"],
      ].map(([title, body], index) => {
        const y = 260 + index * 106;
        return (
          <Fragment key={index}>
            <Text
              style={{
                position: "absolute",
                left: 916,
                top: y,
                width: 294,
                height: 32,
                fontSize: 20,
                color: colors.ink,
              }}
            >
              {title}
            </Text>
            <Text
              style={{
                position: "absolute",
                left: 916,
                top: y + 40,
                width: 294,
                height: 52,
                fontSize: 15,
                lineHeight: 23,
                color: colors.muted,
              }}
            >
              {body}
            </Text>
          </Fragment>
        );
      })}
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 596,
          width: 1148,
          height: 48,
          fontFamily: "display",
          fontSize: 25,
          color: colors.ink,
          align: "center",
        }}
      >
        最初为机器翻译提出，随后成为通用生成模型的主干。
      </Text>
    </Slide>
  );
}
