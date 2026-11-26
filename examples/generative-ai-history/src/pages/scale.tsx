import { Slide, Chart, Text } from "@officedk/pptxdk/react";

import { Rule } from "../components/primitives.tsx";
import { contentLayout } from "../templates/layouts.ts";
import { colors } from "../theme.ts";

export function ScalePage() {
  return (
    <Slide
      id="scale"
      layout={contentLayout}
      placeholders={{
        chapter: "规模｜GPT 系列的三次跨越",
        title: "2018–2020｜规模成为方法",
        source:
          "[6] Radford et al., 2018 · [7] Radford et al., 2019 · [8] Brown et al., 2020",
      }}
    >
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 158,
          width: 740,
          height: 34,
          fontSize: 17,
          color: colors.muted,
        }}
      >
        参数量（十亿，线性刻度）
      </Text>
      <Chart
        id="gpt-parameters"
        style={{
          position: "absolute",
          left: 64,
          top: 196,
          width: 742,
          height: 394,
          fontFamily: "body",
          fontSize: 14,
          color: colors.ink,
        }}
        legend={false}
        plotArea={{ left: 70, top: 24, width: 620, height: 292 }}
        axes={{
          category: {
            kind: "category",
            line: { color: colors.cobalt, width: 1 },
            style: { fontFamily: "body", fontSize: 14, color: colors.ink },
          },
          value: {
            kind: "value",
            min: 0,
            max: 200,
            majorUnit: 50,
            numberFormat: '0"B"',
            gridlines: { color: colors.line, width: 1 },
            line: false,
            style: { fontFamily: "body", fontSize: 12, color: colors.muted },
          },
        }}
        groups={[
          {
            kind: "column",
            data: {
              kind: "shared-categories",
              categories: ["GPT-1 · 2018", "GPT-2 · 2019", "GPT-3 · 2020"],
            },
            series: [
              {
                name: "参数量（十亿）",
                values: [0.117, 1.5, 175],
                color: colors.cobalt,
                points: {
                  0: { color: colors.cyan, labelColor: colors.ink },
                  1: { color: colors.cobalt, labelColor: colors.ink },
                  2: { color: colors.sun, labelColor: colors.ink },
                },
              },
            ],
            gapWidth: 76,
            dataLabels: {
              value: true,
              position: "outsideEnd",
              numberFormat: '0.0##"B"',
              style: { fontFamily: "body", fontSize: 15, color: colors.ink },
            },
          },
        ]}
      />
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 604,
          width: 742,
          height: 38,
          fontSize: 13,
          lineHeight: 20,
          color: colors.muted,
        }}
      >
        GPT-1 约 117M（见 GPT-2 论文表 2），GPT-2 约 1.5B，GPT-3 为
        175B。柱高采用线性刻度。
      </Text>
      <Rule x={856} y={184} width={1} color={colors.line} thickness={434} />
      <Text
        style={{
          position: "absolute",
          left: 900,
          top: 196,
          width: 310,
          height: 82,
          fontFamily: "display",
          fontSize: 60,
          color: colors.cobalt,
        }}
      >
        约1500×
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 902,
          top: 284,
          width: 306,
          height: 72,
          fontSize: 18,
          lineHeight: 29,
          color: colors.ink,
        }}
      >
        两年间，GPT 系列参数量从 1.17 亿增至 1750 亿。
      </Text>
      <Rule x={902} y={386} width={66} color={colors.sun} thickness={5} />
      <Text
        style={{
          position: "absolute",
          left: 902,
          top: 420,
          width: 306,
          height: 36,
          fontFamily: "display",
          fontSize: 26,
          color: colors.ink,
        }}
      >
        规模之外的新行为
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 902,
          top: 474,
          width: 306,
          height: 126,
          fontSize: 17,
          lineHeight: 28,
          color: colors.muted,
        }}
      >
        GPT-3 展示了上下文内学习：不更新参数，也能依照提示中的示例完成新任务。
      </Text>
    </Slide>
  );
}
