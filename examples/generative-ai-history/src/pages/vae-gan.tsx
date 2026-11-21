import { path, preset } from "pptxdk";
import type { ColorValue } from "pptxdk";
import { P, Shape, Text } from "pptxdk/react";

import { Circle, Rule } from "../components/primitives.tsx";
import { SlideFrame } from "../templates/slide-frame.tsx";
import { colors, transparent } from "../theme.ts";

function FlowLabel({
  x,
  y,
  text,
  width = 116,
  kind = "process",
}: {
  readonly x: number;
  readonly y: number;
  readonly text: string;
  readonly width?: number;
  readonly kind?: "process" | "data";
}) {
  return (
    <Shape
      geometry={preset(
        kind === "data" ? "flowChartInputOutput" : "flowChartProcess",
      )}
      x={x}
      y={y}
      width={width}
      height={56}
      fill={transparent}
      style={{ borderColor: colors.cobalt, borderWidth: 1 }}
      text={text}
      textStyle={{
        fontFamily: "body",
        fontSize: 17,
        color: colors.ink,
        align: "center",
        verticalAlign: "middle",
      }}
    />
  );
}

function FlowArrow({
  x,
  y,
  width = 44,
  color = colors.cobalt,
}: {
  readonly x: number;
  readonly y: number;
  readonly width?: number;
  readonly color?: ColorValue;
}) {
  return (
    <Shape
      geometry={preset("rightArrow")}
      x={x}
      y={y}
      width={width}
      height={16}
      fill={color}
    />
  );
}

export function VaeGanPage() {
  return (
    <SlideFrame
      id="vae-gan"
      chapter="架构｜两条不同的生成路线"
      title="2013–2014｜生成有了自己的模型"
      number={4}
      source="[3] Kingma & Welling, 2013 · [4] Goodfellow et al., 2014"
    >
      <Rule x={640} y={174} width={1} color={colors.line} thickness={454} />
      <Text
        x={64}
        y={174}
        width={180}
        height={62}
        style={{ fontFamily: "display", fontSize: 46, color: colors.cobalt }}
      >
        VAE
      </Text>
      <Text
        x={210}
        y={186}
        width={360}
        height={34}
        style={{ fontSize: 20, color: colors.ink }}
      >
        变分自编码器 · 2013
      </Text>
      <Text
        x={64}
        y={246}
        width={540}
        height={62}
        style={{ fontSize: 18, lineHeight: 29, color: colors.muted }}
      >
        <P>把数据编码为潜变量分布，</P>
        <P>再从潜空间采样并解码为新样本。</P>
      </Text>
      <FlowLabel kind="data" x={70} y={342} text="样本 x" width={100} />
      <FlowArrow x={182} y={360} />
      <FlowLabel x={238} y={342} text="编码器" width={100} />
      <FlowArrow x={350} y={360} />
      <FlowLabel kind="data" x={406} y={342} text="μ，σ" width={100} />
      <Circle x={238} y={452} diameter={20} fill={colors.sun} />
      <Text
        x={274}
        y={440}
        width={280}
        height={42}
        style={{ fontFamily: "display", fontSize: 25, color: colors.ink }}
      >
        在潜空间采样 z
      </Text>
      <Rule x={248} y={410} width={1} color={colors.cobalt} thickness={36} />
      <FlowLabel kind="data" x={70} y={520} text="潜变量 z" width={112} />
      <FlowArrow x={194} y={538} />
      <FlowLabel x={250} y={520} text="解码器" width={100} />
      <FlowArrow x={362} y={538} />
      <FlowLabel kind="data" x={418} y={520} text="新样本 x′" width={128} />
      <Text
        x={692}
        y={174}
        width={180}
        height={62}
        style={{ fontFamily: "display", fontSize: 46, color: colors.cobalt }}
      >
        GAN
      </Text>
      <Text
        x={852}
        y={186}
        width={330}
        height={34}
        style={{ fontSize: 20, color: colors.ink }}
      >
        生成对抗网络 · 2014
      </Text>
      <Text
        x={692}
        y={246}
        width={520}
        height={62}
        style={{ fontSize: 18, lineHeight: 29, color: colors.muted }}
      >
        生成器努力制造逼真样本，判别器努力识别真伪；两者在对抗中共同改进。
      </Text>
      <FlowLabel kind="data" x={698} y={342} text="噪声 z" width={100} />
      <FlowArrow x={810} y={360} />
      <FlowLabel x={866} y={342} text="生成器 G" width={112} />
      <FlowArrow x={990} y={360} />
      <FlowLabel kind="data" x={1046} y={342} text="候选样本" width={128} />
      <Shape
        geometry={path("M 1110 398 L 1110 424 L 1054 424 L 1054 448", {
          viewBox: [1280, 720],
          fill: false,
        })}
        x={0}
        y={0}
        width={1280}
        height={720}
        fill={transparent}
        style={{
          borderColor: colors.cobalt,
          borderWidth: 2,
          endArrow: "triangle",
        }}
      />
      <FlowLabel kind="data" x={714} y={448} text="真实样本" width={160} />
      <FlowArrow x={886} y={466} width={40} color={colors.cyan} />
      <FlowLabel x={934} y={448} text="判别器 D：真或假" width={240} />
      <Shape
        geometry={path("M 1054 504 L 1054 528 L 922 528 L 922 404", {
          viewBox: [1280, 720],
          fill: false,
        })}
        x={0}
        y={0}
        width={1280}
        height={720}
        fill={transparent}
        style={{
          borderColor: colors.sun,
          borderWidth: 2,
          endArrow: "triangle",
        }}
      />
      <Text
        x={714}
        y={554}
        width={456}
        height={56}
        style={{
          fontSize: 16,
          lineHeight: 25,
          color: colors.muted,
          align: "center",
        }}
      >
        <P>判别结果回传给生成器，</P>
        <P>成为下一轮对抗的训练信号。</P>
      </Text>
    </SlideFrame>
  );
}
