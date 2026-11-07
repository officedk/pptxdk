import { Shape, Slide, Text } from "@react-pptx/react-pptx";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

export function SlideFrame({
  slideId,
  chapter,
  title,
  number,
  source,
  dark = false,
  children,
}: {
  readonly slideId: string;
  readonly chapter: string;
  readonly title: string;
  readonly number: number;
  readonly source?: string;
  readonly dark?: boolean;
  readonly children: ReactNode;
}) {
  const ink = dark ? colors.paperBright : colors.ink;
  const secondary = dark ? colors.mist : colors.muted;
  return (
    <Slide slideId={slideId} background={dark ? colors.indigo : colors.paper}>
      <Text
        x={64}
        y={30}
        width={700}
        height={26}
        style={{ fontSize: 14, color: secondary, letterSpacing: 1 }}
      >
        {chapter}
      </Text>
      <Text
        x={64}
        y={70}
        width={1110}
        height={70}
        style={{
          fontFamily: "display",
          fontSize: 43,
          color: ink,
          lineHeight: 54,
          textWrap: "nowrap",
        }}
      >
        {title}
      </Text>
      <Shape x={1190} y={34} width={6} height={6} fill={colors.sun} />
      <Shape x={1174} y={34} width={6} height={6} fill={colors.cyan} />
      <Text
        x={1130}
        y={52}
        width={66}
        height={34}
        style={{
          fontFamily: "display",
          fontSize: 22,
          color: secondary,
          align: "right",
        }}
      >
        {String(number).padStart(2, "0")}
      </Text>
      {children}
      <Text
        x={64}
        y={676}
        width={1060}
        height={22}
        style={{ fontSize: 11, color: secondary, textWrap: "nowrap" }}
      >
        {source ?? "生成式人工智能简史"}
      </Text>
      <Text
        x={1130}
        y={674}
        width={66}
        height={22}
        style={{ fontSize: 11, color: secondary, align: "right" }}
      >
        {number} / 14
      </Text>
    </Slide>
  );
}
