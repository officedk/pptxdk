import { Shape, Slide, Text } from "@react-pptx/react-pptx";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

/** 两页卡片内容共享的页首；具体页面仍可独立选择其他构图。 */
export function PageFrame({
  slideId,
  title,
  subtitle,
  children,
}: {
  readonly slideId: string;
  readonly title: string;
  readonly subtitle: string;
  readonly children: ReactNode;
}) {
  return (
    <Slide slideId={slideId} background={colors.paper}>
      <Shape x={80} y={58} width={42} height={5} fill={colors.accent} />
      <Text
        x={80}
        y={88}
        width={760}
        height={72}
        style={{
          fontSize: 52,
          lineHeight: 58,
          fontWeight: 700,
          color: colors.ink,
        }}
      >
        {title}
      </Text>
      <Text
        x={870}
        y={100}
        width={330}
        height={90}
        style={{ fontSize: 18, lineHeight: 25, color: colors.muted }}
      >
        {subtitle}
      </Text>
      <Shape x={80} y={184} width={1120} height={1} fill={colors.soft} />
      {children}
    </Slide>
  );
}
