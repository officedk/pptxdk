import { Shape, Slide, Text } from "pptxdk/react";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

/** Header shared by the two cards pages; individual slides may independently choose other compositions. */
export function PageFrame({
  id,
  title,
  subtitle,
  children,
}: {
  readonly id: string;
  readonly title: string;
  readonly subtitle: string;
  readonly children: ReactNode;
}) {
  return (
    <Slide id={id} background={colors.paper}>
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
