import { Shape, Slide, Text } from "@officedk/pptxdk/react";
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
      <Shape
        style={{
          position: "absolute",
          left: 80,
          top: 58,
          width: 42,
          height: 5,
          background: colors.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 80,
          top: 88,
          width: 760,
          height: 72,
          fontSize: 52,
          lineHeight: 58,
          fontWeight: 700,
          color: colors.ink,
        }}
      >
        {title}
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 870,
          top: 100,
          width: 330,
          height: 90,
          fontSize: 18,
          lineHeight: 25,
          color: colors.muted,
        }}
      >
        {subtitle}
      </Text>
      <Shape
        style={{
          position: "absolute",
          left: 80,
          top: 184,
          width: 1120,
          height: 1,
          background: colors.soft,
        }}
      />
      {children}
    </Slide>
  );
}
