import { Box, Slide, Text } from "@react-pptx/react-pptx";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

/** 标题区与内容区分开，正文既可组合布局，也可放置局部几何构图。 */
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
      <Box
        style={{
          position: "absolute",
          left: 64,
          top: 40,
          width: 1152,
          display: "flex",
          flexDirection: "column",
          gap: 12,
        }}
      >
        <Text width={1152} height="auto" style={{ fontSize: 40 }}>
          {title}
        </Text>
        <Text
          width={1152}
          height="auto"
          style={{ fontSize: 19, color: colors.muted }}
        >
          {subtitle}
        </Text>
      </Box>
      {children}
    </Slide>
  );
}
