import {
  Box,
  MotionScope,
  Text,
  floatUp,
  type LayoutStyle,
  useMotionScope,
  wipe,
} from "@react-pptx/react-pptx";

import { colors, motion } from "../theme.ts";

export type CardContent = {
  readonly title: string;
  readonly body: string;
};

/** 内容卡片按阅读顺序揭示标题与正文，尺寸由父布局或 style 决定。 */
export function InfoCard({
  title,
  body,
  index = 0,
  style,
}: CardContent & {
  readonly index?: number;
  readonly style?: LayoutStyle;
}) {
  const reveal = useMotionScope();
  return (
    <Box
      style={{
        display: "flex",
        flexDirection: "column",
        minHeight: 340,
        padding: 32,
        gap: 20,
        backgroundColor: colors.surface,
        borderWidth: 1,
        borderColor: colors.soft,
        borderRadius: 6,
        ...style,
      }}
    >
      <Box style={{ width: 44, height: 5, backgroundColor: colors.accent }} />
      <Text
        height="auto"
        style={{
          fontSize: 60,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.soft,
        }}
      >
        {String(index + 1).padStart(2, "0")}
      </Text>
      <MotionScope scope={reveal} mode="sequence">
        <Text
          height="auto"
          style={{ fontSize: 30, fontWeight: 700, color: colors.ink }}
          enter={wipe({ from: "left", duration: motion.reveal })}
        >
          {title}
        </Text>
        <Text
          height="auto"
          style={{ fontSize: 20, lineHeight: 28, color: colors.muted }}
          enter={floatUp({ duration: motion.reveal })}
        >
          {body}
        </Text>
      </MotionScope>
    </Box>
  );
}
