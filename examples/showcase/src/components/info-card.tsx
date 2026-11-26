import {
  Box,
  floatUp,
  MotionScope,
  Text,
  useMotionScope,
  wipe,
  type LayoutStyle,
} from "@officedk/pptxdk/react";

import { colors, motion } from "../theme.ts";

export type CardContent = {
  readonly title: string;
  readonly body: string;
};

/** Content cards reveal title and body in reading order; dimensions are determined by parent layout or style. */
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
        background: colors.surface,
        border: { color: colors.soft, width: 1 },
        borderRadius: 6,
        ...style,
      }}
    >
      <Box style={{ width: 44, height: 5, background: colors.accent }} />
      <Text
        style={{
          height: "auto",
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
          style={{
            height: "auto",
            fontSize: 30,
            fontWeight: 700,
            color: colors.ink,
          }}
          enter={wipe({ from: "left", duration: motion.reveal })}
        >
          {title}
        </Text>
        <Text
          style={{
            height: "auto",
            fontSize: 20,
            lineHeight: 28,
            color: colors.muted,
          }}
          enter={floatUp({ duration: motion.reveal })}
        >
          {body}
        </Text>
      </MotionScope>
    </Box>
  );
}
