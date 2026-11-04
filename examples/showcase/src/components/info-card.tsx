import {
  Box,
  MotionScope,
  Text,
  fade,
  useMotionScope,
} from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

export type CardContent = {
  readonly title: string;
  readonly body: string;
  readonly accent: string;
};

/** 卡片只约定内容与宽度；文字按内容定高，内部标题和正文顺序揭示。 */
export function InfoCard({
  title,
  body,
  accent,
  width,
}: CardContent & { readonly width: number }) {
  const reveal = useMotionScope();
  return (
    <Box
      style={{
        display: "flex",
        flexDirection: "column",
        width,
        minHeight: 250,
        padding: 28,
        gap: 22,
        backgroundColor: "#FFFFFF",
        borderWidth: 1,
        borderColor: accent,
      }}
    >
      <MotionScope scope={reveal} mode="sequence">
        <Text
          width={width - 56}
          height="auto"
          style={{ fontSize: 30, color: accent }}
          enter={fade({ duration: 280 })}
        >
          {title}
        </Text>
        <Text
          width={width - 56}
          height="auto"
          style={{ fontSize: 23, color: colors.ink }}
          enter={fade({ duration: 350 })}
        >
          {body}
        </Text>
      </MotionScope>
    </Box>
  );
}
