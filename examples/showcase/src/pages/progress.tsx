import { Shape, Slide, Text } from "@react-pptx/react-pptx";

import { Progress } from "../components/progress.tsx";
import { colors } from "../theme.ts";

export function ProgressPage() {
  return (
    <Slide id="progress" background={colors.ink}>
      <Text
        x={82}
        y={66}
        width={500}
        height={30}
        style={{
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.accentSoft,
        }}
      >
        07 / PROGRESS
      </Text>
      <Shape x={930} y={80} width={268} height={2} fill={colors.darkSurface} />
      <Text
        x={80}
        y={124}
        width={620}
        height={74}
        style={{
          fontSize: 52,
          lineHeight: 62,
          fontWeight: 700,
          color: colors.surface,
        }}
      >
        Steady progress.
      </Text>
      <Progress x={80} y={176} width={1120} completed={3} total={4} />
    </Slide>
  );
}
