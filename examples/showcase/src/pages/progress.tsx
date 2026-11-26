import { Shape, Slide, Text } from "@officedk/pptxdk/react";

import { Progress } from "../components/progress.tsx";
import { colors } from "../theme.ts";

export function ProgressPage() {
  return (
    <Slide id="progress" background={colors.ink}>
      <Text
        style={{
          position: "absolute",
          left: 82,
          top: 66,
          width: 500,
          height: 30,
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.accentSoft,
        }}
      >
        07 / PROGRESS
      </Text>
      <Shape
        style={{
          position: "absolute",
          left: 930,
          top: 80,
          width: 268,
          height: 2,
          background: colors.darkSurface,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 80,
          top: 124,
          width: 620,
          height: 74,
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
