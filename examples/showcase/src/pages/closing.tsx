import { preset } from "pptxdk";
import { Shape, Slide, Text } from "pptxdk/react";

import { colors } from "../theme.ts";

export function ClosingPage() {
  return (
    <Slide id="closing" background={colors.ink}>
      <Shape x={82} y={106} width={52} height={6} fill={colors.current} />
      <Text
        x={76}
        y={184}
        width={740}
        height={190}
        style={{
          fontSize: 78,
          fontWeight: 700,
          lineHeight: 84,
          color: colors.surface,
        }}
      >
        Make room{`\n`}for ideas.
      </Text>
      <Text
        x={82}
        y={426}
        width={480}
        height={66}
        style={{ fontSize: 21, lineHeight: 30, color: colors.accentSoft }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      <Shape
        x={954}
        y={192}
        width={250}
        height={250}
        geometry={preset("ellipse")}
        fill={colors.accent}
      />
      <Shape
        x={994}
        y={232}
        width={170}
        height={170}
        geometry={preset("ellipse")}
        fill={colors.ink}
      />
      <Shape
        x={1059}
        y={297}
        width={40}
        height={40}
        geometry={preset("ellipse")}
        fill={colors.current}
      />
      <Text
        x={82}
        y={642}
        width={180}
        height={28}
        style={{
          fontSize: 15,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.surface,
        }}
      >
        16 — 16
      </Text>
    </Slide>
  );
}
