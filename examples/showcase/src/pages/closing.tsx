import { preset, Shape, Slide, Text } from "@officedk/pptxdk/react";

import { colors } from "../theme.ts";

export function ClosingPage() {
  return (
    <Slide id="closing" background={colors.ink}>
      <Shape
        style={{
          position: "absolute",
          left: 82,
          top: 106,
          width: 52,
          height: 6,
          background: colors.current,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 76,
          top: 184,
          width: 740,
          height: 190,
          fontSize: 78,
          fontWeight: 700,
          lineHeight: 84,
          color: colors.surface,
        }}
      >
        Make room{`\n`}for ideas.
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 82,
          top: 426,
          width: 480,
          height: 66,
          fontSize: 21,
          lineHeight: 30,
          color: colors.accentSoft,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      <Shape
        geometry={preset("ellipse")}
        style={{
          position: "absolute",
          left: 954,
          top: 192,
          width: 250,
          height: 250,
          background: colors.accent,
        }}
      />
      <Shape
        geometry={preset("ellipse")}
        style={{
          position: "absolute",
          left: 994,
          top: 232,
          width: 170,
          height: 170,
          background: colors.ink,
        }}
      />
      <Shape
        geometry={preset("ellipse")}
        style={{
          position: "absolute",
          left: 1059,
          top: 297,
          width: 40,
          height: 40,
          background: colors.current,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 82,
          top: 642,
          width: 180,
          height: 28,
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
