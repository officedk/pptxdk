import {
  Group,
  preset,
  Shape,
  Slide,
  spinIn,
  Step,
  Text,
  wipe,
} from "@officedk/pptxdk/react";

import { colors, motion } from "../theme.ts";

/** The opening establishes the visual language of the entire deck with native geometric logos and staged entrances. */
export function OpeningPage() {
  return (
    <Slide id="opening" background={colors.accent}>
      <Text
        style={{
          position: "absolute",
          left: 82,
          top: 74,
          width: 520,
          height: 30,
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.surface,
        }}
      >
        DESIGN / IN MOTION
      </Text>
      <Shape
        style={{
          position: "absolute",
          left: 82,
          top: 128,
          width: 52,
          height: 6,
          background: colors.current,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 76,
          top: 190,
          width: 660,
          height: 190,
          fontSize: 84,
          fontWeight: 700,
          lineHeight: 88,
          color: colors.surface,
        }}
      >
        Ideas,{`\n`}in motion.
      </Text>
      <Step n={2}>
        <Text
          style={{
            position: "absolute",
            left: 82,
            top: 432,
            width: 510,
            height: 68,
            fontSize: 21,
            lineHeight: 30,
            color: colors.soft,
          }}
          enter={wipe({ from: "left", duration: motion.reveal })}
        >
          Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do
          eiusmod tempor.
        </Text>
      </Step>
      <Step n={1}>
        <Group id="opening-mark" enter={spinIn({ duration: motion.settle })}>
          <Shape
            geometry={preset("ellipse")}
            style={{
              position: "absolute",
              left: 842,
              top: 150,
              width: 360,
              height: 360,
              background: colors.surface,
            }}
          />
          <Shape
            geometry={preset("ellipse")}
            style={{
              position: "absolute",
              left: 874,
              top: 182,
              width: 296,
              height: 296,
              background: colors.accent,
            }}
          />
          <Shape
            geometry={preset("ellipse")}
            style={{
              position: "absolute",
              left: 962,
              top: 272,
              width: 120,
              height: 120,
              background: colors.ink,
            }}
          />
          <Shape
            style={{
              position: "absolute",
              left: 1012,
              top: 242,
              width: 20,
              height: 180,
              background: colors.current,
            }}
          />
          <Shape
            style={{
              position: "absolute",
              left: 932,
              top: 322,
              width: 180,
              height: 20,
              background: colors.current,
            }}
          />
        </Group>
      </Step>
      <Shape
        style={{
          position: "absolute",
          left: 1130,
          top: 568,
          width: 54,
          height: 54,
          background: colors.ink,
          rotate: 12,
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
        01 — 16
      </Text>
    </Slide>
  );
}
