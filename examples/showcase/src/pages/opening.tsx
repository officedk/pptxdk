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
        x={82}
        y={74}
        width={520}
        height={30}
        style={{
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.surface,
        }}
      >
        DESIGN / IN MOTION
      </Text>
      <Shape x={82} y={128} width={52} height={6} fill={colors.current} />
      <Text
        x={76}
        y={190}
        width={660}
        height={190}
        style={{
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
          x={82}
          y={432}
          width={510}
          height={68}
          style={{ fontSize: 21, lineHeight: 30, color: colors.soft }}
          enter={wipe({ from: "left", duration: motion.reveal })}
        >
          Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do
          eiusmod tempor.
        </Text>
      </Step>
      <Step n={1}>
        <Group id="opening-mark" enter={spinIn({ duration: motion.settle })}>
          <Shape
            x={842}
            y={150}
            width={360}
            height={360}
            geometry={preset("ellipse")}
            fill={colors.surface}
          />
          <Shape
            x={874}
            y={182}
            width={296}
            height={296}
            geometry={preset("ellipse")}
            fill={colors.accent}
          />
          <Shape
            x={962}
            y={272}
            width={120}
            height={120}
            geometry={preset("ellipse")}
            fill={colors.ink}
          />
          <Shape
            x={1012}
            y={242}
            width={20}
            height={180}
            fill={colors.current}
          />
          <Shape
            x={932}
            y={322}
            width={180}
            height={20}
            fill={colors.current}
          />
        </Group>
      </Step>
      <Shape
        x={1130}
        y={568}
        width={54}
        height={54}
        fill={colors.ink}
        style={{ rotate: 12 }}
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
        01 — 16
      </Text>
    </Slide>
  );
}
