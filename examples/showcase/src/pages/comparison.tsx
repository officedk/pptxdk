import {
  Box,
  Group,
  Shape,
  Slide,
  Step,
  Text,
  wipe,
} from "@officedk/pptxdk/react";

import { InfoCard } from "../components/info-card.tsx";
import { colors, motion } from "../theme.ts";

export function ComparisonPage() {
  return (
    <Slide id="comparison" background={colors.paper}>
      <Shape
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: 610,
          height: 720,
          background: colors.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 72,
          top: 70,
          width: 430,
          height: 30,
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.soft,
        }}
      >
        05 / COMPARISON
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 70,
          top: 132,
          width: 480,
          height: 130,
          fontSize: 52,
          fontWeight: 700,
          lineHeight: 57,
          color: colors.surface,
        }}
      >
        Same scale.{`\n`}Different weight.
      </Text>
      <Box style={{ position: "absolute", left: 74, top: 330, width: 410 }}>
        <Step n={1}>
          <InfoCard
            index={0}
            title="Lorem ipsum"
            body="Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor."
          />
        </Step>
      </Box>
      <Text
        style={{
          position: "absolute",
          left: 700,
          top: 74,
          width: 480,
          height: 104,
          fontSize: 88,
          lineHeight: 92,
          fontWeight: 700,
          color: colors.ink,
        }}
      >
        A / B
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 706,
          top: 190,
          width: 400,
          height: 44,
          fontSize: 30,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.accent,
        }}
      >
        One baseline.
      </Text>
      <Step n={2}>
        <Group enter={wipe({ from: "left", duration: motion.settle })}>
          <Text
            style={{
              position: "absolute",
              left: 704,
              top: 316,
              width: 48,
              height: 40,
              fontSize: 25,
              fontWeight: 700,
            }}
          >
            A
          </Text>
          <Shape
            style={{
              position: "absolute",
              left: 770,
              top: 320,
              width: 360,
              height: 24,
              background: colors.soft,
            }}
          />
          <Shape
            style={{
              position: "absolute",
              left: 770,
              top: 320,
              width: 310,
              height: 24,
              background: colors.accent,
            }}
          />
          <Text
            style={{
              position: "absolute",
              left: 1140,
              top: 310,
              width: 60,
              height: 42,
              fontSize: 25,
              fontFamily: "medium",
              fontWeight: 500,
            }}
          >
            84
          </Text>
          <Text
            style={{
              position: "absolute",
              left: 704,
              top: 434,
              width: 48,
              height: 40,
              fontSize: 25,
              fontWeight: 700,
            }}
          >
            B
          </Text>
          <Shape
            style={{
              position: "absolute",
              left: 770,
              top: 438,
              width: 360,
              height: 24,
              background: colors.soft,
            }}
          />
          <Shape
            style={{
              position: "absolute",
              left: 770,
              top: 438,
              width: 205,
              height: 24,
              background: colors.ink,
            }}
          />
          <Text
            style={{
              position: "absolute",
              left: 1140,
              top: 428,
              width: 60,
              height: 42,
              fontSize: 25,
              fontFamily: "medium",
              fontWeight: 500,
            }}
          >
            56
          </Text>
          <Shape
            style={{
              position: "absolute",
              left: 1072,
              top: 310,
              width: 8,
              height: 44,
              background: colors.current,
            }}
          />
        </Group>
      </Step>
      <Text
        style={{
          position: "absolute",
          left: 706,
          top: 564,
          width: 420,
          height: 68,
          fontSize: 18,
          lineHeight: 26,
          color: colors.muted,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod
        tempor.
      </Text>
    </Slide>
  );
}
