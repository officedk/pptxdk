import { wipe } from "pptxdk";
import { Box, Group, Shape, Slide, Step, Text } from "pptxdk/react";

import { InfoCard } from "../components/info-card.tsx";
import { colors, motion } from "../theme.ts";

export function ComparisonPage() {
  return (
    <Slide id="comparison" background={colors.paper}>
      <Shape x={0} y={0} width={610} height={720} fill={colors.accent} />
      <Text
        x={72}
        y={70}
        width={430}
        height={30}
        style={{
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.soft,
        }}
      >
        05 / COMPARISON
      </Text>
      <Text
        x={70}
        y={132}
        width={480}
        height={130}
        style={{
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
        x={700}
        y={74}
        width={480}
        height={104}
        style={{
          fontSize: 88,
          lineHeight: 92,
          fontWeight: 700,
          color: colors.ink,
        }}
      >
        A / B
      </Text>
      <Text
        x={706}
        y={190}
        width={400}
        height={44}
        style={{
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
            x={704}
            y={316}
            width={48}
            height={40}
            style={{ fontSize: 25, fontWeight: 700 }}
          >
            A
          </Text>
          <Shape x={770} y={320} width={360} height={24} fill={colors.soft} />
          <Shape x={770} y={320} width={310} height={24} fill={colors.accent} />
          <Text
            x={1140}
            y={310}
            width={60}
            height={42}
            style={{ fontSize: 25, fontFamily: "medium", fontWeight: 500 }}
          >
            84
          </Text>
          <Text
            x={704}
            y={434}
            width={48}
            height={40}
            style={{ fontSize: 25, fontWeight: 700 }}
          >
            B
          </Text>
          <Shape x={770} y={438} width={360} height={24} fill={colors.soft} />
          <Shape x={770} y={438} width={205} height={24} fill={colors.ink} />
          <Text
            x={1140}
            y={428}
            width={60}
            height={42}
            style={{ fontSize: 25, fontFamily: "medium", fontWeight: 500 }}
          >
            56
          </Text>
          <Shape x={1072} y={310} width={8} height={44} fill={colors.current} />
        </Group>
      </Step>
      <Text
        x={706}
        y={564}
        width={420}
        height={68}
        style={{ fontSize: 18, lineHeight: 26, color: colors.muted }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod
        tempor.
      </Text>
    </Slide>
  );
}
