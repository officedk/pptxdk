import { Box, Shape, Slide, Step, Text } from "@officedk/pptxdk/react";

import { InfoCard } from "../components/info-card.tsx";
import { ProcessPath } from "../components/process-path.tsx";
import { colors } from "../theme.ts";

export function CombinationPage() {
  return (
    <Slide id="combination" background={colors.paper}>
      <Shape
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: 410,
          height: 720,
          background: colors.ink,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 60,
          top: 74,
          width: 290,
          height: 30,
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.accentSoft,
        }}
      >
        15 / TOGETHER
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 56,
          top: 126,
          width: 320,
          height: 130,
          fontSize: 52,
          fontWeight: 700,
          lineHeight: 57,
          color: colors.surface,
        }}
      >
        Pieces,{`\n`}together.
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 62,
          top: 274,
          width: 280,
          height: 60,
          fontSize: 18,
          lineHeight: 25,
          color: colors.accentSoft,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      <Box style={{ position: "absolute", left: 62, top: 352, width: 292 }}>
        <Step n={1}>
          <InfoCard
            index={0}
            title="Lorem ipsum"
            body="Lorem ipsum dolor sit amet, consectetur adipiscing elit."
          />
        </Step>
      </Box>
      <Text
        style={{
          position: "absolute",
          left: 480,
          top: 86,
          width: 670,
          height: 70,
          fontSize: 48,
          fontWeight: 700,
        }}
      >
        Lorem ipsum dolor sit.
      </Text>
      <ProcessPath
        x={478}
        y={230}
        width={712}
        firstStep={2}
        nodes={[
          {
            title: "Lorem",
            body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
          },
          {
            title: "Ipsum",
            body: "Sed do eiusmod tempor incididunt ut labore et dolore.",
          },
        ]}
      />
      <Shape
        style={{
          position: "absolute",
          left: 484,
          top: 574,
          width: 44,
          height: 5,
          background: colors.current,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 550,
          top: 558,
          width: 560,
          height: 62,
          fontSize: 20,
          lineHeight: 28,
          color: colors.muted,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod
        tempor.
      </Text>
    </Slide>
  );
}
