import { Box, Shape, Slide, Step, Text } from "@react-pptx/react-pptx";

import { InfoCard } from "../components/info-card.tsx";
import { ProcessPath } from "../components/process-path.tsx";
import { colors } from "../theme.ts";

export function CombinationPage() {
  return (
    <Slide slideId="combination" background={colors.paper}>
      <Shape x={0} y={0} width={410} height={720} fill={colors.ink} />
      <Text
        x={60}
        y={74}
        width={290}
        height={30}
        style={{
          fontSize: 16,
          fontWeight: 700,
          letterSpacing: 2,
          color: colors.accentSoft,
        }}
      >
        15 / TOGETHER
      </Text>
      <Text
        x={56}
        y={126}
        width={320}
        height={130}
        style={{
          fontSize: 52,
          fontWeight: 700,
          lineHeight: 57,
          color: colors.surface,
        }}
      >
        Pieces,{`\n`}together.
      </Text>
      <Text
        x={62}
        y={274}
        width={280}
        height={60}
        style={{ fontSize: 18, lineHeight: 25, color: colors.accentSoft }}
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
        x={480}
        y={86}
        width={670}
        height={70}
        style={{ fontSize: 48, fontWeight: 700 }}
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
      <Shape x={484} y={574} width={44} height={5} fill={colors.current} />
      <Text
        x={550}
        y={558}
        width={560}
        height={62}
        style={{ fontSize: 20, lineHeight: 28, color: colors.muted }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod
        tempor.
      </Text>
    </Slide>
  );
}
