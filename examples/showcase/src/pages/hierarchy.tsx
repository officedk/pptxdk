import { Box, Slide, Text } from "@officedk/pptxdk/react";

import { colors } from "../theme.ts";

export function HierarchyPage() {
  return (
    <Slide id="hierarchy" background={colors.paper}>
      <Text
        x={60}
        y={36}
        width={350}
        height={230}
        style={{
          fontSize: 210,
          lineHeight: 210,
          fontWeight: 700,
          color: colors.soft,
        }}
      >
        02
      </Text>
      <Box
        style={{
          position: "absolute",
          left: 80,
          top: 128,
          width: 1120,
          display: "flex",
          gap: 100,
        }}
      >
        <Box
          style={{
            width: 590,
            display: "flex",
            flexDirection: "column",
            gap: 28,
          }}
        >
          <Text
            width={500}
            height="auto"
            style={{
              fontSize: 16,
              fontWeight: 700,
              letterSpacing: 2,
              color: colors.accent,
            }}
          >
            CLEAR HIERARCHY
          </Text>
          <Text
            width={590}
            height="auto"
            style={{
              fontSize: 66,
              fontWeight: 700,
              lineHeight: 70,
              color: colors.ink,
            }}
          >
            Make the{`\n`}clear choice.
          </Text>
          <Text
            width={440}
            height="auto"
            style={{ fontSize: 21, lineHeight: 30, color: colors.muted }}
          >
            Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do
            eiusmod tempor incididunt.
          </Text>
        </Box>
        <Box
          style={{
            width: 430,
            display: "flex",
            flexDirection: "column",
            gap: 0,
          }}
        >
          <Box
            style={{
              width: 430,
              padding: 30,
              display: "flex",
              flexDirection: "column",
              gap: 16,
              backgroundColor: colors.accent,
            }}
          >
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 18, color: colors.soft }}
            >
              01 / LOREM
            </Text>
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 38, fontWeight: 700, color: colors.surface }}
            >
              First things first.
            </Text>
          </Box>
          <Box
            style={{
              width: 430,
              padding: 30,
              display: "flex",
              flexDirection: "column",
              gap: 12,
              borderWidth: 1,
              borderColor: colors.soft,
            }}
          >
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 17, color: colors.muted }}
            >
              02 / IPSUM
            </Text>
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 28, fontFamily: "medium", fontWeight: 500 }}
            >
              Dolor sit amet
            </Text>
            <Box
              style={{ width: 44, height: 4, backgroundColor: colors.current }}
            />
          </Box>
          <Box
            style={{
              width: 430,
              padding: 30,
              display: "flex",
              flexDirection: "column",
              gap: 12,
            }}
          >
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 17, color: colors.muted }}
            >
              03 / DOLOR
            </Text>
            <Text
              width={370}
              height="auto"
              style={{ fontSize: 24, color: colors.muted }}
            >
              Consectetur elit
            </Text>
          </Box>
        </Box>
      </Box>
    </Slide>
  );
}
