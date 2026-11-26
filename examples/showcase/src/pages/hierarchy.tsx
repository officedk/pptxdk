import { Box, Slide, Text } from "@officedk/pptxdk/react";

import { colors } from "../theme.ts";

export function HierarchyPage() {
  return (
    <Slide id="hierarchy" background={colors.paper}>
      <Text
        style={{
          position: "absolute",
          left: 60,
          top: 36,
          width: 350,
          height: 230,
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
            style={{
              width: 500,
              height: "auto",
              fontSize: 16,
              fontWeight: 700,
              letterSpacing: 2,
              color: colors.accent,
            }}
          >
            CLEAR HIERARCHY
          </Text>
          <Text
            style={{
              width: 590,
              height: "auto",
              fontSize: 66,
              fontWeight: 700,
              lineHeight: 70,
              color: colors.ink,
            }}
          >
            Make the{`\n`}clear choice.
          </Text>
          <Text
            style={{
              width: 440,
              height: "auto",
              fontSize: 21,
              lineHeight: 30,
              color: colors.muted,
            }}
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
              background: colors.accent,
            }}
          >
            <Text
              style={{
                width: 370,
                height: "auto",
                fontSize: 18,
                color: colors.soft,
              }}
            >
              01 / LOREM
            </Text>
            <Text
              style={{
                width: 370,
                height: "auto",
                fontSize: 38,
                fontWeight: 700,
                color: colors.surface,
              }}
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
              border: { color: colors.soft, width: 1 },
            }}
          >
            <Text
              style={{
                width: 370,
                height: "auto",
                fontSize: 17,
                color: colors.muted,
              }}
            >
              02 / IPSUM
            </Text>
            <Text
              style={{
                width: 370,
                height: "auto",
                fontSize: 28,
                fontFamily: "medium",
                fontWeight: 500,
              }}
            >
              Dolor sit amet
            </Text>
            <Box style={{ width: 44, height: 4, background: colors.current }} />
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
              style={{
                width: 370,
                height: "auto",
                fontSize: 17,
                color: colors.muted,
              }}
            >
              03 / DOLOR
            </Text>
            <Text
              style={{
                width: 370,
                height: "auto",
                fontSize: 24,
                color: colors.muted,
              }}
            >
              Consectetur elit
            </Text>
          </Box>
        </Box>
      </Box>
    </Slide>
  );
}
