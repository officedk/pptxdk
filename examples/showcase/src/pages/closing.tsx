import { Box, Slide, Text } from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

export function ClosingPage() {
  return (
    <Slide slideId="closing" background={colors.ink}>
      <Box
        style={{
          position: "absolute",
          left: 96,
          top: 130,
          width: 1080,
          height: 500,
          display: "flex",
          flexDirection: "column",
          gap: 34,
        }}
      >
        <Text
          width={1080}
          height="auto"
          style={{ fontSize: 64, color: colors.paper }}
        >
          让表达回到内容
        </Text>
        <Text
          width={1080}
          height="auto"
          style={{ fontSize: 34, color: colors.coral }}
        >
          层级，让重点可见。
        </Text>
        <Text
          width={1080}
          height="auto"
          style={{ fontSize: 34, color: "#7FC9C1" }}
        >
          关系，让结构清晰。
        </Text>
        <Text
          width={1080}
          height="auto"
          style={{ fontSize: 34, color: colors.paper }}
        >
          节奏，让理解逐步发生。
        </Text>
      </Box>
    </Slide>
  );
}
