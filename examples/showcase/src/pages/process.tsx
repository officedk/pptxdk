import { Shape, Slide, Text } from "@officedk/pptxdk/react";

import { ProcessPath } from "../components/process-path.tsx";
import { colors } from "../theme.ts";

export function ProcessPage() {
  return (
    <Slide id="process" background={colors.paper}>
      <Shape
        style={{
          position: "absolute",
          left: 80,
          top: 62,
          width: 42,
          height: 5,
          background: colors.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 80,
          top: 94,
          width: 760,
          height: 72,
          fontSize: 54,
          lineHeight: 60,
          fontWeight: 700,
        }}
      >
        Follow the thread.
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 880,
          top: 108,
          width: 320,
          height: 60,
          fontSize: 18,
          lineHeight: 26,
          color: colors.muted,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      <ProcessPath
        x={80}
        y={242}
        width={1120}
        nodes={[
          {
            title: "Lorem ipsum",
            body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
          },
          {
            title: "Dolor sit",
            body: "Sed do eiusmod tempor incididunt ut labore et dolore.",
          },
          {
            title: "Amet elit",
            body: "Ut enim ad minim veniam, quis nostrud exercitation.",
          },
        ]}
      />
      <Text
        style={{
          position: "absolute",
          left: 80,
          top: 622,
          width: 250,
          height: 28,
          fontSize: 15,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.muted,
        }}
      >
        06 / PROCESS
      </Text>
    </Slide>
  );
}
