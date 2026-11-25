import { Shape, Slide, Text } from "pptxdk/react";

import { ProcessPath } from "../components/process-path.tsx";
import { colors } from "../theme.ts";

export function ProcessPage() {
  return (
    <Slide id="process" background={colors.paper}>
      <Shape x={80} y={62} width={42} height={5} fill={colors.accent} />
      <Text
        x={80}
        y={94}
        width={760}
        height={72}
        style={{ fontSize: 54, lineHeight: 60, fontWeight: 700 }}
      >
        Follow the thread.
      </Text>
      <Text
        x={880}
        y={108}
        width={320}
        height={60}
        style={{ fontSize: 18, lineHeight: 26, color: colors.muted }}
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
        x={80}
        y={622}
        width={250}
        height={28}
        style={{
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
