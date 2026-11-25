import { Box, Shape, Slide, Text } from "@officedk/pptxdk/react";

import { DeviceFrame } from "../components/device-frame.tsx";
import { colors } from "../theme.ts";

function ExampleContent({ compact = false }: { compact?: boolean }) {
  return (
    <>
      <Text
        width={compact ? 120 : 360}
        height="auto"
        style={{
          fontSize: compact ? 20 : 34,
          fontWeight: 700,
          color: colors.ink,
        }}
      >
        Lorem ipsum
      </Text>
      <Text
        width={compact ? 120 : 390}
        height="auto"
        style={{
          fontSize: compact ? 12 : 17,
          lineHeight: compact ? 17 : 24,
          color: colors.muted,
        }}
      >
        {compact
          ? "Lorem ipsum dolor sit amet."
          : "Lorem ipsum dolor sit amet, consectetur adipiscing elit."}
      </Text>
      <Box
        style={{
          width: compact ? 120 : 390,
          height: compact ? 68 : 128,
          backgroundColor: colors.soft,
        }}
      />
      <Box
        style={{
          width: compact ? 74 : 170,
          height: compact ? 7 : 10,
          backgroundColor: colors.accent,
        }}
      />
    </>
  );
}

export function DevicesPage() {
  return (
    <Slide id="devices" background={colors.paper}>
      <Shape x={80} y={62} width={42} height={5} fill={colors.accent} />
      <Text
        x={80}
        y={112}
        width={420}
        height={146}
        style={{ fontSize: 57, fontWeight: 700, lineHeight: 62 }}
      >
        One story.{`\n`}Every screen.
      </Text>
      <Text
        x={82}
        y={286}
        width={310}
        height={94}
        style={{ fontSize: 19, lineHeight: 27, color: colors.muted }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod
        tempor.
      </Text>
      <DeviceFrame id="device-desktop" x={558} y={126} width={622} height={410}>
        <ExampleContent />
      </DeviceFrame>
      <DeviceFrame id="device-phone" x={408} y={326} width={196} height={326}>
        <ExampleContent compact />
      </DeviceFrame>
      <Text
        x={82}
        y={640}
        width={260}
        height={28}
        style={{
          fontSize: 15,
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.muted,
        }}
      >
        09 / DEVICES
      </Text>
    </Slide>
  );
}
