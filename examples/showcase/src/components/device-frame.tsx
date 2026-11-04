import { Box, Group, Shape, preset } from "@react-pptx/react-pptx";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

/** 子内容在设备内容区内采用普通布局；外框和内容保留在同一个原生组合中。 */
export function DeviceFrame({
  x,
  y,
  width,
  height,
  children,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly children: ReactNode;
}) {
  return (
    <Group>
      <Shape
        x={x}
        y={y}
        width={width}
        height={height}
        fill={colors.ink}
        geometry={preset("roundRect")}
      />
      <Shape
        x={x + 12}
        y={y + 36}
        width={width - 24}
        height={height - 60}
        fill="#FFFFFF"
      />
      <Shape
        x={x + width / 2 - 25}
        y={y + 14}
        width={50}
        height={5}
        fill={colors.muted}
        geometry={preset("roundRect")}
      />
      <Box
        style={{
          position: "absolute",
          left: x + 30,
          top: y + 60,
          width: width - 60,
          height: height - 102,
          display: "flex",
          flexDirection: "column",
          gap: 20,
        }}
      >
        {children}
      </Box>
    </Group>
  );
}
