import { Box, Group, preset, Shape } from "@officedk/pptxdk/react";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

/** Child content uses regular layout inside the screen area; device frame and content are preserved as a single native Group. */
export function DeviceFrame({
  id,
  x,
  y,
  width,
  height,
  children,
}: {
  readonly id: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly children: ReactNode;
}) {
  const compact = width < 300;
  const inset = compact ? 9 : 12;
  const top = compact ? 30 : 38;
  return (
    <Group id={id}>
      <Shape
        geometry={preset("roundRect")}
        style={{
          position: "absolute",
          left: x,
          top: y,
          width,
          height,
          background: colors.ink,
          borderRadius: compact ? 12 : 10,
          boxShadow: { color: colors.ink, blur: 12, distance: 6, angle: 45 },
        }}
      />
      <Shape
        geometry={preset("roundRect")}
        style={{
          position: "absolute",
          left: x + inset,
          top: y + top,
          width: width - inset * 2,
          height: height - top - inset,
          background: colors.surface,
          borderRadius: compact ? 6 : 4,
        }}
      />
      {compact ? (
        <Shape
          geometry={preset("roundRect")}
          style={{
            position: "absolute",
            left: x + width / 2 - 24,
            top: y + 13,
            width: 48,
            height: 5,
            background: colors.muted,
            borderRadius: 2.5,
          }}
        />
      ) : (
        <>
          {[0, 1, 2].map((index) => (
            <Shape
              key={index}
              geometry={preset("ellipse")}
              style={{
                position: "absolute",
                left: x + 24 + index * 14,
                top: y + 15,
                width: 6,
                height: 6,
                background: index === 0 ? colors.current : colors.accentSoft,
              }}
            />
          ))}
        </>
      )}
      <Box
        style={{
          position: "absolute",
          left: x + inset + (compact ? 15 : 28),
          top: y + top + (compact ? 24 : 30),
          width: width - (inset + (compact ? 15 : 28)) * 2,
          height: height - top - inset - (compact ? 44 : 56),
          display: "flex",
          flexDirection: "column",
          gap: compact ? 14 : 18,
        }}
      >
        {children}
      </Box>
    </Group>
  );
}
