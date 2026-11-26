import { preset, Shape, Text, type ColorValue } from "@officedk/pptxdk/react";
import type { ReactNode } from "react";

import { colors, transparent } from "../theme.ts";

export function Rule({
  x,
  y,
  width,
  color = colors.line,
  thickness = 1,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly color?: ColorValue;
  readonly thickness?: number;
}) {
  return (
    <Shape
      style={{
        position: "absolute",
        left: x,
        top: y,
        width,
        height: thickness,
        background: color,
      }}
    />
  );
}

export function Circle({
  x,
  y,
  diameter,
  fill = colors.cobalt,
  stroke,
  strokeWidth = 1,
}: {
  readonly x: number;
  readonly y: number;
  readonly diameter: number;
  readonly fill?: ColorValue;
  readonly stroke?: ColorValue;
  readonly strokeWidth?: number;
}) {
  return (
    <Shape
      geometry={preset("ellipse")}
      style={{
        position: "absolute",
        left: x,
        top: y,
        width: diameter,
        height: diameter,
        background: fill,
        ...(stroke === undefined
          ? {}
          : { border: { color: stroke, width: strokeWidth } }),
      }}
    />
  );
}

export function Arrow({
  x,
  y,
  width = 44,
  color = colors.cobalt,
}: {
  readonly x: number;
  readonly y: number;
  readonly width?: number;
  readonly color?: ColorValue;
}) {
  return (
    <Shape
      geometry={preset("line")}
      style={{
        position: "absolute",
        left: x,
        top: y + 8,
        width,
        height: 1,
        background: transparent,
        border: {
          color,
          width: 2,
          endArrow: "triangle",
        },
      }}
    />
  );
}

export function Tag({
  x,
  y,
  width,
  children,
  dark = false,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly children: ReactNode;
  readonly dark?: boolean;
}) {
  return (
    <Text
      style={{
        position: "absolute",
        left: x,
        top: y,
        width,
        height: 28,
        fontFamily: "body",
        fontSize: 15,
        color: dark ? colors.mist : colors.cobalt,
        letterSpacing: 1,
        textWrap: "nowrap",
      }}
    >
      {children}
    </Text>
  );
}

export function OutlineBox({
  x,
  y,
  width,
  height,
  color = colors.line,
  radius = 0,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly color?: ColorValue;
  readonly radius?: number;
}) {
  return (
    <Shape
      style={{
        position: "absolute",
        left: x,
        top: y,
        width,
        height,
        background: transparent,
        border: { color, width: 1 },
        borderRadius: radius,
      }}
    />
  );
}
