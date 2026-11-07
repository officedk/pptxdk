import { Shape, Text, preset } from "@react-pptx/react-pptx";
import type { ColorValue } from "@react-pptx/react-pptx";
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
  return <Shape x={x} y={y} width={width} height={thickness} fill={color} />;
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
      x={x}
      y={y}
      width={diameter}
      height={diameter}
      fill={fill}
      {...(stroke === undefined
        ? {}
        : { style: { borderColor: stroke, borderWidth: strokeWidth } })}
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
    <>
      <Rule x={x} y={y + 7} width={width - 9} color={color} thickness={2} />
      <Shape
        geometry={preset("triangle")}
        x={x + width - 12}
        y={y}
        width={14}
        height={16}
        fill={color}
        style={{ rotate: 90 }}
      />
    </>
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
      x={x}
      y={y}
      width={width}
      height={28}
      style={{
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
      x={x}
      y={y}
      width={width}
      height={height}
      fill={transparent}
      style={{ borderColor: color, borderWidth: 1, borderRadius: radius }}
    />
  );
}
