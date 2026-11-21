import { path, preset } from "pptxdk";
import { Shape, Text } from "pptxdk/react";

import { colors, transparent } from "../theme.ts";
import { Circle } from "./primitives.tsx";

function noise(index: number, salt: number) {
  const value = Math.sin(index * 91.731 + salt * 17.17) * 43758.5453;
  return value - Math.floor(value);
}

export function SamplingField({
  x,
  y,
  size,
  dark = false,
}: {
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly dark?: boolean;
}) {
  const ink = dark ? colors.mist : colors.cobalt;
  const center = x + size / 2;
  const middle = y + size / 2;
  return (
    <>
      {[0.94, 0.76, 0.58, 0.4].map((ratio, index) => (
        <Shape
          key={ratio}
          geometry={preset("ellipse")}
          x={center - (size * ratio) / 2}
          y={middle - (size * ratio * 0.68) / 2}
          width={size * ratio}
          height={size * ratio * 0.68}
          fill={transparent}
          style={{
            borderColor: index === 2 ? colors.sun : ink,
            borderWidth: index === 2 ? 2 : 1,
            opacity: index === 2 ? 0.9 : 0.35 + index * 0.12,
            rotate: index % 2 === 0 ? -12 : 9,
          }}
        />
      ))}
      {Array.from({ length: 42 }, (_, index) => {
        const angle = noise(index, 1) * Math.PI * 2;
        const radius = size * (0.12 + noise(index, 2) * 0.4);
        const dot = 2 + noise(index, 3) * 5;
        return (
          <Circle
            key={index}
            x={center + Math.cos(angle) * radius - dot / 2}
            y={middle + Math.sin(angle) * radius * 0.65 - dot / 2}
            diameter={dot}
            fill={index % 9 === 0 ? colors.sun : ink}
          />
        );
      })}
      <Shape
        geometry={path("M 0 40 C 40 0 80 80 120 40 C 160 0 200 80 240 40", {
          viewBox: [240, 80],
          fill: false,
        })}
        x={center - size * 0.3}
        y={middle - size * 0.075}
        width={size * 0.6}
        height={size * 0.15}
        fill={transparent}
        style={{ borderColor: colors.cyan, borderWidth: 2 }}
      />
      <Shape
        x={center - 34}
        y={middle - 34}
        width={68}
        height={68}
        fill={colors.sun}
        style={{ rotate: 45 }}
      />
      <Text
        x={center - 34}
        y={middle - 39}
        width={68}
        height={68}
        style={{
          fontFamily: "display",
          fontSize: 39,
          color: colors.indigo,
          align: "center",
          verticalAlign: "middle",
        }}
      >
        生
      </Text>
    </>
  );
}
