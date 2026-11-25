import {
  blockArc,
  Group,
  preset,
  Shape,
  Step,
  Text,
  wipe,
  zoom,
} from "@officedk/pptxdk/react";

import { colors, motion } from "../theme.ts";

/** Values, segmented bars, and rings share the same numerator and denominator; x, y, width specify slide position and total width, composition scales proportionally with width. */
export function Progress({
  completed,
  total,
  x,
  y,
  width,
}: {
  readonly completed: number;
  readonly total: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
}) {
  const fraction = completed / total;
  const scale = width / 1120;
  const size = (value: number) => value * scale;
  const localX = (offset: number) => x + size(offset);
  const localY = (offset: number) => y + size(offset);
  const panelWidth = size(366);
  const panelX = x + width - panelWidth;
  const segmentGap = size(10);
  const segmentWidth = (size(622) - (total - 1) * segmentGap) / total;
  return (
    <Group id="progress-example">
      <Text
        x={x}
        y={localY(56)}
        width={size(570)}
        height={size(160)}
        style={{
          fontSize: size(142),
          lineHeight: size(145),
          fontWeight: 700,
          color: colors.surface,
        }}
      >
        {Math.round(fraction * 100)}%
      </Text>
      <Text
        x={localX(4)}
        y={localY(226)}
        width={size(180)}
        height={size(48)}
        style={{
          fontSize: size(28),
          fontFamily: "medium",
          fontWeight: 500,
          color: colors.accentSoft,
        }}
      >
        {completed} / {total}
      </Text>
      <Text
        x={localX(234)}
        y={localY(224)}
        width={size(388)}
        height={size(78)}
        style={{
          fontSize: size(18),
          lineHeight: size(27),
          color: colors.accentSoft,
        }}
      >
        Lorem ipsum dolor sit amet, consectetur adipiscing elit.
      </Text>
      {Array.from({ length: total }, (_, index) => (
        <Shape
          key={index}
          x={localX(4) + index * (segmentWidth + segmentGap)}
          y={localY(350)}
          width={segmentWidth}
          height={size(12)}
          fill={colors.darkSurface}
          geometry={preset("roundRect")}
          style={{ borderRadius: size(4) }}
        />
      ))}
      <Step n={1}>
        <Group enter={wipe({ from: "left", duration: motion.settle })}>
          {Array.from({ length: completed }, (_, index) => (
            <Shape
              key={index}
              x={localX(4) + index * (segmentWidth + segmentGap)}
              y={localY(350)}
              width={segmentWidth}
              height={size(12)}
              fill={index === completed - 1 ? colors.current : colors.accent}
              geometry={preset("roundRect")}
              style={{ borderRadius: size(4) }}
            />
          ))}
        </Group>
      </Step>
      {Array.from({ length: total }, (_, index) => (
        <Text
          key={index}
          x={localX(4) + index * (segmentWidth + segmentGap)}
          y={localY(382)}
          width={segmentWidth}
          height={size(28)}
          style={{
            fontSize: size(16),
            align: "center",
            color:
              index === completed - 1
                ? colors.current
                : index < completed
                  ? colors.accentSoft
                  : colors.muted,
          }}
        >
          {String(index + 1).padStart(2, "0")}
        </Text>
      ))}
      <Shape
        x={panelX}
        y={y}
        width={panelWidth}
        height={size(468)}
        fill={colors.darkSurface}
        geometry={preset("roundRect")}
        style={{ borderRadius: size(6) }}
      />
      <Shape
        x={panelX + size(40)}
        y={localY(40)}
        width={size(40)}
        height={size(4)}
        fill={colors.current}
      />
      <Shape
        x={panelX + size(53)}
        y={localY(78)}
        width={size(250)}
        height={size(250)}
        fill={colors.ink}
        geometry={preset("ellipse")}
      />
      <Shape
        x={panelX + size(81)}
        y={localY(106)}
        width={size(194)}
        height={size(194)}
        fill={colors.ink}
        geometry={preset("ellipse")}
      />
      <Step n={2}>
        <Group enter={zoom({ duration: motion.settle })}>
          <Shape
            x={panelX + size(53)}
            y={localY(78)}
            width={size(250)}
            height={size(250)}
            fill={colors.current}
            geometry={blockArc({
              startAngle: 270,
              endAngle: 270 + fraction * 360,
              thickness: 0.115,
            })}
          />
        </Group>
      </Step>
      <Text
        x={panelX + size(88)}
        y={localY(156)}
        width={size(180)}
        height={size(72)}
        style={{
          fontSize: size(58),
          lineHeight: size(66),
          fontFamily: "medium",
          fontWeight: 500,
          align: "center",
          color: colors.surface,
        }}
      >
        {completed}/{total}
      </Text>
      <Text
        x={panelX + size(38)}
        y={localY(374)}
        width={size(290)}
        height={size(34)}
        style={{
          fontSize: size(17),
          align: "center",
          color: colors.accentSoft,
        }}
      >
        Lorem ipsum dolor
      </Text>
    </Group>
  );
}
