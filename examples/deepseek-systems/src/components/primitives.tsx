import {
  Group,
  MotionScope,
  preset,
  Shape,
  Text,
  useMotionScope,
  type EntranceEffectDeclaration,
  type ShapeProps,
  type TextStyle,
} from "@officedk/pptxdk/react";
import type { ReactNode } from "react";

import { colors } from "../theme.ts";

export type SemanticTone = "compute" | "comm" | "copper" | "baseline" | "steel";

const tonePalette: Record<
  SemanticTone,
  {
    readonly accent: string;
    readonly soft: string;
    readonly text: string;
  }
> = {
  compute: {
    accent: colors.compute,
    soft: colors.computeSoft,
    text: colors.compute,
  },
  comm: {
    accent: colors.comm,
    soft: colors.commSoft,
    text: colors.commText,
  },
  copper: {
    accent: colors.copper,
    soft: colors.copperSoft,
    text: colors.copperText,
  },
  baseline: {
    accent: colors.baseline,
    soft: colors.baselineSoft,
    text: colors.steel,
  },
  steel: {
    accent: colors.steel,
    soft: colors.slot,
    text: colors.steel,
  },
};

/** 1px or custom-thickness horizontal/vertical structural rule. */
export function Rule({
  x,
  y,
  width,
  height = 1,
  color = colors.line,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height?: number;
  readonly color?: string;
}) {
  return (
    <Shape
      style={{
        position: "absolute",
        left: x,
        top: y,
        width,
        height,
        background: color,
      }}
    />
  );
}

/** Keeps a separately editable fill and label on one entrance moment in its parent's stagger. */
export function EnteredShapeText({
  id,
  x,
  y,
  width,
  height,
  background,
  shapeStyle,
  text,
  textStyle,
  enter,
}: {
  readonly id: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly background: NonNullable<
    NonNullable<ShapeProps["style"]>["background"]
  >;
  readonly shapeStyle?: NonNullable<ShapeProps["style"]>;
  readonly text: string;
  readonly textStyle: TextStyle;
  readonly enter: EntranceEffectDeclaration;
}) {
  const scope = useMotionScope();

  return (
    <MotionScope scope={scope}>
      <Shape
        id={`${id}-background`}
        style={{
          position: "absolute",
          left: x,
          top: y,
          width,
          height,
          background,
          ...shapeStyle,
        }}
        enter={enter}
      />
      <Text
        id={id}
        style={{
          position: "absolute",
          left: x,
          top: y,
          width,
          height,
          ...textStyle,
        }}
        enter={enter}
      >
        {text}
      </Text>
    </MotionScope>
  );
}

/** Elevated wafer-white surface panel with 1px structural hairline border and optional top/left semantic accent bar. */
export function PanelCard({
  x,
  y,
  width,
  height,
  title,
  subtitle,
  rightBadge,
  tone = "compute",
  accentSide = "top",
  background = colors.surface,
  borderColor = colors.line,
  id,
  children,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly title?: string;
  readonly subtitle?: string;
  readonly rightBadge?: string;
  readonly tone?: SemanticTone;
  readonly accentSide?: "top" | "left" | "none";
  readonly background?: string;
  readonly borderColor?: string;
  readonly id?: string;
  readonly children?: ReactNode;
}) {
  const palette = tonePalette[tone];
  return (
    <>
      <Shape
        {...(id === undefined ? {} : { id })}
        style={{
          position: "absolute",
          left: x,
          top: y,
          width,
          height,
          background,
          border: { color: borderColor, width: 1 },
        }}
      />
      {accentSide === "top" && (
        <Shape
          style={{
            position: "absolute",
            left: x,
            top: y,
            width,
            height: 3,
            background: palette.accent,
          }}
        />
      )}
      {accentSide === "left" && (
        <Shape
          style={{
            position: "absolute",
            left: x,
            top: y,
            width: 4,
            height,
            background: palette.accent,
          }}
        />
      )}
      {title !== undefined && (
        <Text
          style={{
            position: "absolute",
            left: x + 18,
            top: y + 12,
            width: rightBadge ? width - 154 : width - 36,
            padding: 0,
            fontSize: 14.5,
            color: colors.steel,
            textWrap: "nowrap",
          }}
        >
          {title}
        </Text>
      )}
      {rightBadge !== undefined && (
        <Shape
          style={{
            position: "absolute",
            left: x + width - 136,
            top: y + 11,
            width: 120,
            height: 22,
            padding: [0, 4],
            background: palette.soft,
            fontSize: 10.5,
            color: palette.text,
            align: "center",
            verticalAlign: "middle",
          }}
        >
          {rightBadge}
        </Shape>
      )}
      {subtitle !== undefined && (
        <Text
          style={{
            position: "absolute",
            left: x + 18,
            top: y + 36,
            width: width - 32,
            padding: 0,
            fontSize: 11,
            color: colors.muted,
          }}
        >
          {subtitle}
        </Text>
      )}
      {children}
    </>
  );
}

/** Compact quantitative KPI card with semantic left border and Jost bold metric value. */
export function MetricCard({
  x,
  y,
  width,
  height,
  label,
  value,
  caption,
  tone = "compute",
  dark = false,
  id,
  enter,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly label: string;
  readonly value: string;
  readonly caption: string;
  readonly tone?: SemanticTone;
  readonly dark?: boolean;
  readonly id?: string;
  readonly enter?: EntranceEffectDeclaration;
}) {
  const palette = tonePalette[tone];
  const valueColor = dark
    ? tone === "copper"
      ? colors.copperBright
      : tone === "comm"
        ? colors.commBright
        : colors.computeBright
    : palette.text;
  const card = (
    <>
      <Shape
        {...(id === undefined ? {} : { id })}
        style={{
          position: "absolute",
          left: x,
          top: y,
          width,
          height,
          background: dark ? colors.obsidianCard : colors.surface,
          border: {
            color: dark ? colors.darkLine : colors.line,
            width: 1,
          },
        }}
      />
      <Shape
        style={{
          position: "absolute",
          left: x,
          top: y,
          width: 4,
          height,
          background: palette.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: x + 14,
          top: y + 10,
          width: width - 20,
          padding: 0,
          fontSize: 11.5,
          color: dark ? colors.darkMuted : colors.muted,
          textWrap: "nowrap",
        }}
      >
        {label}
      </Text>
      <Text
        style={{
          position: "absolute",
          left: x + 14,
          top: y + 28,
          width: width - 20,
          padding: 0,
          fontFamily: "metric",
          fontSize: 21,
          fontWeight: 700,
          color: valueColor,
          textWrap: "nowrap",
        }}
      >
        {value}
      </Text>
      <Text
        style={{
          position: "absolute",
          left: x + 14,
          top: y + height - 24,
          width: width - 20,
          padding: 0,
          fontSize: 10.5,
          lineHeight: 14,
          color: dark ? colors.darkSubtle : colors.muted,
        }}
      >
        {caption}
      </Text>
    </>
  );
  return enter === undefined ? card : <Group enter={enter}>{card}</Group>;
}

/** Semantic pill / tag for pipeline stages, operators, or legend callouts. */
export function SemanticTag({
  x,
  y,
  width,
  height = 24,
  text,
  tone = "compute",
  solid = false,
  rounded = false,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height?: number;
  readonly text: string;
  readonly tone?: SemanticTone;
  readonly solid?: boolean;
  readonly rounded?: boolean;
}) {
  const palette = tonePalette[tone];
  return (
    <Shape
      {...(rounded ? { geometry: preset("roundRect") } : {})}
      style={{
        position: "absolute",
        left: x,
        top: y,
        width,
        height,
        padding: [0, 4],
        background: solid ? palette.accent : palette.soft,
        ...(solid
          ? {}
          : {
              border: { color: palette.accent, width: 1 },
            }),
        fontSize: 11,
        color: solid ? colors.surface : palette.text,
        align: "center",
        verticalAlign: "middle",
      }}
    >
      {text}
    </Shape>
  );
}
