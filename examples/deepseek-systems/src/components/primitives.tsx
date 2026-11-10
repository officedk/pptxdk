import {
  MotionScope,
  Shape,
  Text,
  preset,
  useMotionScope,
} from "@react-pptx/react-pptx";
import type {
  EntranceEffectDeclaration,
  ShapeProps,
  TextStyle,
} from "@react-pptx/react-pptx";
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
  return <Shape x={x} y={y} width={width} height={height} fill={color} />;
}

/** Keeps a separately editable fill and label on one entrance moment in its parent's stagger. */
export function EnteredShapeText({
  testId,
  x,
  y,
  width,
  height,
  fill,
  shapeStyle,
  text,
  textStyle,
  enter,
}: {
  readonly testId: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly fill: NonNullable<ShapeProps["fill"]>;
  readonly shapeStyle?: NonNullable<ShapeProps["style"]>;
  readonly text: string;
  readonly textStyle: TextStyle;
  readonly enter: EntranceEffectDeclaration;
}) {
  const scope = useMotionScope();

  return (
    <MotionScope scope={scope}>
      <Shape
        testId={`${testId}-background`}
        x={x}
        y={y}
        width={width}
        height={height}
        fill={fill}
        {...(shapeStyle === undefined ? {} : { style: shapeStyle })}
        enter={enter}
      />
      <Text
        testId={testId}
        x={x}
        y={y}
        width={width}
        height={height}
        style={textStyle}
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
  fill = colors.surface,
  borderColor = colors.line,
  testId,
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
  readonly fill?: string;
  readonly borderColor?: string;
  readonly testId?: string;
  readonly children?: ReactNode;
}) {
  const palette = tonePalette[tone];
  return (
    <>
      <Shape
        {...(testId === undefined ? {} : { testId })}
        x={x}
        y={y}
        width={width}
        height={height}
        fill={fill}
        style={{ borderColor, borderWidth: 1 }}
      />
      {accentSide === "top" && (
        <Shape x={x} y={y} width={width} height={3} fill={palette.accent} />
      )}
      {accentSide === "left" && (
        <Shape x={x} y={y} width={4} height={height} fill={palette.accent} />
      )}
      {title !== undefined && (
        <Text
          x={x + 18}
          y={y + 12}
          width={rightBadge ? width - 154 : width - 36}
          style={{
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
          x={x + width - 136}
          y={y + 11}
          width={120}
          height={22}
          fill={palette.soft}
          style={{ padding: [0, 4] }}
          text={rightBadge}
          textStyle={{
            fontSize: 10.5,
            color: palette.text,
            align: "center",
            verticalAlign: "middle",
          }}
        />
      )}
      {subtitle !== undefined && (
        <Text
          x={x + 18}
          y={y + 36}
          width={width - 32}
          style={{
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
  testId,
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
  readonly testId?: string;
}) {
  const palette = tonePalette[tone];
  const valueColor = dark
    ? tone === "copper"
      ? colors.copperBright
      : tone === "comm"
        ? colors.commBright
        : colors.computeBright
    : palette.text;
  return (
    <>
      <Shape
        {...(testId === undefined ? {} : { testId })}
        x={x}
        y={y}
        width={width}
        height={height}
        fill={dark ? colors.obsidianCard : colors.surface}
        style={{
          borderColor: dark ? colors.darkLine : colors.line,
          borderWidth: 1,
        }}
      />
      <Shape x={x} y={y} width={4} height={height} fill={palette.accent} />
      <Text
        x={x + 14}
        y={y + 10}
        width={width - 20}
        style={{
          padding: 0,
          fontSize: 11.5,
          color: dark ? colors.darkMuted : colors.muted,
          textWrap: "nowrap",
        }}
      >
        {label}
      </Text>
      <Text
        x={x + 14}
        y={y + 28}
        width={width - 20}
        style={{
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
        x={x + 14}
        y={y + height - 24}
        width={width - 20}
        style={{
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
      x={x}
      y={y}
      width={width}
      height={height}
      fill={solid ? palette.accent : palette.soft}
      {...(rounded ? { geometry: preset("roundRect") } : {})}
      style={
        solid
          ? { padding: [0, 4] }
          : {
              padding: [0, 4],
              borderColor: palette.accent,
              borderWidth: 1,
            }
      }
      text={text}
      textStyle={{
        fontSize: 11,
        color: solid ? colors.surface : palette.text,
        align: "center",
        verticalAlign: "middle",
      }}
    />
  );
}
