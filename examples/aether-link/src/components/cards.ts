import {
  box,
  paragraph,
  preset,
  shape,
  span,
  text,
  type BoxNode,
  type EntranceEffectDeclaration,
  type ThemeColor,
} from "@officedk/pptxdk/vanilla";

import { PALETTE, displayFont, jostFont, jostMediumFont } from "../theme.ts";

/** Props for the closed KPI metric tile built with Vanilla primitives. */
export interface MetricTileProps {
  readonly id?: string | undefined;
  readonly label: string;
  readonly value: string;
  readonly unit?: string | undefined;
  readonly detail: string;
  readonly accentColor?: ThemeColor | undefined;
  readonly dark?: boolean | undefined;
  readonly compact?: boolean | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Props for the closed bottom takeaway strip. */
export interface TakeawayStripProps {
  readonly id?: string | undefined;
  readonly tag: string;
  readonly text: string;
  readonly accentColor?: ThemeColor | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Props for the slotted telemetry panel container returning `{ node, slot }`. */
export interface TelemetryPanelProps {
  readonly id?: string | undefined;
  readonly kicker?: string | undefined;
  readonly title: string;
  readonly badge?: string | undefined;
  readonly accentColor?: ThemeColor | undefined;
  readonly badgeBackground?: ThemeColor | undefined;
  readonly badgeColor?: ThemeColor | undefined;
  readonly dark?: boolean | undefined;
  readonly flex?: number | undefined;
  readonly width?: number | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Return shape for slotted Vanilla container components adopted by `fromVanilla`. */
export interface TelemetryPanelResult {
  readonly node: BoxNode;
  readonly slot: BoxNode;
}

/** Builds a closed KPI metric tile participating in parent Flex/Grid layout and MotionScope. */
export function createMetricTile(props: MetricTileProps): BoxNode {
  const accent = props.accentColor ?? PALETTE.vermilion;
  const bg = props.dark ? PALETTE.obsidianCard : PALETTE.surface;
  const border = props.dark ? PALETTE.obsidianBorder : PALETTE.borderLight;
  const valColor = props.dark ? PALETTE.ivoryText : PALETTE.ink;
  const labelColor = props.dark ? PALETTE.slateLight : PALETTE.inkMuted;
  const detailColor = props.dark ? PALETTE.slateLight : PALETTE.inkBody;

  const valueParagraph = paragraph(
    props.unit
      ? [
          props.value,
          span(` ${props.unit}`, {
            style: {
              fontFamily: jostMediumFont,
              fontSize: props.compact ? 14 : 16,
              fontWeight: 500,
              color: accent,
            },
          }),
        ]
      : props.value,
  );

  return box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      flex: 1,
      flexBasis: 0,
      minWidth: 0,
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      padding: props.compact ? [13, 16, 13, 16] : [16, 18, 16, 18],
      background: bg,
      border: { color: border, width: 1, style: "solid" },
      borderRadius: 10,
    },
  }).append(
    box({
      style: {
        display: "flex",
        flexDirection: "row",
        alignItems: "center",
        gap: 8,
      },
    }).append(
      shape({
        geometry: preset("rect"),
        style: { width: 4, height: 14, background: accent, borderRadius: 2 },
      }),
      text(props.label, {
        style: {
          flex: 1,
          height: "auto",
          fontFamily: jostFont,
          fontSize: props.compact ? 10 : 11,
          fontWeight: 700,
          letterSpacing: 1.1,
          color: labelColor,
        },
      }),
    ),
    text(valueParagraph, {
      style: {
        width: "100%",
        height: "auto",
        fontFamily: displayFont,
        fontSize: props.compact ? 35 : 42,
        fontWeight: 600,
        lineHeight: props.compact ? "36px" : "44px",
        color: valColor,
      },
    }),
    text(props.detail, {
      style: {
        width: "100%",
        height: "auto",
        fontFamily: jostFont,
        fontSize: props.compact ? 11 : 12,
        lineHeight: props.compact ? "15px" : "17px",
        color: detailColor,
      },
    }),
  );
}

/** Builds a closed full-width takeaway summary strip with a tagged pill and assertion copy. */
export function createTakeawayStrip(props: TakeawayStripProps): BoxNode {
  const accent = props.accentColor ?? PALETTE.cobalt;
  return box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      display: "flex",
      flexDirection: "row",
      alignItems: "center",
      gap: 14,
      padding: [10, 16, 10, 16],
      background: PALETTE.paperAlt,
      border: { color: PALETTE.borderLight, width: 1, style: "solid" },
      borderRadius: 8,
    },
  }).append(
    shape(
      paragraph(props.tag, {
        style: { align: "center" },
      }),
      {
        geometry: preset("roundRect"),
        style: {
          width: 124,
          height: 24,
          background: accent,
          borderRadius: 6,
          fontFamily: jostFont,
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 1.1,
          color: "#FFFFFF",
          verticalAlign: "middle",
        },
      },
    ),
    text(props.text, {
      style: {
        flex: 1,
        height: "auto",
        fontFamily: jostMediumFont,
        fontSize: 13,
        fontWeight: 500,
        lineHeight: "18px",
        color: PALETTE.ink,
      },
    }),
  );
}

/**
 * Builds a slotted telemetry card shell with an imperative header bar, optional badge pill,
 * and an exposed flex `slot` container for JSX children (`<Chart />`, `<Mount />`, etc.).
 */
export function createTelemetryPanel(
  props: TelemetryPanelProps,
): TelemetryPanelResult {
  const accent = props.accentColor ?? PALETTE.vermilion;
  const bg = props.dark ? PALETTE.obsidianCard : PALETTE.surface;
  const border = props.dark ? PALETTE.obsidianBorder : PALETTE.borderLight;
  const titleColor = props.dark ? PALETTE.ivoryText : PALETTE.ink;

  const headerChildren = [
    shape({
      geometry: preset("rect"),
      style: { width: 4, height: 18, background: accent, borderRadius: 2 },
    }),
    text(
      props.kicker
        ? [
            paragraph(props.kicker, {
              style: {
                fontFamily: jostFont,
                fontSize: 9.5,
                fontWeight: 700,
                letterSpacing: 1.0,
                color: accent,
              },
            }),
            paragraph(props.title, {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                fontWeight: 700,
                color: titleColor,
              },
            }),
          ]
        : props.title,
      {
        style: {
          flex: 1,
          height: "auto",
          fontFamily: jostFont,
          fontSize: 13,
          fontWeight: 700,
          color: titleColor,
        },
      },
    ),
  ];

  if (props.badge !== undefined) {
    headerChildren.push(
      shape(paragraph(props.badge, { style: { align: "center" } }), {
        geometry: preset("roundRect"),
        style: {
          width: 146,
          height: 24,
          background: props.badgeBackground ?? PALETTE.paperAlt,
          border: { color: border, width: 1, style: "solid" },
          borderRadius: 6,
          fontFamily: jostFont,
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 0.6,
          color: props.badgeColor ?? PALETTE.inkBody,
          verticalAlign: "middle",
        },
      }),
    );
  }

  const slot = box({
    style: {
      flex: 1,
      minHeight: 0,
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      gap: 10,
    },
  });

  const node = box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      ...(props.width === undefined
        ? { flex: props.flex ?? 1, flexBasis: 0, minWidth: 0 }
        : { width: props.width }),
      display: "flex",
      flexDirection: "column",
      gap: 12,
      padding: [16, 18, 16, 18],
      background: bg,
      border: { color: border, width: 1, style: "solid" },
      borderRadius: 12,
    },
  }).append(
    box({
      style: {
        display: "flex",
        flexDirection: "row",
        alignItems: "center",
        gap: 10,
      },
    }).append(...headerChildren),
    slot,
  );

  return { node, slot };
}
