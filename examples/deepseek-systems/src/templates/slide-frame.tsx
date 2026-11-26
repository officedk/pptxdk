import {
  Shape,
  Slide,
  SlideNumber,
  Text,
  transitions,
  type Transition,
} from "@officedk/pptxdk/react";
import type { ReactNode } from "react";

import { Rule, type SemanticTone } from "../components/primitives.tsx";
import { colors } from "../theme.ts";

const badgePalette: Record<
  SemanticTone,
  { readonly bg: string; readonly border: string; readonly text: string }
> = {
  compute: {
    bg: colors.computeSoft,
    border: colors.compute,
    text: colors.compute,
  },
  comm: {
    bg: colors.commSoft,
    border: colors.comm,
    text: colors.commText,
  },
  copper: {
    bg: colors.copperSoft,
    border: colors.copper,
    text: colors.copperText,
  },
  baseline: {
    bg: colors.baselineSoft,
    border: colors.baseline,
    text: colors.steel,
  },
  steel: {
    bg: colors.slot,
    border: colors.steel,
    text: colors.steel,
  },
};

/**
 * Shared 1280x720 page frame for Silicon Telemetry analytical slides.
 * Places an action-title header, semantic takeaway banner, and source/slide-number footer.
 */
export function SlideFrame({
  id,
  kicker,
  title,
  takeaway,
  takeawayTone = "compute",
  source,
  dark = false,
  transition = transitions.fade({ duration: 480 }),
  children,
}: {
  readonly id: string;
  readonly kicker: string;
  readonly title: string;
  readonly takeaway?: string;
  readonly takeawayTone?: SemanticTone;
  readonly source: string;
  readonly dark?: boolean;
  readonly transition?: Transition;
  readonly children: ReactNode;
}) {
  const badge = badgePalette[takeawayTone];
  const primaryInk = dark ? colors.surface : colors.ink;
  const mutedInk = dark ? colors.darkMuted : colors.muted;
  const lineColor = dark ? colors.darkLine : colors.line;

  return (
    <Slide
      id={id}
      style={{ background: dark ? colors.obsidian : colors.paper }}
      transition={transition}
    >
      {/* Top kicker bar */}
      <Shape
        style={{
          position: "absolute",
          left: 48,
          top: 22,
          width: 4,
          height: 16,
          background: dark ? colors.computeBright : colors.compute,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 60,
          top: 20,
          width: 670,
          padding: 0,
          fontSize: 12,
          color: dark ? colors.computeBright : colors.compute,
          textWrap: "nowrap",
        }}
      >
        {kicker}
      </Text>

      {/* Top-right takeaway callout */}
      {takeaway !== undefined && (
        <>
          <Shape
            style={{
              position: "absolute",
              left: 736,
              top: 17,
              width: 496,
              height: 28,
              background: dark ? colors.obsidianCard : badge.bg,
              border: {
                color: dark ? colors.darkLine : badge.border,
                width: 1,
              },
            }}
          />
          <Shape
            style={{
              position: "absolute",
              left: 736,
              top: 17,
              width: 4,
              height: 28,
              background: badge.border,
            }}
          />
          <Text
            style={{
              position: "absolute",
              left: 748,
              top: 22,
              width: 476,
              padding: 0,
              fontSize: 11.5,
              color: dark ? colors.surface : badge.text,
              textWrap: "nowrap",
            }}
          >
            {takeaway}
          </Text>
        </>
      )}

      {/* Action title */}
      <Text
        style={{
          position: "absolute",
          left: 48,
          top: 50,
          width: 1184,
          padding: 0,
          fontSize: 24,
          lineHeight: 32,
          color: primaryInk,
          textWrap: "nowrap",
        }}
      >
        {title}
      </Text>

      {/* Structural header divider */}
      <Rule
        x={48}
        y={94}
        width={1184}
        height={2}
        color={dark ? colors.darkLine : colors.steel}
      />

      {children}

      {/* Structural footer divider & metadata */}
      <Rule x={48} y={676} width={1184} height={1} color={lineColor} />
      <Text
        style={{
          position: "absolute",
          left: 48,
          top: 684,
          width: 1060,
          padding: 0,
          fontSize: 11,
          color: mutedInk,
          textWrap: "nowrap",
        }}
      >
        {source}
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 1120,
          top: 684,
          width: 112,
          padding: 0,
          fontFamily: "metric",
          fontSize: 11,
          fontWeight: 700,
          color: mutedInk,
          align: "right",
          textWrap: "nowrap",
        }}
      >
        <SlideNumber /> / 15
      </Text>
    </Slide>
  );
}
