import {
  box,
  createMotionScopeHandle,
  floatUp,
  group,
  motionScope,
  shape,
  slide,
  step,
  text,
  wipe,
} from "@officedk/pptxdk/vanilla";

import { colors, motion } from "../theme.ts";

export function createComparisonPage() {
  const firstStep = step({ n: 1 });
  const cardScope = createMotionScopeHandle();
  const card = box({
    style: {
      display: "flex",
      flexDirection: "column",
      minHeight: 340,
      padding: 32,
      gap: 20,
      background: colors.surface,
      border: { color: colors.soft, width: 1 },
      borderRadius: 6,
    },
  }).append(
    box({ style: { width: 44, height: 5, background: colors.accent } }),
    text("01", {
      style: {
        height: "auto",
        fontSize: 60,
        fontFamily: "medium",
        fontWeight: 500,
        color: colors.soft,
      },
    }),
    motionScope(cardScope, { mode: "sequence" }).append(
      text("Lorem ipsum", {
        style: {
          height: "auto",
          fontSize: 30,
          fontWeight: 700,
          color: colors.ink,
        },
        enter: wipe({ from: "left", duration: motion.reveal }),
      }),
      text(
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
        {
          style: {
            height: "auto",
            fontSize: 20,
            lineHeight: 28,
            color: colors.muted,
          },
          enter: floatUp({ duration: motion.reveal }),
        },
      ),
    ),
  );
  const firstCardSlot = box({
    style: { position: "absolute", left: 74, top: 330, width: 410 },
  }).append(firstStep.append(card));

  const comparisonStep = step({ n: 2 }).append(
    group({ enter: wipe({ from: "left", duration: motion.settle }) }).append(
      text("A", {
        style: {
          position: "absolute",
          left: 704,
          top: 316,
          width: 48,
          height: 40,
          fontSize: 25,
          fontWeight: 700,
        },
      }),
      shape({
        style: {
          position: "absolute",
          left: 770,
          top: 320,
          width: 360,
          height: 24,
          background: colors.soft,
        },
      }),
      shape({
        style: {
          position: "absolute",
          left: 770,
          top: 320,
          width: 310,
          height: 24,
          background: colors.accent,
        },
      }),
      text("84", {
        style: {
          position: "absolute",
          left: 1140,
          top: 310,
          width: 60,
          height: 42,
          fontSize: 25,
          fontFamily: "medium",
          fontWeight: 500,
        },
      }),
      text("B", {
        style: {
          position: "absolute",
          left: 704,
          top: 434,
          width: 48,
          height: 40,
          fontSize: 25,
          fontWeight: 700,
        },
      }),
      shape({
        style: {
          position: "absolute",
          left: 770,
          top: 438,
          width: 360,
          height: 24,
          background: colors.soft,
        },
      }),
      shape({
        style: {
          position: "absolute",
          left: 770,
          top: 438,
          width: 205,
          height: 24,
          background: colors.ink,
        },
      }),
      text("56", {
        style: {
          position: "absolute",
          left: 1140,
          top: 428,
          width: 60,
          height: 42,
          fontSize: 25,
          fontFamily: "medium",
          fontWeight: 500,
        },
      }),
      shape({
        style: {
          position: "absolute",
          left: 1072,
          top: 310,
          width: 8,
          height: 44,
          background: colors.current,
        },
      }),
    ),
  );

  return slide({ id: "comparison", background: colors.paper }).append(
    shape({
      style: {
        position: "absolute",
        left: 0,
        top: 0,
        width: 610,
        height: 720,
        background: colors.accent,
      },
    }),
    text("05 / COMPARISON", {
      style: {
        position: "absolute",
        left: 72,
        top: 70,
        width: 430,
        height: 30,
        fontSize: 16,
        fontWeight: 700,
        letterSpacing: 2,
        color: colors.soft,
      },
    }),
    text(["Same scale.", "\n", "Different weight."], {
      style: {
        position: "absolute",
        left: 70,
        top: 132,
        width: 480,
        height: 130,
        fontSize: 52,
        fontWeight: 700,
        lineHeight: 57,
        color: colors.surface,
      },
    }),
    firstCardSlot,
    text("A / B", {
      style: {
        position: "absolute",
        left: 700,
        top: 74,
        width: 480,
        height: 104,
        fontSize: 88,
        lineHeight: 92,
        fontWeight: 700,
        color: colors.ink,
      },
    }),
    text("One baseline.", {
      style: {
        position: "absolute",
        left: 706,
        top: 190,
        width: 400,
        height: 44,
        fontSize: 30,
        fontFamily: "medium",
        fontWeight: 500,
        color: colors.accent,
      },
    }),
    comparisonStep,
    text(
      "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
      {
        style: {
          position: "absolute",
          left: 706,
          top: 564,
          width: 420,
          height: 68,
          fontSize: 18,
          lineHeight: 26,
          color: colors.muted,
        },
      },
    ),
  );
}
