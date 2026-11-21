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
} from "pptxdk/vanilla";

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
      backgroundColor: colors.surface,
      borderWidth: 1,
      borderColor: colors.soft,
      borderRadius: 6,
    },
  }).append(
    box({ style: { width: 44, height: 5, backgroundColor: colors.accent } }),
    text("01", {
      height: "auto",
      style: {
        fontSize: 60,
        fontFamily: "medium",
        fontWeight: 500,
        color: colors.soft,
      },
    }),
    motionScope(cardScope, { mode: "sequence" }).append(
      text("Lorem ipsum", {
        height: "auto",
        style: { fontSize: 30, fontWeight: 700, color: colors.ink },
        enter: wipe({ from: "left", duration: motion.reveal }),
      }),
      text(
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
        {
          height: "auto",
          style: { fontSize: 20, lineHeight: 28, color: colors.muted },
          enter: floatUp({ duration: motion.reveal }),
        },
      ),
    ),
  );
  firstStep.append(
    box({
      style: { position: "absolute", left: 74, top: 330, width: 410 },
    }).append(card),
  );

  const comparisonStep = step({ n: 2 }).append(
    group({ enter: wipe({ from: "left", duration: motion.settle }) }).append(
      text("A", {
        x: 704,
        y: 316,
        width: 48,
        height: 40,
        style: { fontSize: 25, fontWeight: 700 },
      }),
      shape({ x: 770, y: 320, width: 360, height: 24, fill: colors.soft }),
      shape({ x: 770, y: 320, width: 310, height: 24, fill: colors.accent }),
      text("84", {
        x: 1140,
        y: 310,
        width: 60,
        height: 42,
        style: { fontSize: 25, fontFamily: "medium", fontWeight: 500 },
      }),
      text("B", {
        x: 704,
        y: 434,
        width: 48,
        height: 40,
        style: { fontSize: 25, fontWeight: 700 },
      }),
      shape({ x: 770, y: 438, width: 360, height: 24, fill: colors.soft }),
      shape({ x: 770, y: 438, width: 205, height: 24, fill: colors.ink }),
      text("56", {
        x: 1140,
        y: 428,
        width: 60,
        height: 42,
        style: { fontSize: 25, fontFamily: "medium", fontWeight: 500 },
      }),
      shape({
        x: 1072,
        y: 310,
        width: 8,
        height: 44,
        fill: colors.current,
      }),
    ),
  );

  return slide({ id: "comparison", background: colors.paper }).append(
    shape({ x: 0, y: 0, width: 610, height: 720, fill: colors.accent }),
    text("05 / COMPARISON", {
      x: 72,
      y: 70,
      width: 430,
      height: 30,
      style: {
        fontSize: 16,
        fontWeight: 700,
        letterSpacing: 2,
        color: colors.soft,
      },
    }),
    text(["Same scale.", "\n", "Different weight."], {
      x: 70,
      y: 132,
      width: 480,
      height: 130,
      style: {
        fontSize: 52,
        fontWeight: 700,
        lineHeight: 57,
        color: colors.surface,
      },
    }),
    firstStep,
    text("A / B", {
      x: 700,
      y: 74,
      width: 480,
      height: 104,
      style: {
        fontSize: 88,
        lineHeight: 92,
        fontWeight: 700,
        color: colors.ink,
      },
    }),
    text("One baseline.", {
      x: 706,
      y: 190,
      width: 400,
      height: 44,
      style: {
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
        x: 706,
        y: 564,
        width: 420,
        height: 68,
        style: { fontSize: 18, lineHeight: 26, color: colors.muted },
      },
    ),
  );
}
