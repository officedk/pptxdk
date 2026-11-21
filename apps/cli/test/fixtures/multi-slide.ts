import { deck, shape, slide } from "pptxdk/vanilla";

export const SpecialIdDeck = deck({ width: 120, height: 96 }).append(
  slide({ id: "sec 1/intro" }).append(
    shape({ id: "s1", x: 8, y: 8, width: 40, height: 24, fill: "#2563EB" }),
  ),
  slide({ id: "sec 2/summary" }).append(
    shape({ id: "s2", x: 8, y: 8, width: 40, height: 24, fill: "#059669" }),
  ),
);

export const HundredPageDeck = deck({ width: 120, height: 96 }).append(
  ...Array.from({ length: 100 }, (_, index) =>
    slide({ id: `slide-${index + 1}` }).append(
      shape({
        id: `box-${index + 1}`,
        x: 4,
        y: 4,
        width: 24,
        height: 16,
        fill: "#334155",
      }),
    ),
  ),
);
