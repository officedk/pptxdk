import { deck, shape, slide } from "@officedk/pptxdk/vanilla";

export const SpecialIdDeck = deck({ width: 120, height: 96 }).append(
  slide({ id: "sec 1/intro" }).append(
    shape({
      id: "s1",
      style: {
        position: "absolute",
        left: 8,
        top: 8,
        width: 40,
        height: 24,
        background: "#2563EB",
      },
    }),
  ),
  slide({ id: "sec 2/summary" }).append(
    shape({
      id: "s2",
      style: {
        position: "absolute",
        left: 8,
        top: 8,
        width: 40,
        height: 24,
        background: "#059669",
      },
    }),
  ),
);

export const HundredPageDeck = deck({ width: 120, height: 96 }).append(
  ...Array.from({ length: 100 }, (_, index) =>
    slide({ id: `slide-${index + 1}` }).append(
      shape({
        id: `box-${index + 1}`,
        style: {
          position: "absolute",
          left: 4,
          top: 4,
          width: 24,
          height: 16,
          background: "#334155",
        },
      }),
    ),
  ),
);

export const SixPageDeck = deck({ width: 120, height: 96 }).append(
  ...Array.from({ length: 6 }, (_, index) =>
    slide({ id: `six-${index + 1}` }).append(
      shape({
        id: `six-box-${index + 1}`,
        style: {
          position: "absolute",
          left: 4 + index * 4,
          top: 4,
          width: 24,
          height: 16,
          background: "#2563EB",
        },
      }),
    ),
  ),
);

export const OneTwentyPageDeck = deck({ width: 120, height: 96 }).append(
  ...Array.from({ length: 120 }, (_, index) =>
    slide({ id: `deck-120-slide-${index + 1}` }).append(
      shape({
        id: `deck-120-box-${index + 1}`,
        style: {
          position: "absolute",
          left: 4 + (index % 10) * 4,
          top: 4 + Math.floor(index / 10) * 4,
          width: 24,
          height: 16,
          background: "#0F766E",
        },
      }),
    ),
  ),
);
