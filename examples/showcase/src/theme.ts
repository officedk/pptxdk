import { readFileSync } from "node:fs";

import { defineTheme, font } from "@officedk/pptxdk/vanilla";

const jost = font(
  "Jost",
  [400, 700].map((weight) => ({
    weight,
    data: new Uint8Array(
      readFileSync(
        new URL(`./assets/fonts/Jost-${weight}.ttf`, import.meta.url),
      ),
    ),
  })),
);
const jostMedium = font("Jost Medium", [
  {
    weight: 500,
    data: new Uint8Array(
      readFileSync(new URL("./assets/fonts/Jost-500.ttf", import.meta.url)),
    ),
  },
]);

export const colors = {
  paper: "#F7F5F0",
  surface: "#FFFFFF",
  ink: "#25243A",
  accent: "#6C63C8",
  current: "#FFC845",
  soft: "#E7E4F6",
  muted: "#706D7A",
  accentSoft: "#C9C5EC",
  darkSurface: "#38364D",
} as const;

export const motion = {
  reveal: 420,
  settle: 600,
  morph: 1000,
} as const;

export const theme = defineTheme({
  name: "Violet editorial",
  colors,
  fonts: { major: jost, minor: jost, medium: jostMedium, chart: jost },
});
