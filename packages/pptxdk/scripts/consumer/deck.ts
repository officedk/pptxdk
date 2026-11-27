import { readFileSync } from "node:fs";

import {
  box,
  cell,
  chart,
  darker,
  deck,
  defaultColorMap,
  defineTheme,
  font,
  group,
  isThemeFontSlot,
  lighter,
  row,
  shape,
  slide,
  table,
  themeFontSlots,
  type ColorMapOverride,
  type FontReference,
  type HexColor,
  type NativeColorSlot,
  type SchemeColorName,
  type Theme,
  type ThemeColor,
  type ThemeFontSlot,
  type ThemeInput,
  type ThemeOverride,
} from "@officedk/pptxdk/vanilla";

const family = font("Jost*", [
  { data: readFileSync(new URL("./Jost.ttf", import.meta.url)) },
]);
const sharedPadding = [8, 12] as const;

const primarySlot: NativeColorSlot = "accent1";
const primaryHex: HexColor = "#1d4ed8";
const schemeText: SchemeColorName = "tx1";
const schemeBg: ThemeColor = "bg1";
const accentShadeLight: ThemeColor = lighter("accent2", 0.85);
const accentShadeDark: ThemeColor = darker("accent1", 0.25);
const majorFontSlot: ThemeFontSlot = themeFontSlots[0];
const minorFontRef: FontReference = isThemeFontSlot("minor")
  ? "minor"
  : "major";

const consumerThemeInput: ThemeInput = {
  name: "Consumer",
  colors: {
    dk1: "#111827",
    lt1: "#fff",
    dk2: "#1e293b",
    lt2: "#f8fafc",
    accent1: primaryHex,
    accent2: "#abc",
    accent3: "#0f766e",
    accent4: "#b45309",
    accent5: "#7c3aed",
    accent6: "#be123c",
    hlink: "#2563eb",
    folHlink: "#6d28d9",
  },
  fonts: { major: family, minor: family },
};
const consumerTheme: Theme = defineTheme(consumerThemeInput);

const slideThemeOverride: ThemeOverride = {
  colors: { accent1: "#0f766e", accent2: "#abc" },
  fonts: { major: family, minor: family },
};
const slideColorMapOverride: ColorMapOverride = {
  ...defaultColorMap,
  tx1: primarySlot,
};

function verifyClosedThemeTypes(): void {
  // @ts-expect-error misspelled native theme color slot must fail type checking
  defineTheme({ colors: { followHlink: "#112233" } });
  // @ts-expect-error slot-to-slot alias in theme colors must fail type checking
  defineTheme({ colors: { dk1: "accent2" } });
  // @ts-expect-error ThemeShade in theme colors must fail type checking
  defineTheme({ colors: { accent1: lighter("accent2", 0.4) } });
  // @ts-expect-error misspelled theme font slot must fail type checking
  defineTheme({ fonts: { minro: "Aptos" } });
  slide({
    id: "bad-ovr",
    themeOverride: {
      colors: {
        // @ts-expect-error misspelled ThemeOverride color slot must fail type checking
        followHlink: "#112233",
      },
    },
  });
  // @ts-expect-error misspelled ThemeOverride font slot must fail type checking
  slide({ id: "bad-font-ovr", themeOverride: { fonts: { minro: "Aptos" } } });
  // @ts-expect-error CSS color keyword must fail type checking
  shape("x", { style: { color: "red" } });
  // @ts-expect-error custom color token name must fail type checking
  shape("x", { style: { color: "tileAccent" } });
  // @ts-expect-error misspelled scheme color name must fail type checking
  shape("x", { style: { color: "acent1" } });
  // @ts-expect-error custom background color token name must fail type checking
  shape("x", { style: { background: "tileAccent" } });
  // @ts-expect-error removed container theme prop on slide must fail type checking
  slide({ id: "s", theme: {} });
  // @ts-expect-error removed container theme prop on box must fail type checking
  box({ theme: {} });
  // @ts-expect-error removed container theme prop on group must fail type checking
  group({ theme: {} });
}
void verifyClosedThemeTypes;

export default deck({
  width: 640,
  height: 360,
  theme: consumerTheme,
}).append(
  slide({
    id: "intro",
    themeOverride: slideThemeOverride,
    colorMapOverride: slideColorMapOverride,
  }).append(
    box({
      id: "container",
      style: {
        position: "absolute",
        left: 24,
        top: 24,
        width: 592,
        height: 312,
        display: "flex",
        flexDirection: "column",
        gap: 8,
        padding: sharedPadding,
        background: schemeBg,
      },
    }).append(
      shape("Packed release", {
        id: "title",
        style: {
          height: 36,
          color: schemeText,
          background: accentShadeLight,
          fontFamily: majorFontSlot,
          border: { color: "#1d4ed8cc", width: 1 },
        },
      }),
      table({
        id: "summary",
        columns: ["1fr", "1fr"],
        style: { width: "100%" },
      }).append(
        row().append(
          cell("Metric", {
            style: {
              padding: sharedPadding,
              color: "#123c",
              fontFamily: minorFontRef,
            },
          }),
          cell("Value", {
            style: {
              padding: sharedPadding,
              color: schemeBg,
              background: accentShadeDark,
              fontFamily: minorFontRef,
            },
          }),
        ),
      ),
      chart({
        id: "trend",
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
            series: [{ name: "Revenue", values: [120, 150], color: "accent2" }],
          },
        ],
        legend: "bottom",
        style: { flex: 1, width: "100%", height: "100%" },
      }),
    ),
  ),
);
