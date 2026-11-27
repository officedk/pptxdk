import { readFileSync } from "node:fs";

import {
  createMorphScope,
  defaultColorMap,
  defineSlideLayout,
  defineTheme,
  font,
  lighter,
  linearGradient,
  type LayoutPlaceholderSpec,
  type LayoutStyle,
} from "@officedk/pptxdk/vanilla";

function fontBytes(name: string): Uint8Array {
  return new Uint8Array(
    readFileSync(new URL(`./assets/fonts/${name}`, import.meta.url)),
  );
}

function imageBytes(name: string): Uint8Array {
  return new Uint8Array(
    readFileSync(new URL(`./assets/${name}`, import.meta.url)),
  );
}

// Static TTF faces omit weight/style so font metadata inference reads OS/2 and post tables.
export const jostFont = font("Jost", [
  { data: fontBytes("Jost-400.ttf") },
  { data: fontBytes("Jost-400-Italic.ttf") },
  { data: fontBytes("Jost-700.ttf") },
  { data: fontBytes("Jost-700-BoldItalic.ttf") },
]);

export const jostMediumFont = font("Jost Medium", [
  { data: fontBytes("Jost-500.ttf") },
]);

// Name ID 1 in DarkerGrotesque-SemiBold.ttf is "Darker Grotesque SemiBold".
export const displayFont = font("Darker Grotesque SemiBold", [
  { data: fontBytes("DarkerGrotesque-SemiBold.ttf") },
]);

export const cassegrainApertureImage = imageBytes("cassegrain-aperture.png");
export const wavefrontStrehlImage = imageBytes("wavefront-strehl.png");
export const deepSpaceBackgroundImage = imageBytes("deep-space-bg.png");

export const PALETTE = {
  obsidian: "#0B101B",
  obsidianCard: "#131C2E",
  obsidianRaised: "#1B2740",
  obsidianBorder: "#2A3B5E",
  paper: "#F6F5F0",
  paperAlt: "#ECEAE2",
  surface: "#FFFFFF",
  borderLight: "#D8D4C8",
  borderStrong: "#BDB7A6",
  ink: "#0F1726",
  inkBody: "#2C3649",
  inkMuted: "#5A6578",
  ivoryText: "#F4F2EC",
  slateLight: "#98A9C7",
  vermilion: "#D94E28",
  vermilionSoft: "#FBEAE4",
  cobalt: "#1D63DC",
  cobaltSoft: "#E6EFFD",
  teal: "#0E8A74",
  tealSoft: "#E2F4F0",
  amber: "#C88A1E",
  amberSoft: "#FBF3E2",
  violet: "#6B4CE6",
} as const;

export const aetherTheme = defineTheme({
  name: "Aether-Link Monograph",
  colors: {
    dk1: PALETTE.ink,
    lt1: PALETTE.paper,
    dk2: PALETTE.obsidianCard,
    lt2: PALETTE.paperAlt,
    accent1: PALETTE.vermilion,
    accent2: PALETTE.cobalt,
    accent3: PALETTE.teal,
    accent4: PALETTE.amber,
    accent5: PALETTE.violet,
    accent6: PALETTE.inkMuted,
    hlink: PALETTE.cobalt,
    folHlink: "#1447A6",
    accent2Muted: lighter("accent2", 0.85),
    tileAccent: PALETTE.vermilion,
  },
  fonts: {
    major: jostFont,
    minor: jostFont,
  },
  formatScheme: {
    name: "Aether Optical Format Scheme",
    bgFillStyles: [
      { kind: "solid", color: { color: "phClr" } },
      {
        kind: "gradient",
        linear: { angle: 145, scaled: true },
        stops: [
          { position: 0, color: { color: "phClr", shade: 0.72 } },
          { position: 0.65, color: { color: "phClr" } },
          { position: 1, color: { color: "phClr", shade: 0.58 } },
        ],
      },
      {
        kind: "solid",
        color: { color: "phClr", shade: 0.8 },
      },
    ],
  },
});

// Shared Morph continuity scope connecting Slide 2 and Slide 3.
export const telemetryMorph = createMorphScope();

const editorialPlaceholders: {
  readonly kicker: LayoutPlaceholderSpec;
  readonly title: LayoutPlaceholderSpec;
} = {
  kicker: {
    type: "subTitle",
    idx: 1,
    style: {
      position: "absolute",
      left: 56,
      top: 26,
      width: 820,
      height: 20,
      fontFamily: jostFont,
      fontSize: 11,
      fontWeight: 700,
      letterSpacing: 1.8,
      color: PALETTE.vermilion,
    },
  },
  title: {
    type: "title",
    style: {
      position: "absolute",
      left: 56,
      top: 46,
      width: 1168,
      height: 52,
      fontFamily: displayFont,
      fontSize: 38,
      fontWeight: 600,
      lineHeight: "42px",
      color: PALETTE.ink,
    },
  },
};

export const heroDarkLayout = defineSlideLayout({
  id: "hero-dark",
  name: "Hero Obsidian",
  type: "title",
  showMasterShapes: false,
  background: linearGradient(
    135,
    { color: "#080C15", position: 0 },
    { color: "#101A2E", position: 0.58 },
    { color: "#18243E", position: 1 },
  ),
  headerFooter: {
    dateTime: false,
    footer: false,
    slideNumber: false,
  },
});

export const editorialLightLayout = defineSlideLayout({
  id: "editorial-light",
  name: "Editorial Light",
  type: "titleOnly",
  showMasterShapes: true,
  background: PALETTE.paper,
  headerFooter: {
    dateTime: true,
    footer: true,
    slideNumber: true,
  },
  placeholders: editorialPlaceholders,
});

export const editorialChartLayout = defineSlideLayout({
  id: "editorial-chart",
  name: "Editorial Chart",
  type: "obj",
  showMasterShapes: true,
  background: PALETTE.paper,
  headerFooter: {
    dateTime: true,
    footer: true,
    slideNumber: true,
  },
  placeholders: {
    ...editorialPlaceholders,
    mainChart: {
      type: "chart",
      idx: 12,
      style: {
        position: "absolute",
        left: 74,
        top: 134,
        width: 750,
        height: 504,
      },
    },
  },
});

export const editorialVisualLayout = defineSlideLayout({
  id: "editorial-visual",
  name: "Editorial Visual",
  type: "picTx",
  showMasterShapes: true,
  background: PALETTE.paper,
  headerFooter: {
    dateTime: true,
    footer: true,
    slideNumber: true,
  },
  placeholders: {
    ...editorialPlaceholders,
    wavefrontMap: {
      type: "pic",
      idx: 10,
      style: {
        position: "absolute",
        left: 528,
        top: 116,
        width: 696,
        height: 416,
      },
    },
  },
});

export const editorialTableLayout = defineSlideLayout({
  id: "editorial-table",
  name: "Editorial Table",
  type: "tbl",
  showMasterShapes: true,
  background: PALETTE.paper,
  headerFooter: {
    dateTime: true,
    footer: true,
    slideNumber: true,
  },
  placeholders: {
    ...editorialPlaceholders,
    matrix: {
      type: "tbl",
      idx: 11,
      style: {
        position: "absolute",
        left: 56,
        top: 116,
        width: 1168,
        height: 472,
      },
    },
  },
});

export const anomalyBriefLayout = defineSlideLayout({
  id: "anomaly-brief",
  name: "Anomaly Brief",
  type: "titleOnly",
  showMasterShapes: true,
  background: PALETTE.paper,
  colorMapOverride: {
    ...defaultColorMap,
    accent6: "accent2",
  },
  themeOverride: {
    colors: {
      accent1: "#C9401C",
      accent3: "#0A7C68",
    },
  },
  headerFooter: {
    dateTime: true,
    footer: true,
    slideNumber: true,
  },
  placeholders: editorialPlaceholders,
});

export const closingDarkLayout = defineSlideLayout({
  id: "closing-dark",
  name: "Closing Obsidian",
  type: "cust",
  showMasterShapes: false,
  background: linearGradient(
    145,
    { color: "#090E19", position: 0 },
    { color: "#121D33", position: 0.65 },
    { color: "#0C1322", position: 1 },
  ),
  headerFooter: {
    dateTime: false,
    footer: false,
    slideNumber: false,
  },
  placeholders: {
    boardMemo: {
      type: "obj",
      idx: 9,
      style: {
        position: "absolute",
        left: 726,
        top: 604,
        width: 280,
        height: 46,
      },
    },
  },
});

export const EDITORIAL_BODY_STYLE: Omit<LayoutStyle, "zIndex"> = {
  display: "flex",
  flexDirection: "column",
  padding: [116, 56, 48, 56],
  gap: 16,
};
