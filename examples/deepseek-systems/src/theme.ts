import { readFileSync } from "node:fs";

import { defineTheme, font } from "@officedk/pptxdk/vanilla";

function fontBytes(name: string): Uint8Array {
  return new Uint8Array(
    readFileSync(new URL(`./assets/fonts/${name}`, import.meta.url)),
  );
}

export const fonts = {
  body: font("Noto Sans SC", [
    { weight: 400, data: fontBytes("NotoSansSC-Regular.ttf") },
  ]),
  metric: font(
    "Jost",
    [400, 700].map((weight) => ({
      weight,
      data: fontBytes(`Jost-${weight}.ttf`),
    })),
  ),
} as const;

/**
 * Silicon Telemetry palette: cool quartz-silicon canvas paired with wafer-white cards
 * and hardware-semantic accents (Ultramarine for compute/model, Teal for RDMA/storage,
 * Copper for breakthrough metrics/hotspots, Pewter for external baselines).
 */
export const colors = {
  paper: "#F2F4F7",
  surface: "#FFFFFF",
  slot: "#E6EAF0",
  obsidian: "#0B101B",
  obsidianCard: "#151E30",
  ink: "#111622",
  steel: "#23304A",
  muted: "#5B677D",
  line: "#CFD6E2",
  darkLine: "#26344F",
  darkMuted: "#94A3B8",
  darkSubtle: "#64748B",
  compute: "#1849C6",
  computeBright: "#93B4FF",
  computeSoft: "#DCE6FD",
  comm: "#0D8A7B",
  commBright: "#5EEAD4",
  commText: "#08665A",
  commSoft: "#D5F2EE",
  copper: "#D95D24",
  copperBright: "#FDBA74",
  copperText: "#B34312",
  copperSoft: "#FCE5D8",
  baseline: "#8492A6",
  baselineSoft: "#E9EDF2",
} as const;

export const motionTiming = {
  fast: 360,
  reveal: 480,
  morph: 900,
} as const;

export const theme = defineTheme({
  name: "晶圆青金 · 超算系统图谱",
  colors: {
    dk1: colors.ink,
    lt1: colors.surface,
    dk2: colors.steel,
    lt2: colors.paper,
    accent1: colors.compute,
    accent2: colors.comm,
    accent3: colors.copper,
    accent4: colors.muted,
    accent5: colors.line,
    accent6: colors.baseline,
    hlink: colors.compute,
    folHlink: colors.comm,
    ...colors,
  },
  fonts: {
    major: fonts.body,
    minor: fonts.body,
    body: fonts.body,
    metric: fonts.metric,
  },
});

export const transparent = { color: colors.paper, alpha: 0 } as const;
