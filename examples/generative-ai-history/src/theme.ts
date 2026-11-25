import { readFileSync } from "node:fs";

import { defineTheme, font } from "@officedk/pptxdk/vanilla";

function fontBytes(name: string) {
  return new Uint8Array(
    readFileSync(new URL(`./assets/fonts/${name}`, import.meta.url)),
  );
}

export const fonts = {
  body: font("Noto Sans SC", [
    { weight: 400, data: fontBytes("NotoSansSC-Regular.ttf") },
  ]),
  display: font("ZCOOL XiaoWei", [
    { weight: 400, data: fontBytes("ZCOOLXiaoWei_400Regular.ttf") },
  ]),
};

export const colors = {
  paper: "#EAF1F2",
  paperBright: "#F7FAF9",
  ink: "#153B4F",
  indigo: "#0E2C3F",
  cobalt: "#276C91",
  cyan: "#5FA7B2",
  sun: "#E9B949",
  muted: "#6F8793",
  line: "#B8CBD1",
  mist: "#D5E3E5",
} as const;

export const theme = defineTheme({
  name: "蓝晒图谱",
  colors: {
    dk1: colors.indigo,
    lt1: colors.paperBright,
    dk2: colors.ink,
    lt2: colors.paper,
    accent1: colors.cobalt,
    accent2: colors.cyan,
    accent3: colors.sun,
    accent4: colors.muted,
    accent5: colors.line,
    accent6: colors.mist,
    ...colors,
  },
  fonts: {
    major: fonts.display,
    minor: fonts.body,
    display: fonts.display,
    body: fonts.body,
  },
});

export const transparent = { color: colors.paper, alpha: 0 } as const;
