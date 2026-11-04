import { readFileSync } from "node:fs";

import { defineTheme, font } from "@react-pptx/react-pptx";

// 字体来自 google/fonts 的 ofl/notosanssc/NotoSansSC[wght].ttf，
// 用 fontTools.varLib.instancer 固定 wght=400 并更新名称表；许可随字体分发。
// 字体随包提供；只使用实际带有度量资源的常规字重，以字号和颜色建立层级。
const bodyFont = font("Noto Sans SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.ttf", import.meta.url),
      ),
    ),
  },
]);

export const colors = {
  paper: "#F7F5EF",
  ink: "#16324F",
  teal: "#138A83",
  coral: "#E98068",
  muted: "#647585",
  pale: "#E4ECE8",
} as const;

export const theme = defineTheme({
  name: "演示表达",
  colors,
  fonts: { major: bodyFont, minor: bodyFont },
});
