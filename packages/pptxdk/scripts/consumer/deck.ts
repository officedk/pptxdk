import { readFileSync } from "node:fs";

import {
  box,
  cell,
  chart,
  deck,
  defineTheme,
  font,
  row,
  shape,
  slide,
  table,
} from "@officedk/pptxdk/vanilla";

const family = font("Jost*", [
  { data: readFileSync(new URL("./Jost.ttf", import.meta.url)) },
]);
const sharedPadding = [8, 12] as const;

export default deck({
  width: 640,
  height: 360,
  theme: defineTheme({
    name: "Consumer",
    fonts: { major: family, minor: family },
  }),
}).append(
  slide({ id: "intro" }).append(
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
      },
    }).append(
      shape("Packed release", {
        id: "title",
        style: { height: 36 },
      }),
      table({
        id: "summary",
        columns: ["1fr", "1fr"],
        style: { width: "100%" },
      }).append(
        row().append(
          cell("Metric", { style: { padding: sharedPadding } }),
          cell("Value", { style: { padding: sharedPadding } }),
        ),
      ),
      chart({
        id: "trend",
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
            series: [{ name: "Revenue", values: [120, 150] }],
          },
        ],
        legend: "bottom",
        style: { flex: 1, width: "100%", height: "100%" },
      }),
    ),
  ),
);
