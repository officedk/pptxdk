import { readFileSync } from "node:fs";

import {
  deck,
  defineTheme,
  font,
  shape,
  slide,
} from "@officedk/pptxdk/vanilla";

const family = font("Jost*", [
  { weight: 400, data: readFileSync(new URL("./Jost.ttf", import.meta.url)) },
]);
export default deck({
  width: 640,
  height: 360,
  theme: defineTheme({
    name: "Consumer",
    fonts: { major: family, minor: family },
  }),
}).append(
  slide({ id: "intro" }).append(
    shape({
      id: "title",
      x: 24,
      y: 24,
      width: 360,
      height: 80,
      text: "Packed release",
    }),
  ),
);
