import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Deck,
  defineTheme,
  font,
  Slide,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

const cjkFont = font("Theme Metrics Sans", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
], { substitute: true });

describe("公开主题入口", () => {
  it("主题字体参与内容度量，文字保留 major 字体与原生颜色槽引用", async () => {
    const result = await build(
      <Deck
        theme={defineTheme({
          name: "Theme Contract",
          colors: { ink: "accent1", accent1: "#336699" },
          fonts: { major: cjkFont },
        })}
        style={{ fontFamily: "major", fontSize: 20, color: "ink" }}
      >
        <Slide slideId="theme">
          <Text testId="measured" x={0} y={0} width={400} height="auto">
            中文 A é
          </Text>
        </Slide>
      </Deck>,
    );

    assert.equal(result.report.slides[0]?.objects[0]?.box?.height, 28.96);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const parts = unzipSync(result.pptx);
    const theme = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slide = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.match(theme, /<a:theme[^>]+name="Theme Contract"/u);
    assert.match(theme, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/u);
    assert.match(theme, /<a:majorFont><a:latin typeface="Theme Metrics Sans"/u);
    assert.match(slide, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.match(slide, /<a:schemeClr val="accent1"><\/a:schemeClr>/u);
  });
});
