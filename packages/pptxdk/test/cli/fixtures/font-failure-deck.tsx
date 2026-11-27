import { readFileSync } from "node:fs";

import { defineTheme } from "@officedk/pptxdk";
import { Deck, Slide, Text } from "@officedk/pptxdk/react";
import { font } from "@officedk/pptxdk/vanilla";

const jost400 = font(
  "Jost",
  [
    {
      weight: 400,
      data: readFileSync(
        new URL("../../assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    },
  ],
  { substitute: true },
);

const jostTheme = defineTheme({
  name: "Jost 400 Only",
  fonts: { major: jost400, minor: jost400 },
});

export function MissingFontWeightDeck() {
  return (
    <Deck width={320} height={180} theme={jostTheme}>
      <Slide id="missing-weight">
        <Text
          id="headline"
          style={{
            position: "absolute",
            left: 20,
            top: 20,
            width: 240,
            fontFamily: jost400,
            fontWeight: 600,
          }}
        >
          SemiBold heading
        </Text>
      </Slide>
    </Deck>
  );
}

export function MissingGlyphDeck() {
  return (
    <Deck width={320} height={180} theme={jostTheme}>
      <Slide id="missing-glyph">
        <Text
          id="glyph-text"
          style={{
            position: "absolute",
            left: 20,
            top: 20,
            width: 240,
            fontFamily: jost400,
          }}
        >
          Hello 汉字
        </Text>
      </Slide>
    </Deck>
  );
}

export default MissingFontWeightDeck;
