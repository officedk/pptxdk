import { Deck, Shape, Slide } from "@officedk/pptxdk/react";

import { sourceLabel } from "./source-label.ts";

export function SourceDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="source">
        <Shape
          id="source-shape"
          name={sourceLabel}
          style={{
            position: "absolute",
            left: 8,
            top: 8,
            width: 80,
            height: 32,
          }}
        >
          {sourceLabel}
        </Shape>
      </Slide>
    </Deck>
  );
}

export function OverflowDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="overflow">
        <Shape
          id="overflow-shape"
          style={{
            position: "absolute",
            left: 100,
            top: 8,
            width: 40,
            height: 32,
          }}
        />
      </Slide>
    </Deck>
  );
}

export default SourceDeck;
