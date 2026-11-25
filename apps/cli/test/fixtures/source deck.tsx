import { Deck, Shape, Slide } from "pptxdk/react";

import { sourceLabel } from "./source-label.ts";

export function SourceDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="source">
        <Shape
          id="source-shape"
          name={sourceLabel}
          x={8}
          y={8}
          width={80}
          height={32}
          text={sourceLabel}
        />
      </Slide>
    </Deck>
  );
}

export function OverflowDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="overflow">
        <Shape id="overflow-shape" x={100} y={8} width={40} height={32} />
      </Slide>
    </Deck>
  );
}

export default SourceDeck;
