import { Deck, Shape, Slide } from "@react-pptx/react-pptx";

import { sourceLabel } from "./source-label.ts";

export function SourceDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide slideId="source">
        <Shape
          testId="source-shape"
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

export default SourceDeck;
