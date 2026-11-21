import { Deck, Shape, Slide } from "pptxdk/react";

export default function BrokenDeck() {
  return (
    <Deck>
      <Slide id="broken">
        <Shape id="broken-shape" x={0} y={0} width={0} height={20} />
      </Slide>
    </Deck>
  );
}

export function UnselectedPageBrokenDeck() {
  return (
    <Deck>
      <Slide id="valid-selected">
        <Shape
          id="ok-shape"
          x={20}
          y={20}
          width={120}
          height={60}
          fill="#2563EB"
        />
      </Slide>
      <Slide id="broken-unselected">
        <Shape id="broken-unselected-shape" x={0} y={0} width={0} height={20} />
      </Slide>
    </Deck>
  );
}
