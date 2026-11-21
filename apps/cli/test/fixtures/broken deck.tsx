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
