import { Deck, Shape, Slide } from "@react-pptx/react-pptx";

export default function BrokenDeck() {
  return (
    <Deck>
      <Slide slideId="broken">
        <Shape testId="broken-shape" x={0} y={0} width={0} height={20} />
      </Slide>
    </Deck>
  );
}
