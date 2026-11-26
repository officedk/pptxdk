import { Deck, Shape, Slide } from "@officedk/pptxdk/react";

export default function BrokenDeck() {
  return (
    <Deck>
      <Slide id="broken">
        <Shape
          id="broken-shape"
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            width: -20,
            height: 20,
          }}
        />
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
          style={{
            position: "absolute",
            left: 20,
            top: 20,
            width: 120,
            height: 60,
            background: "#2563EB",
          }}
        />
      </Slide>
      <Slide id="broken-unselected">
        <Shape
          id="broken-unselected-shape"
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            width: -20,
            height: 20,
          }}
        />
      </Slide>
    </Deck>
  );
}
