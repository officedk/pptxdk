import { Deck, Image, Shape, Slide } from "@officedk/pptxdk/react";

// Passes formal resolve() header validation, then fails paint() on this slide only with a PNG pixel decode error.
const corruptPixelPng = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49,
  0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02,
  0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x01, 0x49, 0x44,
  0x41, 0x54, 0xff, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
  0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

export function PartialFailurePreviewDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="warn-ok">
        <Shape
          id="overflow-shape"
          x={100}
          y={8}
          width={40}
          height={32}
          fill="#2563EB"
        />
      </Slide>
      <Slide id="bad-image">
        <Image
          id="corrupt-img"
          src={corruptPixelPng}
          x={8}
          y={8}
          width={32}
          height={32}
        />
      </Slide>
      <Slide id="good-sibling">
        <Shape
          id="ok-shape"
          x={8}
          y={8}
          width={40}
          height={24}
          fill="#059669"
        />
      </Slide>
    </Deck>
  );
}

export default PartialFailurePreviewDeck;
