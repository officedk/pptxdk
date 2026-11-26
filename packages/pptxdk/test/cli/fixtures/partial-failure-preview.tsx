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
          style={{
            position: "absolute",
            left: 100,
            top: 8,
            width: 40,
            height: 32,
            background: "#2563EB",
          }}
        />
      </Slide>
      <Slide id="bad-image">
        <Image
          id="corrupt-img"
          src={corruptPixelPng}
          style={{
            position: "absolute",
            left: 8,
            top: 8,
            width: 32,
            height: 32,
          }}
        />
      </Slide>
      <Slide id="good-sibling">
        <Shape
          id="ok-shape"
          style={{
            position: "absolute",
            left: 8,
            top: 8,
            width: 40,
            height: 24,
            background: "#059669",
          }}
        />
      </Slide>
    </Deck>
  );
}

export default PartialFailurePreviewDeck;
