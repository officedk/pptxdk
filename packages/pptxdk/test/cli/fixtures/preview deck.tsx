import { readFileSync } from "node:fs";

import { Deck, Image, Shape, Slide, Text } from "@officedk/pptxdk/react";
import { deck, font, shape, slide } from "@officedk/pptxdk/vanilla";
import { useEffect, useMemo, useState } from "react";

import { sourceLabel } from "./source-label.ts";

export function PreviewDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="preview">
        <Shape
          id="preview-shape"
          name={sourceLabel}
          style={{
            position: "absolute",
            left: 8,
            top: 8,
            width: 80,
            height: 32,
            background: "#2563EB",
          }}
        />
      </Slide>
    </Deck>
  );
}

export const AsyncVanillaDeck = (async () => {
  await Promise.resolve();
  return deck({ width: 120, height: 96 }).append(
    slide({ id: "async-vanilla" }).append(
      shape({
        id: "async-shape",
        name: "Async Vanilla Shape",
        style: {
          position: "absolute",
          left: 10,
          top: 12,
          width: 72,
          height: 36,
          background: "#1D4ED8",
        },
      }),
    ),
  );
})();

export const JsxElementDeck = (
  <Deck width={120} height={96}>
    <Slide id="jsx-element">
      <Shape
        id="jsx-shape"
        name="JSX Element Shape"
        style={{
          position: "absolute",
          left: 12,
          top: 14,
          width: 64,
          height: 28,
          background: "#059669",
        }}
      />
    </Slide>
  </Deck>
);

export function HookComponent() {
  const [label, setLabel] = useState("initial");
  const prefix = useMemo(() => "Hook:", []);
  useEffect(() => {
    setLabel("committed");
  }, []);
  return (
    <Deck width={120} height={96}>
      <Slide id="hook-component">
        <Shape
          id="hook-shape"
          name={`${prefix} ${label}`}
          style={{
            position: "absolute",
            left: 16,
            top: 16,
            width: 70,
            height: 30,
            background: "#D97706",
          }}
        />
      </Slide>
    </Deck>
  );
}

export function ExternalPlaceholderDeck() {
  return (
    <Deck width={240} height={140}>
      <Slide id="external-slide">
        <Image
          id="hero-external"
          src={{ externalUrl: "https://cdn.example.com/cli-hero.png" }}
          style={{
            position: "absolute",
            left: 12,
            top: 12,
            width: 120,
            height: 80,
          }}
        />
      </Slide>
    </Deck>
  );
}

const unbundledSans = font("Unbundled Sans");
const substituteSans = font(
  "Unbundled Sans",
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

export function UnembeddedFontPreviewDeck() {
  return (
    <Deck width={240} height={140}>
      <Slide id="font-missing-page">
        <Text
          id="unbundled-label"
          style={{
            position: "absolute",
            left: 12,
            top: 12,
            width: 180,
            height: 40,
            fontFamily: unbundledSans,
          }}
        >
          Needs font
        </Text>
      </Slide>
      <Slide id="blank-sibling">
        <Shape
          id="ok-rect"
          style={{
            position: "absolute",
            left: 12,
            top: 12,
            width: 60,
            height: 32,
            background: "#0D9488",
          }}
        />
      </Slide>
    </Deck>
  );
}

export function SubstituteFontPreviewDeck() {
  return (
    <Deck width={240} height={140}>
      <Slide id="font-missing-page">
        <Text
          id="unbundled-label"
          style={{
            position: "absolute",
            left: 12,
            top: 12,
            width: 180,
            height: 40,
            fontFamily: substituteSans,
          }}
        >
          Needs font
        </Text>
      </Slide>
      <Slide id="blank-sibling">
        <Shape
          id="ok-rect"
          style={{
            position: "absolute",
            left: 12,
            top: 12,
            width: 60,
            height: 32,
            background: "#0D9488",
          }}
        />
      </Slide>
    </Deck>
  );
}

export default PreviewDeck;
