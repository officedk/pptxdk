import { Deck, Shape, Slide } from "@officedk/pptxdk/react";
import { deck, shape, slide } from "@officedk/pptxdk/vanilla";
import { useEffect, useMemo, useState } from "react";

import { sourceLabel } from "./source-label.ts";

export function PreviewDeck() {
  return (
    <Deck width={120} height={96}>
      <Slide id="preview">
        <Shape
          id="preview-shape"
          name={sourceLabel}
          x={8}
          y={8}
          width={80}
          height={32}
          fill="#2563EB"
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
        x: 10,
        y: 12,
        width: 72,
        height: 36,
        fill: "#1D4ED8",
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
        x={12}
        y={14}
        width={64}
        height={28}
        fill="#059669"
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
          x={16}
          y={16}
          width={70}
          height={30}
          fill="#D97706"
        />
      </Slide>
    </Deck>
  );
}

export default PreviewDeck;
