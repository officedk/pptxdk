import { Shape, Slide, Text } from "@officedk/pptxdk/react";

import { WorldMap, type RegionStory } from "../components/world-map.tsx";
import { colors } from "../theme.ts";

const regions: readonly RegionStory[] = [
  {
    region: "northAmerica",
    title: "North America",
    description: "Lorem ipsum dolor sit amet.",
  },
  {
    region: "southAmerica",
    title: "South America",
    description: "Consectetur adipiscing elit.",
  },
  { region: "europe", title: "Europe", description: "Sed do eiusmod tempor." },
  {
    region: "africa",
    title: "Africa",
    description: "Incididunt ut labore et dolore.",
  },
  { region: "asia", title: "Asia", description: "Ut enim ad minim veniam." },
  {
    region: "oceania",
    title: "Oceania",
    description: "Quis nostrud exercitation.",
  },
];

export function WorldPage() {
  return (
    <Slide id="world" background={colors.paper}>
      <Shape
        style={{
          position: "absolute",
          left: 64,
          top: 48,
          width: 42,
          height: 5,
          background: colors.accent,
        }}
      />
      <Text
        style={{
          position: "absolute",
          left: 64,
          top: 76,
          width: 760,
          height: 68,
          fontSize: 50,
          lineHeight: 54,
          fontWeight: 700,
        }}
      >
        Around the world.
      </Text>
      <Text
        style={{
          position: "absolute",
          left: 950,
          top: 88,
          width: 250,
          height: 28,
          fontSize: 15,
          align: "right",
          color: colors.muted,
        }}
      >
        10 / SIX REGIONS
      </Text>
      <WorldMap x={64} y={146} width={1136} regions={regions} />
    </Slide>
  );
}
