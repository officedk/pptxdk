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
      <Shape x={64} y={48} width={42} height={5} fill={colors.accent} />
      <Text
        x={64}
        y={76}
        width={760}
        height={68}
        style={{ fontSize: 50, lineHeight: 54, fontWeight: 700 }}
      >
        Around the world.
      </Text>
      <Text
        x={950}
        y={88}
        width={250}
        height={28}
        style={{ fontSize: 15, align: "right", color: colors.muted }}
      >
        10 / SIX REGIONS
      </Text>
      <WorldMap x={64} y={146} width={1136} regions={regions} />
    </Slide>
  );
}
