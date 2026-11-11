import {
  Group,
  MotionScope,
  Shape,
  Step,
  Text,
  fillColor,
  preset,
  useMotionScope,
  wipe,
} from "@react-pptx/react-pptx";

import { regionBars, type MapRegionId } from "../assets/world-map.ts";
import { colors, motion } from "../theme.ts";

export type RegionStory = {
  readonly region: MapRegionId;
  readonly title: string;
  readonly description: string;
};

function Region({
  story,
  index,
  x,
  y,
  width,
}: {
  story: RegionStory;
  index: number;
  x: number;
  y: number;
  width: number;
}) {
  const highlight = useMotionScope();
  const caption = useMotionScope();
  const scale = (width * 0.72) / 1000;
  const legendX = x + width * 0.77;
  const legendY = y + 24 + index * 76;
  return (
    <Step n={index + 1}>
      <MotionScope scope={highlight}>
        <Group id={`map-${story.region}`}>
          {regionBars[story.region].map(([bx, by, bw, bh], barIndex) => (
            <Shape
              key={barIndex}
              x={x + bx * scale}
              y={y + by * scale}
              width={bw * scale}
              height={bh * scale}
              geometry={preset("roundRect")}
              fill={colors.soft}
              emphasis={fillColor({
                from: colors.soft,
                to: colors.accent,
                duration: motion.settle,
              })}
            />
          ))}
        </Group>
      </MotionScope>
      <MotionScope scope={caption} startAt={highlight.start}>
        <Group
          id={`map-note-${story.region}`}
          enter={wipe({ from: "left", duration: motion.reveal })}
        >
          <Shape
            x={legendX}
            y={legendY + 5}
            width={30}
            height={4}
            fill={colors.accent}
          />
          <Text
            x={legendX + 46}
            y={legendY - 4}
            width={230}
            height={36}
            style={{
              fontSize: 24,
              fontFamily: "medium",
              fontWeight: 500,
              color: colors.ink,
            }}
          >
            {story.title}
          </Text>
          <Text
            x={legendX + 46}
            y={legendY + 35}
            width={230}
            height={42}
            style={{ fontSize: 15, lineHeight: 20, color: colors.muted }}
          >
            {story.description}
          </Text>
        </Group>
      </MotionScope>
    </Step>
  );
}

/** Six regions share coordinated accent colors; the order of regions matches the order of the six click steps. */
export function WorldMap({
  x,
  y,
  width,
  regions,
}: {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly regions: readonly RegionStory[];
}) {
  return (
    <>
      {regions.map((story, index) => (
        <Region
          key={story.region}
          story={story}
          index={index}
          x={x}
          y={y}
          width={width}
        />
      ))}
    </>
  );
}
