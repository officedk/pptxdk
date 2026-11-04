import {
  Group,
  MotionScope,
  Shape,
  Step,
  Text,
  fade,
  fillColor,
  preset,
  useMotionScope,
} from "@react-pptx/react-pptx";

import { regionBars, type MapRegionId } from "../assets/world-map.ts";
import { colors } from "../theme.ts";

export type RegionStory = {
  readonly region: MapRegionId;
  readonly title: string;
  readonly description: string;
  readonly color: string;
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
  const scale = (width * 0.7) / 1000;
  const legendX = x + width * 0.76;
  const legendY = y + index * 78;
  return (
    <Step n={index + 1}>
      <MotionScope scope={highlight}>
        <Group testId={`map-${story.region}`}>
          {regionBars[story.region].map(([bx, by, bw, bh], i) => (
            <Shape
              key={i}
              x={x + bx * scale}
              y={y + by * scale}
              width={bw * scale}
              height={bh * scale}
              geometry={preset("roundRect")}
              fill={colors.pale}
              emphasis={fillColor({
                from: colors.pale,
                to: story.color,
                duration: 650,
              })}
            />
          ))}
        </Group>
      </MotionScope>
      <MotionScope scope={caption} startAt={highlight.start}>
        <Group
          testId={`map-note-${story.region}`}
          enter={fade({ duration: 350 })}
        >
          <Shape
            x={legendX}
            y={legendY + 7}
            width={8}
            height={26}
            fill={story.color}
          />
          <Text
            x={legendX + 24}
            y={legendY}
            width={230}
            height={38}
            style={{ fontSize: 24 }}
          >
            {story.title}
          </Text>
          <Text
            x={legendX + 24}
            y={legendY + 39}
            width={230}
            height={32}
            style={{ fontSize: 16, color: colors.muted }}
          >
            {story.description}
          </Text>
        </Group>
      </MotionScope>
    </Step>
  );
}

/** 六地区的横条地图与同期说明；regions 的顺序就是六个点击步的顺序。 */
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
