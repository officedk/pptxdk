import {
  Chart,
  Group,
  preset,
  Shape,
  Slide,
  Text,
  transitions,
  type MorphScope,
} from "@officedk/pptxdk/react";

import { colors, motion } from "../theme.ts";

export const galleryTopics = [
  {
    id: "contrast",
    title: "Contrast",
    description: "Lorem ipsum dolor sit amet.",
  },
  {
    id: "process",
    title: "Process",
    description: "Consectetur adipiscing elit.",
  },
  {
    id: "hierarchy",
    title: "Hierarchy",
    description: "Sed do eiusmod tempor.",
  },
  { id: "space", title: "Space", description: "Ut enim ad minim veniam." },
] as const;

function GalleryPanel({
  index,
  focus,
  scope,
}: {
  index: number;
  focus: number;
  scope: MorphScope;
}) {
  const topic = galleryTopics[index]!;
  const active = index === focus;
  const panelWidth = active ? 640 : 140;
  const left =
    86 +
    galleryTopics
      .slice(0, index)
      .reduce(
        (x, _, itemIndex) => x + (itemIndex === focus ? 640 : 140) + 16,
        0,
      );
  const scale = active ? 1 : 0.18;
  const contentX = left + (active ? 28 : 16);
  const contentY = active ? 274 : 310;
  const shape = (
    key: string,
    x: number,
    y: number,
    width: number,
    height: number,
    fill: string,
    circle = false,
  ) => (
    <Shape
      key={key}
      x={contentX + x * scale}
      y={contentY + y * scale}
      width={width * scale}
      height={height * scale}
      fill={fill}
      {...(circle ? { geometry: preset("ellipse") } : {})}
    />
  );
  const label = (
    key: string,
    text: string,
    x: number,
    y: number,
    width: number,
  ) => (
    <Text
      key={key}
      x={contentX + x * scale}
      y={contentY + y * scale}
      width={width * scale}
      height={34 * scale}
      style={{
        fontSize: 22 * scale,
        fontWeight: 700,
        align: "center",
        color: colors.ink,
      }}
    >
      {text}
    </Text>
  );
  return (
    <Group
      id={`gallery-panel-${topic.id}`}
      morphKey={scope.handle(`panel-${topic.id}`)}
    >
      <Shape
        x={left}
        y={170}
        width={panelWidth}
        height={366}
        fill={active ? colors.surface : colors.darkSurface}
      />
      <Text
        x={left + (active ? 28 : 12)}
        y={active ? 194 : 482}
        width={active ? 570 : 116}
        height={active ? 42 : 28}
        style={{
          fontSize: active ? 30 : 17,
          lineHeight: active ? 34 : 20,
          fontWeight: 700,
          color: active ? colors.ink : colors.surface,
        }}
      >
        {topic.title}
      </Text>
      <Text
        x={left + 28}
        y={492}
        width={active ? 560 : 84}
        height={26}
        style={{ fontSize: 15, color: colors.muted }}
      >
        {active ? topic.description : ""}
      </Text>
      {index === 0 && (
        <Chart
          id="contrast-chart"
          x={contentX}
          y={contentY}
          width={584 * scale}
          height={196 * scale}
          legend={false}
          plotArea={{
            left: 8 * scale,
            top: 8 * scale,
            width: 568 * scale,
            height: 140 * scale,
          }}
          style={{
            fontSize: 20 * scale,
            fontFamily: "chart",
            color: colors.ink,
          }}
          axes={{
            category: {
              kind: "category",
              visible: active,
              line: { color: colors.soft, width: 1 },
              gridlines: false,
              style: {
                fontSize: 20 * scale,
                fontFamily: "chart",
                fontWeight: 700,
                color: colors.ink,
              },
            },
            value: {
              kind: "value",
              visible: false,
              min: 0,
              max: 100,
              gridlines: false,
            },
          }}
          groups={[
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [
                {
                  name: "Lorem ipsum",
                  values: [84, 56],
                  color: colors.accent,
                  points: { 1: { color: colors.ink } },
                },
              ],
              gapWidth: 80,
              dataLabels: active
                ? {
                    value: true,
                    position: "outsideEnd",
                    style: {
                      fontSize: 16,
                      fontFamily: "medium",
                      fontWeight: 500,
                      color: colors.ink,
                    },
                  }
                : false,
            },
          ]}
        />
      )}
      {index === 1 && (
        <>
          {shape("connector-a", 92, 75, 166, 5, colors.accentSoft)}
          {shape("connector-b", 342, 75, 166, 5, colors.accentSoft)}
          {[0, 1, 2].map((node) => (
            <Group key={node}>
              {shape(
                `node-${node}`,
                node * 250,
                30,
                100,
                100,
                node === 1 ? colors.current : colors.accent,
                true,
              )}
              {label(
                `node-label-${node}`,
                String(node + 1),
                node * 250,
                60,
                100,
              )}
            </Group>
          ))}
        </>
      )}
      {index === 2 && (
        <>
          {shape("stem", 294, 42, 5, 62, colors.accentSoft)}
          {shape("branch", 84, 100, 426, 5, colors.accentSoft)}
          {shape("root", 200, 0, 190, 50, colors.accent)}
          {[0, 1, 2].map((node) => (
            <Group key={node}>
              {shape(
                `leaf-stem-${node}`,
                82 + node * 210,
                100,
                5,
                38,
                colors.accentSoft,
              )}
              {shape(
                `leaf-${node}`,
                node * 210,
                136,
                170,
                62,
                node === 1 ? colors.accent : colors.soft,
              )}
            </Group>
          ))}
        </>
      )}
      {index === 3 && (
        <>
          {shape("space-line", 20, 180, 560, 2, colors.soft)}
          {shape("space-a", 32, 48, 132, 132, colors.accent, true)}
          {shape("space-b", 194, 76, 76, 76, colors.current, true)}
          {shape("space-c", 432, 20, 160, 160, colors.soft, true)}
        </>
      )}
    </Group>
  );
}

/** Four focus states share a continuous canvas; column width changes replace off-screen panels, and navigation supports direct access and return. */
export function GalleryPage({
  focus,
  scope,
}: {
  readonly focus: number;
  readonly scope: MorphScope;
}) {
  return (
    <Slide
      id={`gallery-${galleryTopics[focus]!.id}`}
      background={colors.ink}
      transition={transitions.morph({
        duration: motion.morph,
        fallback: transitions.fade({ duration: 400 }),
      })}
    >
      <Text
        x={80}
        y={58}
        width={680}
        height={58}
        style={{
          fontSize: 44,
          lineHeight: 48,
          fontWeight: 700,
          color: colors.surface,
        }}
      >
        Continuous canvas.
      </Text>
      <Text
        x={960}
        y={72}
        width={240}
        height={30}
        style={{ fontSize: 15, align: "right", color: colors.accentSoft }}
      >
        {11 + focus} / MORPH
      </Text>
      {galleryTopics.map((topic, index) => (
        <GalleryPanel
          key={topic.id}
          index={index}
          focus={focus}
          scope={scope}
        />
      ))}
      {galleryTopics.map((topic, index) => (
        <Shape
          key={topic.id}
          morphKey={scope.handle(`navigation-${topic.id}`)}
          link={{ id: `gallery-${topic.id}` }}
          id={`navigation-${topic.id}`}
          x={86 + index * 284}
          y={584}
          width={260}
          height={54}
          fill={index === focus ? colors.darkSurface : colors.ink}
          geometry={preset("roundRect")}
          text={topic.title}
          textStyle={{
            fontSize: 21,
            fontFamily: index === focus ? "major" : "medium",
            fontWeight: index === focus ? 700 : 500,
            color: index === focus ? colors.surface : colors.accentSoft,
            align: "center",
            verticalAlign: "middle",
          }}
        />
      ))}
      <Shape
        x={86 + focus * 284}
        y={652}
        width={260}
        height={5}
        fill={colors.current}
        morphKey={scope.handle("navigation-selection")}
      />
    </Slide>
  );
}
