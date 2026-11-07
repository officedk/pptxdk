import {
  Group,
  Shape,
  Step,
  Text,
  floatUp,
  preset,
} from "@react-pptx/react-pptx";

import { colors, motion } from "../theme.ts";

/** Numbered nodes and connectors are revealed in content order; each node is preserved as a native Group. */
export function ProcessPath({
  nodes,
  x,
  y,
  width,
  firstStep = 1,
}: {
  readonly nodes: readonly { readonly title: string; readonly body: string }[];
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly firstStep?: number;
}) {
  const columnWidth = width / nodes.length;
  return (
    <>
      {nodes.map((node, index) => {
        const center = x + index * columnWidth + columnWidth / 2;
        return (
          <Step key={`${node.title}-${index}`} n={firstStep + index}>
            <Group enter={floatUp({ duration: motion.reveal })}>
              {index > 0 && (
                <Shape
                  x={center - columnWidth + 48}
                  y={y + 47}
                  width={columnWidth - 96}
                  height={3}
                  fill={colors.accentSoft}
                />
              )}
              <Shape
                x={center - 48}
                y={y}
                width={96}
                height={96}
                geometry={preset("ellipse")}
                fill={colors.accent}
                text={String(index + 1).padStart(2, "0")}
                textStyle={{
                  fontSize: 26,
                  fontWeight: 700,
                  color: colors.surface,
                  align: "center",
                  verticalAlign: "middle",
                }}
              />
              <Text
                x={x + index * columnWidth + 18}
                y={y + 132}
                width={columnWidth - 36}
                height={42}
                style={{
                  fontSize: 27,
                  fontWeight: 700,
                  align: "center",
                  color: colors.ink,
                }}
              >
                {node.title}
              </Text>
              <Text
                x={x + index * columnWidth + 24}
                y={y + 190}
                width={columnWidth - 48}
                height={88}
                style={{
                  fontSize: 18,
                  lineHeight: 26,
                  align: "center",
                  color: colors.muted,
                }}
              >
                {node.body}
              </Text>
            </Group>
          </Step>
        );
      })}
    </>
  );
}
