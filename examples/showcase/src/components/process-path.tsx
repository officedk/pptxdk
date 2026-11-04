import { Group, Shape, Step, Text, fade, preset } from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

/** 连线属于整条路径，节点保留原生组合，节点顺序即点击顺序。 */
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
  const nodeWidth = (width - (nodes.length - 1) * 40) / nodes.length;
  return (
    <>
      <Shape
        x={x + 30}
        y={y + 50}
        width={width - 60}
        height={3}
        fill={colors.pale}
      />
      {nodes.map((node, index) => (
        <Step key={node.title} n={firstStep + index}>
          <Group enter={fade({ duration: 400 })}>
            <Shape
              x={x + index * (nodeWidth + 40)}
              y={y}
              width={nodeWidth}
              height={190}
              fill={index === 1 ? colors.teal : colors.ink}
              geometry={preset("roundRect")}
            />
            <Text
              x={x + index * (nodeWidth + 40) + 24}
              y={y + 24}
              width={nodeWidth - 48}
              height={48}
              style={{ fontSize: 28, color: colors.paper }}
            >
              {node.title}
            </Text>
            <Text
              x={x + index * (nodeWidth + 40) + 24}
              y={y + 90}
              width={nodeWidth - 48}
              height={80}
              style={{ fontSize: 20, color: colors.paper }}
            >
              {node.body}
            </Text>
          </Group>
        </Step>
      ))}
    </>
  );
}
