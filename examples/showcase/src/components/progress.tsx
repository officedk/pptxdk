import { Group, Shape, Text, blockArc, preset } from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

/** 数值、条形和环形使用同一分子与分母，避免图形和文字各自维护比例。 */
export function Progress({
  completed,
  total,
}: {
  readonly completed: number;
  readonly total: number;
}) {
  const fraction = completed / total;
  return (
    <Group testId="progress-example">
      <Text
        x={64}
        y={250}
        width={300}
        height={110}
        style={{ fontSize: 72, color: colors.teal }}
      >
        {Math.round(fraction * 100)}%
      </Text>
      <Text x={64} y={380} width={300} height={60} style={{ fontSize: 26 }}>
        {completed} / {total} 个步骤
      </Text>
      <Shape
        x={410}
        y={318}
        width={400}
        height={30}
        fill={colors.pale}
        geometry={preset("roundRect")}
      />
      <Shape
        x={410}
        y={318}
        width={400 * fraction}
        height={30}
        fill={colors.teal}
        geometry={preset("roundRect")}
      />
      <Text
        x={410}
        y={380}
        width={400}
        height={60}
        style={{ fontSize: 23, color: colors.muted }}
      >
        把完成部分放回整体中理解
      </Text>
      <Shape
        x={920}
        y={250}
        width={230}
        height={230}
        fill={colors.pale}
        geometry={preset("ellipse")}
      />
      <Shape
        x={951}
        y={281}
        width={168}
        height={168}
        fill={colors.paper}
        geometry={preset("ellipse")}
      />
      <Shape
        x={920}
        y={250}
        width={230}
        height={230}
        fill={colors.teal}
        geometry={blockArc({
          startAngle: 270,
          endAngle: 270 + fraction * 360,
          thickness: 0.135,
        })}
      />
    </Group>
  );
}
