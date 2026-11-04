import {
  Box,
  Shape,
  Slide,
  Text,
  Step,
  fade,
  preset,
} from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

/** 开场页用布局维护文字层级，局部几何构图使用绝对定位。 */
export function OpeningPage() {
  return (
    <Slide slideId="opening" background={colors.paper}>
      <Box
        style={{
          position: "absolute",
          left: 80,
          top: 90,
          width: 700,
          display: "flex",
          flexDirection: "column",
          gap: 26,
        }}
      >
        <Text
          width={700}
          height="auto"
          style={{ fontSize: 20, color: colors.teal }}
        >
          演示表达示例集
        </Text>
        <Text width={700} height="auto" style={{ fontSize: 64 }}>
          让信息{`\n`}逐步清晰
        </Text>
        <Step n={1}>
          <Text
            width={650}
            height="auto"
            style={{ fontSize: 26, color: colors.muted }}
            enter={fade()}
          >
            建立层级 · 组织关系 · 留出讲述的节奏
          </Text>
        </Step>
      </Box>
      <Shape
        x={935}
        y={180}
        width={180}
        height={180}
        geometry={preset("ellipse")}
        fill={colors.teal}
      />
      <Shape x={875} y={350} width={180} height={180} fill={colors.ink} />
      <Shape
        x={1060}
        y={365}
        width={90}
        height={90}
        geometry={preset("ellipse")}
        fill={colors.coral}
      />
      <Text
        x={80}
        y={640}
        width={1100}
        height={34}
        style={{ fontSize: 16, color: colors.muted }}
      >
        从一页内容，到一段有节奏的讲述
      </Text>
    </Slide>
  );
}
