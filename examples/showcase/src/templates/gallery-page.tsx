import {
  Group,
  Shape,
  Slide,
  Text,
  preset,
  transitions,
  type MorphScope,
} from "@react-pptx/react-pptx";

import { colors } from "../theme.ts";

export const galleryTopics = [
  { id: "contrast", title: "对比", description: "共同尺度，让差异可见" },
  { id: "process", title: "流程", description: "连接节点，让方向可见" },
  { id: "hierarchy", title: "层级", description: "建立主次，让结构可见" },
  { id: "space", title: "空间", description: "改变位置，让关系可见" },
] as const;

/** 四组内容始终保留在连续画布上；整体移动和缩放由同一连续性身份衔接。 */
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
  const scale = index === focus ? 1 : 0.62;
  const x = 96 + (index - focus) * 1180;
  const y = index === focus ? 180 : 245;
  const shape = (
    key: string,
    left: number,
    top: number,
    width: number,
    height: number,
    fill: string,
  ) => (
    <Shape
      key={key}
      x={x + left * scale}
      y={y + top * scale}
      width={width * scale}
      height={height * scale}
      fill={fill}
      geometry={preset("roundRect")}
    />
  );
  const label = (
    key: string,
    text: string,
    left: number,
    top: number,
    width: number,
    size = 26,
    color: string = colors.paper,
  ) => (
    <Text
      key={key}
      x={x + left * scale}
      y={y + top * scale}
      width={width * scale}
      height={Math.max(60, size * 1.6) * scale}
      style={{ fontSize: size * scale, color }}
    >
      {text}
    </Text>
  );
  return (
    <Group
      testId={`gallery-panel-${topic.id}`}
      morphKey={scope.handle(`panel-${topic.id}`)}
    >
      {label("title", topic.title, 0, 0, 1050, 44)}
      {label("description", topic.description, 0, 80, 1050, 24, "#ACBEC9")}
      {index === 0 && (
        <>
          {shape("left", 0, 160, 490, 210, "#24455F")}
          {shape("right", 560, 160, 490, 210, "#24455F")}
          {label("left-title", "相同的起点", 28, 180, 440)}
          {label("right-title", "清楚的差异", 588, 180, 440)}
          {shape("bar-left", 28, 280, 250, 24, colors.teal)}
          {shape("bar-right", 588, 280, 400, 24, colors.coral)}
        </>
      )}
      {index === 1 && (
        <>
          {shape("connector", 80, 257, 880, 4, "#ACBEC9")}
          {["提出问题", "组织材料", "形成表达"].map((text, i) => (
            <Group key={text}>
              {shape(
                `node-${i}`,
                i * 370,
                178,
                300,
                162,
                i === 1 ? colors.teal : "#24455F",
              )}
              {label(`node-label-${i}`, text, i * 370 + 44, 233, 240, 28)}
            </Group>
          ))}
        </>
      )}
      {index === 2 && (
        <>
          {shape("stem", 518, 160, 4, 120, "#ACBEC9")}
          {shape("branch", 160, 278, 720, 4, "#ACBEC9")}
          {shape("root", 345, 142, 350, 82, colors.teal)}
          {label("root-text", "一个中心观点", 397, 156, 290)}
          {["关键依据", "支撑材料", "补充说明"].map((text, i) => (
            <Group key={text}>
              {shape(`leaf-${i}`, i * 370, 300, 300, 82, "#24455F")}
              {label(`leaf-text-${i}`, text, i * 370 + 44, 313, 250)}
            </Group>
          ))}
        </>
      )}
      {index === 3 && (
        <>
          {shape("field", 0, 148, 1050, 250, "#24455F")}
          {shape("near-a", 80, 208, 132, 132, colors.teal)}
          {shape("near-b", 235, 208, 132, 132, colors.coral)}
          {shape("far", 760, 208, 132, 132, "#ACBEC9")}
          {label("near-text", "靠近，形成一组", 80, 344, 500, 20)}
          {label("far-text", "留白，区分关系", 688, 344, 360, 20)}
        </>
      )}
    </Group>
  );
}

/** 四个聚焦状态共用整页模板，导航直接链接稳定页面标识，支持再次访问。 */
export function GalleryPage({
  focus,
  scope,
}: {
  readonly focus: number;
  readonly scope: MorphScope;
}) {
  return (
    <Slide
      slideId={`gallery-${galleryTopics[focus]!.id}`}
      background={colors.ink}
      transition={transitions.morph({
        duration: 1200,
        fallback: transitions.fade({ duration: 500 }),
      })}
    >
      <Text
        x={64}
        y={44}
        width={1100}
        height={55}
        style={{ fontSize: 24, color: "#ACBEC9" }}
      >
        信息组织画廊 · 从不同角度看同一张画布
      </Text>
      {galleryTopics.map((topic, index) => (
        <GalleryPanel
          key={topic.id}
          index={index}
          focus={focus}
          scope={scope}
        />
      ))}
      <Shape
        x={64 + focus * 288}
        y={686}
        width={264}
        height={4}
        fill={colors.teal}
        geometry={preset("roundRect")}
        morphKey={scope.handle("navigation-selection")}
      />
      {galleryTopics.map((topic, index) => (
        <Shape
          key={topic.id}
          morphKey={scope.handle(`navigation-${topic.id}`)}
          link={{ slideId: `gallery-${topic.id}` }}
          testId={`navigation-${topic.id}`}
          x={64 + index * 288}
          y={624}
          width={264}
          height={56}
          fill={colors.ink}
          text={topic.title}
          textStyle={{
            fontSize: 24,
            color: index === focus ? "#7FC9C1" : colors.paper,
            align: "center",
            verticalAlign: "middle",
          }}
        />
      ))}
    </Slide>
  );
}
