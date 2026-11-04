import { WorldMap, type RegionStory } from "../components/world-map.tsx";
import { PageFrame } from "../templates/page-frame.tsx";

const regions: readonly RegionStory[] = [
  {
    region: "northAmerica",
    title: "北美洲",
    description: "从一个地区开始建立上下文",
    color: "#138A83",
  },
  {
    region: "southAmerica",
    title: "南美洲",
    description: "沿着地理关系继续讲述",
    color: "#E98068",
  },
  {
    region: "europe",
    title: "欧洲",
    description: "小区域也需要清晰的辨识",
    color: "#4679A0",
  },
  {
    region: "africa",
    title: "非洲",
    description: "保留前面的强调，累积全貌",
    color: "#C28B3F",
  },
  {
    region: "asia",
    title: "亚洲",
    description: "让视线跟随当前的重点",
    color: "#7369A4",
  },
  {
    region: "oceania",
    title: "大洋洲",
    description: "六个地区，组成完整的分组",
    color: "#438B65",
  },
];

export function WorldPage() {
  return (
    <PageFrame
      slideId="world"
      title="在世界地图上逐区讲述"
      subtitle="一次关注一个地区，已讲述的内容继续留在画面中。"
    >
      <WorldMap x={64} y={166} width={1152} regions={regions} />
    </PageFrame>
  );
}
