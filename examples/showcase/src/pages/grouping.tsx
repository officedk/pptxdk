import { CardsPage } from "../templates/cards-page.tsx";
import { colors } from "../theme.ts";

export function GroupingPage() {
  return (
    <CardsPage
      slideId="grouping"
      title="把相关内容放在一起"
      subtitle="相同的容器、对齐和间距，帮助听众先看见分组。"
      items={[
        {
          title: "分组",
          body: "相关的信息靠近，让每张卡片只承担一个主题。",
          accent: colors.teal,
        },
        {
          title: "对齐",
          body: "统一标题的起点，减少视线来回寻找的负担。",
          accent: colors.coral,
        },
        {
          title: "一致性",
          body: "相同层级使用相同的样式，让差异留给内容。",
          accent: colors.ink,
        },
      ]}
    />
  );
}

export function RevealPage() {
  return (
    <CardsPage
      slideId="reveal"
      title="一次引入一个重点"
      subtitle="先说清一个观点，再给出依据；让出现的顺序服务讲述。"
      items={[
        {
          title: "先给出中心观点",
          body: "用一句话告诉听众：接下来最值得关注的是什么。",
          accent: colors.teal,
        },
        {
          title: "再补充理解它的依据",
          body: "每次点击只增加一个重点，留出消化信息的时间。",
          accent: colors.coral,
        },
      ]}
    />
  );
}
