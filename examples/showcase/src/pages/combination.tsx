import { Box, Step, Text } from "@react-pptx/react-pptx";

import { InfoCard } from "../components/info-card.tsx";
import { ProcessPath } from "../components/process-path.tsx";
import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

export function CombinationPage() {
  return (
    <PageFrame
      slideId="combination"
      title="把几种表达组合起来"
      subtitle="不同表达各司其职，围绕一个清楚的中心观点展开。"
    >
      <Box style={{ position: "absolute", left: 64, top: 205, width: 360 }}>
        <Step n={1}>
          <InfoCard
            width={360}
            title="一个中心"
            body="卡片凝练核心主张，流程交代展开顺序，留白帮助分组。"
            accent={colors.teal}
          />
        </Step>
      </Box>
      <ProcessPath
        x={490}
        y={240}
        width={710}
        firstStep={2}
        nodes={[
          { title: "建立层级", body: "找到阅读起点" },
          { title: "组织关系", body: "连起关键内容" },
        ]}
      />
      <Text
        x={490}
        y={500}
        width={710}
        height={100}
        style={{ fontSize: 26, color: colors.muted }}
      >
        每增加一种表达，都应该帮助听众更快理解这一页。
      </Text>
    </PageFrame>
  );
}
