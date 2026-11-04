import { Text } from "@react-pptx/react-pptx";

import { Progress } from "../components/progress.tsx";
import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

export function ProgressPage() {
  return (
    <PageFrame
      slideId="progress"
      title="进度需要分母"
      subtitle="数值告诉我们多少，分母告诉我们距离完成还有多远。"
    >
      <Progress completed={3} total={4} />
      <Text
        x={64}
        y={580}
        width={1100}
        height={60}
        style={{ fontSize: 20, color: colors.muted }}
      >
        教学示意：四个讲述步骤，已完成其中三个。三种表达使用同一组数据。
      </Text>
    </PageFrame>
  );
}
