import { Box, Step } from "@react-pptx/react-pptx";

import { InfoCard } from "../components/info-card.tsx";
import { PageFrame } from "../templates/page-frame.tsx";
import { colors } from "../theme.ts";

export function ComparisonPage() {
  return (
    <PageFrame
      slideId="comparison"
      title="让比较拥有共同尺度"
      subtitle="用相同的问题和维度比较，听众才能判断差异意味着什么。"
    >
      <Box
        style={{
          position: "absolute",
          left: 64,
          top: 230,
          width: 1152,
          display: "flex",
          gap: 40,
        }}
      >
        <Step n={1}>
          <InfoCard
            width={420}
            title="并列呈现"
            body="把两个选项放在同一条基线上，先说明共同的比较维度。"
            accent={colors.teal}
          />
        </Step>
        <Box
          style={{
            width: 692,
            padding: 32,
            display: "flex",
            flexDirection: "column",
            gap: 24,
            backgroundColor: colors.pale,
          }}
        >
          <InfoCard
            width={628}
            title="逐项对照"
            body="适用场景、信息密度、讲述顺序：沿着相同维度展开说明。"
            accent={colors.coral}
          />
        </Box>
      </Box>
    </PageFrame>
  );
}
