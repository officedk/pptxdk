import { Box, Step } from "@react-pptx/react-pptx";

import { InfoCard, type CardContent } from "../components/info-card.tsx";
import { PageFrame } from "./page-frame.tsx";

/** 项目数决定等宽列；标题和正文的替换由卡片内部布局吸收。 */
export function CardsPage({
  slideId,
  title,
  subtitle,
  items,
}: {
  readonly slideId: string;
  readonly title: string;
  readonly subtitle: string;
  readonly items: readonly CardContent[];
}) {
  const width = (1152 - (items.length - 1) * 24) / items.length;
  return (
    <PageFrame slideId={slideId} title={title} subtitle={subtitle}>
      <Box
        style={{
          position: "absolute",
          left: 64,
          top: 230,
          width: 1152,
          display: "flex",
          gap: 24,
        }}
      >
        {items.map((item, index) => (
          <Step key={item.title} n={index + 1}>
            <InfoCard {...item} width={width} />
          </Step>
        ))}
      </Box>
    </PageFrame>
  );
}
