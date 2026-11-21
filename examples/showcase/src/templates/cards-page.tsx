import { Box, Step } from "pptxdk/react";

import { InfoCard, type CardContent } from "../components/info-card.tsx";
import { PageFrame } from "./page-frame.tsx";

/** Cards share available width; each card encapsulates its own reading order. */
export function CardsPage({
  id,
  title,
  subtitle,
  items,
}: {
  readonly id: string;
  readonly title: string;
  readonly subtitle: string;
  readonly items: readonly CardContent[];
}) {
  return (
    <PageFrame id={id} title={title} subtitle={subtitle}>
      <Box
        style={{
          position: "absolute",
          left: 80,
          top: 234,
          width: 1120,
          display: "flex",
          gap: 24,
        }}
      >
        {items.map((item, index) => (
          <Step key={`${item.title}-${index}`} n={index + 1}>
            <InfoCard
              {...item}
              index={index}
              style={{ flex: 1, minWidth: 0 }}
            />
          </Step>
        ))}
      </Box>
    </PageFrame>
  );
}
