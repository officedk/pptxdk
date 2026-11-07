import { Deck } from "@react-pptx/react-pptx";

import { AlignmentPage } from "./pages/alignment.tsx";
import { DiffusionPage } from "./pages/diffusion.tsx";
import { FrontierPage } from "./pages/frontier.tsx";
import { OpeningPage } from "./pages/opening.tsx";
import { PredictionPage } from "./pages/prediction.tsx";
import { QuestionsPage } from "./pages/questions.tsx";
import { ScalePage } from "./pages/scale.tsx";
import { SourcesPage } from "./pages/sources.tsx";
import { SynthesisPage } from "./pages/synthesis.tsx";
import { SystemsPage } from "./pages/systems.tsx";
import { TimelinePage } from "./pages/timeline.tsx";
import { TransformerPage } from "./pages/transformer.tsx";
import { VaeGanPage } from "./pages/vae-gan.tsx";
import { colors, theme } from "./theme.ts";

/** Slide sequence is centralized at the entrypoint; each slide organizes copy and diagrams around a historical turning point. */
export function GenerativeAiHistoryDeck() {
  return (
    <Deck
      width={1280}
      height={720}
      title="机器如何学会创造：生成式人工智能简史"
      theme={theme}
      style={{ fontFamily: "body", fontSize: 18, color: colors.ink }}
    >
      <OpeningPage />
      <TimelinePage />
      <PredictionPage />
      <VaeGanPage />
      <TransformerPage />
      <ScalePage />
      <DiffusionPage />
      <AlignmentPage />
      <SystemsPage />
      <FrontierPage />
      <SynthesisPage />
      <QuestionsPage />
      <SourcesPage page={1} />
      <SourcesPage page={2} />
    </Deck>
  );
}
