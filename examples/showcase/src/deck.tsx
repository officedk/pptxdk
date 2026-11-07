import { Deck, createMorphScope } from "@react-pptx/react-pptx";

import { CalendarPage } from "./pages/calendar.tsx";
import { ClosingPage } from "./pages/closing.tsx";
import { CombinationPage } from "./pages/combination.tsx";
import { ComparisonPage } from "./pages/comparison.tsx";
import { DevicesPage } from "./pages/devices.tsx";
import { GroupingPage, RevealPage } from "./pages/grouping.tsx";
import { HierarchyPage } from "./pages/hierarchy.tsx";
import { OpeningPage } from "./pages/opening.tsx";
import { ProcessPage } from "./pages/process.tsx";
import { ProgressPage } from "./pages/progress.tsx";
import { WorldPage } from "./pages/world.tsx";
import { GalleryPage } from "./templates/gallery-page.tsx";
import { colors, theme } from "./theme.ts";

/** Slide sequence is centralized at the entrypoint; each slide maintains its own content and narrative order. */
export function ShowcaseDeck() {
  const gallery = createMorphScope();
  return (
    <Deck
      width={1280}
      height={720}
      theme={theme}
      style={{ fontFamily: "major", fontSize: 24, color: colors.ink }}
    >
      <OpeningPage />
      <HierarchyPage />
      <GroupingPage />
      <RevealPage />
      <ComparisonPage />
      <ProcessPage />
      <ProgressPage />
      <CalendarPage />
      <DevicesPage />
      <WorldPage />
      <GalleryPage focus={0} scope={gallery} />
      <GalleryPage focus={1} scope={gallery} />
      <GalleryPage focus={2} scope={gallery} />
      <GalleryPage focus={3} scope={gallery} />
      <CombinationPage />
      <ClosingPage />
    </Deck>
  );
}
