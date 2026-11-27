import {
  Deck,
  Section,
  Shape,
  SlideLayout,
  SlideMaster,
  defaultColorMap,
  preset,
} from "@officedk/pptxdk/react";

import {
  TelemetryCampaignSlide,
  GroundNetworkSlide,
  SubsystemMatrixSlide,
  AnomalyPostmortemSlide,
  RoadmapCloseSlide,
} from "./pages/flight-campaign.tsx";
import {
  CoverSlide,
  ExecSummarySlide,
  LinkBudgetSlide,
} from "./pages/mission-physics.tsx";
import {
  ArchitectureSlide,
  PointingAcqSlide,
} from "./pages/system-architecture.tsx";
import {
  PALETTE,
  aetherTheme,
  anomalyBriefLayout,
  closingDarkLayout,
  editorialChartLayout,
  editorialLightLayout,
  editorialTableLayout,
  editorialVisualLayout,
  heroDarkLayout,
  jostFont,
  jostMediumFont,
} from "./theme.ts";

/** 10-slide pure-English deep-space optical telemetry engineering review deck. */
export function AetherLinkDeck() {
  return (
    <Deck
      width={1280}
      height={720}
      title="AETHER-LINK: Deep-Space Optical Telemetry & Cislunar Relay Architecture"
      theme={aetherTheme}
      style={{
        fontFamily: jostFont,
        color: PALETTE.ink,
      }}
    >
      <SlideMaster
        id="aether-master"
        name="Aether Monograph Master"
        colorMap={defaultColorMap}
        headerFooter={{
          dateTime: true,
          footer: true,
          slideNumber: true,
        }}
        footerPlaceholders={{
          dateTime: {
            text: "NOV 2026  ·  CDR-4 DOSSIER",
            style: {
              position: "absolute",
              left: 56,
              top: 686,
              width: 240,
              height: 18,
              fontFamily: jostMediumFont,
              fontSize: 10,
              fontWeight: 500,
              letterSpacing: 0.8,
              color: PALETTE.inkMuted,
            },
          },
          footer: {
            text: "AETHER-LINK  ·  1550 nm DEEP-SPACE LASER TELEMETRY & CISLUNAR TRUNK",
            style: {
              position: "absolute",
              left: 340,
              top: 686,
              width: 600,
              height: 18,
              fontFamily: jostMediumFont,
              fontSize: 10,
              fontWeight: 500,
              letterSpacing: 0.8,
              align: "center",
              color: PALETTE.inkMuted,
            },
          },
          slideNumber: {
            style: {
              position: "absolute",
              left: 1124,
              top: 686,
              width: 100,
              height: 18,
              fontFamily: jostFont,
              fontSize: 10,
              fontWeight: 700,
              align: "right",
              color: PALETTE.ink,
            },
          },
        }}
      >
        <Shape
          geometry={preset("rect")}
          style={{
            position: "absolute",
            left: 56,
            top: 102,
            width: 1168,
            height: 1,
            background: PALETTE.borderLight,
          }}
        />
        <Shape
          geometry={preset("rect")}
          style={{
            position: "absolute",
            left: 56,
            top: 102,
            width: 72,
            height: 2,
            background: PALETTE.vermilion,
          }}
        />
        <Shape
          geometry={preset("rect")}
          style={{
            position: "absolute",
            left: 56,
            top: 678,
            width: 1168,
            height: 1,
            background: PALETTE.borderLight,
          }}
        />

        <SlideLayout layout={heroDarkLayout} />
        <SlideLayout layout={editorialLightLayout} />
        <SlideLayout layout={editorialChartLayout} />
        <SlideLayout layout={editorialVisualLayout} />
        <SlideLayout layout={editorialTableLayout} />
        <SlideLayout layout={anomalyBriefLayout} />
        <SlideLayout layout={closingDarkLayout} />
      </SlideMaster>

      <Section name="01 · Mission & Physics">
        <CoverSlide />
        <ExecSummarySlide />
        <LinkBudgetSlide />
      </Section>

      <Section name="02 · System Architecture">
        <ArchitectureSlide />
        <PointingAcqSlide />
      </Section>

      <Section name="03 · Flight Campaign & Roadmap">
        <TelemetryCampaignSlide />
        <GroundNetworkSlide />
        <SubsystemMatrixSlide />
        <AnomalyPostmortemSlide />
        <RoadmapCloseSlide />
      </Section>
    </Deck>
  );
}
