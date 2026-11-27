import {
  Box,
  MotionScope,
  Shape,
  Step,
  fade,
  preset,
  transitions,
  useMotionScope,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import { EditorialSlide, MetricTile } from "../components/chrome.tsx";
import { createCampaignComboChart } from "../components/diagrams.ts";
import { PALETTE, editorialChartLayout } from "../theme.ts";

/** Slide 6: 14-month cruise telemetry using chart-placeholder, staggered MetricTiles, and Step. */
export function TelemetryCampaignSlide() {
  const kpiScope = useMotionScope();
  const mainChart = useMemo(() => createCampaignComboChart(), []);

  return (
    <EditorialSlide
      id="telemetry-campaign"
      layout={editorialChartLayout}
      kicker="03 · FLIGHT CAMPAIGN & ROADMAP  |  14-MONTH CRUISE TELEMETRY"
      title="41.8 Terabits Downlinked Across 0.08 AU to 2.20 AU Cruise"
      transition={transitions.fade({ duration: 500 })}
      extraPlaceholders={{ mainChart }}
      notes="This dual-axis combo chart occupies a native chart-placeholder slot and overlays bi-monthly science volume against peak sustained Mbps."
    >
      {/* Card backdrop positioned behind the native chart-placeholder */}
      <Shape
        geometry={preset("roundRect")}
        style={{
          position: "absolute",
          left: 56,
          top: 116,
          width: 786,
          height: 540,
          zIndex: 0,
          background: PALETTE.surface,
          border: { color: PALETTE.borderLight, width: 1, style: "solid" },
          borderRadius: 12,
        }}
      />

      {/* Right: 3 Stacked Campaign KPI Tiles participating in Step and MotionScope stagger */}
      <Step n={1}>
        <MotionScope scope={kpiScope} mode="stagger" each={100}>
          <Box
            style={{
              position: "absolute",
              left: 862,
              top: 116,
              width: 362,
              height: 540,
              display: "flex",
              flexDirection: "column",
              gap: 12,
            }}
          >
            <MetricTile
              id="camp-tile-1"
              label="TOTAL CAMPAIGN RETURN"
              value="41.8"
              unit="Terabits"
              detail="Exceeds cumulative data returned by all prior Mars Ka-band orbiters combined."
              accentColor={PALETTE.cobalt}
              enter={fade({ duration: 360 })}
            />
            <MetricTile
              id="camp-tile-2"
              label="POST-FEC FRAME INTEGRITY"
              value="99.94"
              unit="%"
              detail="CCSDS SCPPM + CFDP automatic nak-retransmit recovered scintillation fades."
              accentColor={PALETTE.teal}
              enter={fade({ duration: 360 })}
            />
            <MetricTile
              id="camp-tile-3"
              label="HD VIDEO DEMO AT 0.22 AU"
              value="4K60"
              unit="Live"
              detail="Streamed uncompressed multispectral star-field telemetry with 110 s light time."
              accentColor={PALETTE.vermilion}
              dark={true}
              enter={fade({ duration: 360 })}
            />
          </Box>
        </MotionScope>
      </Step>
    </EditorialSlide>
  );
}
