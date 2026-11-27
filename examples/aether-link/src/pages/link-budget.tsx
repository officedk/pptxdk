import {
  Box,
  Chart,
  Mount,
  P,
  Shape,
  Text,
  preset,
  stretch,
  transitions,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import {
  EditorialSlide,
  TakeawayStrip,
  TelemetryPanel,
} from "../components/chrome.tsx";
import { createBeamDivergenceDiagram } from "../components/diagrams.ts";
import { PALETTE, jostFont, jostMediumFont, telemetryMorph } from "../theme.ts";

/** Slide 3: Beam divergence physics with slotted TelemetryPanel, direct Mount adoption, clustered column chart, and stretch motion. */
export function LinkBudgetSlide() {
  const beamDiagram = useMemo(
    () =>
      createBeamDivergenceDiagram({
        id: "divergence-group",
        rfLabel: "32 GHz Ka-Band (3.8 mrad)",
        rfDetail: "570,000 km cone (45x Earth)",
        opticalLabel: "1550 nm Optical (11.4 µrad)",
        opticalDetail: "1,700 km spot (0.13x Earth)",
        caption: "Probe Aperture: 22 cm Optical vs 3.2 m RF",
      }),
    [],
  );

  return (
    <EditorialSlide
      id="link-budget"
      kicker="01 · MISSION & PHYSICS  |  DIFFRACTION & LINK SCALING"
      title="6,000x Carrier Frequency Leap Compresses Beam Divergence"
      transition={transitions.morph({
        duration: 850,
        by: "object",
        fallback: transitions.fade({ duration: 500 }),
      })}
      notes="Notice how the 1550 nm C-band carrier badge morphs seamlessly from Slide 2 into the physics diagram header. Moving from 32 GHz RF to 193 THz near-IR narrows beam divergence by 330x."
    >
      <Box style={{ flex: 1, display: "flex", flexDirection: "row", gap: 20 }}>
        {/* Left: Slotted Vanilla TelemetryPanel wrapping Morph pill, Mount diagram, and diffraction callout */}
        <TelemetryPanel
          id="link-footprint-panel"
          title="BEAM FOOTPRINT AT 1.0 AU"
          accentColor={PALETTE.vermilion}
          width={484}
        >
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              justifyContent: "flex-end",
            }}
          >
            <Shape
              id="morph-laser-pill-s3"
              morphKey={telemetryMorph.handle("laser-carrier-pill")}
              geometry={preset("roundRect")}
              style={{
                width: 244,
                height: 26,
                background: PALETTE.ink,
                borderRadius: 13,
                fontFamily: jostFont,
                fontSize: 10,
                fontWeight: 700,
                letterSpacing: 1.0,
                color: "#FFB094",
                verticalAlign: "middle",
              }}
            >
              <P style={{ align: "center" }}>1550.12 nm · +50.4 dB GAIN</P>
            </Shape>
          </Box>

          <Mount node={beamDiagram} />

          <Box
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 3,
              padding: [10, 12, 10, 12],
              background: PALETTE.paper,
              borderRadius: 8,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 11,
                fontWeight: 700,
                color: PALETTE.vermilion,
              }}
            >
              <P>DIFFRACTION SCALING</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 12,
                lineHeight: "17px",
                color: PALETTE.inkBody,
              }}
            >
              <P>
                Divergence scales with wavelength (1.55 µm vs 9.37 mm), yielding
                330x tighter beam angle.
              </P>
            </Text>
          </Box>
        </TelemetryPanel>

        {/* Right: Slotted Vanilla TelemetryPanel wrapping Clustered Column Chart */}
        <TelemetryPanel
          id="link-rate-panel"
          title="Sustained Downlink Rate vs. Heliocentric Range (Mbps)"
          badge="0.22–2.20 AU"
          accentColor={PALETTE.cobalt}
        >
          <Chart
            id="ka-vs-optical-chart"
            legend="top"
            axes={{
              rangeCat: {
                kind: "category",
                position: "bottom",
                style: {
                  fontFamily: jostMediumFont,
                  fontSize: 12,
                  fontWeight: 500,
                  color: PALETTE.inkBody,
                },
              },
              mbpsVal: {
                kind: "value",
                position: "left",
                min: 0,
                max: 300,
                majorUnit: 50,
                gridlines: { color: "#E5E2D9", width: 1 },
                style: {
                  fontFamily: jostFont,
                  fontSize: 11,
                  color: PALETTE.inkMuted,
                },
                title: {
                  text: "Downlink Rate (Mbps)",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 11,
                    fontWeight: 700,
                    color: PALETTE.inkMuted,
                  },
                },
              },
            }}
            groups={[
              {
                kind: "column",
                axes: { category: "rangeCat", value: "mbpsVal" },
                gapWidth: 110,
                data: {
                  kind: "shared-categories",
                  categories: [
                    "0.22 AU (Lunar Escape)",
                    "0.50 AU (Near-Earth)",
                    "1.52 AU (Mars Opp.)",
                    "2.20 AU (Belt Inner)",
                  ],
                },
                dataLabels: {
                  value: true,
                  position: "outsideEnd",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 11,
                    fontWeight: 700,
                    color: PALETTE.ink,
                  },
                },
                series: [
                  {
                    name: "1550 nm Optical (FLT-220, 76 W)",
                    values: [267, 118, 25, 11.4],
                    color: PALETTE.vermilion,
                  },
                  {
                    name: "32 GHz Ka-Band RF (3.2 m HGA, 180 W)",
                    values: [38, 9.2, 1.4, 0.6],
                    color: PALETTE.cobalt,
                  },
                ],
              },
            ]}
            style={{
              flex: 1,
              fontFamily: jostFont,
              fontSize: 11,
            }}
          />
        </TelemetryPanel>
      </Box>

      <TakeawayStrip
        tag="LINK ADVANTAGE"
        accentColor={PALETTE.vermilion}
        enter={stretch({ duration: 420 })}
        text="At 1.52 AU Mars opposition, FLT-220 delivers 17.8x the data rate of Ka-Band RF while consuming 58% less spacecraft bus power."
      />
    </EditorialSlide>
  );
}
