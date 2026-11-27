import {
  Box,
  Chart,
  P,
  Shape,
  Text,
  preset,
  spinIn,
  transitions,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import { EditorialSlide, TelemetryPanel } from "../components/chrome.tsx";
import { createWavefrontPicture } from "../components/diagrams.ts";
import {
  PALETTE,
  editorialVisualLayout,
  jostFont,
  jostMediumFont,
} from "../theme.ts";

/** Slide 7: Adaptive optics and 4-site ground diversity with pie chart, picture-placeholder, and spinIn emblems. */
export function GroundNetworkSlide() {
  const wavefrontMap = useMemo(() => createWavefrontPicture(), []);

  const sites = [
    {
      site: "PALOMAR, USA",
      ap: "5.1 m  ·  1,712 m",
      clr: "74% Clear Sky",
      c: PALETTE.vermilion,
    },
    {
      site: "TABLE MTN, USA",
      ap: "1.0 m  ·  2,286 m",
      clr: "79% Clear Sky",
      c: PALETTE.cobalt,
    },
    {
      site: "TENERIFE, ESP",
      ap: "1.5 m  ·  2,390 m",
      clr: "83% Clear Sky",
      c: PALETTE.teal,
    },
    {
      site: "CANBERRA, AUS",
      ap: "2.4 m  ·  1,165 m",
      clr: "71% Clear Sky",
      c: PALETTE.amber,
    },
  ] as const;

  return (
    <EditorialSlide
      id="ground-network"
      layout={editorialVisualLayout}
      kicker="03 · FLIGHT CAMPAIGN & ROADMAP  |  GROUND SEGMENT & ADAPTIVE OPTICS"
      title="Cryo-SNSPD Wavefront Control & 4-Site Cloud Diversity"
      transition={transitions.fade({ duration: 450 })}
      extraPlaceholders={{ wavefrontMap }}
      notes="Atmospheric turbulence and cirrus clouds are the primary terrestrial bottlenecks. Combining 64-channel AO with 4 geographically uncorrelated OGS sites raised availability to 96.8%."
    >
      {/* Left: Pie Link Margin Allocation Chart inside slotted TelemetryPanel */}
      <Box
        style={{
          position: "absolute",
          left: 56,
          top: 116,
          width: 452,
          height: 540,
          display: "flex",
          flexDirection: "column",
        }}
      >
        <TelemetryPanel
          id="loss-budget-panel"
          title="12.0 dB Link Loss & Margin Budget"
          accentColor={PALETTE.teal}
        >
          <Chart
            id="loss-budget-pie"
            legend="bottom"
            groups={[
              {
                kind: "pie",
                firstSliceAngle: 0,
                data: {
                  kind: "shared-categories",
                  categories: [
                    "Atmosphere & Cirrus (2.8 dB)",
                    "AO Coupling Loss (1.9 dB)",
                    "Receiver Optics (1.6 dB)",
                    "SNSPD Quantum Loss (1.1 dB)",
                    "Pointing Loss (0.4 dB)",
                    "Net Positive Margin (4.2 dB)",
                  ],
                },
                dataLabels: {
                  percent: true,
                  style: {
                    fontFamily: jostFont,
                    fontSize: 10,
                    fontWeight: 700,
                    color: "#FFFFFF",
                  },
                },
                series: [
                  {
                    name: "dB Share",
                    values: [2.8, 1.9, 1.6, 1.1, 0.4, 4.2],
                    points: {
                      0: { color: PALETTE.cobalt },
                      1: { color: PALETTE.violet },
                      2: { color: PALETTE.amber },
                      3: { color: "#5A6578" },
                      4: { color: PALETTE.vermilion },
                      5: { color: PALETTE.teal },
                    },
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

      {/* Right Bottom: 4-Site Ground Diversity Row below the picture-placeholder */}
      <Box
        style={{
          position: "absolute",
          left: 528,
          top: 546,
          width: 696,
          height: 110,
          display: "grid",
          gridTemplateColumns: "1fr 1fr 1fr 1fr",
          gap: 10,
        }}
      >
        {sites.map((s) => (
          <Box
            key={s.site}
            style={{
              minWidth: 0,
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [10, 12, 10, 12],
              background: PALETTE.surface,
              border: { color: PALETTE.borderLight, width: 1, style: "solid" },
              borderRadius: 10,
            }}
          >
            <Box
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <Shape
                geometry={preset("rect")}
                style={{ width: 22, height: 3, background: s.c }}
              />
              <Shape
                geometry={preset("star4")}
                enter={spinIn({ duration: 420 })}
                style={{ width: 10, height: 10, background: s.c }}
              />
            </Box>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 10,
                fontWeight: 700,
                color: PALETTE.ink,
              }}
            >
              <P>{s.site}</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostMediumFont,
                fontSize: 11,
                fontWeight: 500,
                color: PALETTE.inkBody,
              }}
            >
              <P>{s.ap}</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 11,
                fontWeight: 700,
                color: s.c,
              }}
            >
              <P>{s.clr}</P>
            </Text>
          </Box>
        ))}
      </Box>
    </EditorialSlide>
  );
}
