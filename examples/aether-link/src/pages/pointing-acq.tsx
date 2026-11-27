import {
  Box,
  Chart,
  P,
  Shape,
  Span,
  Text,
  preset,
  transitions,
} from "@officedk/pptxdk/react";

import {
  EditorialSlide,
  TakeawayStrip,
  TelemetryPanel,
} from "../components/chrome.tsx";
import { PALETTE, cassegrainApertureImage, jostFont } from "../theme.ts";

/** Slide 5: Pointing, acquisition & tracking with flip transition, image-fill optical viewport, and scatter error bars. */
export function PointingAcqSlide() {
  return (
    <EditorialSlide
      id="pointing-acq"
      kicker="02 · SYSTEM ARCHITECTURE  |  POINTING, ACQUISITION & TRACKING"
      title="1.5 kHz Piezo FSM Suppresses Micro-Vibration to 0.21 µrad"
      transition={transitions.flip({
        direction: "left",
        duration: 600,
        fallback: transitions.fade({ duration: 450 }),
      })}
      notes="At 11.4 microradians FWHM, a 1 microradian pointing error costs nearly 1.8 dB of link margin. Our 3-tier isolation architecture isolates reaction wheel harmonics by 44 dB."
    >
      <Box style={{ flex: 1, display: "flex", flexDirection: "row", gap: 20 }}>
        {/* Left: PAT Sequence + Timing Strip + Image-Fill Optical Sensor Inset */}
        <Box
          style={{
            width: 472,
            display: "flex",
            flexDirection: "column",
            justifyContent: "space-between",
            gap: 14,
          }}
        >
          <TelemetryPanel
            id="pat-sequence-panel"
            title="CLOSED-LOOP ACQUISITION (< 9.4 SEC)"
            accentColor={PALETTE.vermilion}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 12.5,
                lineHeight: "19px",
                color: PALETTE.inkBody,
              }}
            >
              <P style={{ list: { type: "number" }, spacingAfter: 6 }}>
                <Span style={{ fontWeight: 700, color: PALETTE.ink }}>
                  Blind Ephemeris Scan:{" "}
                </Span>
                Star-tracker attitude aligns 1.2 mrad InGaAs acquisition
                quad-cell.
              </P>
              <P style={{ list: { type: "number" }, spacingAfter: 6 }}>
                <Span style={{ fontWeight: 700, color: PALETTE.ink }}>
                  1064 nm Beacon Pull-In:{" "}
                </Span>
                Ground 5 kW uplink centroid locks coarse hexapod within 18 µrad.
              </P>
              <P style={{ list: { type: "number" }, spacingAfter: 6 }}>
                <Span style={{ fontWeight: 700, color: PALETTE.ink }}>
                  1.5 kHz Piezo FSM Handoff:{" "}
                </Span>
                Voice-coil + PZT tip-tilt mirror drives residual below 0.25
                µrad.
              </P>
              <P style={{ list: { type: "number" } }}>
                <Span style={{ fontWeight: 700, color: PALETTE.ink }}>
                  Point-Ahead Feedforward:{" "}
                </Span>
                2v/c relativistic aberration offset (up to 340 µrad) injected
                digitally.
              </P>
            </Text>

            <Box
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr 1fr",
                gap: 8,
                padding: [10, 12, 10, 12],
                background: PALETTE.paper,
                border: {
                  color: PALETTE.borderLight,
                  width: 1,
                  style: "solid",
                },
                borderRadius: 8,
              }}
            >
              {[
                { ph: "1. SCAN LOCK", dt: "2.9 s" },
                { ph: "2. HEXAPOD", dt: "4.6 s" },
                { ph: "3. PZT FINE", dt: "1.4 s" },
              ].map((t) => (
                <Box
                  key={t.ph}
                  style={{ display: "flex", flexDirection: "column", gap: 2 }}
                >
                  <Text
                    style={{
                      width: "100%",
                      height: "auto",
                      fontSize: 9,
                      fontWeight: 700,
                      color: PALETTE.inkMuted,
                    }}
                  >
                    <P>{t.ph}</P>
                  </Text>
                  <Text
                    style={{
                      width: "100%",
                      height: "auto",
                      fontSize: 15,
                      fontWeight: 700,
                      color: PALETTE.ink,
                    }}
                  >
                    <P>{t.dt}</P>
                  </Text>
                </Box>
              ))}
            </Box>
          </TelemetryPanel>

          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "center",
              gap: 14,
              padding: [12, 16, 12, 16],
              background: PALETTE.cobaltSoft,
              border: { color: "#BDD3FA", width: 1, style: "solid" },
              borderRadius: 12,
            }}
          >
            {/* Native Shape image-fill aperture viewport */}
            <Shape
              id="pat-aperture-fill"
              geometry={preset("roundRect")}
              style={{
                width: 72,
                height: 72,
                flexShrink: 0,
                background: { image: cassegrainApertureImage, fit: "cover" },
                border: { color: PALETTE.cobalt, width: 1.5, style: "solid" },
                borderRadius: 10,
              }}
            />
            <Box
              style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                gap: 6,
              }}
            >
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontFamily: jostFont,
                  fontSize: 11,
                  fontWeight: 700,
                  letterSpacing: 1.2,
                  color: PALETTE.cobalt,
                }}
              >
                <P>DISTURBANCE REJECTION HIGHLIGHTS</P>
              </Text>
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontFamily: jostFont,
                  fontSize: 12,
                  lineHeight: "17px",
                  color: PALETTE.ink,
                }}
              >
                <P
                  style={{
                    list: { type: "bullet", character: "•" },
                    spacingAfter: 3,
                  }}
                >
                  6-strut elastomer mount attenuates &gt; 40 Hz RWA spikes
                </P>
                <P style={{ list: { type: "bullet", character: "•" } }}>
                  Pointing loss held to 0.42 dB vs. 1.50 dB allocation
                </P>
              </Text>
            </Box>
          </Box>
        </Box>

        {/* Right: Native XY Scatter Chart with Smooth Curves, Trendline & Error Bars */}
        <TelemetryPanel
          id="pat-scatter-panel"
          title="Radial Pointing Jitter vs. Spacecraft Disturbance Frequency"
          badge="IN-FLIGHT PSD"
          accentColor={PALETTE.cobalt}
        >
          <Chart
            id="pat-jitter-scatter"
            legend="top"
            axes={{
              freqX: {
                kind: "value",
                position: "bottom",
                min: 0,
                max: 250,
                majorUnit: 50,
                gridlines: { color: "#ECE9DF", width: 1 },
                style: {
                  fontFamily: jostFont,
                  fontSize: 11,
                  color: PALETTE.inkMuted,
                },
                title: {
                  text: "Disturbance Frequency (Hz)",
                  style: {
                    fontFamily: jostFont,
                    fontSize: 11,
                    fontWeight: 700,
                    color: PALETTE.inkMuted,
                  },
                },
              },
              jitterY: {
                kind: "value",
                position: "left",
                min: 0,
                max: 4.0,
                majorUnit: 0.8,
                gridlines: { color: "#E5E2D9", width: 1 },
                style: {
                  fontFamily: jostFont,
                  fontSize: 11,
                  color: PALETTE.inkMuted,
                },
                title: {
                  text: "Radial Jitter (µrad RMS)",
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
                kind: "scatter",
                axes: { x: "freqX", y: "jitterY" },
                data: { kind: "xy" },
                line: "smooth",
                markers: true,
                series: [
                  {
                    name: "Uncompensated Bus + Strut Input",
                    x: [5, 25, 60, 100, 140, 190, 240],
                    y: [3.65, 3.1, 2.45, 1.92, 1.35, 0.82, 0.54],
                    color: PALETTE.cobalt,
                    marker: { symbol: "circle", size: 7 },
                    trendlines: [
                      {
                        type: "linear",
                        line: { color: "#90B2F4", width: 1, dash: "dash" },
                      },
                    ],
                  },
                  {
                    name: "Closed-Loop 1.5 kHz Piezo FSM Output",
                    x: [5, 25, 60, 100, 140, 190, 240],
                    y: [0.09, 0.14, 0.21, 0.24, 0.19, 0.15, 0.11],
                    color: PALETTE.vermilion,
                    marker: { symbol: "diamond", size: 7 },
                    errorBars: {
                      y: {
                        type: "custom",
                        plus: [0.04, 0.05, 0.06, 0.07, 0.05, 0.04, 0.03],
                        minus: [0.03, 0.04, 0.05, 0.05, 0.04, 0.03, 0.03],
                        line: { color: PALETTE.vermilion, width: 1.2 },
                      },
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

      <TakeawayStrip
        tag="CONTROL MARGIN"
        accentColor={PALETTE.teal}
        text="Integrated 0.1–250 Hz radial jitter measured 0.21 µrad RMS in flight — 2.8x tighter than the 0.60 µrad requirement."
      />
    </EditorialSlide>
  );
}
