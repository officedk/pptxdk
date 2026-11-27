import {
  Box,
  Chart,
  P,
  Shape,
  Text,
  preset,
  transitions,
  zoom,
} from "@officedk/pptxdk/react";

import { EditorialSlide, TelemetryPanel } from "../components/chrome.tsx";
import { PALETTE, anomalyBriefLayout, jostFont } from "../theme.ts";

/** Slide 9: Engineering postmortem with layout colorMapOverride + themeOverride and zoom timeline cards. */
export function AnomalyPostmortemSlide() {
  const events = [
    {
      t: "DOY 194 · 04:12Z",
      title: "3.0° Solar Conjunction",
      desc: "Solar flux warmed SiC metering ring +14 K",
      c: PALETTE.amber,
    },
    {
      t: "DOY 194 · 06:40Z",
      title: "2.85 µrad Bias Detected",
      desc: "Bore-sight offset reduced link SNR 3.1 dB",
      c: "accent1",
    },
    {
      t: "DOY 194 · 09:05Z",
      title: "Kalman LUT Uploaded",
      desc: "12-term thermo-elastic LUT patched via Ka",
      c: "accent6",
    },
    {
      t: "DOY 194 · 09:28Z",
      title: "0.21 µrad Nominal Lock",
      desc: "Full 25 Mbps PPM-64 lock restored",
      c: "accent3",
    },
  ] as const;

  return (
    <EditorialSlide
      id="anomaly-postmortem"
      layout={anomalyBriefLayout}
      kicker="03 · FLIGHT CAMPAIGN & ROADMAP  |  ENGINEERING POSTMORTEM #04"
      title="Solar-Conjunction Thermal Bore-Sight Drift & Autonomous Fix"
      transition={transitions.fade({ duration: 450 })}
      notes="During 3-degree solar conjunction on DOY 194, unilateral solar illumination induced a 2.85 microradian bore-sight bias, permanently resolved within 5 hours via an onboard thermo-elastic Kalman lookup table."
    >
      {/* Top 4-Step Incident Timeline Ribbon */}
      <Box
        style={{
          display: "grid",
          gridTemplateColumns: "1fr 1fr 1fr 1fr",
          gap: 12,
          height: 98,
        }}
      >
        {events.map((ev) => (
          <Box
            key={ev.t}
            enter={zoom({ duration: 380 })}
            style={{
              minWidth: 0,
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [10, 14, 10, 14],
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
                gap: 8,
              }}
            >
              <Shape
                geometry={preset("ellipse")}
                style={{ width: 8, height: 8, background: ev.c }}
              />
              <Text
                style={{
                  flex: 1,
                  height: "auto",
                  fontSize: 10,
                  fontWeight: 700,
                  letterSpacing: 0.8,
                  color: ev.c,
                }}
              >
                <P>{ev.t}</P>
              </Text>
            </Box>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 13,
                fontWeight: 700,
                color: PALETTE.ink,
              }}
            >
              <P>{ev.title}</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 11,
                lineHeight: "15px",
                color: PALETTE.inkBody,
              }}
            >
              <P>{ev.desc}</P>
            </Text>
          </Box>
        ))}
      </Box>

      {/* Bottom: Left Area+Line Telemetry Chart (auto plotArea), Right Root Cause Card */}
      <Box style={{ flex: 1, display: "flex", flexDirection: "row", gap: 20 }}>
        <TelemetryPanel
          id="anomaly-chart-panel"
          title="DOY 194 Bore-Sight Pointing Bias Before & After LUT Patch (µrad)"
          badge="SEP = 3.0° PASS"
          accentColor={PALETTE.vermilion}
          flex={1.62}
        >
          <Chart
            id="anomaly-drift-chart"
            legend="top"
            axes={{
              utcCat: {
                kind: "category",
                position: "bottom",
                style: {
                  fontFamily: jostFont,
                  fontSize: 11,
                  color: PALETTE.inkBody,
                },
              },
              biasVal: {
                kind: "value",
                position: "left",
                min: 0,
                max: 3.5,
                majorUnit: 0.5,
                gridlines: { color: "#E7E4DB", width: 1 },
                style: {
                  fontFamily: jostFont,
                  fontSize: 11,
                  color: PALETTE.inkMuted,
                },
                title: {
                  text: "Bore-Sight Bias (µrad)",
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
                kind: "area",
                axes: { category: "utcCat", value: "biasVal" },
                data: {
                  kind: "shared-categories",
                  categories: [
                    "03:00Z",
                    "04:30Z",
                    "06:00Z",
                    "06:40Z",
                    "08:00Z",
                    "09:05Z",
                    "09:30Z",
                    "11:00Z",
                  ],
                },
                series: [
                  {
                    name: "Measured Bore-Sight Drift (µrad)",
                    values: [0.24, 0.95, 2.15, 2.85, 2.62, 1.4, 0.22, 0.2],
                    color: { color: PALETTE.vermilion, alpha: 0.45 },
                    line: { color: PALETTE.vermilion, width: 2.5 },
                  },
                ],
              },
              {
                kind: "line",
                axes: { category: "utcCat", value: "biasVal" },
                data: {
                  kind: "shared-categories",
                  categories: [
                    "03:00Z",
                    "04:30Z",
                    "06:00Z",
                    "06:40Z",
                    "08:00Z",
                    "09:05Z",
                    "09:30Z",
                    "11:00Z",
                  ],
                },
                series: [
                  {
                    name: "0.60 µrad CDR Spec Ceiling",
                    values: [0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
                    color: PALETTE.teal,
                    line: { color: PALETTE.teal, width: 2, dash: "dash" },
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

        <Box
          style={{
            flex: 1,
            flexBasis: 0,
            minWidth: 0,
            display: "flex",
            flexDirection: "column",
            justifyContent: "space-between",
            padding: [18, 20, 18, 20],
            background: PALETTE.obsidianCard,
            borderRadius: 12,
          }}
        >
          <Box style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.2,
                color: "#FF9B7A",
              }}
            >
              <P>ROOT CAUSE ANALYSIS</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 13,
                lineHeight: "19px",
                color: "#DCE6F7",
              }}
            >
              <P>
                At 3.0° Sun-Earth-Probe angle, unilateral solar illumination
                created a 14.2 K azimuthal gradient across the Ti-6Al-4V hexapod
                bipod mounts, tilting the star-tracker frame relative to the
                Cassegrain optical axis.
              </P>
            </Text>
          </Box>

          <Box
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr 1fr",
              gap: 8,
              padding: [10, 12, 10, 12],
              background: "#1A2742",
              border: { color: "#2D426B", width: 1, style: "solid" },
              borderRadius: 8,
            }}
          >
            {[
              { k: "RING GRADIENT", v: "+14.2 K" },
              { k: "PEAK DRIFT", v: "2.85 µrad" },
              { k: "PATCH LOCK", v: "23 min" },
            ].map((m) => (
              <Box
                key={m.k}
                style={{ display: "flex", flexDirection: "column", gap: 2 }}
              >
                <Text
                  style={{
                    width: "100%",
                    height: "auto",
                    fontSize: 9,
                    fontWeight: 700,
                    color: "#8CA2C6",
                  }}
                >
                  <P>{m.k}</P>
                </Text>
                <Text
                  style={{
                    width: "100%",
                    height: "auto",
                    fontSize: 14,
                    fontWeight: 700,
                    color: "#FFFFFF",
                  }}
                >
                  <P>{m.v}</P>
                </Text>
              </Box>
            ))}
          </Box>

          <Box style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.2,
                color: "#6AE3CB",
              }}
            >
              <P>PERMANENT BLOCK-II HARDWARE & FSW FIX</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontSize: 12,
                lineHeight: "18px",
                color: PALETTE.slateLight,
              }}
            >
              <P
                style={{
                  list: { type: "bullet", character: "•" },
                  spacingAfter: 5,
                }}
              >
                Added 8 PRT thermistors feeding real-time 12-state Kalman bias
                estimator
              </P>
              <P
                style={{
                  list: { type: "bullet", character: "•" },
                  spacingAfter: 5,
                }}
              >
                Replaced Ti bipods with low-CTE Invar-36 flexures + Kapton
                sun-shield
              </P>
              <P style={{ list: { type: "bullet", character: "•" } }}>
                Verified across 18 subsequent passes down to 2.8° SEP with zero
                drift
              </P>
            </Text>
          </Box>
        </Box>
      </Box>
    </EditorialSlide>
  );
}
