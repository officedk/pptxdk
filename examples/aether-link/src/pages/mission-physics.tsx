import {
  Box,
  Chart,
  Image,
  MotionScope,
  Mount,
  Notes,
  P,
  Shape,
  Slide,
  Span,
  Step,
  Text,
  fade,
  floatUp,
  linearGradient,
  preset,
  stretch,
  transitions,
  useMotionScope,
  wipe,
  zoom,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import {
  EditorialSlide,
  MetricTile,
  TakeawayStrip,
  TelemetryPanel,
} from "../components/chrome.tsx";
import { createBeamDivergenceDiagram } from "../components/vanilla.ts";
import {
  PALETTE,
  cassegrainApertureImage,
  deepSpaceBackgroundImage,
  displayFont,
  heroDarkLayout,
  jostFont,
  jostMediumFont,
  telemetryMorph,
} from "../theme.ts";

/** Slide 1: Cover dossier with full-bleed starfield background image and staggered KPI strip. */
export function CoverSlide() {
  const kpiScope = useMotionScope();
  const kpis = [
    {
      k: "PEAK DOWNLINK",
      v: "267 Mbps",
      s: "At 0.22 AU lunar-escape range",
      c: PALETTE.vermilion,
    },
    {
      k: "MAX VERIFIED RANGE",
      v: "2.20 AU",
      s: "329M km solar-conjunction lock",
      c: PALETTE.cobalt,
    },
    {
      k: "POINTING JITTER",
      v: "0.21 µrad",
      s: "Closed-loop 1.5 kHz piezo FSM",
      c: PALETTE.teal,
    },
    {
      k: "CUMULATIVE VOLUME",
      v: "41.8 Tb",
      s: "99.94% frame delivery post-FEC",
      c: PALETTE.amber,
    },
  ] as const;

  return (
    <Slide
      id="cover"
      layout={heroDarkLayout}
      background={{ image: deepSpaceBackgroundImage, fit: "cover" }}
      transition={transitions.fade({ duration: 600 })}
      style={{
        position: "relative",
        width: 1280,
        height: 720,
      }}
    >
      {/* Top accent hairline */}
      <Shape
        geometry={preset("rect")}
        style={{
          position: "absolute",
          left: 64,
          top: 48,
          width: 1152,
          height: 2,
          background: linearGradient(
            90,
            { color: PALETTE.vermilion, position: 0 },
            { color: PALETTE.cobalt, position: 0.55 },
            { color: "#1E2D4A", position: 1 },
          ),
        }}
      />

      {/* Left Hero Column */}
      <Box
        style={{
          position: "absolute",
          left: 64,
          top: 78,
          width: 680,
          height: 440,
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
        }}
      >
        <Box style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "center",
              gap: 10,
              width: 440,
              padding: [6, 14, 6, 14],
              background: PALETTE.obsidianRaised,
              border: {
                color: PALETTE.obsidianBorder,
                width: 1,
                style: "solid",
              },
              borderRadius: 20,
            }}
          >
            <Shape
              geometry={preset("ellipse")}
              style={{ width: 8, height: 8, background: PALETTE.vermilion }}
            />
            <Text
              style={{
                flex: 1,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 12,
                fontWeight: 700,
                letterSpacing: 1.2,
                color: PALETTE.ivoryText,
              }}
            >
              <P>
                CDR-4 MISSION REVIEW ·{" "}
                <Span style={{ color: "#FFB094" }}>
                  DEEP-SPACE OPTICAL RELAY
                </Span>
              </P>
            </Text>
          </Box>

          <Text
            id="cover-title"
            style={{
              width: "100%",
              height: "auto",
              fontFamily: displayFont,
              fontSize: 82,
              fontWeight: 600,
              lineHeight: "78px",
              color: PALETTE.ivoryText,
            }}
          >
            <P>AETHER-LINK</P>
          </Text>

          <Text
            style={{
              width: 640,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 22,
              fontWeight: 400,
              lineHeight: "31px",
              color: "#CED9EC",
            }}
          >
            <P>
              1550 nm Coherent Deep-Space Optical Telemetry, Sub-Microradian
              Beam Stabilization & Cislunar Trunk Architecture
            </P>
          </Text>

          <Text
            style={{
              width: 610,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 14,
              lineHeight: "21px",
              color: PALETTE.slateLight,
            }}
          >
            <P>
              Flight Campaign Synthesis (0.08 AU to 2.20 AU) · Jet Propulsion &
              Optical Ground Consortium · November 2026
            </P>
          </Text>
        </Box>

        {/* Interactive internal navigation CTA */}
        <Box
          style={{
            display: "flex",
            flexDirection: "row",
            alignItems: "center",
            gap: 16,
          }}
        >
          <Shape
            id="cover-cta"
            link={{ id: "exec-summary" }}
            geometry={preset("roundRect")}
            style={{
              width: 248,
              height: 42,
              background: PALETTE.vermilion,
              borderRadius: 8,
              fontFamily: jostFont,
              fontSize: 12,
              fontWeight: 700,
              letterSpacing: 1.3,
              color: "#FFFFFF",
              verticalAlign: "middle",
              boxShadow: {
                color: { color: "#000000", alpha: 0.35 },
                blur: 12,
                distance: 4,
                angle: 90,
              },
            }}
          >
            <P style={{ align: "center" }}>OPEN EXECUTIVE TELEMETRY</P>
          </Shape>
          <Text
            style={{
              width: 340,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 12,
              fontStyle: "italic",
              color: PALETTE.slateLight,
            }}
          >
            <P>10-slide engineering & flight verification dossier</P>
          </Text>
        </Box>
      </Box>

      {/* Right Cassegrain Optical Terminal Illustration */}
      <Box
        style={{
          position: "absolute",
          left: 792,
          top: 78,
          width: 424,
          height: 424,
          padding: 10,
          background: PALETTE.obsidianCard,
          border: {
            color: PALETTE.obsidianBorder,
            width: 1.5,
            style: "solid",
          },
          borderRadius: 20,
          boxShadow: {
            color: { color: "#000000", alpha: 0.45 },
            blur: 20,
            distance: 8,
            angle: 90,
          },
        }}
      >
        <Image
          id="cover-aperture-img"
          src={cassegrainApertureImage}
          alt="FLT-220 22 cm SiC Cassegrain Primary Aperture"
          fit="cover"
          geometry={preset("roundRect", { adj: 3500 })}
          style={{
            width: 404,
            height: 404,
          }}
        />
      </Box>

      {/* Bottom 4-column Telemetry Strip using Vanilla MetricTile components */}
      <MotionScope scope={kpiScope} mode="stagger" each={90}>
        <Box
          style={{
            position: "absolute",
            left: 64,
            top: 552,
            width: 1152,
            height: 118,
            display: "flex",
            flexDirection: "row",
            gap: 16,
          }}
        >
          {kpis.map((item, idx) => (
            <MetricTile
              key={item.k}
              id={`cover-kpi-${idx + 1}`}
              label={item.k}
              value={item.v}
              detail={item.s}
              accentColor={item.c}
              dark={true}
              compact={true}
              enter={fade({ duration: 360 })}
            />
          ))}
        </Box>
      </MotionScope>

      <Notes style={{ fontFamily: jostFont, fontSize: 12 }}>
        <P>
          Welcome to CDR-4 for AETHER-LINK. Over the past 14 months, the Flight
          Laser Transceiver demonstrated 10x to 40x higher downlink spectral
          efficiency than deep-space Ka-band RF.
        </P>
      </Notes>
    </Slide>
  );
}

/** Slide 2: Executive telemetry Bento summary with Zoom, Float, Fade, Wipe, and Morph pill. */
export function ExecSummarySlide() {
  return (
    <EditorialSlide
      id="exec-summary"
      kicker="01 · MISSION & PHYSICS  |  EXECUTIVE TELEMETRY SUMMARY"
      title="10.4x Throughput per Watt Over Deep-Space Ka-Band RF"
      transition={transitions.fade({ duration: 500 })}
      notes="Highlight the hero card on the left first, then click through the two progressive reveal steps covering power efficiency, link availability, and cislunar relay readiness."
    >
      {/* Asymmetric 3x2 Bento Grid */}
      <Box
        id="bento-grid"
        style={{
          flex: 1,
          display: "grid",
          gridTemplateColumns: "1.35fr 1fr 1fr",
          gridTemplateRows: "1fr 1fr",
          gap: 16,
        }}
      >
        {/* Left Hero Card spanning both rows */}
        <Box
          id="hero-bento"
          enter={zoom({ duration: 420 })}
          style={{
            gridColumn: "1",
            gridRow: "1 / 3",
            display: "flex",
            flexDirection: "column",
            justifyContent: "space-between",
            padding: [24, 26, 24, 26],
            background: linearGradient(
              150,
              { color: "#0E1628", position: 0 },
              { color: "#192949", position: 1 },
            ),
            borderRadius: 14,
            boxShadow: {
              color: { color: "#0B101B", alpha: 0.14 },
              blur: 16,
              distance: 6,
              angle: 90,
            },
          }}
        >
          <Box style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <Shape
              id="morph-laser-pill-s2"
              morphKey={telemetryMorph.handle("laser-carrier-pill")}
              geometry={preset("roundRect")}
              style={{
                width: 268,
                height: 30,
                background: PALETTE.vermilion,
                borderRadius: 15,
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.2,
                color: "#FFFFFF",
                verticalAlign: "middle",
              }}
            >
              <P style={{ align: "center" }}>1550.12 nm · 193.4 THz C-BAND</P>
            </Shape>

            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: displayFont,
                fontSize: 68,
                fontWeight: 600,
                lineHeight: "66px",
                color: PALETTE.ivoryText,
              }}
            >
              <P>
                267{" "}
                <Span
                  style={{
                    fontFamily: jostMediumFont,
                    fontSize: 24,
                    fontWeight: 500,
                    color: "#FF9B7A",
                  }}
                >
                  Mbps
                </Span>
              </P>
            </Text>

            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostMediumFont,
                fontSize: 16,
                fontWeight: 500,
                lineHeight: "23px",
                color: "#E1E9F8",
              }}
            >
              <P>
                Sustained error-free photon-counting downlink from 0.22 AU (33
                million km)
              </P>
            </Text>
          </Box>

          <Box
            style={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr 1fr",
              gap: 10,
              padding: [12, 14, 12, 14],
              background: "#13203A",
              border: { color: "#293E66", width: 1, style: "solid" },
              borderRadius: 10,
            }}
          >
            {[
              { k: "PHOTON EFFICIENCY", v: "2.18 b/ph" },
              { k: "TX OPTICS / DC", v: "4 W / 76 W" },
              { k: "SCHEDULED PASSES", v: "142 / 142" },
            ].map((m) => (
              <Box
                key={m.k}
                style={{ display: "flex", flexDirection: "column", gap: 3 }}
              >
                <Text
                  style={{
                    width: "100%",
                    height: "auto",
                    fontSize: 9,
                    fontWeight: 700,
                    letterSpacing: 0.8,
                    color: "#8CA2C6",
                  }}
                >
                  <P>{m.k}</P>
                </Text>
                <Text
                  style={{
                    width: "100%",
                    height: "auto",
                    fontSize: 15,
                    fontWeight: 700,
                    color: "#FFFFFF",
                  }}
                >
                  <P>{m.v}</P>
                </Text>
              </Box>
            ))}
          </Box>

          <Box
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 10,
              padding: [14, 16, 14, 16],
              background: "#111B30",
              border: { color: "#2B3E63", width: 1, style: "solid" },
              borderRadius: 10,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: "#8CA2C6",
              }}
            >
              <P>CAMPAIGN VERIFICATION BENCHMARKS</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: "20px",
                color: "#D3DDF0",
              }}
            >
              <P
                style={{
                  list: { type: "bullet", character: "•" },
                  spacingAfter: 4,
                }}
              >
                PPM-64 modulation at 2.18 bits/photon channel capacity
              </P>
              <P
                style={{
                  list: { type: "bullet", character: "•" },
                  spacingAfter: 4,
                }}
              >
                25 Mbps maintained at Mars opposition range (1.52 AU)
              </P>
              <P style={{ list: { type: "bullet", character: "•" } }}>
                Zero lock-drops across 142 scheduled ground passes
              </P>
            </Text>
          </Box>
        </Box>

        {/* Step 1 Click Reveal: Top two engineering efficiency cards */}
        <Step n={1}>
          <Box
            id="bento-card-1"
            enter={floatUp({ duration: 420 })}
            style={{
              gridColumn: "2",
              gridRow: "1",
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [18, 20, 18, 20],
              background: PALETTE.surface,
              border: { color: PALETTE.borderLight, width: 1, style: "solid" },
              borderRadius: 12,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: PALETTE.cobalt,
              }}
            >
              <P>FLIGHT TERMINAL SWaP</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: displayFont,
                fontSize: 48,
                fontWeight: 600,
                lineHeight: "48px",
                color: PALETTE.ink,
              }}
            >
              <P>
                28.4{" "}
                <Span
                  style={{
                    fontFamily: jostMediumFont,
                    fontSize: 18,
                    fontWeight: 500,
                    color: PALETTE.inkMuted,
                  }}
                >
                  kg / 76 W
                </Span>
              </P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: "19px",
                color: PALETTE.inkBody,
              }}
            >
              <P>
                42% lighter than an equivalent 3.2 m deployable Ka-band
                High-Gain Antenna assembly.
              </P>
            </Text>
          </Box>

          <Box
            id="bento-card-2"
            enter={floatUp({ duration: 420, delay: 90 })}
            style={{
              gridColumn: "3",
              gridRow: "1",
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [18, 20, 18, 20],
              background: PALETTE.surface,
              border: { color: PALETTE.borderLight, width: 1, style: "solid" },
              borderRadius: 12,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: PALETTE.teal,
              }}
            >
              <P>DETECTOR EFFICIENCY</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: displayFont,
                fontSize: 48,
                fontWeight: 600,
                lineHeight: "48px",
                color: PALETTE.ink,
              }}
            >
              <P>
                78.2%{" "}
                <Span
                  style={{
                    fontFamily: jostFont,
                    fontSize: 16,
                    fontWeight: 700,
                    color: PALETTE.teal,
                  }}
                >
                  SDE
                </Span>
              </P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: "19px",
                color: PALETTE.inkBody,
              }}
            >
              <P>
                32-pixel WSi Cryo-SNSPD nanowire array at 0.85 K with dark count
                below 120 cps.
              </P>
            </Text>
          </Box>
        </Step>

        {/* Step 2 Click Reveal: Bottom two operational readiness cards */}
        <Step n={2}>
          <Box
            id="bento-card-3"
            enter={fade({ duration: 420 })}
            style={{
              gridColumn: "2",
              gridRow: "2",
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [18, 20, 18, 20],
              background: PALETTE.surface,
              border: { color: PALETTE.borderLight, width: 1, style: "solid" },
              borderRadius: 12,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: PALETTE.amber,
              }}
            >
              <P>GLOBAL WEATHER DIVERSITY</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: displayFont,
                fontSize: 48,
                fontWeight: 600,
                lineHeight: "48px",
                color: PALETTE.ink,
              }}
            >
              <P>
                96.8%{" "}
                <Span
                  style={{
                    fontFamily: jostFont,
                    fontSize: 16,
                    fontWeight: 700,
                    color: PALETTE.amber,
                  }}
                >
                  SLA
                </Span>
              </P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: "19px",
                color: PALETTE.inkBody,
              }}
            >
              <P>
                4-continent optical ground station handover with 18-minute
                predictive cloud routing.
              </P>
            </Text>
          </Box>

          <Box
            id="bento-card-4"
            enter={wipe({ from: "left", duration: 420 })}
            style={{
              gridColumn: "3",
              gridRow: "2",
              display: "flex",
              flexDirection: "column",
              justifyContent: "space-between",
              padding: [18, 20, 18, 20],
              background: PALETTE.vermilionSoft,
              border: { color: "#E8B4A4", width: 1, style: "solid" },
              borderRadius: 12,
            }}
          >
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: PALETTE.vermilion,
              }}
            >
              <P>SUN-EARTH-PROBE ANGLE</P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: displayFont,
                fontSize: 48,
                fontWeight: 600,
                lineHeight: "48px",
                color: PALETTE.ink,
              }}
            >
              <P>
                3.0°{" "}
                <Span
                  style={{
                    fontFamily: jostFont,
                    fontSize: 16,
                    fontWeight: 700,
                    color: PALETTE.vermilion,
                  }}
                >
                  SEP MIN
                </Span>
              </P>
            </Text>
            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 13,
                lineHeight: "19px",
                color: PALETTE.inkBody,
              }}
            >
              <P>
                Narrowband 0.12 nm Fabry-Perot solar-rejection etalon sustained
                daytime lock.
              </P>
            </Text>
          </Box>
        </Step>
      </Box>

      <TakeawayStrip
        tag="CDR-4 GATE"
        accentColor={PALETTE.ink}
        text="All 14 Level-1 link, SWaP, and solar-conjunction thresholds met with >= 3.2 dB positive link margin."
      />
    </EditorialSlide>
  );
}

/** Slide 3: Beam divergence physics with slotted TelemetryPanel, direct Mount adoption, error bars, and stretch motion. */
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
        {/* Left: Slotted Vanilla TelemetryPanel wrapping Morph pill, Mount diagram, and ThemeScope callout */}
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
            theme={{ colors: { tileAccent: PALETTE.vermilion } }}
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
                color: "tileAccent",
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

        {/* Right: Slotted Vanilla TelemetryPanel wrapping Clustered Column Chart with Error Bars */}
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
