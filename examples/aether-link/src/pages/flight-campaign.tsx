import {
  Box,
  Chart,
  MotionScope,
  Mount,
  Notes,
  P,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Step,
  Text,
  fade,
  preset,
  spinIn,
  transitions,
  useMotionScope,
  zoom,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import {
  EditorialSlide,
  MetricTile,
  TakeawayStrip,
  TelemetryPanel,
} from "../components/chrome.tsx";
import {
  createBoardAnnotationSlot,
  createCampaignComboChart,
  createSubsystemMatrixTable,
  createWavefrontPicture,
} from "../components/vanilla.ts";
import {
  PALETTE,
  anomalyBriefLayout,
  closingDarkLayout,
  displayFont,
  editorialChartLayout,
  editorialTableLayout,
  editorialVisualLayout,
  jostFont,
  jostMediumFont,
} from "../theme.ts";

/** Slide 6: 14-month cruise telemetry using chart-placeholder, staggered MetricTiles, Step, and ThemeScope. */
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

      {/* Right: 3 Stacked Campaign KPI Tiles participating in Step, MotionScope stagger, and ThemeScope */}
      <Step n={1}>
        <MotionScope scope={kpiScope} mode="stagger" each={100}>
          <Box
            theme={{ colors: { tileAccent: PALETTE.teal } }}
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

/** Slide 8: Subsystem qualification matrix bound via native table-placeholder. */
export function SubsystemMatrixSlide() {
  const matrix = useMemo(() => createSubsystemMatrixTable(), []);

  return (
    <EditorialSlide
      id="subsystem-matrix"
      layout={editorialTableLayout}
      kicker="03 · FLIGHT CAMPAIGN & ROADMAP  |  VERIFICATION MATRIX"
      title="Subsystem Qualification & Flight Telemetry Compliance"
      transition={transitions.fade({ duration: 450 })}
      extraPlaceholders={{ matrix }}
      notes="This native PowerPoint table is bound into a table-placeholder slot and exercises vertical rowSpan grouping plus horizontal colSpan summary cells."
    >
      <Box
        style={{
          position: "absolute",
          left: 56,
          top: 604,
          width: 1168,
        }}
      >
        <TakeawayStrip
          tag="AUDIT RESULT"
          accentColor={PALETTE.cobalt}
          text="Zero open Class-A waivers; cislunar LRT-90 environmental thermal-vacuum qualification completed October 2026."
        />
      </Box>
    </EditorialSlide>
  );
}

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

/** Slide 10: Closing roadmap with slide-background-ref, unpopulated content-placeholder, external/home links, and SlideNumber. */
export function RoadmapCloseSlide() {
  const boardMemo = useMemo(() => createBoardAnnotationSlot(), []);

  const horizons = [
    {
      yr: "2026 · Q4",
      phase: "CISLUNAR OPS",
      rate: "10 Gbps Trunk",
      gate: "GATE: LRT-90 Flight Integration",
      accent: PALETTE.vermilion,
      items: [
        "Deploy LRT-90 relay in Lunar NRHO",
        "Continuous Artemis south-pole HD optical trunk",
        "CCSDS 141.0-B-1 interoperability lock",
      ],
    },
    {
      yr: "2027 · Q3",
      phase: "MARS OPPOSITION",
      rate: "250 Mbps @ 0.5 AU",
      gate: "GATE: 8 W Space-EDFA Qual",
      accent: PALETTE.cobalt,
      items: [
        "Dual-FLT coherent array on Mars Relay Orbiter",
        "8 W high-power space-qualified fiber amplifier",
        "Sub-second autonomous OGS cloud switchover",
      ],
    },
    {
      yr: "2028 · Q4",
      phase: "JOVIAN CRUISE",
      rate: "25 Mbps @ 5.2 AU",
      gate: "GATE: 50 cm Segmented SiC",
      accent: PALETTE.teal,
      items: [
        "50 cm deployable segmented SiC flight aperture",
        "8-meter synthesized ground photon-bucket array",
        "100x Galileo/Juno science return capacity",
      ],
    },
    {
      yr: "2030+",
      phase: "SOLAR BACKBONE",
      rate: "100 Gbps DWDM",
      gate: "GATE: L4/L5 Relay Ring",
      accent: PALETTE.amber,
      items: [
        "Lagrange L4/L5 solar-conjunction optical relay ring",
        "Eliminates Superior Conjunction blackout windows",
        "Open CCSDS optical DTN backbone across inner system",
      ],
    },
  ] as const;

  return (
    <Slide
      id="roadmap-close"
      layout={closingDarkLayout}
      background={{ ref: 1002, color: "dk2" }}
      transition={transitions.push({ from: "bottom", duration: 650 })}
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 20,
        padding: [44, 56, 38, 56],
      }}
    >
      <Mount node={boardMemo} />
      {/* Header Row */}
      <Box
        style={{
          display: "flex",
          flexDirection: "row",
          justifyContent: "space-between",
          alignItems: "flex-end",
        }}
      >
        <Box
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 6,
            width: 820,
          }}
        >
          <Text
            style={{
              width: "100%",
              height: "auto",
              fontFamily: jostFont,
              fontSize: 11,
              fontWeight: 700,
              letterSpacing: 1.8,
              color: "#FF9B7A",
            }}
          >
            <P>
              03 · HORIZON ROADMAP (2026–2030) · INTERPLANETARY OPTICAL BACKBONE
            </P>
          </Text>
          <Text
            style={{
              width: "100%",
              height: "auto",
              fontFamily: displayFont,
              fontSize: 44,
              fontWeight: 600,
              lineHeight: "46px",
              color: PALETTE.ivoryText,
            }}
          >
            <P>Scaling the Solar System Optical Backbone</P>
          </Text>
        </Box>

        <Box style={{ display: "flex", flexDirection: "row", gap: 12 }}>
          <Shape
            id="ccsds-spec-link"
            link={{ url: "https://public.ccsds.org" }}
            geometry={preset("roundRect")}
            style={{
              width: 158,
              height: 36,
              background: PALETTE.obsidianRaised,
              border: {
                color: PALETTE.obsidianBorder,
                width: 1,
                style: "solid",
              },
              borderRadius: 8,
              fontFamily: jostFont,
              fontSize: 11,
              fontWeight: 700,
              color: "#9CC0FA",
              verticalAlign: "middle",
            }}
          >
            <P style={{ align: "center" }}>CCSDS 142.0 SPEC</P>
          </Shape>
          <Shape
            id="return-home-btn"
            link={{ action: "home" }}
            geometry={preset("roundRect")}
            style={{
              width: 162,
              height: 36,
              background: PALETTE.vermilion,
              borderRadius: 8,
              fontFamily: jostFont,
              fontSize: 11,
              fontWeight: 700,
              color: "#FFFFFF",
              verticalAlign: "middle",
            }}
          >
            <P style={{ align: "center" }}>BACK TO COVER</P>
          </Shape>
        </Box>
      </Box>

      {/* 4 Equal-Width Horizon Columns via CSS Grid */}
      <Box
        style={{
          flex: 1,
          display: "grid",
          gridTemplateColumns: "1fr 1fr 1fr 1fr",
          gridTemplateRows: "1fr",
          gap: 16,
        }}
      >
        {horizons.map((h, i) => (
          <Box
            key={h.yr}
            id={`horizon-col-${i + 1}`}
            style={{
              minWidth: 0,
              display: "flex",
              flexDirection: "column",
              gap: 16,
              padding: [20, 20, 20, 20],
              background: PALETTE.obsidianCard,
              border: {
                color: PALETTE.obsidianBorder,
                width: 1,
                style: "solid",
              },
              borderRadius: 12,
            }}
          >
            <Box style={{ display: "flex", flexDirection: "column", gap: 11 }}>
              <Shape
                geometry={preset("rect")}
                style={{
                  width: "100%",
                  height: 4,
                  background: h.accent,
                  borderRadius: 2,
                }}
              />
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontFamily: jostFont,
                  fontSize: 12,
                  fontWeight: 700,
                  letterSpacing: 1.2,
                  color: h.accent,
                }}
              >
                <P>{h.yr}</P>
              </Text>
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontFamily: displayFont,
                  fontSize: 27,
                  fontWeight: 600,
                  lineHeight: "28px",
                  color: PALETTE.ivoryText,
                }}
              >
                <P>{h.phase}</P>
              </Text>
              <Shape
                geometry={preset("roundRect")}
                style={{
                  width: "100%",
                  height: 32,
                  background: PALETTE.obsidianRaised,
                  borderRadius: 6,
                  fontFamily: jostFont,
                  fontSize: 13,
                  fontWeight: 700,
                  color: "#FFFFFF",
                  verticalAlign: "middle",
                }}
              >
                <P style={{ align: "center" }}>{h.rate}</P>
              </Shape>
            </Box>

            <Box
              style={{
                padding: [8, 10, 8, 10],
                background: "#0E1525",
                border: { color: "#243454", width: 1, style: "solid" },
                borderRadius: 6,
              }}
            >
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontFamily: jostFont,
                  fontSize: 10,
                  fontWeight: 700,
                  letterSpacing: 0.6,
                  color: "#9CC0FA",
                }}
              >
                <P>{h.gate}</P>
              </Text>
            </Box>

            <Text
              style={{
                width: "100%",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 12,
                lineHeight: "18px",
                color: PALETTE.slateLight,
              }}
            >
              {h.items.map((it) => (
                <P
                  key={it}
                  style={{
                    list: { type: "bullet", character: "•" },
                    spacingAfter: 7,
                  }}
                >
                  {it}
                </P>
              ))}
            </Text>
          </Box>
        ))}
      </Box>

      {/* Bottom Sign-off Strip with Board Annotation Content-Placeholder Slot & SlideNumber */}
      <Box
        style={{
          height: 66,
          display: "flex",
          flexDirection: "row",
          justifyContent: "space-between",
          alignItems: "center",
          gap: 16,
          padding: [10, 18, 10, 18],
          background: PALETTE.obsidianCard,
          border: { color: PALETTE.obsidianBorder, width: 1, style: "solid" },
          borderRadius: 10,
        }}
      >
        <Text
          style={{
            width: 632,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 12,
            color: PALETTE.slateLight,
          }}
        >
          <P>
            <Span style={{ fontWeight: 700, color: PALETTE.ivoryText }}>
              CDR-4 BOARD RECOMMENDATION:{" "}
            </Span>
            Approve Block-II Production Release for Cislunar LRT-90 & Mars Relay
            Terminals.
          </P>
        </Text>

        {/* Visual frame around the unpopulated native content-placeholder (726, 604, 280, 46) */}
        <Shape
          geometry={preset("roundRect")}
          style={{
            width: 280,
            height: 46,
            background: "#0E1525",
            border: {
              color: PALETTE.obsidianBorder,
              width: 1,
              style: "dashed",
            },
            borderRadius: 6,
          }}
        />

        <Text
          style={{
            width: 148,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 11,
            fontWeight: 700,
            color: PALETTE.slateLight,
          }}
        >
          <P style={{ align: "right" }}>
            SLIDE <SlideNumber /> / 10
          </P>
        </Text>
      </Box>

      <Notes style={{ fontFamily: jostFont, fontSize: 12 }}>
        <P>
          Closing slide: ask the Review Board for formal authorization to
          proceed to Block-II manufacturing ahead of the Q4 2026 NRHO launch
          manifest.
        </P>
      </Notes>
    </Slide>
  );
}
