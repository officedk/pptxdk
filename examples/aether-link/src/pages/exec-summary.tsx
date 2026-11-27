import {
  Box,
  P,
  Shape,
  Span,
  Step,
  Text,
  fade,
  floatUp,
  linearGradient,
  preset,
  transitions,
  wipe,
  zoom,
} from "@officedk/pptxdk/react";

import { EditorialSlide, TakeawayStrip } from "../components/chrome.tsx";
import {
  PALETTE,
  displayFont,
  jostFont,
  jostMediumFont,
  telemetryMorph,
} from "../theme.ts";

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
