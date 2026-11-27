import {
  Box,
  MotionScope,
  P,
  Shape,
  Text,
  after,
  breathe,
  fade,
  fly,
  preset,
  repeat,
  spin,
  transitions,
  useMotionScope,
} from "@officedk/pptxdk/react";

import { EditorialSlide } from "../components/chrome.tsx";
import { PALETTE, displayFont, jostFont } from "../theme.ts";

/** Slide 4: Four-stage coherent optical pipeline with single stagger MotionScope and spin emphasis. */
export function ArchitectureSlide() {
  const nodesScope = useMotionScope();
  const relayScope = useMotionScope();

  const stages = [
    {
      code: "SEGMENT 01 · DEEP SPACE",
      name: "FLT-220 Transceiver",
      accent: PALETTE.vermilion,
      badge: "22 cm SiC  ·  4 W EDFA",
      iface: "TX: 1550 nm SCPPM-64  |  76 W DC",
      bullets: [
        "1550.12 nm master oscillator + Yb-Er fiber amplifier",
        "CCSDS 142.0-B-1 SCPPM encoder (rate 1/3, 1/2, 2/3)",
        "Piezo fine-steering mirror + point-ahead compensator",
      ],
    },
    {
      code: "SEGMENT 02 · CISLUNAR",
      name: "LRT-90 NRHO Relay",
      accent: PALETTE.cobalt,
      badge: "35 cm Dual-Gimbal",
      iface: "TRUNK: 10 Gbps DWDM + Ka-Band",
      bullets: [
        "Regenerative optical-to-Ka / 100G optical trunk switch",
        "48 TB rad-tolerant NVMe Delay-Tolerant Network store",
        "Bypasses terrestrial cloud cover on critical maneuvers",
      ],
    },
    {
      code: "SEGMENT 03 · GROUND",
      name: "5.1 m Cryo-SNSPD OGS",
      accent: PALETTE.teal,
      badge: "0.85 K He-4 Sorption",
      iface: "RX: 78.2% SDE  |  28 ps Jitter",
      bullets: [
        "64-channel adaptive optics mitigates daytime r0 blur",
        "32-element WSi nanowire array at 0.85 K cryo-bath",
        "Real-time FPGA iterative log-MAP turbo decoder",
      ],
    },
    {
      code: "SEGMENT 04 · MISSION OPS",
      name: "Autonomous Broker",
      accent: PALETTE.amber,
      badge: "18-min Cloud Forecast",
      iface: "SLA: 96.8% Weather Availability",
      bullets: [
        "GOES/Meteosat IR cloud-optical-depth ensemble ingest",
        "Hitless OGS site diversity handover across 4 longitudes",
        "Automated CCSDS CFDP science frame reassembly",
      ],
    },
  ] as const;

  return (
    <EditorialSlide
      id="architecture"
      kicker="02 · SYSTEM ARCHITECTURE  |  END-TO-END SIGNAL CHAIN"
      title="Four-Segment Coherent Optical & Cislunar Relay Pipeline"
      transition={transitions.strips({ direction: "leftDown", duration: 550 })}
      notes="Walk left-to-right across the four architecture segments. A single parent stagger MotionScope sequences all four cards cleanly at 140 ms intervals."
    >
      {/* Single Staggered 4-Stage Pipeline Row */}
      <MotionScope scope={nodesScope} mode="stagger" each={140}>
        <Box
          id="arch-pipeline"
          style={{
            flex: 1,
            display: "flex",
            flexDirection: "row",
            alignItems: "stretch",
            gap: 8,
          }}
        >
          {stages.map((st, idx) => (
            <Box
              key={st.code}
              style={{
                flex: 1,
                flexBasis: 0,
                minWidth: 0,
                display: "flex",
                flexDirection: "row",
                alignItems: "center",
                gap: 8,
              }}
            >
              <Box
                id={`arch-stage-${idx + 1}`}
                enter={fly({ from: "bottom", duration: 400 })}
                style={{
                  flex: 1,
                  flexBasis: 0,
                  minWidth: 0,
                  height: "100%",
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "space-between",
                  padding: [16, 16, 16, 16],
                  background: PALETTE.surface,
                  border: {
                    color: PALETTE.borderLight,
                    width: 1,
                    style: "solid",
                  },
                  borderRadius: 12,
                }}
              >
                <Box
                  style={{ display: "flex", flexDirection: "column", gap: 9 }}
                >
                  <Shape
                    geometry={preset("rect")}
                    style={{
                      width: "100%",
                      height: 4,
                      background: st.accent,
                      borderRadius: 2,
                    }}
                  />
                  <Text
                    style={{
                      width: "100%",
                      height: "auto",
                      fontFamily: jostFont,
                      fontSize: 10,
                      fontWeight: 700,
                      letterSpacing: 1.0,
                      color: st.accent,
                    }}
                  >
                    <P>{st.code}</P>
                  </Text>
                  <Text
                    style={{
                      width: "100%",
                      height: "auto",
                      fontFamily: displayFont,
                      fontSize: 25,
                      fontWeight: 600,
                      lineHeight: "26px",
                      color: PALETTE.ink,
                    }}
                  >
                    <P>{st.name}</P>
                  </Text>
                  <Shape
                    geometry={preset("roundRect")}
                    style={{
                      width: "100%",
                      height: 26,
                      background: PALETTE.paper,
                      border: {
                        color: PALETTE.borderLight,
                        width: 1,
                        style: "solid",
                      },
                      borderRadius: 6,
                      fontFamily: jostFont,
                      fontSize: 11,
                      fontWeight: 700,
                      color: PALETTE.inkBody,
                      verticalAlign: "middle",
                    }}
                  >
                    <P style={{ align: "center" }}>{st.badge}</P>
                  </Shape>
                </Box>

                <Box
                  style={{
                    padding: [8, 10, 8, 10],
                    background: PALETTE.paperAlt,
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
                      color: PALETTE.ink,
                    }}
                  >
                    <P>{st.iface}</P>
                  </Text>
                </Box>

                <Text
                  style={{
                    width: "100%",
                    height: "auto",
                    fontFamily: jostFont,
                    fontSize: 12,
                    lineHeight: "17px",
                    color: PALETTE.inkBody,
                  }}
                >
                  {st.bullets.map((b) => (
                    <P
                      key={b}
                      style={{
                        list: { type: "bullet", character: "•" },
                        spacingAfter: 6,
                      }}
                    >
                      {b}
                    </P>
                  ))}
                </Text>
              </Box>

              {idx < stages.length - 1 ? (
                <Shape
                  geometry={preset("rightArrow")}
                  style={{
                    width: 14,
                    height: 20,
                    flexShrink: 0,
                    background: PALETTE.borderStrong,
                  }}
                />
              ) : null}
            </Box>
          ))}
        </Box>
      </MotionScope>

      {/* Sequenced Lower Wavelength Multiplexing Bus Bar */}
      <MotionScope
        scope={relayScope}
        mode="parallel"
        startAt={after(nodesScope.end, 140)}
      >
        <Box
          id="wavelength-bus"
          enter={fade({ duration: 380 })}
          style={{
            display: "flex",
            flexDirection: "row",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 16,
            padding: [14, 20, 14, 20],
            background: PALETTE.obsidianCard,
            borderRadius: 12,
          }}
        >
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "center",
              gap: 10,
            }}
          >
            <Shape
              id="phase-lock-reticle"
              geometry={preset("star4")}
              emphasis={spin({ by: 360, duration: 1600 })}
              style={{
                width: 24,
                height: 24,
                background: "#FF9B7A",
              }}
            />
            <Shape
              id="active-beacon-pill"
              geometry={preset("roundRect")}
              emphasis={repeat(
                breathe({
                  from: PALETTE.vermilion,
                  to: "#FF8A65",
                  period: 1100,
                }),
                2,
              )}
              style={{
                width: 172,
                height: 34,
                background: PALETTE.vermilion,
                borderRadius: 8,
                fontFamily: jostFont,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 1.1,
                color: "#FFFFFF",
                verticalAlign: "middle",
              }}
            >
              <P style={{ align: "center" }}>OPTICAL DUPLEX PLAN</P>
            </Shape>
          </Box>

          {[
            {
              ch: "DOWNLINK SCIENCE",
              wl: "1550.12 nm  (LHC)",
              desc: "SCPPM-64 · 4 W avg / 320 W peak",
            },
            {
              ch: "UPLINK BEACON",
              wl: "1064.00 nm  (RHC)",
              desc: "5 kW CW ground reference + 256 kbps cmd",
            },
            {
              ch: "CISLUNAR TRUNK",
              wl: "1536.61 nm  (DWDM)",
              desc: "10 Gbps coherent DPSK lunar-Earth hop",
            },
          ].map((item) => (
            <Box
              key={item.ch}
              style={{
                flex: 1,
                display: "flex",
                flexDirection: "column",
                gap: 2,
              }}
            >
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontSize: 10,
                  fontWeight: 700,
                  letterSpacing: 1,
                  color: "#8CA2C6",
                }}
              >
                <P>{item.ch}</P>
              </Text>
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontSize: 13,
                  fontWeight: 700,
                  color: PALETTE.ivoryText,
                }}
              >
                <P>{item.wl}</P>
              </Text>
              <Text
                style={{
                  width: "100%",
                  height: "auto",
                  fontSize: 11,
                  color: PALETTE.slateLight,
                }}
              >
                <P>{item.desc}</P>
              </Text>
            </Box>
          ))}
        </Box>
      </MotionScope>
    </EditorialSlide>
  );
}
