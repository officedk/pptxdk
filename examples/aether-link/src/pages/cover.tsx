import {
  Box,
  Image,
  MotionScope,
  Notes,
  P,
  Shape,
  Slide,
  Span,
  Text,
  fade,
  linearGradient,
  preset,
  transitions,
  useMotionScope,
} from "@officedk/pptxdk/react";

import { MetricTile } from "../components/chrome.tsx";
import {
  PALETTE,
  cassegrainApertureImage,
  deepSpaceBackgroundImage,
  displayFont,
  heroDarkLayout,
  jostFont,
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
