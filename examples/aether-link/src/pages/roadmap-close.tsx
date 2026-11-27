import {
  Box,
  Mount,
  Notes,
  P,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Text,
  preset,
  transitions,
} from "@officedk/pptxdk/react";
import { useMemo } from "react";

import { createBoardAnnotationSlot } from "../components/diagrams.ts";
import { PALETTE, closingDarkLayout, displayFont, jostFont } from "../theme.ts";

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
