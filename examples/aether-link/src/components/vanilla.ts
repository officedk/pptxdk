import {
  box,
  cell,
  chart,
  group,
  image,
  paragraph,
  path,
  placeholder,
  preset,
  row,
  shape,
  span,
  table,
  text,
  type BoxNode,
  type ChartNode,
  type EntranceEffectDeclaration,
  type GroupNode,
  type ImageNode,
  type PlaceholderNode,
  type TableCellStyle,
  type TableNode,
} from "@officedk/pptxdk/vanilla";

import {
  PALETTE,
  displayFont,
  jostFont,
  jostMediumFont,
  wavefrontStrehlImage,
} from "../theme.ts";

/** Props for the closed KPI metric tile built with Vanilla primitives. */
export interface MetricTileProps {
  readonly id?: string | undefined;
  readonly label: string;
  readonly value: string;
  readonly unit?: string | undefined;
  readonly detail: string;
  readonly accentColor?: string | undefined;
  readonly dark?: boolean | undefined;
  readonly compact?: boolean | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Props for the closed bottom takeaway strip. */
export interface TakeawayStripProps {
  readonly id?: string | undefined;
  readonly tag: string;
  readonly text: string;
  readonly accentColor?: string | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Props for the parametric Ka-band vs. 1550 nm optical beam divergence diagram. */
export interface BeamDivergenceDiagramProps {
  readonly id?: string | undefined;
  readonly rfLabel: string;
  readonly rfDetail: string;
  readonly opticalLabel: string;
  readonly opticalDetail: string;
  readonly caption: string;
}

/** Props for the slotted telemetry panel container returning `{ node, slot }`. */
export interface TelemetryPanelProps {
  readonly id?: string | undefined;
  readonly kicker?: string | undefined;
  readonly title: string;
  readonly badge?: string | undefined;
  readonly accentColor?: string | undefined;
  readonly badgeBackground?: string | undefined;
  readonly badgeColor?: string | undefined;
  readonly dark?: boolean | undefined;
  readonly flex?: number | undefined;
  readonly width?: number | undefined;
  readonly enter?: EntranceEffectDeclaration | undefined;
}

/** Return shape for slotted Vanilla container components adopted by `fromVanilla`. */
export interface TelemetryPanelResult {
  readonly node: BoxNode;
  readonly slot: BoxNode;
}

/** Builds a closed KPI metric tile participating in parent Flex/Grid, MotionScope, and ThemeScope. */
export function createMetricTile(props: MetricTileProps): BoxNode {
  const accent = props.accentColor ?? "tileAccent";
  const bg = props.dark ? PALETTE.obsidianCard : PALETTE.surface;
  const border = props.dark ? PALETTE.obsidianBorder : PALETTE.borderLight;
  const valColor = props.dark ? PALETTE.ivoryText : PALETTE.ink;
  const labelColor = props.dark ? PALETTE.slateLight : PALETTE.inkMuted;
  const detailColor = props.dark ? PALETTE.slateLight : PALETTE.inkBody;

  const valueParagraph = paragraph(
    props.unit
      ? [
          props.value,
          span(` ${props.unit}`, {
            style: {
              fontFamily: jostMediumFont,
              fontSize: props.compact ? 14 : 16,
              fontWeight: 500,
              color: accent,
            },
          }),
        ]
      : props.value,
  );

  return box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      flex: 1,
      flexBasis: 0,
      minWidth: 0,
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      padding: props.compact ? [13, 16, 13, 16] : [16, 18, 16, 18],
      background: bg,
      border: { color: border, width: 1, style: "solid" },
      borderRadius: 10,
    },
  }).append(
    box({
      style: {
        display: "flex",
        flexDirection: "row",
        alignItems: "center",
        gap: 8,
      },
    }).append(
      shape({
        geometry: preset("rect"),
        style: { width: 4, height: 14, background: accent, borderRadius: 2 },
      }),
      text(props.label, {
        style: {
          flex: 1,
          height: "auto",
          fontFamily: jostFont,
          fontSize: props.compact ? 10 : 11,
          fontWeight: 700,
          letterSpacing: 1.1,
          color: labelColor,
        },
      }),
    ),
    text(valueParagraph, {
      style: {
        width: "100%",
        height: "auto",
        fontFamily: displayFont,
        fontSize: props.compact ? 35 : 42,
        fontWeight: 600,
        lineHeight: props.compact ? "36px" : "44px",
        color: valColor,
      },
    }),
    text(props.detail, {
      style: {
        width: "100%",
        height: "auto",
        fontFamily: jostFont,
        fontSize: props.compact ? 11 : 12,
        lineHeight: props.compact ? "15px" : "17px",
        color: detailColor,
      },
    }),
  );
}

/** Builds a closed full-width takeaway summary strip with a tagged pill and assertion copy. */
export function createTakeawayStrip(props: TakeawayStripProps): BoxNode {
  const accent = props.accentColor ?? PALETTE.cobalt;
  return box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      display: "flex",
      flexDirection: "row",
      alignItems: "center",
      gap: 14,
      padding: [10, 16, 10, 16],
      background: PALETTE.paperAlt,
      border: { color: PALETTE.borderLight, width: 1, style: "solid" },
      borderRadius: 8,
    },
  }).append(
    shape(
      paragraph(props.tag, {
        style: { align: "center" },
      }),
      {
        geometry: preset("roundRect"),
        style: {
          width: 124,
          height: 24,
          background: accent,
          borderRadius: 6,
          fontFamily: jostFont,
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 1.1,
          color: "#FFFFFF",
          verticalAlign: "middle",
        },
      },
    ),
    text(props.text, {
      style: {
        flex: 1,
        height: "auto",
        fontFamily: jostMediumFont,
        fontSize: 13,
        fontWeight: 500,
        lineHeight: "18px",
        color: PALETTE.ink,
      },
    }),
  );
}

/** Builds a parametric native DrawingML Group comparing RF Ka-band and 1550 nm beam divergence. */
export function createBeamDivergenceDiagram(
  props: BeamDivergenceDiagramProps,
): GroupNode {
  return group({
    ...(props.id === undefined ? {} : { id: props.id }),
    style: {
      position: "relative",
      width: 446,
      height: 278,
    },
  }).append(
    shape({
      geometry: preset("roundRect"),
      style: {
        position: "absolute",
        left: 0,
        top: 0,
        width: 446,
        height: 278,
        background: "#0E1626",
        borderRadius: 10,
      },
    }),
    shape({
      geometry: path("M 48 139 L 392 22 L 392 256 Z", {
        viewBox: [446, 278],
        fill: true,
        stroke: true,
      }),
      style: {
        position: "absolute",
        left: 0,
        top: 0,
        width: 446,
        height: 278,
        background: { color: PALETTE.cobalt, alpha: 0.22 },
        border: { color: "#4D88F0", width: 1.5, style: "dashed" },
      },
    }),
    shape({
      geometry: path("M 48 139 L 392 128 L 392 150 Z", {
        viewBox: [446, 278],
        fill: true,
        stroke: false,
      }),
      style: {
        position: "absolute",
        left: 0,
        top: 0,
        width: 446,
        height: 278,
        background: PALETTE.vermilion,
      },
    }),
    shape({
      geometry: preset("hexagon"),
      style: {
        position: "absolute",
        left: 26,
        top: 121,
        width: 36,
        height: 36,
        background: "#F4F2EC",
        border: { color: PALETTE.vermilion, width: 2, style: "solid" },
      },
    }),
    shape({
      geometry: preset("ellipse"),
      style: {
        position: "absolute",
        left: 382,
        top: 129,
        width: 20,
        height: 20,
        background: "#38B69A",
      },
    }),
    shape({
      geometry: preset("straightConnector1"),
      style: {
        position: "absolute",
        left: 414,
        top: 26,
        width: 1,
        height: 226,
        border: {
          color: "#7BA7F7",
          width: 1.5,
          style: "solid",
          startArrow: "triangle",
          endArrow: "triangle",
        },
      },
    }),
    text(
      [
        paragraph(props.rfLabel),
        paragraph(props.rfDetail, {
          style: { fontSize: 10, fontWeight: 400, color: "#B8CAE8" },
        }),
      ],
      {
        style: {
          position: "absolute",
          left: 222,
          top: 82,
          width: 168,
          height: "auto",
          fontFamily: jostFont,
          fontSize: 11,
          fontWeight: 700,
          color: "#98BEFF",
        },
      },
    ),
    text(
      [
        paragraph(props.opticalLabel),
        paragraph(props.opticalDetail, {
          style: { fontSize: 10, fontWeight: 400, color: "#E4D7D0" },
        }),
      ],
      {
        style: {
          position: "absolute",
          left: 222,
          top: 154,
          width: 168,
          height: "auto",
          fontFamily: jostFont,
          fontSize: 11,
          fontWeight: 700,
          color: "#FF9B7A",
        },
      },
    ),
    text(props.caption, {
      style: {
        position: "absolute",
        left: 18,
        top: 244,
        width: 260,
        height: "auto",
        fontFamily: jostFont,
        fontSize: 10,
        color: "#8CA2C4",
      },
    }),
  );
}

/**
 * Builds a slotted telemetry card shell with an imperative header bar, optional badge pill,
 * and an exposed flex `slot` container for JSX children (`<Chart />`, `<Mount />`, etc.).
 */
export function createTelemetryPanel(
  props: TelemetryPanelProps,
): TelemetryPanelResult {
  const accent = props.accentColor ?? PALETTE.vermilion;
  const bg = props.dark ? PALETTE.obsidianCard : PALETTE.surface;
  const border = props.dark ? PALETTE.obsidianBorder : PALETTE.borderLight;
  const titleColor = props.dark ? PALETTE.ivoryText : PALETTE.ink;

  const headerChildren = [
    shape({
      geometry: preset("rect"),
      style: { width: 4, height: 18, background: accent, borderRadius: 2 },
    }),
    text(
      props.kicker
        ? [
            paragraph(props.kicker, {
              style: {
                fontFamily: jostFont,
                fontSize: 9.5,
                fontWeight: 700,
                letterSpacing: 1.0,
                color: accent,
              },
            }),
            paragraph(props.title, {
              style: {
                fontFamily: jostFont,
                fontSize: 13,
                fontWeight: 700,
                color: titleColor,
              },
            }),
          ]
        : props.title,
      {
        style: {
          flex: 1,
          height: "auto",
          fontFamily: jostFont,
          fontSize: 13,
          fontWeight: 700,
          color: titleColor,
        },
      },
    ),
  ];

  if (props.badge !== undefined) {
    headerChildren.push(
      shape(paragraph(props.badge, { style: { align: "center" } }), {
        geometry: preset("roundRect"),
        style: {
          width: 146,
          height: 24,
          background: props.badgeBackground ?? PALETTE.paperAlt,
          border: { color: border, width: 1, style: "solid" },
          borderRadius: 6,
          fontFamily: jostFont,
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 0.6,
          color: props.badgeColor ?? PALETTE.inkBody,
          verticalAlign: "middle",
        },
      }),
    );
  }

  const slot = box({
    style: {
      flex: 1,
      minHeight: 0,
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      gap: 10,
    },
  });

  const node = box({
    ...(props.id === undefined ? {} : { id: props.id }),
    ...(props.enter === undefined ? {} : { enter: props.enter }),
    style: {
      ...(props.width === undefined
        ? { flex: props.flex ?? 1, flexBasis: 0, minWidth: 0 }
        : { width: props.width }),
      display: "flex",
      flexDirection: "column",
      gap: 12,
      padding: [16, 18, 16, 18],
      background: bg,
      border: { color: border, width: 1, style: "solid" },
      borderRadius: 12,
    },
  }).append(
    box({
      style: {
        display: "flex",
        flexDirection: "row",
        alignItems: "center",
        gap: 10,
      },
    }).append(...headerChildren),
    slot,
  );

  return { node, slot };
}

/** Imperative dual-axis ChartNode factory bound into Slide 6's native `chart-placeholder`. */
export function createCampaignComboChart(): ChartNode {
  return chart({
    id: "campaign-combo-chart",
    title: {
      text: "Bi-Monthly Science Volume (Tb) vs. Peak Link Rate (Mbps)",
      style: {
        fontFamily: jostFont,
        fontSize: 15,
        fontWeight: 700,
        color: PALETTE.ink,
      },
    },
    legend: "top",
    axes: {
      epochCat: {
        kind: "category",
        position: "bottom",
        style: {
          fontFamily: jostMediumFont,
          fontSize: 11,
          fontWeight: 500,
          color: PALETTE.inkBody,
        },
      },
      volLeft: {
        kind: "value",
        position: "left",
        min: 0,
        max: 10,
        majorUnit: 2,
        gridlines: { color: "#E7E4DB", width: 1 },
        style: { fontFamily: jostFont, fontSize: 11, color: PALETTE.cobalt },
        title: {
          text: "Volume per Window (Tb)",
          style: {
            fontFamily: jostFont,
            fontSize: 11,
            fontWeight: 700,
            color: PALETTE.cobalt,
          },
        },
      },
      rateRight: {
        kind: "value",
        position: "right",
        min: 0,
        max: 300,
        majorUnit: 60,
        gridlines: false,
        style: { fontFamily: jostFont, fontSize: 11, color: PALETTE.vermilion },
        title: {
          text: "Peak Link Rate (Mbps)",
          style: {
            fontFamily: jostFont,
            fontSize: 11,
            fontWeight: 700,
            color: PALETTE.vermilion,
          },
        },
      },
    },
    groups: [
      {
        kind: "column",
        axes: { category: "epochCat", value: "volLeft" },
        gapWidth: 85,
        data: {
          kind: "shared-categories",
          categories: [
            "E1 (0.12 AU)",
            "E2 (0.22 AU)",
            "E3 (0.48 AU)",
            "E4 (0.85 AU)",
            "E5 (1.24 AU)",
            "E6 (1.52 AU)",
            "E7 (1.88 AU)",
            "E8 (2.20 AU)",
          ],
        },
        dataLabels: {
          value: true,
          position: "insideEnd",
          style: {
            fontFamily: jostFont,
            fontSize: 10,
            fontWeight: 700,
            color: "#FFFFFF",
          },
        },
        series: [
          {
            name: "Delivered Science Volume (Tb)",
            values: [8.4, 9.1, 7.6, 5.8, 4.2, 3.1, 2.1, 1.5],
            color: PALETTE.cobalt,
          },
        ],
      },
      {
        kind: "line",
        axes: { category: "epochCat", value: "rateRight" },
        data: {
          kind: "shared-categories",
          categories: [
            "E1 (0.12 AU)",
            "E2 (0.22 AU)",
            "E3 (0.48 AU)",
            "E4 (0.85 AU)",
            "E5 (1.24 AU)",
            "E6 (1.52 AU)",
            "E7 (1.88 AU)",
            "E8 (2.20 AU)",
          ],
        },
        series: [
          {
            name: "Peak Downlink Rate (Mbps)",
            values: [267, 267, 124, 62, 36, 25, 16, 11.4],
            color: PALETTE.vermilion,
            smooth: true,
            line: { color: PALETTE.vermilion, width: 2.5 },
            marker: { symbol: "circle", size: 7 },
          },
        ],
      },
    ],
    style: {
      zIndex: 1,
      fontFamily: jostFont,
      fontSize: 11,
    },
  });
}

/** Imperative ImageNode factory bound into Slide 7's native `picture-placeholder`. */
export function createWavefrontPicture(): ImageNode {
  return image({
    id: "ao-wavefront-img",
    src: wavefrontStrehlImage,
    alt: "64-Channel Adaptive Optics Pupil Residual and Focal Plane SNSPD Coupling",
    fit: "cover",
    geometry: preset("roundRect", { adj: 2500 }),
    style: {
      zIndex: 1,
    },
  });
}

/** Imperative TableNode factory bound into Slide 8's native `table-placeholder`. */
export function createSubsystemMatrixTable(): TableNode {
  const rowsData = [
    {
      seg: "Flight Segment",
      rowSpan: 3,
      sub: "FLT-220 Optical Head (SiC)",
      spec: "<= 14.0 kg / 22 cm",
      meas: "12.8 kg / 22 cm",
      margin: "+8.6% mass",
      trl: "TRL-9",
      status: "VERIFIED",
    },
    {
      seg: "",
      rowSpan: 0,
      sub: "1550 nm HP-EDFA Laser Tx",
      spec: "4.0 W avg / 42 W DC",
      meas: "4.2 W avg / 39 W DC",
      margin: "+0.2 dB opt",
      trl: "TRL-9",
      status: "VERIFIED",
    },
    {
      seg: "",
      rowSpan: 0,
      sub: "Piezo FSM + Isolation Struts",
      spec: "<= 0.60 µrad RMS",
      meas: "0.21 µrad RMS",
      margin: "2.8x tighter",
      trl: "TRL-9",
      status: "VERIFIED",
    },
    {
      seg: "Cislunar Relay",
      rowSpan: 2,
      sub: "LRT-90 Dual-Gimbal Telescope",
      spec: "35 cm / 10 Gbps trunk",
      meas: "35 cm / 11.2 Gbps",
      margin: "+1.2 Gbps",
      trl: "TRL-7",
      status: "QUALIFIED",
    },
    {
      seg: "",
      rowSpan: 0,
      sub: "Rad-Hard DTN Bundle Store",
      spec: "32 TB / <= 18 W",
      meas: "48 TB / 15.4 W",
      margin: "+50% cap",
      trl: "TRL-8",
      status: "QUALIFIED",
    },
    {
      seg: "Ground Segment",
      rowSpan: 2,
      sub: "32-px WSi Cryo-SNSPD (0.85 K)",
      spec: ">= 70% SDE / <= 40 ps",
      meas: "78.2% SDE / 28 ps",
      margin: "+0.48 dB",
      trl: "TRL-9",
      status: "OPERATIONAL",
    },
    {
      seg: "",
      rowSpan: 0,
      sub: "64-Ch Daytime AO + Etalon",
      spec: "SEP >= 5.0° lock",
      meas: "SEP 3.0° lock",
      margin: "+2.0° sun",
      trl: "TRL-8",
      status: "OPERATIONAL",
    },
  ] as const;

  const headerCellStyle: TableCellStyle = {
    background: PALETTE.obsidianCard,
    color: PALETTE.ivoryText,
    fontFamily: jostFont,
    fontSize: 11,
    fontWeight: 700,
    letterSpacing: 0.8,
    verticalAlign: "middle",
    padding: [10, 12, 10, 12],
    borderBottom: { color: PALETTE.vermilion, width: 2, style: "solid" },
  };

  const headerRow = row({ style: { height: 40 } }).append(
    cell("SEGMENT", { style: headerCellStyle }),
    cell("SUBSYSTEM ASSEMBLY", { style: headerCellStyle }),
    cell(paragraph("CDR SPEC", { style: { align: "right" } }), {
      style: headerCellStyle,
    }),
    cell(paragraph("MEASURED", { style: { align: "right" } }), {
      style: headerCellStyle,
    }),
    cell(paragraph("MARGIN", { style: { align: "right" } }), {
      style: headerCellStyle,
    }),
    cell(paragraph("MATURITY", { style: { align: "center" } }), {
      style: headerCellStyle,
    }),
    cell(paragraph("GATE STATUS", { style: { align: "center" } }), {
      style: headerCellStyle,
    }),
  );

  const bodyRows = rowsData.map((r, i) => {
    const rowBg = i % 2 === 0 ? PALETTE.surface : "#F9F8F4";
    const cellBase: TableCellStyle = {
      background: rowBg,
      color: PALETTE.inkBody,
      fontFamily: jostFont,
      fontSize: 12,
      verticalAlign: "middle",
      padding: [9, 12, 9, 12],
      borderBottom: { color: "#E4E0D5", width: 1, style: "solid" },
    };
    const statusColor =
      r.status === "VERIFIED"
        ? PALETTE.teal
        : r.status === "OPERATIONAL"
          ? PALETTE.cobalt
          : PALETTE.amber;

    const cells = [
      ...(r.rowSpan > 0
        ? [
            cell(r.seg, {
              rowSpan: r.rowSpan,
              style: {
                ...cellBase,
                background: PALETTE.paperAlt,
                fontWeight: 700,
                color: PALETTE.ink,
                borderRight: {
                  color: PALETTE.borderLight,
                  width: 1,
                  style: "solid",
                },
              },
            }),
          ]
        : []),
      cell(r.sub, {
        style: {
          ...cellBase,
          fontFamily: jostMediumFont,
          fontWeight: 500,
          color: PALETTE.ink,
        },
      }),
      cell(paragraph(r.spec, { style: { align: "right" } }), {
        style: cellBase,
      }),
      cell(paragraph(r.meas, { style: { align: "right" } }), {
        style: { ...cellBase, fontWeight: 700, color: PALETTE.ink },
      }),
      cell(paragraph(r.margin, { style: { align: "right" } }), {
        style: { ...cellBase, fontWeight: 700, color: PALETTE.teal },
      }),
      cell(paragraph(r.trl, { style: { align: "center" } }), {
        style: cellBase,
      }),
      cell(paragraph(r.status, { style: { align: "center" } }), {
        style: { ...cellBase, fontWeight: 700, color: statusColor },
      }),
    ];

    return row({ style: { height: 50 } }).append(...cells);
  });

  const footerRow = row({ style: { height: 42 } }).append(
    cell("COMPOSITE SYSTEM LINK & SWaP COMPLIANCE (ALL 7 ASSEMBLIES)", {
      colSpan: 4,
      style: {
        background: PALETTE.obsidianRaised,
        color: PALETTE.ivoryText,
        fontFamily: jostFont,
        fontSize: 12,
        fontWeight: 700,
        verticalAlign: "middle",
        padding: [8, 12, 8, 12],
      },
    }),
    cell(
      paragraph("100% PASS  ·  +4.2 dB NET LINK MARGIN", {
        style: { align: "right" },
      }),
      {
        colSpan: 3,
        style: {
          background: PALETTE.obsidianRaised,
          color: "#7CE8D2",
          fontFamily: jostFont,
          fontSize: 12,
          fontWeight: 700,
          verticalAlign: "middle",
          padding: [8, 12, 8, 12],
        },
      },
    ),
  );

  return table({
    id: "verification-table",
    columns: [146, "1.35fr", 168, 172, 132, 88, 126],
    style: {
      fontFamily: jostFont,
      fontSize: 12,
      border: { color: PALETTE.borderLight, width: 1, style: "solid" },
    },
  }).append(headerRow, ...bodyRows, footerRow);
}

/** Imperative unpopulated PlaceholderNode bound into Slide 10's native `content-placeholder`. */
export function createBoardAnnotationSlot(): PlaceholderNode {
  return placeholder({
    id: "board-annotation-slot",
    placeholder: "boardMemo",
  });
}
