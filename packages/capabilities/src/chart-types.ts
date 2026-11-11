import type {
  ColorValue,
  LayoutStyle,
  LinkTarget,
} from "@react-pptx/core/capability";

import type { MorphProps } from "./morph-api.ts";
import type { MotionProps } from "./motion-api.ts";
import type { TextStyle } from "./text-service.ts";

export type ChartTextStyle = Omit<TextStyle, "textWrap">;

export type ChartKind = "column" | "bar" | "line" | "area" | "pie" | "doughnut";

/** Native chart or axis title text with an optional text-style override. */
export type ChartTitle =
  | string
  | {
      readonly text: string;
      readonly style?: ChartTextStyle;
    };

/** Native preset stroke patterns for chart axes, gridlines, series, and points. */
export type ChartLineDash =
  | "solid"
  | "dash"
  | "dot"
  | "dashDot"
  | "longDash"
  | "longDashDot"
  | "longDashDotDot";

export type ChartLine = {
  readonly color?: ColorValue;
  /** Line width in px. */
  readonly width?: number;
  readonly dash?: ChartLineDash;
};

/** Native marker symbol; `picture` is excluded until picture markers have a resource contract. */
export type ChartMarkerSymbol =
  | "auto"
  | "circle"
  | "dash"
  | "diamond"
  | "dot"
  | "none"
  | "plus"
  | "square"
  | "star"
  | "triangle"
  | "x";

/** Marker formatting for a line series or one of its points. */
export type ChartMarker = {
  /** Marker symbol; defaults to circle when markers are enabled. */
  readonly symbol?: ChartMarkerSymbol;
  /** Marker size in points, from 2 to 72. */
  readonly size?: number;
  /** Marker fill; omitted leaves the native series or theme default. */
  readonly fill?: ColorValue;
  /** Marker outline; omitted leaves the native series or theme default. */
  readonly line?: ChartLine;
};

export type ChartPoint = {
  readonly color?: ColorValue;
  /** Stroke or border override for this data point. */
  readonly line?: ChartLine;
  /** Marker override for this point; applies only to line charts. */
  readonly marker?: boolean | ChartMarker;
  /** Overrides data label text color for this data point. */
  readonly labelColor?: ColorValue;
  /** Display, position, and number-format overrides for this point's data label. */
  readonly dataLabels?: false | ChartDataLabelOverride;
};

export type ChartSeries = {
  readonly name: string;
  /** One numeric value per category position; null preserves a missing value. */
  readonly values: readonly (number | null)[];
  /** Color of column, area, line, or pie slice. */
  readonly color?: ColorValue;
  /** Stroke or border override; for line charts it controls the series line. */
  readonly line?: ChartLine;
  /** Display, position, and number-format overrides for this series' data labels. */
  readonly dataLabels?: false | ChartDataLabelOverride;
  /** Marker settings; applies only to line charts and is hidden when omitted or false. */
  readonly marker?: boolean | ChartMarker;
  /** Whether the line is drawn with Catmull-Rom smoothing; applies only to line charts. */
  readonly smooth?: boolean;
  /** Overrides color, line, marker, or data label text color for points by position. */
  readonly points?: Readonly<Record<number, ChartPoint>>;
};

/**
 * Ordered categories shared by every series in one chart group.
 *
 * The discriminant leaves independent category domains to a separate future group-data shape.
 * Duplicate labels are meaningful positions, not keys.
 */
export type SharedCategoryData = {
  readonly kind: "shared-categories";
  readonly categories: readonly string[];
};

export type ChartGroupData = SharedCategoryData;

export type ChartTickMark = "none" | "in" | "out" | "cross";

export type ChartTickLabelPosition = "nextTo" | "high" | "low" | "none";

/** Where the crossing axis is placed relative to this axis. */
export type ChartAxisCrossing = "autoZero" | "max" | "min";

/** Chart-wide rendering of null category values; supported values depend on chart kind and stacking. */
export type ChartDisplayBlanksAs = "gap" | "zero" | "span";

type ChartAxisBase = {
  /** Whether to show the axis and its labels; defaults to true. */
  readonly visible?: boolean;
  /** Major gridlines; true uses target application default line style. */
  readonly gridlines?: boolean | ChartLine;
  /** Minor gridlines; true uses target application default line style. */
  readonly minorGridlines?: boolean | ChartLine;
  /** Axis line; false hides axis line. */
  readonly line?: boolean | ChartLine;
  /** Text style for axis labels. */
  readonly style?: ChartTextStyle;
  /** Native axis title; omitted declares no axis title. */
  readonly title?: ChartTitle;
  /** Major tick marks; defaults to none. */
  readonly majorTickMark?: ChartTickMark;
  /** Minor tick marks; defaults to none. */
  readonly minorTickMark?: ChartTickMark;
  /** Label position relative to the axis; defaults to nextTo. */
  readonly tickLabelPosition?: ChartTickLabelPosition;
  /** Reverses the axis direction with native maxMin orientation. */
  readonly reversed?: boolean;
  /** Places the crossing axis using a semantic position. */
  readonly crosses?: ChartAxisCrossing;
  /** Places the crossing axis at a value; category axes use a one-based category position. */
  readonly crossesAt?: number;
};

export type ChartCategoryAxis = ChartAxisBase & {
  readonly kind: "category";
};

export type ChartValueAxis = ChartAxisBase & {
  readonly kind: "value";
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  /** Excel number format, e.g. `#,##0`. */
  readonly numberFormat?: string;
  /** Logarithmic base; only value axes support a logarithmic scale. */
  readonly logBase?: number;
};

export type ChartAxis = ChartCategoryAxis | ChartValueAxis;

/** Chart-level semantic axes. Physical OOXML axis ids and cross references are internal. */
export type ChartAxes = Readonly<Record<string, ChartAxis>>;

/** The category and value axes used by one Cartesian group. */
export type ChartAxisReferences = {
  readonly category: string;
  readonly value: string;
};

export type ChartDataLabelPosition =
  | "center"
  | "insideEnd"
  | "insideBase"
  | "outsideEnd"
  | "bestFit"
  | "left"
  | "right"
  | "bottom"
  | "top";

/** Data label settings; group types narrow position by chart kind and stacking. */
export type ChartDataLabels<
  Position extends ChartDataLabelPosition = ChartDataLabelPosition,
> = {
  readonly value?: boolean;
  readonly percent?: boolean;
  readonly category?: boolean;
  readonly position?: Position;
  readonly numberFormat?: string;
  readonly style?: ChartTextStyle;
};

/**
 * Series- or point-level data-label overrides.
 *
 * Omitted display and position fields inherit from the enclosing group or series. An explicit
 * `false` suppresses labels at that level. Data-label text style remains a group-level field.
 */
export type ChartDataLabelOverride = Omit<ChartDataLabels, "style">;

type ChartGroupBase = {
  readonly data: ChartGroupData;
  readonly series: readonly ChartSeries[];
};

type CartesianGroupBase = ChartGroupBase & {
  /** Omit only when the chart uses its default category and value axes. */
  readonly axes?: ChartAxisReferences;
};

type DataLabelsWithoutPosition =
  | false
  | (Omit<ChartDataLabels, "position"> & { readonly position?: never });

type ColumnOrBarUnstackedOptions = {
  readonly stacking?: undefined;
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "insideBase" | "outsideEnd">;
};

type ColumnOrBarStackedOptions = {
  readonly stacking: "stacked" | "percent";
  readonly dataLabels?:
    | false
    | ChartDataLabels<"center" | "insideEnd" | "insideBase">;
};

type CartesianStacking = {
  readonly stacking?: "stacked" | "percent";
};

export type ColumnOrBarChartGroup = CartesianGroupBase &
  (
    | ({
        readonly kind: "column" | "bar";
        /** Gap between category groups as percentage of bar width, range 0–500. */
        readonly gapWidth?: number;
        /** Overlap percentage for bars in the same group, range -100–100. */
        readonly overlap?: number;
      } & ColumnOrBarUnstackedOptions)
    | ({
        readonly kind: "column" | "bar";
        readonly gapWidth?: number;
        readonly overlap?: number;
      } & ColumnOrBarStackedOptions)
  );

type LineDataLabels =
  | false
  | ChartDataLabels<"left" | "right" | "bottom" | "top" | "center">;

export type LineChartGroup = CartesianGroupBase &
  CartesianStacking & {
    readonly kind: "line";
    readonly dataLabels?: LineDataLabels;
  };

export type AreaChartGroup = CartesianGroupBase &
  CartesianStacking & {
    readonly kind: "area";
    readonly dataLabels?: DataLabelsWithoutPosition;
  };

export type LineOrAreaChartGroup = LineChartGroup | AreaChartGroup;

type PieOptions = {
  /**
   * Radial explosion percentage pulling the slice outward.
   * Rounded integer must be within the 0–2,147,483,647 range supported by current formal validation.
   */
  readonly explosion?: number;
  /** Clockwise starting angle of the first slice in degrees (0–360), serialized as integer. */
  readonly firstSliceAngle?: number;
};

export type PieChartGroup = ChartGroupBase &
  PieOptions & {
    readonly kind: "pie";
    readonly dataLabels?:
      | false
      | ChartDataLabels<"center" | "insideEnd" | "outsideEnd" | "bestFit">;
  };

export type DoughnutChartGroup = ChartGroupBase &
  PieOptions & {
    readonly kind: "doughnut";
    /** Percentage size of doughnut chart center hole, range 10–90. */
    readonly holeSize?: number;
    readonly dataLabels?: DataLabelsWithoutPosition;
  };

export type ChartGroup =
  | ColumnOrBarChartGroup
  | LineOrAreaChartGroup
  | PieChartGroup
  | DoughnutChartGroup;

export type ChartBase = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    /** Chart-level axes referenced by Cartesian groups; omitted uses the default category and value axes. */
    readonly axes?: ChartAxes;
    /** Groups are composed in declared order; each group owns its own series. */
    readonly groups: readonly ChartGroup[];
    /** Native chart title; omitted declares no chart title. */
    readonly title?: ChartTitle;
    /** Native chart-wide blank rendering; omitted leaves the chart-kind default. */
    readonly displayBlanksAs?: ChartDisplayBlanksAs;
    /** When omitted or false, hides the legend. */
    readonly legend?: false | "top" | "bottom" | "left" | "right";
    /** Internal plot area position relative to chart box, in px. */
    readonly plotArea?: {
      readonly left: number;
      readonly top: number;
      readonly width: number;
      readonly height: number;
    };
    /** Text style shared by chart legend, axes, and data labels. */
    readonly style?: ChartTextStyle & LayoutStyle;
  };

export type ChartProps = ChartBase;
