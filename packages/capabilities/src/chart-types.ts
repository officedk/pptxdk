import type {
  ColorValue,
  LayoutStyle,
  LinkTarget,
} from "@react-pptx/core/capability";

import type { MorphProps } from "./morph-api.ts";
import type { MotionProps } from "./motion-api.ts";
import type { TextStyle } from "./text-service.ts";

export type ChartTextStyle = Omit<TextStyle, "textWrap">;

export type ChartKind =
  | "column"
  | "bar"
  | "line"
  | "area"
  | "pie"
  | "doughnut"
  | "scatter";

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

/** Marker formatting for a line or scatter series, or one of their points. */
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
  /** Marker override for this point; applies only to line and scatter charts. */
  readonly marker?: boolean | ChartMarker;
  /**
   * Radial explosion percentage for this pie or doughnut slice.
   * Overrides the series and chart-group defaults for this point.
   */
  readonly explosion?: number;
  /** Overrides data label text color for this data point. */
  readonly labelColor?: ColorValue;
  /** Display, position, and number-format overrides for this point's data label. */
  readonly dataLabels?: false | ChartDataLabelOverride;
};

/** Native chart trendline type; `movingAverage` is the sixth Office trendline choice. */
export type ChartTrendlineType =
  | "linear"
  | "exponential"
  | "logarithmic"
  | "polynomial"
  | "power"
  | "movingAverage";

type ChartTrendlineBase = {
  /** Trendline stroke; omitted leaves the target application default. */
  readonly line?: ChartLine;
  /**
   * Forecast extension after the plotted data, in category positions or scatter X units.
   * Must be greater than 0; category charts require a multiple of 0.5.
   */
  readonly forward?: number;
  /**
   * Forecast extension before the plotted data, in category positions or scatter X units.
   * Must be greater than 0; category charts require a multiple of 0.5.
   */
  readonly backward?: number;
};

type ChartTrendlineDisplay = {
  /** Display the native trendline equation. */
  readonly displayEquation?: boolean;
  /** Display the native R-squared value. */
  readonly displayRSquared?: boolean;
};

/**
 * One native trendline owned by a series.
 *
 * Office supports multiple trendlines on one series, so this is a list on the series rather than a
 * single chart-level analysis option. `order` is required for polynomial, `period` for moving
 * average; `intercept` applies to exponential, linear, and polynomial only. A moving average has
 * no equation or R-squared label, so those fields are not part of its type. The applicable chart
 * kinds and value domains are enforced by the capability parse, not inferred from this union alone.
 */
export type ChartTrendline =
  | (ChartTrendlineBase &
      ChartTrendlineDisplay & {
        readonly type: "linear";
        /** Value where the trendline crosses the value axis; omitted fits it. */
        readonly intercept?: number;
        readonly order?: never;
        readonly period?: never;
      })
  | (ChartTrendlineBase &
      ChartTrendlineDisplay & {
        readonly type: "exponential";
        readonly intercept?: number;
        readonly order?: never;
        readonly period?: never;
      })
  | (ChartTrendlineBase &
      ChartTrendlineDisplay & {
        readonly type: "logarithmic";
        readonly intercept?: never;
        readonly order?: never;
        readonly period?: never;
      })
  | (ChartTrendlineBase &
      ChartTrendlineDisplay & {
        readonly type: "polynomial";
        /** Polynomial order required by Office for this type, an integer from 2 through 6. */
        readonly order: number;
        readonly intercept?: number;
        readonly period?: never;
      })
  | (ChartTrendlineBase &
      ChartTrendlineDisplay & {
        readonly type: "power";
        readonly intercept?: never;
        readonly order?: never;
        readonly period?: never;
      })
  | (ChartTrendlineBase & {
      readonly type: "movingAverage";
      /** Moving-average period, an integer from 2 through 255 and below the series point count. */
      readonly period: number;
      readonly intercept?: never;
      readonly order?: never;
      readonly displayEquation?: never;
      readonly displayRSquared?: never;
    });

type ChartSeriesFormat = {
  readonly name: string;
  /** Color of column, area, line, pie slice, or scatter markers. */
  readonly color?: ColorValue;
  /** Stroke or border override; for line and scatter charts it controls the series line. */
  readonly line?: ChartLine;
  /**
   * Radial explosion percentage for the pie or doughnut series.
   * Overrides the chart-group default; point values override this series default.
   */
  readonly explosion?: number;
  /** Display, position, number-format, leader-line, and hide overrides for this series. */
  readonly dataLabels?: false | ChartSeriesDataLabelOverride;
  /** Marker settings; applies only to line and scatter charts and is hidden when omitted or false. */
  readonly marker?: boolean | ChartMarker;
  /** Overrides color, line, marker, or data label text color for points by position. */
  readonly points?: Readonly<Record<number, ChartPoint>>;
  /** Native analysis trendlines owned by this series, in declaration order. */
  readonly trendlines?: readonly ChartTrendline[];
};

/** Category chart series: one value per ordered group category position. */
export type ChartSeries = ChartSeriesFormat & {
  /** One numeric value per category position; null preserves a missing value. */
  readonly values: readonly (number | null)[];
  /** Whether the line uses the native smoothing flag; applies only to line charts. */
  readonly smooth?: boolean;
};

type ScatterDataLabelContent = {
  /** Display the Y coordinate as the label value. */
  readonly value?: boolean;
  /** Display the X coordinate; `value` keeps its Y meaning when both are enabled. */
  readonly xValue?: boolean;
  /** Label position; omitted inherits the enclosing chart group or series. */
  readonly position?: "left" | "right" | "bottom" | "top" | "center";
  /** Excel number format applied to each displayed coordinate. */
  readonly numberFormat?: string;
};

/**
 * Scatter label settings where `value` is the Y coordinate and `xValue` is the X coordinate.
 *
 * Category and percent content are not applicable to scatter labels and are diagnosed when
 * declared. Position, number format, and the group text style follow the shared label contract.
 */
export type ScatterDataLabels = ScatterDataLabelContent & {
  /** Text style for scatter labels; series and points inherit it. */
  readonly style?: ChartTextStyle;
};

/** Series- and point-level scatter label overrides; text style remains a chart-group field. */
export type ScatterDataLabelOverride = ScatterDataLabelContent;

/** Scatter point formatting; pie and doughnut explosion does not apply. */
export type ScatterChartPoint = Omit<ChartPoint, "explosion" | "dataLabels"> & {
  /** Display, position, and number-format overrides for this point's data label. */
  readonly dataLabels?: false | ScatterDataLabelOverride;
};

/**
 * Scatter series with independent numeric X/Y observations.
 *
 * X and Y share one observation slot at each array position. Either coordinate may be null
 * without changing the slot, its point formatting, or the other coordinate.
 */
export type ScatterChartSeries = Omit<
  ChartSeriesFormat,
  "smooth" | "explosion" | "points" | "dataLabels"
> & {
  /** One numeric X coordinate per observation slot; null preserves a missing X. */
  readonly x: readonly (number | null)[];
  /** One numeric Y coordinate per observation slot; null preserves a missing Y. */
  readonly y: readonly (number | null)[];
  /** Display, position, and number-format overrides for this series' data labels. */
  readonly dataLabels?: false | ScatterDataLabelOverride;
  /** Overrides color, line, marker, or data label text color for points by position. */
  readonly points?: Readonly<Record<number, ScatterChartPoint>>;
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

/** Explicit XY data domain: every scatter series owns its numeric X/Y observations. */
export type ScatterData = {
  readonly kind: "xy";
};

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

/** The category and value axes used by one category Cartesian group. */
export type ChartCategoryAxisReferences = {
  readonly category: string;
  readonly value: string;
};

/** The two numeric axes used by one scatter group. */
export type ChartXYAxisReferences = {
  readonly x: string;
  readonly y: string;
};

export type ChartAxisReferences =
  | ChartCategoryAxisReferences
  | ChartXYAxisReferences;

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

/**
 * Native data-label settings; group types narrow position by chart kind and stacking.
 *
 * Leader lines are only written for pie and doughnut labels that display at least one of
 * value, percent, or category. `true` uses the target application's default leader-line
 * style; an object sets the native line color, width, and dash.
 */
export type ChartDataLabels<
  Position extends ChartDataLabelPosition = ChartDataLabelPosition,
> = {
  readonly value?: boolean;
  readonly percent?: boolean;
  readonly category?: boolean;
  readonly position?: Position;
  readonly numberFormat?: string;
  readonly leaderLines?: boolean | ChartLine;
  readonly style?: ChartTextStyle;
};

/**
 * Point-level data-label overrides.
 *
 * Omitted display and position fields inherit from the enclosing chart group or series. An explicit
 * `false` suppresses labels at that level. Data-label text style and leader lines remain
 * group or series fields because native point labels cannot carry leader-line formatting.
 */
export type ChartDataLabelOverride = Omit<
  ChartDataLabels,
  "style" | "leaderLines"
>;

/** Series-level data-label overrides; leader lines are valid on native series `dLbls`. */
export type ChartSeriesDataLabelOverride = Omit<ChartDataLabels, "style">;

/** Parsed label shape shared across chart kinds; xValue is populated only for scatter. */
export type ChartDataLabelSettings = ChartDataLabels & {
  readonly xValue?: boolean;
};

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

/** Connecting line between scatter observations; `false` draws markers only. */
export type ScatterLine = false | "straight" | "smooth";

/**
 * Scatter group with independent numeric X/Y observations.
 *
 * `line` and `markers` select the five native point-and-line presentations:
 * marker-only, straight or smooth lines with or without markers. Each series may still
 * format its own markers, points, and labels.
 */
export type ScatterChartGroup = {
  readonly kind: "scatter";
  readonly data: ScatterData;
  readonly series: readonly ScatterChartSeries[];
  /** Omit only when the chart uses its default X and Y value axes. */
  readonly axes?: ChartXYAxisReferences;
  /** Connecting line; omitted or false draws markers only. */
  readonly line?: ScatterLine;
  /** Whether markers are drawn; defaults to true. A group must draw markers or a line. */
  readonly markers?: boolean;
  readonly dataLabels?: false | ScatterDataLabels;
};

export type ChartGroup =
  | ColumnOrBarChartGroup
  | LineOrAreaChartGroup
  | PieChartGroup
  | DoughnutChartGroup
  | ScatterChartGroup;

export type ChartBase = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    /**
     * Chart-level axes referenced by groups; omitted uses default category and value axes for
     * category charts or default X and Y value axes for scatter charts.
     */
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
