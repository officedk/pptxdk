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

export type ChartPoint = {
  readonly color?: ColorValue;
  /** Overrides data label text color for this data point. */
  readonly labelColor?: ColorValue;
};

export type ChartSeries = {
  readonly name: string;
  /** One numeric value per category position; null preserves a missing value. */
  readonly values: readonly (number | null)[];
  /** Color of column, area, line, or pie slice. */
  readonly color?: ColorValue;
  /** Overrides color or data label text color for individual points by position. */
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

export type ChartLine = {
  readonly color?: ColorValue;
  /** Line width in px. */
  readonly width?: number;
};

type ChartAxisBase = {
  /** Whether to show the axis and its labels; defaults to true. */
  readonly visible?: boolean;
  /** Major gridlines; true uses target application default line style. */
  readonly gridlines?: boolean | ChartLine;
  /** Axis line; false hides axis line. */
  readonly line?: boolean | ChartLine;
  /** Text style for axis labels. */
  readonly style?: ChartTextStyle;
};

export type ChartCategoryAxis = ChartAxisBase & {
  readonly kind: "category";
};

export type ChartValueAxis = ChartAxisBase & {
  readonly kind: "value";
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  /** Excel number format, e.g. `#,##0`. */
  readonly numberFormat?: string;
};

export type ChartAxis = ChartCategoryAxis | ChartValueAxis;

/** Chart-level semantic axes. Physical OOXML axis ids and cross references are internal. */
export type ChartAxes = Readonly<Record<string, ChartAxis>>;

/** The category and value axes used by one Cartesian group. */
export type ChartAxisReferences = {
  readonly category: string;
  readonly value: string;
};

type ChartDataLabelPosition =
  | "center"
  | "insideEnd"
  | "insideBase"
  | "outsideEnd"
  | "bestFit";

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

export type LineOrAreaChartGroup = CartesianGroupBase &
  CartesianStacking & {
    readonly kind: "line" | "area";
    readonly dataLabels?: DataLabelsWithoutPosition;
  };

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
