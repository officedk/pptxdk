import { z } from "zod";

/** Complete DrawingML preset geometry list from ECMA-376 ST_ShapeType (187 shapes). */
export const presetGeometryTuple = [
  "line",
  "lineInv",
  "triangle",
  "rtTriangle",
  "rect",
  "diamond",
  "parallelogram",
  "trapezoid",
  "nonIsoscelesTrapezoid",
  "pentagon",
  "hexagon",
  "heptagon",
  "octagon",
  "decagon",
  "dodecagon",
  "star4",
  "star5",
  "star6",
  "star7",
  "star8",
  "star10",
  "star12",
  "star16",
  "star24",
  "star32",
  "roundRect",
  "round1Rect",
  "round2SameRect",
  "round2DiagRect",
  "snipRoundRect",
  "snip1Rect",
  "snip2SameRect",
  "snip2DiagRect",
  "plaque",
  "ellipse",
  "teardrop",
  "homePlate",
  "chevron",
  "pieWedge",
  "pie",
  "blockArc",
  "donut",
  "noSmoking",
  "rightArrow",
  "leftArrow",
  "upArrow",
  "downArrow",
  "stripedRightArrow",
  "notchedRightArrow",
  "bentUpArrow",
  "leftRightArrow",
  "upDownArrow",
  "leftUpArrow",
  "leftRightUpArrow",
  "quadArrow",
  "leftArrowCallout",
  "rightArrowCallout",
  "upArrowCallout",
  "downArrowCallout",
  "leftRightArrowCallout",
  "upDownArrowCallout",
  "quadArrowCallout",
  "bentArrow",
  "uturnArrow",
  "circularArrow",
  "leftCircularArrow",
  "leftRightCircularArrow",
  "curvedRightArrow",
  "curvedLeftArrow",
  "curvedUpArrow",
  "curvedDownArrow",
  "swooshArrow",
  "cube",
  "can",
  "lightningBolt",
  "heart",
  "sun",
  "moon",
  "smileyFace",
  "irregularSeal1",
  "irregularSeal2",
  "foldedCorner",
  "bevel",
  "frame",
  "halfFrame",
  "corner",
  "diagStripe",
  "chord",
  "arc",
  "leftBracket",
  "rightBracket",
  "leftBrace",
  "rightBrace",
  "bracketPair",
  "bracePair",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
  "callout1",
  "callout2",
  "callout3",
  "accentCallout1",
  "accentCallout2",
  "accentCallout3",
  "borderCallout1",
  "borderCallout2",
  "borderCallout3",
  "accentBorderCallout1",
  "accentBorderCallout2",
  "accentBorderCallout3",
  "wedgeRectCallout",
  "wedgeRoundRectCallout",
  "wedgeEllipseCallout",
  "cloudCallout",
  "cloud",
  "ribbon",
  "ribbon2",
  "ellipseRibbon",
  "ellipseRibbon2",
  "leftRightRibbon",
  "verticalScroll",
  "horizontalScroll",
  "wave",
  "doubleWave",
  "plus",
  "flowChartProcess",
  "flowChartDecision",
  "flowChartInputOutput",
  "flowChartPredefinedProcess",
  "flowChartInternalStorage",
  "flowChartDocument",
  "flowChartMultidocument",
  "flowChartTerminator",
  "flowChartPreparation",
  "flowChartManualInput",
  "flowChartManualOperation",
  "flowChartConnector",
  "flowChartPunchedCard",
  "flowChartPunchedTape",
  "flowChartSummingJunction",
  "flowChartOr",
  "flowChartCollate",
  "flowChartSort",
  "flowChartExtract",
  "flowChartMerge",
  "flowChartOfflineStorage",
  "flowChartOnlineStorage",
  "flowChartMagneticTape",
  "flowChartMagneticDisk",
  "flowChartMagneticDrum",
  "flowChartDisplay",
  "flowChartDelay",
  "flowChartAlternateProcess",
  "flowChartOffpageConnector",
  "actionButtonBlank",
  "actionButtonHome",
  "actionButtonHelp",
  "actionButtonInformation",
  "actionButtonForwardNext",
  "actionButtonBackPrevious",
  "actionButtonEnd",
  "actionButtonBeginning",
  "actionButtonReturn",
  "actionButtonDocument",
  "actionButtonSound",
  "actionButtonMovie",
  "gear6",
  "gear9",
  "funnel",
  "mathPlus",
  "mathMinus",
  "mathMultiply",
  "mathDivide",
  "mathEqual",
  "mathNotEqual",
  "cornerTabs",
  "squareTabs",
  "plaqueTabs",
  "chartX",
  "chartStar",
  "chartPlus",
] as const;

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

export type PresetGeometryName = (typeof presetGeometryTuple)[number];

export type PresetGeometry = {
  readonly type: "preset";
  readonly name: PresetGeometryName;
  /** Raw DrawingML adjust values indexed by guide name (adj, adj1, adj2, etc.). */
  readonly adjust?: Readonly<Record<string, number>> | undefined;
};

export type PathCommand =
  | { readonly op: "M"; readonly x: number; readonly y: number }
  | { readonly op: "L"; readonly x: number; readonly y: number }
  | {
      readonly op: "C";
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly op: "Q";
      readonly x1: number;
      readonly y1: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: "Z" };

export type GeometryPath = {
  readonly commands: readonly PathCommand[];
  readonly fill?: boolean | undefined;
  readonly stroke?: boolean | undefined;
};

/** Custom geometry uses local coordinates and scales to the bounding box allocated by layout. */
export type CustomGeometry = {
  readonly type: "custom";
  readonly width: number;
  readonly height: number;
  readonly paths: readonly GeometryPath[];
};

export type Geometry = PresetGeometryName | PresetGeometry | CustomGeometry;

const rawPresetGeometryNameSchema = z.enum(presetGeometryTuple);
export const presetGeometryNameSchema: z.ZodType<PresetGeometryName> =
  rawPresetGeometryNameSchema;

export const geometryAdjustKeySchema: z.ZodType<string> = z
  .string()
  .regex(/^adj(?:[1-9]\d*)?$/u, "Geometry adjust key must match adj or adjN");

const rawGeometryAdjustSchema = z
  .record(
    geometryAdjustKeySchema,
    z
      .number()
      .refine(
        Number.isSafeInteger,
        "Geometry adjust must be an integer DrawingML parameter",
      ),
  )
  .readonly();
export const geometryAdjustSchema: z.ZodType<Readonly<Record<string, number>>> =
  rawGeometryAdjustSchema;

const rawPresetGeometrySchema = z
  .strictObject({
    type: z.literal("preset"),
    name: rawPresetGeometryNameSchema,
    /** Raw DrawingML adjust values indexed by guide name (adj, adj1, adj2, etc.). */
    adjust: rawGeometryAdjustSchema.optional(),
  })
  .readonly();
export const presetGeometrySchema: z.ZodType<PresetGeometry> =
  rawPresetGeometrySchema satisfies z.ZodType<
    AssertExact<
      PresetGeometry,
      z.infer<typeof rawPresetGeometrySchema>
    > extends true
      ? PresetGeometry
      : never
  >;

const finiteCoordSchema = z.number().finite();

const rawPathCommandSchema = z.discriminatedUnion("op", [
  z
    .strictObject({
      op: z.literal("M"),
      x: finiteCoordSchema,
      y: finiteCoordSchema,
    })
    .readonly(),
  z
    .strictObject({
      op: z.literal("L"),
      x: finiteCoordSchema,
      y: finiteCoordSchema,
    })
    .readonly(),
  z
    .strictObject({
      op: z.literal("C"),
      x1: finiteCoordSchema,
      y1: finiteCoordSchema,
      x2: finiteCoordSchema,
      y2: finiteCoordSchema,
      x: finiteCoordSchema,
      y: finiteCoordSchema,
    })
    .readonly(),
  z
    .strictObject({
      op: z.literal("Q"),
      x1: finiteCoordSchema,
      y1: finiteCoordSchema,
      x: finiteCoordSchema,
      y: finiteCoordSchema,
    })
    .readonly(),
  z.strictObject({ op: z.literal("Z") }).readonly(),
]);
export const pathCommandSchema: z.ZodType<PathCommand> =
  rawPathCommandSchema satisfies z.ZodType<
    AssertExact<PathCommand, z.infer<typeof rawPathCommandSchema>> extends true
      ? PathCommand
      : never
  >;

const rawGeometryPathSchema = z
  .strictObject({
    commands: z.array(rawPathCommandSchema).min(1).readonly(),
    fill: z.boolean().optional(),
    stroke: z.boolean().optional(),
  })
  .readonly();
export const geometryPathSchema: z.ZodType<GeometryPath> =
  rawGeometryPathSchema satisfies z.ZodType<
    AssertExact<
      GeometryPath,
      z.infer<typeof rawGeometryPathSchema>
    > extends true
      ? GeometryPath
      : never
  >;

const rawCustomGeometrySchema = z
  .strictObject({
    type: z.literal("custom"),
    width: z.number().finite().positive(),
    height: z.number().finite().positive(),
    paths: z.array(rawGeometryPathSchema).min(1).readonly(),
  })
  .readonly();
export const customGeometrySchema: z.ZodType<CustomGeometry> =
  rawCustomGeometrySchema satisfies z.ZodType<
    AssertExact<
      CustomGeometry,
      z.infer<typeof rawCustomGeometrySchema>
    > extends true
      ? CustomGeometry
      : never
  >;

const rawGeometrySchema = z.union([
  rawPresetGeometryNameSchema,
  rawPresetGeometrySchema,
  rawCustomGeometrySchema,
]);
export const geometrySchema: z.ZodType<Geometry> =
  rawGeometrySchema satisfies z.ZodType<
    AssertExact<Geometry, z.infer<typeof rawGeometrySchema>> extends true
      ? Geometry
      : never
  >;
