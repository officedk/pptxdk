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

export type PresetGeometryName = (typeof presetGeometryTuple)[number];

export const presetGeometryNameSchema = z.enum(presetGeometryTuple);

export const geometryAdjustKeySchema = z
  .string()
  .regex(/^adj(?:[1-9]\d*)?$/u, "Geometry adjust key must match adj or adjN");

export const geometryAdjustSchema = z
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

export const presetGeometrySchema = z
  .strictObject({
    type: z.literal("preset"),
    name: presetGeometryNameSchema,
    /** Raw DrawingML adjust values indexed by guide name (adj, adj1, adj2, etc.). */
    adjust: geometryAdjustSchema.optional(),
  })
  .readonly();
export type PresetGeometry = z.infer<typeof presetGeometrySchema>;

const finiteCoordSchema = z.number().finite();

export const pathCommandSchema = z.discriminatedUnion("op", [
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
export type PathCommand = z.infer<typeof pathCommandSchema>;

export const geometryPathSchema = z
  .strictObject({
    commands: z.array(pathCommandSchema).min(1).readonly(),
    fill: z.boolean().optional(),
    stroke: z.boolean().optional(),
  })
  .readonly();
export type GeometryPath = z.infer<typeof geometryPathSchema>;

/** Custom geometry uses local coordinates and scales to the bounding box allocated by layout. */
export const customGeometrySchema = z
  .strictObject({
    type: z.literal("custom"),
    width: z.number().finite().positive(),
    height: z.number().finite().positive(),
    paths: z.array(geometryPathSchema).min(1).readonly(),
  })
  .readonly();
export type CustomGeometry = z.infer<typeof customGeometrySchema>;

export const geometrySchema = z.union([
  presetGeometryNameSchema,
  presetGeometrySchema,
  customGeometrySchema,
]);
export type Geometry = z.infer<typeof geometrySchema>;
