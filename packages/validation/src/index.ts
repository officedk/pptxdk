import { posix } from "node:path";

import validateOpenXml from "@ooxml-tools/validate";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";

const presentationNamespace =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
const drawingNamespace =
  "http://schemas.openxmlformats.org/drawingml/2006/main";
const chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const spreadsheetNamespace =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const tableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const packageRelationshipNamespace =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const contentTypeNamespace =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const p159Namespace =
  "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
const markupCompatibilityNamespace =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
const transitionalStandard = "ECMA-376 5th edition Transitional";
const schemaValidatorPackage = "@ooxml-tools/validate";
const schemaValidatorPackageVersion = "0.4.6";
const openXmlSdkVersion = "3.0.2";
const schemaValidatorOfficeVersion = "Microsoft365";
const schemaValidatorProfile = {
  package: schemaValidatorPackage,
  packageVersion: schemaValidatorPackageVersion,
  sdkVersion: openXmlSdkVersion,
  officeVersion: schemaValidatorOfficeVersion,
} as const;
const validationScope =
  "Open XML SDK 3.0.2 under Microsoft365 profile Transitional package, schema, and constraint validation; local reader independently verifies OPC relationships, content types, and slide, shape, image, group, text, and transition structures, and verifies transition Choice/Fallback Requires and target extension constraints branch by branch. SDK MC selects Choice or Fallback solely according to its support; all compatibility branches and unknown extensions for p14/p159 are not guaranteed to be schema-validated. Validation does not equal full standard certification or application evidence.";
const applicationEvidence = "not-validated";
const minimumSlideLayoutId = 2_147_483_648;
const maximumSlideLayoutId = 4_294_967_295;
const maximumTimeNodeValue = 4_294_967_295;
const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;
const maximumShapeExtentEmu = 2_147_483_647;
const minimumFontSizeCentipoints = 100;
const maximumFontSizeCentipoints = 400_000;
const maximumTextSpacingCentipoints = 158_400;
// The external SDK correlates input using shared file IDs; serializing calls avoids concurrent overwrites, and the queue releases even on failure.
let schemaValidationQueue: Promise<void> = Promise.resolve();

export type ValidationIssue = {
  readonly part: string;
  readonly code: string;
  readonly message: string;
  readonly xpath?: string;
};

export type FileObject = {
  readonly id: number;
  readonly name: string;
  readonly kind: "shape" | "image" | "chart" | "group" | "table";
  readonly geometry?: string;
  readonly boxEmu: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  /** Actual a:xfrm@rot; indicates original angle of 0° when absent. */
  readonly rotationDegrees?: number;
  readonly image?: {
    readonly relationshipId: string;
    readonly part: string;
    readonly contentType: string;
  };
  readonly chart?: FileChart;
  readonly crop?: {
    readonly left?: number;
    readonly top?: number;
    readonly right?: number;
    readonly bottom?: number;
  };
  readonly link?:
    | { readonly slideId: string }
    | { readonly url: string }
    | { readonly action: "home" };
  readonly groupTransform?: {
    readonly childOffset: { readonly x: number; readonly y: number };
    readonly childExtent: { readonly width: number; readonly height: number };
    readonly flipH?: boolean;
    readonly flipV?: boolean;
  };
  readonly members?: readonly FileObject[];
  readonly text?: FileText;
  readonly table?: {
    readonly columnWidthsEmu: readonly number[];
    readonly rowHeightsEmu: readonly number[];
    readonly cells: readonly (readonly string[])[];
    readonly merges: readonly {
      readonly row: number;
      readonly column: number;
      readonly rowSpan: number;
      readonly colSpan: number;
    }[];
    readonly cellTexts: readonly {
      readonly row: number;
      readonly column: number;
      readonly text: FileText;
    }[];
  };
};

/** Native chart line formatting observed from c:spPr/a:ln. */
export type FileChartLine = {
  readonly color?: string;
  readonly widthEmu?: number;
  readonly dash?: string;
};

export type FileChartSeries = {
  readonly name: string;
  readonly categories: readonly string[];
  readonly values: readonly (number | null)[];
  /** Native line-series marker symbol and optional size. */
  readonly marker?: {
    readonly symbol: string;
    readonly size?: number;
  };
  /** Native line-series smoothing flag. */
  readonly smooth?: boolean;
  /** Native c:idx, the chart-wide series identity. */
  readonly index: number;
  /**
   * Native c:order. The formal writer derives it from the chart-wide series index,
   * not from the group-local position.
   */
  readonly order: number;
  /** Position of the owning chart group in c:plotArea order. */
  readonly groupIndex: number;
  /** Native c:explosion on this series, when present. */
  readonly explosion?: number;
  /** Native c:dPt/c:explosion values by point index. */
  readonly pointExplosions?: Readonly<Record<number, number>>;
};

/** Locally observed data-label settings; absent fields are not set by the XML node. */
export type FileChartDataLabel = {
  readonly value?: boolean;
  readonly percent?: boolean;
  readonly category?: boolean;
  readonly position?:
    | "center"
    | "insideEnd"
    | "insideBase"
    | "outsideEnd"
    | "bestFit"
    | "left"
    | "right"
    | "bottom"
    | "top";
  readonly numberFormat?: string;
  /** Native c:showLeaderLines; true or false only when explicitly present. */
  readonly leaderLines?: boolean;
  /** Native c:leaderLines/c:spPr/a:ln formatting, when specified. */
  readonly leaderLine?: FileChartLine;
};

/** Effective series labels and local point overrides reconstructed from chart XML. */
export type FileChartSeriesDataLabels = {
  readonly seriesIndex: number;
  /** Series-level settings merged with any chart-group dLbls base. `false` is c:delete. */
  readonly dataLabels?: false | FileChartDataLabel;
  /** Point-level c:dLbl overrides by native c:idx. `false` is c:delete. */
  readonly pointDataLabels?: Readonly<
    Record<number, false | FileChartDataLabel>
  >;
};

export type FileChartAxis = {
  readonly id: number;
  readonly kind: "category" | "value";
  readonly crossAxisId: number;
  readonly title?: string;
  readonly orientation?: "minMax" | "maxMin";
  readonly logBase?: number;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly minorUnit?: number;
  readonly numberFormat?: string;
  readonly majorTickMark?: string;
  readonly minorTickMark?: string;
  readonly tickLabelPosition?: string;
  readonly crosses?: string;
  readonly crossesAt?: number;
};

export type FileChartGroup = {
  readonly type: string;
  readonly direction?: "col" | "bar";
  readonly grouping?: string;
  /** c:axId references in native order; empty for pie and doughnut groups. */
  readonly axisIds: readonly number[];
  /** Native series ids in c:order order. */
  readonly seriesIndices: readonly number[];
  /** Native c:firstSliceAng for pie and doughnut groups. */
  readonly firstSliceAngle?: number;
  /** Native c:holeSize for doughnut groups. */
  readonly holeSize?: number;
};

export type FileChart = {
  readonly types: readonly string[];
  readonly direction?: "col" | "bar";
  readonly grouping?: string;
  readonly title?: string;
  readonly displayBlanksAs?: "gap" | "zero" | "span";
  /** Flattened series in chart-group order and native c:order order. */
  readonly series: readonly FileChartSeries[];
  readonly groups: readonly FileChartGroup[];
  readonly axes: readonly FileChartAxis[];
  /** Data-label observations in native series order; omitted when no labels are declared. */
  readonly seriesDataLabels?: readonly FileChartSeriesDataLabels[];
  readonly embeddedWorkbook?: {
    readonly part: string;
    readonly sheetName: string;
    readonly categories: readonly string[];
    readonly series: readonly FileChartSeries[];
  };
};

export type FileTextRun = {
  readonly text: string;
  readonly field?: "slidenum";
  readonly fontFamily?: string;
  readonly fontSizePt?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
  readonly lang?: string;
};

export type FileTextList =
  | { readonly type: "bullet"; readonly character: string }
  | {
      readonly type: "number";
      readonly style: string;
      readonly startAt?: number;
    };

export type FileTextParagraph = {
  readonly text: string;
  readonly runs: readonly FileTextRun[];
  readonly alignment?: string;
  readonly lineSpacingPoints?: number;
  readonly spacingAfterPoints?: number;
  readonly list?: FileTextList;
};

export type FileText = {
  readonly text: string;
  readonly paragraphs: readonly FileTextParagraph[];
};

export type FileNotes = {
  readonly part: string;
  readonly masterPart: string;
  readonly themePart?: string;
  readonly placeholder: {
    readonly type: "body";
    readonly index?: number;
  };
  readonly text: FileText;
};

export type FileSlide = {
  readonly id: string;
  readonly hidden?: boolean;
  readonly advanceAfterMs?: number;
  readonly section?: string;
  readonly background?: FileSlideBackground;
  readonly objects: readonly FileObject[];
  readonly transition?: FileTransition;
  readonly animations?: readonly FileAnimationStep[];
  readonly animationContainers?: readonly FileAnimationContainer[];
  readonly notes?: FileNotes;
};

export type FileAnimationContainer = {
  readonly timingId: number;
  readonly kind: "parallel" | "sequence";
  readonly parentTimingId?: number;
  readonly children: readonly {
    readonly kind: "container" | "effect";
    readonly timingId: number;
  }[];
  readonly durationMs?: number;
  readonly endSync?: "all";
  readonly startAfter?: FileAnimationDependency;
};

export type FileAnimationDependency = {
  readonly timingId: number;
  readonly point: "end";
  readonly offsetMs: number;
  readonly source: "effect" | "container";
  readonly targetId?: number;
};

export type FileSlideBackground =
  | { readonly type: "solid"; readonly color?: string }
  | {
      readonly type: "gradient";
      readonly kind: "linear" | "radial";
      readonly stops: readonly {
        readonly position: number;
        readonly color?: string;
      }[];
      readonly angle?: number;
    }
  | {
      readonly type: "image";
      readonly relationshipId: string;
      readonly part: string;
      readonly contentType: string;
    };

export type FileAnimationStep = {
  readonly trigger: "auto" | "click";
  readonly effects: readonly FileAnimationEffect[];
};

export type FileAnimationTargetRange =
  | { readonly type: "paragraph"; readonly start: number; readonly end: number }
  | {
      readonly type: "chart";
      readonly step: "gridLegend" | "series" | "category";
      readonly seriesIndex: number;
      readonly categoryIndex: number;
    };

export type FileAnimationKeyframe = {
  readonly time?: number;
  readonly formula?: string;
  readonly value?: string;
  readonly valueType?: string;
};

/** Actual file behavior node; `from`/`to` are endpoint views; keyframes store the full trajectory when formulas are present. */
export type FileAnimationBehavior = {
  readonly behavior: "set" | "anim" | "animEffect" | "animRot" | "animClr";
  readonly targetId: number;
  readonly targetType: "object" | "paragraph" | "background" | "chart";
  readonly targetRange?: FileAnimationTargetRange;
  readonly attributeNames: readonly string[];
  readonly durationMs?: number;
  readonly from?: string;
  readonly to?: string;
  /** Raw relative angle of animRot, in 1/60000ths of a degree per DrawingML. */
  readonly by?: string;
  readonly calculationMode?: string;
  /** Actual p:tav order in the file; formulas evaluate against numerical key values of adjacent nodes. */
  readonly keyframes?: readonly FileAnimationKeyframe[];
  readonly additive?: string;
  readonly filter?: string;
};

export type FileAnimationEffect = {
  readonly targetId: number;
  readonly timingId: number;
  readonly targetRange?: FileAnimationTargetRange;
  readonly nativeBuild?: "paragraph" | "allAtOnce" | "series" | "category";
  readonly nativeBuildGroupId?: number;
  readonly nativeBuildAnimBg?: boolean;
  readonly nativeBehaviors?: readonly FileAnimationBehavior[];
  readonly kind: "entrance" | "emphasis";
  readonly preset:
    | "fade"
    | "wipe"
    | "zoom"
    | "fly"
    | "floatUp"
    | "floatDown"
    | "spinIn"
    | "stretch"
    | "spin"
    | "breathe"
    | "fillColor";
  readonly durationMs: number;
  readonly delayMs: number;
  readonly repeatCount?: number | "forever";
  readonly by?: "word" | "letter" | number;
  readonly intervalMs?: number;
  readonly completion?: "unknown" | "no-natural-completion";
  readonly startAfter?: FileAnimationDependency;
  readonly from?: "left" | "right" | "top" | "bottom" | number;
  readonly fromColor?: string;
  readonly toColor?: string;
  readonly easing?: {
    readonly accelerate?: number;
    readonly decelerate?: number;
  };
};

const textProjectionParagraphRoles: Partial<
  Record<FileAnimationEffect["preset"], readonly string[]>
> = {
  fade: ["animEffect:fade"],
  wipe: ["wipe"],
  zoom: [
    "anim:ppt_x",
    "anim:ppt_y",
    "anim:ppt_w",
    "anim:ppt_h",
    "animEffect:fade",
  ],
  fly: ["anim:ppt_x", "anim:ppt_y"],
  floatUp: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
  floatDown: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
  spinIn: [
    "anim:ppt_w",
    "anim:ppt_h",
    "anim:style.rotation",
    "anim:ppt_x",
    "anim:ppt_y",
    "animEffect:fade",
  ],
  stretch: ["anim:ppt_x", "anim:ppt_w", "anim:ppt_h"],
  spin: ["animRot:r", "anim:ppt_x", "anim:ppt_y"],
};

const textProjectionBackgroundRoles: Partial<
  Record<FileAnimationEffect["preset"], readonly string[]>
> = {
  fade: ["animEffect:fade"],
  zoom: [
    "anim:ppt_x",
    "anim:ppt_y",
    "anim:ppt_w",
    "anim:ppt_h",
    "animEffect:fade",
  ],
  fly: ["anim:ppt_x", "anim:ppt_y"],
  floatUp: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
  floatDown: ["animEffect:fade", "anim:ppt_x", "anim:ppt_y"],
  spinIn: [
    "anim:ppt_x",
    "anim:ppt_y",
    "anim:ppt_w",
    "anim:ppt_h",
    "animRot:r",
    "animEffect:fade",
  ],
  stretch: ["anim:ppt_x", "anim:ppt_w", "anim:ppt_h"],
  spin: ["animRot:r", "anim:ppt_x", "anim:ppt_y"],
};

function animationBehaviorRole(behavior: FileAnimationBehavior): string {
  return behavior.behavior === "animEffect"
    ? `${behavior.behavior}:${behavior.filter ?? ""}`
    : `${behavior.behavior}:${behavior.attributeNames.join(",")}`;
}

function hasRoles(
  actual: readonly string[],
  expected: readonly string[],
): boolean {
  if (actual.length !== expected.length) return false;
  const normalize = (roles: readonly string[]) =>
    [...roles].sort((left, right) => left.localeCompare(right));
  const expectedRoles = normalize(expected);
  return normalize(actual).every(
    (role, index) => role === expectedRoles[index],
  );
}

function isPositionBehavior(behavior: FileAnimationBehavior): boolean {
  return (
    behavior.behavior === "anim" &&
    behavior.attributeNames.length === 1 &&
    (behavior.attributeNames[0] === "ppt_x" ||
      behavior.attributeNames[0] === "ppt_y")
  );
}

/** Scoped to the common-center rotation formula emitted in this round; does not perform general formula evaluation. */
type CenterRotationFormula = {
  readonly axis: "ppt_x" | "ppt_y";
  readonly centerX: number;
  readonly centerY: number;
  readonly angleDegrees: number;
  readonly progress: "increasing" | "decreasing";
  readonly scaleToCenter: boolean;
  readonly aspectRatio: number;
};

function closeGeometryNumber(left: number, right: number): boolean {
  return (
    Math.abs(left - right) <=
    1e-12 * Math.max(1, Math.abs(left), Math.abs(right))
  );
}

function parseCenterRotationFormula(
  behavior: FileAnimationBehavior,
): CenterRotationFormula | undefined {
  const formula = behavior.keyframes?.[0]?.formula;
  const axis = behavior.attributeNames[0];
  if (formula === undefined || (axis !== "ppt_x" && axis !== "ppt_y")) {
    return undefined;
  }
  const number = "(-?(?:\\d+(?:\\.\\d*)?|\\.\\d+))";
  const progress = "(?:\\$|\\(1-\\$\\))";
  const xPattern = new RegExp(
    `^${number}\\+(\\$\\*)?\\(\\(#ppt_x-${number}\\)\\*cos\\(\\(${number}\\*pi/180\\)\\*(?<firstProgress>${progress})\\)-\\(#ppt_y-${number}\\)\\*${number}\\*sin\\(\\(${number}\\*pi/180\\)\\*(?<secondProgress>${progress})\\)\\)-#ppt_x$`,
    "u",
  );
  const yPattern = new RegExp(
    `^${number}\\+(\\$\\*)?\\(\\(#ppt_x-${number}\\)\\*${number}\\*sin\\(\\(${number}\\*pi/180\\)\\*(?<firstProgress>${progress})\\)\\+\\(#ppt_y-${number}\\)\\*cos\\(\\(${number}\\*pi/180\\)\\*(?<secondProgress>${progress})\\)\\)-#ppt_y$`,
    "u",
  );
  const match = (axis === "ppt_x" ? xPattern : yPattern).exec(formula);
  if (match === null) return undefined;
  const progressMarker = match.groups?.firstProgress;
  if (
    progressMarker === undefined ||
    progressMarker !== match.groups?.secondProgress
  ) {
    return undefined;
  }
  const captures = match.slice(1).filter((value) => value !== progressMarker);
  const values = captures.map((value) =>
    value === undefined || value === "" || value === "$*"
      ? undefined
      : Number(value),
  );
  const scaleToCenter = captures[1] === "$*";
  let centerX: number | undefined;
  let centerY: number | undefined;
  let aspectRatio: number | undefined;
  let angleA: number | undefined;
  let angleB: number | undefined;
  if (axis === "ppt_x") {
    const [
      leadingCenter,
      ,
      repeatedCenterX,
      cosineAngle,
      repeatedCenterY,
      ratio,
      sineAngle,
    ] = values;
    if (
      leadingCenter === undefined ||
      repeatedCenterX === undefined ||
      cosineAngle === undefined ||
      repeatedCenterY === undefined ||
      ratio === undefined ||
      sineAngle === undefined ||
      !closeGeometryNumber(leadingCenter, repeatedCenterX)
    ) {
      return undefined;
    }
    centerX = leadingCenter;
    centerY = repeatedCenterY;
    aspectRatio = ratio;
    angleA = cosineAngle;
    angleB = sineAngle;
  } else {
    const [
      leadingCenterY,
      ,
      repeatedCenterX,
      ratio,
      sineAngle,
      repeatedCenterY,
      cosineAngle,
    ] = values;
    if (
      leadingCenterY === undefined ||
      repeatedCenterX === undefined ||
      ratio === undefined ||
      sineAngle === undefined ||
      repeatedCenterY === undefined ||
      cosineAngle === undefined ||
      !closeGeometryNumber(leadingCenterY, repeatedCenterY)
    ) {
      return undefined;
    }
    centerX = repeatedCenterX;
    centerY = leadingCenterY;
    aspectRatio = ratio;
    angleA = sineAngle;
    angleB = cosineAngle;
  }
  if (
    centerX === undefined ||
    centerY === undefined ||
    aspectRatio === undefined ||
    angleA === undefined ||
    angleB === undefined ||
    !closeGeometryNumber(angleA, angleB)
  ) {
    return undefined;
  }
  return {
    axis,
    centerX,
    centerY,
    angleDegrees: angleA,
    progress: progressMarker === "$" ? "increasing" : "decreasing",
    scaleToCenter,
    aspectRatio,
  };
}

/** Progress carrier for validating shared geometry; does not extrapolate formula positions at arbitrary points in time. */
function validCenterRotationProgress(behavior: FileAnimationBehavior): boolean {
  const keyframes = behavior.keyframes;
  return (
    isPositionBehavior(behavior) &&
    behavior.additive === "sum" &&
    behavior.calculationMode === "lin" &&
    keyframes?.length === 2 &&
    keyframes[0]?.time === 0 &&
    keyframes[1]?.time === 100000 &&
    keyframes[1]?.formula === undefined &&
    keyframes[0]?.valueType === "fltVal" &&
    keyframes[0]?.value === "0" &&
    keyframes[1]?.valueType === "fltVal" &&
    keyframes[1]?.value === "1" &&
    parseCenterRotationFormula(behavior) !== undefined
  );
}

function hasCompletePositionAxes(
  behaviors: readonly FileAnimationBehavior[],
): boolean {
  const axesByTarget = new Map<string, string[]>();
  for (const behavior of behaviors.filter(isPositionBehavior)) {
    const target = JSON.stringify([
      behavior.targetId,
      behavior.targetType,
      behavior.targetRange,
    ]);
    const axes = axesByTarget.get(target) ?? [];
    axes.push(behavior.attributeNames[0] ?? "");
    axesByTarget.set(target, axes);
  }
  return [...axesByTarget.values()].every((axes) =>
    hasRoles(axes, ["ppt_x", "ppt_y"]),
  );
}

function hasConsistentPositionFormulas(
  behaviors: readonly FileAnimationBehavior[],
): boolean {
  // Paragraphs, backgrounds, and members emitted within the same timing node must share parameter relationships.
  const formulas = behaviors
    .filter(isPositionBehavior)
    .map(parseCenterRotationFormula);
  const firstFormula = formulas[0];
  return (
    firstFormula !== undefined &&
    formulas.every(
      (formula) =>
        formula !== undefined &&
        closeGeometryNumber(formula.centerX, firstFormula.centerX) &&
        closeGeometryNumber(formula.centerY, firstFormula.centerY) &&
        closeGeometryNumber(formula.angleDegrees, firstFormula.angleDegrees) &&
        formula.scaleToCenter === firstFormula.scaleToCenter,
    )
  );
}

function sameAnimationTarget(
  left: FileAnimationBehavior,
  right: FileAnimationBehavior,
): boolean {
  return (
    left.targetId === right.targetId &&
    left.targetType === right.targetType &&
    JSON.stringify(left.targetRange) === JSON.stringify(right.targetRange)
  );
}

function oneBehaviorForTarget(
  behaviors: readonly FileAnimationBehavior[],
  target: FileAnimationBehavior,
  behaviorName: FileAnimationBehavior["behavior"],
  attributeName: string,
): FileAnimationBehavior | undefined {
  const matches = behaviors.filter(
    (behavior) =>
      sameAnimationTarget(behavior, target) &&
      behavior.behavior === behaviorName &&
      behavior.attributeNames.length === 1 &&
      behavior.attributeNames[0] === attributeName,
  );
  return matches.length === 1 ? matches[0] : undefined;
}

function numericEndpoints(
  behavior: FileAnimationBehavior | undefined,
): readonly [number, number] | undefined {
  if (behavior?.from === undefined || behavior.to === undefined) {
    return undefined;
  }
  const from = Number(behavior.from);
  const to = Number(behavior.to);
  return Number.isFinite(from) && Number.isFinite(to) ? [from, to] : undefined;
}

function hasLinearNumericEndpoints(
  behavior: FileAnimationBehavior,
  from: string,
  to: string,
): boolean {
  return (
    behavior.calculationMode === "lin" &&
    behavior.keyframes?.length === 2 &&
    behavior.keyframes[0]?.time === 0 &&
    behavior.keyframes[0]?.valueType === "fltVal" &&
    behavior.keyframes[0]?.value === from &&
    behavior.keyframes[1]?.time === 100000 &&
    behavior.keyframes[1]?.valueType === "fltVal" &&
    behavior.keyframes[1]?.value === to
  );
}

function hasSpinInScale(
  behaviors: readonly FileAnimationBehavior[],
  position: FileAnimationBehavior,
  axis: "ppt_w" | "ppt_h",
): boolean {
  const scale = oneBehaviorForTarget(behaviors, position, "anim", axis);
  return (
    scale !== undefined &&
    scale.additive === undefined &&
    scale.calculationMode === "lin" &&
    scale.from === "0" &&
    scale.to === `#${axis}` &&
    scale.durationMs === position.durationMs &&
    scale.keyframes?.length === 2 &&
    scale.keyframes[0]?.time === 0 &&
    scale.keyframes[0]?.valueType === "fltVal" &&
    scale.keyframes[0]?.value === "0" &&
    scale.keyframes[1]?.time === 100000 &&
    scale.keyframes[1]?.valueType === "strVal" &&
    scale.keyframes[1]?.value === `#${axis}`
  );
}

function spinAngleForPosition(
  preset: FileAnimationEffect["preset"],
  position: FileAnimationBehavior,
  behaviors: readonly FileAnimationBehavior[],
): number | undefined {
  if (preset === "spin") {
    const rotation = oneBehaviorForTarget(behaviors, position, "animRot", "r");
    if (
      rotation === undefined ||
      rotation.additive !== "sum" ||
      rotation.durationMs !== position.durationMs
    ) {
      return undefined;
    }
    if (rotation.by !== undefined) {
      const units = Number(rotation.by);
      return Number.isSafeInteger(units) &&
        units >= -2_147_483_648 &&
        units <= 2_147_483_647 &&
        rotation.from === undefined &&
        rotation.to === undefined
        ? units / 60000
        : undefined;
    }
    const endpoints = numericEndpoints(rotation);
    if (endpoints === undefined) return undefined;
    return (endpoints[0] - endpoints[1]) / 60000;
  }
  if (preset !== "spinIn") return undefined;
  if (
    !hasSpinInScale(behaviors, position, "ppt_w") ||
    !hasSpinInScale(behaviors, position, "ppt_h")
  ) {
    return undefined;
  }
  if (position.targetType === "background") {
    const rotation = oneBehaviorForTarget(behaviors, position, "animRot", "r");
    const endpoints = numericEndpoints(rotation);
    if (
      rotation === undefined ||
      rotation.additive !== "sum" ||
      endpoints === undefined ||
      rotation.durationMs !== position.durationMs
    ) {
      return undefined;
    }
    return (endpoints[0] - endpoints[1]) / 60000;
  }
  const rotation = oneBehaviorForTarget(
    behaviors,
    position,
    "anim",
    "style.rotation",
  );
  const endpoints = numericEndpoints(rotation);
  if (
    rotation === undefined ||
    rotation.additive !== undefined ||
    endpoints === undefined ||
    !closeGeometryNumber(endpoints[0], 360) ||
    !closeGeometryNumber(endpoints[1], 0) ||
    !hasLinearNumericEndpoints(rotation, "360", "0") ||
    rotation.durationMs !== position.durationMs
  ) {
    return undefined;
  }
  return endpoints[0] - endpoints[1];
}

/** Checks actual position formulas against slide dimensions, object geometry, and corresponding native rotation/scale parameters. */
function validCenterRotationGeometry(
  preset: FileAnimationEffect["preset"],
  behaviors: readonly FileAnimationBehavior[],
  objects: readonly FileObject[],
  widthEmu: number,
  heightEmu: number,
): boolean {
  const positions = behaviors.filter(isPositionBehavior);
  if (positions.length === 0) return true;
  const parsed = positions.map(parseCenterRotationFormula);
  const firstFormula = parsed[0];
  if (
    firstFormula === undefined ||
    (preset !== "spin" && preset !== "spinIn") ||
    firstFormula.scaleToCenter !== (preset === "spinIn")
  ) {
    return false;
  }
  const expectedRatios = {
    ppt_x: heightEmu / widthEmu,
    ppt_y: widthEmu / heightEmu,
  } as const;
  return positions.every((position, index) => {
    const formula = parsed[index];
    if (
      formula === undefined ||
      formula.scaleToCenter !== (preset === "spinIn") ||
      formula.progress !==
        (preset === "spin" &&
        oneBehaviorForTarget(behaviors, position, "animRot", "r")?.by !==
          undefined
          ? "increasing"
          : "decreasing") ||
      !closeGeometryNumber(formula.centerX, firstFormula.centerX) ||
      !closeGeometryNumber(formula.centerY, firstFormula.centerY) ||
      !closeGeometryNumber(formula.angleDegrees, firstFormula.angleDegrees) ||
      !closeGeometryNumber(formula.aspectRatio, expectedRatios[formula.axis])
    ) {
      return false;
    }
    const geometry = motionTargetGeometry(
      objects,
      position.targetId,
      widthEmu,
      heightEmu,
    );
    if (geometry === undefined || !geometry.supported) return false;
    const centers =
      preset === "spin"
        ? [geometry.objectCenter, ...geometry.containingGroupCenters]
        : [geometry.objectCenter];
    if (
      !centers.some(
        (center) =>
          closeGeometryNumber(formula.centerX, center.x) &&
          closeGeometryNumber(formula.centerY, center.y),
      )
    ) {
      return false;
    }
    const angle = spinAngleForPosition(preset, position, behaviors);
    return (
      angle !== undefined && closeGeometryNumber(formula.angleDegrees, angle)
    );
  });
}

/** Recovers logical whole Text only after complete verification against actual paragraphs, background, and visibility targets. */
function validWholeTextProjection(
  preset: FileAnimationEffect["preset"],
  kind: FileAnimationEffect["kind"],
  paragraphCount: number,
  build: string,
  animBg: boolean,
  behaviors: readonly FileAnimationBehavior[],
): boolean {
  const expectedParagraphRoles = textProjectionParagraphRoles[preset];
  if (expectedParagraphRoles === undefined) return false;

  const visibilitySets = behaviors.filter(({ behavior }) => behavior === "set");
  const visibilityRoles = visibilitySets.map(({ targetType, targetRange }) =>
    targetType === "paragraph" &&
    targetRange?.type === "paragraph" &&
    targetRange.start === targetRange.end
      ? `paragraph:${targetRange.start}`
      : targetType === "background" && targetRange === undefined
        ? "background"
        : "invalid",
  );
  const expectedVisibilityRoles =
    kind === "entrance"
      ? [
          ...Array.from(
            { length: paragraphCount },
            (_, index) => `paragraph:${index}`,
          ),
          ...(animBg ? ["background"] : []),
        ]
      : [];
  if (
    behaviors.some(({ targetType }) => targetType === "object") ||
    !hasRoles(visibilityRoles, expectedVisibilityRoles) ||
    visibilitySets.some(
      ({ attributeNames, durationMs, to }) =>
        attributeNames.length !== 1 ||
        attributeNames[0] !== "style.visibility" ||
        durationMs !== 1 ||
        to !== "visible",
    ) ||
    (kind === "entrance" && build !== "p")
  ) {
    return false;
  }

  for (let index = 0; index < paragraphCount; index += 1) {
    const paragraphBehaviors = behaviors.filter(
      ({ behavior, targetType, targetRange }) =>
        behavior !== "set" &&
        targetType === "paragraph" &&
        targetRange?.type === "paragraph" &&
        targetRange.start === index &&
        targetRange.end === index,
    );
    const actualRoles = paragraphBehaviors.map(animationBehaviorRole);
    if (preset === "wipe") {
      if (
        actualRoles.length !== 1 ||
        !actualRoles[0]?.startsWith("animEffect:wipe(")
      ) {
        return false;
      }
    } else if (!hasRoles(actualRoles, expectedParagraphRoles)) {
      return false;
    }
  }

  const backgroundBehaviors = behaviors.filter(
    ({ behavior, targetType }) =>
      behavior !== "set" && targetType === "background",
  );
  if (backgroundBehaviors.length === 0) return !animBg;
  if (preset === "wipe") {
    const paragraphFilter = behaviors.find(
      ({ behavior, targetType }) =>
        behavior === "animEffect" && targetType === "paragraph",
    )?.filter;
    return (
      animBg &&
      paragraphFilter !== undefined &&
      backgroundBehaviors.every(
        ({ targetRange, filter }) =>
          targetRange === undefined && filter === paragraphFilter,
      ) &&
      hasRoles(backgroundBehaviors.map(animationBehaviorRole), [
        `animEffect:${paragraphFilter}`,
      ])
    );
  }
  const expectedBackgroundRoles = textProjectionBackgroundRoles[preset];
  return (
    animBg &&
    expectedBackgroundRoles !== undefined &&
    backgroundBehaviors.every(({ targetRange }) => targetRange === undefined) &&
    hasRoles(
      backgroundBehaviors.map(animationBehaviorRole),
      expectedBackgroundRoles,
    )
  );
}

export type FileTransition = {
  readonly kind: "fade" | "strips" | "push" | "flip" | "morph";
  readonly speed?: "fast" | "med" | "slow";
  readonly direction?: string;
  readonly durationMs?: number;
  readonly by?: "object" | "word" | "character";
  readonly fallback?: FileTransition;
};

export type ValidationExtension = "p14" | "p159";

export type ValidationOptions = {
  /** Rejects target extensions used in the actual file that are not permitted; performs structural reading only when omitted. */
  readonly allowedExtensions?: readonly ValidationExtension[];
};

export type ValidationResult = {
  readonly profile: {
    readonly standard: typeof transitionalStandard;
    readonly schemaValidator: typeof schemaValidatorProfile;
    readonly scope: string;
    readonly applicationEvidence: typeof applicationEvidence;
  };
  readonly errors: readonly ValidationIssue[];
  readonly extensions: readonly ValidationExtension[];
  readonly presentation: {
    readonly slideCount: number;
    readonly widthEmu?: number;
    readonly heightEmu?: number;
    readonly title?: string;
    readonly sections?: readonly {
      readonly name: string;
      readonly slideIds: readonly string[];
    }[];
  };
  readonly slides: readonly FileSlide[];
};

type XmlFiles = Record<string, Uint8Array>;
type XmlDocument = ReturnType<DOMParser["parseFromString"]>;
type XmlNode = NonNullable<XmlDocument["firstChild"]>;
type XmlElement = NonNullable<XmlDocument["documentElement"]>;
type Documents = Map<string, XmlDocument>;

type TransitionRead = {
  readonly transition?: FileTransition;
  readonly extensions: ReadonlySet<ValidationExtension>;
};

type PresentationSection = {
  readonly name: string;
  readonly slideIds: readonly string[];
};

function issue(
  errors: ValidationIssue[],
  part: string,
  code: string,
  message: string,
): void {
  errors.push({ part, code, message });
}

async function validateStandardStructure(
  input: Uint8Array,
): Promise<ValidationIssue[]> {
  const validation = schemaValidationQueue.then(async () => {
    const results = await validateOpenXml(
      input,
      "pptx",
      schemaValidatorOfficeVersion,
    );
    return results.map(({ id, description, path }) => ({
      part: path.partUri,
      code: id,
      message: description,
      ...(path.xpath === "" ? {} : { xpath: path.xpath }),
    }));
  });
  schemaValidationQueue = validation.then(
    () => undefined,
    () => undefined,
  );

  try {
    return await validation;
  } catch (error) {
    return [
      {
        part: "",
        code: "schema-validator-failed",
        message: `Open XML SDK structural validator failed: ${error instanceof Error ? error.message : String(error)}`,
      },
    ];
  }
}

function elements(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement[] {
  const result: XmlElement[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (
      child.nodeType === 1 &&
      child.namespaceURI === namespace &&
      child.localName === localName
    ) {
      result.push(child as XmlElement);
    }
  }
  return result;
}

function first(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement | undefined {
  return elements(parent, namespace, localName)[0];
}

function attribute(element: XmlElement, name: string): string {
  return element.getAttribute(name) ?? "";
}

function integerAttribute(
  element: XmlElement | undefined,
  name: string,
): number | undefined {
  if (element === undefined) return undefined;
  const value = attribute(element, name);
  if (value === "" || !/^-?\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function decimalAttribute(
  element: XmlElement | undefined,
  name: string,
): number | undefined {
  if (element === undefined) return undefined;
  const value = attribute(element, name);
  if (value === "") return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function readRotationDegrees(
  transform: XmlElement | undefined,
): number | undefined {
  const rotation = integerAttribute(transform, "rot");
  return rotation === undefined ? undefined : rotation / 60000;
}

type ThemeTextReader = {
  readonly colors: ReadonlyMap<string, string>;
  readonly schemeSlots: ReadonlyMap<string, string>;
  readonly fonts: ReadonlyMap<string, string>;
};

const themeColorSlots = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

function themeTextReader(
  documents: Documents,
  themePath = "ppt/theme/theme1.xml",
  masterPath = "ppt/slideMasters/slideMaster1.xml",
): ThemeTextReader {
  const theme = documents.get(themePath)?.documentElement;
  const themeElements =
    theme == null ? undefined : first(theme, drawingNamespace, "themeElements");
  const colorScheme =
    themeElements === undefined
      ? undefined
      : first(themeElements, drawingNamespace, "clrScheme");
  const colors = new Map<string, string>();
  for (const slot of themeColorSlots) {
    const entry =
      colorScheme === undefined
        ? undefined
        : first(colorScheme, drawingNamespace, slot);
    const rgb =
      entry === undefined
        ? undefined
        : first(entry, drawingNamespace, "srgbClr");
    const system =
      entry === undefined
        ? undefined
        : first(entry, drawingNamespace, "sysClr");
    const value =
      rgb !== undefined
        ? attribute(rgb, "val")
        : system === undefined
          ? ""
          : attribute(system, "lastClr") || attribute(system, "val");
    if (/^[0-9a-f]{6}$/iu.test(value)) colors.set(slot, value.toUpperCase());
  }

  const fontScheme =
    themeElements === undefined
      ? undefined
      : first(themeElements, drawingNamespace, "fontScheme");
  const fonts = new Map<string, string>();
  for (const [name, token] of [
    ["majorFont", "major"],
    ["minorFont", "minor"],
  ] as const) {
    const fontSet =
      fontScheme === undefined
        ? undefined
        : first(fontScheme, drawingNamespace, name);
    const latin =
      fontSet === undefined
        ? undefined
        : first(fontSet, drawingNamespace, "latin");
    const family = latin === undefined ? "" : attribute(latin, "typeface");
    if (family !== "") fonts.set(token, family);
  }

  const master = documents.get(masterPath)?.documentElement;
  const colorMap =
    master == null ? undefined : first(master, presentationNamespace, "clrMap");
  const schemeSlots = new Map<string, string>();
  for (const slot of themeColorSlots) schemeSlots.set(slot, slot);
  for (const [scheme, fallback] of [
    ["tx1", "dk1"],
    ["bg1", "lt1"],
    ["tx2", "dk2"],
    ["bg2", "lt2"],
  ] as const) {
    schemeSlots.set(
      scheme,
      (colorMap === undefined ? "" : attribute(colorMap, scheme)) || fallback,
    );
  }
  return { colors, schemeSlots, fonts };
}

function validateThemeFonts(
  documents: Documents,
  errors: ValidationIssue[],
): void {
  const part = "ppt/theme/theme1.xml";
  const theme = documents.get(part)?.documentElement;
  if (theme == null) return;
  const themeElements = first(theme, drawingNamespace, "themeElements");
  if (themeElements === undefined) return;
  const fontScheme = first(themeElements, drawingNamespace, "fontScheme");
  if (fontScheme === undefined) {
    issue(errors, part, "missing-font-scheme", "Theme is missing font scheme");
    return;
  }
  for (const name of ["majorFont", "minorFont"] as const) {
    const fontSet = first(fontScheme, drawingNamespace, name);
    if (fontSet === undefined) {
      issue(
        errors,
        part,
        "missing-theme-font",
        `Theme font scheme is missing ${name}`,
      );
      continue;
    }
    const latin = first(fontSet, drawingNamespace, "latin");
    const ea = first(fontSet, drawingNamespace, "ea");
    const cs = first(fontSet, drawingNamespace, "cs");
    if (latin === undefined || ea === undefined || cs === undefined) {
      issue(
        errors,
        part,
        "invalid-theme-font-slots",
        `Theme ${name} must define latin, ea, and cs font slots`,
      );
      continue;
    }
    const latinTypeface = attribute(latin, "typeface");
    const eaTypeface = attribute(ea, "typeface");
    const csTypeface = attribute(cs, "typeface");
    if (
      latinTypeface === "" ||
      eaTypeface === "" ||
      csTypeface === "" ||
      latinTypeface !== eaTypeface ||
      latinTypeface !== csTypeface
    ) {
      issue(
        errors,
        part,
        "invalid-theme-font-slots",
        `Theme ${name} must define matching non-empty latin, ea, and cs font slots`,
      );
    }
  }
}

function shadedThemeColor(hex: string, fill: XmlElement): string {
  const lumMod = decimalAttribute(
    first(fill, drawingNamespace, "lumMod"),
    "val",
  );
  const lumOff = decimalAttribute(
    first(fill, drawingNamespace, "lumOff"),
    "val",
  );
  if (lumMod === undefined && lumOff === undefined) return hex;
  const [r, g, b] = [0, 2, 4].map(
    (start) => Number.parseInt(hex.slice(start, start + 2), 16) / 255,
  ) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation =
    delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  const hue =
    delta === 0
      ? 0
      : max === r
        ? ((g - b) / delta + 6) % 6
        : max === g
          ? (b - r) / delta + 2
          : (r - g) / delta + 4;
  const shaded = Math.min(
    1,
    Math.max(0, (lightness * (lumMod ?? 100_000) + (lumOff ?? 0)) / 100_000),
  );
  const chroma = (1 - Math.abs(2 * shaded - 1)) * saturation;
  const x = chroma * (1 - Math.abs((hue % 2) - 1));
  const m = shaded - chroma / 2;
  const [r1, g1, b1] =
    hue < 1
      ? [chroma, x, 0]
      : hue < 2
        ? [x, chroma, 0]
        : hue < 3
          ? [0, chroma, x]
          : hue < 4
            ? [0, x, chroma]
            : hue < 5
              ? [x, 0, chroma]
              : [chroma, 0, x];
  return [r1, g1, b1]
    .map((channel) =>
      Math.round((channel + m) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function readRunColor(
  fill: XmlElement | undefined,
  theme: ThemeTextReader,
): string | undefined {
  if (fill === undefined) return undefined;
  const rgb = first(fill, drawingNamespace, "srgbClr");
  if (rgb !== undefined && attribute(rgb, "val") !== "") {
    return `#${attribute(rgb, "val").toUpperCase()}`;
  }
  const scheme = first(fill, drawingNamespace, "schemeClr");
  if (scheme === undefined) return undefined;
  const slot = theme.schemeSlots.get(attribute(scheme, "val"));
  const hex = slot === undefined ? undefined : theme.colors.get(slot);
  return hex === undefined ? undefined : `#${shadedThemeColor(hex, scheme)}`;
}

function readTextRun(
  run: XmlElement,
  theme: ThemeTextReader,
  part: string,
  errors: ValidationIssue[],
): FileTextRun {
  const text = first(run, drawingNamespace, "t")?.textContent ?? "";
  const fieldType = run.localName === "fld" ? attribute(run, "type") : "";
  const fieldId = run.localName === "fld" ? attribute(run, "id") : "";
  if (
    run.localName === "fld" &&
    (fieldType !== "slidenum" ||
      !/^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/iu.test(
        fieldId,
      ))
  ) {
    issue(
      errors,
      part,
      "invalid-text-field",
      "Native text field type or identifier is invalid",
    );
  }
  const properties = first(run, drawingNamespace, "rPr");
  if (properties === undefined)
    return {
      text,
      ...(fieldType === "slidenum" ? { field: "slidenum" as const } : {}),
    };

  const latin = first(properties, drawingNamespace, "latin");
  const ea = first(properties, drawingNamespace, "ea");
  const cs = first(properties, drawingNamespace, "cs");
  const fill = first(properties, drawingNamespace, "solidFill");
  const rawFontFamily = latin === undefined ? "" : attribute(latin, "typeface");
  if (latin !== undefined) {
    if (ea === undefined || cs === undefined) {
      issue(
        errors,
        part,
        "missing-font-slot",
        "Text run is missing East Asian or Complex Script font slot",
      );
    } else {
      const eaTypeface = attribute(ea, "typeface");
      const csTypeface = attribute(cs, "typeface");
      const expectedEa =
        rawFontFamily === "+mj-lt"
          ? "+mj-ea"
          : rawFontFamily === "+mn-lt"
            ? "+mn-ea"
            : rawFontFamily;
      const expectedCs =
        rawFontFamily === "+mj-lt"
          ? "+mj-cs"
          : rawFontFamily === "+mn-lt"
            ? "+mn-cs"
            : rawFontFamily;
      if (eaTypeface !== expectedEa || csTypeface !== expectedCs) {
        issue(
          errors,
          part,
          "mismatched-font-slots",
          "Text run font slots are inconsistent",
        );
      }
    }
  }
  const themeFont =
    rawFontFamily === "+mj-lt"
      ? theme.fonts.get("major")
      : rawFontFamily === "+mn-lt"
        ? theme.fonts.get("minor")
        : undefined;
  const fontFamily = themeFont ?? rawFontFamily;
  const fontSize = decimalAttribute(properties, "sz");
  const bold = attribute(properties, "b");
  const italic = attribute(properties, "i");
  const color = readRunColor(fill, theme);
  const lang = attribute(properties, "lang");
  return {
    text,
    ...(fieldType === "slidenum" ? { field: "slidenum" as const } : {}),
    ...(fontFamily === "" ? {} : { fontFamily }),
    ...(fontSize === undefined ? {} : { fontSizePt: fontSize / 100 }),
    ...(bold === "" ? {} : { bold: bold === "1" || bold === "true" }),
    ...(italic === "" ? {} : { italic: italic === "1" || italic === "true" }),
    ...(color === undefined ? {} : { color }),
    ...(lang === "" ? {} : { lang }),
  };
}

function readParagraph(
  paragraph: XmlElement,
  theme: ThemeTextReader,
  part: string,
  errors: ValidationIssue[],
): FileTextParagraph {
  const runs = childElements(paragraph)
    .filter(
      (run) =>
        run.namespaceURI === drawingNamespace &&
        (run.localName === "r" || run.localName === "fld"),
    )
    .map((run) => readTextRun(run, theme, part, errors));
  const paragraphProperties = first(paragraph, drawingNamespace, "pPr");
  const lineSpacing =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "lnSpc");
  const exactSpacing =
    lineSpacing === undefined
      ? undefined
      : first(lineSpacing, drawingNamespace, "spcPts");
  const lineSpacingValue = decimalAttribute(exactSpacing, "val");
  const spacingAfter =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "spcAft");
  const spacingAfterValue = decimalAttribute(
    spacingAfter === undefined
      ? undefined
      : first(spacingAfter, drawingNamespace, "spcPts"),
    "val",
  );
  const bullet =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "buChar");
  const autoNumber =
    paragraphProperties === undefined
      ? undefined
      : first(paragraphProperties, drawingNamespace, "buAutoNum");
  const alignment =
    paragraphProperties === undefined
      ? ""
      : attribute(paragraphProperties, "algn");
  let list: FileTextList | undefined;
  if (bullet !== undefined) {
    list = { type: "bullet", character: attribute(bullet, "char") };
  } else if (autoNumber !== undefined) {
    const startAt = integerAttribute(autoNumber, "startAt");
    list = {
      type: "number",
      style: attribute(autoNumber, "type"),
      ...(startAt === undefined ? {} : { startAt }),
    };
  }
  return {
    text: runs.map(({ text }) => text).join(""),
    runs,
    ...(alignment === "" ? {} : { alignment }),
    ...(lineSpacingValue === undefined
      ? {}
      : { lineSpacingPoints: lineSpacingValue / 100 }),
    ...(spacingAfterValue === undefined
      ? {}
      : { spacingAfterPoints: spacingAfterValue / 100 }),
    ...(list === undefined ? {} : { list }),
  };
}

function readTextBody(
  shape: XmlElement,
  theme: ThemeTextReader,
  part: string,
  errors: ValidationIssue[],
): FileText | undefined {
  const body = first(shape, presentationNamespace, "txBody");
  if (body === undefined) return undefined;
  const paragraphs = elements(body, drawingNamespace, "p").map((paragraph) =>
    readParagraph(paragraph, theme, part, errors),
  );
  return {
    text: paragraphs.map(({ text }) => text).join("\n"),
    paragraphs,
  };
}

function parseXmlFiles(files: XmlFiles, errors: ValidationIssue[]): Documents {
  const documents: Documents = new Map();
  const parser = new DOMParser({
    onError(level, message) {
      if (level !== "warning") throw new Error(message);
    },
  });
  for (const [part, bytes] of Object.entries(files)) {
    if (!part.endsWith(".xml") && !part.endsWith(".rels")) continue;
    try {
      const document = parser.parseFromString(
        strFromU8(bytes),
        "application/xml",
      );
      if (document.documentElement === null)
        throw new Error("XML missing root element");
      documents.set(part, document);
    } catch (error) {
      issue(
        errors,
        part,
        "invalid-xml",
        error instanceof Error ? error.message : String(error),
      );
    }
  }
  return documents;
}

function sourcePartForRelationships(path: string): string | undefined {
  if (path === "_rels/.rels") return "";
  const marker = "/_rels/";
  const index = path.lastIndexOf(marker);
  if (index === -1 || !path.endsWith(".rels")) return undefined;
  const directory = path.slice(0, index);
  const file = path.slice(index + marker.length, -".rels".length);
  return directory === "" ? file : `${directory}/${file}`;
}

function relationshipTargetPath(source: string, target: string): string {
  const joined = posix.normalize(
    posix.join(posix.dirname(source || "."), target),
  );
  return joined === "." ? "" : joined;
}

function checkRelationships(
  files: XmlFiles,
  documents: Documents,
  errors: ValidationIssue[],
): Map<string, Map<string, XmlElement>> {
  const bySource = new Map<string, Map<string, XmlElement>>();
  for (const [path, document] of documents) {
    if (!path.endsWith(".rels")) continue;
    const source = sourcePartForRelationships(path);
    if (source === undefined) continue;
    const root = document.documentElement;
    if (
      root === null ||
      root.namespaceURI !== packageRelationshipNamespace ||
      root.localName !== "Relationships"
    ) {
      issue(
        errors,
        path,
        "invalid-relationships-root",
        "Relationships part root element does not match OPC namespace",
      );
      continue;
    }
    const entries = new Map<string, XmlElement>();
    for (const relationship of elements(
      root,
      packageRelationshipNamespace,
      "Relationship",
    )) {
      const id = attribute(relationship, "Id");
      const target = attribute(relationship, "Target");
      const mode = attribute(relationship, "TargetMode");
      if (
        id === "" ||
        target === "" ||
        attribute(relationship, "Type") === ""
      ) {
        issue(
          errors,
          path,
          "invalid-relationship",
          "Relationship missing Id, Type, or Target",
        );
        continue;
      }
      if (entries.has(id))
        issue(
          errors,
          path,
          "duplicate-relationship-id",
          `Duplicate relationship identifier ${id}`,
        );
      entries.set(id, relationship);
      if (mode !== "External") {
        const resolved = relationshipTargetPath(source, target);
        if (resolved.startsWith("../") || files[resolved] === undefined) {
          issue(
            errors,
            path,
            "missing-relationship-target",
            `Relationship target ${target} does not exist`,
          );
        }
      }
    }
    bySource.set(source, entries);
  }
  return bySource;
}

function validateContentTypes(
  files: XmlFiles,
  documents: Documents,
  errors: ValidationIssue[],
): { defaults: Map<string, string>; overrides: Map<string, string> } {
  const part = "[Content_Types].xml";
  const document = documents.get(part);
  if (document === undefined) {
    issue(errors, part, "missing-content-types", "Missing content types part");
    return { defaults: new Map(), overrides: new Map() };
  }
  const root = document.documentElement;
  if (
    root === null ||
    root.namespaceURI !== contentTypeNamespace ||
    root.localName !== "Types"
  ) {
    issue(
      errors,
      part,
      "invalid-content-types-root",
      "Content types part root element does not match OPC namespace",
    );
    return { defaults: new Map(), overrides: new Map() };
  }
  const defaults = new Map(
    elements(root, contentTypeNamespace, "Default").map((entry) => [
      attribute(entry, "Extension").toLowerCase(),
      attribute(entry, "ContentType"),
    ]),
  );
  const overrides = new Map(
    elements(root, contentTypeNamespace, "Override").map((entry) => [
      attribute(entry, "PartName"),
      attribute(entry, "ContentType"),
    ]),
  );
  const required: readonly (readonly [string, string])[] = [
    [
      "/ppt/presentation.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    ],
    [
      "/ppt/slideMasters/slideMaster1.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    ],
    [
      "/ppt/slideLayouts/slideLayout1.xml",
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    ],
    [
      "/ppt/theme/theme1.xml",
      "application/vnd.openxmlformats-officedocument.theme+xml",
    ],
  ];
  for (const [path, contentType] of required) {
    if (overrides.get(path) !== contentType) {
      issue(
        errors,
        part,
        "missing-content-type",
        `${path} missing expected content type`,
      );
    }
  }
  for (const path of Object.keys(files).filter((entry) =>
    /^ppt\/slides\/slide\d+\.xml$/u.test(entry),
  )) {
    if (
      overrides.get(`/${path}`) !==
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    ) {
      issue(
        errors,
        part,
        "missing-content-type",
        `/${path} missing slide content type`,
      );
    }
  }
  for (const path of Object.keys(files).filter((entry) =>
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(entry),
  )) {
    if (
      overrides.get(`/${path}`) !==
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
    ) {
      issue(
        errors,
        part,
        "missing-content-type",
        `/${path} missing notes slide content type`,
      );
    }
  }
  for (const path of Object.keys(files).filter((entry) =>
    /^ppt\/notesMasters\/notesMaster\d+\.xml$/u.test(entry),
  )) {
    if (
      overrides.get(`/${path}`) !==
      "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
    ) {
      issue(
        errors,
        part,
        "missing-content-type",
        `/${path} missing notes master content type`,
      );
    }
  }
  return { defaults, overrides };
}

function contentTypeForPart(
  path: string,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
): string | undefined {
  const override = contentTypes.overrides.get(`/${path}`);
  if (override !== undefined) return override;
  const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return contentTypes.defaults.get(extension);
}

function imageSignatureMatches(
  bytes: Uint8Array,
  contentType: string,
): boolean {
  if (contentType === "image/png") {
    return (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47 &&
      bytes[4] === 0x0d &&
      bytes[5] === 0x0a &&
      bytes[6] === 0x1a &&
      bytes[7] === 0x0a
    );
  }
  if (contentType === "image/jpeg") {
    return bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8;
  }
  if (contentType === "image/gif") {
    const signature = String.fromCharCode(...bytes.subarray(0, 6));
    return signature === "GIF87a" || signature === "GIF89a";
  }
  return false;
}

function requireRelationship(
  source: string,
  type: string,
  targetPath: string,
  relationships: Map<string, Map<string, XmlElement>>,
  errors: ValidationIssue[],
): void {
  const found = [...(relationships.get(source)?.values() ?? [])].some(
    (entry) =>
      attribute(entry, "Type") === type &&
      relationshipTargetPath(source, attribute(entry, "Target")) === targetPath,
  );
  if (!found) {
    issue(
      errors,
      source === "" ? "_rels/.rels" : source,
      "missing-required-relationship",
      `Missing relationship pointing to ${targetPath}`,
    );
  }
}

function requireRoot(
  documents: Documents,
  errors: ValidationIssue[],
  path: string,
  namespace: string,
  localName: string,
): void {
  const root = documents.get(path)?.documentElement;
  if (
    root == null ||
    root.namespaceURI !== namespace ||
    root.localName !== localName
  ) {
    issue(
      errors,
      path,
      "invalid-part-root",
      `Part root element should be ${localName}`,
    );
  }
}

function validateSlideLayoutIds(
  documents: Documents,
  errors: ValidationIssue[],
): void {
  const part = "ppt/slideMasters/slideMaster1.xml";
  const root = documents.get(part)?.documentElement;
  if (root == null) return;
  const layoutIds = first(root, presentationNamespace, "sldLayoutIdLst");
  if (layoutIds === undefined) return;
  for (const layoutId of elements(
    layoutIds,
    presentationNamespace,
    "sldLayoutId",
  )) {
    const id = integerAttribute(layoutId, "id");
    if (
      id === undefined ||
      id < minimumSlideLayoutId ||
      id > maximumSlideLayoutId
    ) {
      issue(
        errors,
        part,
        "invalid-slide-layout-id",
        `Master layout identifier ${attribute(layoutId, "id")} must be an integer between ${minimumSlideLayoutId} and ${maximumSlideLayoutId}`,
      );
    }
  }
}

function validateTextValueRange(
  element: XmlElement | undefined,
  name: string,
  minimum: number,
  maximum: number,
  part: string,
  errors: ValidationIssue[],
  code: "invalid-font-size" | "invalid-text-spacing",
  label: string,
): void {
  if (element === undefined || !element.hasAttribute(name)) return;
  const value = integerAttribute(element, name);
  if (value === undefined || value < minimum || value > maximum) {
    issue(
      errors,
      part,
      code,
      `${label} field must be an integer between ${minimum} and ${maximum}`,
    );
  }
}

function isValidBcp47(tag: string): boolean {
  if (
    typeof tag !== "string" ||
    tag.trim() === "" ||
    !/^[A-Za-z0-9]+(-[A-Za-z0-9]+)*$/.test(tag)
  ) {
    return false;
  }
  try {
    const canonical = Intl.getCanonicalLocales(tag);
    return Array.isArray(canonical) && canonical.length === 1;
  } catch {
    return false;
  }
}

function validateTextLangAttribute(
  element: XmlElement | undefined,
  part: string,
  errors: ValidationIssue[],
): void {
  if (element === undefined || !element.hasAttribute("lang")) return;
  const lang = element.getAttribute("lang");
  if (lang === null || !isValidBcp47(lang)) {
    issue(
      errors,
      part,
      "invalid-text-lang",
      "Text language lang attribute must be a valid BCP 47 language tag",
    );
  }
}

function validateTextValueRanges(
  shape: XmlElement,
  part: string,
  errors: ValidationIssue[],
): void {
  const body = first(shape, presentationNamespace, "txBody");
  if (body === undefined) return;
  validateTextBodyValueRanges(body, part, errors);
}

function validateTextBodyValueRanges(
  body: XmlElement,
  part: string,
  errors: ValidationIssue[],
): void {
  for (const paragraph of elements(body, drawingNamespace, "p")) {
    for (const run of elements(paragraph, drawingNamespace, "r")) {
      const rPr = first(run, drawingNamespace, "rPr");
      validateTextLangAttribute(rPr, part, errors);
      validateTextValueRange(
        rPr,
        "sz",
        minimumFontSizeCentipoints,
        maximumFontSizeCentipoints,
        part,
        errors,
        "invalid-font-size",
        "font-size",
      );
    }
    for (const fld of elements(paragraph, drawingNamespace, "fld")) {
      const rPr = first(fld, drawingNamespace, "rPr");
      validateTextLangAttribute(rPr, part, errors);
    }
    const endParaRPr = first(paragraph, drawingNamespace, "endParaRPr");
    validateTextLangAttribute(endParaRPr, part, errors);
    validateTextValueRange(
      endParaRPr,
      "sz",
      minimumFontSizeCentipoints,
      maximumFontSizeCentipoints,
      part,
      errors,
      "invalid-font-size",
      "font-size",
    );
    const properties = first(paragraph, drawingNamespace, "pPr");
    if (properties !== undefined) {
      validateTextLangAttribute(
        first(properties, drawingNamespace, "defRPr"),
        part,
        errors,
      );
    }
    for (const spacingName of ["lnSpc", "spcAft"] as const) {
      const spacing =
        properties === undefined
          ? undefined
          : first(properties, drawingNamespace, spacingName);
      validateTextValueRange(
        spacing === undefined
          ? undefined
          : first(spacing, drawingNamespace, "spcPts"),
        "val",
        0,
        maximumTextSpacingCentipoints,
        part,
        errors,
        "invalid-text-spacing",
        "text-spacing",
      );
    }
  }
}

function readGeometryName(properties: XmlElement | undefined): string {
  if (properties === undefined) return "";
  const preset = first(properties, drawingNamespace, "prstGeom");
  const custom = first(properties, drawingNamespace, "custGeom");
  if (preset !== undefined && custom !== undefined) return "";
  if (preset !== undefined) return attribute(preset, "prst");
  if (custom === undefined) return "";
  const pathList = first(custom, drawingNamespace, "pathLst");
  const paths =
    pathList === undefined ? [] : elements(pathList, drawingNamespace, "path");
  return paths.length === 0 ? "" : "custom";
}

const validLineEndTypes = new Set([
  "none",
  "triangle",
  "stealth",
  "arrow",
  "diamond",
  "oval",
]);

function validateLineProperties(
  properties: XmlElement | undefined,
  part: string,
  errors: ValidationIssue[],
): void {
  if (properties === undefined) return;
  const line = first(properties, drawingNamespace, "ln");
  if (line === undefined) return;
  for (const endName of ["headEnd", "tailEnd"] as const) {
    const end = first(line, drawingNamespace, endName);
    if (end !== undefined && end.hasAttribute("type")) {
      const type = attribute(end, "type");
      if (!validLineEndTypes.has(type)) {
        issue(
          errors,
          part,
          "invalid-line-end-type",
          `Line end ${endName} has invalid type "${type}". Expected one of: none, triangle, stealth, arrow, diamond, oval`,
        );
      }
    }
  }
}

function validateShape(
  shape: XmlElement,
  part: string,
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): FileObject | undefined {
  const shapeProperties = first(shape, presentationNamespace, "nvSpPr");
  const drawingProperties =
    shapeProperties === undefined
      ? undefined
      : first(shapeProperties, presentationNamespace, "cNvPr");
  const id = integerAttribute(drawingProperties, "id");
  const name =
    drawingProperties === undefined ? "" : attribute(drawingProperties, "name");
  if (id === undefined || name === "") {
    issue(
      errors,
      part,
      "invalid-shape-identity",
      "Shape missing valid object identifier or name",
    );
    return undefined;
  }
  const properties = first(shape, presentationNamespace, "spPr");
  const geometryName = readGeometryName(properties);
  if (geometryName === "") {
    issue(
      errors,
      part,
      "missing-shape-geometry",
      `Shape ${name} missing preset geometry`,
    );
    return undefined;
  }
  const transform =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm");
  const rotation = readRotationDegrees(transform);
  const offset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "off");
  const extent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "ext");
  const x = integerAttribute(offset, "x");
  const y = integerAttribute(offset, "y");
  const width = integerAttribute(extent, "cx");
  const height = integerAttribute(extent, "cy");
  if (
    x === undefined ||
    y === undefined ||
    width === undefined ||
    height === undefined ||
    width < 1 ||
    width > maximumShapeExtentEmu ||
    height < 1 ||
    height > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      "invalid-shape-transform",
      `Shape ${name} coordinates or dimensions are invalid`,
    );
    return undefined;
  }
  validateTextValueRanges(shape, part, errors);
  validateLineProperties(properties, part, errors);
  const text = readTextBody(shape, theme, part, errors);
  return {
    id,
    name,
    kind: "shape",
    geometry: geometryName,
    boxEmu: { x, y, width, height },
    ...(rotation === undefined ? {} : { rotationDegrees: rotation }),
    ...(text === undefined ? {} : { text }),
  };
}

function objectIdentity(
  object: XmlElement,
  propertiesName: string,
  part: string,
  kind: string,
  errors: ValidationIssue[],
): { id: number; name: string } | undefined {
  const properties = first(object, presentationNamespace, propertiesName);
  const drawingProperties =
    properties === undefined
      ? undefined
      : first(properties, presentationNamespace, "cNvPr");
  const id = integerAttribute(drawingProperties, "id");
  const name =
    drawingProperties === undefined ? "" : attribute(drawingProperties, "name");
  if (id === undefined || id < 1 || name === "") {
    issue(
      errors,
      part,
      `invalid-${kind}-identity`,
      `${kind} missing valid object identifier or name`,
    );
    return undefined;
  }
  return { id, name };
}

function transformBox(
  transform: XmlElement | undefined,
  part: string,
  name: string,
  kind: string,
  errors: ValidationIssue[],
): FileObject["boxEmu"] | undefined {
  const offset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "off");
  const extent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "ext");
  const x = integerAttribute(offset, "x");
  const y = integerAttribute(offset, "y");
  const width = integerAttribute(extent, "cx");
  const height = integerAttribute(extent, "cy");
  if (
    x === undefined ||
    y === undefined ||
    width === undefined ||
    height === undefined ||
    width < 1 ||
    width > maximumShapeExtentEmu ||
    height < 1 ||
    height > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      `invalid-${kind}-transform`,
      `${kind} ${name} coordinates or dimensions are invalid`,
    );
    return undefined;
  }
  return { x, y, width, height };
}

function validatePicture(
  picture: XmlElement,
  part: string,
  files: XmlFiles,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): FileObject | undefined {
  const identity = objectIdentity(picture, "nvPicPr", part, "picture", errors);
  if (identity === undefined) return undefined;
  const properties = first(picture, presentationNamespace, "spPr");
  const geometryName = readGeometryName(properties);
  if (geometryName === "") {
    issue(
      errors,
      part,
      "missing-picture-geometry",
      `Image ${identity.name} missing geometry`,
    );
    return undefined;
  }
  const transform =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm");
  const boxEmu = transformBox(
    transform,
    part,
    identity.name,
    "picture",
    errors,
  );
  const rotation = readRotationDegrees(transform);
  const fill = first(picture, presentationNamespace, "blipFill");
  const blip =
    fill === undefined ? undefined : first(fill, drawingNamespace, "blip");
  const sourceRect =
    fill === undefined ? undefined : first(fill, drawingNamespace, "srcRect");
  const crop: NonNullable<FileObject["crop"]> = {};
  for (const [field, attributeName] of [
    ["left", "l"],
    ["top", "t"],
    ["right", "r"],
    ["bottom", "b"],
  ] as const) {
    if (sourceRect?.hasAttribute(attributeName) !== true) continue;
    const value = integerAttribute(sourceRect, attributeName);
    if (value === undefined || value < 0 || value > 100000) {
      issue(
        errors,
        part,
        "invalid-picture-crop",
        `Image ${identity.name} crop values are invalid`,
      );
      continue;
    }
    Object.assign(crop, { [field]: value });
  }
  if ((crop.left ?? 0) + (crop.right ?? 0) > 100000) {
    issue(
      errors,
      part,
      "invalid-picture-crop",
      `Image ${identity.name} horizontal crop out of bounds`,
    );
  }
  if ((crop.top ?? 0) + (crop.bottom ?? 0) > 100000) {
    issue(
      errors,
      part,
      "invalid-picture-crop",
      `Image ${identity.name} vertical crop out of bounds`,
    );
  }
  const relationshipId =
    blip?.getAttributeNS(relationshipNamespace, "embed") ?? "";
  const relationship = relationships.get(part)?.get(relationshipId);
  let image: FileObject["image"];
  if (relationship === undefined) {
    issue(
      errors,
      part,
      "missing-image-relationship",
      `Image ${identity.name} references non-existent relationship ${relationshipId}`,
    );
  } else if (
    attribute(relationship, "Type") !== `${relationshipNamespace}/image` ||
    attribute(relationship, "TargetMode") === "External"
  ) {
    issue(
      errors,
      part,
      "invalid-image-relationship",
      `Image ${identity.name} must reference package image part`,
    );
  } else {
    const imagePart = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const bytes = files[imagePart];
    const contentType = contentTypeForPart(imagePart, contentTypes);
    const expectedType = imagePart.endsWith(".png")
      ? "image/png"
      : imagePart.endsWith(".jpeg") || imagePart.endsWith(".jpg")
        ? "image/jpeg"
        : imagePart.endsWith(".gif")
          ? "image/gif"
          : undefined;
    if (
      !imagePart.startsWith("ppt/media/") ||
      bytes === undefined ||
      contentType === undefined ||
      contentType !== expectedType ||
      expectedType === undefined ||
      !imageSignatureMatches(bytes, contentType)
    ) {
      issue(
        errors,
        part,
        "invalid-image-part",
        `Image ${identity.name} media part is missing or content type is invalid`,
      );
    } else {
      image = { relationshipId, part: imagePart, contentType };
    }
  }
  validateLineProperties(properties, part, errors);
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "image",
    geometry: geometryName,
    boxEmu,
    ...(rotation === undefined ? {} : { rotationDegrees: rotation }),
    ...(image === undefined ? {} : { image }),
    ...(Object.keys(crop).length === 0 ? {} : { crop }),
  };
}

type ChartCache<T> = {
  readonly formula: string;
  readonly values: readonly T[];
};

function chartStringCache(
  reference: XmlElement | undefined,
  part: string,
  errors: ValidationIssue[],
): ChartCache<string> | undefined {
  const formulaElement =
    reference === undefined ? undefined : first(reference, chartNamespace, "f");
  const cache =
    reference === undefined
      ? undefined
      : first(reference, chartNamespace, "strCache");
  const formula = formulaElement?.textContent ?? "";
  const count = integerAttribute(
    cache === undefined ? undefined : first(cache, chartNamespace, "ptCount"),
    "val",
  );
  if (formula === "" || count === undefined || count < 0) {
    issue(
      errors,
      part,
      "invalid-chart-cache",
      "Chart string reference missing formula or cache count",
    );
    return undefined;
  }
  const values = Array.from({ length: count }, () => "");
  for (const point of cache === undefined
    ? []
    : elements(cache, chartNamespace, "pt")) {
    const index = integerAttribute(point, "idx");
    const value = first(point, chartNamespace, "v")?.textContent;
    if (
      index === undefined ||
      index < 0 ||
      index >= values.length ||
      value == null
    ) {
      issue(
        errors,
        part,
        "invalid-chart-cache",
        "Chart string cache contains invalid data points",
      );
      continue;
    }
    values[index] = value;
  }
  return { formula, values };
}

function chartNumberCache(
  reference: XmlElement | undefined,
  part: string,
  errors: ValidationIssue[],
): ChartCache<number | null> | undefined {
  const formulaElement =
    reference === undefined ? undefined : first(reference, chartNamespace, "f");
  const cache =
    reference === undefined
      ? undefined
      : first(reference, chartNamespace, "numCache");
  const formula = formulaElement?.textContent ?? "";
  const count = integerAttribute(
    cache === undefined ? undefined : first(cache, chartNamespace, "ptCount"),
    "val",
  );
  if (formula === "" || count === undefined || count < 0) {
    issue(
      errors,
      part,
      "invalid-chart-cache",
      "Chart numeric reference missing formula or cache count",
    );
    return undefined;
  }
  const values: (number | null)[] = Array.from({ length: count }, () => null);
  for (const point of cache === undefined
    ? []
    : elements(cache, chartNamespace, "pt")) {
    const index = integerAttribute(point, "idx");
    const text = first(point, chartNamespace, "v")?.textContent;
    const value = text === undefined ? NaN : Number(text);
    if (
      index === undefined ||
      index < 0 ||
      index >= values.length ||
      !Number.isFinite(value)
    ) {
      issue(
        errors,
        part,
        "invalid-chart-cache",
        "Chart numeric cache contains invalid data points",
      );
      continue;
    }
    values[index] = value;
  }
  return { formula, values };
}

function spreadsheetColumn(index: number): string {
  let column = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    column = String.fromCharCode(65 + ((value - 1) % 26)) + column;
  }
  return column;
}

/** Reads the visible text of a spreadsheet string item, including rich-text runs but excluding phonetic runs. */
function spreadsheetStringText(element: XmlElement): string {
  let text = "";
  for (
    let child = element.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (
      child.nodeType !== 1 ||
      child.namespaceURI !== spreadsheetNamespace ||
      child.localName === "rPh"
    ) {
      continue;
    }
    text +=
      child.localName === "t"
        ? (child.textContent ?? "")
        : spreadsheetStringText(child as XmlElement);
  }
  return text;
}

type ChartFormulaReference = {
  readonly name: string;
  readonly categories: string;
  readonly values: string;
};

type ChartFormulaSeries = {
  readonly series: FileChartSeries;
  readonly formula: ChartFormulaReference;
};

type FormulaRange = {
  readonly sheetName: string;
  readonly startColumn: number;
  readonly startRow: number;
  readonly endColumn: number;
  readonly endRow: number;
  readonly hasRange: boolean;
};

function spreadsheetColumnNumber(column: string): number | undefined {
  let value = 0;
  for (const character of column) {
    const digit = character.charCodeAt(0) - 65;
    if (digit < 0 || digit > 25) return undefined;
    value = value * 26 + digit + 1;
  }
  return value === 0 ? undefined : value;
}

/** Parses only the A1 references emitted by the chart writer, never arbitrary Excel formulas. */
function chartFormulaRange(
  formula: string,
  expectedSheet: string,
): FormulaRange | undefined {
  const match =
    /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/u.exec(
      formula,
    );
  if (match === null) return undefined;
  const rawSheet = match[1] ?? match[2] ?? "";
  const sheetName =
    match[1] === undefined ? rawSheet : rawSheet.replace(/''/gu, "'");
  if (sheetName !== expectedSheet) return undefined;
  const startColumn = spreadsheetColumnNumber(match[3]!);
  const startRow = Number(match[4]);
  const endColumn =
    match[5] === undefined ? startColumn : spreadsheetColumnNumber(match[5]);
  const endRow = match[6] === undefined ? startRow : Number(match[6]);
  if (
    startColumn === undefined ||
    endColumn === undefined ||
    !Number.isSafeInteger(startRow) ||
    !Number.isSafeInteger(endRow) ||
    startRow < 1 ||
    endRow < startRow
  ) {
    return undefined;
  }
  return {
    sheetName,
    startColumn,
    startRow,
    endColumn,
    endRow,
    hasRange: match[6] !== undefined,
  };
}

function readEmbeddedWorkbook(
  part: string,
  chartPart: string,
  bytes: Uint8Array,
  formulaSeries: readonly ChartFormulaSeries[],
  errors: ValidationIssue[],
): FileChart["embeddedWorkbook"] {
  let files: XmlFiles;
  try {
    files = unzipSync(bytes);
  } catch (error) {
    issue(
      errors,
      part,
      "invalid-embedded-workbook",
      error instanceof Error
        ? error.message
        : "Embedded workbook is not a valid ZIP part",
    );
    return undefined;
  }
  const nestedErrors: ValidationIssue[] = [];
  const documents = parseXmlFiles(files, nestedErrors);
  const relationships = checkRelationships(files, documents, nestedErrors);
  for (const nestedError of nestedErrors) {
    issue(
      errors,
      `${part}!/${nestedError.part}`,
      nestedError.code,
      nestedError.message,
    );
  }
  const workbookPart = "xl/workbook.xml";
  const workbook = documents.get(workbookPart)?.documentElement;
  if (
    workbook == null ||
    workbook.namespaceURI !== spreadsheetNamespace ||
    workbook.localName !== "workbook"
  ) {
    issue(
      errors,
      part,
      "invalid-embedded-workbook",
      "Embedded workbook missing xl/workbook.xml",
    );
    return undefined;
  }
  const sheet = elements(
    first(workbook, spreadsheetNamespace, "sheets") ?? workbook,
    spreadsheetNamespace,
    "sheet",
  )[0];
  const sheetName = sheet === undefined ? "" : attribute(sheet, "name");
  const sheetRelationshipId =
    sheet?.getAttributeNS(relationshipNamespace, "id") ?? "";
  const sheetRelationship = relationships
    .get(workbookPart)
    ?.get(sheetRelationshipId);
  if (sheetName === "" || sheetRelationship === undefined) {
    issue(
      errors,
      part,
      "invalid-embedded-workbook",
      "Embedded workbook missing valid worksheet relationship",
    );
    return undefined;
  }
  const sheetPart = relationshipTargetPath(
    workbookPart,
    attribute(sheetRelationship, "Target"),
  );
  const worksheet = documents.get(sheetPart)?.documentElement;
  if (
    worksheet == null ||
    worksheet.namespaceURI !== spreadsheetNamespace ||
    worksheet.localName !== "worksheet"
  ) {
    issue(
      errors,
      part,
      "invalid-embedded-workbook",
      "Embedded workbook worksheet part is invalid",
    );
    return undefined;
  }
  // PowerPoint stores embedded chart workbook strings in sharedStrings and
  // references them from cells with t="s"; the cell value is an index, not text.
  const sharedStringsRelationship = [
    ...(relationships.get(workbookPart)?.values() ?? []),
  ].find(
    (entry) =>
      attribute(entry, "Type") === `${relationshipNamespace}/sharedStrings` &&
      attribute(entry, "TargetMode") !== "External",
  );
  let sharedStrings: readonly string[] | undefined;
  if (sharedStringsRelationship !== undefined) {
    const sharedStringsPart = relationshipTargetPath(
      workbookPart,
      attribute(sharedStringsRelationship, "Target"),
    );
    const sharedStringsRoot = documents.get(sharedStringsPart)?.documentElement;
    if (
      sharedStringsRoot == null ||
      sharedStringsRoot.namespaceURI !== spreadsheetNamespace ||
      sharedStringsRoot.localName !== "sst"
    ) {
      issue(
        errors,
        part,
        "invalid-embedded-workbook",
        "Embedded workbook shared strings part is invalid",
      );
    } else {
      sharedStrings = elements(
        sharedStringsRoot,
        spreadsheetNamespace,
        "si",
      ).map(spreadsheetStringText);
    }
  }
  const sheetData = first(worksheet, spreadsheetNamespace, "sheetData");
  const cells = new Map<string, string>();
  for (const row of sheetData === undefined
    ? []
    : elements(sheetData, spreadsheetNamespace, "row")) {
    for (const cell of elements(row, spreadsheetNamespace, "c")) {
      const reference = attribute(cell, "r");
      const type = attribute(cell, "t");
      const inline = first(cell, spreadsheetNamespace, "is");
      let value = "";
      if (type === "inlineStr") {
        value = inline === undefined ? "" : spreadsheetStringText(inline);
      } else if (type === "s") {
        const text = first(cell, spreadsheetNamespace, "v")?.textContent ?? "";
        const index = text.trim() === "" ? Number.NaN : Number(text);
        const entry = Number.isInteger(index)
          ? sharedStrings?.[index]
          : undefined;
        if (entry === undefined) {
          issue(
            errors,
            part,
            "invalid-embedded-workbook-value",
            `Workbook ${reference} shared string index is invalid`,
          );
        } else {
          value = entry;
        }
      } else {
        value = first(cell, spreadsheetNamespace, "v")?.textContent ?? "";
      }
      if (reference !== "") cells.set(reference, value);
    }
  }

  const cellAt = (column: number, row: number) =>
    cells.get(`${spreadsheetColumn(column - 1)}${row}`);

  const workbookSeries: FileChartSeries[] = [];
  for (const { series, formula } of formulaSeries) {
    const nameRange = chartFormulaRange(formula.name, sheetName);
    const categoryRange = chartFormulaRange(formula.categories, sheetName);
    const valueRange = chartFormulaRange(formula.values, sheetName);
    if (
      nameRange === undefined ||
      categoryRange === undefined ||
      valueRange === undefined ||
      nameRange.hasRange ||
      !categoryRange.hasRange ||
      !valueRange.hasRange ||
      categoryRange.startColumn !== categoryRange.endColumn ||
      valueRange.startColumn !== valueRange.endColumn ||
      categoryRange.startRow !== valueRange.startRow ||
      categoryRange.endRow !== valueRange.endRow
    ) {
      issue(
        errors,
        chartPart,
        "invalid-chart-formula",
        `Chart series ${series.name} does not reference corresponding workbook cell`,
      );
      continue;
    }
    if (
      nameRange.startRow < 1 ||
      categoryRange.startRow < 2 ||
      categoryRange.endRow - categoryRange.startRow + 1 !==
        series.categories.length ||
      valueRange.endRow - valueRange.startRow + 1 !== series.values.length
    ) {
      issue(
        errors,
        chartPart,
        "invalid-chart-formula",
        `Chart series ${series.name} does not reference corresponding workbook cell`,
      );
      continue;
    }
    const nameCell = cellAt(nameRange.startColumn, nameRange.startRow);
    if (nameCell === undefined) {
      issue(
        errors,
        part,
        "incomplete-embedded-workbook",
        `Workbook missing series name ${spreadsheetColumn(nameRange.startColumn - 1)}${nameRange.startRow}`,
      );
    }
    const categories = Array.from(
      { length: categoryRange.endRow - categoryRange.startRow + 1 },
      (_, index) => {
        const row = categoryRange.startRow + index;
        const category = cellAt(categoryRange.startColumn, row);
        if (category === undefined) {
          issue(
            errors,
            part,
            "incomplete-embedded-workbook",
            `Workbook missing category ${spreadsheetColumn(categoryRange.startColumn - 1)}${row}`,
          );
        }
        return category ?? "";
      },
    );
    const values = Array.from(
      { length: valueRange.endRow - valueRange.startRow + 1 },
      (_, index) => {
        const row = valueRange.startRow + index;
        const text = cellAt(valueRange.startColumn, row);
        if (text === undefined) return null;
        const value = Number(text);
        if (!Number.isFinite(value)) {
          issue(
            errors,
            part,
            "invalid-embedded-workbook-value",
            `Workbook ${spreadsheetColumn(valueRange.startColumn - 1)}${row} is not a finite number`,
          );
          return null;
        }
        return value;
      },
    );
    if (
      nameCell !== series.name ||
      categories.some((category, row) => category !== series.categories[row]) ||
      values.some((value, row) => value !== series.values[row])
    ) {
      issue(
        errors,
        chartPart,
        "chart-workbook-mismatch",
        `Chart series ${series.name} does not match embedded workbook content`,
      );
    }
    workbookSeries.push({
      name: nameCell ?? "",
      categories,
      values,
      index: series.index,
      order: series.order,
      groupIndex: series.groupIndex,
      ...(series.marker === undefined ? {} : { marker: series.marker }),
      ...(series.smooth === undefined ? {} : { smooth: series.smooth }),
    });
  }
  return {
    part,
    sheetName,
    categories: workbookSeries[0]?.categories ?? [],
    series: workbookSeries,
  };
}

function booleanAttribute(
  element: XmlElement | undefined,
): boolean | undefined {
  if (element === undefined) return undefined;
  const value = attribute(element, "val");
  if (value === "" || value === "1" || value === "true") return true;
  if (value === "0" || value === "false") return false;
  return undefined;
}

function mergeFileDataLabels(
  base: false | FileChartDataLabel | undefined,
  override: false | FileChartDataLabel | undefined,
): false | FileChartDataLabel | undefined {
  if (override === undefined) return base;
  if (override === false) return false;
  if (base === false || base === undefined) return override;
  return { ...base, ...override };
}

function readChartLine(
  line: XmlElement | undefined,
  theme: ThemeTextReader,
): FileChartLine | undefined {
  if (line === undefined) return undefined;
  const color = readRunColor(first(line, drawingNamespace, "solidFill"), theme);
  const widthEmu = integerAttribute(line, "w");
  const dashValue = attribute(
    first(line, drawingNamespace, "prstDash") ?? line,
    "val",
  );
  if (color === undefined && widthEmu === undefined && dashValue === "") {
    return undefined;
  }
  return {
    ...(color === undefined ? {} : { color }),
    ...(widthEmu === undefined ? {} : { widthEmu }),
    ...(dashValue === "" ? {} : { dash: dashValue }),
  };
}

function hasVisibleDataLabelItem(node: XmlElement): boolean {
  return ["showVal", "showCatName", "showSerName", "showPercent"].some(
    (name) => booleanAttribute(first(node, chartNamespace, name)) === true,
  );
}

function readDataLabelSettings(
  node: XmlElement | undefined,
  chartType: string | null,
  part: string,
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): false | FileChartDataLabel | undefined {
  if (node === undefined) return undefined;
  const deleteNode = first(node, chartNamespace, "delete");
  if (deleteNode !== undefined && booleanAttribute(deleteNode) !== false) {
    return false;
  }
  const settings: {
    value?: boolean;
    percent?: boolean;
    category?: boolean;
    position?: Exclude<FileChartDataLabel["position"], undefined>;
    numberFormat?: string;
    leaderLines?: boolean;
    leaderLine?: FileChartLine;
  } = {};
  const numberFormatNode = first(node, chartNamespace, "numFmt");
  if (numberFormatNode !== undefined) {
    const numberFormat = attribute(numberFormatNode, "formatCode");
    if (numberFormat !== "") settings.numberFormat = numberFormat;
  }
  const positionNode = first(node, chartNamespace, "dLblPos");
  if (positionNode !== undefined) {
    const position = attribute(positionNode, "val");
    if (position === "ctr") settings.position = "center";
    else if (position === "inEnd") settings.position = "insideEnd";
    else if (position === "inBase") settings.position = "insideBase";
    else if (position === "outEnd") settings.position = "outsideEnd";
    else if (position === "bestFit") settings.position = "bestFit";
    else if (position === "l") settings.position = "left";
    else if (position === "r") settings.position = "right";
    else if (position === "b") settings.position = "bottom";
    else if (position === "t") settings.position = "top";
  }
  const value = booleanAttribute(first(node, chartNamespace, "showVal"));
  if (value !== undefined) settings.value = value;
  const category = booleanAttribute(first(node, chartNamespace, "showCatName"));
  if (category !== undefined) settings.category = category;
  const percent = booleanAttribute(first(node, chartNamespace, "showPercent"));
  if (percent !== undefined) {
    settings.percent = percent;
    if (
      percent &&
      chartType !== "pieChart" &&
      chartType !== "doughnutChart" &&
      chartType !== "ofPieChart" &&
      chartType !== "pie3DChart"
    ) {
      issue(
        errors,
        part,
        "invalid-data-label-percent",
        `Office does not allow showPercent in ${chartType ?? "this chart type"}`,
      );
    }
  }
  const showLeaderLines = booleanAttribute(
    first(node, chartNamespace, "showLeaderLines"),
  );
  if (showLeaderLines !== undefined) {
    settings.leaderLines = showLeaderLines;
    if (
      showLeaderLines &&
      chartType !== "pieChart" &&
      chartType !== "doughnutChart" &&
      chartType !== "ofPieChart" &&
      chartType !== "pie3DChart"
    ) {
      issue(
        errors,
        part,
        "invalid-data-label-leader-lines",
        `Office does not allow showLeaderLines in ${chartType ?? "this chart type"}`,
      );
    }
    // The Office constraint applies whenever the element is present, including false.
    if (!hasVisibleDataLabelItem(node)) {
      issue(
        errors,
        part,
        "invalid-data-label-leader-lines",
        "showLeaderLines requires at least one visible sibling value, category, series name, or percent label",
      );
    }
  }
  const leaderLinesNode =
    first(node, chartNamespace, "leaderLines") ?? undefined;
  if (leaderLinesNode !== undefined) {
    const shape = first(leaderLinesNode, chartNamespace, "spPr");
    const line = readChartLine(
      shape === undefined ? undefined : first(shape, drawingNamespace, "ln"),
      theme,
    );
    if (line !== undefined) settings.leaderLine = line;
    if (settings.leaderLines !== true) {
      issue(
        errors,
        part,
        "invalid-data-label-leader-lines",
        "leaderLines requires a sibling showLeaderLines value of true",
      );
    }
  }
  return settings;
}

function officeAllowsSeriesDataLabelPosition(
  chartType: string | null,
  grouping: string,
  position: string,
): boolean {
  if (chartType === "pieChart") {
    return ["bestFit", "outEnd", "inEnd", "ctr"].includes(position);
  }
  if (chartType === "barChart") {
    const supported = ["ctr", "inEnd", "inBase"];
    if (grouping !== "stacked" && grouping !== "percentStacked") {
      supported.push("outEnd");
    }
    return supported.includes(position);
  }
  if (chartType === "lineChart") {
    return ["l", "r", "b", "t", "ctr"].includes(position);
  }
  return false;
}

function nativeChartTitleText(
  title: XmlElement | undefined,
  theme: ThemeTextReader,
  part: string,
  errors: ValidationIssue[],
): string | undefined {
  if (title === undefined) return undefined;
  const tx = first(title, chartNamespace, "tx");
  const rich = tx === undefined ? undefined : first(tx, chartNamespace, "rich");
  if (rich === undefined) return undefined;
  const text = elements(rich, drawingNamespace, "p")
    .map((paragraph) => readParagraph(paragraph, theme, part, errors).text)
    .join("\n");
  return text === "" ? undefined : text;
}

function validateChart(
  frame: XmlElement,
  part: string,
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): FileObject | undefined {
  const identity = objectIdentity(
    frame,
    "nvGraphicFramePr",
    part,
    "chart",
    errors,
  );
  if (identity === undefined) return undefined;
  const transform = first(frame, presentationNamespace, "xfrm");
  const boxEmu = transformBox(transform, part, identity.name, "chart", errors);
  const graphic = first(
    frame,
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "graphic",
  );
  const graphicData =
    graphic === undefined
      ? undefined
      : first(graphic, drawingNamespace, "graphicData");
  const reference =
    graphicData === undefined
      ? undefined
      : first(graphicData, chartNamespace, "chart");
  const relationshipId =
    reference?.getAttributeNS(relationshipNamespace, "id") ?? "";
  const relationship = relationships.get(part)?.get(relationshipId);
  if (
    relationship === undefined ||
    attribute(relationship, "Type") !== `${relationshipNamespace}/chart` ||
    attribute(relationship, "TargetMode") === "External"
  ) {
    issue(
      errors,
      part,
      "invalid-chart-relationship",
      `Chart ${identity.name} missing package chart relationship`,
    );
    return boxEmu === undefined
      ? undefined
      : { id: identity.id, name: identity.name, kind: "chart", boxEmu };
  }
  const chartPart = relationshipTargetPath(
    part,
    attribute(relationship, "Target"),
  );
  const chartXml = documents.get(chartPart)?.documentElement;
  if (
    chartXml == null ||
    chartXml.namespaceURI !== chartNamespace ||
    chartXml.localName !== "chartSpace" ||
    contentTypeForPart(chartPart, contentTypes) !==
      "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
  ) {
    issue(
      errors,
      chartPart,
      "invalid-chart-part",
      `Chart ${identity.name} part or content type is invalid`,
    );
    return boxEmu === undefined
      ? undefined
      : { id: identity.id, name: identity.name, kind: "chart", boxEmu };
  }
  const chartLang = first(chartXml, chartNamespace, "lang");
  if (chartLang !== undefined) {
    const val = attribute(chartLang, "val");
    if (!isValidBcp47(val)) {
      issue(
        errors,
        chartPart,
        "invalid-text-lang",
        "Chart c:lang must be a valid BCP 47 language tag",
      );
    }
  }
  for (const defRPr of descendantsByName(
    chartXml,
    drawingNamespace,
    "defRPr",
  )) {
    validateTextLangAttribute(defRPr, chartPart, errors);
    const latin = first(defRPr, drawingNamespace, "latin");
    const ea = first(defRPr, drawingNamespace, "ea");
    const cs = first(defRPr, drawingNamespace, "cs");
    if (latin !== undefined) {
      const rawFamily = attribute(latin, "typeface");
      if (ea === undefined || cs === undefined) {
        issue(
          errors,
          chartPart,
          "missing-font-slot",
          "Chart defRPr is missing East Asian or Complex Script font slot",
        );
      } else {
        const eaTypeface = attribute(ea, "typeface");
        const csTypeface = attribute(cs, "typeface");
        const expectedEa =
          rawFamily === "+mj-lt"
            ? "+mj-ea"
            : rawFamily === "+mn-lt"
              ? "+mn-ea"
              : rawFamily;
        const expectedCs =
          rawFamily === "+mj-lt"
            ? "+mj-cs"
            : rawFamily === "+mn-lt"
              ? "+mn-cs"
              : rawFamily;
        if (eaTypeface !== expectedEa || csTypeface !== expectedCs) {
          issue(
            errors,
            chartPart,
            "mismatched-font-slots",
            "Chart defRPr font slots are inconsistent",
          );
        }
      }
    }
  }
  for (const endParaRPr of descendantsByName(
    chartXml,
    drawingNamespace,
    "endParaRPr",
  )) {
    validateTextLangAttribute(endParaRPr, chartPart, errors);
  }
  const chart = first(chartXml, chartNamespace, "chart");
  const plotArea =
    chart === undefined ? undefined : first(chart, chartNamespace, "plotArea");
  const chartTypeNames = new Set([
    "barChart",
    "lineChart",
    "areaChart",
    "pieChart",
    "doughnutChart",
  ]);
  const typeNodes: XmlElement[] = [];
  for (
    let child = plotArea?.firstChild ?? null;
    child !== null;
    child = child.nextSibling
  ) {
    if (
      child.nodeType === 1 &&
      child.namespaceURI === chartNamespace &&
      typeof child.localName === "string" &&
      chartTypeNames.has(child.localName)
    ) {
      typeNodes.push(child as XmlElement);
    }
  }
  if (typeNodes.length === 0) {
    issue(
      errors,
      chartPart,
      "missing-chart-type",
      "Chart part missing native chart type",
    );
  }
  const types = typeNodes
    .map(({ localName }) => localName)
    .filter((localName): localName is string => typeof localName === "string");
  const firstType = typeNodes[0];
  const direction =
    firstType?.localName === "barChart"
      ? attribute(
          first(firstType, chartNamespace, "barDir") ?? firstType,
          "val",
        )
      : undefined;
  if (direction !== undefined && direction !== "bar" && direction !== "col") {
    issue(
      errors,
      chartPart,
      "invalid-chart-direction",
      "Chart bar direction is invalid",
    );
  }
  const grouping =
    firstType === undefined
      ? undefined
      : attribute(
          first(firstType, chartNamespace, "grouping") ?? firstType,
          "val",
        );

  const axisNodes: XmlElement[] = [];
  for (
    let child = plotArea?.firstChild ?? null;
    child !== null;
    child = child.nextSibling
  ) {
    if (
      child.nodeType === 1 &&
      child.namespaceURI === chartNamespace &&
      (child.localName === "catAx" || child.localName === "valAx")
    ) {
      axisNodes.push(child as XmlElement);
    }
  }
  const axes: FileChartAxis[] = [];
  const axisById = new Map<number, FileChartAxis>();
  const textTheme = themeTextReader(documents);
  for (const axisNode of axisNodes) {
    const id = integerAttribute(first(axisNode, chartNamespace, "axId"), "val");
    const crossAxisId = integerAttribute(
      first(axisNode, chartNamespace, "crossAx"),
      "val",
    );
    if (id === undefined || id < 0 || crossAxisId === undefined) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        "Chart axis is missing a valid numeric id or crossing axis",
      );
      continue;
    }
    if (axisById.has(id)) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        `Chart axis id ${id} is duplicated`,
      );
      continue;
    }
    const scaling = first(axisNode, chartNamespace, "scaling");
    const orientation =
      scaling === undefined
        ? ""
        : attribute(
            first(scaling, chartNamespace, "orientation") ?? scaling,
            "val",
          );
    if (
      orientation !== "" &&
      orientation !== "minMax" &&
      orientation !== "maxMin"
    ) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        `Chart axis ${id} has an invalid orientation`,
      );
    }
    const logBase =
      scaling === undefined
        ? undefined
        : decimalAttribute(first(scaling, chartNamespace, "logBase"), "val");
    if (logBase !== undefined && (logBase < 2 || logBase > 1000)) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        `Chart axis ${id} has a logarithmic base outside 2 to 1000`,
      );
    }
    const min =
      scaling === undefined
        ? undefined
        : decimalAttribute(first(scaling, chartNamespace, "min"), "val");
    const max =
      scaling === undefined
        ? undefined
        : decimalAttribute(first(scaling, chartNamespace, "max"), "val");
    const majorUnit = decimalAttribute(
      first(axisNode, chartNamespace, "majorUnit"),
      "val",
    );
    const minorUnit = decimalAttribute(
      first(axisNode, chartNamespace, "minorUnit"),
      "val",
    );
    const numberFormatNode = first(axisNode, chartNamespace, "numFmt");
    const numberFormat =
      numberFormatNode === undefined
        ? ""
        : attribute(numberFormatNode, "formatCode");
    const majorTickMark = attribute(
      first(axisNode, chartNamespace, "majorTickMark") ?? axisNode,
      "val",
    );
    const minorTickMark = attribute(
      first(axisNode, chartNamespace, "minorTickMark") ?? axisNode,
      "val",
    );
    const tickLabelPosition = attribute(
      first(axisNode, chartNamespace, "tickLblPos") ?? axisNode,
      "val",
    );
    const crosses = attribute(
      first(axisNode, chartNamespace, "crosses") ?? axisNode,
      "val",
    );
    const crossesAt = decimalAttribute(
      first(axisNode, chartNamespace, "crossesAt"),
      "val",
    );
    const title = nativeChartTitleText(
      first(axisNode, chartNamespace, "title"),
      textTheme,
      chartPart,
      errors,
    );
    const axis: FileChartAxis = {
      id,
      kind: axisNode.localName === "catAx" ? "category" : "value",
      crossAxisId,
      ...(title === undefined ? {} : { title }),
      ...(orientation === "maxMin" ? { orientation } : {}),
      ...(logBase === undefined ? {} : { logBase }),
      ...(min === undefined ? {} : { min }),
      ...(max === undefined ? {} : { max }),
      ...(majorUnit === undefined ? {} : { majorUnit }),
      ...(minorUnit === undefined ? {} : { minorUnit }),
      ...(numberFormat === "" || numberFormat === "General"
        ? {}
        : { numberFormat }),
      ...(majorTickMark === "" || majorTickMark === "none"
        ? {}
        : { majorTickMark }),
      ...(minorTickMark === "" || minorTickMark === "none"
        ? {}
        : { minorTickMark }),
      ...(tickLabelPosition === "" || tickLabelPosition === "nextTo"
        ? {}
        : { tickLabelPosition }),
      ...(crosses === "" || crosses === "autoZero" ? {} : { crosses }),
      ...(crossesAt === undefined ? {} : { crossesAt }),
    };
    axes.push(axis);
    axisById.set(id, axis);
  }
  for (const axis of axes) {
    const crossAxis = axisById.get(axis.crossAxisId);
    if (crossAxis === undefined || crossAxis.crossAxisId !== axis.id) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        `Chart axis ${axis.id} has an invalid crossing axis`,
      );
    }
  }

  const series: FileChartSeries[] = [];
  const formulaSeries: ChartFormulaSeries[] = [];
  const groups: FileChartGroup[] = [];
  const seriesDataLabels: FileChartSeriesDataLabels[] = [];
  const seenSeriesIndexes = new Set<number>();
  const seenSeriesOrders = new Set<number>();
  for (const [groupIndex, typeNode] of typeNodes.entries()) {
    const typeGrouping = attribute(
      first(typeNode, chartNamespace, "grouping") ?? typeNode,
      "val",
    );
    const groupDataLabels = readDataLabelSettings(
      first(typeNode, chartNamespace, "dLbls"),
      typeNode.localName,
      chartPart,
      errors,
      textTheme,
    );
    const axisIds = elements(typeNode, chartNamespace, "axId").flatMap(
      (axisId) => {
        const value = integerAttribute(axisId, "val");
        if (value === undefined || value < 0) {
          issue(
            errors,
            chartPart,
            "invalid-chart-axis",
            "Chart group contains an invalid axis reference",
          );
          return [];
        }
        return [value];
      },
    );
    if (
      typeNode.localName === "pieChart" ||
      typeNode.localName === "doughnutChart"
    ) {
      if (axisIds.length !== 0) {
        issue(
          errors,
          chartPart,
          "invalid-chart-axis",
          "Pie and doughnut chart groups do not reference category or value axes",
        );
      }
    } else if (axisIds.length !== 2) {
      issue(
        errors,
        chartPart,
        "invalid-chart-axis",
        "Cartesian chart group must reference a category axis and a value axis",
      );
    } else {
      const roles = axisIds.map((id) => {
        const axis = axisById.get(id);
        if (axis === undefined) {
          issue(
            errors,
            chartPart,
            "unknown-chart-axis",
            `Chart group references unknown axis id ${id}`,
          );
          return undefined;
        }
        return axis.kind;
      });
      if (
        !roles.includes("category") ||
        !roles.includes("value") ||
        new Set(roles).size !== 2
      ) {
        issue(
          errors,
          chartPart,
          "invalid-chart-axis-role",
          "Chart group axis references must include one category axis and one value axis",
        );
      }
    }
    for (const positionNode of descendantsByName(
      typeNode,
      chartNamespace,
      "dLblPos",
    )) {
      let ancestor = positionNode.parentNode;
      let hasSeriesAncestor = false;
      while (ancestor !== null && ancestor !== typeNode) {
        if (
          ancestor.nodeType === 1 &&
          ancestor.namespaceURI === chartNamespace &&
          ancestor.localName === "ser"
        ) {
          hasSeriesAncestor = true;
          break;
        }
        ancestor = ancestor.parentNode;
      }
      // MS-OI29500 §2.1.1456 applies chart-group restrictions only below c:ser.
      if (!hasSeriesAncestor) continue;

      const posVal = attribute(positionNode, "val");
      if (
        !officeAllowsSeriesDataLabelPosition(
          typeNode.localName,
          typeGrouping,
          posVal,
        )
      ) {
        issue(
          errors,
          chartPart,
          "invalid-data-label-position",
          `Office does not allow dLblPos="${posVal}" in ${typeNode.localName ?? ""}`,
        );
      }
    }

    const pending: {
      readonly order: number;
      readonly series: FileChartSeries;
      readonly formula: ChartFormulaReference;
      readonly labels?: FileChartSeriesDataLabels;
    }[] = [];
    for (const chartSeries of elements(typeNode, chartNamespace, "ser")) {
      const index = integerAttribute(
        first(chartSeries, chartNamespace, "idx"),
        "val",
      );
      const order = integerAttribute(
        first(chartSeries, chartNamespace, "order"),
        "val",
      );
      if (
        index === undefined ||
        index < 0 ||
        order === undefined ||
        order < 0
      ) {
        issue(
          errors,
          chartPart,
          "invalid-chart-series-index",
          "Chart series is missing a valid non-negative c:idx or c:order",
        );
        continue;
      }
      if (seenSeriesIndexes.has(index) || seenSeriesOrders.has(order)) {
        issue(
          errors,
          chartPart,
          "invalid-chart-series-index",
          `Chart series index ${index} or order ${order} is duplicated`,
        );
      }
      seenSeriesIndexes.add(index);
      seenSeriesOrders.add(order);
      const title = first(chartSeries, chartNamespace, "tx");
      const nameCache = chartStringCache(
        title === undefined
          ? undefined
          : first(title, chartNamespace, "strRef"),
        chartPart,
        errors,
      );
      const category = first(chartSeries, chartNamespace, "cat");
      const categoryCache = chartStringCache(
        category === undefined
          ? undefined
          : first(category, chartNamespace, "strRef"),
        chartPart,
        errors,
      );
      const value = first(chartSeries, chartNamespace, "val");
      const valueCache = chartNumberCache(
        value === undefined
          ? undefined
          : first(value, chartNamespace, "numRef"),
        chartPart,
        errors,
      );
      if (
        nameCache === undefined ||
        categoryCache === undefined ||
        valueCache === undefined
      ) {
        continue;
      }
      if (
        categoryCache.values.length !== valueCache.values.length ||
        nameCache.values.length !== 1
      ) {
        issue(
          errors,
          chartPart,
          "invalid-chart-series",
          "Chart series names, categories, and values counts do not match",
        );
      }
      const seriesLabelsNode = first(chartSeries, chartNamespace, "dLbls");
      const seriesLabels = readDataLabelSettings(
        seriesLabelsNode,
        typeNode.localName,
        chartPart,
        errors,
        textTheme,
      );
      const effectiveLabels = mergeFileDataLabels(
        groupDataLabels,
        seriesLabels,
      );
      const pointDataLabels: Record<number, false | FileChartDataLabel> = {};
      for (const pointLabel of elements(
        seriesLabelsNode ?? chartSeries,
        chartNamespace,
        "dLbl",
      )) {
        const pointIndex = integerAttribute(
          first(pointLabel, chartNamespace, "idx"),
          "val",
        );
        if (pointIndex === undefined || pointIndex < 0) continue;
        const pointSettings = readDataLabelSettings(
          pointLabel,
          typeNode.localName,
          chartPart,
          errors,
          textTheme,
        );
        if (pointSettings === undefined) continue;
        pointDataLabels[pointIndex] = pointSettings;
      }
      const name = nameCache.values[0] ?? "";
      const explosion = integerAttribute(
        first(chartSeries, chartNamespace, "explosion"),
        "val",
      );
      const pointExplosions: Record<number, number> = {};
      for (const point of elements(chartSeries, chartNamespace, "dPt")) {
        const pointIndex = integerAttribute(
          first(point, chartNamespace, "idx"),
          "val",
        );
        const pointExplosion = integerAttribute(
          first(point, chartNamespace, "explosion"),
          "val",
        );
        if (pointIndex !== undefined && pointExplosion !== undefined) {
          pointExplosions[pointIndex] = pointExplosion;
        }
      }
      const markerNode = first(chartSeries, chartNamespace, "marker");
      const smoothNode = first(chartSeries, chartNamespace, "smooth");
      if (
        typeNode.localName !== "lineChart" &&
        (markerNode !== undefined || smoothNode !== undefined)
      ) {
        issue(
          errors,
          chartPart,
          "invalid-chart-series-format",
          "Only line chart series support native markers and smoothing",
        );
      }
      let marker: FileChartSeries["marker"];
      if (markerNode !== undefined) {
        const symbol = attribute(
          first(markerNode, chartNamespace, "symbol") ?? markerNode,
          "val",
        );
        const sizeNode = first(markerNode, chartNamespace, "size");
        const size =
          sizeNode === undefined
            ? undefined
            : integerAttribute(sizeNode, "val");
        if (symbol === "") {
          issue(
            errors,
            chartPart,
            "invalid-chart-series-format",
            "Chart series marker is missing a symbol",
          );
        } else if (size !== undefined && (size < 2 || size > 72)) {
          issue(
            errors,
            chartPart,
            "invalid-chart-series-format",
            "Chart series marker size must be between 2 and 72 points",
          );
        } else if (symbol !== "none") {
          marker = {
            symbol,
            ...(size === undefined ? {} : { size }),
          };
        }
      }
      const smoothValue =
        smoothNode === undefined ? "" : attribute(smoothNode, "val");
      const smooth =
        smoothValue === "1" || smoothValue === "true"
          ? true
          : smoothValue === "0" || smoothValue === "false"
            ? false
            : undefined;
      if (smoothValue !== "" && smooth === undefined) {
        issue(
          errors,
          chartPart,
          "invalid-chart-series-format",
          "Chart series smooth must be a boolean value",
        );
      }
      pending.push({
        order,
        ...(effectiveLabels === undefined &&
        Object.keys(pointDataLabels).length === 0
          ? {}
          : {
              labels: {
                seriesIndex: index,
                ...(effectiveLabels === undefined
                  ? {}
                  : { dataLabels: effectiveLabels }),
                ...(Object.keys(pointDataLabels).length === 0
                  ? {}
                  : { pointDataLabels }),
              },
            }),
        series: {
          name,
          categories: categoryCache.values,
          values: valueCache.values,
          index,
          order,
          groupIndex,
          ...(explosion === undefined ? {} : { explosion }),
          ...(Object.keys(pointExplosions).length === 0
            ? {}
            : { pointExplosions }),
          ...(marker === undefined ? {} : { marker }),
          ...(smooth === true ? { smooth: true } : {}),
        },
        formula: {
          name: nameCache.formula,
          categories: categoryCache.formula,
          values: valueCache.formula,
        },
      });
    }
    pending.sort((left, right) => left.order - right.order);
    for (const item of pending) {
      series.push(item.series);
      if (item.labels !== undefined) seriesDataLabels.push(item.labels);
      formulaSeries.push({
        series: item.series,
        formula: item.formula,
      });
    }
    const groupDirection =
      typeNode.localName === "barChart"
        ? attribute(
            first(typeNode, chartNamespace, "barDir") ?? typeNode,
            "val",
          )
        : undefined;
    const firstSliceAngle = integerAttribute(
      first(typeNode, chartNamespace, "firstSliceAng"),
      "val",
    );
    const holeSize = integerAttribute(
      first(typeNode, chartNamespace, "holeSize"),
      "val",
    );
    if (
      firstSliceAngle !== undefined &&
      (firstSliceAngle < 0 || firstSliceAngle > 360)
    ) {
      issue(
        errors,
        chartPart,
        "invalid-chart-option",
        `Chart group firstSliceAng must be between 0 and 360`,
      );
    }
    if (holeSize !== undefined && (holeSize < 10 || holeSize > 90)) {
      issue(
        errors,
        chartPart,
        "invalid-chart-option",
        `Doughnut holeSize must be between 10 and 90`,
      );
    }
    groups.push({
      type: typeNode.localName ?? "",
      ...(groupDirection === "bar" || groupDirection === "col"
        ? { direction: groupDirection }
        : {}),
      ...(typeGrouping === "" ? {} : { grouping: typeGrouping }),
      axisIds,
      seriesIndices: pending.map(({ series: item }) => item.index),
      ...(firstSliceAngle === undefined ? {} : { firstSliceAngle }),
      ...(holeSize === undefined ? {} : { holeSize }),
    });
  }

  const externalData = first(chartXml, chartNamespace, "externalData");
  const workbookRelationshipId =
    externalData?.getAttributeNS(relationshipNamespace, "id") ?? "";
  const workbookRelationship = relationships
    .get(chartPart)
    ?.get(workbookRelationshipId);
  let embeddedWorkbook: FileChart["embeddedWorkbook"];
  if (
    workbookRelationship === undefined ||
    attribute(workbookRelationship, "Type") !==
      `${relationshipNamespace}/package` ||
    attribute(workbookRelationship, "TargetMode") === "External"
  ) {
    issue(
      errors,
      chartPart,
      "invalid-workbook-relationship",
      "Chart missing embedded workbook package relationship",
    );
  } else {
    const workbookPart = relationshipTargetPath(
      chartPart,
      attribute(workbookRelationship, "Target"),
    );
    const workbookBytes = files[workbookPart];
    if (
      workbookBytes === undefined ||
      contentTypeForPart(workbookPart, contentTypes) !==
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ) {
      issue(
        errors,
        workbookPart,
        "invalid-embedded-workbook",
        "Chart embedded workbook part or content type is invalid",
      );
    } else {
      embeddedWorkbook = readEmbeddedWorkbook(
        workbookPart,
        chartPart,
        workbookBytes,
        formulaSeries,
        errors,
      );
    }
  }
  const workbookSeries = embeddedWorkbook?.series ?? [];
  const chartSeriesById = new Map(series.map((item) => [item.index, item]));
  for (const workbookItem of workbookSeries) {
    const chartValue = chartSeriesById.get(workbookItem.index);
    if (
      chartValue !== undefined &&
      (workbookItem.name !== chartValue.name ||
        workbookItem.categories.some(
          (category, row) => category !== chartValue.categories[row],
        ) ||
        workbookItem.values.some(
          (value, row) => value !== chartValue.values[row],
        ))
    ) {
      issue(
        errors,
        chartPart,
        "chart-workbook-mismatch",
        `Chart series ${chartValue.name} cache does not match workbook`,
      );
    }
  }
  if (boxEmu === undefined) return undefined;
  const titleText =
    chart === undefined
      ? undefined
      : nativeChartTitleText(
          first(chart, chartNamespace, "title"),
          textTheme,
          chartPart,
          errors,
        );
  const displayBlanksAsNode =
    chart === undefined
      ? undefined
      : first(chart, chartNamespace, "dispBlanksAs");
  const displayBlanksAsValue =
    displayBlanksAsNode === undefined
      ? ""
      : attribute(displayBlanksAsNode, "val");
  const displayBlanksAs =
    displayBlanksAsValue === "gap" ||
    displayBlanksAsValue === "zero" ||
    displayBlanksAsValue === "span"
      ? displayBlanksAsValue
      : undefined;
  if (displayBlanksAsValue !== "" && displayBlanksAs === undefined) {
    issue(
      errors,
      chartPart,
      "invalid-chart-missing-values",
      "Chart displayBlanksAs must be gap, zero, or span",
    );
  }
  const chartObservation: FileChart = {
    types,
    ...(direction === undefined
      ? {}
      : { direction: direction as "bar" | "col" }),
    ...(grouping === undefined || grouping === "" ? {} : { grouping }),
    ...(titleText === undefined ? {} : { title: titleText }),
    ...(displayBlanksAs === undefined ? {} : { displayBlanksAs }),
    series,
    groups,
    axes,
    ...(seriesDataLabels.length === 0 ? {} : { seriesDataLabels }),
    ...(embeddedWorkbook === undefined ? {} : { embeddedWorkbook }),
  };
  return {
    id: identity.id,
    name: identity.name,
    kind: "chart",
    boxEmu,
    chart: chartObservation,
  };
}

function validateGroup(
  group: XmlElement,
  part: string,
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): FileObject | undefined {
  const identity = objectIdentity(group, "nvGrpSpPr", part, "group", errors);
  if (identity === undefined) return undefined;
  const properties = first(group, presentationNamespace, "grpSpPr");
  const transform =
    properties === undefined
      ? undefined
      : first(properties, drawingNamespace, "xfrm");
  const boxEmu = transformBox(transform, part, identity.name, "group", errors);
  const rotation = readRotationDegrees(transform);
  const childOffset =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "chOff");
  const childExtent =
    transform === undefined
      ? undefined
      : first(transform, drawingNamespace, "chExt");
  const childX = integerAttribute(childOffset, "x");
  const childY = integerAttribute(childOffset, "y");
  const childWidth = integerAttribute(childExtent, "cx");
  const childHeight = integerAttribute(childExtent, "cy");
  if (
    childX === undefined ||
    childY === undefined ||
    childWidth === undefined ||
    childHeight === undefined ||
    childWidth < 1 ||
    childHeight < 1 ||
    childWidth > maximumShapeExtentEmu ||
    childHeight > maximumShapeExtentEmu
  ) {
    issue(
      errors,
      part,
      "invalid-group-transform",
      `Group ${identity.name} missing valid member coordinate bounds`,
    );
  }
  const members = readObjects(
    group,
    part,
    files,
    documents,
    relationships,
    contentTypes,
    errors,
    theme,
  );
  if (members.length === 0) {
    issue(
      errors,
      part,
      "empty-group",
      `Group ${identity.name} has no native members`,
    );
  }
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "group",
    boxEmu,
    ...(childX === undefined || childY === undefined
      ? {}
      : {
          groupTransform: {
            childOffset: { x: childX, y: childY },
            childExtent: {
              width: childWidth ?? 0,
              height: childHeight ?? 0,
            },
            ...(transform !== undefined &&
            ["1", "true"].includes(attribute(transform, "flipH"))
              ? { flipH: true }
              : {}),
            ...(transform !== undefined &&
            ["1", "true"].includes(attribute(transform, "flipV"))
              ? { flipV: true }
              : {}),
          },
        }),
    ...(rotation === undefined ? {} : { rotationDegrees: rotation }),
    members,
  };
}

function tableSpan(
  cell: XmlElement,
  name: "gridSpan" | "rowSpan",
  part: string,
  errors: ValidationIssue[],
): number {
  if (!cell.hasAttribute(name)) return 1;
  const value = integerAttribute(cell, name);
  if (value === undefined || value < 1) {
    issue(
      errors,
      part,
      "invalid-table-merge",
      `${name} must be a positive integer`,
    );
    return 1;
  }
  return value;
}

function tableMergeFlag(
  cell: XmlElement,
  name: "hMerge" | "vMerge",
  part: string,
  errors: ValidationIssue[],
): boolean {
  const value = attribute(cell, name);
  if (value === "" || value === "0" || value === "false") return false;
  if (value === "1" || value === "true") return true;
  issue(errors, part, "invalid-table-merge", `${name} must be a boolean`);
  return false;
}

function tableCellText(
  cell: XmlElement,
  part: string,
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): FileText {
  const body = first(cell, drawingNamespace, "txBody");
  if (body === undefined) {
    issue(
      errors,
      part,
      "missing-table-cell-text",
      "Table cell missing a:txBody",
    );
    return { text: "", paragraphs: [] };
  }
  validateTextBodyValueRanges(body, part, errors);
  const paragraphs = elements(body, drawingNamespace, "p").map((paragraph) =>
    readParagraph(paragraph, theme, part, errors),
  );
  return { text: paragraphs.map(({ text }) => text).join("\n"), paragraphs };
}

/** Independently reads table grid and merge overlays without calling production table module calculation results. */
function validateTable(
  frame: XmlElement,
  part: string,
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): FileObject | undefined {
  const identity = objectIdentity(
    frame,
    "nvGraphicFramePr",
    part,
    "table",
    errors,
  );
  if (identity === undefined) return undefined;
  const transform = first(frame, presentationNamespace, "xfrm");
  const boxEmu = transformBox(transform, part, identity.name, "table", errors);
  const graphic = first(frame, drawingNamespace, "graphic");
  const graphicData =
    graphic === undefined
      ? undefined
      : first(graphic, drawingNamespace, "graphicData");
  const table =
    graphicData === undefined
      ? undefined
      : first(graphicData, drawingNamespace, "tbl");
  if (
    graphicData === undefined ||
    attribute(graphicData, "uri") !== tableUri ||
    table === undefined
  ) {
    issue(
      errors,
      part,
      "invalid-table-graphic",
      `Table ${identity.name} missing DrawingML table content`,
    );
    return undefined;
  }
  const grid = first(table, drawingNamespace, "tblGrid");
  const columns =
    grid === undefined ? [] : elements(grid, drawingNamespace, "gridCol");
  if (columns.length === 0) {
    issue(
      errors,
      part,
      "invalid-table-grid",
      `Table ${identity.name} missing column grid`,
    );
  }
  const columnWidthsEmu = columns.map((column, index) => {
    const width = integerAttribute(column, "w");
    if (width === undefined || width < 1 || width > maximumShapeExtentEmu) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `Table ${identity.name} column ${index + 1} width is invalid`,
      );
      return 0;
    }
    return width;
  });
  const rowNodes = elements(table, drawingNamespace, "tr");
  if (rowNodes.length === 0) {
    issue(
      errors,
      part,
      "invalid-table-grid",
      `Table ${identity.name} requires at least one row`,
    );
  }
  const rowHeightsEmu: number[] = [];
  const rowCells: XmlElement[][] = [];
  for (const [rowIndex, row] of rowNodes.entries()) {
    const height = integerAttribute(row, "h");
    if (height === undefined || height < 1 || height > maximumShapeExtentEmu) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `Table ${identity.name} row ${rowIndex + 1} height is invalid`,
      );
      rowHeightsEmu.push(0);
    } else rowHeightsEmu.push(height);
    const cells = elements(row, drawingNamespace, "tc");
    if (cells.length !== columns.length) {
      issue(
        errors,
        part,
        "invalid-table-grid",
        `Table ${identity.name} row ${rowIndex + 1} contains ${cells.length} cells, while column grid contains ${columns.length} columns`,
      );
    }
    rowCells.push(cells);
  }

  type Owner = {
    readonly row: number;
    readonly column: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  };
  const owners = new Map<string, Owner>();
  const continuations = new Map<
    XmlElement,
    { readonly horizontal: boolean; readonly vertical: boolean }
  >();
  const spans = new Map<
    XmlElement,
    { readonly rowSpan: number; readonly colSpan: number }
  >();
  const merges: {
    row: number;
    column: number;
    rowSpan: number;
    colSpan: number;
  }[] = [];
  const cells: string[][] = rowNodes.map(() =>
    Array.from({ length: columns.length }, () => ""),
  );
  const cellTexts: { row: number; column: number; text: FileText }[] = [];
  for (const [rowIndex, currentRow] of rowCells.entries()) {
    for (
      let column = 0;
      column < Math.min(currentRow.length, columns.length);
      column += 1
    ) {
      const cell = currentRow[column]!;
      const horizontalContinuation = tableMergeFlag(
        cell,
        "hMerge",
        part,
        errors,
      );
      const verticalContinuation = tableMergeFlag(cell, "vMerge", part, errors);
      continuations.set(cell, {
        horizontal: horizontalContinuation,
        vertical: verticalContinuation,
      });
      const colSpan = tableSpan(cell, "gridSpan", part, errors);
      const rowSpan = tableSpan(cell, "rowSpan", part, errors);
      spans.set(cell, { rowSpan, colSpan });
      if (horizontalContinuation || verticalContinuation) continue;
      if (
        column + colSpan > columns.length ||
        rowIndex + rowSpan > rowNodes.length
      ) {
        issue(
          errors,
          part,
          "invalid-table-merge",
          `Table ${identity.name} merge range out of bounds`,
        );
      }
      const owner = { row: rowIndex, column, rowSpan, colSpan };
      for (
        let row = rowIndex;
        row < Math.min(rowNodes.length, rowIndex + rowSpan);
        row += 1
      ) {
        for (
          let col = column;
          col < Math.min(columns.length, column + colSpan);
          col += 1
        ) {
          const key = `${row}:${col}`;
          if (owners.has(key))
            issue(
              errors,
              part,
              "conflicting-table-merge",
              `Table ${identity.name} merge ranges overlap`,
            );
          else owners.set(key, owner);
        }
      }
      if (rowSpan > 1 || colSpan > 1) merges.push(owner);
      const text = tableCellText(cell, part, errors, theme);
      cells[rowIndex]![column] = text.text;
      cellTexts.push({ row: rowIndex, column, text });
    }
  }
  for (const [rowIndex, currentRow] of rowCells.entries()) {
    for (let column = 0; column < columns.length; column += 1) {
      const cell = currentRow[column];
      const owner = owners.get(`${rowIndex}:${column}`);
      if (owner === undefined || cell === undefined) {
        issue(
          errors,
          part,
          "invalid-table-grid",
          `Table ${identity.name} row ${rowIndex + 1} column ${column + 1} has no valid cell`,
        );
        continue;
      }
      const isAnchor = rowIndex === owner.row && column === owner.column;
      const continuation = continuations.get(cell);
      const span = spans.get(cell);
      const horizontalContinuation = continuation?.horizontal ?? false;
      const verticalContinuation = continuation?.vertical ?? false;
      const expectedRowSpan = rowIndex === owner.row ? owner.rowSpan : 1;
      const expectedColSpan = column === owner.column ? owner.colSpan : 1;
      if (
        horizontalContinuation !== (!isAnchor && column > owner.column) ||
        verticalContinuation !== (!isAnchor && rowIndex > owner.row) ||
        span?.rowSpan !== expectedRowSpan ||
        span?.colSpan !== expectedColSpan
      ) {
        issue(
          errors,
          part,
          "invalid-table-merge",
          `Table ${identity.name} merge continuation flags or spans do not match anchor range`,
        );
      }
    }
  }
  // ISO/IEC 29500-1 defines row heights (21.1.3.18) and column widths
  // (21.1.3.2) independently from the graphic frame's displayed extents
  // (20.1.7.3). Native applications may normalize the frame without rewriting
  // the grid, so keep both raw observations and do not reject the combination.
  if (boxEmu === undefined) return undefined;
  return {
    id: identity.id,
    name: identity.name,
    kind: "table",
    boxEmu,
    table: {
      columnWidthsEmu,
      rowHeightsEmu,
      cells,
      merges,
      cellTexts,
    },
  };
}

function readObjects(
  parent: XmlNode,
  part: string,
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
  theme: ThemeTextReader,
): FileObject[] {
  const objects: FileObject[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1 || child.namespaceURI !== presentationNamespace) {
      continue;
    }
    let object: FileObject | undefined;
    switch (child.localName) {
      case "sp":
        object = validateShape(child as XmlElement, part, errors, theme);
        break;
      case "pic":
        object = validatePicture(
          child as XmlElement,
          part,
          files,
          relationships,
          contentTypes,
          errors,
        );
        break;
      case "graphicFrame": {
        const frame = child as XmlElement;
        const graphic = first(frame, drawingNamespace, "graphic");
        const graphicData =
          graphic === undefined
            ? undefined
            : first(graphic, drawingNamespace, "graphicData");
        const uri =
          graphicData === undefined ? undefined : attribute(graphicData, "uri");
        if (uri === chartNamespace) {
          object = validateChart(
            frame,
            part,
            files,
            documents,
            relationships,
            contentTypes,
            errors,
          );
        } else if (uri === tableUri) {
          object = validateTable(frame, part, errors, theme);
        }
        break;
      }
      case "grpSp":
        object = validateGroup(
          child as XmlElement,
          part,
          files,
          documents,
          relationships,
          contentTypes,
          errors,
          theme,
        );
        break;
    }
    if (object !== undefined) {
      const link = readObjectLink(
        child as XmlElement,
        object.id,
        part,
        documents,
        relationships,
        errors,
      );
      objects.push({ ...object, ...(link === undefined ? {} : { link }) });
    }
  }
  return objects;
}

function readObjectLink(
  object: XmlElement,
  objectId: number,
  part: string,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  errors: ValidationIssue[],
): FileObject["link"] {
  const nonVisual = descendantsByName(
    object,
    presentationNamespace,
    "cNvPr",
  ).find((candidate) => integerAttribute(candidate, "id") === objectId);
  const click =
    nonVisual === undefined
      ? undefined
      : descendantsByName(nonVisual, drawingNamespace, "hlinkClick")[0];
  if (click === undefined) return undefined;
  const relationshipId =
    click.getAttributeNS(relationshipNamespace, "id") ?? "";
  const action = attribute(click, "action");
  if (!click.hasAttributeNS(relationshipNamespace, "id")) {
    issue(
      errors,
      part,
      "invalid-object-link",
      "Object link missing relationship ID attribute",
    );
    return undefined;
  }
  if (relationshipId === "") {
    if (action === "ppaction://hlinkshowjump?jump=firstslide") {
      return { action: "home" };
    }
    issue(
      errors,
      part,
      "invalid-object-link",
      "Object link missing valid relationship or home action",
    );
    return undefined;
  }
  const relationship = relationships.get(part)?.get(relationshipId);
  if (relationship === undefined) {
    issue(
      errors,
      part,
      "missing-object-link-relationship",
      `Object link references non-existent relationship ${relationshipId}`,
    );
    return undefined;
  }
  const type = attribute(relationship, "Type");
  const targetMode = attribute(relationship, "TargetMode");
  if (type === `${relationshipNamespace}/hyperlink`) {
    if (targetMode !== "External" || action !== "") {
      issue(
        errors,
        part,
        "invalid-object-link",
        "URL link must use external hyperlink relationship",
      );
      return undefined;
    }
    const url = attribute(relationship, "Target");
    try {
      return { url: new URL(url).href };
    } catch {
      issue(
        errors,
        part,
        "invalid-object-link-url",
        "External link relationship target is not an absolute URL",
      );
      return undefined;
    }
  }
  if (type === `${relationshipNamespace}/slide`) {
    if (targetMode === "External" || action !== "ppaction://hlinksldjump") {
      issue(
        errors,
        part,
        "invalid-object-link",
        "Slide jump must use internal slide relationship and jump action",
      );
      return undefined;
    }
    const targetPart = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const targetSlide = documents.get(targetPart)?.documentElement;
    const commonSlideData =
      targetSlide == null
        ? undefined
        : first(targetSlide, presentationNamespace, "cSld");
    const slideId =
      commonSlideData === undefined ? "" : attribute(commonSlideData, "name");
    if (slideId === "") {
      issue(
        errors,
        part,
        "invalid-object-link-target",
        "Slide jump target missing stable slide identifier",
      );
      return undefined;
    }
    return { slideId };
  }
  issue(
    errors,
    part,
    "invalid-object-link",
    "Object link relationship type is not hyperlink or slide",
  );
  return undefined;
}

function readSlideBackground(
  slide: XmlElement,
  part: string,
  files: XmlFiles,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  theme: ThemeTextReader,
  errors: ValidationIssue[],
): FileSlideBackground | undefined {
  const commonData = first(slide, presentationNamespace, "cSld");
  const background =
    commonData === undefined
      ? undefined
      : first(commonData, presentationNamespace, "bg");
  if (background === undefined) return undefined;
  const properties = first(background, presentationNamespace, "bgPr");
  if (properties === undefined) {
    issue(
      errors,
      part,
      "invalid-slide-background",
      "Slide background missing bgPr",
    );
    return undefined;
  }
  const picture = first(properties, drawingNamespace, "blipFill");
  if (picture !== undefined) {
    const blip = first(picture, drawingNamespace, "blip");
    const relationshipId =
      blip?.getAttributeNS(relationshipNamespace, "embed") ?? "";
    const relationship =
      relationshipId === ""
        ? undefined
        : relationships.get(part)?.get(relationshipId);
    if (
      blip === undefined ||
      relationship === undefined ||
      attribute(relationship, "Type") !== `${relationshipNamespace}/image` ||
      attribute(relationship, "TargetMode") === "External"
    ) {
      issue(
        errors,
        part,
        "invalid-slide-background-image",
        "Slide image background missing valid internal image relationship",
      );
      return undefined;
    }
    const targetPart = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const bytes = files[targetPart];
    const contentType = contentTypeForPart(targetPart, contentTypes);
    if (
      bytes === undefined ||
      contentType === undefined ||
      !imageSignatureMatches(bytes, contentType)
    ) {
      issue(
        errors,
        part,
        "invalid-slide-background-image",
        "Slide image background target missing valid image content or content type",
      );
      return undefined;
    }
    return { type: "image", relationshipId, part: targetPart, contentType };
  }
  const gradient = first(properties, drawingNamespace, "gradFill");
  if (gradient !== undefined) {
    const stops = first(gradient, drawingNamespace, "gsLst");
    const values =
      stops === undefined ? [] : elements(stops, drawingNamespace, "gs");
    const parsedStops = values.map((stop) => {
      const position = integerAttribute(stop, "pos");
      if (position === undefined || position < 0 || position > 100_000) {
        issue(
          errors,
          part,
          "invalid-slide-background-gradient",
          "Gradient stop position is invalid",
        );
      }
      const color = readRunColor(stop, theme);
      return {
        position: (position ?? 0) / 100_000,
        ...(color === undefined ? {} : { color }),
      };
    });
    if (parsedStops.length < 2) {
      issue(
        errors,
        part,
        "invalid-slide-background-gradient",
        "Slide gradient requires at least two stops",
      );
    }
    const linear = first(gradient, drawingNamespace, "lin");
    const radial = first(gradient, drawingNamespace, "path");
    const angleValue =
      linear === undefined ? undefined : integerAttribute(linear, "ang");
    return {
      type: "gradient",
      kind: radial === undefined ? "linear" : "radial",
      stops: parsedStops,
      ...(angleValue === undefined
        ? {}
        : { angle: (((angleValue / 60_000 + 90) % 360) + 360) % 360 }),
    };
  }
  const solid = first(properties, drawingNamespace, "solidFill");
  if (solid !== undefined) {
    const color = readRunColor(solid, theme);
    return { type: "solid", ...(color === undefined ? {} : { color }) };
  }
  issue(
    errors,
    part,
    "invalid-slide-background",
    "Slide background missing supported fill",
  );
  return undefined;
}

function childElements(parent: XmlNode): XmlElement[] {
  const result: XmlElement[] = [];
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType === 1) result.push(child as XmlElement);
  }
  return result;
}

function extensionForNamespace(
  namespace: string | null,
): ValidationExtension | undefined {
  if (namespace === p14Namespace) return "p14";
  if (namespace === p159Namespace) return "p159";
  return undefined;
}

function transitionDuration(value: string): number | undefined {
  const match = /^(\d+(?:\.\d)?)(h|min|s|ms|µs|ns)?$/u.exec(value);
  if (match === null) return undefined;
  const amount = Number(match[1]);
  const multiplier =
    match[2] === "h"
      ? 3_600_000
      : match[2] === "min"
        ? 60_000
        : match[2] === "s"
          ? 1_000
          : match[2] === "µs"
            ? 0.001
            : match[2] === "ns"
              ? 0.000001
              : 1;
  const milliseconds = amount * multiplier;
  return Number.isFinite(milliseconds) ? milliseconds : undefined;
}

function readTransitionElement(
  transition: XmlElement,
  part: string,
  errors: ValidationIssue[],
  insideChoice: boolean,
): {
  readonly transition?: FileTransition;
  readonly extensions: ReadonlySet<ValidationExtension>;
} {
  const extensions = new Set<ValidationExtension>();
  const speedValue = attribute(transition, "spd");
  const speed =
    speedValue === "fast" || speedValue === "med" || speedValue === "slow"
      ? speedValue
      : undefined;
  if (speedValue !== "" && speed === undefined) {
    issue(
      errors,
      part,
      "invalid-transition-speed",
      `Unknown transition speed ${speedValue}`,
    );
  }
  const durationValue = transition.getAttributeNS(p14Namespace, "dur");
  let durationMs: number | undefined;
  if (transition.hasAttributeNS(p14Namespace, "dur")) {
    extensions.add("p14");
    durationMs = transitionDuration(durationValue ?? "");
    if (durationMs === undefined) {
      issue(
        errors,
        part,
        "invalid-transition-duration",
        "p14:dur is not a valid duration",
      );
    }
  }
  const effects = childElements(transition);
  if (effects.length === 0 && transition.hasAttribute("advTm")) {
    return { extensions };
  }
  if (effects.length !== 1) {
    issue(
      errors,
      part,
      "invalid-transition-effect",
      "Transition must contain exactly one effect element",
    );
  }
  const effect = effects[0];
  let kind: FileTransition["kind"] = "fade";
  let direction: string | undefined;
  let by: FileTransition["by"];
  if (effect?.namespaceURI === presentationNamespace) {
    if (effect.localName === "fade") kind = "fade";
    else if (effect.localName === "strips") {
      kind = "strips";
      direction = attribute(effect, "dir");
      if (!["lu", "ru", "ld", "rd"].includes(direction)) {
        issue(
          errors,
          part,
          "invalid-transition-direction",
          "strips.dir is invalid",
        );
      }
    } else if (effect.localName === "push") {
      kind = "push";
      direction = attribute(effect, "dir");
      if (!["l", "r", "u", "d"].includes(direction)) {
        issue(
          errors,
          part,
          "invalid-transition-direction",
          "push.dir is invalid",
        );
      }
    } else {
      issue(
        errors,
        part,
        "unsupported-transition-effect",
        `Unsupported transition effect p:${effect.localName}`,
      );
    }
  } else if (
    effect?.namespaceURI === p14Namespace &&
    effect.localName === "flip"
  ) {
    kind = "flip";
    extensions.add("p14");
    direction = attribute(effect, "dir");
    if (direction !== "l" && direction !== "r") {
      issue(
        errors,
        part,
        "invalid-transition-direction",
        "flip.dir must be l or r",
      );
    }
  } else if (
    effect?.namespaceURI === p159Namespace &&
    effect.localName === "morph"
  ) {
    kind = "morph";
    extensions.add("p159");
    const option = attribute(effect, "option");
    if (option === "byObject") by = "object";
    else if (option === "byWord") by = "word";
    else if (option === "byChar") by = "character";
    else
      issue(
        errors,
        part,
        "invalid-morph-option",
        `Unknown Morph option ${option}`,
      );
  } else if (effect !== undefined) {
    issue(
      errors,
      part,
      "unsupported-transition-effect",
      "Transition effect uses unknown namespace",
    );
  }
  if (!insideChoice && extensions.size > 0) {
    issue(
      errors,
      part,
      "extension-outside-choice",
      "p14/p159 transition mechanisms must be placed within mc:Choice",
    );
  }
  const result: FileTransition = {
    kind,
    ...(speed === undefined ? {} : { speed }),
    ...(direction === undefined ? {} : { direction }),
    ...(durationMs === undefined ? {} : { durationMs }),
    ...(by === undefined ? {} : { by }),
  };
  return { transition: result, extensions };
}

function readAlternateTransition(
  alternate: XmlElement,
  part: string,
  errors: ValidationIssue[],
): TransitionRead {
  const choices = elements(alternate, markupCompatibilityNamespace, "Choice");
  const fallbacks = elements(
    alternate,
    markupCompatibilityNamespace,
    "Fallback",
  );
  if (choices.length !== 1 || fallbacks.length !== 1) {
    issue(
      errors,
      part,
      "invalid-transition-alternate",
      "Transition compatibility branches must contain one Choice and one Fallback",
    );
  }
  const choice = choices[0];
  const fallback = fallbacks[0];
  const choiceTransition =
    choice === undefined
      ? undefined
      : elements(choice, presentationNamespace, "transition")[0];
  const fallbackTransition =
    fallback === undefined
      ? undefined
      : elements(fallback, presentationNamespace, "transition")[0];
  if (
    choice === undefined ||
    fallback === undefined ||
    choiceTransition === undefined ||
    fallbackTransition === undefined
  ) {
    issue(
      errors,
      part,
      "invalid-transition-alternate",
      "Both Choice and Fallback of transition compatibility branches must contain p:transition",
    );
    return { extensions: new Set() };
  }
  const choiceResult = readTransitionElement(
    choiceTransition,
    part,
    errors,
    true,
  );
  const fallbackResult = readTransitionElement(
    fallbackTransition,
    part,
    errors,
    false,
  );
  if (
    fallbackResult.transition !== undefined &&
    fallbackResult.transition.kind !== "fade" &&
    fallbackResult.transition.kind !== "strips" &&
    fallbackResult.transition.kind !== "push"
  ) {
    issue(
      errors,
      part,
      "invalid-transition-fallback",
      "Transition Fallback must use Transitional base effect",
    );
  }
  const required = new Set<ValidationExtension>();
  const requiredPrefixes = attribute(choice, "Requires")
    .trim()
    .split(/\s+/u)
    .filter((prefix) => prefix !== "");
  for (const prefix of requiredPrefixes) {
    const extension = extensionForNamespace(choice.lookupNamespaceURI(prefix));
    if (extension === undefined) {
      issue(
        errors,
        part,
        "unsupported-extension",
        `Transition Choice requires unsupported namespace ${prefix}`,
      );
    } else {
      required.add(extension);
    }
  }
  for (const extension of choiceResult.extensions) {
    if (!required.has(extension)) {
      issue(
        errors,
        part,
        "missing-extension-requirement",
        `Transition Choice uses ${extension}, but Requires does not declare this mechanism`,
      );
    }
  }
  for (const extension of required) {
    if (!choiceResult.extensions.has(extension)) {
      issue(
        errors,
        part,
        "unused-extension-requirement",
        `Transition Choice declares unused ${extension} mechanism`,
      );
    }
  }
  const extensions = new Set<ValidationExtension>([
    ...choiceResult.extensions,
    ...required,
    ...fallbackResult.extensions,
  ]);
  return {
    ...(choiceResult.transition === undefined
      ? {}
      : {
          transition: {
            ...choiceResult.transition,
            ...(fallbackResult.transition === undefined
              ? {}
              : { fallback: fallbackResult.transition }),
          },
        }),
    extensions,
  };
}

function readSlideTransition(
  slide: XmlElement,
  part: string,
  errors: ValidationIssue[],
): TransitionRead {
  const candidates = childElements(slide).filter(
    (child) =>
      (child.namespaceURI === presentationNamespace &&
        child.localName === "transition") ||
      (child.namespaceURI === markupCompatibilityNamespace &&
        child.localName === "AlternateContent"),
  );
  if (candidates.length > 1) {
    issue(
      errors,
      part,
      "duplicate-transition",
      "Slide contains multiple transition declarations",
    );
  }
  const candidate = candidates[0];
  if (candidate === undefined) return { extensions: new Set() };
  if (candidate.namespaceURI === markupCompatibilityNamespace) {
    return readAlternateTransition(candidate, part, errors);
  }
  const result = readTransitionElement(candidate, part, errors, false);
  return {
    ...(result.transition === undefined
      ? {}
      : { transition: result.transition }),
    extensions: result.extensions,
  };
}

function readAdvanceAfter(
  slide: XmlElement,
  part: string,
  errors: ValidationIssue[],
): number | undefined {
  const candidates = childElements(slide).filter(
    (child) =>
      (child.namespaceURI === presentationNamespace &&
        child.localName === "transition") ||
      (child.namespaceURI === markupCompatibilityNamespace &&
        child.localName === "AlternateContent"),
  );
  const transitions = candidates.flatMap((candidate) =>
    candidate.namespaceURI === presentationNamespace
      ? [candidate]
      : descendantsByName(candidate, presentationNamespace, "transition"),
  );
  const hasTiming = transitions.some((transition) =>
    transition.hasAttribute("advTm"),
  );
  if (!hasTiming) return undefined;
  const values = transitions.map((transition) =>
    transition.hasAttribute("advTm")
      ? nonNegativeIntegerValue(attribute(transition, "advTm"))
      : undefined,
  );
  if (
    values.some((value) => value === undefined) ||
    new Set(values).size !== 1 ||
    (values[0] ?? 0) > 4_294_967_295
  ) {
    issue(
      errors,
      part,
      "inconsistent-advance-after",
      "Automatic advance timing for transition compatibility branches must be identical non-negative integer milliseconds",
    );
    return undefined;
  }
  return values[0];
}

function descendantsByName(
  parent: XmlNode,
  namespace: string,
  localName: string,
): XmlElement[] {
  const result: XmlElement[] = [];
  const pending = [parent];
  while (pending.length > 0) {
    const current = pending.pop();
    if (current === undefined) continue;
    for (
      let child = current.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (child.nodeType !== 1) continue;
      const element = child as XmlElement;
      if (
        element.namespaceURI === namespace &&
        element.localName === localName
      ) {
        result.push(element);
      }
      pending.push(element);
    }
  }
  return result;
}

function positiveIntegerValue(value: string): number | undefined {
  if (!/^\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : undefined;
}

function nonNegativeIntegerValue(value: string): number | undefined {
  if (!/^\d+$/u.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

function collectObjectIds(objects: readonly FileObject[]): Set<number> {
  const result = new Set<number>();
  const pending = [...objects];
  while (pending.length > 0) {
    const object = pending.pop();
    if (object === undefined) continue;
    result.add(object.id);
    pending.push(...(object.members ?? []));
  }
  return result;
}

function findObjectById(
  objects: readonly FileObject[],
  id: number,
): FileObject | undefined {
  const pending = [...objects];
  while (pending.length > 0) {
    const object = pending.pop();
    if (object === undefined) continue;
    if (object.id === id) return object;
    pending.push(...(object.members ?? []));
  }
  return undefined;
}

type MotionPoint = { readonly x: number; readonly y: number };
type ObjectCoordinateTransform = {
  readonly offsetX: number;
  readonly offsetY: number;
  readonly scaleX: number;
  readonly scaleY: number;
};
type MotionTargetGeometry = {
  readonly objectCenter: MotionPoint;
  readonly containingGroupCenters: readonly MotionPoint[];
  readonly supported: boolean;
};

/** Maps member center to slide coordinates using actual group off/ext/chOff/chExt. */
function motionTargetGeometry(
  objects: readonly FileObject[],
  targetId: number,
  widthEmu: number,
  heightEmu: number,
): MotionTargetGeometry | undefined {
  if (widthEmu <= 0 || heightEmu <= 0) return undefined;
  const identity: ObjectCoordinateTransform = {
    offsetX: 0,
    offsetY: 0,
    scaleX: 1,
    scaleY: 1,
  };
  let found: MotionTargetGeometry | undefined;
  const visit = (
    candidates: readonly FileObject[],
    transform: ObjectCoordinateTransform,
    containingGroupCenters: readonly MotionPoint[],
    supported: boolean,
  ): void => {
    for (const object of candidates) {
      const box = object.boxEmu;
      const pageX = transform.offsetX + box.x * transform.scaleX;
      const pageY = transform.offsetY + box.y * transform.scaleY;
      const objectCenter = {
        x: (pageX + (box.width * transform.scaleX) / 2) / widthEmu,
        y: (pageY + (box.height * transform.scaleY) / 2) / heightEmu,
      };
      if (object.id === targetId) {
        found = {
          objectCenter,
          containingGroupCenters,
          supported,
        };
        return;
      }
      if (object.kind !== "group") continue;
      const groupCenters = [...containingGroupCenters, objectCenter];
      const groupTransform = object.groupTransform;
      const transformSupported =
        groupTransform !== undefined &&
        groupTransform.childExtent.width > 0 &&
        groupTransform.childExtent.height > 0 &&
        groupTransform.flipH !== true &&
        groupTransform.flipV !== true &&
        closeGeometryNumber((object.rotationDegrees ?? 0) % 360, 0);
      if (!transformSupported || groupTransform === undefined) {
        visit(object.members ?? [], transform, groupCenters, false);
        if (found !== undefined) return;
        continue;
      }
      const localScaleX = box.width / groupTransform.childExtent.width;
      const localScaleY = box.height / groupTransform.childExtent.height;
      const childTransform: ObjectCoordinateTransform = {
        offsetX:
          transform.offsetX +
          (box.x - groupTransform.childOffset.x * localScaleX) *
            transform.scaleX,
        offsetY:
          transform.offsetY +
          (box.y - groupTransform.childOffset.y * localScaleY) *
            transform.scaleY,
        scaleX: transform.scaleX * localScaleX,
        scaleY: transform.scaleY * localScaleY,
      };
      visit(object.members ?? [], childTransform, groupCenters, supported);
      if (found !== undefined) return;
    }
  };
  visit(objects, identity, [], true);
  return found;
}

type TimingTreeNode = {
  readonly wrapper: XmlElement;
  readonly timeNode: XmlElement;
  readonly timingId: number;
  readonly kind: "container" | "effect";
  readonly containerKind?: "parallel" | "sequence";
  readonly parentTimingId?: number;
  readonly children: number[];
};

type SlideAnimationRead = {
  readonly steps: readonly FileAnimationStep[];
  readonly containers: readonly FileAnimationContainer[];
};

function readSlideAnimations(
  slide: XmlElement,
  part: string,
  objects: readonly FileObject[],
  slideSize: { readonly widthEmu: number; readonly heightEmu: number },
  errors: ValidationIssue[],
): SlideAnimationRead | undefined {
  const timingNodes = elements(slide, presentationNamespace, "timing");
  if (timingNodes.length === 0) return undefined;
  if (timingNodes.length > 1) {
    issue(
      errors,
      part,
      "duplicate-timing",
      "Slide contains multiple timing declarations",
    );
  }
  const timing = timingNodes[0];
  if (timing === undefined) return undefined;

  const buildGroupKey = (targetId: number, groupId: number): string =>
    `${targetId}:${groupId}`;
  const timeIds = new Set<number>();
  const effectTimeIds = new Set<number>();
  const effectTargets = new Map<number, number>();
  // Native build groups pair an object with a non-negative group number; group numbers need not be timing ids.
  const effectGroups = new Set<string>();
  const effectsWithoutNaturalCompletion = new Set<number>();
  for (const timeNode of descendantsByName(
    timing,
    presentationNamespace,
    "cTn",
  )) {
    const id = positiveIntegerValue(attribute(timeNode, "id"));
    if (id === undefined) {
      issue(
        errors,
        part,
        "invalid-animation-id",
        "Animation cTn missing positive integer id",
      );
      continue;
    }
    if (timeIds.has(id)) {
      issue(
        errors,
        part,
        "duplicate-animation-id",
        `Duplicate animation timing node id ${id}`,
      );
    }
    timeIds.add(id);
    if (attribute(timeNode, "presetID") !== "") {
      effectTimeIds.add(id);
      if (attribute(timeNode, "repeatCount") === "indefinite") {
        effectsWithoutNaturalCompletion.add(id);
      }
      const targetIds = descendantsByName(
        timeNode,
        presentationNamespace,
        "spTgt",
      )
        .map((target) => positiveIntegerValue(attribute(target, "spid")))
        .filter((targetId): targetId is number => targetId !== undefined);
      const uniqueTargetIds = new Set(targetIds);
      if (uniqueTargetIds.size === 1 && targetIds.length > 0) {
        const [targetId] = uniqueTargetIds;
        if (targetId !== undefined) {
          effectTargets.set(id, targetId);
          const groupId = nonNegativeIntegerValue(
            attribute(timeNode, "grpId") || "0",
          );
          if (groupId !== undefined) {
            effectGroups.add(buildGroupKey(targetId, groupId));
          }
        }
      }
    }
  }

  const knownObjects = collectObjectIds(objects);
  const timingTreeNodes = new Map<number, TimingTreeNode>();
  const structuralWrappers = (parent: XmlNode): XmlElement[] => {
    const result: XmlElement[] = [];
    for (
      let child = parent.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (
        child.nodeType === 1 &&
        child.namespaceURI === presentationNamespace &&
        (child.localName === "par" || child.localName === "seq")
      ) {
        result.push(child as XmlElement);
      }
    }
    return result;
  };
  const visitTimingWrapper = (
    wrapper: XmlElement,
    parentTimingId?: number,
  ): number | undefined => {
    const timeNode = first(wrapper, presentationNamespace, "cTn");
    const timingId = positiveIntegerValue(
      timeNode === undefined ? "" : attribute(timeNode, "id"),
    );
    if (timeNode === undefined || timingId === undefined) {
      issue(
        errors,
        part,
        "invalid-animation-tree",
        "Animation container missing valid cTn",
      );
      return undefined;
    }
    const isEffect = attribute(timeNode, "presetID") !== "";
    const node: TimingTreeNode = {
      wrapper,
      timeNode,
      timingId,
      kind: isEffect ? "effect" : "container",
      ...(isEffect
        ? {}
        : {
            containerKind:
              wrapper.localName === "seq" ? "sequence" : "parallel",
          }),
      ...(parentTimingId === undefined ? {} : { parentTimingId }),
      children: [],
    };
    if (timingTreeNodes.has(timingId)) {
      issue(
        errors,
        part,
        "duplicate-animation-id",
        `Duplicate animation timing node id ${timingId}`,
      );
      return undefined;
    }
    timingTreeNodes.set(timingId, node);
    if (!isEffect) {
      const childList = first(timeNode, presentationNamespace, "childTnLst");
      for (const childWrapper of structuralWrappers(childList ?? timeNode)) {
        const childId = visitTimingWrapper(childWrapper, timingId);
        if (childId !== undefined) node.children.push(childId);
      }
    }
    return timingId;
  };
  const timingList = first(timing, presentationNamespace, "tnLst");
  for (const wrapper of structuralWrappers(timingList ?? timing)) {
    visitTimingWrapper(wrapper);
  }

  const containerEndSync = new Map<number, "all">();
  for (const node of timingTreeNodes.values()) {
    if (node.kind !== "container") continue;
    const endSync = first(node.timeNode, presentationNamespace, "endSync");
    if (endSync === undefined) continue;
    const returnNode = first(endSync, presentationNamespace, "rtn");
    if (
      attribute(endSync, "evt") !== "end" ||
      attribute(returnNode ?? node.timeNode, "val") !== "all"
    ) {
      issue(
        errors,
        part,
        "invalid-animation-end-sync",
        "Animation aggregation container must wait for end events of all members",
      );
      continue;
    }
    containerEndSync.set(node.timingId, "all");
  }

  const containersWithoutNaturalCompletion = new Set<number>();
  const containsUnendingEffect = (
    timingId: number,
    visited: Set<number>,
  ): boolean => {
    if (effectsWithoutNaturalCompletion.has(timingId)) return true;
    if (visited.has(timingId)) return false;
    visited.add(timingId);
    const node = timingTreeNodes.get(timingId);
    return (
      node?.children.some((childId) =>
        containsUnendingEffect(childId, visited),
      ) ?? false
    );
  };
  for (const node of timingTreeNodes.values()) {
    if (
      node.kind === "container" &&
      containsUnendingEffect(node.timingId, new Set())
    ) {
      containersWithoutNaturalCompletion.add(node.timingId);
    }
  }

  const dependencyByTimingId = new Map<number, FileAnimationDependency>();
  for (const node of timingTreeNodes.values()) {
    const conditions = first(node.timeNode, presentationNamespace, "stCondLst");
    const endConditions =
      conditions === undefined
        ? []
        : elements(conditions, presentationNamespace, "cond").filter(
            (condition) => attribute(condition, "evt") === "end",
          );
    if (endConditions.length === 0) continue;
    const condition = endConditions[0];
    const timingReference =
      condition === undefined
        ? undefined
        : first(condition, presentationNamespace, "tn");
    const sourceTimingId = positiveIntegerValue(
      timingReference === undefined ? "" : attribute(timingReference, "val"),
    );
    const offsetMs = nonNegativeIntegerValue(
      condition === undefined ? "0" : attribute(condition, "delay") || "0",
    );
    const source =
      sourceTimingId === undefined
        ? undefined
        : timingTreeNodes.get(sourceTimingId);
    const sourceTargetId =
      sourceTimingId === undefined
        ? undefined
        : effectTargets.get(sourceTimingId);
    if (
      endConditions.length !== 1 ||
      sourceTimingId === undefined ||
      offsetMs === undefined ||
      source === undefined ||
      sourceTimingId === node.timingId ||
      (source.kind === "container" && !containerEndSync.has(sourceTimingId))
    ) {
      issue(
        errors,
        part,
        "invalid-animation-end-dependency",
        "Animation end dependency must point to another naturally completable effect or all-members container",
      );
      continue;
    }
    if (
      effectsWithoutNaturalCompletion.has(sourceTimingId) ||
      containersWithoutNaturalCompletion.has(sourceTimingId)
    ) {
      issue(
        errors,
        part,
        "no-natural-completion",
        "Animation completion dependency cannot reference infinite effects or scopes without natural completion",
      );
      continue;
    }
    dependencyByTimingId.set(node.timingId, {
      timingId: sourceTimingId,
      point: "end",
      offsetMs,
      source: source.kind,
      ...(sourceTargetId === undefined ? {} : { targetId: sourceTargetId }),
    });
  }

  const timingEvent = (timingId: number, point: "start" | "end"): string =>
    `${timingId}:${point}`;
  const eventPrerequisites = new Map<string, string[]>();
  const requireBefore = (before: string, after: string): void => {
    const prerequisites = eventPrerequisites.get(after) ?? [];
    prerequisites.push(before);
    eventPrerequisites.set(after, prerequisites);
    if (!eventPrerequisites.has(before)) eventPrerequisites.set(before, []);
  };
  for (const node of timingTreeNodes.values()) {
    const startEvent = timingEvent(node.timingId, "start");
    const endEvent = timingEvent(node.timingId, "end");
    eventPrerequisites.set(
      startEvent,
      eventPrerequisites.get(startEvent) ?? [],
    );
    requireBefore(startEvent, endEvent);
    if (node.parentTimingId !== undefined) {
      // Child nodes inherit parent start prerequisites; parent completion does not imply child completion.
      requireBefore(timingEvent(node.parentTimingId, "start"), startEvent);
    }
    const startDependency = dependencyByTimingId.get(node.timingId);
    if (startDependency !== undefined) {
      requireBefore(timingEvent(startDependency.timingId, "end"), startEvent);
    }
    if (node.kind === "container" && containerEndSync.has(node.timingId)) {
      for (const childId of node.children) {
        requireBefore(timingEvent(childId, "end"), endEvent);
      }
    }
  }
  const completedEvents = new Set<string>();
  const visitingEvents = new Set<string>();
  let hasTimingCycle = false;
  const visitTimingEvent = (event: string): void => {
    if (completedEvents.has(event) || hasTimingCycle) return;
    if (visitingEvents.has(event)) {
      hasTimingCycle = true;
      return;
    }
    visitingEvents.add(event);
    for (const prerequisite of eventPrerequisites.get(event) ?? []) {
      visitTimingEvent(prerequisite);
    }
    visitingEvents.delete(event);
    completedEvents.add(event);
  };
  for (const event of eventPrerequisites.keys()) {
    visitTimingEvent(event);
  }
  if (hasTimingCycle) {
    issue(
      errors,
      part,
      "invalid-animation-end-dependency",
      "Animation completion dependency contains wait cycle and cannot end naturally",
    );
  }

  const containers: FileAnimationContainer[] = [...timingTreeNodes.values()]
    .filter(
      (
        node,
      ): node is TimingTreeNode & {
        readonly containerKind: "parallel" | "sequence";
      } => node.kind === "container" && node.containerKind !== undefined,
    )
    .map((node) => {
      const startAfter = dependencyByTimingId.get(node.timingId);
      const durationMs = integerAttribute(node.timeNode, "dur");
      return {
        timingId: node.timingId,
        kind: node.containerKind,
        ...(node.parentTimingId === undefined
          ? {}
          : { parentTimingId: node.parentTimingId }),
        children: node.children.flatMap((childId) => {
          const child = timingTreeNodes.get(childId);
          return child === undefined
            ? []
            : [{ kind: child.kind, timingId: childId }];
        }),
        ...(durationMs === undefined ? {} : { durationMs }),
        ...(containerEndSync.get(node.timingId) === undefined
          ? {}
          : { endSync: "all" as const }),
        ...(startAfter === undefined ? {} : { startAfter }),
      };
    });

  const buildList = first(timing, presentationNamespace, "bldLst");
  const backgroundBuildGroups = new Set<string>();
  const wholeObjectBuildGroups = new Set<string>();
  for (const node of timingTreeNodes.values()) {
    if (node.kind !== "effect") continue;
    const groupId = nonNegativeIntegerValue(
      attribute(node.timeNode, "grpId") || "0",
    );
    if (groupId === undefined) continue;
    // An effect can carry several behavior targets; classify every target so traversal order
    // cannot change the group's background role.
    for (const target of descendantsByName(
      node.timeNode,
      presentationNamespace,
      "spTgt",
    )) {
      const targetId = positiveIntegerValue(attribute(target, "spid"));
      if (targetId === undefined) continue;
      const key = buildGroupKey(targetId, groupId);
      if (first(target, presentationNamespace, "bg") !== undefined) {
        backgroundBuildGroups.add(key);
      } else if (
        first(target, presentationNamespace, "txEl") === undefined &&
        findObjectById(objects, targetId)?.text !== undefined
      ) {
        // A whole-object text effect keeps the shape fill and text in one native target and may
        // declare animBg without a separate spTgt/bg branch.
        wholeObjectBuildGroups.add(key);
      }
    }
  }
  const paragraphBuilds = new Map<
    string,
    { readonly build: string; readonly animBg?: boolean }
  >();
  const chartBuilds = new Map<string, "all" | "series" | "category">();
  const declaredChartBuildGroups = new Set<string>();
  if (buildList !== undefined) {
    for (const build of elements(buildList, presentationNamespace, "bldP")) {
      const targetId = positiveIntegerValue(attribute(build, "spid"));
      const groupText = attribute(build, "grpId");
      const groupId = nonNegativeIntegerValue(
        groupText === "" ? "0" : groupText,
      );
      if (targetId === undefined || groupId === undefined) {
        issue(
          errors,
          part,
          "invalid-paragraph-build-group",
          "Text build declaration requires valid native object and non-negative group identifier",
        );
        continue;
      }
      const key = buildGroupKey(targetId, groupId);
      if (paragraphBuilds.has(key)) {
        issue(
          errors,
          part,
          "duplicate-paragraph-build-group",
          `Duplicate build group ${groupId} for text object ${targetId}`,
        );
      }
      const animBgText = attribute(build, "animBg");
      paragraphBuilds.set(key, {
        build: attribute(build, "build"),
        ...(animBgText === "1" || animBgText === "true"
          ? { animBg: true }
          : animBgText === "0" || animBgText === "false"
            ? { animBg: false }
            : {}),
      });
    }
    for (const [key, build] of paragraphBuilds) {
      const hasBackgroundBehavior = backgroundBuildGroups.has(key);
      if (
        (build.animBg === true) !== hasBackgroundBehavior &&
        !(wholeObjectBuildGroups.has(key) && !hasBackgroundBehavior)
      ) {
        issue(
          errors,
          part,
          "inconsistent-animation-background-role",
          `animBg for text build group ${key} must match actual spTgt/bg behavior of the group`,
        );
      }
    }
    for (const key of backgroundBuildGroups) {
      if (!paragraphBuilds.has(key)) {
        issue(
          errors,
          part,
          "missing-animation-background-build",
          `Text build group ${key} with background behavior missing bldP declaration`,
        );
      }
    }
    for (const build of elements(
      buildList,
      presentationNamespace,
      "bldGraphic",
    )) {
      const targetId = positiveIntegerValue(attribute(build, "spid"));
      const groupText = attribute(build, "grpId");
      const groupId = nonNegativeIntegerValue(
        groupText === "" ? "0" : groupText,
      );
      if (targetId === undefined || groupId === undefined) {
        issue(
          errors,
          part,
          "invalid-chart-build-group",
          "Chart build declaration requires valid native object and non-negative group identifier",
        );
        continue;
      }
      const key = buildGroupKey(targetId, groupId);
      if (declaredChartBuildGroups.has(key)) {
        issue(
          errors,
          part,
          "duplicate-chart-build-group",
          `Duplicate build group ${groupId} for chart object ${targetId}`,
        );
        continue;
      }
      declaredChartBuildGroups.add(key);
      const targetObject = findObjectById(objects, targetId);
      if (targetObject === undefined) {
        issue(
          errors,
          part,
          "invalid-chart-build-group",
          `Chart build declaration target ${targetId} does not exist`,
        );
        continue;
      }
      if (targetObject.kind !== "chart") {
        issue(
          errors,
          part,
          "invalid-chart-build-group",
          `Chart build declaration target ${targetId} is not a chart object`,
        );
        continue;
      }
      if (!effectGroups.has(key)) {
        issue(
          errors,
          part,
          "invalid-chart-build-group",
          `Chart build declaration must associate with an actual chart animation effect of group ${groupId}`,
        );
        continue;
      }
      const subBuild = first(build, presentationNamespace, "bldSub");
      const chartBuild =
        subBuild === undefined
          ? undefined
          : first(subBuild, drawingNamespace, "bldChart");
      const mode =
        chartBuild === undefined
          ? first(build, presentationNamespace, "bldAsOne") === undefined
            ? undefined
            : "all"
          : attribute(chartBuild, "bld");
      if (mode === "all" || mode === "series" || mode === "category") {
        chartBuilds.set(key, mode);
      } else {
        issue(
          errors,
          part,
          "invalid-chart-build-group",
          "Chart build declaration is missing a supported native build mode",
        );
      }
    }
  }
  const mainSequence = descendantsByName(
    timing,
    presentationNamespace,
    "cTn",
  ).find((timeNode) => attribute(timeNode, "nodeType") === "mainSeq");
  const sequenceChildren =
    mainSequence === undefined
      ? undefined
      : first(mainSequence, presentationNamespace, "childTnLst");
  if (sequenceChildren === undefined) {
    issue(
      errors,
      part,
      "invalid-animation-tree",
      "Animation timing missing mainSeq click step list",
    );
    return { steps: [], containers };
  }

  type EffectEntry = {
    readonly effectNode: XmlElement;
    readonly effectTimeNode: XmlElement;
    readonly delayMs: number;
    readonly startAfter?: FileAnimationDependency;
    readonly order: number;
  };
  const effectEntries = (
    timingId: number,
    parent: {
      readonly delayMs: number;
      readonly startAfter?: FileAnimationDependency;
    },
    output: EffectEntry[],
  ): void => {
    const node = timingTreeNodes.get(timingId);
    if (node === undefined) return;
    const dependency = dependencyByTimingId.get(timingId);
    let delayMs = parent.delayMs;
    let startAfter = parent.startAfter;
    if (dependency !== undefined) {
      delayMs = dependency.offsetMs;
      startAfter = dependency;
    } else {
      const conditionList = first(
        node.timeNode,
        presentationNamespace,
        "stCondLst",
      );
      const ordinaryCondition =
        conditionList === undefined
          ? undefined
          : elements(conditionList, presentationNamespace, "cond").find(
              (condition) => attribute(condition, "evt") !== "end",
            );
      const delayText =
        ordinaryCondition === undefined
          ? "0"
          : attribute(ordinaryCondition, "delay") || "0";
      const localDelay =
        delayText === "indefinite" ? 0 : nonNegativeIntegerValue(delayText);
      if (localDelay === undefined) {
        issue(
          errors,
          part,
          "invalid-animation-delay",
          "Animation container delay must be non-negative integer milliseconds",
        );
        return;
      }
      delayMs += localDelay;
    }
    if (node.kind === "effect") {
      output.push({
        effectNode: node.wrapper,
        effectTimeNode: node.timeNode,
        delayMs,
        ...(startAfter === undefined ? {} : { startAfter }),
        order: output.length,
      });
      return;
    }
    for (const childId of node.children) {
      effectEntries(
        childId,
        { delayMs, ...(startAfter === undefined ? {} : { startAfter }) },
        output,
      );
    }
  };
  const nativeBehaviorsOf = (
    effectTimeNode: XmlElement,
  ): FileAnimationBehavior[] => {
    const children = first(effectTimeNode, presentationNamespace, "childTnLst");
    if (children === undefined) return [];
    const behaviorElements: XmlElement[] = [];
    for (
      let child = children.firstChild;
      child !== null;
      child = child.nextSibling
    ) {
      if (
        child.nodeType === 1 &&
        child.namespaceURI === presentationNamespace &&
        child.localName !== null &&
        ["set", "anim", "animEffect", "animRot", "animClr"].includes(
          child.localName,
        )
      ) {
        behaviorElements.push(child as XmlElement);
      }
    }
    const rangeOf = (
      target: XmlElement,
    ): FileAnimationTargetRange | undefined => {
      const textElement = first(target, presentationNamespace, "txEl");
      if (textElement !== undefined) {
        const range = first(textElement, presentationNamespace, "pRg");
        const start = nonNegativeIntegerValue(
          range === undefined ? "" : attribute(range, "st"),
        );
        const end = nonNegativeIntegerValue(
          range === undefined ? "" : attribute(range, "end"),
        );
        return start === undefined || end === undefined || start > end
          ? undefined
          : { type: "paragraph", start, end };
      }
      const graphicElement = first(target, presentationNamespace, "graphicEl");
      const chart =
        graphicElement === undefined
          ? undefined
          : first(graphicElement, drawingNamespace, "chart");
      if (chart === undefined) return undefined;
      const seriesIndex = integerAttribute(chart, "seriesIdx");
      const categoryIndex = integerAttribute(chart, "categoryIdx");
      const step = attribute(chart, "bldStep");
      if (
        seriesIndex === undefined ||
        categoryIndex === undefined ||
        (step !== "gridLegend" && step !== "series" && step !== "category")
      ) {
        return undefined;
      }
      return {
        type: "chart",
        step,
        seriesIndex,
        categoryIndex,
      };
    };
    const valueText = (wrapper: XmlElement | undefined): string | undefined => {
      if (wrapper === undefined) return undefined;
      const value =
        first(wrapper, presentationNamespace, "fltVal") ??
        first(wrapper, presentationNamespace, "strVal");
      const text = value === undefined ? "" : attribute(value, "val");
      return text === "" ? undefined : text;
    };
    return behaviorElements.flatMap((behaviorNode) => {
      const cBhvr = first(behaviorNode, presentationNamespace, "cBhvr");
      if (cBhvr === undefined) return [];
      const targets = descendantsByName(cBhvr, presentationNamespace, "spTgt");
      const target = targets[0];
      if (target === undefined) return [];
      const targetId = positiveIntegerValue(attribute(target, "spid"));
      if (targetId === undefined) return [];
      const targetText = first(target, presentationNamespace, "txEl");
      const targetGraphic = first(target, presentationNamespace, "graphicEl");
      const targetRange = rangeOf(target);
      const durationMs = integerAttribute(
        first(cBhvr, presentationNamespace, "cTn"),
        "dur",
      );
      const targetType =
        first(target, presentationNamespace, "bg") !== undefined
          ? "background"
          : targetText !== undefined
            ? "paragraph"
            : targetGraphic !== undefined
              ? "chart"
              : "object";
      const attributeList = first(cBhvr, presentationNamespace, "attrNameLst");
      const attributeNames =
        attributeList === undefined
          ? []
          : elements(attributeList, presentationNamespace, "attrName").map(
              ({ textContent }) => textContent ?? "",
            );
      const tavList = first(behaviorNode, presentationNamespace, "tavLst");
      const tavs =
        tavList === undefined
          ? []
          : elements(tavList, presentationNamespace, "tav");
      const keyframes = tavs.map((tav): FileAnimationKeyframe => {
        const time = nonNegativeIntegerValue(attribute(tav, "tm"));
        const formula = attribute(tav, "fmla");
        const valueNode = first(tav, presentationNamespace, "val");
        const typedValue =
          valueNode === undefined
            ? undefined
            : (first(valueNode, presentationNamespace, "fltVal") ??
              first(valueNode, presentationNamespace, "strVal"));
        const valueType = typedValue?.localName ?? undefined;
        const value =
          typedValue === undefined ? "" : attribute(typedValue, "val");
        return {
          ...(time === undefined ? {} : { time }),
          ...(formula === "" ? {} : { formula }),
          ...(value === "" ? {} : { value }),
          ...(valueType === undefined ? {} : { valueType }),
        };
      });
      const firstTav = tavs[0];
      const lastTav = tavs.at(-1);
      const fromValue =
        valueText(
          firstTav === undefined
            ? undefined
            : first(firstTav, presentationNamespace, "val"),
        ) ??
        (() => {
          const from = first(behaviorNode, presentationNamespace, "from");
          const color =
            from === undefined
              ? undefined
              : first(from, drawingNamespace, "srgbClr");
          return (
            attribute(behaviorNode, "from") ||
            (color === undefined ? "" : attribute(color, "val")) ||
            undefined
          );
        })();
      const toValue =
        valueText(
          lastTav === undefined
            ? undefined
            : first(lastTav, presentationNamespace, "val"),
        ) ??
        (() => {
          const to = first(behaviorNode, presentationNamespace, "to");
          const color =
            to === undefined
              ? undefined
              : first(to, drawingNamespace, "srgbClr");
          return (
            attribute(behaviorNode, "to") ||
            (color === undefined ? "" : attribute(color, "val")) ||
            valueText(to) ||
            undefined
          );
        })();
      const additive = attribute(cBhvr, "additive");
      const by = attribute(behaviorNode, "by");
      const filter = attribute(behaviorNode, "filter");
      return [
        {
          behavior: behaviorNode.localName as FileAnimationBehavior["behavior"],
          targetId,
          targetType,
          ...(targetRange === undefined ? {} : { targetRange }),
          attributeNames,
          ...(durationMs === undefined ? {} : { durationMs }),
          ...(fromValue === undefined ? {} : { from: fromValue }),
          ...(toValue === undefined ? {} : { to: toValue }),
          ...(by === "" ? {} : { by }),
          ...(attribute(behaviorNode, "calcmode") === ""
            ? {}
            : { calculationMode: attribute(behaviorNode, "calcmode") }),
          ...(tavList === undefined ? {} : { keyframes }),
          ...(additive === "" ? {} : { additive }),
          ...(filter === "" ? {} : { filter }),
        },
      ];
    });
  };

  const result: FileAnimationStep[] = [];
  for (const stepNode of elements(
    sequenceChildren,
    presentationNamespace,
    "par",
  )) {
    const stepTimeNode = first(stepNode, presentationNamespace, "cTn");
    if (stepTimeNode === undefined) {
      issue(errors, part, "invalid-animation-step", "Click step missing cTn");
      continue;
    }
    const stepConditions = first(
      stepTimeNode,
      presentationNamespace,
      "stCondLst",
    );
    const isAutomatic =
      stepConditions !== undefined &&
      descendantsByName(stepConditions, presentationNamespace, "cond").some(
        (condition) => attribute(condition, "evt") === "onBegin",
      );
    const trigger: FileAnimationStep["trigger"] = isAutomatic
      ? "auto"
      : "click";
    const effects: FileAnimationEffect[] = [];
    const entries: EffectEntry[] = [];
    const stepId = positiveIntegerValue(attribute(stepTimeNode, "id"));
    const stepTimingNode =
      stepId === undefined ? undefined : timingTreeNodes.get(stepId);
    for (const childId of stepTimingNode?.children ?? []) {
      effectEntries(childId, { delayMs: 0 }, entries);
    }
    entries.sort(
      (left, right) =>
        (left.startAfter === undefined
          ? left.delayMs
          : Number.MAX_SAFE_INTEGER) -
          (right.startAfter === undefined
            ? right.delayMs
            : Number.MAX_SAFE_INTEGER) || left.order - right.order,
    );
    for (const entry of entries) {
      const { effectTimeNode } = entry;
      const presetId = integerAttribute(effectTimeNode, "presetID");
      if (presetId === undefined) continue;
      const presetClass = attribute(effectTimeNode, "presetClass");
      const colorAnimationNodes = descendantsByName(
        effectTimeNode,
        presentationNamespace,
        "animClr",
      );
      const presetById: Readonly<
        Record<number, FileAnimationEffect["preset"]>
      > = {
        2: "fly",
        8: "spin",
        9: "fillColor",
        10: "fade",
        17: "stretch",
        22: "wipe",
        42: "floatUp",
        47: "floatDown",
        49: "spinIn",
        53: "zoom",
      };
      const preset =
        presetById[presetId] ??
        (presetId === 0 && colorAnimationNodes.length === 1
          ? "breathe"
          : undefined);
      const kind =
        presetClass === "entr"
          ? "entrance"
          : presetClass === "emph"
            ? "emphasis"
            : undefined;
      if (preset === undefined || kind === undefined) {
        issue(
          errors,
          part,
          "unsupported-animation",
          `Unsupported animation preset ${presetClass}#${presetId}`,
        );
        continue;
      }
      const expectedKind =
        preset === "spin" || preset === "breathe" || preset === "fillColor"
          ? "emphasis"
          : "entrance";
      if (kind !== expectedKind) {
        issue(
          errors,
          part,
          "invalid-animation-kind",
          `presetClass for ${preset} must be ${expectedKind === "emphasis" ? "emph" : "entr"}`,
        );
        continue;
      }
      const effectConditionList = first(
        effectTimeNode,
        presentationNamespace,
        "stCondLst",
      );
      const effectCondition =
        effectConditionList === undefined
          ? undefined
          : first(effectConditionList, presentationNamespace, "cond");
      const effectDelay = nonNegativeIntegerValue(
        effectCondition === undefined
          ? "0"
          : attribute(effectCondition, "delay") || "0",
      );
      if (effectDelay === undefined) {
        issue(
          errors,
          part,
          "invalid-animation-delay",
          "Animation effect delay must be non-negative integer milliseconds",
        );
        continue;
      }
      const repeatText = attribute(effectTimeNode, "repeatCount");
      let repeatCount: number | "forever" | undefined;
      if (repeatText === "indefinite") {
        repeatCount = "forever";
      } else if (repeatText !== "") {
        const nativeRepeatCount = nonNegativeIntegerValue(repeatText);
        if (
          nativeRepeatCount === undefined ||
          nativeRepeatCount > maximumTimeNodeValue ||
          nativeRepeatCount < 1000 ||
          nativeRepeatCount % 1000 !== 0
        ) {
          issue(
            errors,
            part,
            "invalid-animation-repeat",
            "Animation repeatCount must be a whole positive integer number of cycles in thousandths",
          );
          continue;
        }
        repeatCount = nativeRepeatCount / 1000;
      }
      const autoReverse = ["1", "true"].includes(
        attribute(effectTimeNode, "autoRev"),
      );
      const behaviorNodes = descendantsByName(
        effectTimeNode,
        presentationNamespace,
        "cBhvr",
      );
      const durations = behaviorNodes
        .map((behavior) =>
          integerAttribute(
            first(behavior, presentationNamespace, "cTn"),
            "dur",
          ),
        )
        .filter((duration): duration is number => duration !== undefined);
      const traversalDurationMs =
        durations.length === 0 ? undefined : Math.max(...durations);
      if (
        traversalDurationMs === undefined ||
        traversalDurationMs <= 0 ||
        (preset === "breathe" &&
          (!autoReverse || traversalDurationMs > maximumTimeNodeValue))
      ) {
        issue(
          errors,
          part,
          preset === "breathe"
            ? "invalid-breathe-cycle"
            : "invalid-animation-duration",
          preset === "breathe"
            ? "Breathe cycle must consist of forward color animation and native reverse"
            : "Animation behavior missing positive integer duration",
        );
        continue;
      }
      const durationMs = traversalDurationMs * (preset === "breathe" ? 2 : 1);
      const targetNodes = descendantsByName(
        effectTimeNode,
        presentationNamespace,
        "spTgt",
      );
      const targetIds = targetNodes.map((target) =>
        positiveIntegerValue(attribute(target, "spid")),
      );
      const validTargetIds = targetIds.filter(
        (targetId): targetId is number => targetId !== undefined,
      );
      if (
        targetNodes.length === 0 ||
        validTargetIds.length !== targetNodes.length ||
        new Set(validTargetIds).size !== 1
      ) {
        issue(
          errors,
          part,
          "invalid-animation-target",
          "An animation effect must target a single native object",
        );
        continue;
      }
      const targetId = validTargetIds[0];
      if (targetId === undefined || !knownObjects.has(targetId)) {
        issue(
          errors,
          part,
          "missing-animation-target",
          `Animation target ${String(targetId)} does not exist`,
        );
        continue;
      }
      let invalidRange = false;
      const ranges: (FileAnimationTargetRange | undefined)[] = targetNodes.map(
        (target) => {
          const textElement = first(target, presentationNamespace, "txEl");
          if (textElement !== undefined) {
            const range = first(textElement, presentationNamespace, "pRg");
            const start = nonNegativeIntegerValue(
              range === undefined ? "" : attribute(range, "st"),
            );
            const end = nonNegativeIntegerValue(
              range === undefined ? "" : attribute(range, "end"),
            );
            if (start === undefined || end === undefined || start > end) {
              issue(
                errors,
                part,
                "invalid-animation-target-range",
                "Text animation paragraph range must be a valid non-negative closed interval",
              );
              invalidRange = true;
              return undefined;
            }
            return { type: "paragraph" as const, start, end };
          }
          const graphicElement = first(
            target,
            presentationNamespace,
            "graphicEl",
          );
          const chart =
            graphicElement === undefined
              ? undefined
              : first(graphicElement, drawingNamespace, "chart");
          if (chart !== undefined) {
            const seriesIndex = integerAttribute(chart, "seriesIdx");
            const categoryIndex = integerAttribute(chart, "categoryIdx");
            const step = attribute(chart, "bldStep");
            if (
              seriesIndex === undefined ||
              categoryIndex === undefined ||
              (step !== "gridLegend" &&
                step !== "series" &&
                step !== "category")
            ) {
              issue(
                errors,
                part,
                "invalid-animation-target-range",
                "Chart animation range missing valid series, category, or expand step",
              );
              invalidRange = true;
              return undefined;
            }
            return {
              type: "chart" as const,
              step: step,
              seriesIndex,
              categoryIndex,
            };
          }
          return undefined;
        },
      );
      if (invalidRange) continue;
      const rangeSignatures = new Set(
        ranges.map((range) => JSON.stringify(range)),
      );
      const targetObject = findObjectById(objects, targetId);
      const nativeBuildGroupId = nonNegativeIntegerValue(
        attribute(effectTimeNode, "grpId") || "0",
      );
      const effectBuildGroup =
        nativeBuildGroupId === undefined
          ? undefined
          : buildGroupKey(targetId, nativeBuildGroupId);
      const paragraphBuild =
        effectBuildGroup === undefined
          ? undefined
          : paragraphBuilds.get(effectBuildGroup);
      const nativeBehaviors = nativeBehaviorsOf(effectTimeNode);
      if (
        nativeBehaviors.some(({ behavior, durationMs: behaviorDuration }) =>
          behavior === "set"
            ? behaviorDuration !== 1
            : behaviorDuration !== traversalDurationMs,
        )
      ) {
        issue(
          errors,
          part,
          "inconsistent-animation-behavior-duration",
          "Native behaviors of the same effect must retain identical duration, with visibility initialization fixed at 1ms",
        );
        continue;
      }
      const paragraphRanges = nativeBehaviors
        .map(({ targetType, targetRange }) =>
          targetType === "paragraph" && targetRange?.type === "paragraph"
            ? targetRange
            : undefined,
        )
        .filter(
          (
            range,
          ): range is Extract<
            FileAnimationTargetRange,
            { type: "paragraph" }
          > => range !== undefined,
        );
      const paragraphIndexes = new Set(
        paragraphRanges.map(({ start }) => start),
      );
      // build=p with grpId=1 represents full paragraph projection; must be validated even if all pRg are omitted.
      const requiresFullParagraphProjection =
        nativeBuildGroupId === 1 &&
        paragraphBuild?.build === "p" &&
        targetObject?.text !== undefined &&
        targetObject.text.paragraphs.length > 0;
      const fullParagraphProjection =
        nativeBuildGroupId === 1 &&
        (paragraphBuild?.build === "allAtOnce" ||
          paragraphBuild?.build === "p") &&
        targetObject?.text !== undefined &&
        paragraphRanges.length > 0 &&
        ranges.every(
          (range) => range === undefined || range.type === "paragraph",
        ) &&
        paragraphRanges.every(({ start, end }) => start === end) &&
        paragraphIndexes.size === targetObject.text.paragraphs.length &&
        targetObject.text.paragraphs.every((_, index) =>
          paragraphIndexes.has(index),
        );
      if (rangeSignatures.size > 1 && !fullParagraphProjection) {
        issue(
          errors,
          part,
          "inconsistent-animation-target-range",
          "Behavior nodes of the same animation must use consistent ranges; whole text projection must cover all paragraphs",
        );
        continue;
      }
      if (
        !fullParagraphProjection &&
        ranges.some((range) => range?.type === "paragraph") &&
        (nativeBuildGroupId === undefined || paragraphBuild?.build !== "p")
      ) {
        issue(
          errors,
          part,
          "invalid-animation-target-range",
          "Paragraph-by-paragraph text targets must associate with native build declaration of the same group with build=p",
        );
        continue;
      }
      if (
        (fullParagraphProjection || requiresFullParagraphProjection) &&
        !validWholeTextProjection(
          preset,
          kind,
          targetObject?.text?.paragraphs.length ?? 0,
          paragraphBuild?.build ?? "",
          paragraphBuild?.animBg === true,
          nativeBehaviors,
        )
      ) {
        issue(
          errors,
          part,
          "invalid-animation-projection-target",
          "Initialization, paragraph properties, or background role of whole text projection do not match corresponding preset",
        );
        continue;
      }
      if (
        (preset === "spin" || preset === "spinIn") &&
        nativeBehaviors.some(isPositionBehavior) &&
        (!hasCompletePositionAxes(nativeBehaviors) ||
          !hasConsistentPositionFormulas(nativeBehaviors) ||
          nativeBehaviors
            .filter(isPositionBehavior)
            .some((behavior) => !validCenterRotationProgress(behavior)) ||
          !validCenterRotationGeometry(
            preset,
            nativeBehaviors,
            objects,
            slideSize.widthEmu,
            slideSize.heightEmu,
          ))
      ) {
        issue(
          errors,
          part,
          "invalid-animation-progress-formula",
          "Position formula for common-center spin is unsupported or does not match actual geometry, angle, scale, and progress parameters",
        );
      }
      const targetRange = fullParagraphProjection ? undefined : ranges[0];
      let fromColor: string | undefined;
      let toColor: string | undefined;
      if (preset === "breathe" || preset === "fillColor") {
        const colorAnimation = colorAnimationNodes[0];
        const colorBehavior =
          colorAnimation === undefined
            ? undefined
            : first(colorAnimation, presentationNamespace, "cBhvr");
        const colorAttributeList =
          colorBehavior === undefined
            ? undefined
            : first(colorBehavior, presentationNamespace, "attrNameLst");
        const colorAttribute =
          colorAttributeList === undefined
            ? undefined
            : first(colorAttributeList, presentationNamespace, "attrName");
        const colorTargetIds =
          colorBehavior === undefined
            ? []
            : descendantsByName(
                colorBehavior,
                presentationNamespace,
                "spTgt",
              ).map((target) =>
                positiveIntegerValue(attribute(target, "spid")),
              );
        const colorValue = (name: "from" | "to"): string | undefined => {
          if (colorAnimation === undefined) return undefined;
          const wrapper = first(colorAnimation, presentationNamespace, name);
          const rgb =
            wrapper === undefined
              ? undefined
              : first(wrapper, drawingNamespace, "srgbClr");
          const hex = rgb === undefined ? "" : attribute(rgb, "val");
          return /^[0-9a-f]{6}$/iu.test(hex)
            ? `#${hex.toUpperCase()}`
            : undefined;
        };
        fromColor = colorValue("from");
        toColor = colorValue("to");
        if (
          colorAnimation === undefined ||
          attribute(colorAnimation, "clrSpc") !== "rgb" ||
          colorAttribute?.textContent !== "fillcolor" ||
          colorTargetIds.length !== 1 ||
          colorTargetIds[0] !== targetId ||
          fromColor === undefined ||
          toColor === undefined ||
          targetObject?.kind !== "shape" ||
          (preset === "fillColor" &&
            (autoReverse ||
              repeatCount !== undefined ||
              attribute(effectTimeNode, "fill") !== "hold" ||
              colorBehavior === undefined ||
              childElements(colorBehavior)
                .filter((node) => node.localName === "cTn")
                .some((node) => attribute(node, "fill") !== "hold")))
        ) {
          issue(
            errors,
            part,
            preset === "breathe"
              ? "invalid-breathe-animation"
              : "invalid-fill-color-animation",
            preset === "breathe"
              ? "Breathe effect must be forward color animation on Shape fill with native reverse"
              : "Fill color emphasis must retain final state and not reverse or repeat",
          );
          continue;
        }
      }
      const chartSeriesCount = targetObject?.chart?.series.length ?? 0;
      // Independent rings may have different point counts; keep every real slot addressable.
      const chartCategoryCount = Math.max(
        0,
        ...(targetObject?.chart?.series.map(
          (series) => series.categories.length,
        ) ?? []),
      );
      if (
        targetRange?.type === "paragraph" &&
        (targetObject?.text === undefined ||
          paragraphBuild?.build !== "p" ||
          targetRange.end >= targetObject.text.paragraphs.length)
      ) {
        issue(
          errors,
          part,
          "invalid-animation-target-range",
          "Text animation paragraph range exceeds target native text object",
        );
        continue;
      }
      if (targetRange?.type === "chart") {
        if (targetObject?.kind !== "chart") {
          issue(
            errors,
            part,
            "invalid-animation-target-range",
            "Chart animation range targets a non-chart native object",
          );
          continue;
        }
        if (
          effectBuildGroup === undefined ||
          !declaredChartBuildGroups.has(effectBuildGroup)
        ) {
          issue(
            errors,
            part,
            "invalid-chart-build-group",
            "Chart animation effect must associate with a build declaration of the same chart group",
          );
          continue;
        }
        const chartBuild = chartBuilds.get(effectBuildGroup);
        if (chartBuild === undefined) {
          // The declaration was invalid and already diagnosed.
          continue;
        }
        if (targetRange.step === "gridLegend") {
          if (chartBuild === "all") {
            issue(
              errors,
              part,
              "invalid-animation-target-range",
              "Grid and legend animation requires a series or category chart build",
            );
            continue;
          }
          if (
            targetRange.seriesIndex !== -3 ||
            targetRange.categoryIndex !== -3
          ) {
            issue(
              errors,
              part,
              "invalid-animation-target-range",
              "Chart gridLegend animation must use native -3 sentinel indices",
            );
            continue;
          }
        } else if (chartBuild !== targetRange.step) {
          issue(
            errors,
            part,
            "invalid-animation-target-range",
            "Chart animation step does not match native chart build mode",
          );
          continue;
        } else if (
          targetRange.step === "series" &&
          (targetRange.seriesIndex < 0 ||
            targetRange.categoryIndex !== -4 ||
            targetRange.seriesIndex >= chartSeriesCount)
        ) {
          issue(
            errors,
            part,
            "invalid-animation-target-range",
            "Chart series animation range exceeds series in native chart",
          );
          continue;
        } else if (
          targetRange.step === "category" &&
          (targetRange.seriesIndex !== -4 ||
            targetRange.categoryIndex < 0 ||
            targetRange.categoryIndex >= chartCategoryCount)
        ) {
          issue(
            errors,
            part,
            "invalid-animation-target-range",
            "Chart category animation range exceeds categories in native chart",
          );
          continue;
        }
      }
      const iterateNodes = elements(
        effectTimeNode,
        presentationNamespace,
        "iterate",
      );
      if (iterateNodes.length > 1) {
        issue(
          errors,
          part,
          "invalid-animation-iteration",
          "An animation effect can contain only one native iterate declaration",
        );
        continue;
      }
      const iterate = iterateNodes[0];
      let by: "word" | "letter" | number | undefined;
      let intervalMs: number | undefined;
      if (iterate !== undefined) {
        const type = attribute(iterate, "type");
        const interval = first(iterate, presentationNamespace, "tmAbs");
        const value = nonNegativeIntegerValue(
          interval === undefined ? "" : attribute(interval, "val"),
        );
        if (
          (type !== "wd" && type !== "lt") ||
          value === undefined ||
          value < 1 ||
          targetRange !== undefined ||
          targetObject?.text === undefined ||
          effectBuildGroup === undefined ||
          !paragraphBuilds.has(effectBuildGroup)
        ) {
          issue(
            errors,
            part,
            "invalid-animation-iteration",
            "Text iterate must be word-by-word or letter-by-letter with positive integer absolute interval targeting native text object",
          );
          continue;
        }
        by = type === "wd" ? "word" : "letter";
        intervalMs = value;
      }
      const effectConditions =
        effectConditionList === undefined
          ? []
          : elements(effectConditionList, presentationNamespace, "cond");
      const endConditions = effectConditions.filter(
        (condition) => attribute(condition, "evt") === "end",
      );
      const timingId = positiveIntegerValue(attribute(effectTimeNode, "id"));
      if (timingId === undefined) continue;
      const startAfter = dependencyByTimingId.get(timingId) ?? entry.startAfter;
      if (endConditions.length > 1) continue;
      const easing: { accelerate?: number; decelerate?: number } = {};
      const accelerate = integerAttribute(effectTimeNode, "accel");
      const decelerate = integerAttribute(effectTimeNode, "decel");
      if (accelerate !== undefined && accelerate !== 0) {
        easing.accelerate = accelerate / 100000;
      }
      if (decelerate !== undefined && decelerate !== 0) {
        easing.decelerate = decelerate / 100000;
      }
      const directionSubtype = integerAttribute(
        effectTimeNode,
        "presetSubtype",
      );
      const direction =
        preset === "wipe" || preset === "fly"
          ? ({ 1: "top", 2: "right", 4: "bottom", 8: "left" } as const)[
              directionSubtype ?? -1
            ]
          : undefined;
      if ((preset === "wipe" || preset === "fly") && direction === undefined) {
        issue(
          errors,
          part,
          "invalid-animation-direction",
          `${preset} missing valid direction`,
        );
        continue;
      }
      const rotation = descendantsByName(
        effectTimeNode,
        presentationNamespace,
        "animRot",
      )[0];
      const rotationFrom =
        rotation === undefined ? undefined : decimalAttribute(rotation, "from");
      const rotationByAttribute =
        rotation === undefined ? "" : attribute(rotation, "by");
      const rotationBy =
        rotationByAttribute === ""
          ? undefined
          : decimalAttribute(rotation, "by");
      const rotationByIsValid =
        rotationBy !== undefined &&
        Number.isSafeInteger(rotationBy) &&
        rotationBy >= -2_147_483_648 &&
        rotationBy <= 2_147_483_647;
      if (
        preset === "spin" &&
        ((rotationFrom === undefined && !rotationByIsValid) ||
          (rotationFrom !== undefined && rotationByAttribute !== ""))
      ) {
        issue(
          errors,
          part,
          "invalid-animation-rotation",
          "spin must declare valid starting angle or relative angle, and only one of them",
        );
        continue;
      }
      if (preset === "spin" && rotationBy !== undefined && rotationByIsValid) {
        if (by !== undefined) {
          issue(
            errors,
            part,
            "invalid-animation-iteration",
            "Relative rotation for spin cannot combine with text iterate",
          );
          continue;
        }
        by = rotationBy / 60000;
      }
      const from =
        direction === undefined
          ? rotationFrom === undefined
            ? undefined
            : rotationFrom / 60000
          : direction;
      const chartBuild =
        effectBuildGroup === undefined
          ? undefined
          : chartBuilds.get(effectBuildGroup);
      const nativeBuild =
        paragraphBuild?.build === "p"
          ? "paragraph"
          : paragraphBuild?.build === "allAtOnce"
            ? "allAtOnce"
            : chartBuild === "series" || chartBuild === "category"
              ? chartBuild
              : undefined;
      effects.push({
        targetId,
        timingId,
        ...(targetRange === undefined ? {} : { targetRange }),
        ...(nativeBuild === undefined ? {} : { nativeBuild }),
        ...(paragraphBuild === undefined || nativeBuildGroupId === undefined
          ? {}
          : {
              nativeBuildGroupId,
              ...(paragraphBuild.animBg === undefined
                ? {}
                : { nativeBuildAnimBg: paragraphBuild.animBg }),
            }),
        ...(fullParagraphProjection ||
        nativeBehaviors.some(
          ({ targetType, additive }) =>
            targetType === "background" || additive === "sum",
        )
          ? { nativeBehaviors }
          : {}),
        kind,
        preset,
        durationMs,
        delayMs: entry.delayMs,
        ...(repeatCount === undefined ? {} : { repeatCount }),
        ...(by === undefined ? {} : { by }),
        ...(intervalMs === undefined ? {} : { intervalMs }),
        ...(repeatCount === "forever"
          ? { completion: "no-natural-completion" as const }
          : by === "word" || by === "letter"
            ? { completion: "unknown" as const }
            : {}),
        ...(startAfter === undefined ? {} : { startAfter }),
        ...(from === undefined ? {} : { from }),
        ...(fromColor === undefined ? {} : { fromColor }),
        ...(toColor === undefined ? {} : { toColor }),
        ...(Object.keys(easing).length === 0 ? {} : { easing }),
      });
    }
    if (effects.length > 0) result.push({ trigger, effects });
  }
  return { steps: result, containers };
}

function readCoreTitle(
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  errors: ValidationIssue[],
): string | undefined {
  const coreRelationshipType =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
  const relationship = [...(relationships.get("")?.values() ?? [])].find(
    (entry) => attribute(entry, "Type") === coreRelationshipType,
  );
  if (relationship === undefined) return undefined;
  const part = relationshipTargetPath("", attribute(relationship, "Target"));
  const root = documents.get(part)?.documentElement;
  if (root == null || root.localName !== "coreProperties") {
    issue(
      errors,
      part,
      "invalid-core-properties",
      "Presentation metadata part root element is invalid",
    );
    return undefined;
  }
  const title = first(root, "http://purl.org/dc/elements/1.1/", "title");
  return title?.textContent ?? undefined;
}

function readSections(
  root: XmlElement,
  slideIds: readonly XmlElement[],
  slideIdByPresentationId: ReadonlyMap<string, string>,
  part: string,
  errors: ValidationIssue[],
): {
  readonly sections: readonly PresentationSection[];
  readonly used: boolean;
} {
  const list = descendantsByName(root, p14Namespace, "sectionLst")[0];
  if (list === undefined) return { sections: [], used: false };
  const assigned = new Set<string>();
  const sections = elements(list, p14Namespace, "section").map((section) => {
    const name = attribute(section, "name");
    if (name.trim() === "") {
      issue(
        errors,
        part,
        "invalid-section-name",
        "Section name cannot be empty",
      );
    }
    const ids = first(section, p14Namespace, "sldIdLst");
    const sectionSlideIds = (
      ids === undefined ? [] : elements(ids, p14Namespace, "sldId")
    ).flatMap((entry) => {
      const presentationId = attribute(entry, "id");
      const slideId = slideIdByPresentationId.get(presentationId);
      if (slideId === undefined) {
        issue(
          errors,
          part,
          "missing-section-slide",
          `Section references non-existent slide ID ${presentationId}`,
        );
        return [];
      }
      if (assigned.has(presentationId)) {
        issue(
          errors,
          part,
          "duplicate-section-slide",
          `Slide ID ${presentationId} appears in multiple sections`,
        );
      }
      assigned.add(presentationId);
      return [slideId];
    });
    return { name, slideIds: sectionSlideIds };
  });
  for (const slide of slideIds) {
    const id = attribute(slide, "id");
    if (!assigned.has(id)) {
      issue(
        errors,
        part,
        "unassigned-section-slide",
        `Slide ID ${id} not assigned to any section`,
      );
    }
  }
  return { sections, used: true };
}

function notesRelationship(
  relationships: Map<string, Map<string, XmlElement>>,
  source: string,
  type: string,
): XmlElement[] {
  return [...(relationships.get(source)?.values() ?? [])].filter(
    (relationship) => attribute(relationship, "Type") === type,
  );
}

function readNotes(
  slidePath: string,
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  theme: ThemeTextReader,
  errors: ValidationIssue[],
): FileNotes | undefined {
  const notesSlides = notesRelationship(
    relationships,
    slidePath,
    `${relationshipNamespace}/notesSlide`,
  );
  if (notesSlides.length === 0) return undefined;
  if (notesSlides.length > 1) {
    issue(
      errors,
      slidePath,
      "duplicate-notes-slide-relationship",
      "Slide has more than one notes slide relationship",
    );
  }
  const notesPath = relationshipTargetPath(
    slidePath,
    attribute(notesSlides[0]!, "Target"),
  );
  const notesRoot = documents.get(notesPath)?.documentElement;
  if (
    notesRoot == null ||
    notesRoot.localName !== "notes" ||
    notesRoot.namespaceURI !== presentationNamespace
  ) {
    issue(
      errors,
      notesPath,
      "invalid-notes-slide",
      "Notes slide relationship target is not a valid Transitional notes slide part",
    );
    return undefined;
  }
  const slideRelations = notesRelationship(
    relationships,
    notesPath,
    `${relationshipNamespace}/slide`,
  );
  if (
    slideRelations.length !== 1 ||
    relationshipTargetPath(
      notesPath,
      attribute(slideRelations[0]!, "Target"),
    ) !== slidePath
  ) {
    issue(
      errors,
      notesPath,
      "notes-slide-target-mismatch",
      "Notes slide must reference its owning slide exactly once",
    );
  }
  const masters = notesRelationship(
    relationships,
    notesPath,
    `${relationshipNamespace}/notesMaster`,
  );
  if (masters.length !== 1) {
    issue(
      errors,
      notesPath,
      "notes-master-relationship",
      "Notes slide must reference exactly one notes master",
    );
    return undefined;
  }
  const masterPath = relationshipTargetPath(
    notesPath,
    attribute(masters[0]!, "Target"),
  );
  const masterRoot = documents.get(masterPath)?.documentElement;
  if (
    masterRoot == null ||
    masterRoot.localName !== "notesMaster" ||
    masterRoot.namespaceURI !== presentationNamespace
  ) {
    issue(
      errors,
      masterPath,
      "invalid-notes-master",
      "Notes master relationship target is not a valid Transitional notes master part",
    );
    return undefined;
  }
  const masterThemes = notesRelationship(
    relationships,
    masterPath,
    `${relationshipNamespace}/theme`,
  );
  const themePath =
    masterThemes.length === 1
      ? relationshipTargetPath(
          masterPath,
          attribute(masterThemes[0]!, "Target"),
        )
      : undefined;
  if (themePath === undefined) {
    issue(
      errors,
      masterPath,
      "notes-master-theme-relationship",
      "Notes master must reference exactly one theme",
    );
  }
  const notesTheme =
    themePath === undefined
      ? theme
      : themeTextReader(documents, themePath, masterPath);
  const commonSlideData = first(notesRoot, presentationNamespace, "cSld");
  const slideTree =
    commonSlideData === undefined
      ? undefined
      : first(commonSlideData, presentationNamespace, "spTree");
  if (slideTree === undefined) {
    issue(
      errors,
      notesPath,
      "missing-notes-shape-tree",
      "Notes slide missing shape tree",
    );
    return undefined;
  }
  const bodies = elements(slideTree, presentationNamespace, "sp").filter(
    (shape) => {
      const nonVisual = first(shape, presentationNamespace, "nvSpPr");
      const properties =
        nonVisual === undefined
          ? undefined
          : first(nonVisual, presentationNamespace, "nvPr");
      const placeholder =
        properties === undefined
          ? undefined
          : first(properties, presentationNamespace, "ph");
      return (
        placeholder !== undefined && attribute(placeholder, "type") === "body"
      );
    },
  );
  if (bodies.length !== 1) {
    issue(
      errors,
      notesPath,
      "notes-body-placeholder",
      "Notes slide must contain exactly one body placeholder",
    );
    if (bodies.length === 0) return undefined;
  }
  const body = readTextBody(bodies[0]!, notesTheme, notesPath, errors);
  if (body === undefined) {
    issue(
      errors,
      notesPath,
      "missing-notes-text-body",
      "Notes body placeholder missing text body",
    );
    return undefined;
  }
  const nonVisual = first(bodies[0]!, presentationNamespace, "nvSpPr")!;
  const properties = first(nonVisual, presentationNamespace, "nvPr")!;
  const placeholderElement = first(properties, presentationNamespace, "ph")!;
  const placeholderIndex = integerAttribute(placeholderElement, "idx");
  return {
    part: notesPath,
    masterPart: masterPath,
    ...(themePath === undefined ? {} : { themePart: themePath }),
    placeholder: {
      type: "body",
      ...(placeholderIndex === undefined ? {} : { index: placeholderIndex }),
    },
    text: body,
  };
}

function readPresentation(
  files: XmlFiles,
  documents: Documents,
  relationships: Map<string, Map<string, XmlElement>>,
  contentTypes: {
    defaults: Map<string, string>;
    overrides: Map<string, string>;
  },
  errors: ValidationIssue[],
): {
  presentation: ValidationResult["presentation"];
  slides: FileSlide[];
  extensions: ValidationExtension[];
} {
  const textTheme = themeTextReader(documents);
  const part = "ppt/presentation.xml";
  const document = documents.get(part);
  const root = document?.documentElement;
  if (
    root == null ||
    root.localName !== "presentation" ||
    root.namespaceURI !== presentationNamespace
  ) {
    issue(
      errors,
      part,
      "invalid-presentation",
      "Missing Transitional PresentationML root part",
    );
    return {
      presentation: { slideCount: 0 },
      slides: [],
      extensions: [],
    };
  }
  const size = first(root, presentationNamespace, "sldSz");
  const widthEmu = integerAttribute(size, "cx");
  const heightEmu = integerAttribute(size, "cy");
  if (
    widthEmu === undefined ||
    heightEmu === undefined ||
    widthEmu < minimumSlideSizeEmu ||
    widthEmu > maximumSlideSizeEmu ||
    heightEmu < minimumSlideSizeEmu ||
    heightEmu > maximumSlideSizeEmu
  ) {
    issue(
      errors,
      part,
      "invalid-slide-size",
      `Slide size must be within the closed interval of ${minimumSlideSizeEmu} to ${maximumSlideSizeEmu} EMU`,
    );
  }
  const ids = first(root, presentationNamespace, "sldIdLst");
  const slideIds =
    ids === undefined ? [] : elements(ids, presentationNamespace, "sldId");
  const presentationRelationships = relationships.get(part) ?? new Map();
  const slides: FileSlide[] = [];
  const usedExtensions = new Set<ValidationExtension>();
  const presentationSlideIds = new Map<string, string>();
  const stableSlideIds = new Set<string>();
  slideIds.forEach((slideId, index) => {
    const presentationId = attribute(slideId, "id");
    if (!/^\d+$/u.test(presentationId)) {
      issue(
        errors,
        part,
        "invalid-slide-id",
        `Slide ${index + 1} missing numeric ID`,
      );
    }
    if (presentationSlideIds.has(presentationId)) {
      issue(
        errors,
        part,
        "duplicate-slide-id",
        `Duplicate slide ID ${presentationId}`,
      );
    }
    const relationshipId =
      slideId.getAttributeNS(relationshipNamespace, "id") ?? "";
    const relationship = presentationRelationships.get(relationshipId);
    if (relationship === undefined) {
      issue(
        errors,
        part,
        "missing-slide-relationship",
        `Slide ${index + 1} missing corresponding relationship`,
      );
      return;
    }
    const slidePath = relationshipTargetPath(
      part,
      attribute(relationship, "Target"),
    );
    const slideDocument = documents.get(slidePath);
    const slideRoot = slideDocument?.documentElement;
    if (
      slideRoot == null ||
      slideRoot.localName !== "sld" ||
      slideRoot.namespaceURI !== presentationNamespace
    ) {
      issue(
        errors,
        slidePath,
        "invalid-slide",
        "Slide relationship target is not a valid Transitional slide part",
      );
      return;
    }
    const commonSlideData = first(slideRoot, presentationNamespace, "cSld");
    const slideTree =
      commonSlideData === undefined
        ? undefined
        : first(commonSlideData, presentationNamespace, "spTree");
    if (slideTree === undefined) {
      issue(
        errors,
        slidePath,
        "missing-shape-tree",
        "Slide missing shape tree",
      );
      return;
    }
    const objects = readObjects(
      slideTree,
      slidePath,
      files,
      documents,
      relationships,
      contentTypes,
      errors,
      textTheme,
    );
    const transition = readSlideTransition(slideRoot, slidePath, errors);
    const advanceAfterMs = readAdvanceAfter(slideRoot, slidePath, errors);
    const background = readSlideBackground(
      slideRoot,
      slidePath,
      files,
      relationships,
      contentTypes,
      textTheme,
      errors,
    );
    const animationRead = readSlideAnimations(
      slideRoot,
      slidePath,
      objects,
      {
        widthEmu: widthEmu ?? 0,
        heightEmu: heightEmu ?? 0,
      },
      errors,
    );
    const notes = readNotes(
      slidePath,
      files,
      documents,
      relationships,
      textTheme,
      errors,
    );
    for (const extension of transition.extensions) {
      usedExtensions.add(extension);
    }
    const seenIds = new Set<number>();
    const pending = [...objects];
    while (pending.length > 0) {
      const object = pending.pop();
      if (object === undefined) continue;
      if (seenIds.has(object.id)) {
        issue(
          errors,
          slidePath,
          "duplicate-object-id",
          `Duplicate object identifier ${object.id}`,
        );
      }
      seenIds.add(object.id);
      pending.push(...(object.members ?? []));
    }
    const stableSlideId =
      commonSlideData === undefined ? "" : attribute(commonSlideData, "name");
    if (stableSlideId === "") {
      issue(
        errors,
        slidePath,
        "missing-stable-slide-id",
        "Slide missing stable slide identifier",
      );
    }
    if (stableSlideIds.has(stableSlideId)) {
      issue(
        errors,
        slidePath,
        "duplicate-stable-slide-id",
        `Duplicate slide identifier ${stableSlideId}`,
      );
    }
    stableSlideIds.add(stableSlideId);
    presentationSlideIds.set(presentationId, stableSlideId);
    slides.push({
      id: stableSlideId,
      ...(attribute(slideRoot, "show") === "false" ? { hidden: true } : {}),
      ...(advanceAfterMs === undefined ? {} : { advanceAfterMs }),
      ...(background === undefined ? {} : { background }),
      objects,
      ...(transition.transition === undefined
        ? {}
        : { transition: transition.transition }),
      ...(animationRead === undefined
        ? {}
        : {
            animations: animationRead.steps,
            animationContainers: animationRead.containers,
          }),
      ...(notes === undefined ? {} : { notes }),
    });
  });
  const sectionResult = readSections(
    root,
    slideIds,
    presentationSlideIds,
    part,
    errors,
  );
  if (sectionResult.used) usedExtensions.add("p14");
  const sectionBySlideId = new Map<string, string>();
  for (const section of sectionResult.sections) {
    for (const slideId of section.slideIds)
      sectionBySlideId.set(slideId, section.name);
  }
  const slidesWithSections = slides.map((slide) => {
    const section = sectionBySlideId.get(slide.id);
    return { ...slide, ...(section === undefined ? {} : { section }) };
  });
  const title = readCoreTitle(documents, relationships, errors);
  if (slideIds.length === 0)
    issue(errors, part, "missing-slides", "Presentation contains no slides");
  if (files["ppt/theme/theme1.xml"] === undefined) {
    issue(
      errors,
      "ppt/theme/theme1.xml",
      "missing-theme",
      "Missing required theme part",
    );
  }
  return {
    presentation: {
      slideCount: slideIds.length,
      ...(widthEmu === undefined ? {} : { widthEmu }),
      ...(heightEmu === undefined ? {} : { heightEmu }),
      ...(title === undefined ? {} : { title }),
      ...(sectionResult.sections.length === 0
        ? {}
        : { sections: sectionResult.sections }),
    },
    slides: slidesWithSections,
    extensions: [...usedExtensions],
  };
}

/**
 * Executes SDK structural validation on the same PPTX bytes and independently reads slide objects, geometry, and native text,
 * without reading build reports. allowedExtensions constrains all transition compatibility branches discovered by the local reader;
 * SDK MC branch selection and extension coverage are limited; validation results do not represent application evidence.
 */
export async function validatePptx(
  input: Uint8Array,
  options: ValidationOptions = {},
): Promise<ValidationResult> {
  const errors = await validateStandardStructure(input);
  let files: XmlFiles;
  try {
    files = unzipSync(input);
  } catch (error) {
    issue(
      errors,
      "",
      "invalid-package",
      error instanceof Error ? error.message : String(error),
    );
    return {
      profile: {
        standard: transitionalStandard,
        schemaValidator: schemaValidatorProfile,
        scope: validationScope,
        applicationEvidence,
      },
      errors,
      extensions: [],
      presentation: { slideCount: 0 },
      slides: [],
    };
  }
  const documents = parseXmlFiles(files, errors);
  for (const requiredPart of [
    "[Content_Types].xml",
    "_rels/.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/theme/theme1.xml",
  ]) {
    if (files[requiredPart] === undefined)
      issue(errors, requiredPart, "missing-part", "PPTX missing required part");
  }
  const relationships = checkRelationships(files, documents, errors);
  const contentTypes = validateContentTypes(files, documents, errors);
  requireRoot(
    documents,
    errors,
    "ppt/slideMasters/slideMaster1.xml",
    presentationNamespace,
    "sldMaster",
  );
  validateSlideLayoutIds(documents, errors);
  requireRoot(
    documents,
    errors,
    "ppt/slideLayouts/slideLayout1.xml",
    presentationNamespace,
    "sldLayout",
  );
  requireRoot(
    documents,
    errors,
    "ppt/theme/theme1.xml",
    drawingNamespace,
    "theme",
  );
  validateThemeFonts(documents, errors);
  requireRelationship(
    "",
    `${relationshipNamespace}/officeDocument`,
    "ppt/presentation.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/presentation.xml",
    `${relationshipNamespace}/slideMaster`,
    "ppt/slideMasters/slideMaster1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideMasters/slideMaster1.xml",
    `${relationshipNamespace}/slideLayout`,
    "ppt/slideLayouts/slideLayout1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideMasters/slideMaster1.xml",
    `${relationshipNamespace}/theme`,
    "ppt/theme/theme1.xml",
    relationships,
    errors,
  );
  requireRelationship(
    "ppt/slideLayouts/slideLayout1.xml",
    `${relationshipNamespace}/slideMaster`,
    "ppt/slideMasters/slideMaster1.xml",
    relationships,
    errors,
  );
  const { presentation, slides, extensions } = readPresentation(
    files,
    documents,
    relationships,
    contentTypes,
    errors,
  );
  for (const [index] of slides.entries()) {
    requireRelationship(
      `ppt/slides/slide${index + 1}.xml`,
      `${relationshipNamespace}/slideLayout`,
      "ppt/slideLayouts/slideLayout1.xml",
      relationships,
      errors,
    );
  }
  for (const extension of extensions) {
    if (
      options.allowedExtensions !== undefined &&
      !options.allowedExtensions.includes(extension)
    ) {
      issue(
        errors,
        "ppt/slides",
        "forbidden-extension",
        `Target forbids actual file usage of ${extension}`,
      );
    }
  }
  return {
    profile: {
      standard: transitionalStandard,
      schemaValidator: schemaValidatorProfile,
      scope: validationScope,
      applicationEvidence,
    },
    errors,
    extensions,
    presentation,
    slides,
  };
}
