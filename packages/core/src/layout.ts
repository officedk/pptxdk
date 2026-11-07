import Yoga, {
  Align,
  Direction,
  Display,
  Edge,
  FlexDirection,
  Gutter,
  Justify,
  MeasureMode,
  PositionType,
  Wrap,
  type MeasureFunction,
  type Node as YogaNode,
} from "yoga-layout";

import { BuildError, CapabilityError } from "#/diagnostics.ts";
import { textStyleFrom } from "#/theme.ts";
import type {
  BuildOptions,
  CapabilityChild,
  CapabilityElement,
  CapabilityLayoutContract,
  CapabilityLayoutMeasurement,
  CapabilityParseResult,
  RegisteredCapabilityModule,
  Diagnostic,
  GeometryBox,
  LayoutLength,
  LayoutStyle,
  NativeElementKind,
  ThemeContext,
  ThemeScope,
  ObjectCapabilityModule,
} from "#/types.ts";

type Dimensions = { readonly width: number; readonly height: number };
type Edges = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};
type GridTrack =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "auto"; readonly min: number }
  | {
      readonly kind: "fr";
      readonly value: number;
      readonly min: number;
      readonly contentMinimum?: boolean;
    };
type GridPlacement = {
  readonly node: LayoutNode;
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly columnSpan: number;
};
type GridModel = {
  readonly placements: readonly GridPlacement[];
  readonly widths: readonly number[];
  readonly heights: readonly number[];
  readonly columnGap: number;
  readonly rowGap: number;
  readonly spanWidth: (start: number, span: number) => number;
  readonly spanHeight: (start: number, span: number) => number;
  readonly width: number;
  readonly height: number;
};
type LayoutNode = {
  readonly yoga: YogaNode;
  readonly element?: CapabilityElement;
  readonly style: LayoutStyle;
  readonly theme: ThemeContext;
  readonly path: string;
  readonly module?: ObjectCapabilityModule;
  parent?: LayoutNode;
  readonly children: LayoutNode[];
  readonly escaped: LayoutNode[];
  gridItem?: boolean;
  /** Retain allocated box height when the grid parent has not yet recalculated Yoga child nodes. */
  gridAreaHeight?: number;
  /** When text needs to provide a baseline, delegated from outer object node to inner measurement leaf node. */
  measureFunction?: MeasureFunction;
  baselineContent?: YogaNode;
  baseline?: number | undefined;
  parsed?: unknown;
  contract?: CapabilityLayoutContract | undefined;
};

export type LayoutPassResult = {
  readonly boxes: WeakMap<CapabilityElement, GeometryBox>;
  readonly themes: WeakMap<CapabilityElement, ThemeContext>;
  readonly hidden: WeakSet<CapabilityElement>;
  readonly warnings: readonly Diagnostic[];
  readonly parsed?: WeakMap<CapabilityElement, CapabilityParseResult>;
};

export const paddingStyleKeys = [
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
] as const;

const layoutProperties = [
  "display",
  "position",
  "inset",
  "left",
  "top",
  "right",
  "bottom",
  "width",
  "height",
  "minWidth",
  "minHeight",
  "maxWidth",
  "maxHeight",
  "margin",
  "marginTop",
  "marginRight",
  "marginBottom",
  "marginLeft",
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
  "gap",
  "rowGap",
  "columnGap",
  "flexDirection",
  "flexWrap",
  "justifyContent",
  "alignItems",
  "alignSelf",
  "alignContent",
  "flex",
  "flexGrow",
  "flexShrink",
  "flexBasis",
  "aspectRatio",
  "gridTemplateColumns",
  "gridTemplateRows",
  "gridColumn",
  "gridRow",
  "justifyItems",
  "justifySelf",
  "zIndex",
] as const;

const layoutPropertySet = new Set<string>(layoutProperties);

export const layoutStyleKeys = layoutProperties.filter(
  (key) => !(paddingStyleKeys as readonly string[]).includes(key),
);
const textProperties = new Set([
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
  "textWrap",
  "lang",
]);
const visualProperties = new Set([
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "borderRadius",
  "opacity",
  "boxShadow",
  "softEdge",
  "rotate",
]);
const groupVisualProperties = new Set(["rotate"]);
const containerProperties = [
  "flexDirection",
  "flexWrap",
  "justifyContent",
  "alignItems",
  "alignContent",
  "gap",
  "rowGap",
  "columnGap",
] as const;
const gridContainerProperties = [
  "gridTemplateColumns",
  "gridTemplateRows",
  "justifyItems",
] as const;
const gridItemProperties = ["gridColumn", "gridRow", "justifySelf"] as const;
const alignItems = [
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
  "baseline",
] as const;
const contentAlignments = [
  "start",
  "end",
  "center",
  "flex-start",
  "flex-end",
  "space-between",
  "space-around",
  "space-evenly",
] as const;

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

function fail(path: string, code: string, message: string): never {
  throw new BuildError([diagnostic(path, code, message)]);
}

function deriveTheme(
  parent: ThemeContext,
  scope: ThemeScope | undefined,
  style: unknown,
  path: string,
): ThemeContext {
  try {
    const textStyle = style === undefined ? {} : textStyleFrom(style);
    return parent.with(scope, textStyle);
  } catch (error) {
    return fail(
      path,
      error instanceof CapabilityError ? error.code : "invalid-theme-scope",
      error instanceof Error ? error.message : String(error),
    );
  }
}

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function hasVisibleBoxStyle(value: unknown): boolean {
  if (!record(value)) return false;
  return (
    value["backgroundColor"] !== undefined ||
    value["backgroundImage"] !== undefined ||
    value["boxShadow"] !== undefined ||
    (value["borderColor"] !== undefined && value["borderWidth"] !== 0)
  );
}

function describe(value: unknown): string {
  if (typeof value === "string") return `"${value}"`;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return typeof value;
}

function element(child: CapabilityChild): child is CapabilityElement {
  return "props" in child;
}

function gridTokens(input: string, path: string): string[] {
  const tokens: string[] = [];
  let depth = 0;
  let current = "";
  for (const character of input.trim()) {
    if (character === "(") depth += 1;
    if (character === ")") depth -= 1;
    if (depth < 0)
      fail(
        path,
        "invalid-grid-tracks",
        "Mismatched parentheses in grid tracks",
      );
    if (/\s/u.test(character) && depth === 0) {
      if (current !== "") tokens.push(current);
      current = "";
    } else current += character;
  }
  if (depth !== 0)
    fail(path, "invalid-grid-tracks", "Mismatched parentheses in grid tracks");
  if (current !== "") tokens.push(current);
  return tokens;
}

function parseGridPixels(token: string, path: string): number {
  const match = /^(\d+(?:\.\d+)?)(?:px)?$/u.exec(token);
  if (match === null)
    fail(path, "invalid-grid-tracks", `Unsupported grid track ${token}`);
  return Number(match[1]);
}

function parseGridBasicTrack(token: string, path: string): GridTrack {
  if (token === "auto") return { kind: "auto", min: 0 };
  const fraction = /^(\d+(?:\.\d+)?)fr$/u.exec(token);
  if (fraction !== null) {
    const value = Number(fraction[1]);
    if (value <= 0)
      fail(
        path,
        "invalid-grid-tracks",
        "fr track fraction must be greater than 0",
      );
    return { kind: "fr", value, min: 0 };
  }
  return { kind: "px", value: parseGridPixels(token, path) };
}

function gridTracks(
  input: string | undefined,
  path: string,
  defaultFr: boolean,
): GridTrack[] {
  if (input === undefined || input.trim() === "") {
    return defaultFr ? [{ kind: "fr", value: 1, min: 0 }] : [];
  }
  const parseOne = (token: string): GridTrack[] => {
    const repeat = /^repeat\(\s*(\d+)\s*,\s*(.+)\)$/u.exec(token);
    if (repeat !== null) {
      const count = Number(repeat[1]);
      if (!Number.isInteger(count) || count < 1 || count > 128) {
        fail(
          path,
          "invalid-grid-tracks",
          "repeat count must be an integer between 1 and 128",
        );
      }
      const repeated = gridTokens(repeat[2]!, path).flatMap(parseOne);
      if (repeated.length * count > 128) {
        fail(
          path,
          "invalid-grid-tracks",
          "Expanded grid tracks cannot exceed 128 tracks",
        );
      }
      return Array.from({ length: count }, () => repeated).flat();
    }
    const minmax = /^minmax\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)$/u.exec(token);
    if (minmax !== null) {
      const lower = minmax[1]!.trim();
      const minValue = lower === "auto" ? 0 : parseGridPixels(lower, path);
      const maximum = parseGridBasicTrack(minmax[2]!.trim(), path);
      if (maximum.kind === "fr") {
        return [
          {
            ...maximum,
            min: Math.max(minValue, maximum.min),
            contentMinimum: lower === "auto",
          },
        ];
      }
      if (maximum.kind === "auto") return [{ kind: "auto", min: minValue }];
      return [{ kind: "px", value: Math.max(minValue, maximum.value) }];
    }
    return [parseGridBasicTrack(token, path)];
  };
  const tracks = gridTokens(input, path).flatMap(parseOne);
  if (tracks.length === 0 || tracks.length > 128) {
    fail(
      path,
      "invalid-grid-tracks",
      "Grid tracks must contain between 1 and 128 tracks",
    );
  }
  return tracks;
}

function distributeFrTracks(
  sizes: number[],
  tracks: readonly GridTrack[],
  availableSize: number,
  gap: number,
): void {
  const trackAt = (index: number): GridTrack =>
    tracks[index] ?? { kind: "auto", min: 0 };
  const flexible = new Set(
    sizes.flatMap((_, index) => (trackAt(index).kind === "fr" ? [index] : [])),
  );
  if (flexible.size === 0) return;

  const gapSize = gap * Math.max(0, sizes.length - 1);
  const fixedSize = sizes.reduce(
    (sum, size, index) => sum + (trackAt(index).kind === "fr" ? 0 : size),
    0,
  );

  // When a lower bound exceeds the initial share, freeze that track first, then allocate remaining space to unfrozen tracks.
  while (flexible.size > 0) {
    const factorTotal = [...flexible].reduce((sum, index) => {
      const track = trackAt(index);
      return sum + (track.kind === "fr" ? track.value : 0);
    }, 0);
    const frozenSize = sizes.reduce(
      (sum, size, index) =>
        sum +
        (flexible.has(index) ? 0 : trackAt(index).kind === "fr" ? size : 0),
      0,
    );
    const freeSize = Math.max(
      0,
      availableSize - fixedSize - gapSize - frozenSize,
    );
    const undersized = [...flexible].filter((index) => {
      const track = trackAt(index);
      return (
        track.kind === "fr" &&
        freeSize * (track.value / factorTotal) < (sizes[index] ?? track.min)
      );
    });
    if (undersized.length > 0) {
      for (const index of undersized) flexible.delete(index);
      continue;
    }
    for (const index of flexible) {
      const track = trackAt(index);
      if (track.kind === "fr")
        sizes[index] = freeSize * (track.value / factorTotal);
    }
    return;
  }
}

function gridPlacement(
  value: number | string | undefined,
  path: string,
): { readonly start?: number; readonly span: number } {
  if (value === undefined) return { span: 1 };
  const positiveIndex = (raw: string): number => {
    const result = Number(raw);
    if (!Number.isInteger(result) || result < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid line numbers must be positive integers",
      );
    }
    return result - 1;
  };
  if (typeof value === "number") {
    if (!Number.isInteger(value) || value < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid line numbers must be positive integers",
      );
    }
    return { start: value - 1, span: 1 };
  }
  const text = value.trim();
  let match = /^span\s+(\d+)$/u.exec(text);
  if (match !== null) {
    const span = Number(match[1]);
    if (span < 1)
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    return { span };
  }
  match = /^(\d+)\s*\/\s*span\s+(\d+)$/u.exec(text);
  if (match !== null) {
    const span = Number(match[2]);
    if (span < 1)
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    return { start: positiveIndex(match[1]!), span };
  }
  match = /^(\d+)\s*\/\s*(\d+)$/u.exec(text);
  if (match !== null) {
    const start = positiveIndex(match[1]!);
    const end = positiveIndex(match[2]!);
    if (end <= start)
      fail(
        path,
        "invalid-grid-placement",
        "Grid end line number must be greater than start line number",
      );
    return { start, span: end - start };
  }
  if (/^\d+$/u.test(text)) return { start: positiveIndex(text), span: 1 };
  return fail(
    path,
    "invalid-grid-placement",
    `Unsupported grid placement ${value}`,
  );
}

function gridAlignment(
  value:
    | LayoutStyle["alignItems"]
    | LayoutStyle["justifyItems"]
    | LayoutStyle["justifySelf"],
): "start" | "end" | "center" | "stretch" {
  if (value === "center") return "center";
  if (value === "end" || value === "flex-end") return "end";
  if (value === "stretch") return "stretch";
  return "start";
}

function resolvePercent(
  value: LayoutLength | undefined,
  basis: number,
): number | undefined {
  if (typeof value !== "string") return undefined;
  const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u.exec(value);
  return match === null ? undefined : (Number(match[1]) * basis) / 100;
}

// Only allow shared text style properties through layout whitelist here; specific validation remains the responsibility of capability modules.
function allowsSharedTextStyle(type: string): boolean {
  return (
    type === "rp:text" ||
    type === "rp:table" ||
    type === "rp:chart" ||
    type === "rp:box"
  );
}

function identity(element: CapabilityElement): string | undefined {
  const testId = element.props["testId"];
  return typeof testId === "string" && testId.trim() !== ""
    ? testId
    : undefined;
}

function label(element: CapabilityElement, index: number): string {
  const type = element.type.startsWith("rp:")
    ? element.type.slice(3)
    : element.type;
  return `${type}[${identity(element) ?? String(index + 1)}]`;
}

function number(value: unknown, name: string, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail(path, "invalid-layout-style", `${name} must be a finite pixel number`);
  }
  return value;
}

function length(
  value: unknown,
  name: string,
  path: string,
  nonNegative = true,
): number | undefined {
  if (value === undefined || value === "auto") return undefined;
  if (typeof value === "string") {
    const percent = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u.exec(value);
    if (percent === null) {
      fail(
        path,
        "invalid-layout-style",
        `${name} must be a pixel number, percentage, or auto`,
      );
    }
    const amount = Number(percent[1]);
    if (!Number.isFinite(amount) || (nonNegative && amount < 0)) {
      fail(path, "invalid-layout-style", `${name} cannot be less than 0`);
    }
    return undefined;
  }
  const amount = number(value, name, path);
  if (nonNegative && amount < 0) {
    fail(path, "invalid-layout-style", `${name} cannot be less than 0`);
  }
  return amount;
}

function spacing(value: unknown, name: string, path: string): Edges {
  if (value === undefined) return { top: 0, right: 0, bottom: 0, left: 0 };
  if (typeof value === "number") {
    const amount = number(value, name, path);
    return { top: amount, right: amount, bottom: amount, left: amount };
  }
  if (!Array.isArray(value) || (value.length !== 2 && value.length !== 4)) {
    fail(
      path,
      "invalid-layout-style",
      `${name} must be a number, 2-tuple, or 4-tuple`,
    );
  }
  const values = value.map((item) => number(item, name, path));
  if (values.length === 2) {
    return {
      top: values[0]!,
      right: values[1]!,
      bottom: values[0]!,
      left: values[1]!,
    };
  }
  return {
    top: values[0]!,
    right: values[1]!,
    bottom: values[2]!,
    left: values[3]!,
  };
}

function yogaJustify(
  value: NonNullable<LayoutStyle["justifyContent"]>,
): Justify {
  switch (value) {
    case "start":
    case "flex-start":
      return Justify.FlexStart;
    case "end":
    case "flex-end":
      return Justify.FlexEnd;
    case "center":
      return Justify.Center;
    case "space-between":
      return Justify.SpaceBetween;
    case "space-around":
      return Justify.SpaceAround;
    case "space-evenly":
      return Justify.SpaceEvenly;
    default: {
      const exhaustiveCheck: never = value;
      throw new Error(`Unsupported justifyContent: ${String(exhaustiveCheck)}`);
    }
  }
}

function yogaAlign(
  value: NonNullable<LayoutStyle["alignItems"]> | "auto",
): Align {
  switch (value) {
    case "auto":
      return Align.Auto;
    case "start":
    case "flex-start":
      return Align.FlexStart;
    case "end":
    case "flex-end":
      return Align.FlexEnd;
    case "center":
      return Align.Center;
    case "stretch":
      return Align.Stretch;
    case "baseline":
      return Align.Baseline;
    default: {
      const exhaustiveCheck: never = value;
      throw new Error(`Unsupported alignItems: ${String(exhaustiveCheck)}`);
    }
  }
}

function yogaAlignContent(
  value: NonNullable<LayoutStyle["alignContent"]>,
): Align {
  switch (value) {
    case "start":
    case "flex-start":
      return Align.FlexStart;
    case "end":
    case "flex-end":
      return Align.FlexEnd;
    case "center":
      return Align.Center;
    case "space-between":
      return Align.SpaceBetween;
    case "space-around":
      return Align.SpaceAround;
    case "space-evenly":
      return Align.SpaceEvenly;
    case "stretch":
      return Align.Stretch;
    default: {
      const exhaustiveCheck: never = value;
      throw new Error(`Unsupported alignContent: ${String(exhaustiveCheck)}`);
    }
  }
}

function rowDirection(style: LayoutStyle): boolean {
  const direction = style.flexDirection ?? "row";
  return direction === "row" || direction === "row-reverse";
}

function reverseDirection(style: LayoutStyle): boolean {
  return (style.flexDirection ?? "row").endsWith("-reverse");
}

function readStyle(
  value: unknown,
  path: string,
  isContainer: boolean,
  isText: boolean,
  allowsPadding = isContainer,
  allowedPresentation: ReadonlySet<string> = new Set(),
): LayoutStyle {
  if (value === undefined) return {};
  if (!record(value)) {
    fail(path, "invalid-layout-style", "style must be an object");
  }
  const unknown = Object.keys(value).find(
    (key) =>
      !layoutPropertySet.has(key) &&
      !(isText && textProperties.has(key)) &&
      !allowedPresentation.has(key),
  );
  if (unknown !== undefined) {
    fail(
      path,
      "unsupported-layout-style",
      `style contains unsupported property ${unknown}`,
    );
  }
  if (
    !isContainer &&
    ((value["display"] !== undefined &&
      value["display"] !== "none" &&
      value["display"] !== "block") ||
      Object.keys(value).some((key) =>
        containerProperties.includes(
          key as (typeof containerProperties)[number],
        ),
      ))
  ) {
    fail(
      path,
      "invalid-layout-context",
      "Container-only styles cannot be used on objects without layout children",
    );
  }
  const display = value["display"];
  if (
    display !== undefined &&
    display !== "block" &&
    display !== "flex" &&
    display !== "grid" &&
    display !== "none"
  ) {
    fail(
      path,
      "unsupported-layout-value",
      "display only supports block, flex, grid, or none",
    );
  }
  const resolvedDisplay = display ?? "block";
  if (resolvedDisplay === "grid" && value["alignItems"] === "baseline") {
    fail(
      path,
      "unsupported-grid-alignment",
      "Grid containers do not support baseline alignment",
    );
  }
  for (const key of containerProperties) {
    if (
      value[key] !== undefined &&
      !(key === "alignItems" && resolvedDisplay === "grid") &&
      !(key === "gap" && resolvedDisplay === "grid") &&
      !(key === "rowGap" && resolvedDisplay === "grid") &&
      !(key === "columnGap" && resolvedDisplay === "grid") &&
      resolvedDisplay !== "flex"
    ) {
      fail(
        path,
        "invalid-layout-context",
        `${key} is only applicable to display: "flex"${key === "alignItems" ? " or grid" : ""} containers`,
      );
    }
  }
  for (const key of gridContainerProperties) {
    if (value[key] !== undefined && resolvedDisplay !== "grid") {
      fail(
        path,
        "invalid-layout-context",
        `${key} is only applicable to display: "grid" containers`,
      );
    }
  }
  const allowedValue = (
    name: string,
    candidate: unknown,
    values: readonly string[],
  ) => {
    if (
      candidate !== undefined &&
      !values.includes(typeof candidate === "string" ? candidate : "")
    ) {
      fail(
        path,
        "unsupported-layout-value",
        `${name} does not support value ${describe(candidate)}`,
      );
    }
  };
  allowedValue("flexDirection", value["flexDirection"], [
    "row",
    "column",
    "row-reverse",
    "column-reverse",
  ]);
  allowedValue("flexWrap", value["flexWrap"], ["nowrap", "wrap"]);
  allowedValue("justifyContent", value["justifyContent"], contentAlignments);
  allowedValue("alignItems", value["alignItems"], alignItems);
  allowedValue("alignSelf", value["alignSelf"], ["auto", ...alignItems]);
  allowedValue("alignContent", value["alignContent"], [
    ...contentAlignments,
    "stretch",
  ]);
  allowedValue("position", value["position"], [
    "static",
    "relative",
    "absolute",
  ]);
  allowedValue("justifyItems", value["justifyItems"], [
    "start",
    "end",
    "center",
    "stretch",
    "flex-start",
    "flex-end",
  ]);
  allowedValue("justifySelf", value["justifySelf"], [
    "auto",
    "start",
    "end",
    "center",
    "stretch",
    "flex-start",
    "flex-end",
  ]);
  const position =
    (value["position"] as LayoutStyle["position"] | undefined) ?? "static";
  if (
    position === "static" &&
    ["inset", "left", "top", "right", "bottom"].some(
      (key) => value[key] !== undefined,
    )
  ) {
    fail(
      path,
      "invalid-layout-context",
      "Position offsets are only applicable to relative or absolute elements",
    );
  }
  for (const key of ["flex", "flexGrow", "flexShrink"] as const) {
    const candidate = value[key];
    if (candidate !== undefined && number(candidate, key, path) < 0) {
      fail(path, "invalid-layout-style", `${key} cannot be less than 0`);
    }
  }
  for (const key of ["gap", "rowGap", "columnGap"] as const) {
    const candidate = value[key];
    if (candidate !== undefined && number(candidate, key, path) < 0) {
      fail(path, "invalid-layout-style", `${key} cannot be less than 0`);
    }
  }
  for (const key of [
    "width",
    "height",
    "minWidth",
    "minHeight",
    "maxWidth",
    "maxHeight",
    "flexBasis",
  ] as const) {
    length(value[key], key, path);
  }
  const ratio = value["aspectRatio"];
  if (
    ratio !== undefined &&
    (typeof ratio !== "number" || !Number.isFinite(ratio) || ratio <= 0)
  ) {
    fail(
      path,
      "invalid-layout-style",
      "aspectRatio must be a finite positive number",
    );
  }
  const zIndex = value["zIndex"];
  if (
    zIndex !== undefined &&
    (typeof zIndex !== "number" || !Number.isFinite(zIndex))
  ) {
    fail(path, "invalid-layout-style", "zIndex must be a finite number");
  }
  for (const key of ["inset", "left", "top", "right", "bottom"] as const) {
    const candidate = value[key];
    if (candidate !== undefined) length(candidate, key, path, false);
  }
  for (const key of ["gridTemplateColumns", "gridTemplateRows"] as const) {
    if (value[key] !== undefined && typeof value[key] !== "string") {
      fail(path, "invalid-layout-style", `${key} must be a grid track string`);
    }
  }
  for (const key of gridItemProperties) {
    const candidate = value[key];
    if (
      candidate !== undefined &&
      key !== "justifySelf" &&
      typeof candidate !== "number" &&
      typeof candidate !== "string"
    ) {
      fail(
        path,
        "invalid-layout-style",
        `${key} must be a line number or span declaration`,
      );
    }
  }
  const padding = value["padding"];
  if (padding !== undefined) {
    const values = spacing(padding, "padding", path);
    if (Object.values(values).some((item) => item < 0)) {
      fail(path, "invalid-layout-style", "padding cannot be less than 0");
    }
  }
  for (const key of [
    "paddingTop",
    "paddingRight",
    "paddingBottom",
    "paddingLeft",
  ] as const) {
    const candidate = value[key];
    if (candidate !== undefined && number(candidate, key, path) < 0) {
      fail(path, "invalid-layout-style", `${key} cannot be less than 0`);
    }
  }
  for (const key of [
    "marginTop",
    "marginRight",
    "marginBottom",
    "marginLeft",
  ] as const) {
    if (value[key] !== undefined) number(value[key], key, path);
  }
  if (value["margin"] !== undefined) spacing(value["margin"], "margin", path);
  if (
    !allowsPadding &&
    [
      "padding",
      "paddingTop",
      "paddingRight",
      "paddingBottom",
      "paddingLeft",
    ].some((key) => value[key] !== undefined)
  ) {
    fail(
      path,
      "invalid-layout-context",
      "padding is only applicable to layout containers",
    );
  }
  const minWidth = length(value["minWidth"], "minWidth", path);
  const maxWidth = length(value["maxWidth"], "maxWidth", path);
  const minHeight = length(value["minHeight"], "minHeight", path);
  const maxHeight = length(value["maxHeight"], "maxHeight", path);
  if (
    typeof value["minWidth"] === "number" &&
    typeof value["maxWidth"] === "number" &&
    minWidth !== undefined &&
    maxWidth !== undefined &&
    minWidth > maxWidth
  ) {
    fail(
      path,
      "invalid-layout-style",
      "minWidth cannot be greater than maxWidth",
    );
  }
  if (
    typeof value["minHeight"] === "number" &&
    typeof value["maxHeight"] === "number" &&
    minHeight !== undefined &&
    maxHeight !== undefined &&
    minHeight > maxHeight
  ) {
    fail(
      path,
      "invalid-layout-style",
      "minHeight cannot be greater than maxHeight",
    );
  }
  return value;
}

/**
 * Parses and validates host props into a CapabilityLayoutContract for object capabilities.
 */
export function parseCapabilityLayoutContract(options: {
  readonly nativeKind: NativeElementKind;
  readonly props: Readonly<Record<string, unknown>>;
  readonly allowedStyleKeys: ReadonlySet<string>;
  readonly allowPadding?: boolean;
  readonly path?: string;
}): CapabilityLayoutContract {
  const { nativeKind, props, allowedStyleKeys, allowPadding = false } = options;
  const path = options.path ?? "style";
  const rawStyle = props["style"];
  if (rawStyle !== undefined) {
    if (!record(rawStyle)) {
      throw new CapabilityError(
        "invalid-style",
        "style must be an object",
        options.path,
      );
    }
    const unknown = Object.keys(rawStyle).find(
      (key) => rawStyle[key] !== undefined && !allowedStyleKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-style",
        `style contains unsupported property ${unknown}`,
        options.path,
      );
    }
    if (!allowPadding) {
      const paddingKey = (paddingStyleKeys as readonly string[]).find(
        (key) => rawStyle[key] !== undefined,
      );
      if (paddingKey !== undefined) {
        throw new CapabilityError(
          "invalid-layout-context",
          "padding is only applicable to layout containers",
          options.path,
        );
      }
    }
  }

  const layoutStyle = readStyle(
    rawStyle,
    path,
    false,
    false,
    allowPadding,
    allowedStyleKeys,
  );

  let padding:
    | {
        readonly top: number;
        readonly right: number;
        readonly bottom: number;
        readonly left: number;
      }
    | undefined;
  if (allowPadding && rawStyle !== undefined) {
    const hasPadding = (paddingStyleKeys as readonly string[]).some(
      (key) => rawStyle[key] !== undefined,
    );
    if (hasPadding) {
      const edges = expandedSpacing(layoutStyle, "padding");
      padding = {
        top: edges.top,
        right: edges.right,
        bottom: edges.bottom,
        left: edges.left,
      };
    }
  }

  const x = typeof props["x"] === "number" ? props["x"] : undefined;
  const y = typeof props["y"] === "number" ? props["y"] : undefined;
  const styleWidth = rawStyle !== undefined ? rawStyle["width"] : undefined;
  const styleHeight = rawStyle !== undefined ? rawStyle["height"] : undefined;
  const width = (styleWidth ?? props["width"]) as
    | number
    | `${number}%`
    | "auto"
    | undefined;
  const height = (styleHeight ?? props["height"]) as
    | number
    | `${number}%`
    | "auto"
    | undefined;

  return {
    nativeKind,
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
    layoutStyle,
    ...(padding === undefined ? {} : { padding }),
  };
}

function expandedSpacing(
  style: LayoutStyle,
  kind: "margin" | "padding",
): Edges {
  const base = spacing(style[kind], kind, "layout");
  if (kind === "margin") {
    return {
      top: style.marginTop ?? base.top,
      right: style.marginRight ?? base.right,
      bottom: style.marginBottom ?? base.bottom,
      left: style.marginLeft ?? base.left,
    };
  }
  return {
    top: style.paddingTop ?? base.top,
    right: style.paddingRight ?? base.right,
    bottom: style.paddingBottom ?? base.bottom,
    left: style.paddingLeft ?? base.left,
  };
}

function styleLength(value: LayoutLength | undefined): number | undefined {
  return typeof value === "number" ? value : undefined;
}

function setLength(
  yoga: YogaNode,
  property:
    | "width"
    | "height"
    | "minWidth"
    | "minHeight"
    | "maxWidth"
    | "maxHeight"
    | "flexBasis",
  value: LayoutLength | undefined,
): void {
  const setters = {
    width: ["setWidth", "setWidthPercent"],
    height: ["setHeight", "setHeightPercent"],
    minWidth: ["setMinWidth", "setMinWidthPercent"],
    minHeight: ["setMinHeight", "setMinHeightPercent"],
    maxWidth: ["setMaxWidth", "setMaxWidthPercent"],
    maxHeight: ["setMaxHeight", "setMaxHeightPercent"],
    flexBasis: ["setFlexBasis", "setFlexBasisPercent"],
  } as const;
  const [pointSetter, percentSetter] = setters[property];
  const percent =
    typeof value === "string" ? /^(\d+(?:\.\d*)?|\.\d+)%$/u.exec(value) : null;
  if (percent !== null) {
    yoga[percentSetter](Number(percent[1]));
    return;
  }
  if (typeof value === "number") {
    yoga[pointSetter](value);
    return;
  }
  if (value === "auto") {
    const autoSetter = {
      width: "setWidthAuto",
      height: "setHeightAuto",
      minWidth: "setMinWidth",
      minHeight: "setMinHeight",
      maxWidth: "setMaxWidth",
      maxHeight: "setMaxHeight",
      flexBasis: "setFlexBasisAuto",
    }[property] as keyof YogaNode;
    const setter = yoga[autoSetter] as ((amount?: number) => void) | undefined;
    if (setter !== undefined) setter.call(yoga);
  }
}

function setPosition(
  yoga: YogaNode,
  edge: Edge,
  value: LayoutLength | undefined,
): void {
  if (value === undefined || value === "auto") return;
  yoga.setPosition(edge, value);
}

function layoutPixel(value: number): number {
  // Yoga stores dimensions as single-precision floats; normalized to hundredths of a pixel per text service precision.
  return Math.round(value * 100) / 100;
}

function objectModule(
  module: RegisteredCapabilityModule | undefined,
): module is ObjectCapabilityModule {
  return module !== undefined && "compile" in module;
}

/** Resolve Box, block, and flex/grid layouts before capability modules write native objects. */
export class LayoutPass {
  readonly #modules: ReadonlyMap<string, RegisteredCapabilityModule>;
  readonly #dimensions: Dimensions;
  readonly #slideNumber: number;
  readonly #options: BuildOptions;
  readonly #config = Yoga.Config.create();
  readonly #boxes = new WeakMap<CapabilityElement, GeometryBox>();
  readonly #themes = new WeakMap<CapabilityElement, ThemeContext>();
  readonly #hidden = new WeakSet<CapabilityElement>();
  readonly #parsedResults = new WeakMap<
    CapabilityElement,
    CapabilityParseResult
  >();
  readonly #warnings: Diagnostic[] = [];
  readonly #warningKeys = new Set<string>();

  constructor(
    modules: ReadonlyMap<string, RegisteredCapabilityModule>,
    dimensions: Dimensions,
    slideNumber = 1,
    options: BuildOptions = {},
  ) {
    this.#modules = modules;
    this.#dimensions = dimensions;
    this.#slideNumber = slideNumber;
    this.#options = options;
    this.#config.setPointScaleFactor(0);
  }

  run(
    children: readonly CapabilityChild[],
    parentPath: string,
    theme: ThemeContext,
    slideStyle?: unknown,
  ): LayoutPassResult {
    const rootYoga = Yoga.Node.create(this.#config);
    const rootStyle = readStyle(slideStyle, parentPath, true, true);
    const root: LayoutNode = {
      yoga: rootYoga,
      style: rootStyle,
      theme,
      path: parentPath,
      children: [],
      escaped: [],
    };
    rootYoga.setWidth(this.#dimensions.width);
    rootYoga.setHeight(this.#dimensions.height);
    rootYoga.setPositionType(PositionType.Relative);
    this.#applyStyle(rootYoga, rootStyle, { display: "block" }, true, true);
    try {
      this.#appendChildren(
        children,
        parentPath,
        root,
        slideStyle !== undefined,
        theme,
      );
      this.#prepareTextBaselines(root);
      // Fixed-size parent nodes may skip creating a new measure tree; solve initially, then mark measurement leaves dirty and recalculate.
      rootYoga.calculateLayout(
        this.#dimensions.width,
        this.#dimensions.height,
        Direction.LTR,
      );
      for (let pass = 0; pass < 3; pass += 1) {
        this.#resolveGridTree(root);
        rootYoga.calculateLayout(
          this.#dimensions.width,
          this.#dimensions.height,
          Direction.LTR,
        );
      }
      if (this.#markBaselineMeasurementsDirty(root)) {
        rootYoga.calculateLayout(
          this.#dimensions.width,
          this.#dimensions.height,
          Direction.LTR,
        );
      }
      if (this.#applyAutomaticMinimums(root)) {
        rootYoga.calculateLayout(
          this.#dimensions.width,
          this.#dimensions.height,
          Direction.LTR,
        );
        this.#resolveGridTree(root);
        rootYoga.calculateLayout(
          this.#dimensions.width,
          this.#dimensions.height,
          Direction.LTR,
        );
      }
      this.#commit(root, { x: 0, y: 0 });
    } finally {
      rootYoga.freeRecursive();
      this.#freeEscaped(root);
      this.#config.free();
    }
    return {
      boxes: this.#boxes,
      themes: this.#themes,
      hidden: this.#hidden,
      warnings: this.#warnings,
      parsed: this.#parsedResults,
    };
  }

  #validateLoose(input: CapabilityElement, path: string): LayoutStyle {
    const module = this.#modules.get(input.type);
    if (!objectModule(module)) return {};
    // Text, Table, and Chart consume font properties via shared text service and remain leaf objects in layout.
    const allowsTextStyle =
      allowsSharedTextStyle(input.type) ||
      (input.type === "rp:shape" &&
        typeof input.props["text"] === "string" &&
        input.props["text"] !== "");
    const isGroup = input.type === "rp:group";
    const presentation =
      input.type === "rp:group"
        ? groupVisualProperties
        : ["rp:box", "rp:text", "rp:shape", "rp:image"].includes(input.type)
          ? visualProperties
          : undefined;
    const allowsPadding =
      input.type === "rp:text" ||
      (input.type === "rp:shape" &&
        typeof input.props["text"] === "string" &&
        input.props["text"] !== "");
    return readStyle(
      input.props["style"],
      path,
      isGroup,
      allowsTextStyle,
      allowsPadding,
      presentation,
    );
  }

  // Transparent Groups are not in the layout tree, so theme context must be passed separately along the declaration tree.
  #appendChildren(
    children: readonly CapabilityChild[],
    parentPath: string,
    parent: LayoutNode,
    includeObjects: boolean,
    themeContext: ThemeContext,
  ): void {
    children.forEach((child, index) => {
      if (!element(child)) return;
      const childPath = `${parentPath}/${label(child, index)}`;
      if (parent.style.display === "none") {
        this.#markHidden(child);
        return;
      }
      if (child.type === "rp:box") {
        const childStyle = readStyle(
          child.props["style"],
          childPath,
          true,
          true,
          true,
          visualProperties,
        );
        if (childStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        this.#validateItemStyle(childStyle, parent, childPath);
        const built = this.#buildBox(child, childPath, parent, themeContext);
        this.#appendLayoutChild(parent, built);
        return;
      }
      if (child.type === "rp:group") {
        const childStyle = readStyle(
          child.props["style"],
          childPath,
          true,
          true,
          true,
          groupVisualProperties,
        );
        if (childStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        this.#validateItemStyle(childStyle, parent, childPath);
        if (!includeObjects) {
          // Group text defaults apply only to descendants, not indicating it should become a layout item.
          const optsIntoLayout = Object.keys(childStyle).some(
            (key) => key !== "zIndex" && !textProperties.has(key),
          );
          if (!optsIntoLayout) {
            const groupTheme = deriveTheme(
              themeContext,
              child.props["theme"] as ThemeScope | undefined,
              child.props["style"],
              childPath,
            );
            this.#themes.set(child, groupTheme);
            this.#appendChildren(
              child.children,
              childPath,
              parent,
              false,
              groupTheme,
            );
            return;
          }
        }
        const built = this.#buildGroup(child, childPath, parent, themeContext);
        this.#appendLayoutChild(parent, built);
        return;
      }
      const module = this.#modules.get(child.type);
      if (module !== undefined && "compileDeclaration" in module) {
        this.#themes.set(child, themeContext);
        this.#appendChildren(
          child.children,
          childPath,
          parent,
          includeObjects,
          themeContext,
        );
        return;
      }
      if (objectModule(module) && module.parse !== undefined) {
        let parseResult = this.#parsedResults.get(child);
        if (parseResult === undefined) {
          try {
            parseResult = module.parse(child, {
              path: childPath,
              theme: themeContext,
              options: this.#options,
              slideNumber: this.#slideNumber,
            });
            this.#parsedResults.set(child, parseResult);
          } catch (error) {
            if (error instanceof BuildError) throw error;
            const code =
              error instanceof CapabilityError
                ? error.code
                : "invalid-capability-input";
            const message =
              error instanceof Error ? error.message : String(error);
            return fail(
              error instanceof CapabilityError && error.path !== undefined
                ? error.path
                : childPath,
              code,
              message,
            );
          }
        }
        const { contract, parsed } = parseResult;
        if (contract.layoutStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        this.#validateItemStyle(contract.layoutStyle, parent, childPath);
        const built = this.#buildObject(
          child,
          childPath,
          parent,
          contract.layoutStyle,
          themeContext,
          parsed,
          contract,
        );
        if (!includeObjects) {
          if (
            built !== undefined &&
            this.#requiresLooseLayout(contract.layoutStyle)
          ) {
            this.#positionLooseGeometry(built);
            this.#appendLayoutChild(parent, built);
          }
          return;
        }
        if (built !== undefined) {
          this.#appendLayoutChild(parent, built);
        }
        return;
      }
      const looseStyle = includeObjects
        ? undefined
        : this.#validateLoose(child, childPath);
      if (looseStyle?.display === "none") {
        this.#markHidden(child);
        return;
      }
      if (!includeObjects) {
        this.#themes.set(child, themeContext);
        if (looseStyle !== undefined)
          this.#validateItemStyle(looseStyle, parent, childPath);
        if (looseStyle !== undefined && this.#requiresLooseLayout(looseStyle)) {
          const built = this.#buildObject(
            child,
            childPath,
            parent,
            looseStyle,
            themeContext,
          );
          if (built !== undefined) {
            this.#positionLooseGeometry(built);
            this.#appendLayoutChild(parent, built);
          }
        }
        return;
      }
      if (!objectModule(module)) return;
      const allowsTextStyle =
        allowsSharedTextStyle(child.type) ||
        (child.type === "rp:shape" &&
          typeof child.props["text"] === "string" &&
          child.props["text"] !== "");
      const presentation = ["rp:text", "rp:shape", "rp:image"].includes(
        child.type,
      )
        ? visualProperties
        : undefined;
      const childStyle = readStyle(
        child.props["style"],
        childPath,
        false,
        allowsTextStyle,
        child.type === "rp:text" ||
          (child.type === "rp:shape" &&
            typeof child.props["text"] === "string" &&
            child.props["text"] !== ""),
        presentation,
      );
      if (childStyle.display === "none") {
        this.#markHidden(child);
        return;
      }
      this.#validateItemStyle(childStyle, parent, childPath);
      const layoutChild = this.#buildObject(
        child,
        childPath,
        parent,
        childStyle,
        themeContext,
      );
      if (layoutChild === undefined) return;
      this.#appendLayoutChild(parent, layoutChild);
    });
  }

  #validateItemStyle(
    style: LayoutStyle,
    parent: LayoutNode,
    path: string,
  ): void {
    if (
      parent.style.display === "grid" &&
      style.position === "absolute" &&
      (style.alignSelf !== undefined ||
        style.justifySelf !== undefined ||
        style.gridColumn !== undefined ||
        style.gridRow !== undefined)
    ) {
      fail(
        path,
        "unsupported-grid-positioning",
        "Absolutely positioned grid items currently do not support grid placement or self alignment",
      );
    }
    if (
      parent.style.display === "grid" &&
      style.position !== "absolute" &&
      style.alignSelf === "baseline"
    ) {
      fail(
        path,
        "unsupported-grid-alignment",
        "Grid items do not support baseline alignment",
      );
    }
  }

  #appendLayoutChild(parent: LayoutNode, child: LayoutNode): void {
    child.parent = parent;
    if (
      child.style.position === "absolute" &&
      this.#containingBlock(parent) !== parent
    ) {
      this.#containingBlock(parent).escaped.push(child);
      return;
    }
    this.#markGridItem(child, parent);
    parent.yoga.insertChild(child.yoga, parent.children.length);
    parent.children.push(child);
  }

  #containingBlock(parent: LayoutNode): LayoutNode {
    let current: LayoutNode | undefined = parent;
    while (current !== undefined) {
      if (current.element === undefined) return current;
      if (
        current.style.position === "relative" ||
        current.style.position === "absolute"
      ) {
        return current;
      }
      current = current.parent;
    }
    return fail(
      parent.path,
      "layout-unavailable",
      "Positioned object is not within slide canvas",
    );
  }

  #placeEscaped(node: LayoutNode, containingBlock: GeometryBox): void {
    const container = Yoga.Node.create(this.#config);
    container.setWidth(containingBlock.width);
    container.setHeight(containingBlock.height);
    container.setPositionType(PositionType.Relative);
    container.insertChild(node.yoga, 0);
    try {
      this.#markBaselineMeasurementsDirty(node);
      for (let pass = 0; pass < 3; pass += 1) {
        container.calculateLayout(
          containingBlock.width,
          containingBlock.height,
          Direction.LTR,
        );
        this.#resolveGridTree(node);
      }
      if (this.#applyAutomaticMinimums(node)) {
        container.calculateLayout(
          containingBlock.width,
          containingBlock.height,
          Direction.LTR,
        );
        this.#resolveGridTree(node);
        container.calculateLayout(
          containingBlock.width,
          containingBlock.height,
          Direction.LTR,
        );
      }
      container.calculateLayout(
        containingBlock.width,
        containingBlock.height,
        Direction.LTR,
      );
      this.#commit(node, { x: containingBlock.x, y: containingBlock.y });
    } finally {
      container.removeChild(node.yoga);
      container.free();
    }
  }

  #freeEscaped(node: LayoutNode): void {
    for (const child of node.escaped) {
      child.yoga.freeRecursive();
      this.#freeEscaped(child);
    }
    for (const child of node.children) this.#freeEscaped(child);
  }

  #buildBox(
    input: CapabilityElement,
    path: string,
    parent: LayoutNode,
    themeContext: ThemeContext,
  ): LayoutNode {
    const unknown = Object.keys(input.props).find(
      (key) =>
        ![
          "testId",
          "part",
          "style",
          "theme",
          "morphKey",
          "link",
          "__motionContext",
        ].includes(key),
    );
    if (unknown !== undefined) {
      fail(
        path,
        "unknown-parameter",
        `Box contains unknown parameter ${unknown}`,
      );
    }
    const testId = input.props["testId"];
    if (
      testId !== undefined &&
      (typeof testId !== "string" || testId.trim() === "")
    ) {
      fail(path, "invalid-parameter", "testId must be a non-empty string");
    }
    const part = input.props["part"];
    if (
      part !== undefined &&
      (typeof part !== "string" || part.trim() === "")
    ) {
      fail(path, "invalid-parameter", "part must be a non-empty string");
    }
    const rawStyle = input.props["style"];
    const style = readStyle(rawStyle, path, true, true, true, visualProperties);
    if (
      input.props["morphKey"] !== undefined &&
      !hasVisibleBoxStyle(rawStyle)
    ) {
      fail(
        path,
        "morph-needs-object",
        "Only visible Box elements generating native background shapes can participate in Morph",
      );
    }
    const theme = deriveTheme(
      themeContext,
      input.props["theme"] as ThemeScope | undefined,
      rawStyle,
      path,
    );
    const yoga = Yoga.Node.create(this.#config);
    const node: LayoutNode = {
      yoga,
      element: input,
      style,
      theme,
      path,
      children: [],
      parent,
      escaped: [],
    };
    this.#themes.set(input, theme);
    this.#applyStyle(yoga, style, parent.style, false, true);
    this.#appendChildren(input.children, path, node, true, theme);
    return node;
  }

  #buildGroup(
    input: CapabilityElement,
    path: string,
    parent: LayoutNode,
    themeContext: ThemeContext,
  ): LayoutNode {
    const unknown = Object.keys(input.props).find(
      (key) =>
        ![
          "testId",
          "part",
          "name",
          "style",
          "theme",
          "enter",
          "emphasis",
          "motion",
          "__motionContext",
          "__motionEnabled",
          "__motionStep",
        ].includes(key),
    );
    if (unknown !== undefined)
      fail(
        path,
        "unknown-parameter",
        `Group contains unknown parameter ${unknown}`,
      );
    const style = readStyle(
      input.props["style"],
      path,
      true,
      true,
      true,
      groupVisualProperties,
    );
    const theme = deriveTheme(
      themeContext,
      input.props["theme"] as ThemeScope | undefined,
      input.props["style"],
      path,
    );
    const yoga = Yoga.Node.create(this.#config);
    const node: LayoutNode = {
      yoga,
      element: input,
      style,
      theme,
      path,
      children: [],
      parent,
      escaped: [],
    };
    this.#themes.set(input, theme);
    this.#applyStyle(yoga, style, parent.style, false, true);
    this.#appendChildren(input.children, path, node, true, theme);
    return node;
  }

  #markHidden(child: CapabilityElement): void {
    this.#hidden.add(child);
    for (const descendant of child.children) {
      if (element(descendant)) this.#markHidden(descendant);
    }
  }

  #markGridItem(child: LayoutNode, parent: LayoutNode): void {
    if (
      parent.style.display !== "grid" ||
      child.style.position === "absolute"
    ) {
      return;
    }
    child.gridItem = true;
    child.yoga.setPositionType(PositionType.Absolute);
  }

  #requiresLooseLayout(style: LayoutStyle): boolean {
    return (
      style.position === "absolute" ||
      style.position === "relative" ||
      style.aspectRatio !== undefined ||
      style.width !== undefined ||
      style.height !== undefined ||
      style.minWidth !== undefined ||
      style.minHeight !== undefined ||
      style.maxWidth !== undefined ||
      style.maxHeight !== undefined
    );
  }

  #buildObject(
    input: CapabilityElement,
    path: string,
    parent: LayoutNode,
    style: LayoutStyle,
    themeContext: ThemeContext,
    parsed?: unknown,
    contract?: CapabilityLayoutContract,
  ): LayoutNode | undefined {
    const module = this.#modules.get(input.type);
    if (!objectModule(module)) return undefined;
    const yoga = Yoga.Node.create(this.#config);
    const node: LayoutNode = {
      yoga,
      element: input,
      style,
      theme: themeContext,
      path,
      module,
      children: [],
      parent,
      escaped: [],
      ...(parsed !== undefined ? { parsed } : {}),
      ...(contract !== undefined ? { contract } : {}),
    };
    this.#themes.set(input, themeContext);
    this.#applyStyle(yoga, style, parent.style, false, false);
    if (contract?.padding !== undefined) {
      yoga.setPadding(Edge.Top, contract.padding.top);
      yoga.setPadding(Edge.Right, contract.padding.right);
      yoga.setPadding(Edge.Bottom, contract.padding.bottom);
      yoga.setPadding(Edge.Left, contract.padding.left);
    }
    this.#setDeclaredSize(node);
    if (style.position === "absolute") {
      const x =
        contract?.x ??
        (typeof input.props["x"] === "number" ? input.props["x"] : undefined);
      const y =
        contract?.y ??
        (typeof input.props["y"] === "number" ? input.props["y"] : undefined);
      if (
        style.left === undefined &&
        style.inset === undefined &&
        style.right === undefined &&
        x !== undefined
      ) {
        yoga.setPosition(Edge.Left, x);
      }
      if (
        style.top === undefined &&
        style.inset === undefined &&
        style.bottom === undefined &&
        y !== undefined
      ) {
        yoga.setPosition(Edge.Top, y);
      }
    }
    if (module.measureLayout !== undefined) {
      const measureFunction: MeasureFunction = (width, widthMode) => {
        const measured = this.#measure(
          node,
          widthMode === MeasureMode.Undefined
            ? {}
            : {
                availableWidth: width,
                widthMode:
                  widthMode === MeasureMode.Exactly ? "exactly" : "at-most",
              },
        );
        return {
          width:
            widthMode === MeasureMode.Exactly ? width : (measured?.width ?? 0),
          height: measured?.height ?? 0,
        };
      };
      node.measureFunction = measureFunction;
      yoga.setMeasureFunc(measureFunction);
    }
    return node;
  }

  #positionLooseGeometry(node: LayoutNode): void {
    const element = node.element;
    if (element === undefined) return;
    const props = element.props;
    const x = typeof props["x"] === "number" ? props["x"] : 0;
    const y = typeof props["y"] === "number" ? props["y"] : 0;
    if (node.style.position === "absolute") {
      if (
        node.style.left === undefined &&
        node.style.inset === undefined &&
        node.style.right === undefined
      )
        node.yoga.setPosition(Edge.Left, x);
      if (
        node.style.top === undefined &&
        node.style.inset === undefined &&
        node.style.bottom === undefined
      )
        node.yoga.setPosition(Edge.Top, y);
      return;
    }
    const offset = (
      start: LayoutLength | undefined,
      end: LayoutLength | undefined,
      basis: number,
    ) => {
      if (typeof start === "number") return start;
      if (typeof start === "string" && start.endsWith("%")) {
        return (Number(start.slice(0, -1)) * basis) / 100;
      }
      if (typeof end === "number") return -end;
      if (typeof end === "string" && end.endsWith("%")) {
        return -(Number(end.slice(0, -1)) * basis) / 100;
      }
      return 0;
    };
    const left = offset(
      node.style.left ?? node.style.inset,
      node.style.right ?? node.style.inset,
      this.#dimensions.width,
    );
    const top = offset(
      node.style.top ?? node.style.inset,
      node.style.bottom ?? node.style.inset,
      this.#dimensions.height,
    );
    node.yoga.setPositionType(PositionType.Absolute);
    node.yoga.setPositionAuto(Edge.Right);
    node.yoga.setPositionAuto(Edge.Bottom);
    node.yoga.setPosition(Edge.Left, x + left);
    node.yoga.setPosition(Edge.Top, y + top);
    node.yoga.setAlwaysFormsContainingBlock(true);
  }

  #gridPlacements(
    node: LayoutNode,
    explicitColumns: readonly GridTrack[],
  ): {
    readonly placements: readonly GridPlacement[];
    readonly columns: GridTrack[];
  } {
    const columns = [...explicitColumns];
    const children = node.children.filter((child) => child.gridItem === true);
    const occupied = new Set<string>();
    const placements: GridPlacement[] = [];
    let cursorRow = 0;
    let cursorColumn = 0;
    const ensureColumns = (count: number) => {
      if (count > 128)
        fail(
          node.path,
          "invalid-grid-placement",
          "Implicit grid columns cannot exceed 128 columns",
        );
      while (columns.length < count) columns.push({ kind: "auto", min: 0 });
    };
    for (const child of children) {
      const column = gridPlacement(child.style.gridColumn, child.path);
      const row = gridPlacement(child.style.gridRow, child.path);
      const columnSpan = column.span;
      if (column.start !== undefined) ensureColumns(column.start + columnSpan);
      else if (columnSpan > columns.length) ensureColumns(columnSpan);
      let currentRow = row.start ?? cursorRow;
      let currentColumn = column.start ?? cursorColumn;
      if (
        row.start === undefined &&
        column.start !== undefined &&
        column.start < cursorColumn
      )
        currentRow += 1;
      const fits = (rowIndex: number, columnIndex: number) => {
        if (row.start === undefined && column.start === undefined) {
          if (columnIndex + columnSpan > columns.length) return false;
        } else {
          ensureColumns(columnIndex + columnSpan);
        }
        for (let rowOffset = 0; rowOffset < row.span; rowOffset += 1) {
          for (
            let columnOffset = 0;
            columnOffset < columnSpan;
            columnOffset += 1
          ) {
            if (
              occupied.has(
                `${rowIndex + rowOffset}:${columnIndex + columnOffset}`,
              )
            )
              return false;
          }
        }
        return true;
      };
      const fixed = row.start !== undefined && column.start !== undefined;
      if (!fixed) {
        while (!fits(currentRow, currentColumn)) {
          if (column.start !== undefined) currentRow += 1;
          else {
            currentColumn += 1;
            if (currentColumn + columnSpan > columns.length) {
              currentColumn = 0;
              currentRow += 1;
            }
          }
        }
      }
      if (currentRow + row.span > 128) {
        fail(
          child.path,
          "invalid-grid-placement",
          "Implicit grid rows cannot exceed 128 rows",
        );
      }
      for (let rowOffset = 0; rowOffset < row.span; rowOffset += 1) {
        for (
          let columnOffset = 0;
          columnOffset < columnSpan;
          columnOffset += 1
        ) {
          occupied.add(
            `${currentRow + rowOffset}:${currentColumn + columnOffset}`,
          );
        }
      }
      placements.push({
        node: child,
        row: currentRow,
        column: currentColumn,
        rowSpan: row.span,
        columnSpan,
      });
      if (row.start === undefined) {
        cursorRow = currentRow;
        cursorColumn = currentColumn + columnSpan;
      }
    }
    return { placements, columns };
  }

  #naturalSize(node: LayoutNode, width?: number): Dimensions {
    node.yoga.calculateLayout(width, undefined, Direction.LTR);
    return {
      width: node.yoga.getComputedWidth(),
      height: node.yoga.getComputedHeight(),
    };
  }

  #gridModel(
    node: LayoutNode,
    innerWidth: number | undefined,
    innerHeight: number | undefined,
  ): GridModel {
    const columnTracks = gridTracks(
      node.style.gridTemplateColumns,
      node.path,
      true,
    );
    const rowTracks = gridTracks(node.style.gridTemplateRows, node.path, false);
    const { placements, columns } = this.#gridPlacements(node, columnTracks);
    const columnGap = node.style.columnGap ?? node.style.gap ?? 0;
    const rowGap = node.style.rowGap ?? node.style.gap ?? 0;
    const rowCount = Math.max(
      rowTracks.length,
      ...placements.map((placement) => placement.row + placement.rowSpan),
      0,
    );
    const minContentWidth = (item: LayoutNode): number =>
      this.#minContentContribution(item, "width");
    const columnWidths = columns.map((track) =>
      track.kind === "px" ? track.value : track.min,
    );
    for (const placement of placements) {
      if (placement.columnSpan !== 1) continue;
      const track = columns[placement.column];
      if (
        track === undefined ||
        track.kind === "px" ||
        (track.kind === "fr" && track.contentMinimum !== true)
      )
        continue;
      columnWidths[placement.column] = Math.max(
        columnWidths[placement.column] ?? 0,
        Math.max(track.min, minContentWidth(placement.node)),
      );
    }
    const columnSpanWidth = (start: number, span: number) =>
      columnWidths
        .slice(start, start + span)
        .reduce((sum, width) => sum + width, 0) +
      columnGap * Math.max(0, span - 1);
    for (const placement of placements) {
      if (placement.columnSpan === 1) continue;
      const required = minContentWidth(placement.node);
      const missing =
        required - columnSpanWidth(placement.column, placement.columnSpan);
      if (missing <= 0) continue;
      const growable = Array.from(
        { length: placement.columnSpan },
        (_, index) => placement.column + index,
      ).filter(
        (index) =>
          columns[index]?.kind === "auto" ||
          (columns[index]?.kind === "fr" &&
            columns[index]?.contentMinimum === true),
      );
      if (growable.length === 0) continue;
      for (const index of growable) {
        columnWidths[index] =
          (columnWidths[index] ?? 0) + missing / growable.length;
      }
    }
    if (innerWidth !== undefined) {
      distributeFrTracks(columnWidths, columns, innerWidth, columnGap);
    }
    const spanWidth = (start: number, span: number) =>
      columnWidths
        .slice(start, start + span)
        .reduce((sum, width) => sum + width, 0) +
      columnGap * Math.max(0, span - 1);
    const rowHeights = Array.from({ length: rowCount }, (_, index) => {
      const track = rowTracks[index] ?? { kind: "auto" as const, min: 0 };
      return track.kind === "px" ? track.value : track.min;
    });
    for (const placement of placements) {
      if (placement.rowSpan !== 1) continue;
      const track = rowTracks[placement.row] ?? {
        kind: "auto" as const,
        min: 0,
      };
      if (track.kind === "px" || (track.kind === "fr" && !track.contentMinimum))
        continue;
      const margin = expandedSpacing(placement.node.style, "margin");
      const size = this.#naturalSize(
        placement.node,
        spanWidth(placement.column, placement.columnSpan),
      );
      rowHeights[placement.row] = Math.max(
        rowHeights[placement.row] ?? 0,
        track.min,
        size.height + margin.top + margin.bottom,
      );
    }
    const spanHeight = (start: number, span: number) =>
      rowHeights
        .slice(start, start + span)
        .reduce((sum, height) => sum + height, 0) +
      rowGap * Math.max(0, span - 1);
    for (const placement of placements) {
      if (placement.rowSpan === 1) continue;
      const natural = this.#naturalSize(
        placement.node,
        spanWidth(placement.column, placement.columnSpan),
      );
      const margin = expandedSpacing(placement.node.style, "margin");
      const missing =
        natural.height +
        margin.top +
        margin.bottom -
        spanHeight(placement.row, placement.rowSpan);
      if (missing <= 0) continue;
      const growable = Array.from(
        { length: placement.rowSpan },
        (_, index) => placement.row + index,
      ).filter((index) => (rowTracks[index]?.kind ?? "auto") !== "px");
      if (growable.length === 0) continue;
      for (const index of growable) {
        rowHeights[index] =
          (rowHeights[index] ?? 0) + missing / growable.length;
      }
    }
    if (innerHeight !== undefined) {
      distributeFrTracks(rowHeights, rowTracks, innerHeight, rowGap);
    }
    const spanHeightFinal = (start: number, span: number) =>
      rowHeights
        .slice(start, start + span)
        .reduce((sum, height) => sum + height, 0) +
      rowGap * Math.max(0, span - 1);
    return {
      placements,
      widths: columnWidths,
      heights: rowHeights,
      columnGap,
      rowGap,
      spanWidth,
      spanHeight: spanHeightFinal,
      width: spanWidth(0, columnWidths.length),
      height: rowCount === 0 ? 0 : spanHeightFinal(0, rowCount),
    };
  }

  #placeGrid(node: LayoutNode): void {
    const padding = expandedSpacing(node.style, "padding");
    const outerWidth = node.yoga.getComputedWidth();
    const outerHeight =
      node.yoga.getComputedHeight() || node.gridAreaHeight || 0;
    const hasWidth =
      node.element === undefined ||
      (node.style.width !== undefined && node.style.width !== "auto");
    const hasResolvedHeight =
      node.element === undefined ||
      (node.style.height !== undefined && node.style.height !== "auto") ||
      outerHeight > 0;
    const innerWidth =
      outerWidth > 0
        ? Math.max(0, outerWidth - padding.left - padding.right)
        : undefined;
    const innerHeight = hasResolvedHeight
      ? Math.max(0, outerHeight - padding.top - padding.bottom)
      : undefined;
    const model = this.#gridModel(node, innerWidth, innerHeight);
    if (innerWidth !== undefined && model.width > innerWidth + 0.01) {
      this.#warnOnce(
        node.path,
        "grid-overflow",
        `Grid columns minimum size total ${model.width}px exceeds available width ${innerWidth}px; build retains track dimensions`,
      );
    }
    if (innerHeight !== undefined && model.height > innerHeight + 0.01) {
      this.#warnOnce(
        node.path,
        "grid-overflow",
        `Grid rows minimum size total ${model.height}px exceeds available height ${innerHeight}px; build retains track dimensions`,
      );
    }
    if (!hasWidth && outerWidth <= 0) {
      node.yoga.setWidth(model.width + padding.left + padding.right);
    }
    if (!hasResolvedHeight) {
      node.yoga.setHeight(model.height + padding.top + padding.bottom);
    }
    for (const placement of model.placements) {
      const child = placement.node;
      const childStyle = child.style;
      const cellWidth = model.spanWidth(placement.column, placement.columnSpan);
      const cellHeight = model.spanHeight(placement.row, placement.rowSpan);
      const cellX =
        padding.left +
        model.spanWidth(0, placement.column) +
        (placement.column > 0 ? model.columnGap : 0);
      const cellY =
        padding.top +
        model.spanHeight(0, placement.row) +
        (placement.row > 0 ? model.rowGap : 0);
      const margin = expandedSpacing(childStyle, "margin");
      const availableWidth = Math.max(
        0,
        cellWidth - margin.left - margin.right,
      );
      const availableHeight = Math.max(
        0,
        cellHeight - margin.top - margin.bottom,
      );
      const justify = gridAlignment(
        childStyle.justifySelf === undefined ||
          childStyle.justifySelf === "auto"
          ? (node.style.justifyItems ?? "stretch")
          : childStyle.justifySelf,
      );
      const alignValue =
        childStyle.alignSelf === undefined || childStyle.alignSelf === "auto"
          ? (node.style.alignItems ?? "stretch")
          : childStyle.alignSelf;
      if (alignValue === "baseline") {
        fail(
          child.path,
          "unsupported-grid-alignment",
          "Text baseline alignment in grid does not yet have a reliable row baseline",
        );
      }
      const align = gridAlignment(alignValue);
      const declaredWidth = childStyle.width ?? child.element?.props["width"];
      const declaredHeight =
        childStyle.height ?? child.element?.props["height"];
      const widthPercent = resolvePercent(
        declaredWidth as LayoutLength | undefined,
        cellWidth,
      );
      const heightPercent = resolvePercent(
        declaredHeight as LayoutLength | undefined,
        cellHeight,
      );
      const widthSpecified =
        declaredWidth !== undefined && declaredWidth !== "auto";
      let heightSpecified =
        declaredHeight !== undefined && declaredHeight !== "auto";
      if (widthPercent !== undefined) child.yoga.setWidth(widthPercent);
      if (heightPercent !== undefined && availableHeight > 0) {
        child.yoga.setHeight(heightPercent);
      } else if (heightPercent !== undefined) {
        child.yoga.setHeightAuto();
        heightSpecified = false;
      }
      const minWidth = resolvePercent(childStyle.minWidth, cellWidth);
      const maxWidth = resolvePercent(childStyle.maxWidth, cellWidth);
      const minHeight = resolvePercent(childStyle.minHeight, cellHeight);
      const maxHeight = resolvePercent(childStyle.maxHeight, cellHeight);
      if (minWidth !== undefined) child.yoga.setMinWidth(minWidth);
      if (maxWidth !== undefined) child.yoga.setMaxWidth(maxWidth);
      if (minHeight !== undefined) child.yoga.setMinHeight(minHeight);
      if (maxHeight !== undefined) child.yoga.setMaxHeight(maxHeight);
      const intrinsic = this.#naturalSize(
        child,
        typeof declaredWidth === "string" && declaredWidth.endsWith("%")
          ? cellWidth
          : undefined,
      );
      let width = widthSpecified
        ? intrinsic.width
        : justify === "stretch"
          ? availableWidth
          : intrinsic.width;
      if (
        childStyle.aspectRatio !== undefined &&
        heightSpecified &&
        !widthSpecified
      ) {
        width = intrinsic.height * childStyle.aspectRatio;
      }
      if (!widthSpecified || widthPercent !== undefined)
        child.yoga.setWidth(Math.max(0, width));
      const naturalAtWidth = this.#naturalSize(child, width);
      let height = heightSpecified
        ? naturalAtWidth.height
        : align === "stretch" && childStyle.aspectRatio === undefined
          ? availableHeight
          : naturalAtWidth.height;
      if (!heightSpecified && childStyle.aspectRatio === undefined) {
        child.yoga.setHeight(Math.max(0, height));
      }
      const usedWidth = child.yoga.getComputedWidth() || width;
      const usedHeight = child.yoga.getComputedHeight() || height;
      if (child.style.display === "grid" && !heightSpecified) {
        child.gridAreaHeight = height;
      }
      if (
        usedWidth > availableWidth + 0.01 ||
        usedHeight > availableHeight + 0.01
      ) {
        this.#warnOnce(
          child.path,
          "grid-item-overflow",
          "Grid item overflows its track area; build retains declared dimensions",
        );
      }
      const xAlign =
        justify === "center"
          ? Math.max(0, (availableWidth - usedWidth) / 2)
          : justify === "end"
            ? Math.max(0, availableWidth - usedWidth)
            : 0;
      const yAlign =
        align === "center"
          ? Math.max(0, (availableHeight - usedHeight) / 2)
          : align === "end"
            ? Math.max(0, availableHeight - usedHeight)
            : 0;
      const relative = (
        start: LayoutLength | undefined,
        end: LayoutLength | undefined,
        basis: number,
      ) => {
        const resolve = (value: LayoutLength | undefined) =>
          typeof value === "number"
            ? value
            : typeof value === "string" && value.endsWith("%")
              ? (Number(value.slice(0, -1)) * basis) / 100
              : undefined;
        const startValue = resolve(start);
        return startValue ?? -(resolve(end) ?? 0);
      };
      const dx =
        childStyle.position === "relative"
          ? relative(
              childStyle.left ?? childStyle.inset,
              childStyle.right ?? childStyle.inset,
              cellWidth,
            )
          : 0;
      const dy =
        childStyle.position === "relative"
          ? relative(
              childStyle.top ?? childStyle.inset,
              childStyle.bottom ?? childStyle.inset,
              cellHeight,
            )
          : 0;
      child.yoga.setPositionType(PositionType.Absolute);
      child.yoga.setPosition(Edge.Left, cellX + xAlign + dx);
      child.yoga.setPosition(Edge.Top, cellY + yAlign + dy);
    }
  }

  #resolveGridTree(node: LayoutNode): void {
    if (node.style.display === "grid") this.#placeGrid(node);
    for (const child of node.children) this.#resolveGridTree(child);
  }

  #warnOnce(path: string, code: string, message: string): void {
    const key = `${path}\u0000${code}\u0000${message}`;
    if (this.#warningKeys.has(key)) return;
    this.#warningKeys.add(key);
    this.#warnings.push({ severity: "warning", path, code, message });
  }

  #setDeclaredSize(node: LayoutNode): void {
    const props = node.element?.props ?? {};
    const style = node.style;
    const rawWidth = node.contract?.width ?? style.width ?? props["width"];
    const width = rawWidth;
    const rawHeight = node.contract?.height ?? style.height ?? props["height"];
    const height = rawHeight;
    if (width !== undefined)
      setLength(node.yoga, "width", width as LayoutLength);
    if (height !== undefined)
      setLength(node.yoga, "height", height as LayoutLength);
  }

  #applyStyle(
    yoga: YogaNode,
    style: LayoutStyle,
    parent: LayoutStyle,
    isRoot: boolean,
    isContainer: boolean,
  ): void {
    const display = style.display ?? "block";
    yoga.setDisplay(display === "none" ? Display.None : Display.Flex);
    yoga.setPositionType(
      style.position === "absolute"
        ? PositionType.Absolute
        : style.position === "relative" || isRoot
          ? PositionType.Relative
          : PositionType.Static,
    );
    for (const [edge, value] of [
      [Edge.Left, style.left ?? style.inset],
      [Edge.Top, style.top ?? style.inset],
      [Edge.Right, style.right ?? style.inset],
      [Edge.Bottom, style.bottom ?? style.inset],
    ] as const) {
      setPosition(yoga, edge, value);
    }
    if (style.aspectRatio !== undefined) yoga.setAspectRatio(style.aspectRatio);
    const flex = style.display === "flex";
    yoga.setFlexDirection(
      flex
        ? {
            row: FlexDirection.Row,
            column: FlexDirection.Column,
            "row-reverse": FlexDirection.RowReverse,
            "column-reverse": FlexDirection.ColumnReverse,
          }[style.flexDirection ?? "row"]
        : FlexDirection.Column,
    );
    if (flex && style.flexWrap === "wrap") yoga.setFlexWrap(Wrap.Wrap);
    if (flex && style.justifyContent !== undefined) {
      const align = style.justifyContent;
      const isReversed = reverseDirection(style);
      const physicalAlign: NonNullable<LayoutStyle["justifyContent"]> =
        align === "start"
          ? isReversed
            ? "flex-end"
            : "flex-start"
          : align === "end"
            ? isReversed
              ? "flex-start"
              : "flex-end"
            : align;
      yoga.setJustifyContent(yogaJustify(physicalAlign));
    }
    if (flex && style.alignItems !== undefined) {
      yoga.setAlignItems(yogaAlign(style.alignItems));
    }
    if (flex && style.alignContent !== undefined) {
      yoga.setAlignContent(yogaAlignContent(style.alignContent));
    } else if (flex && style.flexWrap === "wrap") {
      // Yoga stacks flex lines from start edge by default; CSS stretches lines with remaining space by default.
      yoga.setAlignContent(Align.Stretch);
    }
    const grid = style.display === "grid";
    if ((flex || grid) && style.gap !== undefined)
      yoga.setGap(Gutter.All, style.gap);
    if ((flex || grid) && style.rowGap !== undefined)
      yoga.setGap(Gutter.Row, style.rowGap);
    if ((flex || grid) && style.columnGap !== undefined)
      yoga.setGap(Gutter.Column, style.columnGap);
    for (const key of [
      "width",
      "height",
      "minWidth",
      "minHeight",
      "maxWidth",
      "maxHeight",
    ] as const) {
      setLength(yoga, key, style[key]);
    }
    const padding = expandedSpacing(style, "padding");
    yoga.setPadding(Edge.Top, padding.top);
    yoga.setPadding(Edge.Right, padding.right);
    yoga.setPadding(Edge.Bottom, padding.bottom);
    yoga.setPadding(Edge.Left, padding.left);
    if (!isRoot) {
      const margin = expandedSpacing(style, "margin");
      yoga.setMargin(Edge.Top, margin.top);
      yoga.setMargin(Edge.Right, margin.right);
      yoga.setMargin(Edge.Bottom, margin.bottom);
      yoga.setMargin(Edge.Left, margin.left);
      if (parent.display === "flex") {
        yoga.setFlexShrink(style.flexShrink ?? 1);
        if (style.flex !== undefined) {
          yoga.setFlexGrow(style.flex);
          yoga.setFlexShrink(style.flexShrink ?? 1);
          yoga.setFlexBasis(0);
        }
        if (style.flexGrow !== undefined) yoga.setFlexGrow(style.flexGrow);
        if (style.flexBasis !== undefined)
          setLength(yoga, "flexBasis", style.flexBasis);
        if (style.alignSelf !== undefined)
          yoga.setAlignSelf(yogaAlign(style.alignSelf));
      } else {
        yoga.setFlexShrink(0);
      }
    }
    if (isContainer && style.display !== "flex") {
      yoga.setAlignItems(Align.Stretch);
    }
  }

  #measure(
    node: LayoutNode,
    request: {
      readonly availableWidth?: number;
      readonly widthMode?: "at-most" | "exactly";
    },
  ): CapabilityLayoutMeasurement | undefined {
    if (
      node.module?.measureLayout === undefined ||
      node.element === undefined
    ) {
      return undefined;
    }
    try {
      return node.module.measureLayout(node.parsed ?? node.element, {
        ...request,
        theme: node.theme,
        slideNumber: this.#slideNumber,
      });
    } catch (error) {
      if (error instanceof BuildError) throw error;
      const code =
        error instanceof CapabilityError
          ? error.code
          : "layout-measurement-failed";
      const message = error instanceof Error ? error.message : String(error);
      return fail(node.path, code, message);
    }
  }

  #applyAutomaticMinimums(root: LayoutNode): boolean {
    let changed = false;
    const flex = root.style.display === "flex";
    if (flex) {
      const horizontal = rowDirection(root.style);
      for (const child of root.children) {
        const explicitMinimum = horizontal
          ? child.style.minWidth
          : child.style.minHeight;
        if (
          child.style.position === "absolute" ||
          (explicitMinimum !== undefined && explicitMinimum !== "auto")
        )
          continue;
        const axis = horizontal ? "width" : "height";
        const minimum = this.#automaticMinimum(child, axis);
        if (minimum === undefined || minimum <= 0) continue;
        if (horizontal) child.yoga.setMinWidth(minimum);
        else child.yoga.setMinHeight(minimum);
        changed = true;
      }
    }
    for (const child of root.children) {
      if (this.#applyAutomaticMinimums(child)) changed = true;
    }
    return changed;
  }

  #automaticMinimum(
    node: LayoutNode,
    axis: "width" | "height",
  ): number | undefined {
    const content = this.#minContent(node, axis);
    if (content === undefined) return undefined;
    const maximum =
      axis === "width"
        ? styleLength(node.style.maxWidth)
        : styleLength(node.style.maxHeight);
    const rawDimension = node.style[axis] ?? node.element?.props[axis];
    const declared = length(rawDimension, axis, node.path);
    return Math.min(
      content,
      ...(declared === undefined ? [] : [declared]),
      ...(maximum === undefined ? [] : [maximum]),
    );
  }

  #resolvedChildAlignment(node: LayoutNode, child: LayoutNode): Align {
    // Use properties applied to Yoga to replicate its column baseline fallback to flex-start.
    const childAlignment = child.yoga.getAlignSelf();
    const alignment =
      childAlignment === Align.Auto
        ? node.yoga.getAlignItems()
        : childAlignment;
    const direction = node.yoga.getFlexDirection();
    if (
      alignment === Align.Baseline &&
      (direction === FlexDirection.Column ||
        direction === FlexDirection.ColumnReverse)
    ) {
      return Align.FlexStart;
    }
    return alignment;
  }

  #baselineParticipants(node: LayoutNode): LayoutNode[] {
    // Resolve participants using actual properties of layout node, ignoring raw styles dropped by non-flex parents.
    return node.children.filter(
      (child) => this.#resolvedChildAlignment(node, child) === Align.Baseline,
    );
  }

  #isTextBearing(node: LayoutNode): boolean {
    return (
      node.element?.type === "rp:text" ||
      (typeof node.element?.props["text"] === "string" &&
        node.element.props["text"] !== "")
    );
  }

  #containsText(node: LayoutNode): boolean {
    return (
      this.#isTextBearing(node) ||
      node.children.some((child) => this.#containsText(child))
    );
  }

  #baselineTextLeaf(node: LayoutNode): LayoutNode {
    if (this.#isTextBearing(node)) return node;
    if (node.element?.type !== "rp:box" || node.children.length !== 1) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Text baseline currently only propagates along single-child Box chains",
      );
    }
    const child = node.children[0]!;
    const simpleRow =
      node.style.display === "flex" &&
      node.style.flexWrap !== "wrap" &&
      rowDirection(node.style) &&
      !reverseDirection(node.style);
    const childAlign = child.style.alignSelf;
    const parentAlign = node.style.alignItems;
    const alignedAtStart =
      childAlign === "start" ||
      childAlign === "flex-start" ||
      ((childAlign === undefined || childAlign === "auto") &&
        (parentAlign === "start" || parentAlign === "flex-start"));
    const stretchedFixedHeight =
      (childAlign === undefined ||
        childAlign === "auto" ||
        childAlign === "stretch") &&
      (parentAlign === undefined || parentAlign === "stretch") &&
      typeof (child.style.height ?? child.element?.props["height"]) ===
        "number";
    if (!simpleRow || (!alignedAtStart && !stretchedFixedHeight)) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Text baseline requires a nowrap row Box with the only child fixed at the top of the cross axis",
      );
    }
    return this.#baselineTextLeaf(child);
  }

  #prepareTextBaselines(node: LayoutNode): void {
    const flexRow = node.style.display === "flex" && rowDirection(node.style);
    const participants = this.#baselineParticipants(node);
    if (flexRow && participants.some((child) => this.#containsText(child))) {
      if (node.style.flexWrap === "wrap") {
        fail(
          node.path,
          "unsupported-baseline-context",
          "Text baseline currently requires a nowrap row flex container",
        );
      }
      for (const participant of participants) {
        if (!this.#containsText(participant)) continue;
        this.#installBaselineMarker(this.#baselineTextLeaf(participant));
      }
    }
    for (const child of node.children) this.#prepareTextBaselines(child);
    for (const child of node.escaped) this.#prepareTextBaselines(child);
  }

  #installBaselineMarker(node: LayoutNode): void {
    if (node.baselineContent !== undefined) return;
    const measureFunction = node.measureFunction;
    if (measureFunction === undefined) {
      fail(
        node.path,
        "baseline-measurement-unavailable",
        "Baseline alignment requires reliable font measurement from text capability",
      );
    }
    const rawWidth = node.style.width ?? node.element?.props["width"];
    const availableWidth = length(rawWidth, "width", node.path);
    const padding = expandedSpacing(node.style, "padding");
    const measurement = this.#measure(
      node,
      availableWidth === undefined
        ? {}
        : {
            availableWidth: Math.max(
              0,
              availableWidth - padding.left - padding.right,
            ),
          },
    );
    if (
      measurement?.baseline === "unavailable" ||
      measurement?.baseline === undefined
    ) {
      fail(
        node.path,
        "baseline-measurement-unavailable",
        "Baseline alignment requires a single text run, top anchor, and explicit font resource",
      );
    }

    const content = Yoga.Node.create(this.#config);
    content.setFlexShrink(1);
    content.setMeasureFunc(measureFunction);
    const marker = Yoga.Node.create(this.#config);
    marker.setWidth(0);
    marker.setHeight(0);
    marker.setPositionType(PositionType.Relative);
    marker.setPosition(Edge.Top, measurement.baseline);
    // Zero-size relative positioning marker only provides content baseline to Yoga; does not enter semantic layout tree or PPTX.
    marker.setIsReferenceBaseline(true);

    node.yoga.unsetMeasureFunc();
    node.yoga.setFlexDirection(FlexDirection.Row);
    node.yoga.setAlignItems(Align.FlexStart);
    node.yoga.insertChild(content, 0);
    node.yoga.insertChild(marker, 1);
    node.baseline = measurement.baseline + padding.top;
    node.baselineContent = content;
  }

  #markBaselineMeasurementsDirty(node: LayoutNode): boolean {
    let changed = false;
    if (node.baselineContent !== undefined) {
      node.baselineContent.markDirty();
      changed = true;
    }
    for (const child of node.children) {
      if (this.#markBaselineMeasurementsDirty(child)) changed = true;
    }
    return changed;
  }

  #baselineOffsetWithin(node: LayoutNode): number {
    if (node.baseline !== undefined) return node.baseline;
    if (node.children.length === 0) return node.yoga.getComputedHeight();

    // Same as Yoga Baseline.cpp: take baseline or reference child of this line first, else first item.
    const participants = new Set(this.#baselineParticipants(node));
    const baselineChild = node.children.find(
      (child) => child.yoga.isReferenceBaseline() || participants.has(child),
    );
    const child = baselineChild ?? node.children[0];
    if (child === undefined) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Layout subtree has no propagatable text baseline",
      );
    }
    if (
      node.style.flexWrap === "wrap" &&
      baselineChild !== undefined &&
      baselineChild !== node.children[0]
    ) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Cannot reliably determine text baseline row in wrapping Box",
      );
    }
    return child.yoga.getComputedTop() + this.#baselineOffsetWithin(child);
  }

  #baselineContentHeight(node: LayoutNode): number | undefined {
    const participants = this.#baselineParticipants(node);
    if (!participants.some((child) => this.#containsText(child)))
      return undefined;
    // Baseline line height is composed from max ascent and max descent, not written back as a fixed height constraint.
    let ascent = 0;
    let descent = 0;
    for (const child of participants) {
      const margin = expandedSpacing(child.style, "margin");
      const contribution = this.#minContentContribution(child, "height");
      const height = Math.max(0, contribution - margin.top - margin.bottom);
      const baseline = this.#baselineOffsetWithin(child);
      ascent = Math.max(ascent, margin.top + baseline);
      descent = Math.max(descent, margin.bottom + height - baseline);
    }
    const nonParticipantsHeight = Math.max(
      0,
      ...node.children
        .filter((child) => !participants.includes(child))
        .map((child) => this.#minContentContribution(child, "height")),
    );
    return Math.max(ascent + descent, nonParticipantsHeight);
  }

  #minContentContribution(node: LayoutNode, axis: "width" | "height"): number {
    const content = this.#minContent(node, axis) ?? 0;
    const rawDimension = node.style[axis] ?? node.element?.props[axis];
    let size = length(rawDimension, axis, node.path) ?? content;
    const minimum =
      axis === "width"
        ? styleLength(node.style.minWidth)
        : styleLength(node.style.minHeight);
    const maximum =
      axis === "width"
        ? styleLength(node.style.maxWidth)
        : styleLength(node.style.maxHeight);
    if (minimum !== undefined) size = Math.max(size, minimum);
    if (maximum !== undefined) size = Math.min(size, maximum);
    const margin = expandedSpacing(node.style, "margin");
    return (
      size +
      (axis === "width"
        ? margin.left + margin.right
        : margin.top + margin.bottom)
    );
  }

  #minContent(node: LayoutNode, axis: "width" | "height"): number | undefined {
    const isContainer =
      node.element === undefined ||
      node.element.type === "rp:box" ||
      node.element.type === "rp:group";
    if (!isContainer) {
      const padding = expandedSpacing(node.style, "padding");
      const computedWidth = node.yoga.getComputedWidth();
      const rawWidth = node.style.width ?? node.element?.props["width"];
      const measuredWidth =
        typeof rawWidth === "string" && rawWidth.endsWith("%")
          ? undefined
          : Math.max(0, computedWidth - padding.left - padding.right);
      const measured = this.#measure(
        node,
        measuredWidth === undefined ? {} : { availableWidth: measuredWidth },
      );
      const result =
        axis === "width"
          ? measured?.minContentWidth
          : measured?.minContentHeight;
      if (result === undefined) return undefined;
      return (
        result +
        (axis === "width"
          ? padding.left + padding.right
          : padding.top + padding.bottom)
      );
    }
    const horizontal = axis === "width";
    if (node.style.display === "grid") {
      const model = this.#gridModel(node, undefined, undefined);
      return (
        (horizontal ? model.width : model.height) +
        (horizontal
          ? expandedSpacing(node.style, "padding").left +
            expandedSpacing(node.style, "padding").right
          : expandedSpacing(node.style, "padding").top +
            expandedSpacing(node.style, "padding").bottom)
      );
    }
    const flexRow =
      node.style.display === "flex" &&
      rowDirection(node.style) &&
      node.style.flexWrap !== "wrap";
    const flexColumn =
      node.style.display === "flex" && !rowDirection(node.style);
    let content: number;
    if (horizontal && flexRow) {
      content = node.children.reduce(
        (total, child) => total + this.#minContentContribution(child, axis),
        0,
      );
      content +=
        (node.style.columnGap ?? node.style.gap ?? 0) *
        Math.max(0, node.children.length - 1);
    } else if (!horizontal && (flexColumn || node.style.display !== "flex")) {
      content = node.children.reduce(
        (total, child) => total + this.#minContentContribution(child, axis),
        0,
      );
      if (node.style.display === "flex") {
        content +=
          (node.style.rowGap ?? node.style.gap ?? 0) *
          Math.max(0, node.children.length - 1);
      }
    } else {
      content = Math.max(
        0,
        ...node.children.map((child) =>
          this.#minContentContribution(child, axis),
        ),
      );
    }
    if (!horizontal && flexRow) {
      content = this.#baselineContentHeight(node) ?? content;
    }
    if (horizontal) {
      const padding = expandedSpacing(node.style, "padding");
      content += padding.left + padding.right;
      const declared = styleLength(node.style.width);
      const maximum = styleLength(node.style.maxWidth);
      if (declared !== undefined) content = Math.min(content, declared);
      if (maximum !== undefined) content = Math.min(content, maximum);
    } else {
      const padding = expandedSpacing(node.style, "padding");
      content += padding.top + padding.bottom;
      const declared = styleLength(node.style.height);
      const maximum = styleLength(node.style.maxHeight);
      if (declared !== undefined) content = Math.min(content, declared);
      if (maximum !== undefined) content = Math.min(content, maximum);
    }
    return content;
  }

  #commit(
    node: LayoutNode,
    origin: { readonly x: number; readonly y: number },
  ): void {
    const box = {
      x: layoutPixel(origin.x + node.yoga.getComputedLeft()),
      y: layoutPixel(origin.y + node.yoga.getComputedTop()),
      width: layoutPixel(node.yoga.getComputedWidth()),
      height: layoutPixel(node.yoga.getComputedHeight()),
    };
    if (node.element?.type === "rp:box") {
      this.#boxes.set(node.element, box);
      if (
        box.x < 0 ||
        box.y < 0 ||
        box.x + box.width > this.#dimensions.width ||
        box.y + box.height > this.#dimensions.height
      ) {
        this.#warnings.push({
          severity: "warning",
          path: node.path,
          code: "overflow",
          message:
            "Layout container overflows slide canvas; build retains original declaration",
        });
      }
    } else if (node.element !== undefined) {
      this.#boxes.set(node.element, box);
    }
    for (const child of node.children)
      this.#commit(child, { x: box.x, y: box.y });
    for (const child of node.escaped) this.#placeEscaped(child, box);
  }
}
