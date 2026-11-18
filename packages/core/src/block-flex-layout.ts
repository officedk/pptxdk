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

import { BuildError } from "#/diagnostics.ts";
import type {
  CustomContainerSolver,
  LayoutEdges,
  LayoutNodeMeasure,
  LayoutSolver,
  LayoutSolverNode,
  LayoutSolverPort,
} from "#/layout-protocol.ts";
import type {
  Diagnostic,
  GeometryBox,
  LayoutLength,
  LayoutStyle,
} from "#/types.ts";

export const paddingStyleKeys = [
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
] as const;

const blockFlexContainerProperties = [
  "flexDirection",
  "flexWrap",
  "justifyContent",
  "alignItems",
  "alignContent",
  "gap",
  "rowGap",
  "columnGap",
] as const;

export const blockFlexLayoutProperties = [
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
  "zIndex",
] as const;

const alignItemsValues = [
  "start",
  "end",
  "center",
  "stretch",
  "flex-start",
  "flex-end",
  "baseline",
] as const;

const contentAlignmentValues = [
  "start",
  "end",
  "center",
  "flex-start",
  "flex-end",
  "space-between",
  "space-around",
  "space-evenly",
] as const;

function fail(path: string, code: string, message: string): never {
  const diagnostic: Diagnostic = { severity: "error", path, code, message };
  throw new BuildError([diagnostic]);
}

export function isStyleRecord(
  value: unknown,
): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function describeStyleValue(value: unknown): string {
  if (typeof value === "string") return `"${value}"`;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return typeof value;
}

export function parseLayoutNumber(
  value: unknown,
  name: string,
  path: string,
): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail(path, "invalid-layout-style", `${name} must be a finite pixel number`);
  }
  return value;
}

export function parseLayoutLength(
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
  const amount = parseLayoutNumber(value, name, path);
  if (nonNegative && amount < 0) {
    fail(path, "invalid-layout-style", `${name} cannot be less than 0`);
  }
  return amount;
}

export function parseLayoutSpacing(
  value: unknown,
  name: string,
  path: string,
): LayoutEdges {
  if (value === undefined) return { top: 0, right: 0, bottom: 0, left: 0 };
  if (typeof value === "number") {
    const amount = parseLayoutNumber(value, name, path);
    return { top: amount, right: amount, bottom: amount, left: amount };
  }
  if (!Array.isArray(value) || (value.length !== 2 && value.length !== 4)) {
    fail(
      path,
      "invalid-layout-style",
      `${name} must be a number, 2-tuple, or 4-tuple`,
    );
  }
  const values = value.map((item) => parseLayoutNumber(item, name, path));
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

export function expandedSpacing(
  style: LayoutStyle,
  kind: "margin" | "padding",
): LayoutEdges {
  const base = parseLayoutSpacing(style[kind], kind, "layout");
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

export function resolvePercent(
  value: LayoutLength | undefined,
  basis: number,
): number | undefined {
  if (typeof value !== "string") return undefined;
  const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u.exec(value);
  return match === null ? undefined : (Number(match[1]) * basis) / 100;
}

export function styleLength(
  value: LayoutLength | undefined,
): number | undefined {
  return typeof value === "number" ? value : undefined;
}

/**
 * Validates container-level Block/Flex display and container-only property eligibility.
 */
export function validateBlockFlexContainerContext(
  value: Record<string, unknown>,
  path: string,
  isContainer: boolean,
): "block" | "flex" | "grid" | "none" {
  if (
    !isContainer &&
    ((value["display"] !== undefined &&
      value["display"] !== "none" &&
      value["display"] !== "block") ||
      Object.keys(value).some((key) =>
        blockFlexContainerProperties.includes(
          key as (typeof blockFlexContainerProperties)[number],
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
  return display ?? "block";
}

export function validateBlockFlexContainerProperties(
  value: Record<string, unknown>,
  resolvedDisplay: "block" | "flex" | "grid" | "none",
  path: string,
): void {
  for (const key of blockFlexContainerProperties) {
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
}

export function validateBlockFlexEnumValues(
  value: Record<string, unknown>,
  path: string,
): void {
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
        `${name} does not support value ${describeStyleValue(candidate)}`,
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
  allowedValue(
    "justifyContent",
    value["justifyContent"],
    contentAlignmentValues,
  );
  allowedValue("alignItems", value["alignItems"], alignItemsValues);
  allowedValue("alignSelf", value["alignSelf"], ["auto", ...alignItemsValues]);
  allowedValue("alignContent", value["alignContent"], [
    ...contentAlignmentValues,
    "stretch",
  ]);
  allowedValue("position", value["position"], [
    "static",
    "relative",
    "absolute",
  ]);
}

export function validateBlockFlexDimensionsAndOffsets(
  value: Record<string, unknown>,
  allowsPadding: boolean,
  path: string,
): void {
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
    if (
      candidate !== undefined &&
      parseLayoutNumber(candidate, key, path) < 0
    ) {
      fail(path, "invalid-layout-style", `${key} cannot be less than 0`);
    }
  }
  for (const key of ["gap", "rowGap", "columnGap"] as const) {
    const candidate = value[key];
    if (
      candidate !== undefined &&
      parseLayoutNumber(candidate, key, path) < 0
    ) {
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
    parseLayoutLength(value[key], key, path);
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
    if (candidate !== undefined) parseLayoutLength(candidate, key, path, false);
  }
}

export function validateBlockFlexSpacingAndBounds(
  value: Record<string, unknown>,
  allowsPadding: boolean,
  path: string,
): void {
  const padding = value["padding"];
  if (padding !== undefined) {
    const values = parseLayoutSpacing(padding, "padding", path);
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
    if (
      candidate !== undefined &&
      parseLayoutNumber(candidate, key, path) < 0
    ) {
      fail(path, "invalid-layout-style", `${key} cannot be less than 0`);
    }
  }
  for (const key of [
    "marginTop",
    "marginRight",
    "marginBottom",
    "marginLeft",
  ] as const) {
    if (value[key] !== undefined) parseLayoutNumber(value[key], key, path);
  }
  if (value["margin"] !== undefined) {
    parseLayoutSpacing(value["margin"], "margin", path);
  }
  if (
    !allowsPadding &&
    paddingStyleKeys.some((key) => value[key] !== undefined)
  ) {
    fail(
      path,
      "invalid-layout-context",
      "padding is only applicable to layout containers",
    );
  }
  const minWidth = parseLayoutLength(value["minWidth"], "minWidth", path);
  const maxWidth = parseLayoutLength(value["maxWidth"], "maxWidth", path);
  const minHeight = parseLayoutLength(value["minHeight"], "minHeight", path);
  const maxHeight = parseLayoutLength(value["maxHeight"], "maxHeight", path);
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

type BlockFlexNode = {
  readonly view: LayoutSolverNode;
  readonly path: string;
  readonly style: LayoutStyle;
  readonly yoga: YogaNode;
  readonly isRoot: boolean;
  readonly isContainer: boolean;
  readonly propagatesBaseline: boolean;
  readonly contractWidth?: LayoutLength | undefined;
  readonly contractHeight?: LayoutLength | undefined;
  readonly contractX?: number | undefined;
  readonly contractY?: number | undefined;
  readonly measure?: LayoutNodeMeasure | undefined;
  readonly customHandler?: CustomContainerSolver | undefined;
  parent?: BlockFlexNode;
  readonly children: BlockFlexNode[];
  readonly escaped: BlockFlexNode[];
  measureFunction?: MeasureFunction;
  baselineContent?: YogaNode;
  baseline?: number | undefined;
};

function applyBlockFlexStyle(
  yoga: YogaNode,
  style: LayoutStyle,
  parent: LayoutStyle,
  isRoot: boolean,
  isContainer: boolean,
  isCustomContainer: boolean,
): void {
  const display = style.display ?? "block";
  yoga.setDisplay(display === "none" ? Display.None : Display.Flex);
  const positioned =
    style.position === "absolute" || style.position === "relative" || isRoot;
  yoga.setPositionType(
    style.position === "absolute"
      ? PositionType.Absolute
      : positioned
        ? PositionType.Relative
        : PositionType.Static,
  );
  if (!positioned && isCustomContainer) {
    // Delegated custom containers (such as Grid) place items with local absolute positions.
    // A static container must still form a containing block so those positions resolve locally,
    // while author absolute descendants escape to the nearest positioned ancestor before layout.
    yoga.setAlwaysFormsContainingBlock(true);
  }
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
  if ((flex || isCustomContainer) && style.gap !== undefined) {
    yoga.setGap(Gutter.All, style.gap);
  }
  if ((flex || isCustomContainer) && style.rowGap !== undefined) {
    yoga.setGap(Gutter.Row, style.rowGap);
  }
  if ((flex || isCustomContainer) && style.columnGap !== undefined) {
    yoga.setGap(Gutter.Column, style.columnGap);
  }
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
      if (style.flexBasis !== undefined) {
        setLength(yoga, "flexBasis", style.flexBasis);
      }
      if (style.alignSelf !== undefined) {
        yoga.setAlignSelf(yogaAlign(style.alignSelf));
      }
    } else {
      yoga.setFlexShrink(0);
    }
  }
  if (isContainer && style.display !== "flex") {
    yoga.setAlignItems(Align.Stretch);
  }
}

/**
 * Creates the Block & Flex layout solver backed by Yoga, handling CSS-subset block/flex defaults,
 * automatic minimum sizes, text baseline alignment & propagation, containing-block escaping,
 * loose canvas geometry, and delegated custom container collaboration.
 */
export function createBlockFlexLayoutSolver(options: {
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly customContainers?: readonly CustomContainerSolver[];
}): LayoutSolver {
  const { dimensions, customContainers = [] } = options;
  const config = Yoga.Config.create();
  config.setPointScaleFactor(0);
  const internalNodes = new WeakMap<LayoutSolverNode, BlockFlexNode>();
  let rootNode: BlockFlexNode | undefined;

  const findCustomHandler = (
    display: LayoutStyle["display"],
  ): CustomContainerSolver | undefined =>
    customContainers.find((handler) => handler.handlesDisplay(display));

  const asInternal = (node: LayoutSolverNode): BlockFlexNode => {
    const internal = internalNodes.get(node);
    if (internal === undefined) {
      throw new Error(`Unknown layout solver node: ${node.path}`);
    }
    return internal;
  };

  const containingBlock = (parent: BlockFlexNode): BlockFlexNode => {
    let current: BlockFlexNode | undefined = parent;
    while (current !== undefined) {
      if (current.isRoot) return current;
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
  };

  const minContent = (
    node: BlockFlexNode,
    axis: "width" | "height",
  ): number | undefined => {
    if (!node.isContainer) {
      const padding = expandedSpacing(node.style, "padding");
      const computedWidth = node.yoga.getComputedWidth();
      const rawWidth = node.style.width ?? node.contractWidth;
      const measuredWidth =
        typeof rawWidth === "string" && rawWidth.endsWith("%")
          ? undefined
          : Math.max(0, computedWidth - padding.left - padding.right);
      const measured = node.measure?.(
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
    if (node.customHandler !== undefined) {
      return node.customHandler.measureMinContent(node.view, axis, port);
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
        (total, child) => total + minContentContribution(child, axis),
        0,
      );
      content +=
        (node.style.columnGap ?? node.style.gap ?? 0) *
        Math.max(0, node.children.length - 1);
    } else if (!horizontal && (flexColumn || node.style.display !== "flex")) {
      content = node.children.reduce(
        (total, child) => total + minContentContribution(child, axis),
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
        ...node.children.map((child) => minContentContribution(child, axis)),
      );
    }
    if (!horizontal && flexRow) {
      content = baselineContentHeight(node) ?? content;
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
  };

  const minContentContribution = (
    node: BlockFlexNode,
    axis: "width" | "height",
  ): number => {
    const content = minContent(node, axis) ?? 0;
    const rawDimension =
      node.style[axis] ??
      (axis === "width" ? node.contractWidth : node.contractHeight);
    let size = parseLayoutLength(rawDimension, axis, node.path) ?? content;
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
  };

  const automaticMinimum = (
    node: BlockFlexNode,
    axis: "width" | "height",
  ): number | undefined => {
    const content = minContent(node, axis);
    if (content === undefined) return undefined;
    const maximum =
      axis === "width"
        ? styleLength(node.style.maxWidth)
        : styleLength(node.style.maxHeight);
    const rawDimension =
      node.style[axis] ??
      (axis === "width" ? node.contractWidth : node.contractHeight);
    const declared = parseLayoutLength(rawDimension, axis, node.path);
    return Math.min(
      content,
      ...(declared === undefined ? [] : [declared]),
      ...(maximum === undefined ? [] : [maximum]),
    );
  };

  const applyAutomaticMinimums = (root: BlockFlexNode): boolean => {
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
        ) {
          continue;
        }
        const axis = horizontal ? "width" : "height";
        const minimum = automaticMinimum(child, axis);
        if (minimum === undefined || minimum <= 0) continue;
        if (horizontal) child.yoga.setMinWidth(minimum);
        else child.yoga.setMinHeight(minimum);
        changed = true;
      }
    }
    for (const child of root.children) {
      if (applyAutomaticMinimums(child)) changed = true;
    }
    return changed;
  };

  const resolvedChildAlignment = (
    node: BlockFlexNode,
    child: BlockFlexNode,
  ): Align => {
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
  };

  const baselineParticipants = (node: BlockFlexNode): BlockFlexNode[] => {
    // Resolve participants using actual properties of layout node, ignoring raw styles dropped by non-flex parents.
    return node.children.filter(
      (child) => resolvedChildAlignment(node, child) === Align.Baseline,
    );
  };

  const hasVerticalOffset = (node: BlockFlexNode): boolean => {
    return (
      node.style.top !== undefined ||
      node.style.marginTop !== undefined ||
      node.style.bottom !== undefined ||
      node.style.marginBottom !== undefined
    );
  };

  const resolveBaselineLeaf = (node: BlockFlexNode): BlockFlexNode => {
    if (!node.isContainer || !node.propagatesBaseline) {
      if (node.measure === undefined) {
        fail(
          node.path,
          "unsupported-baseline-context",
          "Baseline alignment requires a text-bearing capability with baseline support",
        );
      }
      return node;
    }
    if (node.children.length !== 1) {
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
    const childHeight = child.contractHeight ?? child.style.height;
    const stretchedFixedHeight =
      (childAlign === undefined ||
        childAlign === "auto" ||
        childAlign === "stretch") &&
      (parentAlign === undefined || parentAlign === "stretch") &&
      typeof childHeight === "number";
    if (!simpleRow || (!alignedAtStart && !stretchedFixedHeight)) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Text baseline requires a nowrap row Box with the only child fixed at the top of the cross axis",
      );
    }
    return resolveBaselineLeaf(child);
  };

  const installBaselineMarker = (node: BlockFlexNode): void => {
    if (node.baselineContent !== undefined) return;
    const measureFunction = node.measureFunction;
    if (measureFunction === undefined || node.measure === undefined) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Baseline alignment requires a text-bearing capability with baseline support",
      );
    }
    const rawWidth = node.contractWidth ?? node.style.width;
    const availableWidth = parseLayoutLength(rawWidth, "width", node.path);
    const padding = expandedSpacing(node.style, "padding");
    const measurement = node.measure(
      availableWidth === undefined
        ? {}
        : {
            availableWidth: Math.max(
              0,
              availableWidth - padding.left - padding.right,
            ),
          },
    );
    if (measurement?.baseline === undefined) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "Baseline alignment requires a text-bearing capability with baseline support",
      );
    }
    if (measurement.baseline === "unavailable" || hasVerticalOffset(node)) {
      fail(
        node.path,
        "baseline-measurement-unavailable",
        "Baseline alignment requires a single text run, top anchor, and explicit font resource",
      );
    }

    const content = Yoga.Node.create(config);
    content.setFlexShrink(1);
    content.setMeasureFunc(measureFunction);
    const marker = Yoga.Node.create(config);
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
  };

  const prepareTextBaselines = (node: BlockFlexNode): void => {
    const flexRow = node.style.display === "flex" && rowDirection(node.style);
    const participants = baselineParticipants(node);
    if (flexRow && participants.length > 0) {
      if (node.style.flexWrap === "wrap") {
        fail(
          node.path,
          "unsupported-baseline-context",
          "Text baseline currently requires a nowrap row flex container",
        );
      }
      for (const participant of participants) {
        installBaselineMarker(resolveBaselineLeaf(participant));
      }
    }
    for (const child of node.children) prepareTextBaselines(child);
    for (const child of node.escaped) prepareTextBaselines(child);
  };

  const markBaselineMeasurementsDirty = (node: BlockFlexNode): boolean => {
    let changed = false;
    if (node.baselineContent !== undefined) {
      node.baselineContent.markDirty();
      changed = true;
    }
    for (const child of node.children) {
      if (markBaselineMeasurementsDirty(child)) changed = true;
    }
    return changed;
  };

  const baselineOffsetWithin = (node: BlockFlexNode): number => {
    if (node.baseline !== undefined) return node.baseline;
    if (node.children.length === 0) return node.yoga.getComputedHeight();

    // Same as Yoga Baseline.cpp: take baseline or reference child of this line first, else first item.
    const participants = new Set(baselineParticipants(node));
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
    return child.yoga.getComputedTop() + baselineOffsetWithin(child);
  };

  const baselineContentHeight = (node: BlockFlexNode): number | undefined => {
    const participants = baselineParticipants(node);
    if (participants.length === 0) return undefined;
    // Baseline line height is composed from max ascent and max descent, not written back as a fixed height constraint.
    let ascent = 0;
    let descent = 0;
    for (const child of participants) {
      const margin = expandedSpacing(child.style, "margin");
      const contribution = minContentContribution(child, "height");
      const height = Math.max(0, contribution - margin.top - margin.bottom);
      const baseline = baselineOffsetWithin(child);
      ascent = Math.max(ascent, margin.top + baseline);
      descent = Math.max(descent, margin.bottom + height - baseline);
    }
    const nonParticipantsHeight = Math.max(
      0,
      ...node.children
        .filter((child) => !participants.includes(child))
        .map((child) => minContentContribution(child, "height")),
    );
    return Math.max(ascent + descent, nonParticipantsHeight);
  };

  const port: LayoutSolverPort = {
    getComputedSize(node) {
      const internal = asInternal(node);
      return {
        width: internal.yoga.getComputedWidth(),
        height: internal.yoga.getComputedHeight(),
      };
    },
    setContainerSize(node, size) {
      const internal = asInternal(node);
      if (size.width !== undefined) internal.yoga.setWidth(size.width);
      if (size.height !== undefined) internal.yoga.setHeight(size.height);
    },
    measureNaturalSize(node, availableWidth) {
      const internal = asInternal(node);
      internal.yoga.calculateLayout(availableWidth, undefined, Direction.LTR);
      return {
        width: internal.yoga.getComputedWidth(),
        height: internal.yoga.getComputedHeight(),
      };
    },
    measureMinContentContribution(node, axis) {
      return minContentContribution(asInternal(node), axis);
    },
    setChildConstraints(node, constraints) {
      const internal = asInternal(node);
      if (constraints.width !== undefined) {
        internal.yoga.setWidth(constraints.width);
      }
      if (constraints.height === "auto") {
        internal.yoga.setHeightAuto();
      } else if (constraints.height !== undefined) {
        internal.yoga.setHeight(constraints.height);
      }
      if (constraints.minWidth !== undefined) {
        internal.yoga.setMinWidth(constraints.minWidth);
      }
      if (constraints.maxWidth !== undefined) {
        internal.yoga.setMaxWidth(constraints.maxWidth);
      }
      if (constraints.minHeight !== undefined) {
        internal.yoga.setMinHeight(constraints.minHeight);
      }
      if (constraints.maxHeight !== undefined) {
        internal.yoga.setMaxHeight(constraints.maxHeight);
      }
    },
    placeChild(node, position) {
      const internal = asInternal(node);
      internal.yoga.setPositionType(PositionType.Absolute);
      internal.yoga.setPosition(Edge.Left, position.x);
      internal.yoga.setPosition(Edge.Top, position.y);
    },
  };

  const resolveCustomContainers = (node: BlockFlexNode): void => {
    node.customHandler?.placeContainer(node.view, port);
    for (const child of node.children) resolveCustomContainers(child);
  };

  const placeEscaped = (
    node: BlockFlexNode,
    containingBox: GeometryBox,
    onCommit: (committed: LayoutSolverNode, box: GeometryBox) => void,
  ): void => {
    const container = Yoga.Node.create(config);
    container.setWidth(containingBox.width);
    container.setHeight(containingBox.height);
    container.setPositionType(PositionType.Relative);
    container.insertChild(node.yoga, 0);
    try {
      markBaselineMeasurementsDirty(node);
      for (let pass = 0; pass < 3; pass += 1) {
        container.calculateLayout(
          containingBox.width,
          containingBox.height,
          Direction.LTR,
        );
        resolveCustomContainers(node);
      }
      if (applyAutomaticMinimums(node)) {
        container.calculateLayout(
          containingBox.width,
          containingBox.height,
          Direction.LTR,
        );
        resolveCustomContainers(node);
        container.calculateLayout(
          containingBox.width,
          containingBox.height,
          Direction.LTR,
        );
      }
      container.calculateLayout(
        containingBox.width,
        containingBox.height,
        Direction.LTR,
      );
      commit(node, { x: containingBox.x, y: containingBox.y }, onCommit);
    } finally {
      container.removeChild(node.yoga);
      container.free();
    }
  };

  const commit = (
    node: BlockFlexNode,
    origin: { readonly x: number; readonly y: number },
    onCommit: (committed: LayoutSolverNode, box: GeometryBox) => void,
  ): void => {
    // Keep the solver's precision: capabilities compile and write the same geometry, so display rounding here would shrink native output.
    const box: GeometryBox = {
      x: origin.x + node.yoga.getComputedLeft(),
      y: origin.y + node.yoga.getComputedTop(),
      width: node.yoga.getComputedWidth(),
      height: node.yoga.getComputedHeight(),
    };
    onCommit(node.view, box);
    for (const child of node.children) {
      commit(child, { x: box.x, y: box.y }, onCommit);
    }
    for (const child of node.escaped) {
      placeEscaped(child, box, onCommit);
    }
  };

  const freeEscaped = (node: BlockFlexNode): void => {
    for (const child of node.escaped) {
      child.yoga.freeRecursive();
      freeEscaped(child);
    }
    for (const child of node.children) freeEscaped(child);
  };

  return {
    createRoot(spec) {
      const yoga = Yoga.Node.create(config);
      const customHandler = findCustomHandler(spec.style.display);
      yoga.setWidth(dimensions.width);
      yoga.setHeight(dimensions.height);
      yoga.setPositionType(PositionType.Relative);
      applyBlockFlexStyle(
        yoga,
        spec.style,
        { display: "block" },
        true,
        true,
        customHandler !== undefined,
      );
      const childViews: LayoutSolverNode[] = [];
      const view: LayoutSolverNode = {
        path: spec.path,
        style: spec.style,
        isRoot: true,
        isContainer: true,
        children: childViews,
      };
      const root: BlockFlexNode = {
        view,
        yoga,
        path: spec.path,
        style: spec.style,
        isRoot: true,
        isContainer: true,
        propagatesBaseline: true,
        customHandler,
        children: [],
        escaped: [],
      };
      internalNodes.set(view, root);
      rootNode = root;
      return view;
    },

    createNode(spec, parent) {
      const parentInternal = asInternal(parent);
      const yoga = Yoga.Node.create(config);
      const customHandler = spec.isContainer
        ? findCustomHandler(spec.style.display)
        : undefined;
      applyBlockFlexStyle(
        yoga,
        spec.style,
        parentInternal.style,
        false,
        spec.isContainer,
        customHandler !== undefined,
      );
      if (!spec.isContainer && spec.padding !== undefined) {
        yoga.setPadding(Edge.Top, spec.padding.top);
        yoga.setPadding(Edge.Right, spec.padding.right);
        yoga.setPadding(Edge.Bottom, spec.padding.bottom);
        yoga.setPadding(Edge.Left, spec.padding.left);
      }
      const width = spec.contractWidth ?? spec.style.width;
      const height = spec.contractHeight ?? spec.style.height;
      if (width !== undefined) setLength(yoga, "width", width);
      if (height !== undefined) setLength(yoga, "height", height);
      if (spec.style.position === "absolute") {
        if (
          spec.style.left === undefined &&
          spec.style.inset === undefined &&
          spec.style.right === undefined &&
          spec.contractX !== undefined
        ) {
          yoga.setPosition(Edge.Left, spec.contractX);
        }
        if (
          spec.style.top === undefined &&
          spec.style.inset === undefined &&
          spec.style.bottom === undefined &&
          spec.contractY !== undefined
        ) {
          yoga.setPosition(Edge.Top, spec.contractY);
        }
      }
      const childViews: LayoutSolverNode[] = [];
      const view: LayoutSolverNode = {
        path: spec.path,
        style: spec.style,
        isRoot: false,
        isContainer: spec.isContainer,
        ...(spec.contractWidth === undefined
          ? {}
          : { contractWidth: spec.contractWidth }),
        ...(spec.contractHeight === undefined
          ? {}
          : { contractHeight: spec.contractHeight }),
        ...(spec.padding === undefined ? {} : { padding: spec.padding }),
        children: childViews,
      };
      const node: BlockFlexNode = {
        view,
        yoga,
        path: spec.path,
        style: spec.style,
        isRoot: false,
        isContainer: spec.isContainer,
        propagatesBaseline: spec.propagatesBaseline ?? false,
        ...(spec.contractWidth === undefined
          ? {}
          : { contractWidth: spec.contractWidth }),
        ...(spec.contractHeight === undefined
          ? {}
          : { contractHeight: spec.contractHeight }),
        ...(spec.contractX === undefined ? {} : { contractX: spec.contractX }),
        ...(spec.contractY === undefined ? {} : { contractY: spec.contractY }),
        ...(spec.measure === undefined ? {} : { measure: spec.measure }),
        ...(customHandler === undefined ? {} : { customHandler }),
        parent: parentInternal,
        children: [],
        escaped: [],
      };
      internalNodes.set(view, node);
      if (spec.measure !== undefined) {
        const measure = spec.measure;
        const measureFunction: MeasureFunction = (availWidth, widthMode) => {
          const measured = measure(
            widthMode === MeasureMode.Undefined
              ? {}
              : {
                  availableWidth: availWidth,
                  widthMode:
                    widthMode === MeasureMode.Exactly ? "exactly" : "at-most",
                },
          );
          return {
            width:
              widthMode === MeasureMode.Exactly
                ? availWidth
                : (measured?.width ?? 0),
            height: measured?.height ?? 0,
          };
        };
        node.measureFunction = measureFunction;
        yoga.setMeasureFunc(measureFunction);
      }
      return view;
    },

    appendChild(parent, child) {
      const parentInternal = asInternal(parent);
      const childInternal = asInternal(child);
      childInternal.parent = parentInternal;
      if (childInternal.style.position === "absolute") {
        const cb = containingBlock(parentInternal);
        if (cb !== parentInternal) {
          cb.escaped.push(childInternal);
          return;
        }
      }
      if (
        parentInternal.customHandler !== undefined &&
        childInternal.style.position !== "absolute"
      ) {
        childInternal.yoga.setPositionType(PositionType.Absolute);
      }
      parentInternal.yoga.insertChild(
        childInternal.yoga,
        parentInternal.children.length,
      );
      parentInternal.children.push(childInternal);
      (parentInternal.view.children as LayoutSolverNode[]).push(child);
    },

    requiresLooseLayout(style) {
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
    },

    positionLooseGeometry(node) {
      const internal = asInternal(node);
      if (internal.isRoot) return;
      const x = internal.contractX ?? 0;
      const y = internal.contractY ?? 0;
      if (internal.style.position === "absolute") {
        if (
          internal.style.left === undefined &&
          internal.style.inset === undefined &&
          internal.style.right === undefined
        ) {
          internal.yoga.setPosition(Edge.Left, x);
        }
        if (
          internal.style.top === undefined &&
          internal.style.inset === undefined &&
          internal.style.bottom === undefined
        ) {
          internal.yoga.setPosition(Edge.Top, y);
        }
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
        internal.style.left ?? internal.style.inset,
        internal.style.right ?? internal.style.inset,
        dimensions.width,
      );
      const top = offset(
        internal.style.top ?? internal.style.inset,
        internal.style.bottom ?? internal.style.inset,
        dimensions.height,
      );
      internal.yoga.setPositionType(PositionType.Absolute);
      internal.yoga.setPositionAuto(Edge.Right);
      internal.yoga.setPositionAuto(Edge.Bottom);
      internal.yoga.setPosition(Edge.Left, x + left);
      internal.yoga.setPosition(Edge.Top, y + top);
      internal.yoga.setAlwaysFormsContainingBlock(true);
    },

    solve(onCommit) {
      if (rootNode === undefined) {
        throw new Error("Layout solver root has not been initialized");
      }
      const root = rootNode;
      try {
        prepareTextBaselines(root);
        // Fixed-size parent nodes may skip creating a new measure tree; solve initially, then mark measurement leaves dirty and recalculate.
        root.yoga.calculateLayout(
          dimensions.width,
          dimensions.height,
          Direction.LTR,
        );
        for (let pass = 0; pass < 3; pass += 1) {
          resolveCustomContainers(root);
          root.yoga.calculateLayout(
            dimensions.width,
            dimensions.height,
            Direction.LTR,
          );
        }
        if (markBaselineMeasurementsDirty(root)) {
          root.yoga.calculateLayout(
            dimensions.width,
            dimensions.height,
            Direction.LTR,
          );
        }
        if (applyAutomaticMinimums(root)) {
          root.yoga.calculateLayout(
            dimensions.width,
            dimensions.height,
            Direction.LTR,
          );
          resolveCustomContainers(root);
          root.yoga.calculateLayout(
            dimensions.width,
            dimensions.height,
            Direction.LTR,
          );
        }
        commit(root, { x: 0, y: 0 }, onCommit);
      } finally {
        root.yoga.freeRecursive();
        freeEscaped(root);
        config.free();
      }
    },
  };
}
