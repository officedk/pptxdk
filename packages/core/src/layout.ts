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
import type {
  CapabilityChild,
  CapabilityElement,
  CapabilityLayoutMeasurement,
  CapabilityModule,
  Diagnostic,
  GeometryBox,
  LayoutLength,
  LayoutStyle,
  ObjectCapabilityModule,
} from "#/types.ts";

type Dimensions = { readonly width: number; readonly height: number };
type Edges = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};
type LayoutNode = {
  readonly yoga: YogaNode;
  readonly element?: CapabilityElement;
  readonly style: LayoutStyle;
  readonly path: string;
  readonly module?: ObjectCapabilityModule;
  readonly children: LayoutNode[];
  /** 文字需要提供基线时，由外部对象节点转交给内部测量叶节点。 */
  measureFunction?: MeasureFunction;
  baselineContent?: YogaNode;
  baseline?: number;
};

export type LayoutPassResult = {
  readonly boxes: WeakMap<CapabilityElement, GeometryBox>;
  readonly warnings: readonly Diagnostic[];
};

const layoutProperties = [
  "display",
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
] as const;

const layoutPropertySet = new Set<string>(layoutProperties);
const textProperties = new Set([
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
]);
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

function record(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
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

// 这里只允许共享文字样式属性通过布局白名单；具体校验仍由能力模块负责。
function allowsSharedTextStyle(type: string): boolean {
  return type === "rp:text" || type === "rp:table" || type === "rp:chart";
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
    fail(path, "invalid-layout-style", `${name} 必须是有限像素数值`);
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
  const amount = number(value, name, path);
  if (nonNegative && amount < 0) {
    fail(path, "invalid-layout-style", `${name} 不能小于 0`);
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
    fail(path, "invalid-layout-style", `${name} 必须是数字、二元组或四元组`);
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
): LayoutStyle {
  if (value === undefined) return {};
  if (!record(value)) {
    fail(path, "invalid-layout-style", "style 必须是对象");
  }
  const unknown = Object.keys(value).find(
    (key) =>
      !layoutPropertySet.has(key) && !(isText && textProperties.has(key)),
  );
  if (unknown !== undefined) {
    fail(path, "unsupported-layout-style", `style 含未实现属性 ${unknown}`);
  }
  if (
    !isContainer &&
    (value["display"] !== undefined ||
      Object.keys(value).some((key) =>
        containerProperties.includes(
          key as (typeof containerProperties)[number],
        ),
      ))
  ) {
    fail(
      path,
      "invalid-layout-context",
      "容器专属样式不能用于没有布局子项的对象",
    );
  }
  const display = value["display"];
  if (display !== undefined && display !== "block" && display !== "flex") {
    fail(path, "unsupported-layout-value", "display 只支持 block 或 flex");
  }
  const resolvedDisplay =
    (display as LayoutStyle["display"] | undefined) ?? "block";
  for (const key of containerProperties) {
    if (value[key] !== undefined && resolvedDisplay !== "flex") {
      fail(
        path,
        "invalid-layout-context",
        `${key} 只适用于 display: "flex" 的容器`,
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
        `${name} 不支持取值 ${describe(candidate)}`,
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
  for (const key of ["flex", "flexGrow", "flexShrink"] as const) {
    const candidate = value[key];
    if (candidate !== undefined && number(candidate, key, path) < 0) {
      fail(path, "invalid-layout-style", `${key} 不能小于 0`);
    }
  }
  for (const key of ["gap", "rowGap", "columnGap"] as const) {
    const candidate = value[key];
    if (candidate !== undefined && number(candidate, key, path) < 0) {
      fail(path, "invalid-layout-style", `${key} 不能小于 0`);
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
  const padding = value["padding"];
  if (padding !== undefined) {
    const values = spacing(padding, "padding", path);
    if (Object.values(values).some((item) => item < 0)) {
      fail(path, "invalid-layout-style", "padding 不能小于 0");
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
      fail(path, "invalid-layout-style", `${key} 不能小于 0`);
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
    !isContainer &&
    [
      "padding",
      "paddingTop",
      "paddingRight",
      "paddingBottom",
      "paddingLeft",
    ].some((key) => value[key] !== undefined)
  ) {
    fail(path, "invalid-layout-context", "padding 只适用于布局容器");
  }
  const minWidth = length(value["minWidth"], "minWidth", path);
  const maxWidth = length(value["maxWidth"], "maxWidth", path);
  const minHeight = length(value["minHeight"], "minHeight", path);
  const maxHeight = length(value["maxHeight"], "maxHeight", path);
  if (minWidth !== undefined && maxWidth !== undefined && minWidth > maxWidth) {
    fail(path, "invalid-layout-style", "minWidth 不能大于 maxWidth");
  }
  if (
    minHeight !== undefined &&
    maxHeight !== undefined &&
    minHeight > maxHeight
  ) {
    fail(path, "invalid-layout-style", "minHeight 不能大于 maxHeight");
  }
  return value as LayoutStyle;
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

function layoutPixel(value: number): number {
  // Yoga 以单精度浮点保存尺寸；按文字服务精度归一到百分之一像素。
  return Math.round(value * 100) / 100;
}

function objectModule(
  module: CapabilityModule | undefined,
): module is ObjectCapabilityModule {
  return module !== undefined && "compile" in module;
}

/** 在能力模块写出原生对象前求解 Box、块与弹性布局。 */
export class LayoutPass {
  readonly #modules: ReadonlyMap<string, CapabilityModule>;
  readonly #dimensions: Dimensions;
  readonly #config = Yoga.Config.create();
  readonly #boxes = new WeakMap<CapabilityElement, GeometryBox>();
  readonly #warnings: Diagnostic[] = [];

  constructor(
    modules: ReadonlyMap<string, CapabilityModule>,
    dimensions: Dimensions,
  ) {
    this.#modules = modules;
    this.#dimensions = dimensions;
    this.#config.setPointScaleFactor(0);
  }

  run(
    children: readonly CapabilityChild[],
    parentPath: string,
  ): LayoutPassResult {
    const rootYoga = Yoga.Node.create(this.#config);
    const root: LayoutNode = {
      yoga: rootYoga,
      style: { display: "block" },
      path: parentPath,
      children: [],
    };
    rootYoga.setDisplay(Display.Flex);
    rootYoga.setFlexDirection(FlexDirection.Column);
    rootYoga.setWidth(this.#dimensions.width);
    rootYoga.setHeight(this.#dimensions.height);
    try {
      this.#appendChildren(children, parentPath, root, false);
      this.#prepareTextBaselines(root);
      // 固定尺寸父节点可能跳过新建测量树；先初次求解，再标脏测量叶并重算。
      rootYoga.calculateLayout(
        this.#dimensions.width,
        this.#dimensions.height,
        Direction.LTR,
      );
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
      }
      this.#commit(root, { x: 0, y: 0 });
    } finally {
      rootYoga.freeRecursive();
      this.#config.free();
    }
    return { boxes: this.#boxes, warnings: this.#warnings };
  }

  #validateLoose(input: CapabilityElement, path: string): void {
    const module = this.#modules.get(input.type);
    if (!objectModule(module)) return;
    // Text、Table 与 Chart 以共享文字服务消费字体属性，布局中仍是叶对象。
    const allowsTextStyle = allowsSharedTextStyle(input.type);
    readStyle(input.props["style"], path, false, allowsTextStyle);
  }

  #appendChildren(
    children: readonly CapabilityChild[],
    parentPath: string,
    parent: LayoutNode,
    includeObjects: boolean,
  ): void {
    children.forEach((child, index) => {
      if (!element(child)) return;
      const childPath = `${parentPath}/${label(child, index)}`;
      if (child.type === "rp:box") {
        const built = this.#buildBox(child, childPath, parent.style);
        parent.yoga.insertChild(built.yoga, parent.children.length);
        parent.children.push(built);
        return;
      }
      if (child.type === "rp:group") {
        // Group 只保留原生成员关系；其内容继续使用外层布局上下文。
        this.#appendChildren(child.children, childPath, parent, includeObjects);
        return;
      }
      if (!includeObjects) {
        this.#validateLoose(child, childPath);
        return;
      }
      const module = this.#modules.get(child.type);
      if (!objectModule(module)) return;
      const allowsTextStyle = allowsSharedTextStyle(child.type);
      const childStyle = readStyle(
        child.props["style"],
        childPath,
        false,
        allowsTextStyle,
      );
      const childYoga = Yoga.Node.create(this.#config);
      const layoutChild: LayoutNode = {
        yoga: childYoga,
        element: child,
        style: childStyle,
        path: childPath,
        module,
        children: [],
      };
      this.#applyStyle(childYoga, childStyle, parent.style, false, false);
      this.#setDeclaredSize(layoutChild);
      if (module.measureLayout !== undefined) {
        const measureFunction: MeasureFunction = (width, widthMode) => {
          const measured = this.#measure(
            layoutChild,
            widthMode === MeasureMode.Undefined
              ? {}
              : { availableWidth: width },
          );
          return {
            width:
              widthMode === MeasureMode.Exactly
                ? width
                : (measured?.width ?? 0),
            height: measured?.height ?? 0,
          };
        };
        layoutChild.measureFunction = measureFunction;
        childYoga.setMeasureFunc(measureFunction);
      }
      parent.yoga.insertChild(childYoga, parent.children.length);
      parent.children.push(layoutChild);
    });
  }

  #buildBox(
    input: CapabilityElement,
    path: string,
    parentStyle: LayoutStyle,
  ): LayoutNode {
    const unknown = Object.keys(input.props).find(
      (key) => key !== "testId" && key !== "style",
    );
    if (unknown !== undefined) {
      fail(path, "unknown-parameter", `Box 含未实现参数 ${unknown}`);
    }
    const testId = input.props["testId"];
    if (
      testId !== undefined &&
      (typeof testId !== "string" || testId.trim() === "")
    ) {
      fail(path, "invalid-parameter", "testId 必须是非空字符串");
    }
    const style = readStyle(input.props["style"], path, true, false);
    const yoga = Yoga.Node.create(this.#config);
    const node: LayoutNode = {
      yoga,
      element: input,
      style,
      path,
      children: [],
    };
    this.#applyStyle(yoga, style, parentStyle, false, true);
    this.#appendChildren(input.children, path, node, true);
    return node;
  }

  #setDeclaredSize(node: LayoutNode): void {
    const props = node.element?.props ?? {};
    const style = node.style;
    const rawWidth = style.width ?? props["width"];
    const width = length(rawWidth, "width", node.path);
    const rawHeight = style.height ?? props["height"];
    const height =
      rawHeight === "auto" ? undefined : length(rawHeight, "height", node.path);
    if (width !== undefined) node.yoga.setWidth(width);
    if (height !== undefined) node.yoga.setHeight(height);
  }

  #applyStyle(
    yoga: YogaNode,
    style: LayoutStyle,
    parent: LayoutStyle,
    isRoot: boolean,
    isContainer: boolean,
  ): void {
    yoga.setDisplay(Display.Flex);
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
      // Yoga 默认从起始边堆叠弹性行；CSS 默认会拉伸有剩余空间的行。
      yoga.setAlignContent(Align.Stretch);
    }
    if (flex && style.gap !== undefined) yoga.setGap(Gutter.All, style.gap);
    if (flex && style.rowGap !== undefined)
      yoga.setGap(Gutter.Row, style.rowGap);
    if (flex && style.columnGap !== undefined)
      yoga.setGap(Gutter.Column, style.columnGap);
    for (const key of [
      "width",
      "height",
      "minWidth",
      "minHeight",
      "maxWidth",
      "maxHeight",
    ] as const) {
      const amount = styleLength(style[key]);
      if (amount === undefined) continue;
      const setter = {
        width: "setWidth",
        height: "setHeight",
        minWidth: "setMinWidth",
        minHeight: "setMinHeight",
        maxWidth: "setMaxWidth",
        maxHeight: "setMaxHeight",
      }[key] as keyof YogaNode;
      (yoga[setter] as (value: number) => void)(amount);
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
          const basis = styleLength(style.flexBasis);
          if (basis === undefined) yoga.setFlexBasisAuto();
          else yoga.setFlexBasis(basis);
        }
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
    request: { readonly availableWidth?: number },
  ): CapabilityLayoutMeasurement | undefined {
    if (
      node.module?.measureLayout === undefined ||
      node.element === undefined
    ) {
      return undefined;
    }
    try {
      return node.module.measureLayout(node.element, request);
    } catch (error) {
      if (error instanceof BuildError) throw error;
      const code =
        error instanceof CapabilityError
          ? error.code
          : "layout-measurement-failed";
      const message = error instanceof Error ? error.message : String(error);
      fail(node.path, code, message);
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
        if (typeof explicitMinimum === "number") continue;
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
    // 使用已应用到 Yoga 的属性，复现其列向 baseline 到 flex-start 的回退。
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
    if (node.style.display !== "flex") return [];
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
        "文字 baseline 目前只沿单子项 Box 链传播",
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
        "文字 baseline 需要不换行的行 Box，且唯一子项的交叉轴位置固定在顶部",
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
          "文字 baseline 目前要求不换行的行弹性容器",
        );
      }
      for (const participant of participants) {
        if (!this.#containsText(participant)) continue;
        this.#installBaselineMarker(this.#baselineTextLeaf(participant));
      }
    }
    for (const child of node.children) this.#prepareTextBaselines(child);
  }

  #installBaselineMarker(node: LayoutNode): void {
    if (node.baselineContent !== undefined) return;
    const measureFunction = node.measureFunction;
    if (measureFunction === undefined) {
      fail(
        node.path,
        "baseline-measurement-unavailable",
        "baseline 对齐需要文字能力提供可靠的字体度量",
      );
    }
    const rawWidth = node.style.width ?? node.element?.props["width"];
    const availableWidth = length(rawWidth, "width", node.path);
    const measurement = this.#measure(
      node,
      availableWidth === undefined ? {} : { availableWidth },
    );
    if (measurement?.baseline === undefined) {
      fail(
        node.path,
        "baseline-measurement-unavailable",
        "baseline 对齐需要单一文字 run、top anchor 与显式字体资源",
      );
    }

    const content = Yoga.Node.create(this.#config);
    content.setFlexShrink(1);
    content.setMeasureFunc(measureFunction);
    const marker = Yoga.Node.create(this.#config);
    marker.setWidth(0);
    marker.setHeight(0);
    marker.setPositionType(PositionType.Relative);
    marker.setPosition(
      Edge.Top,
      measurement.baseline - expandedSpacing(node.style, "padding").top,
    );
    // 零尺寸相对定位标记只向 Yoga 提供内容基线，不进入语义布局树或 PPTX。
    marker.setIsReferenceBaseline(true);

    node.yoga.unsetMeasureFunc();
    node.yoga.setFlexDirection(FlexDirection.Row);
    node.yoga.setAlignItems(Align.FlexStart);
    node.yoga.insertChild(content, 0);
    node.yoga.insertChild(marker, 1);
    node.baseline = measurement.baseline;
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

    // 与 Yoga Baseline.cpp 相同：先取本行的基线或参考子项，否则取首项。
    const participants = new Set(this.#baselineParticipants(node));
    const baselineChild = node.children.find(
      (child) => child.yoga.isReferenceBaseline() || participants.has(child),
    );
    const child = baselineChild ?? node.children[0];
    if (child === undefined) {
      fail(
        node.path,
        "unsupported-baseline-context",
        "布局子树没有可传播的文字 baseline",
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
        "无法在换行的 Box 中可靠确定文字 baseline 所在行",
      );
    }
    return child.yoga.getComputedTop() + this.#baselineOffsetWithin(child);
  }

  #baselineContentHeight(node: LayoutNode): number | undefined {
    const participants = this.#baselineParticipants(node);
    if (!participants.some((child) => this.#containsText(child)))
      return undefined;
    // 基线行高按最大上升部与最大下降部合成，不回写为固定高度约束。
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
    if (node.element?.type !== "rp:box") {
      const measured = this.#measure(node, {
        availableWidth: node.yoga.getComputedWidth(),
      });
      const result =
        axis === "width"
          ? measured?.minContentWidth
          : measured?.minContentHeight;
      return result;
    }
    const horizontal = axis === "width";
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
          message: "布局容器超出页面画布，构建保留原始声明",
        });
      }
    } else if (node.element !== undefined) {
      this.#boxes.set(node.element, box);
    }
    for (const child of node.children)
      this.#commit(child, { x: box.x, y: box.y });
  }
}
