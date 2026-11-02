import type { ReactNode } from "react";

export const formatBaseline = "ECMA-376 5th edition Transitional" as const;

export type Extension = "p14" | "p159";

export type Target = {
  readonly name: string;
  readonly format: typeof formatBaseline;
  readonly extensions: readonly Extension[];
};

export const targets = Object.freeze({
  transitional: Object.freeze({
    name: "transitional",
    format: formatBaseline,
    extensions: Object.freeze(["p14", "p159"] as const),
  }),
  baselineOnly: Object.freeze({
    name: "baseline-only",
    format: formatBaseline,
    extensions: Object.freeze([] as const),
  }),
});

export type BuildOptions = {
  readonly target?: Target;
};

export type GeometryBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/** 固定画布长度；数值按像素解释，auto 由布局容器结合内容决定。 */
export type LayoutLength = number | "auto";

export type LayoutSpacing =
  | number
  | readonly [vertical: number, horizontal: number]
  | readonly [top: number, right: number, bottom: number, left: number];

/** 固定画布支持的块与弹性样式子集，不包含网格或绝对定位。 */
export type LayoutStyle = {
  readonly display?: "block" | "flex";
  readonly width?: LayoutLength;
  readonly height?: LayoutLength;
  readonly minWidth?: LayoutLength;
  readonly minHeight?: LayoutLength;
  readonly maxWidth?: LayoutLength;
  readonly maxHeight?: LayoutLength;
  readonly margin?: LayoutSpacing;
  readonly marginTop?: number;
  readonly marginRight?: number;
  readonly marginBottom?: number;
  readonly marginLeft?: number;
  readonly padding?: LayoutSpacing;
  readonly paddingTop?: number;
  readonly paddingRight?: number;
  readonly paddingBottom?: number;
  readonly paddingLeft?: number;
  readonly gap?: number;
  readonly rowGap?: number;
  readonly columnGap?: number;
  readonly flexDirection?: "row" | "column" | "row-reverse" | "column-reverse";
  readonly flexWrap?: "nowrap" | "wrap";
  readonly justifyContent?:
    | "start"
    | "end"
    | "center"
    | "flex-start"
    | "flex-end"
    | "space-between"
    | "space-around"
    | "space-evenly";
  readonly alignItems?:
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | "baseline";
  readonly alignSelf?:
    | "auto"
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end"
    | "baseline";
  readonly alignContent?:
    | "start"
    | "end"
    | "center"
    | "flex-start"
    | "flex-end"
    | "space-between"
    | "space-around"
    | "space-evenly"
    | "stretch";
  readonly flex?: number;
  readonly flexGrow?: number;
  readonly flexShrink?: number;
  readonly flexBasis?: LayoutLength;
};

/** 透明布局容器；它只安排子项，不要求生成对应的原生演示对象。 */
export type BoxProps = ElementProps & {
  readonly testId?: string;
  readonly style?: LayoutStyle;
};

export type ObjectObservation = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly geometry?: string;
  readonly box?: GeometryBox;
  readonly text?: string;
  readonly members?: readonly ObjectObservation[];
};

export type LayoutChildReference = {
  readonly kind: "layout" | "object";
  readonly id: number;
};

export type LayoutObservation = {
  readonly id: number;
  readonly testId?: string;
  readonly box: GeometryBox;
  readonly children: readonly LayoutChildReference[];
};

export type SlideObservation = {
  readonly id: string;
  readonly index: number;
  readonly objects: readonly ObjectObservation[];
  readonly layouts?: readonly LayoutObservation[];
  readonly capabilities?: Readonly<Record<string, unknown>>;
};

export type BuildReport = {
  readonly evidence: "generation";
  readonly target: {
    readonly name: string;
    readonly format: typeof formatBaseline;
    readonly extensions: readonly Extension[];
  };
  readonly extensions: readonly Extension[];
  readonly slides: readonly SlideObservation[];
  readonly diagnostics: readonly Diagnostic[];
};

export type BuildResult = {
  readonly pptx: Uint8Array;
  readonly report: BuildReport;
};

/** React 宿主树中的通用节点，不暴露 React 或渲染器内部对象。 */
export type CapabilityElement = {
  readonly type: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly CapabilityChild[];
};

export type CapabilityChild =
  | CapabilityElement
  | { readonly type: "#text"; readonly text: string };

export type CapabilityInput = CapabilityElement;

export type CapabilityChildren = {
  readonly xml: string;
  readonly objects: readonly ObjectObservation[];
};

/** 能力模块随原生对象提交、供同名页面模块消费的声明数据。 */
export type CapabilityAnnotation = {
  readonly capability: string;
  readonly value: unknown;
};

/** 核心补充对象身份和诊断位置后传给页面模块的声明数据。 */
export type CapabilityContribution = CapabilityAnnotation & {
  readonly path: string;
  readonly targetId: number;
};

export type CapabilityResource = {
  readonly bytes: Uint8Array;
  readonly extension: string;
  readonly contentType: string;
  readonly relationshipType: string;
};

export type CapabilityContext = {
  readonly id: number;
  readonly path: string;
  readonly layoutBox?: GeometryBox;
  /** 使用同一组已注册模块递归编译原生子对象。 */
  compileChildren(children: readonly CapabilityChild[]): CapabilityChildren;
  /** 将二进制资源加入 PPTX，并返回当前页面对应的关系标识。 */
  registerResource(resource: CapabilityResource): string;
};

export type CapabilityLayoutRequest = {
  /** 当前布局算法可提供的内容宽度，单位为像素。 */
  readonly availableWidth?: number;
};

/** 能力模块可提供的固有尺寸及主轴 min-content 尺寸，均以像素表示。 */
export type CapabilityLayoutMeasurement = {
  /** 按请求宽度排版后的自然宽度。 */
  readonly width?: number;
  /** 按请求宽度排版后的自然高度。 */
  readonly height?: number;
  /** 计算宽度方向自动最小尺寸时采用的最小内容宽度。 */
  readonly minContentWidth?: number;
  /** 计算高度方向自动最小尺寸时采用的最小内容高度。 */
  readonly minContentHeight?: number;
};

export type CapabilityOutput = {
  /** 一个 Transitional PresentationML 对象，作为正式 slide shape tree 的子节点写入。 */
  readonly xml: string;
  readonly observation: ObjectObservation;
  readonly extensions: readonly Extension[];
  /** 由对应页面能力编译的对象级声明，例如动效。 */
  readonly annotations?: readonly CapabilityAnnotation[];
  readonly resourceClaims?: readonly {
    readonly kind: string;
    readonly key: string;
    readonly identity: string;
  }[];
  readonly warnings?: readonly {
    readonly code: string;
    readonly message: string;
  }[];
};

export type ObjectCapabilityModule = {
  /** 声明组件的宿主类型，例如 `rp:shape`。 */
  readonly declaration: string;
  /** 布局需要固有尺寸或主轴最小内容尺寸时，使用能力自身的内容度量提供依据。 */
  measureLayout?(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined;
  compile(input: CapabilityInput, context: CapabilityContext): CapabilityOutput;
};

export type SlideCapabilitySlot = "transition" | "timing";

export type SlideCapabilityContext = {
  readonly path: string;
  readonly contributions: readonly CapabilityContribution[];
};

export type SlideCapabilityOutput = {
  /** 由核心放入页面属性区的格式内容。 */
  readonly xml: string;
  readonly extensions: readonly Extension[];
  readonly observation?: unknown;
};

export type SlideCapabilityModule = {
  /** 该模块负责的 Slide 属性名。 */
  readonly slideProp: string;
  /** 按 PresentationML 的固定子元素顺序写入页面。 */
  readonly slot: SlideCapabilitySlot;
  /** 无论 Slide 是否提供对应属性，均运行一次。 */
  readonly always?: boolean;
  /** 报告中的稳定能力键。 */
  readonly declaration: string;
  compileSlide(
    input: CapabilityInput,
    context: SlideCapabilityContext,
  ): SlideCapabilityOutput;
};

export type CapabilityModule = ObjectCapabilityModule | SlideCapabilityModule;

export type Diagnostic = {
  readonly severity: "error" | "warning";
  readonly path: string;
  readonly code: string;
  readonly message: string;
};

export type ElementProps = {
  readonly children?: ReactNode;
};

export type DeckProps = ElementProps & {
  readonly width?: number;
  readonly height?: number;
};

export type SlideProps = ElementProps & {
  readonly slideId: string;
  readonly transition?: unknown;
};
