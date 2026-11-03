import type { ReactNode } from "react";

import type { MorphKey } from "#/morph.ts";
import type {
  ThemeContext,
  ThemeInput,
  ThemeScope,
  TextStyle,
  VisualStyle,
} from "#/theme.ts";

export type {
  ColorValue,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  GradientStop,
  ImageVisualStyle,
  LinearGradient,
  FontToken,
  NativeColorSlot,
  ResolvedColor,
  ResolvedThemeFont,
  TextAlignment,
  TextStyle,
  VisualStyle,
  RadialGradient,
  Shadow,
  Theme,
  ThemeColor,
  ThemeContext,
  ThemeInput,
  ThemeScope,
  ThemeShade,
} from "#/theme.ts";

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
  /** 明确请求嵌入文稿实际使用的完整字体；默认不嵌入。 */
  readonly embedFonts?: "full";
  /**
   * 单次构建的协作式取消信号；未指定时没有默认超时。
   * 完成、未处理错误与取消以首次确定的终态为准，完成后取消不会改变结果。
   */
  readonly signal?: AbortSignal;
};

export type GeometryBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/** 固定画布长度；数值按像素解释，百分比按所属布局区域求值。 */
export type LayoutLength = number | `${number}%` | "auto";

export type LayoutSpacing =
  | number
  | readonly [vertical: number, horizontal: number]
  | readonly [top: number, right: number, bottom: number, left: number];

/** 固定画布 CSS 布局子集；数值长度按像素解释。 */
export type LayoutStyle = {
  readonly display?: "block" | "flex" | "grid" | "none";
  readonly position?: "static" | "relative" | "absolute";
  readonly inset?: LayoutLength;
  readonly left?: LayoutLength;
  readonly top?: LayoutLength;
  readonly right?: LayoutLength;
  readonly bottom?: LayoutLength;
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
  readonly aspectRatio?: number;
  readonly gridTemplateColumns?: string;
  readonly gridTemplateRows?: string;
  readonly gridColumn?: number | string;
  readonly gridRow?: number | string;
  readonly justifyItems?:
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end";
  readonly justifySelf?:
    | "auto"
    | "start"
    | "end"
    | "center"
    | "stretch"
    | "flex-start"
    | "flex-end";
  readonly zIndex?: number;
};

/** 透明布局容器；它只安排子项，不要求生成对应的原生演示对象。 */
export type BoxProps = ElementProps & {
  readonly testId?: string;
  readonly morphKey?: MorphKey;
  readonly style?: LayoutStyle & TextStyle & VisualStyle;
  readonly theme?: ThemeScope;
};

export type ObjectObservation = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly geometry?: string;
  readonly box?: GeometryBox;
  readonly text?: string;
  readonly table?: NativeTableObservation;
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

/** 表格报告保留零基行列索引，与其原生合并范围一一对应。 */
export type NativeTableObservation = {
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
  /** 被合并覆盖的格子留空，合并锚点保存单元格正文。 */
  readonly cells: readonly (readonly string[])[];
  readonly merges: readonly {
    readonly row: number;
    readonly column: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  }[];
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
  /** 原生对象声明才带有此身份；透明能力声明不生成对象 ID。 */
  readonly targetId?: number;
  /** 页面宿主树中的前序位置，用于保留透明声明和对象成员顺序。 */
  readonly order: number;
};

export type CapabilityResource = {
  readonly bytes: Uint8Array;
  readonly extension: string;
  readonly contentType: string;
  readonly relationshipType: string;
};

/** 一个页面原生部件，以及它通过部件关系引用的二进制资源。 */
export type CapabilityPackagePart = {
  readonly part: {
    /** 包内相对路径模板，必须恰含一个 `{n}` 部件序号。 */
    readonly pathTemplate: string;
    readonly contentType: string;
  };
  readonly relationshipType: string;
  /** 接收相关部件关系标识，按 relatedParts 顺序对应。 */
  readonly content: (
    relatedRelationshipIds: readonly string[],
  ) => string | Uint8Array;
  readonly relatedParts: readonly {
    readonly pathTemplate: string;
    readonly contentType: string;
    readonly relationshipType: string;
    readonly bytes: Uint8Array;
  }[];
};

export type CapabilityContext = {
  readonly id: number;
  readonly path: string;
  /** 当前元素可用的文稿主题、局部颜色作用域与容器文字默认值。 */
  readonly theme: ThemeContext;
  /** 本次构建选项；能力模块自行解释其负责的选项。 */
  readonly options: BuildOptions;
  readonly layoutBox?: GeometryBox;
  /** 将显式连续性句柄转换成 OOXML 对象名称；未声明连续性的对象不带此值。 */
  readonly morphName?: string;
  /** 使用同一组已注册模块递归编译原生子对象。 */
  compileChildren(children: readonly CapabilityChild[]): CapabilityChildren;
  /** 将二进制资源加入 PPTX，并返回当前页面对应的关系标识。 */
  registerResource(resource: CapabilityResource): string;
  /** 注册当前页面引用的原生部件及其直接关联的二进制部件。 */
  registerPackagePart(part: CapabilityPackagePart): string;
};

/** 一个文稿级 XML 记录中的资源项，由核心统一分配部件与关系标识。 */
export type CapabilityPresentationContribution = {
  readonly collection: string;
  readonly presentationAttributes?: Readonly<Record<string, string>>;
  readonly record: {
    readonly key: string;
    readonly identity: string;
    readonly tagName: string;
    readonly content: string;
  };
  readonly item: {
    readonly key: string;
    readonly identity: string;
    /** 同一记录中兄弟关系项的模式顺序。 */
    readonly order: number;
    readonly relationshipType: string;
    readonly pathTemplate: string;
    readonly contentType: string;
    readonly bytes: Uint8Array;
    readonly content: (relationshipId: string) => string;
  };
};

export type CapabilityLayoutRequest = {
  /** 当前布局算法可提供的内容宽度，单位为像素。 */
  readonly availableWidth?: number;
  /** 可用宽度是可收缩的上限，还是布局必须分配的精确宽度。 */
  readonly widthMode?: "at-most" | "exactly";
  /** 度量与对象写出共享的主题、局部作用域和容器文字默认值。 */
  readonly theme: ThemeContext;
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
  /** 有可靠文字度量时，首行基线距对象顶边的像素距离。 */
  readonly baseline?: number;
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
    /** 来自共享文字服务的实际输出文字，供后续字体资源处理使用。 */
    readonly characters?: string;
  }[];
  /** 由对象实际使用的资源构成、需要挂接到 presentation 的记录项。 */
  readonly presentationContributions?: readonly CapabilityPresentationContribution[];
  readonly warnings?: readonly {
    readonly code: string;
    readonly message: string;
    /** 相对原生对象的子路径，用于精确定位单元格诊断。 */
    readonly path?: string;
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

/** 不生成原生对象、只向所属页面能力提交声明的透明宿主节点。 */
export type TransparentCapabilityModule = {
  /** 声明组件生成的宿主类型，例如 `rp:motion-scope`。 */
  readonly declaration: string;
  compileDeclaration(
    input: CapabilityInput,
    context: { readonly path: string },
  ): readonly CapabilityAnnotation[];
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
  /** 由该能力自行接线、供内部透明声明宿主使用的处理器。 */
  readonly transparentDeclarations?: readonly TransparentCapabilityModule[];
  compileSlide(
    input: CapabilityInput,
    context: SlideCapabilityContext,
  ): SlideCapabilityOutput;
};

export type CapabilityModule = ObjectCapabilityModule | SlideCapabilityModule;

/** Builder 内部注册的模块，透明声明只能由页面能力模块拥有。 */
export type RegisteredCapabilityModule =
  | CapabilityModule
  | TransparentCapabilityModule;

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
  readonly theme?: ThemeInput;
  readonly style?: TextStyle;
};

export type SlideProps = ElementProps & {
  readonly slideId: string;
  readonly transition?: unknown;
  readonly theme?: ThemeScope;
  readonly style?: TextStyle & Omit<LayoutStyle, "zIndex">;
};
