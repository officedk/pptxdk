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
  /**
   * 单次构建的协作式取消信号；未指定时没有默认超时。
   * 完成、未处理错误与取消以首次确定的终态为准，完成后取消不会改变结果。
   */
  readonly signal?: AbortSignal;
};

export type Box = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

export type ObjectObservation = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly geometry?: string;
  readonly box?: Box;
  readonly text?: string;
  readonly table?: NativeTableObservation;
  readonly members?: readonly ObjectObservation[];
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
  /** 使用同一组已注册模块递归编译原生子对象。 */
  compileChildren(children: readonly CapabilityChild[]): CapabilityChildren;
  /** 将二进制资源加入 PPTX，并返回当前页面对应的关系标识。 */
  registerResource(resource: CapabilityResource): string;
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
    /** 相对原生对象的子路径，用于精确定位单元格诊断。 */
    readonly path?: string;
  }[];
};

export type ObjectCapabilityModule = {
  /** 声明组件的宿主类型，例如 `rp:shape`。 */
  readonly declaration: string;
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
