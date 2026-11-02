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
};

export type SlideObservation = {
  readonly id: string;
  readonly index: number;
  readonly objects: readonly ObjectObservation[];
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

export type CapabilityInput = {
  readonly type: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly unknown[];
};

export type CapabilityContext = {
  readonly id: number;
  readonly path: string;
};

export type CapabilityOutput = {
  /** 一个 Transitional PresentationML 对象，作为正式 slide shape tree 的子节点写入。 */
  readonly xml: string;
  readonly observation: ObjectObservation;
  readonly extensions: readonly Extension[];
};

export type CapabilityModule = {
  /** 声明组件的宿主类型，例如 `rp:shape`。 */
  readonly declaration: string;
  compile(input: CapabilityInput, context: CapabilityContext): CapabilityOutput;
};

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
};
