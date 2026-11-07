import type { ReactNode } from "react";

import type { MorphKey } from "#/morph.ts";
import type {
  ThemeContext,
  ThemeInput,
  ThemeScope,
  TextStyle,
  VisualStyle,
  ColorValue,
  LinearGradient,
  RadialGradient,
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

/** Slide background can use theme colors, gradients supported by the shared visual writer, or formal image resource inputs. */
export type SlideBackground =
  | ColorValue
  | LinearGradient
  | RadialGradient
  | { readonly image: string | URL | Uint8Array };

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
  /** Explicitly request embedding full files or character subsets of fonts actually used by the deck; no embedding by default. */
  readonly embedFonts?: "full" | "subset";
  /**
   * Cooperative cancellation signal for a single build; no default timeout when unspecified.
   * Completion, unhandled errors, and cancellation follow the first established terminal state; cancelling after completion does not alter the result.
   */
  readonly signal?: AbortSignal;
};

export type GeometryBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

/** Fixed canvas length; numeric values are interpreted in pixels, percentages are evaluated against the parent layout region. */
export type LayoutLength = number | `${number}%` | "auto";

export type LayoutSpacing =
  | number
  | readonly [vertical: number, horizontal: number]
  | readonly [top: number, right: number, bottom: number, left: number];

/** Fixed canvas CSS layout subset; numeric lengths are interpreted in pixels. */
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

/** Transparent layout Box; it arranges children only and does not require generating a corresponding native presentation object. */
export type BoxProps = ElementProps & {
  readonly testId?: string;
  readonly part?: string;
  readonly morphKey?: MorphKey;
  readonly link?: LinkTarget;
  readonly style?: LayoutStyle & TextStyle & VisualStyle;
  readonly theme?: ThemeScope;
};

/** Target declaration interpreted by PowerPoint upon clicking an object; internal targets are resolved by stable slideId. */
export type LinkTarget =
  | { readonly slideId: string }
  | { readonly url: string }
  | { readonly action: "home" };

/** Link properties registered by core as relationships or slide actions for native object capabilities to write out. */
export type NativeLink = {
  readonly relationshipId?: string;
  readonly action?: string;
};

export type ObjectObservation = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  readonly testId?: string;
  readonly part?: string;
  readonly geometry?: string;
  readonly box?: GeometryBox;
  readonly text?: string;
  readonly link?: string;
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
  readonly part?: string;
  readonly box: GeometryBox;
  readonly children: readonly LayoutChildReference[];
};

/** Table report retains zero-based row and column indexes, corresponding one-to-one with native merge extents. */
export type NativeTableObservation = {
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
  /** Cells covered by a merge remain empty, while merge anchors retain cell body text. */
  readonly cells: readonly (readonly string[])[];
  readonly merges: readonly {
    readonly row: number;
    readonly column: number;
    readonly rowSpan: number;
    readonly colSpan: number;
  }[];
  /** Identities and final geometries of rows and cells are submitted by the table module in shared layout results. */
  readonly parts?: readonly NativeTablePartObservation[];
};

export type NativeTablePartObservation = {
  readonly kind: "row" | "cell";
  readonly testId?: string;
  readonly part?: string;
  readonly row: number;
  readonly column?: number;
  readonly rowSpan?: number;
  readonly colSpan?: number;
  readonly box: GeometryBox;
  readonly text: string;
};

export type SlideObservation = {
  readonly id: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
  readonly objects: readonly ObjectObservation[];
  readonly layouts?: readonly LayoutObservation[];
  readonly capabilities?: Readonly<Record<string, unknown>>;
};

export type BuildReport = {
  readonly evidence: "generation";
  readonly title?: string;
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
  readonly query: BuildQuery;
};

export type LocatorObservation = {
  readonly evidence: "generation";
  readonly testId?: string;
  readonly part?: string;
  readonly box?: GeometryBox;
  readonly text?: string;
  readonly objects: readonly ObjectObservation[];
  readonly resources?: readonly LocatorResourceObservation[];
};

export type ResourceObservation = {
  readonly part: string;
  readonly contentType: string;
  readonly byteLength: number;
};

export type LocatorResourceObservation = ResourceObservation & {
  readonly objectId: number;
};

/** Read-only locator and field observation for formal content in a successful build. */
export type Locator = {
  getByTestId(testId: string): Locator;
  getAllByTestId(testId: string): readonly Locator[];
  part(name: string): Locator;
  parts(name: string): readonly Locator[];
  children(): readonly Locator[];
  descendants(): readonly Locator[];
  observe(): LocatorObservation;
};

export type SlideQuery = Locator & {
  readonly id: string;
  readonly index: number;
  readonly observation: SlideObservation;
};

export type BuildQuery = {
  readonly evidence: "generation";
  readonly slides: readonly SlideQuery[];
  slide(slideId: string): SlideQuery;
};

/** Generic node in the React host tree, without exposing React or renderer internal objects. */
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

/** Declaration data submitted by capability modules alongside native objects for consumption by slide modules of the same name. */
export type CapabilityAnnotation = {
  readonly capability: string;
  readonly value: unknown;
};

/** Declaration data passed to slide modules after core supplements object identity and diagnostic locations. */
export type CapabilityContribution = CapabilityAnnotation & {
  readonly path: string;
  /** Only native object declarations carry this identity; transparent capability declarations do not generate object IDs. */
  readonly targetId?: number;
  /** Preorder index in the slide host tree to preserve ordering of transparent declarations and object members. */
  readonly order: number;
};

export type CapabilityResource = {
  readonly bytes: Uint8Array;
  readonly extension: string;
  readonly contentType: string;
  readonly relationshipType: string;
};

/** A slide native part and binary resources it references via part relationships. */
export type CapabilityPackagePart = {
  readonly part: {
    /** Relative package path template, which must contain exactly one `{n}` part sequence number. */
    readonly pathTemplate: string;
    readonly contentType: string;
  };
  readonly relationshipType: string;
  /** Receives related part relationship IDs corresponding to relatedParts order. */
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
  /** Deck theme, local color scope, and container text defaults available to the current element. */
  readonly theme: ThemeContext;
  /** Options for the current build; capability modules interpret their respective options. */
  readonly options: BuildOptions;
  /** 1-based slide number in the deck, used by shared text service to write slide number fields. */
  readonly slideNumber: number;
  /** Allocate stable and unique IDs for native text fields in the current deck. */
  readonly createTextFieldId: () => string;
  /** Object click relationship resolved and registered by core, written by capability into p:cNvPr. */
  readonly nativeLink?: NativeLink;
  readonly layoutBox?: GeometryBox;
  /** Convert an explicit continuity handle into an OOXML object name; objects without declared continuity omit this value. */
  readonly morphName?: string;
  /** Recursively compile native child objects using the same registered modules. */
  compileChildren(children: readonly CapabilityChild[]): CapabilityChildren;
  /** Add binary resource to the PPTX package and return the relationship ID for the current slide. */
  registerResource(resource: CapabilityResource): string;
  /** Register a native part referenced by the current slide and its directly associated binary parts. */
  registerPackagePart(part: CapabilityPackagePart): string;
};

/** A resource item within a presentation-level XML record, with part and relationship IDs assigned uniformly by core. */
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
    /** Schema order of sibling relationship items within the same record. */
    readonly order: number;
    readonly relationshipType: string;
    readonly pathTemplate: string;
    readonly contentType: string;
    readonly bytes?: Uint8Array;
    /** Points to the corresponding shared resource claim when awaiting completion of deck character collection. */
    readonly resourceClaim?: { readonly kind: string; readonly key: string };
    /** Called once after core aggregates characters for this resource claim, returning final part bytes to write. */
    readonly resolveBytes?: (characters: string) => Uint8Array;
    readonly content: (relationshipId: string) => string;
  };
};

export type CapabilityLayoutRequest = {
  /** Content width provided by the current layout algorithm in pixels. */
  readonly availableWidth?: number;
  /** Whether available width is a shrinkable upper bound or an exact width the layout must allocate. */
  readonly widthMode?: "at-most" | "exactly";
  /** Theme, local scope, and container text defaults shared between measurement and object compilation. */
  readonly theme: ThemeContext;
  /** 1-based slide number in the deck, used by shared text service to measure native slide number fields. */
  readonly slideNumber: number;
};

/** Intrinsic size and main-axis min-content size provided by capability modules, in pixels. */
export type CapabilityLayoutMeasurement = {
  /** Natural width after typesetting at the requested width. */
  readonly width?: number;
  /** Natural height after typesetting at the requested width. */
  readonly height?: number;
  /** Minimum content width used when computing automatic minimum size along the width axis. */
  readonly minContentWidth?: number;
  /** Minimum content height used when computing automatic minimum size along the height axis. */
  readonly minContentHeight?: number;
  /** Distance in pixels from the object top edge to the first line baseline when reliable text measurement is available. */
  readonly baseline?: number;
};

export type CapabilityOutput = {
  /** A Transitional PresentationML object written as a child of the slide shape tree. */
  readonly xml: string;
  readonly observation: ObjectObservation;
  readonly extensions: readonly Extension[];
  /** Object-level declarations compiled by the corresponding slide capability, such as motion. */
  readonly annotations?: readonly CapabilityAnnotation[];
  readonly resourceClaims?: readonly CapabilityResourceClaim[];
  /** Record items composed of resources actually used by objects that must attach to the presentation. */
  readonly presentationContributions?: readonly CapabilityPresentationContribution[];
  readonly warnings?: readonly {
    readonly code: string;
    readonly message: string;
    /** Subpath relative to the native object, used for precisely locating cell diagnostics. */
    readonly path?: string;
  }[];
};

/** Claim provided by capability to deck-scoped resource processing; character content is generated by the shared text service. */
export type CapabilityResourceClaim = {
  readonly kind: string;
  readonly key: string;
  readonly identity: string;
  /** Actual output text from the shared text service, used for subsequent font resource processing. */
  readonly characters?: string;
};

export type ObjectCapabilityModule = {
  /** Host type of the declared component, such as `rp:shape`. */
  readonly declaration: string;
  /** Provides measurement basis using the capability's own content measurement when layout requires intrinsic or main-axis min-content size. */
  measureLayout?(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined;
  compile(input: CapabilityInput, context: CapabilityContext): CapabilityOutput;
};

/** Transparent host node that does not generate native objects, submitting declarations only to its parent slide capability. */
export type TransparentCapabilityModule = {
  /** Host type generated by the declared component, such as `rp:motion-scope`. */
  readonly declaration: string;
  compileDeclaration(
    input: CapabilityInput,
    context: { readonly path: string },
  ): readonly CapabilityAnnotation[];
};

export type SlideCapabilitySlot = "background" | "transition" | "timing";

export type SlideCapabilityContext = {
  readonly path: string;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly contributions: readonly CapabilityContribution[];
  readonly theme: ThemeContext;
  readonly registerResource: (resource: CapabilityResource) => string;
};

export type SlideCapabilityOutput = {
  /** Formatted content placed into the slide property region by core. */
  readonly xml: string;
  readonly extensions: readonly Extension[];
  readonly observation?: unknown;
};

export type SlideCapabilityModule = {
  /** Slide property name handled by this module. */
  readonly slideProp: string;
  /** Slide properties jointly consumed by the same module, such as transition and advance timing. */
  readonly slideProps?: readonly string[];
  /** Written to the slide following PresentationML fixed child element order. */
  readonly slot: SlideCapabilitySlot;
  /** Runs once regardless of whether Slide provides the corresponding property. */
  readonly always?: boolean;
  /** Stable capability key in reports. */
  readonly declaration: string;
  /** Handlers wired by this capability for internal transparent declaration hosts. */
  readonly transparentDeclarations?: readonly TransparentCapabilityModule[];
  compileSlide(
    input: CapabilityInput,
    context: SlideCapabilityContext,
  ): SlideCapabilityOutput;
};

export type CapabilityModule = ObjectCapabilityModule | SlideCapabilityModule;

/** Modules registered inside Builder; transparent declarations can only be owned by slide capability modules. */
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
  readonly title?: string;
};

export type SlideProps = ElementProps & {
  readonly slideId: string;
  readonly transition?: unknown;
  readonly theme?: ThemeScope;
  readonly style?: TextStyle & Omit<LayoutStyle, "zIndex">;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
  readonly background?: SlideBackground;
};

/** Organizes slides into collapsible flat sections in PowerPoint. */
export type SectionProps = ElementProps & {
  readonly name: string;
};
