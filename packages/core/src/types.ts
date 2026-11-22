import type { LinkTarget } from "#/author-schemas.ts";
import type { LayoutStyle } from "#/layout-schemas.ts";
import type { NativeMechanismDefinition } from "#/mechanisms.ts";
import type { MorphKey } from "#/morph.ts";
import {
  formatBaseline,
  targets,
  type BuildOptions,
  type Diagnostic,
  type DiagnosticSeverity,
  type EmitOptions,
  type Extension,
  type FontEmbeddingMode,
  type ResolveOptions,
  type Target,
  type TargetName,
} from "#/stage-options.ts";
import type { ThemeContext } from "#/theme.ts";

export type { AuthorId } from "#/author-ids.ts";
export type {
  DeckProps,
  LinkTarget,
  SectionProps,
  SlideBackground,
  SlideProps,
  SlideStyle,
} from "#/author-schemas.ts";
export type {
  LayoutLength,
  LayoutSpacing,
  LayoutStyle,
  PaddingStyle,
} from "#/layout-schemas.ts";

export type {
  ColorValue,
  Font,
  FontFace,
  FontFaceStyle,
  FontOptions,
  FontReference,
  FontSize,
  GradientStop,
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

export { formatBaseline, targets };
export type {
  BuildOptions,
  EmitOptions,
  Extension,
  FontEmbeddingMode,
  ResolveOptions,
  Target,
  TargetName,
};

// SlideBackground is derived from slideBackgroundSchema in author-schemas.ts.

export type GeometryBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};

// LayoutLength, LayoutSpacing, LayoutStyle, PaddingStyle, and LinkTarget are derived from Zod schemas.

/** Link properties registered by core as relationships or slide actions for native object capabilities to write out. */
export type NativeLink = {
  readonly relationshipId?: string;
  readonly action?: string;
};

export type ObjectObservation = {
  readonly id: number;
  readonly kind: string;
  readonly name: string;
  /** Author id declared on the source node; numeric `id` remains the build observation identity. */
  readonly authorId?: string;
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
  readonly authorId?: string;
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
  readonly authorId?: string;
  readonly row: number;
  readonly column?: number;
  readonly rowSpan?: number;
  readonly colSpan?: number;
  readonly box: GeometryBox;
  readonly text: string;
};

export type SlideObservation = {
  /** Page author id; also the stable identifier written to the slide part. */
  readonly id: string;
  /** Author id declared on the Slide, retained as the explicit author observation for pages. */
  readonly authorId?: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden?: boolean;
  readonly objects: readonly ObjectObservation[];
  readonly layouts?: readonly LayoutObservation[];
  readonly capabilities?: Readonly<Record<string, unknown>>;
};

export type ResolutionReport = {
  readonly evidence: "generation";
  readonly title?: string;
  readonly extensions: readonly Extension[];
  readonly slides: readonly SlideObservation[];
  readonly diagnostics: readonly Diagnostic[];
};

export type Resolution = {
  readonly report: ResolutionReport;
  readonly query: BuildQuery;
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

export type EmitResult = {
  readonly pptx: Uint8Array;
  readonly report: BuildReport;
};

export type LocatorObservation = {
  readonly evidence: "generation";
  readonly authorId?: string;
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
  /** Finds one author id in this locator scope without crossing the next named scope. */
  get(id: string): Locator;
  /** Lists matches for an author id in this locator scope; duplicates are already rejected by the build. */
  getAll(id: string): readonly Locator[];
  /** Direct structural child locators, including anonymous and named nodes. */
  children(): readonly Locator[];
  descendants(): readonly Locator[];
  observe(): LocatorObservation;
};

export type SlideQuery = Locator & {
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly observation: SlideObservation;
};

export type BuildQuery = {
  readonly evidence: "generation";
  readonly slides: readonly SlideQuery[];
  /** Page lookup by author id in this build. */
  getSlide(id: string): SlideQuery;
};

/** Generic node in the React host tree, without exposing React or renderer internal objects. */
export type CapabilityElement = {
  readonly type: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly children: readonly CapabilityChild[];
  /** Page-affiliated declaration that remains part of the page below hidden canvas parents. */
  readonly pageAffiliated?: boolean;
};

export type CapabilityChild =
  | CapabilityElement
  | { readonly type: "#text"; readonly text: string };

export type CapabilityInput = CapabilityElement;

/** Declaration data submitted by capability modules alongside native objects for consumption by slide modules of the same name. */
export type CapabilityAnnotation = {
  readonly capability: string;
  readonly value: unknown;
  /**
   * Annotation that may remain unconsumed when its capability's writer is trimmed.
   *
   * Object capabilities emit optional local declarations (for example an object awaiting an inherited
   * default effect). Capability-owned contribution resolution runs before missing-capability checks and
   * turns declarations with actual effects into required contributions; required annotations keep the
   * existing missing-capability diagnostic.
   */
  readonly optional?: boolean;
  /**
   * Resolver owned by the capability that produced this annotation.
   *
   * Producers attach it to every local declaration, including optional ones, so the formal build finds
   * the resolver from effective contributions rather than from registered modules or their order. The
   * build applies it once per capability before slide output and missing-capability checks, so a trimmed
   * writer does not change how the remaining declarations are interpreted.
   */
  readonly resolveContributions?: ContributionResolver;
};

/** Declaration data passed to slide modules after core supplements object identity and diagnostic locations. */
export type CapabilityContribution = CapabilityAnnotation & {
  readonly path: string;
  /** Only native object declarations carry this identity; transparent capability declarations do not generate object IDs. */
  readonly targetId?: number;
  /** Preorder index in the slide host tree to preserve ordering of transparent declarations and object members. */
  readonly order: number;
  /** Build-local declaration view that produced this contribution; never retained on an editable node. */
  readonly declaration?: CapabilityElement;
};

/**
 * Derives effective annotations for one capability from the current build tree.
 *
 * The build calls this before slide output and missing-capability checks, with only that capability's
 * contributions. The result must contain the effective replacements; hiding an annotation here also
 * means it no longer makes a trimmed capability mandatory.
 */
export type ContributionResolver = (
  children: readonly CapabilityChild[],
  contributions: readonly CapabilityContribution[],
) => readonly CapabilityContribution[];

export type CapabilityResource =
  | {
      readonly bytes: Uint8Array;
      readonly extension: string;
      readonly contentType: string;
      readonly relationshipType: string;
    }
  | {
      readonly externalTarget: string;
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

/** A package relationship target resolved by core while packaging a capability-provided part. */
export type CapabilityPartRelationshipTarget =
  | { readonly kind: "slide" }
  | { readonly kind: "part"; readonly key: string }
  | { readonly kind: "path"; readonly path: string };

export type CapabilityPartRelationship = {
  readonly relationshipType: string;
  readonly target: CapabilityPartRelationshipTarget;
};

/** A slide-owned part whose content, content type, and relationship declarations stay with the capability. */
export type CapabilitySlidePackagePart = {
  readonly part: {
    /** Relative package path template, which must contain exactly one `{n}` part sequence number. */
    readonly pathTemplate: string;
    readonly contentType: string;
  };
  /** Relationship type from the current slide part to this part. */
  readonly slideRelationshipType: string;
  readonly relationships?: readonly CapabilityPartRelationship[];
  /** Receives relationship IDs corresponding to relationships order. */
  readonly content: (relationshipIds: readonly string[]) => string | Uint8Array;
};

/**
 * A deck-shared part registered once per build, such as a notes master.
 *
 * Core allocates its path and presentation relationship. The optional presentation element carries
 * the relationship ID so the capability owns the XML element while core owns schema placement.
 */
export type CapabilitySharedPackagePart = {
  readonly part: {
    readonly pathTemplate: string;
    readonly contentType: string;
  };
  /**
   * Relationship type from the presentation part to this part.
   *
   * Omit for a shared part that is only referenced by another capability part, such as a second
   * theme component. Such a part still receives a unique package path and can be targeted by a
   * `{ kind: "part", key }` relationship, but it has no presentation-level relationship.
   */
  readonly presentationRelationshipType?: string;
  readonly relationships?: readonly CapabilityPartRelationship[];
  readonly presentationElement?: {
    readonly position: "before-slides" | "after-notes-size";
    readonly content: (relationshipId: string) => string;
  };
  readonly content: (relationshipIds: readonly string[]) => string | Uint8Array;
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

export type NativeElementKind =
  | "sp"
  | "pic"
  | "grpSp"
  | "graphicFrame"
  | "cxnSp";

export type CapabilityLayoutContract = {
  readonly nativeKind: NativeElementKind;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number | `${number}%` | "auto";
  readonly height?: number | `${number}%` | "auto";
  readonly layoutStyle: LayoutStyle;
  readonly padding?: {
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
    readonly left: number;
  };
};

export type CapabilityParseContext = {
  readonly path: string;
  readonly theme: ThemeContext;
  readonly options: ResolveOptions;
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
  /** Distance in pixels from the object top edge to the first line baseline when reliable text measurement is available, or `"unavailable"` when the object bears text whose baseline cannot be measured. */
  readonly baseline?: number | "unavailable";
};

/**
 * Generic sub-part observation node contributed by a composite object capability
 * for read-only locator traversal and observation without core knowing capability-specific part kinds.
 */
export type CapabilitySubPartObservation = {
  readonly authorId?: string;
  readonly box?: GeometryBox;
  readonly text?: string;
  readonly parts?: readonly CapabilitySubPartObservation[];
};

/** Claim provided by capability to deck-scoped resource processing; character content is generated by the shared text service. */
export type CapabilityResourceClaim = {
  readonly kind: string;
  readonly key: string;
  readonly identity: string;
  /** Actual output text from the shared text service, used for subsequent font resource processing. */
  readonly characters?: string;
};

export type PreviewTransform = {
  readonly box: GeometryBox;
  readonly angle: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
};

export type PreviewWarning = {
  readonly message: string;
  readonly path?: string;
};

/**
 * Generic native object collaboration facts submitted by a bound object capability.
 * Core coordinates Morph continuity, explicit name conflicts, and click relationships without reading leaf props.
 */
export type BoundObjectNativeCollaboration = {
  readonly morphKey?: MorphKey;
  /** True when the capability claims an explicit native object name, conflicting with morphKey. */
  readonly hasExplicitName?: boolean;
  /** Whether this bound instance produces a native slide object; defaults to true when omitted. */
  readonly emitsObject?: boolean;
  readonly link?: LinkTarget;
};

/**
 * Generic container collaboration facts submitted by a bound object capability.
 * Core coordinates child layout/compilation, scoped theme propagation, layout observation,
 * baseline propagation, and native object qualification without hardcoding container names.
 */
export type BoundContainerCollaboration = {
  /** Child declarations owned by this layout container. */
  readonly children: readonly CapabilityChild[];
  /** Scoped theme and accumulated text defaults applied to descendants and container compilation. */
  readonly theme?: ThemeContext;
  /** True when this container produces a native PresentationML object rather than remaining layout-only. */
  readonly emitsObject: boolean;
  /**
   * Controls whether this container enters the Yoga layout tree when placed inside a loose parent
   * without declaring non-text layout styles. Defaults to `"always"`.
   */
  readonly layoutMode?: "always" | "when-styled";
  /**
   * Controls whether resolved children remain siblings of the container or are enclosed as native
   * members inside `bound.resolve` via `context.resolveChildren`. Defaults to `"siblings"`.
   */
  readonly childComposition?: "siblings" | "members";
  /** True when members inside this container cannot declare continuity keys (e.g. native group members). Defaults to false. */
  readonly restrictsMemberContinuity?: boolean;
  /** True when this container contributes a layout-region observation (`kind: "layout"`). Defaults to true. */
  readonly observesLayout?: boolean;
  /** True when single-child flex rows propagate text baseline through this container. Defaults to true. */
  readonly propagatesBaseline?: boolean;
  /** True when overflowing the slide canvas emits a layout container overflow warning. Defaults to true. */
  readonly checksCanvasOverflow?: boolean;
};

export type CapabilityWarning = {
  readonly code: string;
  readonly message: string;
  /** Subpath relative to the native object, used for precisely locating cell diagnostics. */
  readonly path?: string;
};

export type ResolvedLink = {
  readonly target: LinkTarget;
  readonly summary: string;
};

export type ResolvedObjectChildren = {
  readonly objects: readonly ObjectObservation[];
};

export type ObjectResolveContext = {
  readonly id: number;
  readonly path: string;
  /** Deck theme, local color scope, and container text defaults available to the current element. */
  readonly theme: ThemeContext;
  readonly options: ResolveOptions;
  readonly slideNumber: number;
  readonly layoutBox?: GeometryBox;
  /** Convert an explicit continuity handle into an OOXML object name; objects without declared continuity omit this value. */
  readonly morphName?: string;
  /** Link target already validated during resolve before relationship ID allocation in emit. */
  readonly link?: ResolvedLink;
  /** Recursively resolve member child objects using the same registered modules. */
  resolveChildren(children: readonly CapabilityChild[]): ResolvedObjectChildren;
};

export type EmitContext = {
  readonly id: number;
  readonly path: string;
  readonly theme: ThemeContext;
  readonly options: EmitOptions;
  readonly slideNumber: number;
  readonly nativeLink?: NativeLink;
  readonly childrenXml: string;
  readonly createTextFieldId: () => string;
  registerResource(resource: CapabilityResource): string;
  registerPackagePart(part: CapabilityPackagePart): string;
};

export type EmitOutput = {
  readonly xml: string;
  readonly presentationContributions?: readonly CapabilityPresentationContribution[];
};

export type PaintContext = {
  readonly path: string;
  readonly slideNumber: number;
  readonly objectId: number;
  readonly box?: GeometryBox;
  readonly theme: ThemeContext;
  readonly childrenSvg?: string;
};

export type PaintOutput = {
  readonly svg: string;
  readonly warnings?: readonly PreviewWarning[];
};

export type ResolvedObject = {
  readonly observation: ObjectObservation;
  readonly transform?: PreviewTransform;
  readonly parts?: readonly CapabilitySubPartObservation[];
  readonly extensions: readonly Extension[];
  readonly annotations?: readonly CapabilityAnnotation[];
  readonly resourceClaims?: readonly CapabilityResourceClaim[];
  readonly warnings?: readonly CapabilityWarning[];
  emit(context: EmitContext): EmitOutput;
  paint?(context: PaintContext): PaintOutput;
};

/**
 * Per-build bound object instance whose private parsed state remains encapsulated inside the module closure.
 */
export type BoundObjectCapability = {
  readonly contract: CapabilityLayoutContract;
  readonly native?: BoundObjectNativeCollaboration;
  readonly container?: BoundContainerCollaboration;
  measureLayout?(
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement | undefined;
  resolve(context: ObjectResolveContext): ResolvedObject;
};

/**
 * Object capability module that binds per-build declaration state into a closed handle without exposing private parsed state.
 */
export type ObjectCapabilityModule<
  Bound extends BoundObjectCapability = BoundObjectCapability,
> = {
  readonly declaration: string;
  readonly mechanisms?: readonly NativeMechanismDefinition[];
  bind(input: CapabilityInput, context: CapabilityParseContext): Bound;
};

/** Alias for the unified bound object capability contract. */
export type BoundObjectCapabilityModule = ObjectCapabilityModule;

/** Transparent host node that does not generate native objects, submitting declarations only to its parent slide capability. */
export type TransparentCapabilityModule = {
  /** Host type generated by the declared component, such as `motion-scope`. */
  readonly declaration: string;
  /**
   * Whether the parent slide capability traverses this declaration's children as ordinary content.
   *
   * `"owned"` means the capability consumes the subtree itself: children never become layout nodes,
   * native objects, paint layers, or motion targets. Core still validates author ids and reports the
   * declaration to its owner, and owned declarations are collected even below hidden canvas parents.
   */
  readonly children?: "traverse" | "owned";
  compileDeclaration(
    input: CapabilityInput,
    context: {
      readonly path: string;
      readonly theme?: ThemeContext;
      readonly slideNumber?: number;
    },
  ): readonly CapabilityAnnotation[];
};

export type SlideCapabilitySlot = "background" | "transition" | "timing";

export type SlideCapabilityParseContext = {
  readonly path: string;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly theme: ThemeContext;
  readonly options: ResolveOptions;
  readonly slideNumber: number;
  readonly contributions: readonly CapabilityContribution[];
};

/**
 * Generic cross-slide continuity collaboration facts submitted by a bound slide capability.
 */
export type BoundSlideContinuityCollaboration = {
  /**
   * When true, core validates that each MorphKey object declared on this slide
   * has the same native element kind as matching objects on reachable source slides.
   */
  readonly matchSourceKinds?: boolean;
};

export type SlideResolveContext = {
  readonly path: string;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly contributions: readonly CapabilityContribution[];
  readonly children: readonly CapabilityChild[];
  readonly theme: ThemeContext;
  readonly options: ResolveOptions;
  readonly slideNumber: number;
};

export type SlideEmitContext = {
  readonly path: string;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly theme: ThemeContext;
  readonly options: EmitOptions;
  readonly slideNumber: number;
  readonly createTextFieldId: () => string;
  readonly registerResource: (resource: CapabilityResource) => string;
  readonly registerPackagePart: (part: CapabilitySlidePackagePart) => string;
  readonly registerSharedPart: (
    key: string,
    part: CapabilitySharedPackagePart,
  ) => { readonly path: string; readonly relationshipId?: string };
};

export type SlideEmitOutput = {
  readonly xml: string;
  readonly presentationContributions?: readonly CapabilityPresentationContribution[];
};

export type SlidePaintContext = {
  readonly path: string;
  readonly dimensions: { readonly width: number; readonly height: number };
  readonly theme: ThemeContext;
  readonly slideNumber: number;
};

export type SlidePaintOutput = {
  readonly svg: string;
  readonly warnings?: readonly PreviewWarning[];
};

export type ResolvedSlide = {
  readonly observation?: unknown;
  readonly extensions: readonly Extension[];
  readonly resourceClaims?: readonly CapabilityResourceClaim[];
  readonly warnings?: readonly CapabilityWarning[];
  emit(context: SlideEmitContext): SlideEmitOutput;
  paint?(context: SlidePaintContext): SlidePaintOutput;
};

export type BoundSlideCapability = {
  readonly continuity?: BoundSlideContinuityCollaboration;
  resolveSlide(context: SlideResolveContext): ResolvedSlide;
};

/** Slide capability module that binds per-build slide declaration state into a closed handle. */
export type SlideCapabilityModule<
  Bound extends BoundSlideCapability = BoundSlideCapability,
> = {
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
  readonly mechanisms?: readonly NativeMechanismDefinition[];
  bindSlide(
    input: CapabilityInput,
    context: SlideCapabilityParseContext,
  ): Bound;
};

/** Alias for the unified bound slide capability contract. */
export type BoundSlideCapabilityModule = SlideCapabilityModule;

export type CapabilityModule = ObjectCapabilityModule | SlideCapabilityModule;

/** Modules registered inside a resolver; transparent declarations can only be owned by slide capability modules. */
export type RegisteredCapabilityModule =
  | CapabilityModule
  | TransparentCapabilityModule;

export type { Diagnostic, DiagnosticSeverity };

// DeckProps, SlideProps, and SectionProps are derived from Zod schemas in author-schemas.ts.
