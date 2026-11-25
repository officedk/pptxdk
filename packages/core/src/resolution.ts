/**
 * Internal backend protocol for reading opaque `Resolution` snapshots.
 *
 * Solely for internal stage backend packages (`@pptxdk/emit` and `@pptxdk/preview`);
 * this subpath is not an author API and must never be re-exported from `.` or `./document`.
 */
import type { HeaderFooter, SlideLayoutType } from "#/masters.ts";
import type { MorphKey } from "#/morph.ts";
import type {
  MasterPlaceholderMap,
  MasterTextStyles,
  PlaceholderMap,
} from "#/placeholders.ts";
import type { QueryContentNode } from "#/query.ts";
import type {
  ColorMap,
  ColorMapOverride,
  Theme,
  ThemeContext,
  ThemeOverride,
} from "#/theme.ts";
import type {
  Diagnostic,
  EmitContext,
  EmitOutput,
  EmittedSlideLayoutSpec,
  EmittedSlideMasterSpec,
  Extension,
  GeometryBox,
  MastersEmitDriver,
  PaintContext,
  PaintOutput,
  PptxRelationship,
  PreviewTransform,
  Resolution,
  ResolvedLink,
  SlideCapabilityModule,
  SlideCapabilitySlot,
  SlideEmitContext,
  SlideEmitOutput,
  SlidePaintContext,
  SlidePaintOutput,
} from "#/types.ts";

export type {
  EmittedSlideLayoutSpec,
  EmittedSlideMasterSpec,
  MastersEmitDriver,
  PptxRelationship,
};

export type ResolutionDimensions = {
  readonly width: number;
  readonly height: number;
};

export type ResolvedSectionRecord = {
  readonly name: string;
  readonly slideIndexes: readonly number[];
};

export type ResolvedObjectLayer = {
  readonly declarationType: string;
  readonly sourceOrder: readonly number[];
  readonly zIndex: number;
  readonly path: string;
  readonly objectId: number;
  readonly paintObjectId: number;
  readonly theme: ThemeContext;
  readonly box?: GeometryBox;
  readonly transform?: PreviewTransform;
  readonly emitsObject: boolean;
  readonly hasObservation: boolean;
  readonly childComposition: "siblings" | "members" | "leaf";
  readonly resolvedLink?: ResolvedLink;
  readonly morphKey?: MorphKey;
  readonly extensions: readonly Extension[];
  readonly memberLayers: readonly ResolvedObjectLayer[];
  readonly siblingLayers: readonly ResolvedObjectLayer[];
  readonly emit: (context: EmitContext) => EmitOutput;
  readonly paint?: (context: PaintContext) => PaintOutput;
};

export type ResolvedSlideCapabilityRecord = {
  readonly declaration: string;
  readonly slot: SlideCapabilitySlot;
  readonly hasSlideProps: boolean;
  readonly path: string;
  readonly theme: ThemeContext;
  readonly extensions: readonly Extension[];
  readonly emit: (context: SlideEmitContext) => SlideEmitOutput;
  readonly paint?: (context: SlidePaintContext) => SlidePaintOutput;
};

export type InternalResolvedSlideLayout = {
  readonly key: string;
  readonly masterKey: string;
  readonly id: string;
  readonly name?: string;
  readonly type?: SlideLayoutType;
  readonly preserve?: boolean;
  readonly matchingName?: string;
  readonly userDrawn?: boolean;
  readonly showMasterShapes?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly placeholders?: PlaceholderMap;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
  readonly headerFooter?: HeaderFooter;
  readonly background?: ResolvedSlideCapabilityRecord;
  readonly footerPlaceholders?: ResolvedSlideCapabilityRecord;
  readonly placeholderVisuals?: ResolvedSlideCapabilityRecord;
  readonly layers?: readonly ResolvedObjectLayer[];
  readonly slideCapabilities?: readonly ResolvedSlideCapabilityRecord[];
};

export type InternalResolvedSlideMaster = {
  readonly key: string;
  readonly id: string;
  readonly name?: string;
  readonly preserve?: boolean;
  readonly theme: Theme;
  readonly themeContext: ThemeContext;
  readonly colorMap: ColorMap;
  readonly headerFooter?: HeaderFooter;
  readonly background?: ResolvedSlideCapabilityRecord;
  readonly footerPlaceholders?: ResolvedSlideCapabilityRecord;
  readonly placeholderVisuals?: ResolvedSlideCapabilityRecord;
  readonly layers?: readonly ResolvedObjectLayer[];
  readonly placeholders?: MasterPlaceholderMap;
  readonly textStyles?: MasterTextStyles;
  readonly layouts: readonly InternalResolvedSlideLayout[];
  readonly slideCapabilities?: readonly ResolvedSlideCapabilityRecord[];
};

export type InternalResolvedSlide = {
  readonly id: string;
  readonly name?: string;
  readonly authorId?: string;
  readonly index: number;
  readonly theme: ThemeContext;
  readonly layoutKey?: string;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
  readonly effectiveThemeOverride?: ThemeOverride;
  readonly effectiveColorMapOverride?: ColorMapOverride;
  readonly hasCustomColorMap?: boolean;
  readonly section?: string;
  readonly hidden: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly showMasterShapes?: boolean;
  readonly inheritedBackground?: ResolvedSlideCapabilityRecord;
  readonly masterLayers?: readonly ResolvedObjectLayer[];
  readonly layoutLayers?: readonly ResolvedObjectLayer[];
  readonly content: readonly QueryContentNode[];
  readonly layers: readonly ResolvedObjectLayer[];
  readonly slideCapabilities: readonly ResolvedSlideCapabilityRecord[];
};

export type InternalResolution = {
  readonly dimensions: ResolutionDimensions;
  readonly title?: string;
  readonly theme: ThemeContext;
  readonly masters?: readonly InternalResolvedSlideMaster[];
  readonly mastersCanvas?: NonNullable<SlideCapabilityModule["mastersCanvas"]>;
  readonly mastersEmitDriver?: MastersEmitDriver;
  readonly sections: readonly ResolvedSectionRecord[];
  readonly sectionsEmitter?: {
    readonly extensions: readonly Extension[];
    emit(): SlideEmitOutput;
  };
  readonly slides: readonly InternalResolvedSlide[];
  readonly resourceClaims: ReadonlyMap<
    string,
    {
      readonly identity: string;
      readonly path: string;
      readonly characters: string;
    }
  >;
  readonly warnings: readonly Diagnostic[];
};

// Keep resolved pipeline internals out of the public Resolution object and unreachable from author code.
const resolutionStore = new WeakMap<Resolution, InternalResolution>();

export function registerInternalResolution(
  resolution: Resolution,
  internal: InternalResolution,
): void {
  resolutionStore.set(resolution, internal);
}

export function getInternalResolution(
  resolution: Resolution,
): InternalResolution {
  const internal = resolutionStore.get(resolution);
  if (internal === undefined) {
    throw new TypeError("Expected a Resolution produced by resolve()");
  }
  return internal;
}
