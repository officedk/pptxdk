import type { MorphKey } from "#/morph.ts";
import type { QueryContentNode } from "#/query.ts";
import type { ThemeContext } from "#/theme.ts";
import type {
  Diagnostic,
  EmitContext,
  EmitOutput,
  Extension,
  GeometryBox,
  PaintContext,
  PaintOutput,
  PreviewTransform,
  Resolution,
  ResolvedLink,
  SlideCapabilitySlot,
  SlideEmitContext,
  SlideEmitOutput,
  SlidePaintContext,
  SlidePaintOutput,
} from "#/types.ts";

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

export type InternalResolvedSlide = {
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden: boolean;
  readonly content: readonly QueryContentNode[];
  readonly layers: readonly ResolvedObjectLayer[];
  readonly slideCapabilities: readonly ResolvedSlideCapabilityRecord[];
};

export type InternalResolution = {
  readonly dimensions: ResolutionDimensions;
  readonly title?: string;
  readonly theme: ThemeContext;
  readonly sections: readonly ResolvedSectionRecord[];
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
