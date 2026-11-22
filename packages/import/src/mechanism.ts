import type {
  CanvasMechanismDefinition,
  EmitContext,
  EmitOutput,
  NativeColorSlot,
  PaintContext,
  PaintOutput,
  PlaceholderMap,
  ShowMechanismDefinition,
  SlideEmitContext,
  SlideEmitOutput,
} from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";

import type { ImportDiagnostic, ImportLocation } from "#/diagnostics.ts";
import type { ParsedRelationshipEntry, PptxPackage } from "#/opc.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

export type DeferredSlideLink = {
  readonly targetPart: string;
  readonly relationshipId: string;
  readonly location: ImportLocation;
  apply(targetSlideId: string): void;
};

export type ShapeTreeImportContext = {
  readonly cursor: XmlElementCursor;
  readonly part: string;
  readonly slideIndex: number;
  readonly package: PptxPackage;
  readonly containerNode?: CommonNode;
  readonly relationships?: ReadonlyMap<string, ParsedRelationshipEntry>;
  readonly registerDeferredSlideLink?: (link: DeferredSlideLink) => void;
  readonly registerImportedObject?: (spid: number, node: CommonNode) => void;
  readonly layoutPlaceholders?: PlaceholderMap;
  readonly clrMap?: Readonly<Record<string, NativeColorSlot>>;
};

export type ShapeTreeImportOutcome = {
  readonly node?: CommonNode;
  readonly diagnostics?: readonly ImportDiagnostic[];
};

export type ShapeTreeImportHandler = {
  readonly target: "shape-tree";
  readonly localName: "sp" | "pic" | "grpSp" | "graphicFrame" | "cxnSp";
  /** Refines matching when multiple mechanisms share an element localName (such as `p:sp`). */
  matches?(cursor: XmlElementCursor): boolean;
  read(context: ShapeTreeImportContext): ShapeTreeImportOutcome;
};

export type OmittedSlideObjectInfo = {
  readonly objectName?: string;
  readonly reason: "unsupported" | "corrupt";
};

export type SlideImportContext = {
  readonly slideCursor: XmlElementCursor;
  readonly slideNode: CommonNode;
  readonly part: string;
  readonly slideIndex?: number;
  readonly package: PptxPackage;
  readonly objectBySpid?: ReadonlyMap<number, CommonNode>;
  readonly omittedSpids?: ReadonlyMap<number, OmittedSlideObjectInfo>;
  readonly placeholderKeyBySpid?: ReadonlyMap<number, string>;
};

export type SlideImportOutcome = {
  readonly diagnostics?: readonly ImportDiagnostic[];
};

export type SlideImportHandler = {
  readonly target: "slide";
  readonly slot:
    | "background"
    | "header-footer"
    | "transition"
    | "timing"
    | "notes";
  /** Refines matching when multiple mechanisms share a slide slot. */
  matches?(cursor: XmlElementCursor): boolean;
  read(context: SlideImportContext): SlideImportOutcome;
};

export type DeckImportContext = {
  readonly presentationCursor: XmlElementCursor;
  readonly deckNode: CommonNode;
  readonly package: PptxPackage;
};

export type DeckImportOutcome = {
  readonly diagnostics?: readonly ImportDiagnostic[];
};

export type DeckImportHandler = {
  readonly target: "deck";
  read(context: DeckImportContext): DeckImportOutcome;
};

export type TextInlineImportContext = {
  readonly cursor: XmlElementCursor;
  readonly part: string;
  readonly slideIndex: number;
};

export type TextInlineImportHandler = {
  readonly target: "text-inline";
  readonly localName: "fld";
  matches?(cursor: XmlElementCursor): boolean;
  read(context: TextInlineImportContext): {
    readonly kind: "slide-number" | "unsupported";
    readonly fieldType: string;
    readonly rPrCursor?: XmlElementCursor;
    readonly fallbackText: string;
    readonly diagnostics: readonly ImportDiagnostic[];
  };
};

/** Concrete import-direction handler bound onto production mechanism definitions. */
export type ImportMechanismHandler =
  | ShapeTreeImportHandler
  | SlideImportHandler
  | DeckImportHandler
  | TextInlineImportHandler;

/** Defines a native canvas mechanism requiring production `emit`, `paint`, and `import` directions. */
export function defineCanvasMechanism<
  const Key extends string,
  Model,
  EmitCtx = EmitContext,
  EmitOut = EmitOutput,
  PaintCtx = PaintContext,
  PaintOut = PaintOutput,
>(
  definition: CanvasMechanismDefinition<
    Key,
    Model,
    ImportMechanismHandler,
    EmitCtx,
    EmitOut,
    PaintCtx,
    PaintOut
  >,
): CanvasMechanismDefinition<
  Key,
  Model,
  ImportMechanismHandler,
  EmitCtx,
  EmitOut,
  PaintCtx,
  PaintOut
> {
  return Object.freeze(definition);
}

/** Defines a native show mechanism requiring production `emit` and `import` directions without `paint`. */
export function defineShowMechanism<
  const Key extends string,
  Model,
  EmitCtx = SlideEmitContext,
  EmitOut = SlideEmitOutput,
>(
  definition: ShowMechanismDefinition<
    Key,
    Model,
    ImportMechanismHandler,
    EmitCtx,
    EmitOut
  >,
): ShowMechanismDefinition<
  Key,
  Model,
  ImportMechanismHandler,
  EmitCtx,
  EmitOut
> {
  return Object.freeze(definition);
}
