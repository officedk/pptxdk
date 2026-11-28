import {
  matchLayoutPlaceholderKey,
  matchMasterPlaceholderSpec,
  nativePlaceholderRefSchema,
  placeholderRoleSchema,
  resolvePlaceholderBoxFromLayers,
  type Geometry,
  type LayoutPlaceholderSpec,
  type MasterPlaceholderSpec,
  type NativePlaceholderRef,
} from "@pptxdk/core/capability";
import { pxToEmu } from "@pptxdk/emit";
import {
  ooxmlBoolSchema,
  ooxmlNonNegativeIntSchema,
  P_NS,
  type ImportDiagnostic,
  type ImportLocation,
  type ShapeTreeImportContext,
  type XmlElementCursor,
} from "@pptxdk/import";

import {
  resolvePlaceholderVisualStyle,
  type PlaceholderInheritanceLayers,
} from "./style-inheritance.ts";

const P14_MOD_ID_EXT_URI = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";

export type XmlElement = XmlElementCursor["element"];

/** Finds a direct child XML element matching namespace URI and local name without mutating cursor tracking. */
export function directElementChild(
  parent: XmlElement,
  namespaceUri: string,
  localName: string,
): XmlElement | undefined {
  for (
    let child = parent.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if (
      (el.namespaceURI ?? "") === namespaceUri &&
      (el.localName ?? el.tagName) === localName
    ) {
      return el;
    }
  }
  return undefined;
}

/** Side-effect-free lookup of `<p:ph>` inside `<p:{nvContainer}>/<p:nvPr>`. */
export function directPlaceholderElement(
  cursor: XmlElementCursor,
  nvContainerLocalName: "nvSpPr" | "nvPicPr" | "nvGraphicFramePr",
): XmlElement | undefined {
  const nvContainer = directElementChild(
    cursor.element,
    P_NS,
    nvContainerLocalName,
  );
  if (nvContainer === undefined) return undefined;
  const nvPr = directElementChild(nvContainer, P_NS, "nvPr");
  if (nvPr === undefined) return undefined;
  return directElementChild(nvPr, P_NS, "ph");
}

/** Side-effect-free predicate identifying `<p:sp>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderSpElement(cursor: XmlElementCursor): boolean {
  return directPlaceholderElement(cursor, "nvSpPr") !== undefined;
}

/** Side-effect-free predicate identifying `<p:pic>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderPicElement(cursor: XmlElementCursor): boolean {
  return directPlaceholderElement(cursor, "nvPicPr") !== undefined;
}

/** Side-effect-free predicate identifying `<p:graphicFrame>` elements bound to a placeholder (`<p:ph>`). */
export function isPlaceholderGraphicFrameElement(
  cursor: XmlElementCursor,
): boolean {
  return directPlaceholderElement(cursor, "nvGraphicFramePr") !== undefined;
}

function consumeSubtree(cursor: XmlElementCursor): void {
  const attrs = cursor.element.attributes;
  for (let i = 0; i < attrs.length; i += 1) {
    const attr = attrs.item(i);
    if (attr === null) continue;
    const ns = attr.namespaceURI ?? "";
    const local = attr.localName ?? attr.name;
    if (
      ns === "http://www.w3.org/2000/xmlns/" ||
      attr.name === "xmlns" ||
      attr.prefix === "xmlns"
    ) {
      continue;
    }
    cursor.consumeAttr(local, ns);
  }
  for (const child of cursor.remainingElementChildren()) {
    consumeSubtree(child);
  }
}

export type NvPrPlaceholderReadResult =
  | { readonly kind: "none" }
  | { readonly kind: "ok"; readonly nativeRef: NativePlaceholderRef }
  | { readonly kind: "error" };

/**
 * Reads optional `<p:ph>` and benign `p14:modId` extensions inside `<p:nvPr>`.
 * Shared across text, image, table, chart, and unpopulated placeholder importers.
 */
export function readNvPrPlaceholder(
  nvPr: XmlElementCursor | undefined,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): NvPrPlaceholderReadResult {
  if (nvPr === undefined) return { kind: "none" };

  const extLst = nvPr.child(P_NS, "extLst");
  if (extLst !== undefined) {
    for (const ext of extLst.children(P_NS, "ext")) {
      if (ext.attr("uri") === P14_MOD_ID_EXT_URI) {
        consumeSubtree(ext);
      }
    }
  }

  const ph = nvPr.child(P_NS, "ph");
  if (ph === undefined) {
    return { kind: "none" };
  }

  const rawPh = ph.attrsObject([
    "type",
    "idx",
    "sz",
    "orient",
    "hasCustomPrompt",
  ]);
  const explicitType = rawPh["type"] !== undefined;
  const parsedRole = placeholderRoleSchema.safeParse(rawPh["type"] ?? "obj");
  if (!parsedRole.success) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: ph.xpath },
      namespaceUri: P_NS,
      localName: "ph",
      mechanismHint: `ph:${rawPh["type"] ?? ""}`,
    });
    return { kind: "error" };
  }

  const idxParsed = ooxmlNonNegativeIntSchema.safeParse(rawPh["idx"]);
  const hasPromptParsed = ooxmlBoolSchema.safeParse(rawPh["hasCustomPrompt"]);
  const parsedNativeRef = nativePlaceholderRefSchema.safeParse({
    type: parsedRole.data,
    explicitType,
    ...(idxParsed.success ? { idx: idxParsed.data } : {}),
    ...(rawPh["sz"] !== undefined ? { size: rawPh["sz"] } : {}),
    ...(rawPh["orient"] !== undefined ? { orient: rawPh["orient"] } : {}),
    ...(hasPromptParsed.success && hasPromptParsed.data
      ? { hasCustomPrompt: true }
      : {}),
  });
  if (!parsedNativeRef.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: ph.xpath },
      attributeOrElement: "p:ph",
      rawValue: JSON.stringify(rawPh),
    });
    return { kind: "error" };
  }

  return { kind: "ok", nativeRef: parsedNativeRef.data };
}

export type PlaceholderClaimResult =
  | {
      readonly bound: true;
      readonly key?: string;
      readonly layoutPlaceholder?: LayoutPlaceholderSpec;
      readonly masterPlaceholder?: MasterPlaceholderSpec;
    }
  | {
      readonly bound: false;
      readonly reason: "orphan" | "duplicate" | "incompatible";
      readonly layoutPlaceholder?: LayoutPlaceholderSpec;
      readonly masterPlaceholder?: MasterPlaceholderSpec;
      readonly inheritanceContext: PlaceholderInheritanceLayers;
    };

/**
 * Evaluates whether a slide object carrying `<p:ph>` can keep a 1:1 `placeholder` slot binding,
 * or must materialize its 5-layer master/layout cascade and unbind into a free slide object.
 */
export function resolvePlaceholderImportClaim(args: {
  readonly context: ShapeTreeImportContext;
  readonly nativeRef: NativePlaceholderRef;
  readonly compatibleRoles?: ReadonlySet<string>;
  readonly allowMasterDirectFallback?: boolean;
}): PlaceholderClaimResult {
  const {
    context,
    nativeRef,
    compatibleRoles,
    allowMasterDirectFallback = false,
  } = args;
  const matchedKey = matchLayoutPlaceholderKey(context.layoutPlaceholders, {
    type: nativeRef.type,
    idx: nativeRef.idx,
  });
  const layoutPlaceholder =
    matchedKey === undefined
      ? undefined
      : context.layoutPlaceholders?.[matchedKey];
  const effectiveRole = layoutPlaceholder?.type ?? nativeRef.type;
  const masterPlaceholder = matchMasterPlaceholderSpec(
    context.masterPlaceholders,
    effectiveRole,
  );
  const roleCompatible =
    compatibleRoles === undefined || compatibleRoles.has(effectiveRole);
  const alreadyClaimed =
    matchedKey !== undefined &&
    context.claimedPlaceholderKeys?.has(matchedKey) === true;

  if (
    matchedKey !== undefined &&
    layoutPlaceholder !== undefined &&
    roleCompatible &&
    !alreadyClaimed
  ) {
    context.claimedPlaceholderKeys?.add(matchedKey);
    return {
      bound: true,
      key: matchedKey,
      layoutPlaceholder,
      ...(masterPlaceholder === undefined ? {} : { masterPlaceholder }),
    };
  }

  // OOXML sentinel idx=4294967295 (0xFFFFFFFF) inherits directly from the master placeholder
  // via nativePlaceholder when the bound layout does not define a matching slot.
  if (
    allowMasterDirectFallback &&
    matchedKey === undefined &&
    nativeRef.idx === 4_294_967_295 &&
    roleCompatible &&
    masterPlaceholder !== undefined
  ) {
    return {
      bound: true,
      masterPlaceholder,
    };
  }

  const inheritanceContext: PlaceholderInheritanceLayers = {
    ...(layoutPlaceholder === undefined ? {} : { layoutPlaceholder }),
    ...(masterPlaceholder === undefined ? {} : { masterPlaceholder }),
    ...(context.masterTextStyles === undefined
      ? {}
      : { masterTextStyles: context.masterTextStyles }),
    nativeRef,
    slideDimensions: context.slideDimensions ?? { width: 960, height: 540 },
  };
  return {
    bound: false,
    reason:
      matchedKey === undefined
        ? "orphan"
        : alreadyClaimed
          ? "duplicate"
          : "incompatible",
    ...(layoutPlaceholder === undefined ? {} : { layoutPlaceholder }),
    ...(masterPlaceholder === undefined ? {} : { masterPlaceholder }),
    inheritanceContext,
  };
}

/**
 * Resolves effective free-object box coordinates for an unbound placeholder across
 * master placeholder, layout placeholder, and local `<a:xfrm>`.
 */
export function resolveUnboundPlaceholderBox(args: {
  readonly inheritanceContext: PlaceholderInheritanceLayers;
  readonly localBox?: {
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
  };
  readonly requireHeight: boolean;
}):
  | {
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height?: number;
    }
  | undefined {
  const { inheritanceContext, localBox, requireHeight } = args;
  const localStyle: Record<string, unknown> = {};
  if (localBox?.x !== undefined) localStyle["left"] = localBox.x;
  if (localBox?.y !== undefined) localStyle["top"] = localBox.y;
  if (localBox?.width !== undefined) localStyle["width"] = localBox.width;
  if (localBox?.height !== undefined) localStyle["height"] = localBox.height;

  const resolved = resolvePlaceholderBoxFromLayers(
    [
      inheritanceContext.masterPlaceholder,
      inheritanceContext.layoutPlaceholder,
      Object.keys(localStyle).length === 0 ? undefined : { style: localStyle },
    ],
    inheritanceContext.slideDimensions,
  );
  const validCoord = (v: unknown): v is number =>
    typeof v === "number" && Number.isFinite(v) && Math.abs(v) <= 1_000_000;
  const validExtent = (v: unknown): v is number => {
    if (typeof v !== "number" || !Number.isFinite(v) || v <= 0) return false;
    const emu = pxToEmu(v);
    return emu >= 1 && emu <= 2_147_483_647;
  };
  if (
    !validCoord(resolved.x) ||
    !validCoord(resolved.y) ||
    !validExtent(resolved.width)
  ) {
    return undefined;
  }
  const hasHeight = validExtent(resolved.height);
  if (requireHeight && !hasHeight) {
    return undefined;
  }
  return {
    x: resolved.x,
    y: resolved.y,
    width: resolved.width,
    ...(hasHeight ? { height: resolved.height as number } : {}),
  };
}

/** Removes `"none"` placeholder-clear sentinels from a style record materialized onto a free object. */
export function stripPlaceholderClearSentinels(
  style: Record<string, unknown>,
): Record<string, unknown> {
  const cleaned: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(style)) {
    if (value === undefined) continue;
    if (
      (key === "background" ||
        key === "border" ||
        key === "boxShadow" ||
        key === "textShadow") &&
      value === "none"
    ) {
      continue;
    }
    cleaned[key] = value;
  }
  return cleaned;
}

/** Computes corner borderRadius when a geometry specifies DrawingML `roundRect`. */
export function roundRectBorderRadiusForBox(
  geometry: Geometry | undefined,
  boxSize: { readonly width: number; readonly height: number } | undefined,
): number | undefined {
  if (geometry === undefined || boxSize === undefined) return undefined;
  if (geometry === "roundRect") {
    return (16667 / 100_000) * Math.min(boxSize.width, boxSize.height);
  }
  if (
    typeof geometry === "object" &&
    geometry.type === "preset" &&
    geometry.name === "roundRect"
  ) {
    const adj = geometry.adjust?.["adj"] ?? 16667;
    if (adj > 0) {
      return (adj / 100_000) * Math.min(boxSize.width, boxSize.height);
    }
  }
  return undefined;
}

/** Materializes 5-layer visual (`spPr`) overrides onto an unbound free slide object. */
export function materializeUnboundVisualStyle(args: {
  readonly inheritanceContext: PlaceholderInheritanceLayers;
  readonly localVisualStyle: Readonly<Record<string, unknown>>;
  readonly localGeometry?: Geometry;
  readonly boxSize: { readonly width: number; readonly height: number };
}): Record<string, unknown> {
  const { inheritanceContext, localVisualStyle, localGeometry, boxSize } = args;
  const resolved = stripPlaceholderClearSentinels({
    ...resolvePlaceholderVisualStyle(inheritanceContext, localVisualStyle),
  });
  if (resolved["borderRadius"] === undefined) {
    const effectiveGeom =
      localGeometry ??
      inheritanceContext.layoutPlaceholder?.geometry ??
      inheritanceContext.masterPlaceholder?.geometry;
    const radius = roundRectBorderRadiusForBox(effectiveGeom, boxSize);
    if (radius !== undefined) {
      resolved["borderRadius"] = radius;
    }
  }
  return resolved;
}

/** Records a `compatibility-fallback` diagnostic when an over-rule placeholder unbinds into a free object. */
export function emitPlaceholderUnbindDiagnostic(args: {
  readonly diagnostics: ImportDiagnostic[];
  readonly location: ImportLocation;
  readonly retainedAs: "text" | "shape" | "image" | "table" | "chart";
}): void {
  args.diagnostics.push({
    code: "compatibility-fallback",
    category: "unsupported",
    disposition: "fallback",
    location: args.location,
    preferredMechanism: "placeholder",
    selectedFallback: args.retainedAs,
    loss: {
      lost: "master-layout-inheritance",
      retainedAs: args.retainedAs,
    },
  });
}
