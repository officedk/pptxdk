import type {
  CascadedOmission,
  ImportDiagnostic,
  ImportLocation,
} from "#/diagnostics.ts";
import {
  A_NS,
  DIAGRAM_NS,
  P_NS,
  R_NS,
  REL_DIAGRAM_COLORS,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  REL_DIAGRAM_LAYOUT,
  REL_DIAGRAM_QUICK_STYLE,
} from "#/namespaces.ts";
import {
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

export type RelationshipIndex = {
  readonly byId: ReadonlyMap<string, ParsedRelationshipEntry>;
  readonly entries: readonly ParsedRelationshipEntry[];
  readonly diagnostics: readonly ImportDiagnostic[];
};

/** Reads and indexes all OPC relationships owned by `sourcePartPath`. */
export function buildRelationshipIndex(
  pkg: PptxPackage,
  sourcePartPath: string,
): RelationshipIndex {
  const { entries, diagnostics } = pkg.readRelationships(sourcePartPath);
  const byId = new Map<string, ParsedRelationshipEntry>();
  for (const entry of entries) {
    byId.set(entry.Id, entry);
  }
  return { byId, entries, diagnostics };
}

export type ResolvedInternalResource =
  | {
      readonly ok: true;
      readonly targetPart: string;
      readonly relationship: ParsedRelationshipEntry;
      readonly bytes: Uint8Array;
    }
  | {
      readonly ok: false;
      readonly diagnostic: Extract<
        ImportDiagnostic,
        { readonly code: "broken-relationship" | "missing-target-part" }
      >;
    };

/**
 * Resolves an internal OPC relationship ID to detached target part bytes without touching external URLs.
 */
export function resolveInternalRelationshipTarget(
  pkg: PptxPackage,
  relationships: ReadonlyMap<string, ParsedRelationshipEntry>,
  sourcePartPath: string,
  relationshipId: string,
  location: ImportLocation,
  expectedType?: string,
): ResolvedInternalResource {
  const rel = relationships.get(relationshipId);
  if (
    rel === undefined ||
    rel.TargetMode === "External" ||
    (expectedType !== undefined && rel.Type !== expectedType)
  ) {
    return {
      ok: false,
      diagnostic: {
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location,
        relationshipId,
      },
    };
  }
  const targetPart = resolvePartPath(sourcePartPath, rel.Target);
  if (targetPart === undefined) {
    return {
      ok: false,
      diagnostic: {
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location,
        relationshipId,
      },
    };
  }
  const bytes = pkg.readDetachedBytes(targetPart);
  if (bytes === undefined) {
    return {
      ok: false,
      diagnostic: {
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location,
        relationshipId,
        targetPart,
      },
    };
  }
  return { ok: true, targetPart, relationship: rel, bytes };
}

export type ResolvedExternalResource =
  | {
      readonly ok: true;
      readonly target: string;
      readonly relationship: ParsedRelationshipEntry;
    }
  | {
      readonly ok: false;
      readonly diagnostic: Extract<
        ImportDiagnostic,
        { readonly code: "broken-relationship" }
      >;
    };

/** Resolves an external OPC relationship (`TargetMode="External"`) without performing any network access. */
export function resolveExternalRelationshipTarget(
  relationships: ReadonlyMap<string, ParsedRelationshipEntry>,
  relationshipId: string,
  location: ImportLocation,
  expectedType?: string,
): ResolvedExternalResource {
  const rel = relationships.get(relationshipId);
  if (
    rel === undefined ||
    rel.TargetMode !== "External" ||
    (expectedType !== undefined && rel.Type !== expectedType)
  ) {
    return {
      ok: false,
      diagnostic: {
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location,
        relationshipId,
      },
    };
  }
  return { ok: true, target: rel.Target, relationship: rel };
}

const diagramRelRoleByAttr = [
  ["dm", "diagram-data", REL_DIAGRAM_DATA],
  ["lo", "diagram-layout", REL_DIAGRAM_LAYOUT],
  ["qs", "diagram-quick-style", REL_DIAGRAM_QUICK_STYLE],
  ["cs", "diagram-colors", REL_DIAGRAM_COLORS],
] as const;

type XmlElement = XmlElementCursor["element"];

function firstDirectChild(
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

/** Extracts non-visual `cNvPr` id and name from any shape-tree element without mutating cursor tracking. */
export function peekShapeTreeObjectIdentity(element: XmlElement): {
  readonly objectId?: number;
  readonly objectName?: string;
} {
  for (
    let child = element.firstChild;
    child !== null;
    child = child.nextSibling
  ) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    if ((el.namespaceURI ?? "") !== P_NS) continue;
    const local = el.localName ?? el.tagName;
    if (
      local !== "nvSpPr" &&
      local !== "nvPicPr" &&
      local !== "nvGraphicFramePr" &&
      local !== "nvGrpSpPr" &&
      local !== "nvCxnSpPr"
    ) {
      continue;
    }
    const cNvPr = firstDirectChild(el, P_NS, "cNvPr");
    if (cNvPr === undefined) return {};
    const rawId = cNvPr.getAttribute("id");
    const parsedId =
      rawId !== null && /^\d+$/u.test(rawId) ? Number(rawId) : undefined;
    const rawName = cNvPr.getAttribute("name");
    return {
      ...(parsedId !== undefined && Number.isSafeInteger(parsedId)
        ? { objectId: parsedId }
        : {}),
      ...(rawName !== null && rawName.trim() !== ""
        ? { objectName: rawName }
        : {}),
    };
  }
  return {};
}

/** Inspects an omitted `<p:graphicFrame>` for `graphicData/@uri` and cascaded diagram relationships. */
export function inspectUnsupportedGraphicFrame(
  cursor: XmlElementCursor,
  sourcePartPath: string,
  relationships: ReadonlyMap<string, ParsedRelationshipEntry>,
  relationshipEntries: readonly ParsedRelationshipEntry[],
): {
  readonly mechanismHint?: string;
  readonly cascaded: readonly CascadedOmission[];
} {
  const graphic = firstDirectChild(cursor.element, A_NS, "graphic");
  const graphicData =
    graphic === undefined
      ? undefined
      : firstDirectChild(graphic, A_NS, "graphicData");
  const uri = graphicData?.getAttribute("uri") ?? undefined;
  const cascaded: CascadedOmission[] = [];
  const seenParts = new Set<string>();

  if (graphicData !== undefined && uri === DIAGRAM_NS) {
    const relIds = firstDirectChild(graphicData, DIAGRAM_NS, "relIds");
    if (relIds !== undefined) {
      for (const [attrName, role] of diagramRelRoleByAttr) {
        const rId = relIds.getAttributeNS(R_NS, attrName);
        if (!rId) continue;
        const rel = relationships.get(rId);
        if (rel === undefined || rel.TargetMode === "External") continue;
        const targetPart = resolvePartPath(sourcePartPath, rel.Target);
        if (targetPart === undefined) continue;
        seenParts.add(targetPart);
        cascaded.push({ location: { part: targetPart }, role });
      }
    }
    for (const rel of relationshipEntries) {
      if (rel.Type === REL_DIAGRAM_DRAWING && rel.TargetMode !== "External") {
        const targetPart = resolvePartPath(sourcePartPath, rel.Target);
        if (targetPart !== undefined && !seenParts.has(targetPart)) {
          seenParts.add(targetPart);
          cascaded.push({
            location: { part: targetPart },
            role: "diagram-drawing",
          });
        }
      }
    }
  }

  return {
    ...(uri !== undefined && uri !== "" ? { mechanismHint: uri } : {}),
    cascaded,
  };
}
