import { z } from "zod";

import type { ImportDiagnostic } from "#/diagnostics.ts";
import {
  appendUnconsumedDiagnostics,
  consumeIdentityShapeTreeBoilerplate,
} from "#/masters.ts";
import {
  A_NS,
  P14_CREATION_ID_EXT_URI,
  P_NS,
  REL_HANDOUT_MASTER,
  REL_NOTES_MASTER,
  R_NS,
} from "#/namespaces.ts";
import {
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import { peekShapeTreeObjectIdentity } from "#/resources.ts";
import { OoxmlNonNegativeIntSchema } from "#/xml-fragments.ts";
import { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";

const standardNotesMasterPlaceholders = new Set([
  "sldImg",
  "body",
  "hdr",
  "ftr",
  "dt",
  "sldNum",
]);

const standardHandoutMasterPlaceholders = new Set([
  "hdr",
  "ftr",
  "dt",
  "sldNum",
]);

const defaultColorMap: Readonly<Record<string, string>> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const PlaceholderTypeXmlSchema = z.object({
  type: z.string().optional(),
  idx: OoxmlNonNegativeIntSchema.optional(),
  sz: z.enum(["full", "half", "quarter"]).optional(),
  orient: z.enum(["horz", "vert"]).optional(),
});

const BgRefAttrsSchema = z.object({
  idx: OoxmlNonNegativeIntSchema,
});

function isStandardColorMap(clrMap: XmlElementCursor): boolean {
  const attrs = clrMap.element.attributes;
  const observed = new Map<string, string>();
  for (let i = 0; i < attrs.length; i += 1) {
    const attr = attrs.item(i);
    if (attr === null) continue;
    const ns = attr.namespaceURI ?? "";
    const local = attr.localName ?? attr.name;
    if (ns === "" && local in defaultColorMap) {
      observed.set(local, attr.value);
    }
  }
  const matches = Object.entries(defaultColorMap).every(
    ([slot, target]) => observed.get(slot) === target,
  );
  if (matches) {
    for (const slot of Object.keys(defaultColorMap)) {
      clrMap.consumeAttr(slot);
    }
  }
  return matches;
}

function isBaselineBgRef(bg: XmlElementCursor): boolean {
  const bgRef = bg.child(P_NS, "bgRef");
  if (bgRef === undefined) return false;
  const raw = bgRef.attrsObject(["idx"]);
  const parsed = BgRefAttrsSchema.safeParse(raw);
  if (!parsed.success || parsed.data.idx !== 1001) {
    return false;
  }
  const schemeClr = bgRef.child(A_NS, "schemeClr");
  if (schemeClr === undefined) return false;
  return schemeClr.attr("val") === "bg1";
}

function consumeBenignCreationIdExtLst(cSld: XmlElementCursor): void {
  const extLst = cSld.child(P_NS, "extLst");
  if (extLst === undefined) return;
  for (const ext of extLst.children(P_NS, "ext")) {
    if (ext.element.getAttribute("uri") !== P14_CREATION_ID_EXT_URI) {
      continue;
    }
    ext.consumeAttr("uri");
    ext.suppressUnconsumed();
  }
}

function readPlaceholderTypeOnSp(
  spCursor: XmlElementCursor,
): string | undefined {
  const nvSpPr = spCursor.child(P_NS, "nvSpPr");
  if (nvSpPr === undefined) return undefined;
  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    cNvPr.consumeAttr("id");
    cNvPr.consumeAttr("name");
  }
  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    cNvSpPr.suppressUnconsumed();
  }
  const nvPr = nvSpPr.child(P_NS, "nvPr");
  if (nvPr === undefined) return undefined;
  const ph = nvPr.child(P_NS, "ph");
  if (ph === undefined) return undefined;
  const parsed = PlaceholderTypeXmlSchema.safeParse(
    ph.attrsObject(["type", "idx", "sz", "orient"]),
  );
  return parsed.success ? parsed.data.type : undefined;
}

function inspectPrintMasterPart(
  pkg: PptxPackage,
  masterPartPath: string,
  kind: "notesMaster" | "handoutMaster",
  diagnostics: ImportDiagnostic[],
): void {
  const xmlText = pkg.readTextPart(masterPartPath);
  if (xmlText === undefined) return;

  let rootEl: ReturnType<typeof parseXmlPart>;
  try {
    rootEl = parseXmlPart(masterPartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return;
  }

  if (
    (rootEl.namespaceURI ?? "") !== P_NS ||
    (rootEl.localName ?? rootEl.tagName) !== kind
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return;
  }

  const retainedAs =
    kind === "notesMaster" ? "default-notes-master" : "omitted";
  const allowedPlaceholders =
    kind === "notesMaster"
      ? standardNotesMasterPlaceholders
      : standardHandoutMasterPlaceholders;

  const masterCursor = XmlElementCursor.createRoot(
    rootEl,
    masterPartPath,
    `/p:${kind}`,
  );
  const rels = pkg.readRelationships(masterPartPath);
  diagnostics.push(...rels.diagnostics);

  let reportedCustomLoss = false;

  const cSld = masterCursor.child(P_NS, "cSld");
  if (cSld !== undefined) {
    cSld.consumeAttr("name");
    const bg = cSld.child(P_NS, "bg");
    if (bg !== undefined) {
      if (!isBaselineBgRef(bg)) {
        reportedCustomLoss = true;
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: bg.xpath },
          namespaceUri: bg.namespaceUri,
          localName: bg.localName,
          mechanismHint: `${kind}-custom-background`,
          loss: {
            lost: "custom-print-layout",
            retainedAs,
          },
        });
        bg.suppressUnconsumed();
      }
    }

    const spTree = cSld.child(P_NS, "spTree");
    if (spTree !== undefined) {
      consumeIdentityShapeTreeBoilerplate(spTree);
      for (const item of spTree.remainingElementChildren()) {
        if (item.namespaceUri === P_NS && item.localName === "sp") {
          const phType = readPlaceholderTypeOnSp(item);
          if (phType !== undefined && allowedPlaceholders.has(phType)) {
            // Standard baseline placeholder template (geometry/prompt/lstStyle): consume cleanly.
            item.suppressUnconsumed();
            continue;
          }
        }
        reportedCustomLoss = true;
        const identity = peekShapeTreeObjectIdentity(item.element);
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: masterPartPath,
            xpath: item.xpath,
            ...identity,
          },
          namespaceUri: item.namespaceUri,
          localName: item.localName,
          mechanismHint: `${kind}-custom-object`,
          loss: {
            lost: "custom-print-layout",
            retainedAs,
          },
        });
        item.suppressUnconsumed();
      }
    }

    consumeBenignCreationIdExtLst(cSld);
  }

  const clrMap = masterCursor.child(P_NS, "clrMap");
  if (clrMap !== undefined && !isStandardColorMap(clrMap)) {
    reportedCustomLoss = true;
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { part: masterPartPath, xpath: clrMap.xpath },
      namespaceUri: clrMap.namespaceUri,
      localName: clrMap.localName,
      mechanismHint: `${kind}-custom-color-map`,
      loss: {
        lost: "custom-print-layout",
        retainedAs,
      },
    });
    clrMap.suppressUnconsumed();
  }

  const styleElementName =
    kind === "notesMaster" ? "notesStyle" : "handoutStyle";
  const masterStyle = masterCursor.child(P_NS, styleElementName);
  if (masterStyle !== undefined) {
    masterStyle.suppressUnconsumed();
  }

  if (kind === "handoutMaster" && !reportedCustomLoss) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { part: masterPartPath, xpath: masterCursor.xpath },
      namespaceUri: P_NS,
      localName: kind,
      mechanismHint: "handout-master-print-layout",
      loss: {
        lost: "custom-print-layout",
        retainedAs: "omitted",
      },
    });
  }

  appendUnconsumedDiagnostics(masterCursor, undefined, diagnostics);
}

/**
 * Inspects `<p:notesMasterIdLst>` and `<p:handoutMasterIdLst>` from `presentation.xml`.
 *
 * Standard baseline `notesMaster` parts produce zero false-loss diagnostics, while custom
 * non-placeholder print graphics/backgrounds and `handoutMaster` layouts report exact part
 * locations and `custom-print-layout` editability loss without blocking slide notes import.
 */
export function inspectNotesAndHandoutMasters(
  presCursor: XmlElementCursor,
  pkg: PptxPackage,
  presRelById: ReadonlyMap<string, ParsedRelationshipEntry>,
  diagnostics: ImportDiagnostic[],
): void {
  const visitedParts = new Set<string>();

  const inspectList = (
    listTag: "notesMasterIdLst" | "handoutMasterIdLst",
    entryTag: "notesMasterId" | "handoutMasterId",
    relType: string,
    kind: "notesMaster" | "handoutMaster",
  ): void => {
    const lst = presCursor.child(P_NS, listTag);
    if (lst === undefined) return;
    for (const idCursor of lst.children(P_NS, entryTag)) {
      const rId = idCursor.attr("id", R_NS);
      if (rId === undefined) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: idCursor.xpath },
          relationshipId: "",
        });
        continue;
      }
      const rel = presRelById.get(rId);
      if (
        rel === undefined ||
        rel.Type !== relType ||
        rel.TargetMode === "External"
      ) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: idCursor.xpath },
          relationshipId: rId,
        });
        continue;
      }
      const targetPart = resolvePartPath(pkg.presentationPath, rel.Target);
      if (targetPart === undefined || !pkg.hasPart(targetPart)) {
        diagnostics.push({
          code: "missing-target-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: pkg.presentationPath, xpath: idCursor.xpath },
          relationshipId: rId,
          targetPart: targetPart ?? rel.Target,
        });
        continue;
      }
      if (visitedParts.has(targetPart)) continue;
      visitedParts.add(targetPart);
      inspectPrintMasterPart(pkg, targetPart, kind, diagnostics);
    }
  };

  inspectList(
    "notesMasterIdLst",
    "notesMasterId",
    REL_NOTES_MASTER,
    "notesMaster",
  );
  inspectList(
    "handoutMasterIdLst",
    "handoutMasterId",
    REL_HANDOUT_MASTER,
    "handoutMaster",
  );
}
