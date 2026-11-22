import { textStyleKeys } from "@pptxdk/core/capability";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import {
  A_NS,
  P14_CREATION_ID_EXT_URI,
  P_NS,
  REL_NOTES_SLIDE,
  appendUnconsumedDiagnostics,
  consumeIdentityShapeTreeBoilerplate,
  importPresentationSections,
  parseXmlPart,
  peekShapeTreeObjectIdentity,
  resolvePartPath,
  type DeckImportContext,
  type DeckImportHandler,
  type DeckImportOutcome,
  type ImportDiagnostic,
  type SlideImportContext,
  type SlideImportHandler,
  type SlideImportOutcome,
  type XmlElementCursor,
  XmlElementCursor as CursorClass,
} from "@pptxdk/import";

import { populateRichTextNode, readRichTextBody } from "../text/text-import.ts";
import { createNotesNode } from "./notes-declaration.ts";
import {
  NotesPlaceholderXmlSchema,
  NotesRootAttrsXmlSchema,
  NotesShapeIdentityXmlSchema,
} from "./notes-schemas.ts";

const baselineNotesSlidePlaceholders = new Set([
  "sldImg",
  "hdr",
  "ftr",
  "dt",
  "sldNum",
]);

function consumeNotesCreationIdExtLst(cSld: XmlElementCursor): void {
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

function readNotesPlaceholderInfo(
  spCursor: XmlElementCursor,
  part: string,
  slideIndex: number,
  diagnostics: ImportDiagnostic[],
): {
  readonly phType?: string;
  readonly objectId?: number;
  readonly objectName?: string;
} {
  const nvSpPr = spCursor.child(P_NS, "nvSpPr");
  if (nvSpPr === undefined) return {};

  let objectId: number | undefined;
  let objectName: string | undefined;
  const cNvPr = nvSpPr.child(P_NS, "cNvPr");
  if (cNvPr !== undefined) {
    const rawCNvPr = cNvPr.attrsObject(["id", "name"]);
    const parsedCNvPr = NotesShapeIdentityXmlSchema.safeParse(rawCNvPr);
    if (parsedCNvPr.success) {
      objectId = parsedCNvPr.data.id;
      if (
        parsedCNvPr.data.name !== undefined &&
        parsedCNvPr.data.name.trim() !== ""
      ) {
        objectName = parsedCNvPr.data.name;
      }
    }
  }

  const cNvSpPr = nvSpPr.child(P_NS, "cNvSpPr");
  if (cNvSpPr !== undefined) {
    cNvSpPr.suppressUnconsumed();
  }

  const nvPr = nvSpPr.child(P_NS, "nvPr");
  if (nvPr === undefined) {
    return {
      ...(objectId === undefined ? {} : { objectId }),
      ...(objectName === undefined ? {} : { objectName }),
    };
  }
  const ph = nvPr.child(P_NS, "ph");
  if (ph === undefined) {
    return {
      ...(objectId === undefined ? {} : { objectId }),
      ...(objectName === undefined ? {} : { objectName }),
    };
  }

  const rawPh = ph.attrsObject([
    "type",
    "idx",
    "sz",
    "orient",
    "hasCustomPrompt",
  ]);
  const parsedPh = NotesPlaceholderXmlSchema.safeParse(rawPh);
  if (!parsedPh.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part, xpath: ph.xpath, slideIndex },
      attributeOrElement: "p:ph",
      rawValue: JSON.stringify(rawPh),
    });
    return {
      ...(rawPh["type"] === undefined ? {} : { phType: rawPh["type"] }),
      ...(objectId === undefined ? {} : { objectId }),
      ...(objectName === undefined ? {} : { objectName }),
    };
  }

  return {
    ...(parsedPh.data.type === undefined ? {} : { phType: parsedPh.data.type }),
    ...(objectId === undefined ? {} : { objectId }),
    ...(objectName === undefined ? {} : { objectName }),
  };
}

function consumeNotesBodyShapeProperties(spPr: XmlElementCursor): void {
  // Notes body placeholders inherit geometry from the notes master; standard empty
  // <p:spPr/> or baseline rect placeholder geometry does not map to author props.
  spPr.suppressUnconsumed();
}

/**
 * Slide-level import handler for `speaker-notes`, reading `ppt/notesSlides/notesSlide{n}.xml`
 * body rich text into a page-affiliated `NotesNode` on the owning `SlideNode`.
 */
export const speakerNotesImportHandler: SlideImportHandler = {
  target: "slide",
  slot: "notes",
  read(context: SlideImportContext): SlideImportOutcome {
    const { slideNode, part, slideIndex, package: pkg } = context;
    if (slideIndex === undefined) {
      return {};
    }
    const diagnostics: ImportDiagnostic[] = [];

    const slideRels = pkg.readRelationships(part);
    const notesRel = slideRels.entries.find(
      (entry) => entry.Type === REL_NOTES_SLIDE,
    );
    if (notesRel === undefined) {
      return {};
    }

    const notesPartPath =
      notesRel.TargetMode === "External"
        ? undefined
        : resolvePartPath(part, notesRel.Target);
    if (notesPartPath === undefined) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: notesRel.xpath, slideIndex },
        relationshipId: notesRel.Id,
      });
      return { diagnostics };
    }
    if (!pkg.hasPart(notesPartPath)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part, xpath: notesRel.xpath, slideIndex },
        relationshipId: notesRel.Id,
        targetPart: notesPartPath,
      });
      return { diagnostics };
    }

    const xmlText = pkg.readTextPart(notesPartPath);
    if (xmlText === undefined) {
      return { diagnostics };
    }

    let notesEl: ReturnType<typeof parseXmlPart>;
    try {
      notesEl = parseXmlPart(notesPartPath, xmlText);
    } catch {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: notesPartPath, slideIndex },
      });
      return { diagnostics };
    }

    if (
      (notesEl.namespaceURI ?? "") !== P_NS ||
      (notesEl.localName ?? notesEl.tagName) !== "notes"
    ) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: notesPartPath, slideIndex },
      });
      return { diagnostics };
    }

    const notesRels = pkg.readRelationships(notesPartPath);
    diagnostics.push(...notesRels.diagnostics);

    const notesCursor = CursorClass.createRoot(
      notesEl,
      notesPartPath,
      "/p:notes",
    );
    const rawRootAttrs = notesCursor.attrsObject([
      "showMasterSp",
      "showMasterPhAnim",
    ]);
    if (!NotesRootAttrsXmlSchema.safeParse(rawRootAttrs).success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: { part: notesPartPath, xpath: notesCursor.xpath, slideIndex },
        attributeOrElement: "p:notes",
        rawValue: JSON.stringify(rawRootAttrs),
      });
    }

    const clrMapOvr = notesCursor.child(P_NS, "clrMapOvr");
    if (clrMapOvr !== undefined) {
      clrMapOvr.ignoreEmptyChild(A_NS, "masterClrMapping");
    }

    const cSld = notesCursor.child(P_NS, "cSld");
    if (cSld === undefined) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: notesPartPath, xpath: notesCursor.xpath, slideIndex },
      });
      return { diagnostics };
    }
    cSld.consumeAttr("name");
    consumeNotesCreationIdExtLst(cSld);

    const spTree = cSld.child(P_NS, "spTree");
    if (spTree !== undefined) {
      consumeIdentityShapeTreeBoilerplate(spTree);
      let importedNotesBody = false;

      for (const item of spTree.remainingElementChildren()) {
        if (item.namespaceUri === P_NS && item.localName === "sp") {
          const info = readNotesPlaceholderInfo(
            item,
            notesPartPath,
            slideIndex,
            diagnostics,
          );
          if (info.phType === "body" && !importedNotesBody) {
            importedNotesBody = true;
            const spPr = item.child(P_NS, "spPr");
            if (spPr !== undefined) {
              consumeNotesBodyShapeProperties(spPr);
            }
            const txBody = item.child(P_NS, "txBody");
            if (txBody !== undefined) {
              const { boxStyle, importedParagraphs } = readRichTextBody(
                txBody,
                notesPartPath,
                slideIndex,
                diagnostics,
              );
              const notesStyle: Record<string, unknown> = {};
              const boxRecord = boxStyle as Record<string, unknown>;
              for (const key of textStyleKeys) {
                if (
                  Object.hasOwn(boxRecord, key) &&
                  boxRecord[key] !== undefined
                ) {
                  notesStyle[key] = boxRecord[key];
                }
              }
              const hasNotesStyle = Object.keys(notesStyle).length > 0;
              const isEmptyPlaceholder =
                !hasNotesStyle &&
                (importedParagraphs.length === 0 ||
                  (importedParagraphs.length === 1 &&
                    importedParagraphs[0]!.runs.length === 0 &&
                    Object.keys(importedParagraphs[0]!.style).length === 0));
              if (!isEmptyPlaceholder) {
                const notesNode = createNotesNode(
                  hasNotesStyle ? { style: notesStyle } : {},
                );
                populateRichTextNode(notesNode, importedParagraphs);
                if (slideNode instanceof CommonContainer) {
                  slideNode.append(notesNode);
                }
              }
            }
            continue;
          }
          if (
            info.phType !== undefined &&
            baselineNotesSlidePlaceholders.has(info.phType)
          ) {
            item.suppressUnconsumed();
            continue;
          }
        }

        const identity = peekShapeTreeObjectIdentity(item.element);
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: notesPartPath,
            xpath: item.xpath,
            slideIndex,
            ...identity,
          },
          namespaceUri: item.namespaceUri,
          localName: item.localName,
          mechanismHint: "notes-slide-custom-object",
          loss: {
            lost: "custom-print-layout",
            retainedAs: "notes-body-only",
          },
        });
        item.suppressUnconsumed();
      }
    }

    appendUnconsumedDiagnostics(notesCursor, slideIndex, diagnostics);
    return diagnostics.length === 0 ? {} : { diagnostics };
  },
};

/**
 * Deck-level import handler for `presentation-sections` (`<p14:sectionLst>`).
 *
 * When invoked inside `createImporter`, `importPresentationSections` has already grouped
 * slides into `SectionNode`s; this handler remains idempotent and also supports standalone
 * mechanism calls against an ungrouped `DeckContainer`.
 */
export const presentationSectionsImportHandler: DeckImportHandler = {
  target: "deck",
  stage: "post-slides",
  read(context: DeckImportContext): DeckImportOutcome {
    const {
      presentationCursor,
      deckNode,
      package: pkg,
      slideNodeByPresentationSldId,
    } = context;
    const alreadySectioned = deckNode.children.some(
      (child) => child instanceof CommonNode && child.type === "section",
    );
    if (alreadySectioned) {
      return {};
    }

    let slideNodeBySldId = slideNodeByPresentationSldId;
    if (slideNodeBySldId === undefined) {
      const fallbackMap = new Map<number, CommonNode>();
      const slideNodes = deckNode.children.filter(
        (child): child is CommonNode =>
          child instanceof CommonNode && child.type === "slide",
      );
      const sldIdLst = presentationCursor.child(P_NS, "sldIdLst");
      if (sldIdLst !== undefined) {
        const sldIds = sldIdLst.children(P_NS, "sldId");
        for (let index = 0; index < sldIds.length; index += 1) {
          const rawId = sldIds[index]?.attr("id");
          const slideNode = slideNodes[index];
          if (rawId !== undefined && slideNode !== undefined) {
            const parsedId = Number(rawId);
            if (Number.isInteger(parsedId)) {
              fallbackMap.set(parsedId, slideNode);
            }
          }
        }
      }
      slideNodeBySldId = fallbackMap;
    }
    const diagnostics: ImportDiagnostic[] = [];
    importPresentationSections(
      presentationCursor,
      deckNode,
      slideNodeBySldId,
      pkg.presentationPath,
      diagnostics,
    );
    return diagnostics.length === 0 ? {} : { diagnostics };
  },
};
