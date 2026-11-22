import type { RegisteredCapabilityModule } from "@pptxdk/core/capability";
import {
  type CommonNode,
  type SlideLayoutContainer,
  createDeckNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { z } from "zod";

import {
  ImportOptionsSchema,
  type ImportDiagnostic,
  type ImportOptions,
  type ImportReport,
} from "#/diagnostics.ts";
import {
  allocateScopedAuthorId,
  appendUnconsumedDiagnostics,
  consumeIdentityShapeTreeBoilerplate,
  importMastersAndLayouts,
} from "#/masters.ts";
import type {
  DeckImportHandler,
  ImportMechanismHandler,
  ShapeTreeImportHandler,
  SlideImportHandler,
} from "#/mechanism.ts";
import {
  A_NS,
  P_NS,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  R_NS,
  emuToPx,
} from "#/namespaces.ts";
import {
  openPptxPackage,
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import {
  ExtentEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlIntSchema,
} from "#/xml-fragments.ts";
import {
  MalformedXmlPartError,
  XmlElementCursor,
  parseXmlPart,
} from "#/xml-reader.ts";

export type ImportResult = {
  readonly document: CommonNode;
  readonly report: ImportReport;
};

const defaultSlideWidthPx = 960;
const defaultSlideHeightPx = 540;

const SlideSizeFragmentSchema = ExtentEmuFragmentSchema.extend({
  type: z.string().optional(),
});

const SlideShowAttrSchema = z.object({
  show: OoxmlBoolSchema.optional(),
});

type RegisteredImportHandlers = {
  readonly shapeTree: readonly ShapeTreeImportHandler[];
  readonly slide: readonly SlideImportHandler[];
  readonly deck: readonly DeckImportHandler[];
};

function isImportMechanismHandler(
  value: unknown,
): value is ImportMechanismHandler {
  if (typeof value !== "object" || value === null || !("target" in value)) {
    return false;
  }
  const target = (value as { readonly target?: unknown }).target;
  return target === "shape-tree" || target === "slide" || target === "deck";
}

function collectHandlers(
  capabilities: readonly RegisteredCapabilityModule[],
): RegisteredImportHandlers {
  const shapeTree: ShapeTreeImportHandler[] = [];
  const slide: SlideImportHandler[] = [];
  const deck: DeckImportHandler[] = [];
  for (const mod of capabilities) {
    if (!("mechanisms" in mod) || mod.mechanisms === undefined) continue;
    for (const mechanism of mod.mechanisms) {
      if (!isImportMechanismHandler(mechanism.import)) continue;
      const handler = mechanism.import;
      switch (handler.target) {
        case "shape-tree":
          shapeTree.push(handler);
          break;
        case "slide":
          slide.push(handler);
          break;
        case "deck":
          deck.push(handler);
          break;
      }
    }
  }
  return { shapeTree, slide, deck };
}

function consumePresentationHeaderBoilerplate(
  presCursor: XmlElementCursor,
): void {
  const notesSz = presCursor.child(P_NS, "notesSz");
  if (notesSz !== undefined) {
    notesSz.consumeAttr("cx");
    notesSz.consumeAttr("cy");
  }
}

function readSlideSize(
  presCursor: XmlElementCursor,
  diagnostics: ImportDiagnostic[],
): { readonly width: number; readonly height: number } {
  const sldSz = presCursor.child(P_NS, "sldSz");
  if (sldSz === undefined) {
    return { width: defaultSlideWidthPx, height: defaultSlideHeightPx };
  }
  const raw = sldSz.attrsObject(["cx", "cy", "type"]);
  const parsed = SlideSizeFragmentSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part: presCursor.part, xpath: sldSz.xpath },
      attributeOrElement: "p:sldSz",
      rawValue: JSON.stringify(raw),
    });
    return { width: defaultSlideWidthPx, height: defaultSlideHeightPx };
  }
  return {
    width: emuToPx(parsed.data.cx),
    height: emuToPx(parsed.data.cy),
  };
}

function readSlidePart(
  pkg: PptxPackage,
  slidePartPath: string,
  slideIndex: number,
  usedSlideIds: Set<string>,
  layoutNodeByPartPath: ReadonlyMap<string, SlideLayoutContainer>,
  handlers: RegisteredImportHandlers,
  diagnostics: ImportDiagnostic[],
): CommonNode | undefined {
  const xmlText = pkg.readTextPart(slidePartPath);
  if (xmlText === undefined) return undefined;

  let slideEl: ReturnType<typeof parseXmlPart>;
  try {
    slideEl = parseXmlPart(slidePartPath, xmlText);
  } catch (error) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: slidePartPath,
        slideIndex,
      },
    });
    void error;
    return undefined;
  }

  if (
    (slideEl.namespaceURI ?? "") !== P_NS ||
    (slideEl.localName ?? slideEl.tagName) !== "sld"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: slidePartPath, slideIndex },
    });
    return undefined;
  }

  const slideCursor = XmlElementCursor.createRoot(
    slideEl,
    slidePartPath,
    "/p:sld",
  );
  const rawShow = slideCursor.attrsObject(["show"]);
  const parsedShow = SlideShowAttrSchema.safeParse(rawShow);
  let hidden = false;
  if (!parsedShow.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: slidePartPath,
        xpath: `${slideCursor.xpath}/@show`,
        slideIndex,
      },
      attributeOrElement: "@show",
      rawValue: rawShow["show"] ?? "",
    });
  } else if (parsedShow.data.show === false) {
    hidden = true;
  }

  const cSld = slideCursor.child(P_NS, "cSld");
  if (cSld === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: slidePartPath, xpath: slideCursor.xpath, slideIndex },
    });
    return undefined;
  }

  const slideRels = pkg.readRelationships(slidePartPath);
  diagnostics.push(...slideRels.diagnostics);
  const layoutRel = slideRels.entries.find(
    (entry) =>
      entry.Type === REL_SLIDE_LAYOUT && entry.TargetMode !== "External",
  );
  const resolvedLayoutPath =
    layoutRel === undefined
      ? undefined
      : resolvePartPath(slidePartPath, layoutRel.Target);
  const layoutNode =
    resolvedLayoutPath === undefined
      ? undefined
      : layoutNodeByPartPath.get(resolvedLayoutPath);

  const rawSlideName = cSld.attr("name");
  const slideId = allocateScopedAuthorId(
    rawSlideName,
    usedSlideIds,
    `slide-${slideIndex + 1}`,
  );
  const slideNode = createSlideNode({
    id: slideId,
    ...(layoutNode === undefined ? {} : { layout: layoutNode }),
    ...(hidden ? { hidden: true } : {}),
  });

  const clrMapOvr = slideCursor.child(P_NS, "clrMapOvr");
  if (clrMapOvr !== undefined) {
    clrMapOvr.ignoreEmptyChild(A_NS, "masterClrMapping");
  }

  const spTree = cSld.child(P_NS, "spTree");
  if (spTree !== undefined) {
    consumeIdentityShapeTreeBoilerplate(spTree);
    for (const itemCursor of spTree.remainingElementChildren()) {
      const match = handlers.shapeTree.find(
        (handler) =>
          itemCursor.namespaceUri === P_NS &&
          handler.localName === itemCursor.localName &&
          (handler.matches === undefined || handler.matches(itemCursor)),
      );
      if (match === undefined) {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: slidePartPath,
            xpath: itemCursor.xpath,
            slideIndex,
          },
          namespaceUri: itemCursor.namespaceUri,
          localName: itemCursor.localName,
        });
        continue;
      }
      const outcome = match.read({
        cursor: itemCursor,
        part: slidePartPath,
        slideIndex,
        package: pkg,
      });
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
      if (outcome.node !== undefined) {
        slideNode.append(outcome.node);
      }
    }
  }

  for (const slideHandler of handlers.slide) {
    const outcome = slideHandler.read({
      slideCursor,
      slideNode,
      part: slidePartPath,
      slideIndex,
      package: pkg,
    });
    if (outcome.diagnostics !== undefined) {
      diagnostics.push(...outcome.diagnostics);
    }
  }

  appendUnconsumedDiagnostics(slideCursor, slideIndex, diagnostics);
  return slideNode;
}

/**
 * Creates a Transitional PPTX importer bound to the caller's registered capability modules.
 *
 * The returned function owns OPC package traversal, XML cursor tracking, and structured diagnostics,
 * while delegating native object/slide parsing to mechanisms declared by `capabilities`.
 */
export function createImporter(
  capabilities: readonly RegisteredCapabilityModule[],
): (input: Uint8Array, options?: ImportOptions) => ImportResult {
  const handlers = collectHandlers(capabilities);

  return function importPptx(
    input: Uint8Array,
    options: ImportOptions = {},
  ): ImportResult {
    ImportOptionsSchema.parse(options);
    if (options.signal?.aborted) {
      throw options.signal.reason;
    }

    const pkg = openPptxPackage(input);
    const diagnostics: ImportDiagnostic[] = [...pkg.relDiagnostics];

    const presCursor = pkg.presentationRoot;
    consumePresentationHeaderBoilerplate(presCursor);
    const size = readSlideSize(presCursor, diagnostics);
    const deckNode = createDeckNode({
      width: size.width,
      height: size.height,
    });

    const relById = new Map<string, ParsedRelationshipEntry>();
    for (const rel of pkg.presentationRels) {
      relById.set(rel.Id, rel);
    }

    const usedDeckIds = new Set<string>();
    const layoutNodeByPartPath = importMastersAndLayouts(
      presCursor,
      pkg,
      relById,
      deckNode,
      usedDeckIds,
      diagnostics,
    );

    let slideCount = 0;
    const sldIdLst = presCursor.child(P_NS, "sldIdLst");
    if (sldIdLst !== undefined) {
      const sldIds = sldIdLst.children(P_NS, "sldId");
      for (let index = 0; index < sldIds.length; index += 1) {
        if (options.signal?.aborted) {
          throw options.signal.reason;
        }
        const sldIdCursor = sldIds[index]!;
        const rawId = sldIdCursor.attr("id");
        if (rawId !== undefined && !OoxmlIntSchema.safeParse(rawId).success) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "recovered",
            location: {
              part: pkg.presentationPath,
              xpath: `${sldIdCursor.xpath}/@id`,
              slideIndex: index,
            },
            attributeOrElement: "@id",
            rawValue: rawId,
          });
        }
        const rId = sldIdCursor.attr("id", R_NS);
        if (rId === undefined) {
          diagnostics.push({
            code: "broken-relationship",
            category: "corrupt-input",
            disposition: "omitted",
            location: {
              part: pkg.presentationPath,
              xpath: sldIdCursor.xpath,
              slideIndex: index,
            },
            relationshipId: "",
          });
          continue;
        }
        const rel = relById.get(rId);
        if (
          rel === undefined ||
          rel.Type !== REL_SLIDE ||
          rel.TargetMode === "External"
        ) {
          diagnostics.push({
            code: "broken-relationship",
            category: "corrupt-input",
            disposition: "omitted",
            location: {
              part: pkg.presentationPath,
              xpath: sldIdCursor.xpath,
              slideIndex: index,
            },
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
            location: {
              part: pkg.presentationPath,
              xpath: sldIdCursor.xpath,
              slideIndex: index,
            },
            relationshipId: rId,
            targetPart: targetPart ?? rel.Target,
          });
          continue;
        }

        const slideNode = readSlidePart(
          pkg,
          targetPart,
          index,
          usedDeckIds,
          layoutNodeByPartPath,
          handlers,
          diagnostics,
        );
        if (slideNode !== undefined) {
          deckNode.append(slideNode);
          slideCount += 1;
        }
      }
    }

    for (const deckHandler of handlers.deck) {
      const outcome = deckHandler.read({
        presentationCursor: presCursor,
        deckNode,
        package: pkg,
      });
      if (outcome.diagnostics !== undefined) {
        diagnostics.push(...outcome.diagnostics);
      }
    }

    appendUnconsumedDiagnostics(presCursor, undefined, diagnostics);

    const report: ImportReport = Object.freeze({
      evidence: "import",
      sourceFormat: "ECMA-376 5th edition Transitional",
      slideCount,
      diagnostics: Object.freeze(diagnostics),
    });

    return { document: deckNode, report };
  };
}

export { MalformedXmlPartError };
