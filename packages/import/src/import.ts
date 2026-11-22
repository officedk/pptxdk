import type { ColorMapOverride, ThemeOverride } from "@pptxdk/core";
import type {
  NativeColorSlot,
  PlaceholderMap,
  RegisteredCapabilityModule,
} from "@pptxdk/core/capability";
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
  consumeCommonSlideCreationId,
  importMastersAndLayouts,
  readShapeTreeItems,
} from "#/masters.ts";
import type {
  DeckImportHandler,
  DeferredSlideLink,
  ImportMechanismHandler,
  ShapeTreeImportHandler,
  SlideImportHandler,
} from "#/mechanism.ts";
import {
  P_NS,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_THEME_OVERRIDE,
  R_NS,
  emuToPx,
} from "#/namespaces.ts";
import { inspectNotesAndHandoutMasters } from "#/notes-masters.ts";
import {
  openPptxPackage,
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import { buildRelationshipIndex } from "#/resources.ts";
import { importPresentationSections } from "#/sections.ts";
import { readColorMapOverride, readThemeOverridePart } from "#/themes.ts";
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
  showMasterSp: OoxmlBoolSchema.optional(),
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
        case "text-inline":
          // Consumed directly by rich-text paragraph readers inside shape-tree/slide mechanisms.
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
  clrMapByLayoutPartPath: ReadonlyMap<
    string,
    Readonly<Record<string, NativeColorSlot>>
  >,
  handlers: RegisteredImportHandlers,
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
): { readonly node: CommonNode; readonly slideId: string } | undefined {
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

  const relIndex = buildRelationshipIndex(pkg, slidePartPath);
  diagnostics.push(...relIndex.diagnostics);

  const slideCursor = XmlElementCursor.createRoot(
    slideEl,
    slidePartPath,
    "/p:sld",
  );
  const rawShow = slideCursor.attrsObject(["show", "showMasterSp"]);
  const parsedShow = SlideShowAttrSchema.safeParse(rawShow);
  let hidden = false;
  let showMasterShapes: boolean | undefined;
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
  } else {
    if (parsedShow.data.show === false) {
      hidden = true;
    }
    showMasterShapes = parsedShow.data.showMasterSp;
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
  consumeCommonSlideCreationId(cSld);

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
  const layoutPlaceholders = layoutNode?.props["placeholders"] as
    | PlaceholderMap
    | undefined;
  const layoutClrMap =
    resolvedLayoutPath === undefined
      ? undefined
      : clrMapByLayoutPartPath.get(resolvedLayoutPath);

  const clrMapOvr = slideCursor.child(P_NS, "clrMapOvr");
  let colorMapOverride: ColorMapOverride | undefined;
  if (clrMapOvr !== undefined) {
    const parsedOvr = readColorMapOverride(
      clrMapOvr,
      slidePartPath,
      slideIndex,
      diagnostics,
    );
    if (parsedOvr !== undefined && parsedOvr !== "master") {
      colorMapOverride = parsedOvr;
    }
  }

  let themeOverride: ThemeOverride | undefined;
  const themeOverrideRel = slideRels.entries.find(
    (entry) =>
      entry.Type === REL_THEME_OVERRIDE && entry.TargetMode !== "External",
  );
  if (themeOverrideRel !== undefined) {
    const overridePart = resolvePartPath(
      slidePartPath,
      themeOverrideRel.Target,
    );
    if (overridePart === undefined || !pkg.hasPart(overridePart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: {
          part: slidePartPath,
          xpath: themeOverrideRel.xpath,
          slideIndex,
        },
        relationshipId: themeOverrideRel.Id,
        targetPart: overridePart ?? themeOverrideRel.Target,
      });
    } else {
      themeOverride = readThemeOverridePart(
        pkg,
        overridePart,
        slideIndex,
        diagnostics,
      );
    }
  }

  const rawSlideName = cSld.attr("name");
  const slideId = allocateScopedAuthorId(
    rawSlideName,
    usedSlideIds,
    `slide-${slideIndex + 1}`,
  );
  const slideNode = createSlideNode({
    id: slideId,
    ...(layoutNode === undefined ? {} : { layout: layoutNode }),
    ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
    ...(themeOverride === undefined ? {} : { themeOverride }),
    ...(hidden ? { hidden: true } : {}),
    ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
  });

  const spTree = cSld.child(P_NS, "spTree");
  if (spTree !== undefined) {
    readShapeTreeItems({
      spTree,
      pkg,
      partPath: slidePartPath,
      slideIndex,
      relIndex,
      shapeTreeHandlers: handlers.shapeTree,
      registerDeferredSlideLink,
      appendNode: (node) => slideNode.append(node),
      diagnostics,
      containerNode: slideNode,
      ...(layoutPlaceholders === undefined ? {} : { layoutPlaceholders }),
      ...(layoutClrMap === undefined ? {} : { clrMap: layoutClrMap }),
    });
  }

  for (const slideHandler of handlers.slide) {
    if (
      slideHandler.matches !== undefined &&
      !slideHandler.matches(slideCursor)
    ) {
      continue;
    }
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
  return { node: slideNode, slideId };
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
    const clrMapByLayoutPartPath = new Map<
      string,
      Readonly<Record<string, NativeColorSlot>>
    >();
    const slideIdByPartPath = new Map<string, string>();
    const slideNodeByPresentationSldId = new Map<number, CommonNode>();
    const deferredSlideLinks: DeferredSlideLink[] = [];
    const registerDeferredSlideLink = (link: DeferredSlideLink): void => {
      deferredSlideLinks.push(link);
    };
    const layoutNodeByPartPath = importMastersAndLayouts(
      presCursor,
      pkg,
      relById,
      deckNode,
      usedDeckIds,
      diagnostics,
      handlers.slide,
      handlers.shapeTree,
      registerDeferredSlideLink,
      clrMapByLayoutPartPath,
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
        const parsedSlideNumId =
          rawId === undefined ? undefined : OoxmlIntSchema.safeParse(rawId);
        if (parsedSlideNumId !== undefined && !parsedSlideNumId.success) {
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

        const slideResult = readSlidePart(
          pkg,
          targetPart,
          index,
          usedDeckIds,
          layoutNodeByPartPath,
          clrMapByLayoutPartPath,
          handlers,
          registerDeferredSlideLink,
          diagnostics,
        );
        if (slideResult !== undefined) {
          slideIdByPartPath.set(targetPart, slideResult.slideId);
          if (parsedSlideNumId?.success) {
            slideNodeByPresentationSldId.set(
              parsedSlideNumId.data,
              slideResult.node,
            );
          }
          deckNode.append(slideResult.node);
          slideCount += 1;
        }
      }
    }

    for (const pending of deferredSlideLinks) {
      const targetSlideId = slideIdByPartPath.get(pending.targetPart);
      if (targetSlideId === undefined) {
        diagnostics.push({
          code: "missing-target-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: pending.location,
          relationshipId: pending.relationshipId,
          targetPart: pending.targetPart,
        });
        continue;
      }
      pending.apply(targetSlideId);
    }

    inspectNotesAndHandoutMasters(presCursor, pkg, relById, diagnostics);
    importPresentationSections(
      presCursor,
      deckNode,
      slideNodeByPresentationSldId,
      pkg.presentationPath,
      diagnostics,
    );

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
