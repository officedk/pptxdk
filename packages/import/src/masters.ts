import {
  slideLayoutTypeSchema,
  type ColorMap,
  type ColorMapOverride,
  type HeaderFooter,
  type SlideLayoutProps,
  type SlideMasterProps,
  type Theme,
  type ThemeOverride,
} from "@pptxdk/core";
import {
  createSlideLayoutNode,
  createSlideMasterNode,
  type CommonNode,
  type DeckContainer,
  type SlideLayoutContainer,
} from "@pptxdk/core/document";
import { z } from "zod";

import { readAlternateContent } from "#/alternate-content.ts";
import type { ImportDiagnostic } from "#/diagnostics.ts";
import type {
  DeferredSlideLink,
  ShapeTreeImportHandler,
  SlideImportHandler,
} from "#/mechanism.ts";
import {
  A_NS,
  P_NS,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  REL_THEME_OVERRIDE,
  R_NS,
} from "#/namespaces.ts";
import {
  resolvePartPath,
  type ParsedRelationshipEntry,
  type PptxPackage,
} from "#/opc.ts";
import {
  inspectUnsupportedGraphicFrame,
  peekShapeTreeObjectIdentity,
  type RelationshipIndex,
} from "#/resources.ts";
import {
  isNonDefaultColorMap,
  readColorMapElement,
  readColorMapOverride,
  readThemeOverridePart,
  readThemePart,
} from "#/themes.ts";
import { OoxmlBoolSchema, OoxmlNonNegativeIntSchema } from "#/xml-fragments.ts";
import { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";

const minimumMasterLayoutId = 2_147_483_648;
const maximumMasterLayoutId = 4_294_967_295;

const SlideMasterAttrsSchema = z.object({
  preserve: OoxmlBoolSchema.optional(),
});

const SlideLayoutAttrsSchema = z.object({
  type: slideLayoutTypeSchema.optional(),
  preserve: OoxmlBoolSchema.optional(),
  matchingName: z.string().optional(),
  userDrawn: OoxmlBoolSchema.optional(),
  showMasterSp: OoxmlBoolSchema.optional(),
  showMasterPhAnim: OoxmlBoolSchema.optional(),
});

const HeaderFooterAttrsSchema = z.object({
  sldNum: OoxmlBoolSchema.optional(),
  hdr: OoxmlBoolSchema.optional(),
  ftr: OoxmlBoolSchema.optional(),
  dt: OoxmlBoolSchema.optional(),
});

function readHeaderFooterElement(
  parentCursor: XmlElementCursor,
  partPath: string,
  diagnostics: ImportDiagnostic[],
): HeaderFooter | undefined {
  const hf = parentCursor.child(P_NS, "hf");
  if (hf === undefined) return undefined;
  const raw = hf.attrsObject(["sldNum", "hdr", "ftr", "dt"]);
  const parsed = HeaderFooterAttrsSchema.safeParse(raw);
  if (!parsed.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: partPath, xpath: hf.xpath },
      attributeOrElement: "p:hf",
      rawValue: JSON.stringify(raw),
    });
    return undefined;
  }
  const result: HeaderFooter = {
    ...(parsed.data.dt === undefined ? {} : { dateTime: parsed.data.dt }),
    ...(parsed.data.ftr === undefined ? {} : { footer: parsed.data.ftr }),
    ...(parsed.data.hdr === undefined ? {} : { header: parsed.data.hdr }),
    ...(parsed.data.sldNum === undefined
      ? {}
      : { slideNumber: parsed.data.sldNum }),
  };
  return Object.keys(result).length === 0 ? {} : result;
}

/** Allocates a deterministic unique author id within `used`, disambiguating collisions with `-2`, `-3`, ... */
export function allocateScopedAuthorId(
  requested: string | undefined,
  used: Set<string>,
  fallbackBase: string,
): string {
  const base =
    requested !== undefined && requested.trim().length > 0
      ? requested.trim()
      : fallbackBase;
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  let suffix = 2;
  while (used.has(`${base}-${suffix}`)) {
    suffix += 1;
  }
  const unique = `${base}-${suffix}`;
  used.add(unique);
  return unique;
}

/** Consumes standard identity `<p:nvGrpSpPr>` and `<p:grpSpPr>` transform boilerplate on an `spTree`. */
export function consumeIdentityShapeTreeBoilerplate(
  spTree: XmlElementCursor,
): void {
  const nvGrpSpPr = spTree.child(P_NS, "nvGrpSpPr");
  if (nvGrpSpPr !== undefined) {
    const cNvPr = nvGrpSpPr.child(P_NS, "cNvPr");
    if (cNvPr !== undefined) {
      cNvPr.consumeAttr("id");
      cNvPr.consumeAttr("name");
    }
    nvGrpSpPr.ignoreEmptyChild(P_NS, "cNvGrpSpPr");
    nvGrpSpPr.ignoreEmptyChild(P_NS, "nvPr");
  }
  const grpSpPr = spTree.child(P_NS, "grpSpPr");
  if (grpSpPr !== undefined) {
    const xfrm = grpSpPr.child(A_NS, "xfrm");
    if (xfrm !== undefined) {
      const off = xfrm.child(A_NS, "off");
      if (off !== undefined && off.attr("x") === "0" && off.attr("y") === "0") {
        // standard identity offset
      }
      const ext = xfrm.child(A_NS, "ext");
      if (
        ext !== undefined &&
        ext.attr("cx") === "0" &&
        ext.attr("cy") === "0"
      ) {
        // standard identity extent
      }
      const chOff = xfrm.child(A_NS, "chOff");
      if (
        chOff !== undefined &&
        chOff.attr("x") === "0" &&
        chOff.attr("y") === "0"
      ) {
        // standard identity child offset
      }
      const chExt = xfrm.child(A_NS, "chExt");
      if (
        chExt !== undefined &&
        chExt.attr("cx") === "0" &&
        chExt.attr("cy") === "0"
      ) {
        // standard identity child extent
      }
    }
  }
}

/** Records `unconsumed-xml` diagnostics for all unconsumed attributes/elements under a root cursor. */
export function appendUnconsumedDiagnostics(
  cursor: XmlElementCursor,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): void {
  for (const item of cursor.collectUnconsumed()) {
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: item.kind,
      location: {
        ...item.location,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
      namespaceUri: item.namespaceUri,
      localName: item.localName,
    });
  }
}

function validateMasterLayoutNumericId(
  rawId: string | undefined,
  part: string,
  xpath: string,
  diagnostics: ImportDiagnostic[],
): void {
  if (rawId === undefined) return;
  const parsed = OoxmlNonNegativeIntSchema.safeParse(rawId);
  if (
    !parsed.success ||
    parsed.data < minimumMasterLayoutId ||
    parsed.data > maximumMasterLayoutId
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { part, xpath: `${xpath}/@id` },
      attributeOrElement: "@id",
      rawValue: rawId,
    });
  }
}

const p14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const creationIdExtensionUri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

/** Consumes benign PowerPoint `<p:cSld>/<p:extLst>` creationId metadata without altering author models. */
export function consumeCommonSlideCreationId(cSld: XmlElementCursor): void {
  const extLst = cSld.child(P_NS, "extLst");
  if (extLst === undefined) return;
  for (const ext of extLst.children(P_NS, "ext")) {
    if (ext.attr("uri") !== creationIdExtensionUri) continue;
    const creationId = ext.child(p14Namespace, "creationId");
    if (creationId !== undefined) {
      creationId.consumeAttr("val");
    }
  }
}

function stripSlideIndexFromDiagnostics(
  entries: readonly ImportDiagnostic[] | undefined,
): ImportDiagnostic[] {
  if (entries === undefined || entries.length === 0) return [];
  return entries.map((entry) => {
    const { slideIndex: _ignored, ...locationRest } = entry.location;
    return {
      ...entry,
      location: locationRest,
    };
  });
}

const supportedShapeTreeNamespaces = new Set<string>([P_NS, A_NS, R_NS]);

function readSingleShapeTreeItem(
  itemCursor: XmlElementCursor,
  pkg: PptxPackage,
  partPath: string,
  slideIndex: number | undefined,
  relIndex: RelationshipIndex,
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
  containerNode?: CommonNode,
): CommonNode | undefined {
  const match = shapeTreeHandlers.find(
    (handler) =>
      itemCursor.namespaceUri === P_NS &&
      handler.localName === itemCursor.localName &&
      (handler.matches === undefined || handler.matches(itemCursor)),
  );
  if (match === undefined) {
    const identity = peekShapeTreeObjectIdentity(itemCursor.element);
    const frameInfo =
      itemCursor.namespaceUri === P_NS &&
      itemCursor.localName === "graphicFrame"
        ? inspectUnsupportedGraphicFrame(
            itemCursor,
            partPath,
            relIndex.byId,
            relIndex.entries,
          )
        : undefined;
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: {
        part: partPath,
        xpath: itemCursor.xpath,
        ...(slideIndex === undefined ? {} : { slideIndex }),
        ...identity,
      },
      namespaceUri: itemCursor.namespaceUri,
      localName: itemCursor.localName,
      ...(frameInfo?.mechanismHint === undefined
        ? {}
        : { mechanismHint: frameInfo.mechanismHint }),
      ...(frameInfo === undefined || frameInfo.cascaded.length === 0
        ? {}
        : { cascaded: frameInfo.cascaded }),
    });
    itemCursor.suppressUnconsumed();
    return undefined;
  }
  const outcome = match.read({
    cursor: itemCursor,
    part: partPath,
    slideIndex: slideIndex ?? 0,
    package: pkg,
    ...(containerNode === undefined ? {} : { containerNode }),
    relationships: relIndex.byId,
    registerDeferredSlideLink,
  });
  if (outcome.diagnostics !== undefined) {
    diagnostics.push(
      ...(slideIndex === undefined
        ? stripSlideIndexFromDiagnostics(outcome.diagnostics)
        : outcome.diagnostics),
    );
  }
  return outcome.node;
}

/** Shared `<p:spTree>` child reader for slides, slide layouts, and slide masters. */
export function readShapeTreeItems(args: {
  readonly spTree: XmlElementCursor;
  readonly pkg: PptxPackage;
  readonly partPath: string;
  readonly slideIndex: number | undefined;
  readonly relIndex: RelationshipIndex;
  readonly shapeTreeHandlers: readonly ShapeTreeImportHandler[];
  readonly registerDeferredSlideLink: (link: DeferredSlideLink) => void;
  readonly appendNode: (node: CommonNode) => void;
  readonly diagnostics: ImportDiagnostic[];
  readonly containerNode?: CommonNode;
}): void {
  const {
    spTree,
    pkg,
    partPath,
    slideIndex,
    relIndex,
    shapeTreeHandlers,
    registerDeferredSlideLink,
    appendNode,
    diagnostics,
    containerNode,
  } = args;
  consumeIdentityShapeTreeBoilerplate(spTree);
  for (const itemCursor of spTree.remainingElementChildren()) {
    if (
      itemCursor.namespaceUri === MC_NS_IMPORT_PLACEHOLDER &&
      itemCursor.localName === "AlternateContent"
    ) {
      const selection = readAlternateContent(itemCursor, (ns) =>
        supportedShapeTreeNamespaces.has(ns),
      );
      if (selection.branch === "choice") {
        for (const childCursor of selection.children) {
          const node = readSingleShapeTreeItem(
            childCursor,
            pkg,
            partPath,
            slideIndex,
            relIndex,
            shapeTreeHandlers,
            registerDeferredSlideLink,
            diagnostics,
            containerNode,
          );
          if (node !== undefined) appendNode(node);
        }
      } else if (selection.branch === "fallback") {
        let adoptedFallback = false;
        for (const childCursor of selection.children) {
          const identity = peekShapeTreeObjectIdentity(childCursor.element);
          const node = readSingleShapeTreeItem(
            childCursor,
            pkg,
            partPath,
            slideIndex,
            relIndex,
            shapeTreeHandlers,
            registerDeferredSlideLink,
            diagnostics,
            containerNode,
          );
          if (node !== undefined) {
            adoptedFallback = true;
            appendNode(node);
            diagnostics.push({
              code: "compatibility-fallback",
              category: "unsupported",
              disposition: "fallback",
              location: {
                part: partPath,
                xpath: itemCursor.xpath,
                ...(slideIndex === undefined ? {} : { slideIndex }),
                ...identity,
              },
              preferredMechanism: selection.preferredMechanism,
              selectedFallback: node.kind,
              loss: {
                lost: "native-mechanism",
                retainedAs: node.kind,
              },
            });
          }
        }
        if (!adoptedFallback) {
          itemCursor.suppressUnconsumed();
        }
      } else {
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: partPath,
            xpath: itemCursor.xpath,
            ...(slideIndex === undefined ? {} : { slideIndex }),
          },
          namespaceUri: itemCursor.namespaceUri,
          localName: itemCursor.localName,
          mechanismHint: selection.preferredMechanism,
        });
        itemCursor.suppressUnconsumed();
      }
      continue;
    }

    const node = readSingleShapeTreeItem(
      itemCursor,
      pkg,
      partPath,
      slideIndex,
      relIndex,
      shapeTreeHandlers,
      registerDeferredSlideLink,
      diagnostics,
      containerNode,
    );
    if (node !== undefined) appendNode(node);
  }
}

const MC_NS_IMPORT_PLACEHOLDER =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";

function readSlideLayoutPart(
  pkg: PptxPackage,
  layoutPartPath: string,
  ownerMasterPartPath: string,
  layoutOrder: number,
  usedLayoutIds: Set<string>,
  masterLevelSlideHandlers: readonly SlideImportHandler[],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
): SlideLayoutContainer | undefined {
  const xmlText = pkg.readTextPart(layoutPartPath);
  if (xmlText === undefined) return undefined;

  let layoutEl: ReturnType<typeof parseXmlPart>;
  try {
    layoutEl = parseXmlPart(layoutPartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath },
    });
    return undefined;
  }

  if (
    (layoutEl.namespaceURI ?? "") !== P_NS ||
    (layoutEl.localName ?? layoutEl.tagName) !== "sldLayout"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath },
    });
    return undefined;
  }

  const layoutCursor = XmlElementCursor.createRoot(
    layoutEl,
    layoutPartPath,
    "/p:sldLayout",
  );
  const rawAttrs = layoutCursor.attrsObject([
    "type",
    "preserve",
    "matchingName",
    "userDrawn",
    "showMasterSp",
    "showMasterPhAnim",
  ]);
  const parsedAttrs = SlideLayoutAttrsSchema.safeParse(rawAttrs);
  let layoutType: SlideLayoutProps["type"];
  let preserve: boolean | undefined;
  let matchingName: string | undefined;
  let userDrawn: boolean | undefined;
  let showMasterShapes: boolean | undefined;
  let showMasterPlaceholderAnimations: boolean | undefined;

  if (parsedAttrs.success) {
    layoutType = parsedAttrs.data.type;
    preserve = parsedAttrs.data.preserve;
    matchingName = parsedAttrs.data.matchingName;
    userDrawn = parsedAttrs.data.userDrawn;
    showMasterShapes = parsedAttrs.data.showMasterSp;
    showMasterPlaceholderAnimations = parsedAttrs.data.showMasterPhAnim;
  } else {
    for (const issue of parsedAttrs.error.issues) {
      const key = String(issue.path[0] ?? "sldLayout");
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "recovered",
        location: {
          part: layoutPartPath,
          xpath: `${layoutCursor.xpath}/@${key}`,
        },
        attributeOrElement: `@${key}`,
        rawValue: rawAttrs[key] ?? "",
      });
    }
    const typeRes = slideLayoutTypeSchema.safeParse(rawAttrs["type"]);
    if (typeRes.success) layoutType = typeRes.data;
    const preserveRes = OoxmlBoolSchema.safeParse(rawAttrs["preserve"]);
    if (preserveRes.success) preserve = preserveRes.data;
    if (rawAttrs["matchingName"] !== undefined) {
      matchingName = rawAttrs["matchingName"];
    }
    const userDrawnRes = OoxmlBoolSchema.safeParse(rawAttrs["userDrawn"]);
    if (userDrawnRes.success) userDrawn = userDrawnRes.data;
    const showSpRes = OoxmlBoolSchema.safeParse(rawAttrs["showMasterSp"]);
    if (showSpRes.success) showMasterShapes = showSpRes.data;
    const showPhAnimRes = OoxmlBoolSchema.safeParse(
      rawAttrs["showMasterPhAnim"],
    );
    if (showPhAnimRes.success) {
      showMasterPlaceholderAnimations = showPhAnimRes.data;
    }
  }

  const cSld = layoutCursor.child(P_NS, "cSld");
  let layoutName: string | undefined;
  if (cSld === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: layoutPartPath, xpath: layoutCursor.xpath },
    });
    return undefined;
  }
  layoutName = cSld.attr("name");
  consumeCommonSlideCreationId(cSld);
  const spTree = cSld.child(P_NS, "spTree");
  if (spTree !== undefined) {
    consumeIdentityShapeTreeBoilerplate(spTree);
  }
  const clrMapOvr = layoutCursor.child(P_NS, "clrMapOvr");
  let colorMapOverride: ColorMapOverride | undefined;
  if (clrMapOvr !== undefined) {
    const parsedOvr = readColorMapOverride(
      clrMapOvr,
      layoutPartPath,
      undefined,
      diagnostics,
    );
    if (parsedOvr !== undefined && parsedOvr !== "master") {
      colorMapOverride = parsedOvr;
    }
  }
  const layoutHeaderFooter = readHeaderFooterElement(
    layoutCursor,
    layoutPartPath,
    diagnostics,
  );

  const layoutRels = pkg.readRelationships(layoutPartPath);
  diagnostics.push(...layoutRels.diagnostics);
  const backMasterRel = layoutRels.entries.find(
    (entry) =>
      entry.Type === REL_SLIDE_MASTER && entry.TargetMode !== "External",
  );
  if (backMasterRel !== undefined) {
    const resolvedMaster = resolvePartPath(
      layoutPartPath,
      backMasterRel.Target,
    );
    if (resolvedMaster !== ownerMasterPartPath) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: layoutPartPath, xpath: backMasterRel.xpath },
        relationshipId: backMasterRel.Id,
      });
    }
  }

  let themeOverride: ThemeOverride | undefined;
  const themeOverrideRel = layoutRels.entries.find(
    (entry) =>
      entry.Type === REL_THEME_OVERRIDE && entry.TargetMode !== "External",
  );
  if (themeOverrideRel !== undefined) {
    const overridePart = resolvePartPath(
      layoutPartPath,
      themeOverrideRel.Target,
    );
    if (overridePart === undefined || !pkg.hasPart(overridePart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: layoutPartPath, xpath: themeOverrideRel.xpath },
        relationshipId: themeOverrideRel.Id,
        targetPart: overridePart ?? themeOverrideRel.Target,
      });
    } else {
      themeOverride = readThemeOverridePart(
        pkg,
        overridePart,
        undefined,
        diagnostics,
      );
    }
  }

  const layoutId = allocateScopedAuthorId(
    layoutName,
    usedLayoutIds,
    `layout-${layoutOrder}`,
  );
  const props: SlideLayoutProps = {
    id: layoutId,
    ...(layoutName === undefined ? {} : { name: layoutName }),
    ...(layoutType === undefined ? {} : { type: layoutType }),
    ...(preserve === undefined ? {} : { preserve }),
    ...(matchingName === undefined ? {} : { matchingName }),
    ...(userDrawn === undefined ? {} : { userDrawn }),
    ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
    ...(showMasterPlaceholderAnimations === undefined
      ? {}
      : { showMasterPlaceholderAnimations }),
    ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
    ...(themeOverride === undefined ? {} : { themeOverride }),
    ...(layoutHeaderFooter === undefined
      ? {}
      : { headerFooter: layoutHeaderFooter }),
  };
  const layoutNode = createSlideLayoutNode(props);
  if (spTree !== undefined) {
    if (shapeTreeHandlers.length > 0) {
      const relIndex: RelationshipIndex = {
        entries: layoutRels.entries,
        byId: new Map(layoutRels.entries.map((entry) => [entry.Id, entry])),
        diagnostics: [],
      };
      readShapeTreeItems({
        spTree,
        pkg,
        partPath: layoutPartPath,
        slideIndex: undefined,
        relIndex,
        shapeTreeHandlers,
        registerDeferredSlideLink,
        appendNode: (node) => layoutNode.append(node),
        diagnostics,
        containerNode: layoutNode,
      });
    } else {
      consumeIdentityShapeTreeBoilerplate(spTree);
    }
  }
  for (const slideHandler of masterLevelSlideHandlers) {
    const outcome = slideHandler.read({
      slideCursor: layoutCursor,
      slideNode: layoutNode,
      part: layoutPartPath,
      slideIndex: 0,
      package: pkg,
    });
    diagnostics.push(...stripSlideIndexFromDiagnostics(outcome.diagnostics));
  }

  appendUnconsumedDiagnostics(layoutCursor, undefined, diagnostics);
  return layoutNode;
}

function readSlideMasterPart(
  pkg: PptxPackage,
  masterPartPath: string,
  masterOrder: number,
  usedDeckIds: Set<string>,
  themesByPartPath: Map<string, Theme>,
  layoutNodeByPartPath: Map<string, SlideLayoutContainer>,
  masterLevelSlideHandlers: readonly SlideImportHandler[],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void,
  diagnostics: ImportDiagnostic[],
) {
  const xmlText = pkg.readTextPart(masterPartPath);
  if (xmlText === undefined) return undefined;

  let masterEl: ReturnType<typeof parseXmlPart>;
  try {
    masterEl = parseXmlPart(masterPartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return undefined;
  }

  if (
    (masterEl.namespaceURI ?? "") !== P_NS ||
    (masterEl.localName ?? masterEl.tagName) !== "sldMaster"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath },
    });
    return undefined;
  }

  const masterCursor = XmlElementCursor.createRoot(
    masterEl,
    masterPartPath,
    "/p:sldMaster",
  );
  const rawAttrs = masterCursor.attrsObject(["preserve"]);
  const parsedAttrs = SlideMasterAttrsSchema.safeParse(rawAttrs);
  let preserve: boolean | undefined;
  if (!parsedAttrs.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: masterPartPath,
        xpath: `${masterCursor.xpath}/@preserve`,
      },
      attributeOrElement: "@preserve",
      rawValue: rawAttrs["preserve"] ?? "",
    });
  } else {
    preserve = parsedAttrs.data.preserve;
  }

  const cSld = masterCursor.child(P_NS, "cSld");
  if (cSld === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: masterPartPath, xpath: masterCursor.xpath },
    });
    return undefined;
  }
  const masterName = cSld.attr("name");
  consumeCommonSlideCreationId(cSld);
  const spTree = cSld.child(P_NS, "spTree");

  const clrMap = masterCursor.child(P_NS, "clrMap");
  let colorMap: ColorMap | undefined;
  if (clrMap !== undefined) {
    const parsedMap = readColorMapElement(
      clrMap,
      masterPartPath,
      undefined,
      diagnostics,
    );
    if (isNonDefaultColorMap(parsedMap)) {
      colorMap = parsedMap;
    }
  }

  const txStyles = masterCursor.child(P_NS, "txStyles");
  if (txStyles !== undefined) {
    txStyles.ignoreEmptyChild(P_NS, "titleStyle");
    txStyles.ignoreEmptyChild(P_NS, "bodyStyle");
    txStyles.ignoreEmptyChild(P_NS, "otherStyle");
  }
  const masterHeaderFooter = readHeaderFooterElement(
    masterCursor,
    masterPartPath,
    diagnostics,
  );

  const masterRels = pkg.readRelationships(masterPartPath);
  diagnostics.push(...masterRels.diagnostics);
  const masterRelById = new Map<string, ParsedRelationshipEntry>();
  for (const rel of masterRels.entries) {
    masterRelById.set(rel.Id, rel);
  }

  let masterTheme: Theme | undefined;
  const themeRel = masterRels.entries.find(
    (entry) => entry.Type === REL_THEME && entry.TargetMode !== "External",
  );
  if (themeRel !== undefined) {
    const themePartPath = resolvePartPath(masterPartPath, themeRel.Target);
    if (themePartPath === undefined || !pkg.hasPart(themePartPath)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: masterPartPath, xpath: themeRel.xpath },
        relationshipId: themeRel.Id,
        targetPart: themePartPath ?? themeRel.Target,
      });
    } else {
      masterTheme = readThemePart(
        pkg,
        themePartPath,
        themesByPartPath,
        diagnostics,
      );
    }
  }

  const masterId = allocateScopedAuthorId(
    masterName,
    usedDeckIds,
    `master-${masterOrder}`,
  );
  const masterProps: SlideMasterProps = {
    id: masterId,
    ...(masterName === undefined ? {} : { name: masterName }),
    ...(preserve === undefined ? {} : { preserve }),
    ...(masterTheme === undefined ? {} : { theme: masterTheme }),
    ...(colorMap === undefined ? {} : { colorMap }),
    ...(masterHeaderFooter === undefined
      ? {}
      : { headerFooter: masterHeaderFooter }),
  };
  const masterNode = createSlideMasterNode(masterProps);
  if (spTree !== undefined) {
    if (shapeTreeHandlers.length > 0) {
      readShapeTreeItems({
        spTree,
        pkg,
        partPath: masterPartPath,
        slideIndex: undefined,
        relIndex: {
          entries: masterRels.entries,
          byId: masterRelById,
          diagnostics: [],
        },
        shapeTreeHandlers,
        registerDeferredSlideLink,
        appendNode: (node) => masterNode.append(node),
        diagnostics,
        containerNode: masterNode,
      });
    } else {
      consumeIdentityShapeTreeBoilerplate(spTree);
    }
  }
  for (const slideHandler of masterLevelSlideHandlers) {
    const outcome = slideHandler.read({
      slideCursor: masterCursor,
      slideNode: masterNode,
      part: masterPartPath,
      slideIndex: 0,
      package: pkg,
    });
    diagnostics.push(...stripSlideIndexFromDiagnostics(outcome.diagnostics));
  }

  const usedLayoutIds = new Set<string>();
  const sldLayoutIdLst = masterCursor.child(P_NS, "sldLayoutIdLst");
  if (sldLayoutIdLst !== undefined) {
    const layoutIdEntries = sldLayoutIdLst.children(P_NS, "sldLayoutId");
    for (let index = 0; index < layoutIdEntries.length; index += 1) {
      const layoutIdCursor = layoutIdEntries[index]!;
      validateMasterLayoutNumericId(
        layoutIdCursor.attr("id"),
        masterPartPath,
        layoutIdCursor.xpath,
        diagnostics,
      );
      const rId = layoutIdCursor.attr("id", R_NS);
      if (rId === undefined) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: "",
        });
        continue;
      }
      const rel = masterRelById.get(rId);
      if (
        rel === undefined ||
        rel.Type !== REL_SLIDE_LAYOUT ||
        rel.TargetMode === "External"
      ) {
        diagnostics.push({
          code: "broken-relationship",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: rId,
        });
        continue;
      }
      const layoutPart = resolvePartPath(masterPartPath, rel.Target);
      if (layoutPart === undefined || !pkg.hasPart(layoutPart)) {
        diagnostics.push({
          code: "missing-target-part",
          category: "corrupt-input",
          disposition: "omitted",
          location: { part: masterPartPath, xpath: layoutIdCursor.xpath },
          relationshipId: rId,
          targetPart: layoutPart ?? rel.Target,
        });
        continue;
      }
      const layoutNode = readSlideLayoutPart(
        pkg,
        layoutPart,
        masterPartPath,
        index + 1,
        usedLayoutIds,
        masterLevelSlideHandlers,
        shapeTreeHandlers,
        registerDeferredSlideLink,
        diagnostics,
      );
      if (layoutNode !== undefined) {
        masterNode.append(layoutNode);
        layoutNodeByPartPath.set(layoutPart, layoutNode);
      }
    }
  }

  appendUnconsumedDiagnostics(masterCursor, undefined, diagnostics);
  return masterNode;
}

/**
 * Traverses `<p:sldMasterIdLst>` and `<p:sldLayoutIdLst>` relationships to reconstruct explicit
 * `SlideMaster` and `SlideLayout` ownership (including idle layouts) onto `deckNode`.
 */
export function importMastersAndLayouts(
  presCursor: XmlElementCursor,
  pkg: PptxPackage,
  presRelById: ReadonlyMap<string, ParsedRelationshipEntry>,
  deckNode: DeckContainer,
  usedDeckIds: Set<string>,
  diagnostics: ImportDiagnostic[],
  slideHandlers: readonly SlideImportHandler[] = [],
  shapeTreeHandlers: readonly ShapeTreeImportHandler[] = [],
  registerDeferredSlideLink: (link: DeferredSlideLink) => void = () => {},
): ReadonlyMap<string, SlideLayoutContainer> {
  const masterLevelSlideHandlers = slideHandlers.filter(
    (handler) =>
      handler.slot === "background" || handler.slot === "header-footer",
  );
  const layoutNodeByPartPath = new Map<string, SlideLayoutContainer>();
  const themesByPartPath = new Map<string, Theme>();
  const masterIdLst = presCursor.child(P_NS, "sldMasterIdLst");
  if (masterIdLst === undefined) {
    return layoutNodeByPartPath;
  }
  const masterIdEntries = masterIdLst.children(P_NS, "sldMasterId");
  for (let index = 0; index < masterIdEntries.length; index += 1) {
    const masterIdCursor = masterIdEntries[index]!;
    validateMasterLayoutNumericId(
      masterIdCursor.attr("id"),
      pkg.presentationPath,
      masterIdCursor.xpath,
      diagnostics,
    );
    const rId = masterIdCursor.attr("id", R_NS);
    if (rId === undefined) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: "",
      });
      continue;
    }
    const rel = presRelById.get(rId);
    if (
      rel === undefined ||
      rel.Type !== REL_SLIDE_MASTER ||
      rel.TargetMode === "External"
    ) {
      diagnostics.push({
        code: "broken-relationship",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: rId,
      });
      continue;
    }
    const masterPart = resolvePartPath(pkg.presentationPath, rel.Target);
    if (masterPart === undefined || !pkg.hasPart(masterPart)) {
      diagnostics.push({
        code: "missing-target-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { part: pkg.presentationPath, xpath: masterIdCursor.xpath },
        relationshipId: rId,
        targetPart: masterPart ?? rel.Target,
      });
      continue;
    }
    const masterNode = readSlideMasterPart(
      pkg,
      masterPart,
      index + 1,
      usedDeckIds,
      themesByPartPath,
      layoutNodeByPartPath,
      masterLevelSlideHandlers,
      shapeTreeHandlers,
      registerDeferredSlideLink,
      diagnostics,
    );
    if (masterNode !== undefined) {
      deckNode.append(masterNode);
    }
  }
  return layoutNodeByPartPath;
}
