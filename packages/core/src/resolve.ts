import { readAuthorId, validateAuthorIdScopes } from "#/author-ids.ts";
import {
  parseDeckProps,
  parseLinkTarget,
  parseSectionProps,
  parseSlideCoreProps,
  xml10StringSchema,
} from "#/author-schemas.ts";
import {
  componentLabel,
  declarationIdentity,
  isDeclarationNode,
  type DeclarationChild,
  type DeclarationContainer,
  type DeclarationNode,
} from "#/declaration.ts";
import { BuildError, CapabilityError } from "#/diagnostics.ts";
import {
  LayoutPass,
  validateBoundObjectCapability,
  type LayoutPassResult,
} from "#/layout.ts";
import {
  footerPlaceholdersSchema,
  isSlideLayoutDefinition,
  slideLayoutDataPropsSchema,
  slideMasterDataPropsSchema,
  type FooterPlaceholderSpec,
  type FooterPlaceholders,
  type HeaderFooter,
  type SlideLayoutDefinition,
  type SlideLayoutObservation,
  type SlideMasterObservation,
} from "#/masters.ts";
import { getMorphKeyLabel, isMorphKey, type MorphKey } from "#/morph.ts";
import {
  matchMasterPlaceholderSpec,
  nativePlaceholderRefSchema,
  type MasterPlaceholderMap,
  type MasterTextStyles,
  type PlaceholderMap,
} from "#/placeholders.ts";
import { createBuildQuery, type QueryContentNode } from "#/query.ts";
import {
  registerInternalResolution,
  type InternalResolution,
  type InternalResolvedSlide,
  type InternalResolvedSlideLayout,
  type InternalResolvedSlideMaster,
  type ResolutionDimensions,
  type ResolvedObjectLayer,
  type ResolvedSectionRecord,
  type ResolvedSlideCapabilityRecord,
} from "#/resolution.ts";
import { resolveOptionsSchema } from "#/stage-options.ts";
import {
  createThemeContext,
  defaultColorMap,
  defineTheme,
  isTheme,
  textStyleFrom,
  type ColorMap,
  type ColorMapOverride,
  type Theme,
  type ThemeContext,
  type ThemeInput,
  type ThemeOverride,
  type ThemeScope,
} from "#/theme.ts";
import type {
  BoundContainerCollaboration,
  BoundObjectCapability,
  BoundSlideCapability,
  CapabilityChild,
  CapabilityContribution,
  CapabilityElement,
  CapabilityInput,
  CapabilityLayoutContract,
  CapabilityModule,
  CapabilityResourceClaim,
  CapabilitySubPartObservation,
  CapabilityWarning,
  Diagnostic,
  Extension,
  GeometryBox,
  LayoutChildReference,
  LayoutObservation,
  NativeElementKind,
  ObjectCapabilityModule,
  ObjectObservation,
  ObjectResolveContext,
  PlaceholderResolutionContext,
  RegisteredCapabilityModule,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  ResolvedLink,
  ResolvedObjectChildren,
  ResolvedSlide as ResolvedSlideCapabilityOutput,
  SlideCapabilityModule,
  SlideMasterInheritanceContext,
  SlideObservation,
  TransparentCapabilityModule,
} from "#/types.ts";

type ParsedSlideNode = {
  readonly node: DeclarationNode;
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly layoutKey?: string;
  readonly hasExplicitLayout: boolean;
  readonly colorMapOverride?: ColorMapOverride;
  readonly themeOverride?: ThemeOverride;
  readonly section?: string;
  readonly hidden?: boolean;
  readonly showMasterPlaceholderAnimations?: boolean;
  readonly showMasterShapes?: boolean;
};

type ParsedLayoutSpec = {
  readonly node: DeclarationNode;
  readonly data: Omit<
    InternalResolvedSlideLayout,
    | "key"
    | "masterKey"
    | "placeholders"
    | "background"
    | "footerPlaceholders"
    | "layers"
    | "slideCapabilities"
  >;
  readonly placeholders?: PlaceholderMap;
  readonly background?: unknown;
  readonly headerFooter?: HeaderFooter;
  readonly footerPlaceholders?: FooterPlaceholders;
  readonly key: string;
  readonly path: string;
  readonly props: Readonly<Record<string, unknown>>;
};

type ParsedMasterSpec = {
  readonly node: DeclarationNode;
  readonly canvasChildren: readonly DeclarationNode[];
  readonly key: string;
  readonly id: string;
  readonly name?: string;
  readonly preserve?: boolean;
  readonly theme: Theme;
  readonly colorMap: ColorMap;
  readonly explicitColorMap?: ColorMap;
  readonly style?: Readonly<Record<string, unknown>>;
  readonly placeholders?: MasterPlaceholderMap;
  readonly textStyles?: MasterTextStyles;
  readonly background?: unknown;
  readonly headerFooter?: HeaderFooter;
  readonly footerPlaceholders?: FooterPlaceholders;
  readonly path: string;
  readonly props: Readonly<Record<string, unknown>>;
  readonly layouts: readonly ParsedLayoutSpec[];
};

type ParsedDocument = {
  readonly masters: readonly ParsedMasterSpec[];
  readonly layoutIdByKey: ReadonlyMap<string, string>;
  readonly slides: readonly ParsedSlideNode[];
  readonly sections: readonly ResolvedSectionRecord[];
};

type MorphObject = {
  readonly kind: NativeElementKind;
  readonly name: string;
  readonly path: string;
};

type DocumentResolveState = {
  readonly usedExtensions: Set<Extension>;
  readonly warnings: Diagnostic[];
  readonly resourceClaims: Map<
    string,
    { identity: string; path: string; characters: string }
  >;
  readonly morphNames: Map<MorphKey, string>;
  nextMorphName: number;
  readonly reservedObjectNames: ReadonlySet<string>;
  readonly slideIndexById: ReadonlyMap<string, number>;
};

type SlideResolveState = {
  readonly index: number;
  nextId: number;
  nextNodeOrder: number;
  readonly diagnostics: Diagnostic[];
  readonly layouts: LayoutObservation[];
  readonly contributions: CapabilityContribution[];
  readonly morphObjects: Map<MorphKey, MorphObject>;
  matchContinuityKinds: boolean;
  readonly linkedSlideIndexes: Set<number>;
  readonly layoutDerivedObjectIds: Set<number>;
};

type BoundCapabilityRegistry = {
  readonly declarations: ReadonlyMap<
    string,
    ObjectCapabilityModule | TransparentCapabilityModule
  >;
  readonly slideModules: ReadonlyMap<string, SlideCapabilityModule>;
};

type ResolvePassContext = {
  readonly document: DocumentResolveState;
  readonly slide: SlideResolveState;
  readonly modules: BoundCapabilityRegistry;
  readonly dimensions: ResolutionDimensions;
  readonly options: ResolveOptions;
  readonly theme: ThemeContext;
  readonly layout: LayoutPassResult;
  readonly slideNumber: number;
  readonly placeholderContexts: WeakMap<
    CapabilityElement,
    PlaceholderResolutionContext
  >;
};

type ResolvedLayerEntry = {
  readonly layer: ResolvedObjectLayer;
  readonly objects: readonly ObjectObservation[];
  readonly reference: LayoutChildReference;
};

type ResolvedChildrenResult = {
  readonly objects: readonly ObjectObservation[];
  readonly references: readonly LayoutChildReference[];
  readonly content: readonly QueryContentNode[];
  readonly layers: readonly ResolvedObjectLayer[];
  readonly entries: readonly ResolvedLayerEntry[];
};

type ResolvedSlideResult = {
  readonly internal: InternalResolvedSlide;
  readonly observation: SlideObservation;
  readonly morphObjects: ReadonlyMap<MorphKey, MorphObject>;
  readonly matchContinuityKinds: boolean;
  readonly linkedSlideIndexes: ReadonlySet<number>;
};

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== "object" || value === null) return value;
  // Capability observations might be shallow-frozen; still recursively freeze any unfrozen child values.
  if (!Object.isFrozen(value)) Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

function reportBox(box: GeometryBox): GeometryBox {
  // Reports show layout-derived geometry at hundredths of a pixel; capability measurement fallbacks keep their measured values.
  return {
    x: Math.round(box.x * 100) / 100,
    y: Math.round(box.y * 100) / 100,
    width: Math.round(box.width * 100) / 100,
    height: Math.round(box.height * 100) / 100,
  };
}

function projectObservation(
  observation: ObjectObservation,
  layoutDerivedObjectIds: ReadonlySet<number>,
): ObjectObservation {
  const box =
    observation.box !== undefined && layoutDerivedObjectIds.has(observation.id)
      ? reportBox(observation.box)
      : observation.box;
  const members =
    observation.members === undefined
      ? undefined
      : observation.members.map((member) =>
          projectObservation(member, layoutDerivedObjectIds),
        );
  return {
    ...observation,
    ...(box === undefined || box === observation.box ? {} : { box }),
    ...(members === undefined ? {} : { members }),
  };
}

function projectObservations(
  observations: readonly ObjectObservation[],
  layoutDerivedObjectIds: ReadonlySet<number>,
): readonly ObjectObservation[] {
  return observations.map((observation) =>
    projectObservation(observation, layoutDerivedObjectIds),
  );
}

function isObjectCapability(
  module: RegisteredCapabilityModule,
): module is ObjectCapabilityModule {
  return "bind" in module;
}

function nonWhitespaceText(children: readonly DeclarationChild[]): boolean {
  return children.some(
    (child) => !isDeclarationNode(child) && child.text.trim() !== "",
  );
}

function childNodes(children: readonly DeclarationChild[]): DeclarationNode[] {
  return children.filter(isDeclarationNode);
}

function capabilityChild(child: DeclarationChild): CapabilityChild {
  if (!isDeclarationNode(child)) return { type: "#text", text: child.text };
  return {
    type: child.type,
    props: child.props,
    children: child.children.map(capabilityChild),
    ...(child.pageAffiliated === true ? { pageAffiliated: true } : {}),
  };
}

function capabilityLabel(child: CapabilityChild, index: number): string {
  if (!("props" in child)) return `text[${index + 1}]`;
  const identity = declarationIdentity(child.props);
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${child.type}[${label}]`;
}

function isHiddenByStyle(node: DeclarationNode): boolean {
  const style = node.props["style"];
  return (
    typeof style === "object" &&
    style !== null &&
    !Array.isArray(style) &&
    "display" in style &&
    style.display === "none"
  );
}

function collectExplicitObjectNamesFromChildren(
  children: readonly DeclarationChild[],
  modules: ReadonlyMap<
    string,
    ObjectCapabilityModule | TransparentCapabilityModule
  >,
  names: Set<string>,
): void {
  for (const child of children) {
    if (!isDeclarationNode(child) || isHiddenByStyle(child)) continue;
    const module = modules.get(child.type);
    if (module !== undefined && "compileDeclaration" in module) {
      if (module.children === "owned") continue;
      collectExplicitObjectNamesFromChildren(child.children, modules, names);
      continue;
    }
    const rawName = child.props["name"];
    if (typeof rawName === "string" && rawName.trim() !== "") {
      names.add(rawName);
    }
    collectExplicitObjectNamesFromChildren(child.children, modules, names);
  }
}

function collectExplicitObjectNames(
  slides: readonly ParsedSlideNode[],
  modules: ReadonlyMap<
    string,
    ObjectCapabilityModule | TransparentCapabilityModule
  >,
  extraChildren: readonly DeclarationNode[] = [],
): ReadonlySet<string> {
  const names = new Set<string>();
  for (const slide of slides) {
    collectExplicitObjectNamesFromChildren(slide.node.children, modules, names);
  }
  collectExplicitObjectNamesFromChildren(extraChildren, modules, names);
  return names;
}

function nativeObjectName(
  key: MorphKey,
  document: DocumentResolveState,
): string {
  let name = document.morphNames.get(key);
  if (name === undefined) {
    // Skip any explicit object name authored anywhere in the deck so auto-generated Morph
    // handles never collide with or accidentally pair with manually named native objects.
    do {
      name = `!!morph-${document.nextMorphName++}`;
    } while (document.reservedObjectNames.has(name));
    document.morphNames.set(key, name);
  }
  return name;
}

function validateResolveOptions(options: ResolveOptions): void {
  const parsed = resolveOptionsSchema.safeParse(options);
  if (parsed.success) return;

  const issue = parsed.error.issues[0];
  if (issue === undefined) return;
  if (issue.path.length === 0 && issue.code === "unrecognized_keys") {
    throw new BuildError([
      diagnostic(
        "resolve",
        "unknown-parameter",
        `Resolve options contain unknown parameter ${issue.keys[0]}`,
      ),
    ]);
  }
  if (issue.path[0] === "signal") {
    throw new BuildError([
      diagnostic(
        "resolve.signal",
        "invalid-options",
        "Resolve options signal must be an AbortSignal",
      ),
    ]);
  }
  throw new BuildError([
    diagnostic(
      "resolve",
      "invalid-options",
      "Resolve options must be an object",
    ),
  ]);
}

// Deck dimensions and props validated via parseDeckProps(deck.props).

function unknownPropertyKey(
  props: Readonly<Record<string, unknown>>,
  allowed: readonly string[],
): string | undefined {
  const allowedSet = new Set(allowed);
  return Object.keys(props).find((key) => !allowedSet.has(key));
}

function resolveSlides(
  container: DeclarationContainer,
  modules: BoundCapabilityRegistry,
  defaultDeckTheme: Theme,
): ParsedDocument {
  const decks = childNodes(container.children);
  if (decks.length !== 1 || decks[0]?.type !== "deck") {
    throw new BuildError([
      diagnostic(
        "deck",
        "document-root",
        "Build content must have exactly one Deck root component",
      ),
    ]);
  }
  const deck = decks[0];
  if (deck === undefined)
    throw new BuildError([
      diagnostic("deck", "document-root", "Missing Deck root component"),
    ]);
  if (nonWhitespaceText(container.children)) {
    throw new BuildError([
      diagnostic(
        "deck",
        "unexpected-text",
        "No text allowed outside the Deck root component",
      ),
    ]);
  }
  if (nonWhitespaceText(deck.children)) {
    throw new BuildError([
      diagnostic(
        "deck",
        "unexpected-text",
        "Deck can only contain SlideMaster, Slide, and Section components",
      ),
    ]);
  }
  const topLevel = childNodes(deck.children);
  if (
    topLevel.length === 0 ||
    topLevel.length !== deck.children.filter(isDeclarationNode).length
  ) {
    throw new BuildError([
      diagnostic("deck", "missing-slide", "Deck requires at least one Slide"),
    ]);
  }

  const masterNodes = topLevel.filter((node) => node.type === "slide-master");
  const slideOrSectionNodes = topLevel.filter(
    (node) => node.type !== "slide-master",
  );

  const authorIds = new Set<string>();
  const slideProperties = [...modules.slideModules.values()].flatMap(
    (module) => module.slideProps ?? [module.slideProp],
  );
  const slidePropertySet = new Set(slideProperties);
  const knownTrimmedSlideProps = new Set([
    "transition",
    "advanceAfter",
    "motion",
  ]);
  const masters: ParsedMasterSpec[] = [];
  const layoutKeyByNode = new Map<DeclarationNode, string>();
  const layoutKeysByDefinition = new Map<SlideLayoutDefinition, string[]>();
  const layoutIdByKey = new Map<string, string>();
  let defaultBlankLayoutKey: string | undefined;
  let firstLayoutKey: string | undefined;
  let globalLayoutCounter = 0;

  for (const [masterIndex, masterNode] of masterNodes.entries()) {
    const rawMasterPath = `deck/slide-master[${masterIndex + 1}]`;
    const masterAuthorId = readAuthorId(masterNode.props["id"], rawMasterPath);
    const masterKey = `master-${masterIndex + 1}`;
    const masterId = masterAuthorId ?? masterKey;
    const masterPath = `deck/slide-master[${masterAuthorId ?? masterIndex + 1}]`;
    const unknownMasterProp = unknownPropertyKey(masterNode.props, [
      "id",
      "name",
      "preserve",
      "theme",
      "colorMap",
      "style",
      "background",
      "headerFooter",
      "footerPlaceholders",
      "placeholders",
      "textStyles",
      ...slideProperties,
    ]);
    if (unknownMasterProp !== undefined) {
      if (knownTrimmedSlideProps.has(unknownMasterProp)) {
        throw new BuildError([
          diagnostic(
            masterPath,
            "missing-capability",
            `No capability module registered for SlideMaster.${unknownMasterProp}`,
          ),
        ]);
      }
      throw new BuildError([
        diagnostic(
          masterPath,
          "unknown-parameter",
          `SlideMaster contains unknown parameter ${unknownMasterProp}`,
        ),
      ]);
    }
    const masterDataInput: Record<string, unknown> = {};
    for (const key of [
      "id",
      "name",
      "preserve",
      "colorMap",
      "background",
      "headerFooter",
      "footerPlaceholders",
      "placeholders",
      "textStyles",
    ] as const) {
      if (Object.hasOwn(masterNode.props, key)) {
        masterDataInput[key] = masterNode.props[key];
      }
    }
    const parsedMaster = slideMasterDataPropsSchema.safeParse(masterDataInput);
    if (!parsedMaster.success) {
      throw new BuildError([
        diagnostic(
          masterPath,
          "invalid-slide-master",
          parsedMaster.error.issues[0]?.message ??
            "Invalid SlideMaster declaration",
        ),
      ]);
    }
    let masterTheme: Theme;
    const rawMasterTheme = masterNode.props["theme"];
    if (rawMasterTheme === undefined) {
      masterTheme = defaultDeckTheme;
    } else if (isTheme(rawMasterTheme)) {
      masterTheme = rawMasterTheme;
    } else {
      try {
        masterTheme = defineTheme(rawMasterTheme as ThemeInput);
      } catch (error) {
        if (error instanceof CapabilityError) {
          throw new BuildError([
            diagnostic(masterPath, error.code, error.message),
          ]);
        }
        throw new BuildError([
          diagnostic(
            masterPath,
            "invalid-theme",
            error instanceof Error ? error.message : String(error),
          ),
        ]);
      }
    }
    if (masterAuthorId !== undefined) {
      if (!xml10StringSchema.safeParse(masterAuthorId).success) {
        throw new BuildError([
          diagnostic(
            masterPath,
            "invalid-xml-character",
            "SlideMaster.id contains characters not allowed in XML 1.0",
          ),
        ]);
      }
      if (authorIds.has(masterAuthorId)) {
        throw new BuildError([
          diagnostic(
            masterPath,
            "duplicate-author-id",
            `Duplicate author id ${JSON.stringify(masterAuthorId)} in the deck`,
          ),
        ]);
      }
      authorIds.add(masterAuthorId);
    }
    if (
      parsedMaster.data.name !== undefined &&
      !xml10StringSchema.safeParse(parsedMaster.data.name).success
    ) {
      throw new BuildError([
        diagnostic(
          masterPath,
          "invalid-xml-character",
          "SlideMaster.name contains characters not allowed in XML 1.0",
        ),
      ]);
    }
    for (const phSpec of Object.values(parsedMaster.data.placeholders ?? {})) {
      for (const phText of [phSpec.name, phSpec.promptText]) {
        if (phText === undefined) continue;
        if (!xml10StringSchema.safeParse(phText).success) {
          throw new BuildError([
            diagnostic(
              masterPath,
              "invalid-xml-character",
              "SlideMaster placeholder text contains characters not allowed in XML 1.0",
            ),
          ]);
        }
      }
    }
    if (nonWhitespaceText(masterNode.children)) {
      throw new BuildError([
        diagnostic(
          masterPath,
          "unexpected-text",
          "SlideMaster cannot directly contain text",
        ),
      ]);
    }
    const masterChildren = childNodes(masterNode.children);
    const layoutNodes: DeclarationNode[] = [];
    const masterCanvasNodes: DeclarationNode[] = [];
    for (const child of masterChildren) {
      if (child.type === "slide-layout") {
        layoutNodes.push(child);
      } else {
        masterCanvasNodes.push(child);
      }
    }
    if (layoutNodes.length === 0) {
      throw new BuildError([
        diagnostic(
          masterPath,
          "empty-slide-master",
          "SlideMaster requires at least one SlideLayout",
        ),
      ]);
    }

    const masterLayoutIds = new Set<string>();
    const resolvedLayouts: ParsedLayoutSpec[] = [];

    for (const [layoutIndex, layoutNode] of layoutNodes.entries()) {
      const rawLayoutPath = `${masterPath}/slide-layout[${layoutIndex + 1}]`;
      const rawLayoutDef = layoutNode.props["layout"];
      let layoutDef: SlideLayoutDefinition | undefined;
      if (rawLayoutDef !== undefined) {
        if (!isSlideLayoutDefinition(rawLayoutDef)) {
          throw new BuildError([
            diagnostic(
              rawLayoutPath,
              "invalid-slide-layout",
              "SlideLayout.layout must be created by defineSlideLayout()",
            ),
          ]);
        }
        layoutDef = rawLayoutDef;
      }
      const effectiveLayoutProps: Record<string, unknown> = {
        ...layoutDef?.props,
      };
      for (const [key, value] of Object.entries(layoutNode.props)) {
        if (key === "layout" || value === undefined) continue;
        effectiveLayoutProps[key] = value;
      }
      const unknownLayoutProp = unknownPropertyKey(effectiveLayoutProps, [
        "id",
        "name",
        "type",
        "preserve",
        "matchingName",
        "userDrawn",
        "showMasterShapes",
        "showMasterPlaceholderAnimations",
        "placeholders",
        "colorMapOverride",
        "themeOverride",
        "background",
        "headerFooter",
        "footerPlaceholders",
        ...slideProperties,
      ]);
      if (unknownLayoutProp !== undefined) {
        if (knownTrimmedSlideProps.has(unknownLayoutProp)) {
          throw new BuildError([
            diagnostic(
              rawLayoutPath,
              "missing-capability",
              `No capability module registered for SlideLayout.${unknownLayoutProp}`,
            ),
          ]);
        }
        throw new BuildError([
          diagnostic(
            rawLayoutPath,
            "unknown-parameter",
            `SlideLayout contains unknown parameter ${unknownLayoutProp}`,
          ),
        ]);
      }
      const layoutDataInput: Record<string, unknown> = {};
      for (const key of [
        "id",
        "name",
        "type",
        "preserve",
        "matchingName",
        "userDrawn",
        "showMasterShapes",
        "showMasterPlaceholderAnimations",
        "placeholders",
        "colorMapOverride",
        "themeOverride",
        "background",
        "headerFooter",
        "footerPlaceholders",
      ] as const) {
        if (Object.hasOwn(effectiveLayoutProps, key)) {
          layoutDataInput[key] = effectiveLayoutProps[key];
        }
      }
      const parsedLayout =
        slideLayoutDataPropsSchema.safeParse(layoutDataInput);
      if (!parsedLayout.success) {
        throw new BuildError([
          diagnostic(
            rawLayoutPath,
            "invalid-slide-layout",
            parsedLayout.error.issues[0]?.message ??
              "Invalid SlideLayout declaration",
          ),
        ]);
      }
      const layoutAuthorId = readAuthorId(parsedLayout.data.id, rawLayoutPath);
      const layoutId = layoutAuthorId ?? `layout-${layoutIndex + 1}`;
      const layoutPath = `${masterPath}/slide-layout[${layoutAuthorId ?? layoutIndex + 1}]`;
      if (layoutAuthorId !== undefined) {
        if (masterLayoutIds.has(layoutAuthorId)) {
          throw new BuildError([
            diagnostic(
              layoutPath,
              "duplicate-author-id",
              `Duplicate SlideLayout.id ${JSON.stringify(layoutAuthorId)} in ${masterPath}`,
            ),
          ]);
        }
        masterLayoutIds.add(layoutAuthorId);
      }
      for (const textField of [
        parsedLayout.data.id,
        parsedLayout.data.name,
        parsedLayout.data.matchingName,
      ]) {
        if (textField === undefined) continue;
        if (!xml10StringSchema.safeParse(textField).success) {
          throw new BuildError([
            diagnostic(
              layoutPath,
              "invalid-xml-character",
              "SlideLayout text property contains characters not allowed in XML 1.0",
            ),
          ]);
        }
      }
      for (const phSpec of Object.values(
        parsedLayout.data.placeholders ?? {},
      )) {
        for (const phText of [phSpec.name, phSpec.promptText]) {
          if (phText === undefined) continue;
          if (!xml10StringSchema.safeParse(phText).success) {
            throw new BuildError([
              diagnostic(
                layoutPath,
                "invalid-xml-character",
                "SlideLayout placeholder text contains characters not allowed in XML 1.0",
              ),
            ]);
          }
        }
      }
      validateAuthorIdScopes(layoutNode, layoutPath, new Set());
      if (nonWhitespaceText(layoutNode.children)) {
        throw new BuildError([
          diagnostic(
            layoutPath,
            "unexpected-text",
            "SlideLayout cannot directly contain text",
          ),
        ]);
      }

      const layoutKey = `layout-${++globalLayoutCounter}`;
      layoutKeyByNode.set(layoutNode, layoutKey);
      if (layoutDef !== undefined) {
        const existing = layoutKeysByDefinition.get(layoutDef) ?? [];
        existing.push(layoutKey);
        layoutKeysByDefinition.set(layoutDef, existing);
      }
      layoutIdByKey.set(layoutKey, layoutId);
      firstLayoutKey ??= layoutKey;
      if (
        parsedLayout.data.type === "blank" &&
        defaultBlankLayoutKey === undefined
      ) {
        defaultBlankLayoutKey = layoutKey;
      }

      const layoutData: ParsedLayoutSpec["data"] = {
        id: layoutId,
        ...(parsedLayout.data.name === undefined
          ? {}
          : { name: parsedLayout.data.name }),
        ...(parsedLayout.data.type === undefined
          ? {}
          : { type: parsedLayout.data.type }),
        ...(parsedLayout.data.preserve === undefined
          ? {}
          : { preserve: parsedLayout.data.preserve }),
        ...(parsedLayout.data.matchingName === undefined
          ? {}
          : { matchingName: parsedLayout.data.matchingName }),
        ...(parsedLayout.data.userDrawn === undefined
          ? {}
          : { userDrawn: parsedLayout.data.userDrawn }),
        ...(parsedLayout.data.showMasterShapes === undefined
          ? {}
          : { showMasterShapes: parsedLayout.data.showMasterShapes }),
        ...(parsedLayout.data.showMasterPlaceholderAnimations === undefined
          ? {}
          : {
              showMasterPlaceholderAnimations:
                parsedLayout.data.showMasterPlaceholderAnimations,
            }),
        ...(parsedLayout.data.colorMapOverride === undefined
          ? {}
          : { colorMapOverride: parsedLayout.data.colorMapOverride }),
        ...(parsedLayout.data.themeOverride === undefined
          ? {}
          : { themeOverride: parsedLayout.data.themeOverride }),
        ...(parsedLayout.data.headerFooter === undefined
          ? {}
          : { headerFooter: parsedLayout.data.headerFooter }),
      };
      const layoutSlideProps: Record<string, unknown> =
        parsedLayout.data.placeholders === undefined
          ? {}
          : { placeholders: parsedLayout.data.placeholders };
      for (const prop of slidePropertySet) {
        if (prop === "background") continue;
        if (Object.hasOwn(effectiveLayoutProps, prop)) {
          layoutSlideProps[prop] = effectiveLayoutProps[prop];
        }
      }
      resolvedLayouts.push({
        node: layoutNode,
        key: layoutKey,
        path: layoutPath,
        data: layoutData,
        ...(parsedLayout.data.placeholders === undefined
          ? {}
          : { placeholders: parsedLayout.data.placeholders }),
        ...(parsedLayout.data.background === undefined
          ? {}
          : { background: parsedLayout.data.background }),
        ...(parsedLayout.data.headerFooter === undefined
          ? {}
          : { headerFooter: parsedLayout.data.headerFooter }),
        ...(parsedLayout.data.footerPlaceholders === undefined
          ? {}
          : { footerPlaceholders: parsedLayout.data.footerPlaceholders }),
        props: layoutSlideProps,
      });
    }

    if (masterCanvasNodes.length > 0) {
      validateAuthorIdScopes(
        {
          type: "slide-master",
          props: masterNode.props,
          children: masterCanvasNodes,
        },
        masterPath,
        new Set(masterLayoutIds),
      );
    }

    const masterSlideProps: Record<string, unknown> =
      parsedMaster.data.placeholders === undefined
        ? {}
        : { placeholders: parsedMaster.data.placeholders };
    for (const prop of slidePropertySet) {
      if (prop === "background") continue;
      if (Object.hasOwn(masterNode.props, prop)) {
        masterSlideProps[prop] = masterNode.props[prop];
      }
    }
    const masterStyleProp = masterNode.props["style"];
    masters.push({
      node: masterNode,
      canvasChildren: masterCanvasNodes,
      key: masterKey,
      id: masterId,
      path: masterPath,
      ...(parsedMaster.data.name === undefined
        ? {}
        : { name: parsedMaster.data.name }),
      ...(parsedMaster.data.preserve === undefined
        ? {}
        : { preserve: parsedMaster.data.preserve }),
      theme: masterTheme,
      colorMap: parsedMaster.data.colorMap ?? defaultColorMap,
      ...(parsedMaster.data.colorMap === undefined
        ? {}
        : { explicitColorMap: parsedMaster.data.colorMap }),
      ...(masterStyleProp !== undefined &&
      typeof masterStyleProp === "object" &&
      masterStyleProp !== null &&
      !Array.isArray(masterStyleProp)
        ? { style: masterStyleProp as Readonly<Record<string, unknown>> }
        : {}),
      ...(parsedMaster.data.placeholders === undefined
        ? {}
        : { placeholders: parsedMaster.data.placeholders }),
      ...(parsedMaster.data.textStyles === undefined
        ? {}
        : { textStyles: parsedMaster.data.textStyles }),
      ...(parsedMaster.data.background === undefined
        ? {}
        : { background: parsedMaster.data.background }),
      ...(parsedMaster.data.headerFooter === undefined
        ? {}
        : { headerFooter: parsedMaster.data.headerFooter }),
      ...(parsedMaster.data.footerPlaceholders === undefined
        ? {}
        : { footerPlaceholders: parsedMaster.data.footerPlaceholders }),
      props: masterSlideProps,
      layouts: resolvedLayouts,
    });
  }

  const sectioned = slideOrSectionNodes.some((node) => node.type === "section");
  const candidates: {
    node: DeclarationNode;
    section?: string;
    sectionIndex?: number;
  }[] = [];
  const sections: { name: string; slideIndexes: number[] }[] = [];
  if (sectioned) {
    for (const [index, node] of slideOrSectionNodes.entries()) {
      if (node.type !== "section") {
        throw new BuildError([
          diagnostic(
            `deck/${componentLabel(node, index)}`,
            "mixed-sections",
            "When using Section, every slide must belong to a Section",
          ),
        ]);
      }
      const { name } = parseSectionProps(node.props, index);
      if (nonWhitespaceText(node.children)) {
        throw new BuildError([
          diagnostic(
            `deck/section[${name}]`,
            "unexpected-section-text",
            "Section can only contain Slide components",
          ),
        ]);
      }
      const sectionIndex = sections.length;
      const section = { name, slideIndexes: [] as number[] };
      sections.push(section);
      for (const [memberIndex, member] of childNodes(node.children).entries()) {
        if (member.type !== "slide") {
          throw new BuildError([
            diagnostic(
              `deck/section[${name}]/${componentLabel(member, memberIndex)}`,
              "nested-section",
              "Section can only contain Slide components and cannot nest Section components",
            ),
          ]);
        }
        candidates.push({ node: member, section: name, sectionIndex });
      }
    }
  } else {
    for (const [index, node] of slideOrSectionNodes.entries()) {
      if (node.type !== "slide") {
        throw new BuildError([
          diagnostic(
            `deck/${componentLabel(node, index)}`,
            "unexpected-declaration",
            "Direct children of Deck must be SlideMaster, Slide, or Section",
          ),
        ]);
      }
      candidates.push({ node });
    }
  }
  if (candidates.length === 0 && masters.length === 0) {
    throw new BuildError([
      diagnostic("deck", "missing-slide", "Deck requires at least one Slide"),
    ]);
  }
  const slides = candidates.map(({ node, section, sectionIndex }, index) => {
    const {
      id,
      hidden,
      showMasterPlaceholderAnimations,
      showMasterShapes,
      colorMapOverride,
      themeOverride,
    } = parseSlideCoreProps(node.props, index, slideProperties);
    const authorId = id;
    if (authorIds.has(authorId)) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${authorId}]`,
          "duplicate-author-id",
          `Duplicate Slide.id ${JSON.stringify(authorId)} in the deck`,
        ),
      ]);
    }
    authorIds.add(authorId);
    validateAuthorIdScopes(node, `deck/slide[${id}]`, new Set());

    let layoutKey: string | undefined;
    const rawLayout = node.props["layout"];
    if (rawLayout === undefined) {
      if (masters.length > 0) {
        layoutKey = defaultBlankLayoutKey ?? firstLayoutKey;
      }
    } else if (isSlideLayoutDefinition(rawLayout)) {
      const keys = layoutKeysByDefinition.get(rawLayout);
      const targetLabel = rawLayout.id ?? rawLayout.name ?? "<unnamed>";
      if (keys === undefined || keys.length === 0) {
        throw new BuildError([
          diagnostic(
            `deck/slide[${id}]`,
            "foreign-slide-layout",
            `Slide ${JSON.stringify(id)} references slide layout ${JSON.stringify(targetLabel)}, which is not owned by this deck`,
          ),
        ]);
      }
      if (keys.length > 1) {
        throw new BuildError([
          diagnostic(
            `deck/slide[${id}]`,
            "ambiguous-slide-layout",
            `Slide ${JSON.stringify(id)} references slide layout ${JSON.stringify(targetLabel)} mounted multiple times in this deck`,
          ),
        ]);
      }
      layoutKey = keys[0];
    } else if (isDeclarationNode(rawLayout)) {
      if (rawLayout.type !== "slide-layout") {
        throw new BuildError([
          diagnostic(
            `deck/slide[${id}]`,
            "invalid-slide-layout",
            "Slide.layout must be a SlideLayout or defineSlideLayout specification",
          ),
        ]);
      }
      const matchedKey = layoutKeyByNode.get(rawLayout);
      if (matchedKey === undefined) {
        const rawTarget = rawLayout.props["id"] ?? rawLayout.props["name"];
        const targetLabel =
          typeof rawTarget === "string" ? rawTarget : "<unnamed>";
        throw new BuildError([
          diagnostic(
            `deck/slide[${id}]`,
            "foreign-slide-layout",
            `Slide ${JSON.stringify(id)} references slide layout ${JSON.stringify(targetLabel)}, which is not owned by this deck`,
          ),
        ]);
      }
      layoutKey = matchedKey;
    } else {
      throw new BuildError([
        diagnostic(
          `deck/slide[${id}]`,
          "invalid-slide-layout",
          "Slide.layout must be a SlideLayout or defineSlideLayout specification",
        ),
      ]);
    }

    const resolved: ParsedSlideNode = {
      node,
      id,
      ...(authorId === undefined ? {} : { authorId }),
      index: index + 1,
      ...(layoutKey === undefined ? {} : { layoutKey }),
      hasExplicitLayout: rawLayout !== undefined,
      ...(colorMapOverride === undefined ? {} : { colorMapOverride }),
      ...(themeOverride === undefined ? {} : { themeOverride }),
      ...(section === undefined ? {} : { section }),
      ...(hidden === undefined ? {} : { hidden }),
      ...(showMasterPlaceholderAnimations === undefined
        ? {}
        : { showMasterPlaceholderAnimations }),
      ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
    };
    if (sectionIndex !== undefined) {
      sections[sectionIndex]?.slideIndexes.push(index + 1);
    }
    return resolved;
  });
  return {
    masters,
    layoutIdByKey,
    slides,
    sections: sections.map(({ name, slideIndexes }) => ({
      name,
      slideIndexes,
    })),
  };
}

function validateLink(
  value: unknown,
  context: ResolvePassContext,
  path: string,
): ResolvedLink | undefined {
  let parsed: ReturnType<typeof parseLinkTarget>;
  try {
    parsed = parseLinkTarget(value, path);
  } catch (error) {
    if (error instanceof CapabilityError) {
      context.slide.diagnostics.push(
        diagnostic(path, error.code, error.message),
      );
      return undefined;
    }
    throw error;
  }
  if (parsed === undefined) return undefined;
  if ("id" in parsed) {
    const targetIndex = context.document.slideIndexById.get(parsed.id);
    if (targetIndex === undefined) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-link-target",
          `No target slide with author id ${JSON.stringify(parsed.id)}`,
        ),
      );
      return undefined;
    }
    if (targetIndex !== context.slide.index) {
      context.slide.linkedSlideIndexes.add(targetIndex);
    }
    return { target: { id: parsed.id }, summary: `slide:${parsed.id}` };
  }
  if ("url" in parsed) {
    let url: URL;
    try {
      url = new URL(parsed.url);
    } catch {
      context.slide.diagnostics.push(
        diagnostic(path, "invalid-link", "link.url must be an absolute URL"),
      );
      return undefined;
    }
    return { target: { url: url.href }, summary: `url:${url.href}` };
  }
  return { target: { action: "home" }, summary: "home" };
}

function recordResourceClaims(
  claims: readonly CapabilityResourceClaim[] | undefined,
  document: DocumentResolveState,
  diagnostics: Diagnostic[],
  path: string,
): void {
  for (const claim of claims ?? []) {
    const claimKey = `${claim.kind}\u0000${claim.key}`;
    const existing = document.resourceClaims.get(claimKey);
    if (existing !== undefined && existing.identity !== claim.identity) {
      diagnostics.push(
        diagnostic(
          path,
          "conflicting-resource",
          `Resource ${claim.key} has conflicting definition with ${existing.path}`,
        ),
      );
    } else {
      document.resourceClaims.set(claimKey, {
        identity: claim.identity,
        path: existing?.path ?? path,
        characters: (existing?.characters ?? "") + (claim.characters ?? ""),
      });
    }
  }
}

function recordCapabilityWarnings(
  warnings: readonly CapabilityWarning[] | undefined,
  document: DocumentResolveState,
  path: string,
): void {
  for (const warning of warnings ?? []) {
    document.warnings.push({
      severity: "warning",
      path: warning.path === undefined ? path : `${path}/${warning.path}`,
      code: warning.code,
      message: warning.message,
    });
  }
}

function checkCanvasOverflow(
  box: GeometryBox | undefined,
  dimensions: ResolutionDimensions,
  document: DocumentResolveState,
  path: string,
): void {
  if (box === undefined) return;
  if (
    box.x < 0 ||
    box.y < 0 ||
    box.x + box.width > dimensions.width ||
    box.y + box.height > dimensions.height
  ) {
    document.warnings.push({
      severity: "warning",
      path,
      code: "overflow",
      message:
        "Object overflows slide canvas; build retains original declaration",
    });
  }
}

function recordObjectMorph(args: {
  readonly morphKey: MorphKey | undefined;
  readonly morphName: string | undefined;
  readonly nativeKind: NativeElementKind;
  readonly observationName: string;
  readonly hasOutput: boolean;
  readonly slide: SlideResolveState;
  readonly declarationType: string;
  readonly path: string;
}): void {
  const {
    morphKey,
    morphName,
    nativeKind,
    observationName,
    hasOutput,
    slide,
    declarationType,
    path,
  } = args;
  if (morphKey === undefined || morphName === undefined) return;
  if (!hasOutput) {
    slide.diagnostics.push(
      diagnostic(
        path,
        "morph-needs-object",
        "morphKey can only be used with capabilities that produce native objects",
      ),
    );
  } else if (observationName !== morphName) {
    slide.diagnostics.push(
      diagnostic(
        path,
        "unsupported-morph-object",
        `${declarationType} did not write morphKey to native object name, cannot participate in Morph matching`,
      ),
    );
  } else if (!slide.morphObjects.has(morphKey)) {
    slide.morphObjects.set(morphKey, {
      kind: nativeKind,
      name: morphName,
      path,
    });
  }
}

function createObjectResolveContext(args: {
  readonly resolveContext: ResolvePassContext;
  readonly id: number;
  readonly path: string;
  readonly layoutBox: GeometryBox | undefined;
  readonly theme: ThemeContext;
  readonly placeholder: PlaceholderResolutionContext | undefined;
  readonly insideGroup: boolean;
  readonly restrictsMemberContinuity: boolean;
  readonly morphName: string | undefined;
  readonly link: ResolvedLink | undefined;
  readonly onChildrenResolved?: (resolved: ResolvedChildrenResult) => void;
}): ObjectResolveContext {
  const {
    resolveContext,
    id,
    path,
    layoutBox,
    theme,
    placeholder,
    insideGroup,
    restrictsMemberContinuity,
    morphName,
    link,
    onChildrenResolved,
  } = args;
  return {
    id,
    path,
    theme,
    options: resolveContext.options,
    slideNumber: resolveContext.slideNumber,
    ...(layoutBox === undefined ? {} : { layoutBox }),
    ...(placeholder === undefined ? {} : { placeholder }),
    ...(morphName === undefined ? {} : { morphName }),
    ...(link === undefined ? {} : { link }),
    resolveChildren(nested): ResolvedObjectChildren {
      const resolved = resolveChildrenPass(
        nested,
        resolveContext,
        path,
        insideGroup || restrictsMemberContinuity,
      );
      onChildrenResolved?.(resolved);
      return { objects: resolved.objects };
    },
  };
}

function resolveChildrenPass(
  children: readonly CapabilityChild[],
  context: ResolvePassContext,
  parentPath: string,
  insideGroup = false,
): ResolvedChildrenResult {
  const layers: {
    readonly layer: ResolvedObjectLayer;
    readonly objects: readonly ObjectObservation[];
    readonly reference: LayoutChildReference;
  }[] = [];
  const content: QueryContentNode[] = [];

  children.forEach((child, index) => {
    if (!("props" in child)) {
      if (child.text.trim() !== "") {
        context.slide.diagnostics.push(
          diagnostic(
            `${parentPath}/${capabilityLabel(child, index)}`,
            "unexpected-text",
            "Native object container cannot directly contain text",
          ),
        );
      }
      return;
    }

    if (context.layout.hidden.has(child)) return;
    const path = `${parentPath}/${capabilityLabel(child, index)}`;
    const authorId = readAuthorId(child.props["id"], path);
    const order = context.slide.nextNodeOrder++;
    const module = context.modules.declarations.get(child.type);
    if (module === undefined) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-capability",
          `No capability module registered for ${child.type}`,
        ),
      );
      return;
    }
    if ("compileDeclaration" in module) {
      if (module.children === "owned") {
        // Owned subtrees belong to their capability and never become slide content. The same
        // declarations were already reported by the structure-only collection pass.
        return;
      }
      try {
        for (const annotation of module.compileDeclaration(child, {
          path,
          theme: context.theme,
          slideNumber: context.slideNumber,
        })) {
          context.slide.contributions.push({
            ...annotation,
            path,
            order,
            declaration: child,
          });
        }
      } catch (error) {
        context.slide.diagnostics.push(
          error instanceof CapabilityError
            ? diagnostic(
                error.path === undefined ? path : `${path}/${error.path}`,
                error.code,
                error.message,
              )
            : diagnostic(
                path,
                "capability-failure",
                error instanceof Error ? error.message : String(error),
              ),
        );
      }
      const childResolved = resolveChildrenPass(
        child.children,
        context,
        path,
        insideGroup,
      );
      content.push(...childResolved.content);
      for (const entry of childResolved.entries) {
        layers.push({
          ...entry,
          layer: {
            ...entry.layer,
            sourceOrder: [index, ...entry.layer.sourceOrder],
          },
        });
      }
      return;
    }
    if (!isObjectCapability(module)) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-capability",
          `No capability module registered for ${child.type}`,
        ),
      );
      return;
    }
    const layoutBox = context.layout.boxes.get(child);
    const placeholder = context.placeholderContexts.get(child);
    const theme =
      placeholder !== undefined
        ? placeholder.masterTheme
        : (context.layout.themes.get(child) ?? context.theme);
    try {
      const bound =
        context.layout.bound.get(child) ??
        validateBoundObjectCapability(
          module.bind(child, {
            path,
            theme,
            options: context.options,
            slideNumber: context.slideNumber,
            ...(placeholder === undefined ? {} : { placeholder }),
          }),
        );
      if (bound.container !== undefined) {
        resolveBoundContainer({
          child,
          index,
          path,
          authorId,
          order,
          bound,
          container: bound.container,
          theme,
          context,
          insideGroup,
          layers,
          content,
        });
        return;
      }
      const id = context.slide.nextId++;
      const requestedMorphKey = bound.native?.morphKey;
      const hasExplicitName = bound.native?.hasExplicitName === true;
      const emitsObject = bound.native?.emitsObject !== false;
      let morphKey: MorphKey | undefined;
      let morphName: string | undefined;
      if (requestedMorphKey !== undefined) {
        if (!isMorphKey(requestedMorphKey)) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "invalid-morph-key",
              "morphKey must be created by createMorphScope().handle(name)",
            ),
          );
        } else if (insideGroup) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "morph-inside-group",
              "Morph only matches top-level slide objects; place morphKey on Group to move the entire group",
            ),
          );
        } else if (!emitsObject) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "morph-needs-object",
              "morphKey can only be used with capabilities that produce native objects",
            ),
          );
        } else if (hasExplicitName) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "morph-object-name-conflict",
              "Both morphKey and name claim native object names; specify only one",
            ),
          );
        } else {
          morphKey = requestedMorphKey;
          morphName = nativeObjectName(morphKey, context.document);
          const previous = context.slide.morphObjects.get(morphKey);
          if (previous !== undefined) {
            context.slide.diagnostics.push(
              diagnostic(
                path,
                "duplicate-morph-key",
                `Morph handle "${getMorphKeyLabel(morphKey)}" reused on this slide; can only be declared once per slide`,
              ),
            );
          }
        }
      }
      const rawLink = bound.native?.link;
      const resolvedLink =
        rawLink === undefined
          ? undefined
          : emitsObject
            ? validateLink(rawLink, context, path)
            : (context.slide.diagnostics.push(
                diagnostic(
                  path,
                  "link-needs-object",
                  "Only elements producing native objects can attach click links",
                ),
              ),
              undefined);

      const resolvedObj = bound.resolve(
        createObjectResolveContext({
          resolveContext: context,
          id,
          path,
          layoutBox,
          theme,
          placeholder,
          insideGroup,
          restrictsMemberContinuity: bound.contract.nativeKind === "grpSp",
          morphName,
          link: resolvedLink,
        }),
      );
      recordObjectMorph({
        morphKey,
        morphName,
        nativeKind: bound.contract.nativeKind,
        observationName: resolvedObj.observation.name,
        hasOutput: emitsObject,
        slide: context.slide,
        declarationType: child.type,
        path,
      });
      for (const extension of resolvedObj.extensions) {
        context.document.usedExtensions.add(extension);
      }
      recordResourceClaims(
        resolvedObj.resourceClaims,
        context.document,
        context.slide.diagnostics,
        path,
      );
      recordCapabilityWarnings(resolvedObj.warnings, context.document, path);
      for (const annotation of resolvedObj.annotations ?? []) {
        context.slide.contributions.push({
          ...annotation,
          path,
          targetId: id,
          order,
          declaration: child,
        });
      }
      const observation: ObjectObservation = {
        ...resolvedObj.observation,
        id,
        ...(authorId === undefined ? {} : { authorId }),
        ...(placeholder?.key === undefined
          ? {}
          : resolvedObj.observation.placeholder === undefined
            ? { placeholder: placeholder.key }
            : {}),
        ...(resolvedLink === undefined ? {} : { link: resolvedLink.summary }),
      };
      if (layoutBox !== undefined) context.slide.layoutDerivedObjectIds.add(id);
      content.push({
        ...(authorId === undefined ? {} : { authorId }),
        reference: { kind: "object", id: observation.id },
        children: subPartContentNodes(resolvedObj.parts),
      });
      checkCanvasOverflow(
        resolvedObj.observation.box,
        context.dimensions,
        context.document,
        path,
      );
      const layer: ResolvedObjectLayer = {
        declarationType: child.type,
        sourceOrder: [index],
        zIndex: paintOrder(bound.contract),
        path,
        objectId: observation.id,
        paintObjectId: observation.id,
        theme,
        ...(resolvedObj.observation.box === undefined
          ? {}
          : { box: resolvedObj.observation.box }),
        ...(resolvedObj.transform === undefined
          ? {}
          : { transform: resolvedObj.transform }),
        emitsObject,
        hasObservation: true,
        childComposition: "leaf",
        ...(resolvedLink === undefined ? {} : { resolvedLink }),
        ...(morphKey === undefined ? {} : { morphKey }),
        extensions: resolvedObj.extensions,
        memberLayers: [],
        siblingLayers: [],
        emit: (emitContext) => resolvedObj.emit(emitContext),
        ...(resolvedObj.paint === undefined
          ? {}
          : { paint: (paintContext) => resolvedObj.paint!(paintContext) }),
      };
      layers.push({
        layer,
        objects: [observation],
        reference: { kind: "object", id: observation.id },
      });
    } catch (error) {
      context.slide.diagnostics.push(
        error instanceof CapabilityError
          ? diagnostic(
              error.path === undefined ? path : `${path}/${error.path}`,
              error.code,
              error.message,
            )
          : diagnostic(
              path,
              "capability-failure",
              error instanceof Error ? error.message : String(error),
            ),
      );
    }
  });

  layers.sort(
    (left, right) =>
      left.layer.zIndex - right.layer.zIndex ||
      compareSourceOrder(left.layer.sourceOrder, right.layer.sourceOrder),
  );
  return {
    objects: layers.flatMap(({ objects }) => objects),
    references: layers.map(({ reference }) => reference),
    content,
    layers: layers.map(({ layer }) => layer),
    entries: layers,
  };
}

function resolveBoundContainer(args: {
  readonly child: CapabilityElement;
  readonly index: number;
  readonly path: string;
  readonly authorId: string | undefined;
  readonly order: number;
  readonly bound: BoundObjectCapability;
  readonly container: BoundContainerCollaboration;
  readonly theme: ThemeContext;
  readonly context: ResolvePassContext;
  readonly insideGroup: boolean;
  readonly layers: {
    readonly layer: ResolvedObjectLayer;
    readonly objects: readonly ObjectObservation[];
    readonly reference: LayoutChildReference;
  }[];
  readonly content: QueryContentNode[];
}): void {
  const {
    child,
    index,
    path,
    authorId,
    order,
    bound,
    container,
    theme,
    context,
    insideGroup,
    layers,
    content,
  } = args;
  const box = context.layout.boxes.get(child);
  if (box === undefined && container.layoutMode !== "when-styled") {
    context.slide.diagnostics.push(
      diagnostic(
        path,
        "layout-unavailable",
        "Container is missing layout result",
      ),
    );
    return;
  }
  const enclosesChildren = container.childComposition === "members";
  const restrictsMemberContinuity =
    container.restrictsMemberContinuity === true ||
    bound.contract.nativeKind === "grpSp";
  const observesLayout =
    container.observesLayout !== false && box !== undefined;
  const layoutId = observesLayout ? context.slide.nextId++ : undefined;
  const emitsObject = container.emitsObject;
  const requestedMorphKey = bound.native?.morphKey;
  const hasExplicitName = bound.native?.hasExplicitName === true;
  const rawLink = bound.native?.link;
  const containerLink =
    rawLink === undefined
      ? undefined
      : emitsObject
        ? validateLink(rawLink, context, path)
        : (context.slide.diagnostics.push(
            diagnostic(
              path,
              "link-needs-object",
              "Transparent container does not produce clickable objects; attach link to a visible container or object",
            ),
          ),
          undefined);
  let morphKey: MorphKey | undefined;
  let morphName: string | undefined;
  if (requestedMorphKey !== undefined) {
    if (!isMorphKey(requestedMorphKey)) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "invalid-morph-key",
          "morphKey must be created by createMorphScope().handle(name)",
        ),
      );
    } else if (insideGroup) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "morph-inside-group",
          "Morph only matches top-level slide objects; place morphKey on Group to move the entire group",
        ),
      );
    } else if (!emitsObject) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "morph-needs-object",
          "Only containers generating native objects can participate in Morph",
        ),
      );
    } else if (hasExplicitName) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "morph-object-name-conflict",
          "Both morphKey and name claim native object names; specify only one",
        ),
      );
    } else {
      morphKey = requestedMorphKey;
      morphName = nativeObjectName(morphKey, context.document);
      if (context.slide.morphObjects.has(morphKey)) {
        context.slide.diagnostics.push(
          diagnostic(
            path,
            "duplicate-morph-key",
            `Morph handle "${getMorphKeyLabel(morphKey)}" reused on this slide; can only be declared once per slide`,
          ),
        );
      }
    }
  }
  let containerObservation: ObjectObservation | undefined;
  let memberResolved: ResolvedChildrenResult | undefined;
  let memberLayers: readonly ResolvedObjectLayer[] = [];
  let containerLayer: ResolvedObjectLayer | undefined;
  const objectId = context.slide.nextId++;

  try {
    const resolvedObj = bound.resolve(
      createObjectResolveContext({
        resolveContext: context,
        id: objectId,
        path,
        layoutBox: box,
        theme,
        placeholder: undefined,
        insideGroup,
        restrictsMemberContinuity,
        morphName,
        link: containerLink,
        onChildrenResolved: (resolved) => {
          memberResolved = resolved;
        },
      }),
    );
    if (emitsObject) {
      containerObservation = {
        ...resolvedObj.observation,
        id: objectId,
        ...(authorId === undefined ? {} : { authorId }),
        ...(containerLink === undefined ? {} : { link: containerLink.summary }),
      };
      if (box !== undefined) context.slide.layoutDerivedObjectIds.add(objectId);
      if (container.checksCanvasOverflow === false) {
        checkCanvasOverflow(
          resolvedObj.observation.box,
          context.dimensions,
          context.document,
          path,
        );
      }
      recordObjectMorph({
        morphKey,
        morphName,
        nativeKind: bound.contract.nativeKind,
        observationName: resolvedObj.observation.name,
        hasOutput: true,
        slide: context.slide,
        declarationType: child.type,
        path,
      });
      for (const extension of resolvedObj.extensions) {
        context.document.usedExtensions.add(extension);
      }
      recordResourceClaims(
        resolvedObj.resourceClaims,
        context.document,
        context.slide.diagnostics,
        path,
      );
    } else if (morphKey !== undefined) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "morph-needs-object",
          "Only containers generating native objects can participate in Morph",
        ),
      );
    }
    for (const annotation of resolvedObj.annotations ?? []) {
      context.slide.contributions.push({
        ...annotation,
        path,
        ...(containerObservation === undefined
          ? {}
          : { targetId: containerObservation.id }),
        order,
        declaration: child,
      });
    }
    recordCapabilityWarnings(resolvedObj.warnings, context.document, path);
    memberLayers = enclosesChildren ? (memberResolved?.layers ?? []) : [];
    containerLayer = {
      declarationType: child.type,
      sourceOrder: [index],
      zIndex: paintOrder(bound.contract),
      path,
      objectId: containerObservation?.id ?? objectId,
      paintObjectId: layoutId ?? objectId,
      theme,
      ...(containerObservation?.box === undefined
        ? {}
        : { box: containerObservation.box }),
      ...(resolvedObj.transform === undefined
        ? {}
        : { transform: resolvedObj.transform }),
      emitsObject,
      hasObservation: containerObservation !== undefined,
      childComposition: enclosesChildren ? "members" : "siblings",
      ...(containerLink === undefined ? {} : { resolvedLink: containerLink }),
      ...(morphKey === undefined ? {} : { morphKey }),
      extensions: resolvedObj.extensions,
      memberLayers,
      siblingLayers: [],
      emit: (emitContext) => resolvedObj.emit(emitContext),
      ...(resolvedObj.paint === undefined
        ? {}
        : { paint: (paintContext) => resolvedObj.paint!(paintContext) }),
    };
  } catch (error) {
    context.slide.diagnostics.push(
      error instanceof CapabilityError
        ? diagnostic(
            error.path === undefined ? path : `${path}/${error.path}`,
            error.code,
            error.message,
          )
        : diagnostic(
            path,
            "capability-failure",
            error instanceof Error ? error.message : String(error),
          ),
    );
  }

  const resolvedChildren = enclosesChildren
    ? (memberResolved ?? {
        objects: [],
        references: [],
        content: [],
        layers: memberLayers,
        entries: [],
      })
    : resolveChildrenPass(
        container.children,
        context,
        path,
        insideGroup || restrictsMemberContinuity,
      );

  const finalLayer: ResolvedObjectLayer =
    containerLayer !== undefined
      ? {
          ...containerLayer,
          ...(!enclosesChildren
            ? { siblingLayers: resolvedChildren.layers }
            : {}),
        }
      : {
          declarationType: child.type,
          sourceOrder: [index],
          zIndex: paintOrder(bound.contract),
          path,
          objectId,
          paintObjectId: layoutId ?? objectId,
          theme,
          emitsObject,
          hasObservation: false,
          childComposition: enclosesChildren ? "members" : "siblings",
          extensions: [],
          memberLayers: enclosesChildren ? resolvedChildren.layers : [],
          siblingLayers: enclosesChildren ? [] : resolvedChildren.layers,
          emit: () => ({ xml: "" }),
        };

  const layerObservedObjects = [
    ...(containerObservation === undefined ? [] : [containerObservation]),
    ...(enclosesChildren ? [] : resolvedChildren.objects),
  ];

  if (layoutId !== undefined && box !== undefined) {
    context.slide.layouts.push({
      id: layoutId,
      ...(authorId === undefined ? {} : { authorId }),
      box: reportBox(box),
      children: [
        ...(containerObservation === undefined
          ? []
          : [{ kind: "object" as const, id: containerObservation.id }]),
        ...resolvedChildren.references,
      ],
    });
    content.push({
      ...(authorId === undefined ? {} : { authorId }),
      reference: {
        kind: "layout",
        id: layoutId,
        objectIds:
          containerObservation === undefined ? [] : [containerObservation.id],
      },
      children: resolvedChildren.content,
    });
    layers.push({
      layer: finalLayer,
      objects: layerObservedObjects,
      reference: { kind: "layout", id: layoutId },
    });
    return;
  }
  if (containerObservation !== undefined) {
    content.push({
      ...(authorId === undefined ? {} : { authorId }),
      reference: { kind: "object", id: containerObservation.id },
      children: resolvedChildren.content,
    });
  } else {
    content.push(...resolvedChildren.content);
  }
  layers.push({
    layer: finalLayer,
    objects: layerObservedObjects,
    reference:
      containerObservation !== undefined
        ? { kind: "object", id: containerObservation.id }
        : { kind: "layout", id: objectId },
  });
}

function compareSourceOrder(
  left: readonly number[],
  right: readonly number[],
): number {
  for (let index = 0; index < Math.min(left.length, right.length); index += 1) {
    const difference = left[index]! - right[index]!;
    if (difference !== 0) return difference;
  }
  return left.length - right.length;
}

function paintOrder(contract?: CapabilityLayoutContract): number {
  const value = contract?.layoutStyle.zIndex;
  return typeof value === "number" && Number.isFinite(value) ? value : 0;
}

function subPartContentNodes(
  parts: readonly CapabilitySubPartObservation[] | undefined,
): QueryContentNode[] {
  if (parts === undefined || parts.length === 0) return [];
  return parts.map((part) => ({
    ...(part.authorId === undefined ? {} : { authorId: part.authorId }),
    reference: {
      kind: "part",
      ...(part.box === undefined ? {} : { box: part.box }),
      ...(part.text === undefined ? {} : { text: part.text }),
    },
    children: subPartContentNodes(part.parts),
  }));
}

function resolveContributions(
  children: readonly CapabilityChild[],
  contributions: readonly CapabilityContribution[],
): readonly CapabilityContribution[] {
  let resolved = contributions;
  const applied = new Set<string>();
  for (const contribution of contributions) {
    const capability = contribution.capability;
    // A resolver is carried by the actual declarations of its capability; module registration and
    // order do not decide whether structure-derived semantics apply.
    if (applied.has(capability)) continue;
    const resolver = contribution.resolveContributions;
    if (resolver === undefined) continue;
    const owned = resolved.filter(
      (candidate) => candidate.capability === capability,
    );
    if (owned.length === 0) continue;
    applied.add(capability);
    const effective = resolver(children, owned);
    const next: CapabilityContribution[] = [];
    let inserted = false;
    for (const candidate of resolved) {
      if (candidate.capability !== capability) {
        next.push(candidate);
      } else if (!inserted) {
        next.push(...effective);
        inserted = true;
      }
    }
    resolved = next;
  }
  return resolved;
}

function collectOwnedContributions(
  children: readonly CapabilityChild[],
  modules: ReadonlyMap<
    string,
    ObjectCapabilityModule | TransparentCapabilityModule
  >,
  slide: SlideResolveState,
  parentPath: string,
  order: { value: number },
  hidden: WeakSet<CapabilityElement>,
  theme: ThemeContext,
  slideNumber: number,
): void {
  children.forEach((child, index) => {
    if (!("props" in child)) return;
    const path = `${parentPath}/${capabilityLabel(child, index)}`;
    const module = modules.get(child.type);
    if (
      module !== undefined &&
      "compileDeclaration" in module &&
      module.children === "owned"
    ) {
      try {
        for (const annotation of module.compileDeclaration(child, {
          path,
          theme,
          slideNumber,
        })) {
          slide.contributions.push({
            ...annotation,
            path,
            order: order.value++,
          });
        }
      } catch (error) {
        slide.diagnostics.push(
          error instanceof CapabilityError
            ? diagnostic(
                error.path === undefined ? path : `${path}/${error.path}`,
                error.code,
                error.message,
              )
            : diagnostic(
                path,
                "capability-failure",
                error instanceof Error ? error.message : String(error),
              ),
        );
      }
      // The owning capability consumes this subtree; do not treat its children as slide content.
      return;
    }
    if (child.pageAffiliated === true) {
      // A page-affiliated declaration is not canvas content. Layout prunes hidden canvas parents,
      // so this structure-only walk reports a trimmed owner that layout can no longer reach.
      if (module === undefined && hidden.has(child)) {
        slide.diagnostics.push(
          diagnostic(
            path,
            "missing-capability",
            `No capability module registered for ${child.type}`,
          ),
        );
      }
      return;
    }
    // Owned declarations are page-affiliated even below a hidden canvas parent, so this
    // structure-only walk still reaches them.
    collectOwnedContributions(
      child.children,
      modules,
      slide,
      path,
      order,
      hidden,
      theme,
      slideNumber,
    );
  });
}

function collectNestedPlaceholders(
  children: readonly CapabilityChild[],
  modules: BoundCapabilityRegistry,
  parentPath: string,
  allowDirect: boolean,
  diagnostics: Diagnostic[],
  onDirect?: (element: CapabilityElement, path: string) => void,
): void {
  children.forEach((child, index) => {
    if (!("props" in child)) return;
    const childPath = `${parentPath}/${capabilityLabel(child, index)}`;
    const module = modules.declarations.get(child.type);
    if (module !== undefined && "compileDeclaration" in module) {
      if (module.children === "owned") return;
      collectNestedPlaceholders(
        child.children,
        modules,
        childPath,
        allowDirect,
        diagnostics,
        onDirect,
      );
      return;
    }
    const hasPhProp =
      child.props["placeholder"] !== undefined ||
      child.props["nativePlaceholder"] !== undefined;
    if (hasPhProp) {
      if (allowDirect) {
        onDirect?.(child, childPath);
      } else {
        diagnostics.push(
          diagnostic(
            childPath,
            "placeholder-inside-container",
            "Placeholders can only be bound directly on Slide",
          ),
        );
      }
    }
    collectNestedPlaceholders(
      child.children,
      modules,
      childPath,
      false,
      diagnostics,
    );
  });
}

function buildSlidePlaceholderContexts(args: {
  readonly slideId: string;
  readonly slidePath: string;
  readonly hasExplicitLayout: boolean;
  readonly layoutSpec: InternalResolvedSlideLayout | undefined;
  readonly masterSpec: InternalResolvedSlideMaster | undefined;
  readonly children: readonly CapabilityChild[];
  readonly rawPlaceholdersProp: unknown;
  readonly modules: BoundCapabilityRegistry;
  readonly diagnostics: Diagnostic[];
}): WeakMap<CapabilityElement, PlaceholderResolutionContext> {
  const {
    slideId,
    slidePath,
    hasExplicitLayout,
    layoutSpec,
    masterSpec,
    children,
    rawPlaceholdersProp,
    modules,
    diagnostics,
  } = args;
  const contexts = new WeakMap<
    CapabilityElement,
    PlaceholderResolutionContext
  >();
  const seenKeys = new Set<string>();

  if (rawPlaceholdersProp !== undefined && !hasExplicitLayout) {
    diagnostics.push(
      diagnostic(
        slidePath,
        "placeholder-requires-layout",
        `Slide ${JSON.stringify(slideId)} declares placeholders without a bound SlideLayout`,
      ),
    );
  }

  collectNestedPlaceholders(
    children,
    modules,
    slidePath,
    true,
    diagnostics,
    (child, childPath) => {
      const rawKey = child.props["placeholder"];
      const rawNativePh = child.props["nativePlaceholder"];
      let parsedNativePh:
        | ReturnType<typeof nativePlaceholderRefSchema.parse>
        | undefined;
      if (rawNativePh !== undefined) {
        const parsed = nativePlaceholderRefSchema.safeParse(rawNativePh);
        if (!parsed.success) {
          diagnostics.push(
            diagnostic(
              childPath,
              "invalid-slide-placeholder",
              parsed.error.issues[0]?.message ??
                "Invalid nativePlaceholder specification",
            ),
          );
          return;
        }
        parsedNativePh = parsed.data;
      }

      if (rawKey !== undefined) {
        if (typeof rawKey !== "string" || rawKey.trim() === "") {
          diagnostics.push(
            diagnostic(
              childPath,
              "invalid-slide-placeholder",
              "placeholder key must be a non-empty string",
            ),
          );
          return;
        }
        if (seenKeys.has(rawKey)) {
          diagnostics.push(
            diagnostic(
              childPath,
              "duplicate-slide-placeholder",
              `Placeholder ${JSON.stringify(rawKey)} can only be bound once per slide`,
            ),
          );
          return;
        }
        seenKeys.add(rawKey);
      }

      if (!hasExplicitLayout && parsedNativePh === undefined) {
        diagnostics.push(
          diagnostic(
            childPath,
            "placeholder-requires-layout",
            `Placeholder ${JSON.stringify(rawKey)} requires a bound SlideLayout`,
          ),
        );
        return;
      }

      const layoutPlaceholder =
        typeof rawKey === "string"
          ? layoutSpec?.placeholders?.[rawKey]
          : undefined;
      if (layoutPlaceholder === undefined && parsedNativePh === undefined) {
        const layoutLabel = layoutSpec?.id ?? layoutSpec?.name ?? "<unnamed>";
        diagnostics.push(
          diagnostic(
            childPath,
            "unknown-slide-placeholder",
            `SlideLayout ${JSON.stringify(layoutLabel)} does not define placeholder ${JSON.stringify(rawKey)}`,
          ),
        );
        return;
      }

      if (masterSpec === undefined) {
        diagnostics.push(
          diagnostic(
            childPath,
            "placeholder-requires-layout",
            `Placeholder ${JSON.stringify(rawKey)} requires a SlideLayout owned by a SlideMaster`,
          ),
        );
        return;
      }

      const nativeRef =
        parsedNativePh ??
        ({
          type: layoutPlaceholder!.type,
          ...(layoutPlaceholder!.explicitType === undefined
            ? {}
            : { explicitType: layoutPlaceholder!.explicitType }),
          ...(layoutPlaceholder!.idx === undefined
            ? {}
            : { idx: layoutPlaceholder!.idx }),
          ...(layoutPlaceholder!.sz === undefined
            ? {}
            : { sz: layoutPlaceholder!.sz }),
          ...(layoutPlaceholder!.orient === undefined
            ? {}
            : { orient: layoutPlaceholder!.orient }),
          ...(layoutPlaceholder!.hasCustomPrompt === undefined
            ? {}
            : { hasCustomPrompt: layoutPlaceholder!.hasCustomPrompt }),
        } as const);
      const role = layoutPlaceholder?.type ?? nativeRef.type;
      const masterPlaceholder = matchMasterPlaceholderSpec(
        masterSpec.placeholders,
        role,
      );
      contexts.set(child, {
        ...(typeof rawKey === "string" ? { key: rawKey } : {}),
        nativeRef,
        masterTheme: masterSpec.themeContext,
        ...(masterSpec.textStyles === undefined
          ? {}
          : { masterTextStyles: masterSpec.textStyles }),
        ...(masterPlaceholder === undefined ? {} : { masterPlaceholder }),
        ...(layoutPlaceholder === undefined ? {} : { layoutPlaceholder }),
      });
    },
  );
  return contexts;
}

type OwnerSlideCapabilityResult = {
  readonly slideCapabilities: readonly ResolvedSlideCapabilityRecord[];
  readonly capabilities: Readonly<Record<string, unknown>>;
  readonly continuityByDeclaration: ReadonlyMap<string, boolean | undefined>;
};

function resolveOwnerSlideCapabilities(args: {
  readonly type: "slide-master" | "slide-layout";
  readonly props: Readonly<Record<string, unknown>>;
  readonly path: string;
  readonly modules: BoundCapabilityRegistry;
  readonly dimensions: ResolutionDimensions;
  readonly document: DocumentResolveState;
  readonly options: ResolveOptions;
  readonly theme: ThemeContext;
}): OwnerSlideCapabilityResult {
  const { type, props, path, modules, dimensions, document, options, theme } =
    args;
  const slideCapabilities: ResolvedSlideCapabilityRecord[] = [];
  const capabilities: Record<string, unknown> = {};
  const continuityByDeclaration = new Map<string, boolean | undefined>();
  const diagnostics: Diagnostic[] = [];
  const input: CapabilityInput = { type, props, children: [] };

  for (const module of modules.slideModules.values()) {
    const slideProps = module.slideProps ?? [module.slideProp];
    const hasSlideProps = slideProps.some((property) =>
      Object.hasOwn(props, property),
    );
    if (!hasSlideProps) continue;
    try {
      const boundSlide = validateBoundSlideCapability(
        module.bindSlide(input, {
          path,
          dimensions,
          theme,
          options,
          slideNumber: 0,
          contributions: [],
        }),
      );
      continuityByDeclaration.set(
        module.declaration,
        boundSlide.continuity?.matchSourceKinds,
      );
      const resolvedSlideCap: ResolvedSlideCapabilityOutput =
        boundSlide.resolveSlide({
          path,
          dimensions,
          contributions: [],
          children: [],
          theme,
          options,
          slideNumber: 0,
        });
      if (resolvedSlideCap.observation !== undefined) {
        capabilities[module.declaration] = resolvedSlideCap.observation;
      }
      for (const extension of resolvedSlideCap.extensions) {
        document.usedExtensions.add(extension);
      }
      recordResourceClaims(
        resolvedSlideCap.resourceClaims,
        document,
        diagnostics,
        path,
      );
      recordCapabilityWarnings(resolvedSlideCap.warnings, document, path);
      slideCapabilities.push({
        declaration: module.declaration,
        slot: module.slot,
        hasSlideProps,
        path,
        theme,
        extensions: resolvedSlideCap.extensions,
        emit: (emitContext) => resolvedSlideCap.emit(emitContext),
        ...(resolvedSlideCap.paint === undefined
          ? {}
          : {
              paint: (paintContext) => resolvedSlideCap.paint!(paintContext),
            }),
      });
    } catch (error) {
      diagnostics.push(
        error instanceof CapabilityError
          ? diagnostic(error.path ?? path, error.code, error.message)
          : diagnostic(
              path,
              "capability-failure",
              error instanceof Error ? error.message : String(error),
            ),
      );
    }
  }
  if (diagnostics.length > 0) throw new BuildError(diagnostics);
  return { slideCapabilities, capabilities, continuityByDeclaration };
}

function resolveSlideDeclaration(
  parsed: ParsedSlideNode,
  modules: BoundCapabilityRegistry,
  dimensions: ResolutionDimensions,
  document: DocumentResolveState,
  options: ResolveOptions,
  documentTheme: ThemeContext,
  layoutIdByKey: ReadonlyMap<string, string>,
  layoutByKey: ReadonlyMap<string, InternalResolvedSlideLayout>,
  masterByKey: ReadonlyMap<string, InternalResolvedSlideMaster>,
  inheritedContinuity?: ReadonlyMap<string, boolean | undefined>,
  inheritedBackground?: ResolvedSlideCapabilityRecord,
  masterLayers?: readonly ResolvedObjectLayer[],
  layoutLayers?: readonly ResolvedObjectLayer[],
  inheritance?: SlideMasterInheritanceContext,
): ResolvedSlideResult {
  const {
    node,
    id,
    authorId,
    index,
    layoutKey,
    hasExplicitLayout,
    section,
    hidden = false,
    showMasterShapes,
  } = parsed;
  const slide: SlideResolveState = {
    index,
    nextId: 2,
    nextNodeOrder: 0,
    diagnostics: [],
    layouts: [],
    contributions: [],
    morphObjects: new Map(),
    matchContinuityKinds: false,
    linkedSlideIndexes: new Set(),
    layoutDerivedObjectIds: new Set(),
  };
  const path = `deck/slide[${id}]`;
  if (nonWhitespaceText(node.children)) {
    slide.diagnostics.push(
      diagnostic(path, "unexpected-text", "Slide cannot directly contain text"),
    );
  }
  const children = node.children.map(capabilityChild);
  const layoutSpec =
    layoutKey === undefined ? undefined : layoutByKey.get(layoutKey);
  const masterSpec =
    layoutSpec === undefined
      ? undefined
      : masterByKey.get(layoutSpec.masterKey);
  const placeholderContexts = buildSlidePlaceholderContexts({
    slideId: id,
    slidePath: path,
    hasExplicitLayout,
    layoutSpec,
    masterSpec,
    children,
    rawPlaceholdersProp: node.props["placeholders"],
    modules,
    diagnostics: slide.diagnostics,
  });
  if (slide.diagnostics.length > 0) {
    throw new BuildError(slide.diagnostics);
  }
  let theme: ThemeContext;
  try {
    theme = documentTheme.with(
      node.props["theme"] as ThemeScope | undefined,
      textStyleFrom(node.props["style"]),
    );
  } catch (error) {
    if (!(error instanceof CapabilityError)) throw error;
    throw new BuildError([diagnostic(path, error.code, error.message)]);
  }
  const layout = new LayoutPass(
    modules.declarations,
    dimensions,
    index,
    options,
  ).run(children, path, theme, node.props["style"], (element) =>
    placeholderContexts.get(element),
  );
  document.warnings.push(...layout.warnings);
  const ownedOrder = { value: 0 };
  collectOwnedContributions(
    children,
    modules.declarations,
    slide,
    path,
    ownedOrder,
    layout.hidden,
    theme,
    index,
  );
  const resolvedChildren = resolveChildrenPass(
    children,
    {
      document,
      slide,
      modules,
      dimensions,
      options,
      theme,
      layout,
      slideNumber: index,
      placeholderContexts,
    },
    path,
  );
  let contributions: readonly CapabilityContribution[] = slide.contributions;
  try {
    contributions = resolveContributions(children, slide.contributions);
  } catch (error) {
    slide.diagnostics.push(
      error instanceof CapabilityError
        ? diagnostic(error.path ?? path, error.code, error.message)
        : diagnostic(
            path,
            "capability-failure",
            error instanceof Error ? error.message : String(error),
          ),
    );
  }

  const capabilities: Record<string, unknown> = {};
  const slideCapabilities: ResolvedSlideCapabilityRecord[] = [];
  const input: CapabilityInput = {
    type: node.type,
    props: node.props,
    children: node.children.map(capabilityChild),
  };

  for (const module of modules.slideModules.values()) {
    const slideProps = module.slideProps ?? [module.slideProp];
    const hasSlideProps = slideProps.some((property) =>
      Object.hasOwn(node.props, property),
    );
    if (!module.always && !hasSlideProps) {
      if (inheritedContinuity?.get(module.declaration) === true) {
        slide.matchContinuityKinds = true;
      }
      continue;
    }
    try {
      const boundSlide = validateBoundSlideCapability(
        module.bindSlide(input, {
          path,
          dimensions,
          theme,
          options,
          slideNumber: index,
          contributions,
          ...(inheritance === undefined ? {} : { inheritance }),
        }),
      );
      const effectiveContinuity =
        boundSlide.continuity?.matchSourceKinds ??
        inheritedContinuity?.get(module.declaration);
      if (effectiveContinuity === true) {
        slide.matchContinuityKinds = true;
      }
      const resolvedSlideCap: ResolvedSlideCapabilityOutput =
        boundSlide.resolveSlide({
          path,
          dimensions,
          contributions,
          children,
          theme,
          options,
          slideNumber: index,
          ...(inheritance === undefined ? {} : { inheritance }),
        });
      if (resolvedSlideCap.observation !== undefined) {
        capabilities[module.declaration] = resolvedSlideCap.observation;
      }
      for (const extension of resolvedSlideCap.extensions) {
        document.usedExtensions.add(extension);
      }
      recordResourceClaims(
        resolvedSlideCap.resourceClaims,
        document,
        slide.diagnostics,
        path,
      );
      recordCapabilityWarnings(resolvedSlideCap.warnings, document, path);
      slideCapabilities.push({
        declaration: module.declaration,
        slot: module.slot,
        hasSlideProps,
        path,
        theme,
        extensions: resolvedSlideCap.extensions,
        emit: (emitContext) => resolvedSlideCap.emit(emitContext),
        ...(resolvedSlideCap.paint === undefined
          ? {}
          : {
              paint: (paintContext) => resolvedSlideCap.paint!(paintContext),
            }),
      });
    } catch (error) {
      slide.diagnostics.push(
        error instanceof CapabilityError
          ? diagnostic(error.path ?? path, error.code, error.message)
          : diagnostic(
              path,
              "capability-failure",
              error instanceof Error ? error.message : String(error),
            ),
      );
    }
  }
  for (const contribution of contributions) {
    if (contribution.optional) continue;
    const consumer = modules.slideModules.get(contribution.capability);
    if (
      consumer === undefined ||
      (!consumer.always &&
        !(consumer.slideProps ?? [consumer.slideProp]).some((property) =>
          Object.hasOwn(node.props, property),
        ))
    ) {
      slide.diagnostics.push(
        diagnostic(
          contribution.path,
          "missing-capability",
          `No capability module registered for ${contribution.capability}`,
        ),
      );
    }
  }
  if (slide.diagnostics.length > 0) throw new BuildError(slide.diagnostics);

  const projectedObjects = projectObservations(
    resolvedChildren.objects,
    slide.layoutDerivedObjectIds,
  );
  const sortedLayouts = [...slide.layouts].sort(
    (left, right) => left.id - right.id,
  );
  const layoutId =
    layoutKey === undefined ? undefined : layoutIdByKey.get(layoutKey);
  const observation: SlideObservation = {
    id,
    ...(authorId === undefined ? {} : { authorId }),
    index,
    ...(layoutId === undefined ? {} : { layout: layoutId }),
    ...(section === undefined ? {} : { section }),
    ...(hidden ? { hidden: true } : {}),
    ...(parsed.showMasterPlaceholderAnimations === undefined
      ? {}
      : {
          showMasterPlaceholderAnimations:
            parsed.showMasterPlaceholderAnimations,
        }),
    objects: projectedObjects,
    ...(sortedLayouts.length > 0 && { layouts: sortedLayouts }),
    ...(Object.keys(capabilities).length > 0 && { capabilities }),
  };

  return {
    observation,
    morphObjects: slide.morphObjects,
    matchContinuityKinds: slide.matchContinuityKinds,
    linkedSlideIndexes: slide.linkedSlideIndexes,
    internal: {
      id,
      ...(authorId === undefined ? {} : { authorId }),
      index,
      ...(layoutKey === undefined ? {} : { layoutKey }),
      ...(parsed.colorMapOverride === undefined
        ? {}
        : { colorMapOverride: parsed.colorMapOverride }),
      ...(parsed.themeOverride === undefined
        ? {}
        : { themeOverride: parsed.themeOverride }),
      ...(section === undefined ? {} : { section }),
      hidden,
      ...(parsed.showMasterPlaceholderAnimations === undefined
        ? {}
        : {
            showMasterPlaceholderAnimations:
              parsed.showMasterPlaceholderAnimations,
          }),
      ...(showMasterShapes === undefined ? {} : { showMasterShapes }),
      ...(inheritedBackground === undefined ? {} : { inheritedBackground }),
      ...(masterLayers === undefined ? {} : { masterLayers }),
      ...(layoutLayers === undefined ? {} : { layoutLayers }),
      content: resolvedChildren.content,
      layers: resolvedChildren.layers,
      slideCapabilities,
    },
  };
}

function resolveSharedCanvasLayers(args: {
  readonly children: readonly CapabilityChild[];
  readonly path: string;
  readonly style: unknown;
  readonly modules: BoundCapabilityRegistry;
  readonly dimensions: ResolutionDimensions;
  readonly options: ResolveOptions;
  readonly theme: ThemeContext;
  readonly document: DocumentResolveState;
  readonly initialNextId?: number;
}): readonly ResolvedObjectLayer[] {
  const {
    children,
    path,
    style,
    modules,
    dimensions,
    options,
    theme,
    document,
    initialNextId = 2,
  } = args;
  if (children.length === 0) return [];
  const state: SlideResolveState = {
    index: 1,
    nextId: initialNextId,
    nextNodeOrder: 0,
    diagnostics: [],
    layouts: [],
    contributions: [],
    morphObjects: new Map(),
    matchContinuityKinds: false,
    linkedSlideIndexes: new Set(),
    layoutDerivedObjectIds: new Set(),
  };
  const layout = new LayoutPass(
    modules.declarations,
    dimensions,
    1,
    options,
  ).run(children, path, theme, style);
  document.warnings.push(...layout.warnings);
  const resolved = resolveChildrenPass(
    children,
    {
      document,
      slide: state,
      modules,
      dimensions,
      options,
      theme,
      layout,
      slideNumber: 1,
      placeholderContexts: new Map(),
    },
    path,
  );
  for (const [, morphObj] of state.morphObjects) {
    state.diagnostics.push(
      diagnostic(
        morphObj.path,
        "morph-on-master",
        "Morph handles are only supported on Slide objects",
      ),
    );
  }
  for (const contribution of state.contributions) {
    if (contribution.optional) continue;
    state.diagnostics.push(
      diagnostic(
        contribution.path,
        "unsupported-declaration",
        `${contribution.capability} declarations are only supported inside Slide`,
      ),
    );
  }
  if (state.diagnostics.length > 0) {
    throw new BuildError(state.diagnostics);
  }
  return resolved.layers;
}

function resolveMasterOrLayoutSlideCapability(args: {
  readonly module: SlideCapabilityModule;
  readonly targetType: "slide-master" | "slide-layout" | "slide";
  readonly props: Readonly<Record<string, unknown>>;
  readonly path: string;
  readonly dimensions: ResolutionDimensions;
  readonly theme: ThemeContext;
  readonly options: ResolveOptions;
  readonly slideNumber: number;
  readonly document: DocumentResolveState;
}): ResolvedSlideCapabilityRecord {
  const {
    module,
    targetType,
    props,
    path,
    dimensions,
    theme,
    options,
    slideNumber,
    document,
  } = args;
  try {
    const input: CapabilityInput = {
      type: targetType,
      props,
      children: [],
    };
    const bound = validateBoundSlideCapability(
      module.bindSlide(input, {
        path,
        dimensions,
        theme,
        options,
        slideNumber,
        contributions: [],
      }),
    );
    const resolvedCap = bound.resolveSlide({
      path,
      dimensions,
      contributions: [],
      children: [],
      theme,
      options,
      slideNumber,
    });
    for (const extension of resolvedCap.extensions) {
      document.usedExtensions.add(extension);
    }
    const claimDiagnostics: Diagnostic[] = [];
    recordResourceClaims(
      resolvedCap.resourceClaims,
      document,
      claimDiagnostics,
      path,
    );
    if (claimDiagnostics.length > 0) {
      throw new BuildError(claimDiagnostics);
    }
    recordCapabilityWarnings(resolvedCap.warnings, document, path);
    return {
      declaration: module.declaration,
      slot: module.slot,
      hasSlideProps: true,
      path,
      theme,
      extensions: resolvedCap.extensions,
      emit: (emitContext) => resolvedCap.emit(emitContext),
      ...(resolvedCap.paint === undefined
        ? {}
        : { paint: (paintContext) => resolvedCap.paint!(paintContext) }),
    };
  } catch (error) {
    if (error instanceof BuildError) throw error;
    throw new BuildError([
      error instanceof CapabilityError
        ? diagnostic(error.path ?? path, error.code, error.message)
        : diagnostic(
            path,
            "capability-failure",
            error instanceof Error ? error.message : String(error),
          ),
    ]);
  }
}

function deriveFooterPlaceholdersFromMap(
  placeholders: PlaceholderMap | undefined,
): FooterPlaceholders | undefined {
  if (placeholders === undefined) return undefined;
  const result: {
    dateTime?: FooterPlaceholderSpec;
    footer?: FooterPlaceholderSpec;
    slideNumber?: FooterPlaceholderSpec;
  } = {};
  for (const spec of Object.values(placeholders)) {
    const slot =
      spec.type === "dt"
        ? "dateTime"
        : spec.type === "ftr"
          ? "footer"
          : spec.type === "sldNum"
            ? "slideNumber"
            : undefined;
    if (slot === undefined || result[slot] !== undefined) continue;
    const mergedStyle =
      spec.style !== undefined || spec.listStyles?.[1] !== undefined
        ? { ...spec.style, ...spec.listStyles?.[1] }
        : undefined;
    const candidate = {
      ...(spec.x === undefined ? {} : { x: spec.x }),
      ...(spec.y === undefined ? {} : { y: spec.y }),
      ...(spec.width === undefined ? {} : { width: spec.width }),
      ...(spec.height === undefined ? {} : { height: spec.height }),
      ...(spec.idx === undefined ? {} : { idx: spec.idx }),
      ...(spec.sz === "full" || spec.sz === "half" || spec.sz === "quarter"
        ? { size: spec.sz }
        : {}),
      ...(spec.promptText === undefined ? {} : { text: spec.promptText }),
      ...(mergedStyle === undefined ? {} : { style: mergedStyle }),
      visible: false,
    };
    result[slot] = candidate;
  }
  const parsed = footerPlaceholdersSchema.safeParse(result);
  if (!parsed.success) return undefined;
  return Object.keys(parsed.data).length > 0 ? parsed.data : undefined;
}

function resolveDocument(
  container: DeclarationContainer,
  modules: BoundCapabilityRegistry,
  options: ResolveOptions,
): Resolution {
  const decks = childNodes(container.children);
  if (decks.length !== 1 || decks[0]?.type !== "deck") {
    throw new BuildError([
      diagnostic(
        "deck",
        "document-root",
        "Build content must have exactly one Deck root component",
      ),
    ]);
  }
  const deck = decks[0];
  const { width, height, title } = parseDeckProps(deck.props);
  const size: ResolutionDimensions = { width, height };
  let theme: ThemeContext;
  try {
    theme = createThemeContext(
      deck.props["theme"] as ThemeInput | Theme | undefined,
      deck.props["style"] as Record<string, unknown> | undefined,
    );
  } catch (error) {
    if (!(error instanceof CapabilityError)) throw error;
    throw new BuildError([diagnostic("deck", error.code, error.message)]);
  }
  const parsed = resolveSlides(container, modules, theme.theme);
  const { masters: parsedMasters, layoutIdByKey, slides, sections } = parsed;
  const slideIndexById = new Map(
    slides.map((slide) => [slide.id, slide.index] as const),
  );
  const sharedCanvasRoots: DeclarationNode[] = [
    ...parsedMasters.flatMap((m) => m.canvasChildren),
    ...parsedMasters.flatMap((m) =>
      m.layouts.flatMap((l) => childNodes(l.node.children)),
    ),
  ];
  const document: DocumentResolveState = {
    usedExtensions: new Set(sections.length > 0 ? ["p14"] : []),
    warnings: [],
    resourceClaims: new Map(),
    morphNames: new Map(),
    nextMorphName: 1,
    reservedObjectNames: collectExplicitObjectNames(
      slides,
      modules.declarations,
      sharedCanvasRoots,
    ),
    slideIndexById,
  };

  const bgModule =
    [...modules.slideModules.values()].find(
      (module) =>
        module.slot === "background" &&
        module.declaration === "slide-background",
    ) ??
    [...modules.slideModules.values()].find(
      (module) => module.slot === "background",
    );
  const footerModule = [...modules.slideModules.values()].find(
    (module) =>
      module.slot === "shape-tree" &&
      module.declaration === "footer-placeholder",
  );

  const resolvedMasters: InternalResolvedSlideMaster[] = [];
  const masterObservations: SlideMasterObservation[] = [];
  const masterByKey = new Map<string, InternalResolvedSlideMaster>();
  const masterByLayoutKey = new Map<string, InternalResolvedSlideMaster>();
  const layoutByKey = new Map<string, InternalResolvedSlideLayout>();
  const parsedMasterByKey = new Map<string, ParsedMasterSpec>();
  const parsedLayoutByKey = new Map<
    string,
    { readonly master: ParsedMasterSpec; readonly layout: ParsedLayoutSpec }
  >();
  const continuityByLayoutKey = new Map<
    string,
    ReadonlyMap<string, boolean | undefined>
  >();

  for (const parsedMaster of parsedMasters) {
    parsedMasterByKey.set(parsedMaster.key, parsedMaster);
    const mergedMasterStyle =
      parsedMaster.style === undefined
        ? (deck.props["style"] as Record<string, unknown> | undefined)
        : {
            ...(deck.props["style"] as Record<string, unknown> | undefined),
            ...parsedMaster.style,
          };
    let masterThemeBase: ThemeContext;
    try {
      masterThemeBase = createThemeContext(
        parsedMaster.theme,
        mergedMasterStyle,
        { colorMap: parsedMaster.colorMap },
      );
    } catch (error) {
      if (!(error instanceof CapabilityError)) throw error;
      throw new BuildError([
        diagnostic(parsedMaster.path, error.code, error.message),
      ]);
    }
    const masterCapResult = resolveOwnerSlideCapabilities({
      type: "slide-master",
      props: parsedMaster.props,
      path: parsedMaster.path,
      modules,
      dimensions: size,
      document,
      options,
      theme: masterThemeBase,
    });
    let masterBackground: ResolvedSlideCapabilityRecord | undefined;
    if (parsedMaster.background !== undefined) {
      if (bgModule === undefined) {
        throw new BuildError([
          diagnostic(
            parsedMaster.path,
            "missing-capability",
            "No capability module registered for SlideMaster.background",
          ),
        ]);
      }
      masterBackground = resolveMasterOrLayoutSlideCapability({
        module: bgModule,
        targetType: "slide-master",
        props: { background: parsedMaster.background },
        path: parsedMaster.path,
        dimensions: size,
        theme: masterThemeBase,
        options,
        slideNumber: 1,
        document,
      });
    }
    let masterFooterPlaceholders: ResolvedSlideCapabilityRecord | undefined;
    if (parsedMaster.footerPlaceholders !== undefined) {
      if (footerModule === undefined) {
        throw new BuildError([
          diagnostic(
            parsedMaster.path,
            "missing-capability",
            "No capability module registered for SlideMaster.footerPlaceholders",
          ),
        ]);
      }
      masterFooterPlaceholders = resolveMasterOrLayoutSlideCapability({
        module: footerModule,
        targetType: "slide-master",
        props: { footerPlaceholders: parsedMaster.footerPlaceholders },
        path: parsedMaster.path,
        dimensions: size,
        theme: masterThemeBase,
        options,
        slideNumber: 1,
        document,
      });
    }
    const masterLayers = resolveSharedCanvasLayers({
      children: parsedMaster.canvasChildren.map(capabilityChild),
      path: parsedMaster.path,
      style: parsedMaster.node.props["style"],
      modules,
      dimensions: size,
      options,
      theme: masterThemeBase,
      document,
      initialNextId: 2 + Object.keys(parsedMaster.placeholders ?? {}).length,
    });

    const resolvedLayouts: InternalResolvedSlideLayout[] = [];
    const observedLayouts: SlideLayoutObservation[] = [];

    for (const parsedLayout of parsedMaster.layouts) {
      parsedLayoutByKey.set(parsedLayout.key, {
        master: parsedMaster,
        layout: parsedLayout,
      });
      let layoutThemeBase = masterThemeBase;
      if (
        parsedLayout.data.themeOverride !== undefined ||
        parsedLayout.data.colorMapOverride !== undefined
      ) {
        layoutThemeBase = layoutThemeBase.withNativeOverrides(
          parsedLayout.data.themeOverride,
          parsedLayout.data.colorMapOverride,
        );
      }
      const layoutCapResult = resolveOwnerSlideCapabilities({
        type: "slide-layout",
        props: parsedLayout.props,
        path: parsedLayout.path,
        modules,
        dimensions: size,
        document,
        options,
        theme: layoutThemeBase,
      });
      const mergedContinuity = new Map<string, boolean | undefined>(
        masterCapResult.continuityByDeclaration,
      );
      for (const [decl, match] of layoutCapResult.continuityByDeclaration) {
        if (match !== undefined) {
          mergedContinuity.set(decl, match);
        }
      }
      continuityByLayoutKey.set(parsedLayout.key, mergedContinuity);

      let layoutBackground: ResolvedSlideCapabilityRecord | undefined;
      if (parsedLayout.background !== undefined) {
        if (bgModule === undefined) {
          throw new BuildError([
            diagnostic(
              parsedLayout.path,
              "missing-capability",
              "No capability module registered for SlideLayout.background",
            ),
          ]);
        }
        layoutBackground = resolveMasterOrLayoutSlideCapability({
          module: bgModule,
          targetType: "slide-layout",
          props: { background: parsedLayout.background },
          path: parsedLayout.path,
          dimensions: size,
          theme: layoutThemeBase,
          options,
          slideNumber: 1,
          document,
        });
      }
      let layoutFooterPlaceholders: ResolvedSlideCapabilityRecord | undefined;
      if (parsedLayout.footerPlaceholders !== undefined) {
        if (footerModule === undefined) {
          throw new BuildError([
            diagnostic(
              parsedLayout.path,
              "missing-capability",
              "No capability module registered for SlideLayout.footerPlaceholders",
            ),
          ]);
        }
        layoutFooterPlaceholders = resolveMasterOrLayoutSlideCapability({
          module: footerModule,
          targetType: "slide-layout",
          props: { footerPlaceholders: parsedLayout.footerPlaceholders },
          path: parsedLayout.path,
          dimensions: size,
          theme: layoutThemeBase,
          options,
          slideNumber: 1,
          document,
        });
      }
      const layoutLayers = resolveSharedCanvasLayers({
        children: parsedLayout.node.children.map(capabilityChild),
        path: parsedLayout.path,
        style: parsedLayout.node.props["style"],
        modules,
        dimensions: size,
        options,
        theme: layoutThemeBase,
        document,
        initialNextId: 2 + Object.keys(parsedLayout.placeholders ?? {}).length,
      });

      const resolvedLayout: InternalResolvedSlideLayout = {
        key: parsedLayout.key,
        masterKey: parsedMaster.key,
        ...parsedLayout.data,
        ...(parsedLayout.placeholders === undefined
          ? {}
          : { placeholders: parsedLayout.placeholders }),
        ...(layoutBackground === undefined
          ? {}
          : { background: layoutBackground }),
        ...(layoutFooterPlaceholders === undefined
          ? {}
          : { footerPlaceholders: layoutFooterPlaceholders }),
        ...(layoutLayers.length === 0 ? {} : { layers: layoutLayers }),
        ...(layoutCapResult.slideCapabilities.length > 0
          ? { slideCapabilities: layoutCapResult.slideCapabilities }
          : {}),
      };
      resolvedLayouts.push(resolvedLayout);
      layoutByKey.set(parsedLayout.key, resolvedLayout);
      observedLayouts.push({
        ...parsedLayout.data,
        ...(Object.keys(layoutCapResult.capabilities).length > 0
          ? { capabilities: layoutCapResult.capabilities }
          : {}),
      });
    }

    const resolvedMaster: InternalResolvedSlideMaster = {
      key: parsedMaster.key,
      id: parsedMaster.id,
      ...(parsedMaster.name === undefined ? {} : { name: parsedMaster.name }),
      ...(parsedMaster.preserve === undefined
        ? {}
        : { preserve: parsedMaster.preserve }),
      theme: parsedMaster.theme,
      themeContext: masterThemeBase,
      colorMap: parsedMaster.colorMap,
      ...(parsedMaster.headerFooter === undefined
        ? {}
        : { headerFooter: parsedMaster.headerFooter }),
      ...(masterBackground === undefined
        ? {}
        : { background: masterBackground }),
      ...(masterFooterPlaceholders === undefined
        ? {}
        : { footerPlaceholders: masterFooterPlaceholders }),
      ...(masterLayers.length === 0 ? {} : { layers: masterLayers }),
      ...(parsedMaster.placeholders === undefined
        ? {}
        : { placeholders: parsedMaster.placeholders }),
      ...(parsedMaster.textStyles === undefined
        ? {}
        : { textStyles: parsedMaster.textStyles }),
      ...(masterCapResult.slideCapabilities.length > 0
        ? { slideCapabilities: masterCapResult.slideCapabilities }
        : {}),
      layouts: resolvedLayouts,
    };
    masterByKey.set(parsedMaster.key, resolvedMaster);
    resolvedMasters.push(resolvedMaster);
    for (const resolvedLayout of resolvedLayouts) {
      masterByLayoutKey.set(resolvedLayout.key, resolvedMaster);
    }
    masterObservations.push({
      id: parsedMaster.id,
      ...(parsedMaster.name === undefined ? {} : { name: parsedMaster.name }),
      ...(parsedMaster.preserve === undefined
        ? {}
        : { preserve: parsedMaster.preserve }),
      ...(parsedMaster.explicitColorMap === undefined
        ? {}
        : { colorMap: parsedMaster.explicitColorMap }),
      ...(parsedMaster.headerFooter === undefined
        ? {}
        : { headerFooter: parsedMaster.headerFooter }),
      ...(Object.keys(masterCapResult.capabilities).length > 0
        ? { capabilities: masterCapResult.capabilities }
        : {}),
      layouts: observedLayouts,
    });
  }

  const resolvedSlides = slides.map((parsedSlide) => {
    const boundMaster =
      parsedSlide.layoutKey === undefined
        ? undefined
        : masterByLayoutKey.get(parsedSlide.layoutKey);
    const boundLayout =
      parsedSlide.layoutKey === undefined
        ? undefined
        : layoutByKey.get(parsedSlide.layoutKey);
    let slideThemeBase: ThemeContext;
    if (boundMaster !== undefined) {
      slideThemeBase = createThemeContext(
        boundMaster.theme,
        deck.props["style"] as Record<string, unknown> | undefined,
        { colorMap: boundMaster.colorMap },
      );
      if (
        boundLayout?.themeOverride !== undefined ||
        boundLayout?.colorMapOverride !== undefined
      ) {
        slideThemeBase = slideThemeBase.withNativeOverrides(
          boundLayout.themeOverride,
          boundLayout.colorMapOverride,
        );
      }
    } else {
      slideThemeBase = theme;
    }
    if (
      parsedSlide.themeOverride !== undefined ||
      parsedSlide.colorMapOverride !== undefined
    ) {
      slideThemeBase = slideThemeBase.withNativeOverrides(
        parsedSlide.themeOverride,
        parsedSlide.colorMapOverride,
      );
    }

    const parsedOwner =
      parsedSlide.layoutKey === undefined
        ? undefined
        : parsedLayoutByKey.get(parsedSlide.layoutKey);
    const parsedMaster = parsedOwner?.master;
    const parsedLayout = parsedOwner?.layout;

    let inheritedBackground: ResolvedSlideCapabilityRecord | undefined;
    if (
      !Object.hasOwn(parsedSlide.node.props, "background") &&
      parsedSlide.layoutKey !== undefined &&
      bgModule !== undefined
    ) {
      const inheritedDecl =
        parsedLayout?.background ?? parsedMaster?.background;
      if (inheritedDecl !== undefined) {
        inheritedBackground = resolveMasterOrLayoutSlideCapability({
          module: bgModule,
          targetType: "slide",
          props: { background: inheritedDecl },
          path: `deck/slide[${parsedSlide.id}]`,
          dimensions: size,
          theme: slideThemeBase,
          options,
          slideNumber: parsedSlide.index,
          document,
        });
      }
    }

    const masterFooter =
      parsedMaster?.footerPlaceholders ??
      deriveFooterPlaceholdersFromMap(parsedMaster?.placeholders);
    const layoutFooter =
      parsedLayout?.footerPlaceholders ??
      deriveFooterPlaceholdersFromMap(parsedLayout?.placeholders);
    const inheritance: SlideMasterInheritanceContext | undefined =
      parsedMaster !== undefined || parsedLayout !== undefined
        ? {
            ...(parsedMaster === undefined
              ? {}
              : {
                  master: {
                    ...(parsedMaster.headerFooter === undefined
                      ? {}
                      : { headerFooter: parsedMaster.headerFooter }),
                    ...(masterFooter === undefined
                      ? {}
                      : { footerPlaceholders: masterFooter }),
                  },
                }),
            ...(parsedLayout === undefined
              ? {}
              : {
                  layout: {
                    ...(parsedLayout.headerFooter === undefined
                      ? {}
                      : { headerFooter: parsedLayout.headerFooter }),
                    ...(layoutFooter === undefined
                      ? {}
                      : { footerPlaceholders: layoutFooter }),
                  },
                }),
          }
        : undefined;

    const slideShowsMaster = parsedSlide.showMasterShapes ?? true;
    const layoutShowsMaster = boundLayout?.showMasterShapes ?? true;
    const activeMasterLayers =
      slideShowsMaster &&
      layoutShowsMaster &&
      (boundMaster?.layers?.length ?? 0) > 0
        ? boundMaster!.layers
        : undefined;
    const activeLayoutLayers =
      slideShowsMaster && (boundLayout?.layers?.length ?? 0) > 0
        ? boundLayout!.layers
        : undefined;

    return resolveSlideDeclaration(
      parsedSlide,
      modules,
      size,
      document,
      options,
      slideThemeBase,
      layoutIdByKey,
      layoutByKey,
      masterByKey,
      parsedSlide.layoutKey === undefined
        ? undefined
        : continuityByLayoutKey.get(parsedSlide.layoutKey),
      inheritedBackground,
      activeMasterLayers,
      activeLayoutLayers,
      inheritance,
    );
  });
  void parsedMasterByKey;

  const morphDiagnostics: Diagnostic[] = [];
  for (let index = 0; index < slides.length; index += 1) {
    const current = slides[index];
    const currentResult = resolvedSlides[index];
    if (
      current === undefined ||
      currentResult === undefined ||
      !currentResult.matchContinuityKinds
    ) {
      continue;
    }
    const sourceIndexes = new Set<number>();
    if (!current.hidden) {
      let priorVisible = index - 1;
      while (priorVisible >= 0 && slides[priorVisible]?.hidden === true)
        priorVisible -= 1;
      if (priorVisible >= 0) sourceIndexes.add(priorVisible);
    }
    for (
      let sourceIndex = 0;
      sourceIndex < resolvedSlides.length;
      sourceIndex += 1
    ) {
      if (
        sourceIndex !== index &&
        resolvedSlides[sourceIndex]?.linkedSlideIndexes.has(current.index)
      ) {
        sourceIndexes.add(sourceIndex);
      }
    }
    for (const [key, object] of currentResult.morphObjects) {
      for (const sourceIndex of sourceIndexes) {
        const source = resolvedSlides[sourceIndex]?.morphObjects.get(key);
        if (source === undefined || source.kind === object.kind) continue;
        morphDiagnostics.push(
          diagnostic(
            object.path,
            "morph-object-kind-mismatch",
            `Native object ${object.name} (handle "${getMorphKeyLabel(key)}") corresponds to ${source.kind} on source slide ${slides[sourceIndex]?.id} and ${object.kind} on current slide; only matching native object kinds are supported`,
          ),
        );
      }
    }
  }
  if (morphDiagnostics.length > 0) throw new BuildError(morphDiagnostics);

  const slideObservations: SlideObservation[] = resolvedSlides.map(
    ({ observation }) => observation,
  );
  const report: ResolutionReport = deepFreeze({
    evidence: "generation",
    ...(typeof title === "string" ? { title } : {}),
    extensions: [...document.usedExtensions],
    ...(masterObservations.length > 0 ? { masters: masterObservations } : {}),
    slides: slideObservations,
    diagnostics: document.warnings,
  });
  const query = createBuildQuery(
    resolvedSlides.map((slide, index) => ({
      observation: report.slides[index]!,
      content: slide.internal.content,
    })),
  );
  const resolution: Resolution = Object.freeze({ report, query });
  const internal: InternalResolution = {
    dimensions: size,
    ...(typeof title === "string" ? { title } : {}),
    theme,
    ...(resolvedMasters.length > 0 ? { masters: resolvedMasters } : {}),
    sections,
    slides: resolvedSlides.map(({ internal: slideInternal }) => slideInternal),
    resourceClaims: document.resourceClaims,
    warnings: [...document.warnings],
  };
  registerInternalResolution(resolution, internal);
  return resolution;
}

function normalizeDeclarationInput(
  input: DeclarationContainer | DeclarationNode,
): DeclarationContainer {
  if (isDeclarationNode(input)) return { children: [input] };
  if (
    typeof input === "object" &&
    input !== null &&
    "children" in input &&
    Array.isArray(input.children)
  ) {
    return input;
  }
  return { children: [] };
}

export function createResolver(modules: readonly CapabilityModule[]) {
  const bound = bindCapabilityModules(modules);

  return function resolve(
    declaration: DeclarationContainer | DeclarationNode,
    options: ResolveOptions = {},
  ): Resolution {
    validateResolveOptions(options);
    if (options.signal?.aborted) throw options.signal.reason;
    return resolveDocument(
      normalizeDeclarationInput(declaration),
      bound,
      options,
    );
  };
}

function validateBoundSlideCapability(bound: unknown): BoundSlideCapability {
  if (typeof bound !== "object" || bound === null) {
    throw new CapabilityError(
      "invalid-capability-binding",
      "Slide capability bindSlide() must return a resolveSlide handler",
    );
  }
  const candidate = bound as {
    readonly resolveSlide?: unknown;
  };
  if (typeof candidate.resolveSlide !== "function") {
    throw new CapabilityError(
      "invalid-capability-binding",
      "Slide capability bindSlide() must return a resolveSlide handler",
    );
  }
  return bound as BoundSlideCapability;
}

function bindCapabilityModules(
  modules: readonly CapabilityModule[],
): BoundCapabilityRegistry {
  const declarations = new Map<
    string,
    ObjectCapabilityModule | TransparentCapabilityModule
  >();
  const slideModules = new Map<string, SlideCapabilityModule>();
  for (const module of modules) {
    if (typeof module !== "object" || module === null) {
      throw new BuildError([
        diagnostic(
          "build.modules",
          "invalid-capability-module",
          "Capability module must be an object",
        ),
      ]);
    }
    const slideModule =
      "bindSlide" in module &&
      typeof module.bindSlide === "function" &&
      !("bind" in module) &&
      !("parse" in module)
        ? (module as SlideCapabilityModule)
        : undefined;
    const objectModule =
      "bind" in module &&
      typeof module.bind === "function" &&
      !("parse" in module) &&
      !("bindSlide" in module)
        ? (module as ObjectCapabilityModule)
        : undefined;
    if (slideModule === undefined && objectModule === undefined) {
      const rawDeclaration = (module as { declaration?: unknown }).declaration;
      const declarationLabel =
        typeof rawDeclaration === "string" && rawDeclaration.trim() !== ""
          ? rawDeclaration
          : "<unknown>";
      throw new BuildError([
        diagnostic(
          "build.modules",
          "invalid-capability-module",
          `Capability module ${declarationLabel} must implement the bound bind() or bindSlide() contract`,
        ),
      ]);
    }
    if (
      typeof module.declaration !== "string" ||
      module.declaration.trim() === ""
    ) {
      throw new BuildError([
        diagnostic(
          "build.modules",
          "invalid-capability-module",
          "Capability module declaration must be a non-empty string",
        ),
      ]);
    }
    if (slideModule !== undefined) {
      if (slideModules.has(slideModule.declaration)) {
        throw new BuildError([
          diagnostic(
            "build.modules",
            "duplicate-capability",
            "Duplicate capability module declaration",
          ),
        ]);
      }
      slideModules.set(slideModule.declaration, slideModule);
      for (const transparent of slideModule.transparentDeclarations ?? []) {
        if (
          typeof transparent.declaration !== "string" ||
          transparent.declaration.trim() === ""
        ) {
          throw new BuildError([
            diagnostic(
              "build.modules",
              "invalid-capability-module",
              "Capability module declaration must be a non-empty string",
            ),
          ]);
        }
        if (
          !("compileDeclaration" in transparent) ||
          typeof transparent.compileDeclaration !== "function"
        ) {
          throw new BuildError([
            diagnostic(
              "build.modules",
              "invalid-capability-module",
              `Transparent declaration ${transparent.declaration} must implement compileDeclaration()`,
            ),
          ]);
        }
        if (declarations.has(transparent.declaration)) {
          throw new BuildError([
            diagnostic(
              "build.modules",
              "duplicate-capability",
              "Duplicate capability module declaration",
            ),
          ]);
        }
        declarations.set(transparent.declaration, transparent);
      }
    } else if (objectModule !== undefined) {
      if (declarations.has(objectModule.declaration)) {
        throw new BuildError([
          diagnostic(
            "build.modules",
            "duplicate-capability",
            "Duplicate capability module declaration",
          ),
        ]);
      }
      declarations.set(objectModule.declaration, objectModule);
    }
  }
  return { declarations, slideModules };
}
