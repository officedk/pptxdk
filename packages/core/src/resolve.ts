import { readAuthorId, validateAuthorIdScopes } from "#/author-ids.ts";
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
import { getMorphKeyLabel, isMorphKey, type MorphKey } from "#/morph.ts";
import { createBuildQuery, type QueryContentNode } from "#/query.ts";
import {
  registerInternalResolution,
  type InternalResolution,
  type InternalResolvedSlide,
  type ResolutionDimensions,
  type ResolvedObjectLayer,
  type ResolvedSectionRecord,
  type ResolvedSlideCapabilityRecord,
} from "#/resolution.ts";
import { resolveOptionsSchema } from "#/stage-options.ts";
import {
  createThemeContext,
  textStyleFrom,
  type ThemeContext,
  type ThemeInput,
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
  RegisteredCapabilityModule,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  ResolvedLink,
  ResolvedObjectChildren,
  ResolvedSlide as ResolvedSlideCapabilityOutput,
  SlideCapabilityModule,
  SlideObservation,
  TransparentCapabilityModule,
} from "#/types.ts";

type ParsedSlideNode = {
  readonly node: DeclarationNode;
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden?: boolean;
};

type ParsedDocument = {
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

const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;
const emuPerPx = 9525;

function pxToEmu(value: number): number {
  return Math.round(value * emuPerPx);
}

function assertXml10Characters(value: string): void {
  for (const character of value) {
    const codePoint = character.codePointAt(0);
    if (
      codePoint !== undefined &&
      codePoint !== 0x9 &&
      codePoint !== 0xa &&
      codePoint !== 0xd &&
      (codePoint < 0x20 ||
        (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
        codePoint === 0xfffe ||
        codePoint === 0xffff)
    ) {
      throw new Error("String contains characters not allowed in XML 1.0");
    }
  }
}

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

function ownKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
): string | undefined {
  return Object.keys(value).find((key) => !allowed.includes(key));
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
): ReadonlySet<string> {
  const names = new Set<string>();
  for (const slide of slides) {
    collectExplicitObjectNamesFromChildren(slide.node.children, modules, names);
  }
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

function positiveDimension(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new BuildError([
      diagnostic(
        "deck",
        "invalid-dimension",
        `${label} must be a finite positive pixel value`,
      ),
    ]);
  }
  const sizeEmu = pxToEmu(value);
  if (sizeEmu < minimumSlideSizeEmu || sizeEmu > maximumSlideSizeEmu) {
    throw new BuildError([
      diagnostic(
        "deck",
        "invalid-dimension",
        `${label} converted slide size must be within the closed interval of ${minimumSlideSizeEmu} to ${maximumSlideSizeEmu} EMU`,
      ),
    ]);
  }
  return value;
}

function deckDimensions(node: DeclarationNode): ResolutionDimensions {
  const unknown = ownKeys(node.props, [
    "width",
    "height",
    "theme",
    "style",
    "title",
  ]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic(
        "deck",
        "unknown-parameter",
        `Deck contains unknown parameter ${unknown}`,
      ),
    ]);
  }
  return {
    width: positiveDimension(node.props["width"] ?? 1280, "Deck.width"),
    height: positiveDimension(node.props["height"] ?? 720, "Deck.height"),
  };
}

function resolveSlides(
  container: DeclarationContainer,
  modules: BoundCapabilityRegistry,
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
        "Deck can only contain Slide and Section components",
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
  const sectioned = topLevel.some((node) => node.type === "section");
  const candidates: {
    node: DeclarationNode;
    section?: string;
    sectionIndex?: number;
  }[] = [];
  const sections: { name: string; slideIndexes: number[] }[] = [];
  if (sectioned) {
    for (const [index, node] of topLevel.entries()) {
      if (node.type !== "section") {
        throw new BuildError([
          diagnostic(
            `deck/${componentLabel(node, index)}`,
            "mixed-sections",
            "When using Section, every slide must belong to a Section",
          ),
        ]);
      }
      const unknown = ownKeys(node.props, ["name"]);
      if (unknown !== undefined) {
        throw new BuildError([
          diagnostic(
            `deck/section[${index + 1}]`,
            "unknown-section-property",
            `Section does not support property ${unknown}`,
          ),
        ]);
      }
      const name = node.props["name"];
      if (typeof name !== "string" || name.trim() === "") {
        throw new BuildError([
          diagnostic(
            `deck/section[${index + 1}]`,
            "invalid-section-name",
            "Section.name must be a non-empty string",
          ),
        ]);
      }
      try {
        assertXml10Characters(name);
      } catch {
        throw new BuildError([
          diagnostic(
            `deck/section[${name}]`,
            "invalid-xml-character",
            "Section.name contains characters not allowed in XML 1.0",
          ),
        ]);
      }
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
    for (const [index, node] of topLevel.entries()) {
      if (node.type !== "slide") {
        throw new BuildError([
          diagnostic(
            `deck/${componentLabel(node, index)}`,
            "unexpected-declaration",
            "Direct children of Deck must be Slide or Section",
          ),
        ]);
      }
      candidates.push({ node });
    }
  }
  if (candidates.length === 0) {
    throw new BuildError([
      diagnostic("deck", "missing-slide", "Deck requires at least one Slide"),
    ]);
  }
  const authorIds = new Set<string>();
  const slideProperties = [...modules.slideModules.values()].flatMap(
    (module) => module.slideProps ?? [module.slideProp],
  );
  const slides = candidates.map(({ node, section, sectionIndex }, index) => {
    const unknown = ownKeys(node.props, [
      "id",
      "theme",
      "style",
      "hidden",
      ...slideProperties,
    ]);
    const authorId = readAuthorId(node.props["id"], `deck/slide[${index + 1}]`);
    if (unknown !== undefined) {
      const label = authorId ?? index + 1;
      throw new BuildError([
        diagnostic(
          `deck/slide[${label}]`,
          "missing-capability",
          `No capability module registered for Slide.${unknown}`,
        ),
      ]);
    }
    if (authorId === undefined) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-author-id",
          "Slide must declare id as a non-empty string",
        ),
      ]);
    }
    try {
      assertXml10Characters(authorId);
    } catch {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-xml-character",
          "Slide.id contains characters not allowed in XML 1.0",
        ),
      ]);
    }
    const id = authorId;
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
    const hidden = node.props["hidden"];
    if (hidden !== undefined && typeof hidden !== "boolean") {
      throw new BuildError([
        diagnostic(
          `deck/slide[${id}]`,
          "invalid-hidden-slide",
          "Slide.hidden must be a boolean",
        ),
      ]);
    }
    const resolved = {
      node,
      id,
      ...(authorId === undefined ? {} : { authorId }),
      index: index + 1,
      ...(section === undefined ? {} : { section }),
      ...(hidden === undefined ? {} : { hidden }),
    };
    if (sectionIndex !== undefined) {
      sections[sectionIndex]?.slideIndexes.push(index + 1);
    }
    return resolved;
  });
  return {
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
  if (value === undefined) return undefined;
  const reportInvalid = (message: string) => {
    context.slide.diagnostics.push(diagnostic(path, "invalid-link", message));
  };
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    reportInvalid("link must specify id, url, or home action");
    return undefined;
  }
  const entries = Object.entries(value).filter(
    ([, target]) => target !== undefined,
  );
  if (entries.length !== 1) {
    reportInvalid("link must specify exactly one target");
    return undefined;
  }
  const [kind, target] = entries[0] ?? [];
  if (kind === "id" && typeof target === "string") {
    const targetIndex = context.document.slideIndexById.get(target);
    if (targetIndex === undefined) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-link-target",
          `No target slide with author id ${JSON.stringify(target)}`,
        ),
      );
      return undefined;
    }
    if (targetIndex !== context.slide.index) {
      context.slide.linkedSlideIndexes.add(targetIndex);
    }
    return { target: { id: target }, summary: `slide:${target}` };
  }
  if (kind === "url" && typeof target === "string") {
    let url: URL;
    try {
      url = new URL(target);
    } catch {
      reportInvalid("link.url must be an absolute URL");
      return undefined;
    }
    return { target: { url: url.href }, summary: `url:${url.href}` };
  }
  if (kind === "action" && target === "home") {
    return { target: { action: "home" }, summary: "home" };
  }
  reportInvalid("link supports { id }, { url }, or { action: 'home' }");
  return undefined;
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
    const theme = context.layout.themes.get(child) ?? context.theme;
    try {
      const bound =
        context.layout.bound.get(child) ??
        validateBoundObjectCapability(
          module.bind(child, {
            path,
            theme,
            options: context.options,
            slideNumber: context.slideNumber,
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

function resolveSlideDeclaration(
  parsed: ParsedSlideNode,
  modules: BoundCapabilityRegistry,
  dimensions: ResolutionDimensions,
  document: DocumentResolveState,
  options: ResolveOptions,
  documentTheme: ThemeContext,
): ResolvedSlideResult {
  const { node, id, authorId, index, section, hidden = false } = parsed;
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
  ).run(children, path, theme, node.props["style"]);
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
    if (!module.always && !hasSlideProps) continue;
    try {
      const boundSlide = validateBoundSlideCapability(
        module.bindSlide(input, {
          path,
          dimensions,
          theme,
          options,
          slideNumber: index,
          contributions,
        }),
      );
      if (boundSlide.continuity?.matchSourceKinds === true) {
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
  const observation: SlideObservation = {
    id,
    ...(authorId === undefined ? {} : { authorId }),
    index,
    ...(section === undefined ? {} : { section }),
    ...(hidden ? { hidden: true } : {}),
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
      ...(section === undefined ? {} : { section }),
      hidden,
      content: resolvedChildren.content,
      layers: resolvedChildren.layers,
      slideCapabilities,
    },
  };
}

function resolveDocument(
  container: DeclarationContainer,
  modules: BoundCapabilityRegistry,
  options: ResolveOptions,
): Resolution {
  const parsed = resolveSlides(container, modules);
  const { slides, sections } = parsed;
  const deck = childNodes(container.children)[0];
  if (deck === undefined)
    throw new BuildError([
      diagnostic("deck", "document-root", "Missing Deck root component"),
    ]);
  const size = deckDimensions(deck);
  const title = deck.props["title"];
  if (title !== undefined && typeof title !== "string") {
    throw new BuildError([
      diagnostic("deck", "invalid-title", "Deck.title must be a string"),
    ]);
  }
  if (typeof title === "string") {
    try {
      assertXml10Characters(title);
    } catch {
      throw new BuildError([
        diagnostic(
          "deck",
          "invalid-title",
          "Deck.title contains characters not allowed in XML 1.0",
        ),
      ]);
    }
  }
  let theme: ThemeContext;
  try {
    theme = createThemeContext(
      deck.props["theme"] as ThemeInput | undefined,
      deck.props["style"] as Record<string, unknown> | undefined,
    );
  } catch (error) {
    if (!(error instanceof CapabilityError)) throw error;
    throw new BuildError([diagnostic("deck", error.code, error.message)]);
  }
  const slideIndexById = new Map(
    slides.map((slide) => [slide.id, slide.index] as const),
  );
  const document: DocumentResolveState = {
    usedExtensions: new Set(sections.length > 0 ? ["p14"] : []),
    warnings: [],
    resourceClaims: new Map(),
    morphNames: new Map(),
    nextMorphName: 1,
    reservedObjectNames: collectExplicitObjectNames(
      slides,
      modules.declarations,
    ),
    slideIndexById,
  };

  const resolvedSlides = slides.map((parsedSlide) =>
    resolveSlideDeclaration(
      parsedSlide,
      modules,
      size,
      document,
      options,
      theme,
    ),
  );

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
