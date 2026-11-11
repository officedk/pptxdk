import { createHash } from "node:crypto";

import { strToU8, zipSync } from "fflate";

import type {
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
} from "#/declaration.ts";
import { CapabilityError, BuildError } from "#/diagnostics.ts";
import { LayoutPass, type LayoutPassResult } from "#/layout.ts";
import { getMorphKeyLabel, isMorphKey, type MorphKey } from "#/morph.ts";
import {
  assertXml10Characters,
  pptxParts,
  pxToEmu,
  type PptxPackagePart,
  type PptxPackageRelationshipSet,
  type PptxPackageResource,
  type PptxRelationship,
} from "#/ooxml.ts";
import { createBuildQuery, type QueryContentNode } from "#/query.ts";
import {
  createThemeContext,
  textStyleFrom,
  type ThemeContext,
  type ThemeInput,
  type ThemeScope,
} from "#/theme.ts";
import {
  formatBaseline,
  targets,
  type BuildOptions,
  type BuildReport,
  type BuildResult,
  type CapabilityChild,
  type CapabilityElement,
  type CapabilityChildren,
  type CapabilityInput,
  type CapabilityPackagePart,
  type CapabilityPresentationContribution,
  type CapabilityContribution,
  type CapabilityContext,
  type CapabilityLayoutContract,
  type CapabilityModule,
  type CapabilityParseResult,
  type RegisteredCapabilityModule,
  type CapabilityResource,
  type CapabilityResourceClaim,
  type Diagnostic,
  type Extension,
  type LayoutChildReference,
  type LayoutObservation,
  type NativeElementKind,
  type ObjectCapabilityModule,
  type ObjectObservation,
  type NativeLink,
  type ResourceObservation,
  type SlideCapabilityModule,
  type SlideObservation,
  type SlideCapabilitySlot,
  type Target,
} from "#/types.ts";

type SlideBuild = {
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden: boolean;
  readonly advanceAfter?: number;
  readonly objects: readonly ObjectObservation[];
  readonly layouts: readonly LayoutObservation[];
  readonly content: readonly QueryContentNode[];
  readonly resources: readonly (ResourceObservation & {
    readonly objectId: number;
  })[];
  readonly shapeXml: string;
  readonly backgroundXml: string;
  readonly slideXml: string;
  readonly capabilities: Readonly<Record<string, unknown>>;
  readonly relationships: readonly PptxRelationship[];
  readonly morphObjects: ReadonlyMap<MorphKey, MorphObject>;
  readonly linkedSlideIndexes: ReadonlySet<number>;
};

type ResolvedSlide = {
  readonly node: DeclarationNode;
  readonly id: string;
  readonly authorId?: string;
  readonly index: number;
  readonly section?: string;
  readonly hidden?: boolean;
  readonly advanceAfter?: number;
};

type ResolvedSection = {
  readonly name: string;
  readonly slideIndexes: readonly number[];
};

type ResolvedDocument = {
  readonly slides: readonly ResolvedSlide[];
  readonly sections: readonly ResolvedSection[];
};

type MorphObject = {
  readonly kind: NativeElementKind;
  readonly name: string;
  readonly path: string;
};

type CompiledLink = {
  readonly nativeLink: NativeLink;
  readonly summary: string;
};

type StoredPresentationItem = {
  readonly identity: string;
  readonly order: number;
  readonly preview: string;
  readonly content: string;
  readonly partPath: string;
  readonly relationshipId: string;
  readonly relationshipType: string;
  readonly contentType: string;
  readonly path: string;
  readonly resourceOwners: Map<SlideBuildState, Set<number>>;
  readonly bytes?: Uint8Array;
  readonly resourceClaim?: Pick<CapabilityResourceClaim, "kind" | "key">;
  readonly resolveBytes?: (characters: string) => Uint8Array;
};

type DocumentBuildState = {
  readonly usedExtensions: Set<Extension>;
  readonly warnings: Diagnostic[];
  readonly resourceClaims: Map<
    string,
    { identity: string; path: string; characters: string }
  >;
  readonly resources: Map<string, PptxPackageResource>;
  readonly resourceTypes: Map<string, string>;
  readonly packageParts: Map<string, PptxPackagePart>;
  readonly packageRelationships: Map<string, PptxRelationship[]>;
  readonly presentationAttributes: Map<string, string>;
  readonly presentationPartPaths: Set<string>;
  readonly presentationCollections: Map<
    string,
    {
      readonly tagName: string;
      readonly records: Map<
        string,
        {
          readonly identity: string;
          readonly tagName: string;
          readonly content: string;
          readonly items: Map<string, StoredPresentationItem>;
        }
      >;
    }
  >;
  nextPackagePartId: number;
  nextPresentationRelationshipId: number;
  readonly morphNames: Map<MorphKey, string>;
  nextMorphName: number;
  readonly slideIndexById: ReadonlyMap<string, number>;
  nextTextFieldId: number;
};

type SlideBuildState = {
  readonly index: number;
  nextId: number;
  nextNodeOrder: number;
  readonly diagnostics: Diagnostic[];
  readonly relationships: PptxRelationship[];
  readonly relationshipIds: Map<string, string>;
  readonly layouts: LayoutObservation[];
  readonly resourceObservations: (ResourceObservation & { objectId: number })[];
  readonly contentChildrenByObjectId: Map<number, readonly QueryContentNode[]>;
  readonly contributions: CapabilityContribution[];
  readonly morphObjects: Map<MorphKey, MorphObject>;
  readonly linkedSlideIndexes: Set<number>;
};

type BuildContext = {
  readonly document: DocumentBuildState;
  readonly slide: SlideBuildState;
  readonly modules: ReadonlyMap<string, RegisteredCapabilityModule>;
  readonly target: Target;
  readonly dimensions: Dimensions;
  readonly options: BuildOptions;
  readonly theme: ThemeContext;
  readonly layout: LayoutPassResult;
  readonly slideNumber: number;
};

type CompiledChildren = CapabilityChildren & {
  readonly references: readonly LayoutChildReference[];
  readonly content: readonly QueryContentNode[];
  readonly paintLayers: readonly PaintLayer[];
};

type PaintLayer = {
  readonly sourceOrder: readonly number[];
  readonly zIndex: number;
  readonly xml: string;
  readonly objects: readonly ObjectObservation[];
  readonly reference: LayoutChildReference;
};

type Dimensions = { readonly width: number; readonly height: number };
const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;

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

function ownKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
): string | undefined {
  return Object.keys(value).find((key) => !allowed.includes(key));
}

function isDeclarationNode(child: DeclarationChild): child is DeclarationNode {
  return "props" in child;
}

function isSlideCapability(
  module: RegisteredCapabilityModule,
): module is SlideCapabilityModule {
  return "compileSlide" in module;
}

function isObjectCapability(
  module: RegisteredCapabilityModule,
): module is ObjectCapabilityModule {
  return "compile" in module;
}

function nonWhitespaceText(children: readonly DeclarationChild[]): boolean {
  return children.some(
    (child) => !isDeclarationNode(child) && child.text.trim() !== "",
  );
}

function childNodes(children: readonly DeclarationChild[]): DeclarationNode[] {
  return children.filter(isDeclarationNode);
}

function valueIdentity(props: Readonly<Record<string, unknown>>): unknown {
  const testId = props["testId"];
  return testId === undefined ? props["id"] : testId;
}

function componentLabel(node: DeclarationNode, index: number): string {
  const type = node.type.startsWith("rp:") ? node.type.slice(3) : node.type;
  const identity = valueIdentity(node.props);
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${type}[${label}]`;
}

function capabilityChild(child: DeclarationChild): CapabilityChild {
  if (!isDeclarationNode(child)) return { type: "#text", text: child.text };
  return {
    type: child.type,
    props: child.props,
    children: child.children.map(capabilityChild),
  };
}

function capabilityLabel(child: CapabilityChild, index: number): string {
  if (!("props" in child)) return `text[${index + 1}]`;
  const type = child.type.startsWith("rp:") ? child.type.slice(3) : child.type;
  const identity = valueIdentity(child.props);
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${type}[${label}]`;
}

function nativeObjectName(key: MorphKey, document: DocumentBuildState): string {
  let name = document.morphNames.get(key);
  if (name === undefined) {
    name = `!!react-pptx-morph-${document.nextMorphName++}`;
    document.morphNames.set(key, name);
  }
  return name;
}

function usesMorphTransition(node: DeclarationNode): boolean {
  const transition = node.props["transition"];
  return (
    typeof transition === "object" &&
    transition !== null &&
    !Array.isArray(transition) &&
    (transition as Record<string, unknown>)["type"] === "transition" &&
    (transition as Record<string, unknown>)["kind"] === "morph"
  );
}

function validateTarget(target: Target | undefined): Target {
  const value = target ?? targets.transitional;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new BuildError([
      diagnostic("build.target", "invalid-target", "Target must be an object"),
    ]);
  }
  const record = value as unknown as Record<string, unknown>;
  const unknown = ownKeys(record, ["name", "format", "extensions"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "unknown-parameter",
        `Target contains unknown parameter ${unknown}`,
      ),
    ]);
  }
  if (typeof record["name"] !== "string" || record["name"].trim() === "") {
    throw new BuildError([
      diagnostic(
        "build.target",
        "invalid-target",
        "Target requires a non-empty name",
      ),
    ]);
  }
  if (record["format"] !== formatBaseline) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "unsupported-format",
        `This build only emits ${formatBaseline}`,
      ),
    ]);
  }
  if (!Array.isArray(record["extensions"])) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "invalid-target",
        "Target extensions must be an array",
      ),
    ]);
  }
  const extensions: Extension[] = [];
  for (const extension of record["extensions"]) {
    if (extension !== "p14" && extension !== "p159") {
      throw new BuildError([
        diagnostic(
          "build.target",
          "unsupported-extension",
          `Unsupported target extension ${String(extension)}`,
        ),
      ]);
    }
    if (extensions.includes(extension)) {
      throw new BuildError([
        diagnostic(
          "build.target",
          "duplicate-extension",
          `Duplicate target extension declaration ${extension}`,
        ),
      ]);
    }
    extensions.push(extension);
  }
  return Object.freeze({
    name: record["name"],
    format: formatBaseline,
    extensions: Object.freeze(extensions),
  });
}

export function validateOptions(options: BuildOptions): Target {
  if (
    typeof options !== "object" ||
    options === null ||
    Array.isArray(options)
  ) {
    throw new BuildError([
      diagnostic("build", "invalid-options", "Build options must be an object"),
    ]);
  }
  const unknown = ownKeys(options, ["target", "signal", "embedFonts"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic(
        "build",
        "unknown-parameter",
        `Build options contain unknown parameter ${unknown}`,
      ),
    ]);
  }
  if (
    options.embedFonts !== undefined &&
    options.embedFonts !== "full" &&
    options.embedFonts !== "subset"
  ) {
    throw new BuildError([
      diagnostic(
        "build.embedFonts",
        "unsupported-font-embedding",
        "embedFonts currently only supports 'full' or 'subset'",
      ),
    ]);
  }
  return validateTarget(options.target);
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

function deckDimensions(node: DeclarationNode): Dimensions {
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

function readAuthorId(value: unknown, path: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new BuildError([
      diagnostic(path, "invalid-author-id", "id must be a non-empty string"),
    ]);
  }
  return value;
}

/**
 * Validates new author ids on one slide using the same scopes the author query traverses.
 *
 * A node's own id belongs to its parent scope; its descendants belong to a new scope only when the
 * node declares an id. Anonymous layout nodes therefore stay transparent, while a page remains the
 * root scope even for legacy slides without a new id.
 */
function validateAuthorIdScopes(
  node: DeclarationNode,
  path: string,
  scope: Set<string>,
): void {
  node.children.forEach((child, index) => {
    if (!isDeclarationNode(child)) return;
    const childPath = `${path}/${componentLabel(child, index)}`;
    const id = readAuthorId(child.props["id"], childPath);
    if (id !== undefined) {
      if (scope.has(id)) {
        throw new BuildError([
          diagnostic(
            childPath,
            "duplicate-author-id",
            `Duplicate author id ${JSON.stringify(id)} in the same scope`,
          ),
        ]);
      }
      scope.add(id);
    }
    validateAuthorIdScopes(
      child,
      childPath,
      id === undefined ? scope : new Set(),
    );
  });
}

function resolveSlides(
  container: DeclarationContainer,
  modules: ReadonlyMap<string, RegisteredCapabilityModule>,
): ResolvedDocument {
  const decks = childNodes(container.children);
  if (decks.length !== 1 || decks[0]?.type !== "rp:deck") {
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
  const sectioned = topLevel.some((node) => node.type === "rp:section");
  const candidates: {
    node: DeclarationNode;
    section?: string;
    sectionIndex?: number;
  }[] = [];
  const sections: { name: string; slideIndexes: number[] }[] = [];
  if (sectioned) {
    for (const [index, node] of topLevel.entries()) {
      if (node.type !== "rp:section") {
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
        if (member.type !== "rp:slide") {
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
      if (node.type !== "rp:slide") {
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
  const ids = new Set<string>();
  const authorIds = new Set<string>();
  const slideProperties = [...modules.values()]
    .filter(isSlideCapability)
    .flatMap((module) => module.slideProps ?? [module.slideProp]);
  const slides = candidates.map(({ node, section, sectionIndex }, index) => {
    const unknown = ownKeys(node.props, [
      "id",
      "slideId",
      "theme",
      "style",
      "hidden",
      ...slideProperties,
    ]);
    const authorId = readAuthorId(node.props["id"], `deck/slide[${index + 1}]`);
    const rawSlideId = node.props["slideId"];
    if (unknown !== undefined) {
      const label =
        authorId ?? (typeof rawSlideId === "string" ? rawSlideId : index + 1);
      throw new BuildError([
        diagnostic(
          `deck/slide[${label}]`,
          "missing-capability",
          `No capability module registered for Slide.${unknown}`,
        ),
      ]);
    }
    if (
      rawSlideId !== undefined &&
      (typeof rawSlideId !== "string" || rawSlideId.trim() === "")
    ) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-slide-id",
          "Slide.slideId must be a non-empty string",
        ),
      ]);
    }
    if (rawSlideId === undefined && authorId === undefined) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-slide-id",
          "Slide must declare id or slideId as a non-empty string",
        ),
      ]);
    }
    if (typeof rawSlideId === "string") {
      try {
        assertXml10Characters(rawSlideId);
      } catch {
        throw new BuildError([
          diagnostic(
            `deck/slide[${index + 1}]`,
            "invalid-xml-character",
            "Slide.slideId contains characters not allowed in XML 1.0",
          ),
        ]);
      }
    }
    const id = typeof rawSlideId === "string" ? rawSlideId : authorId!;
    if (authorId !== undefined) {
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
    }
    if (ids.has(id)) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${id}]`,
          "duplicate-slide-id",
          `Duplicate slide identity "${id}"`,
        ),
      ]);
    }
    ids.add(id);
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
    const advanceAfter = node.props["advanceAfter"];
    if (
      advanceAfter !== undefined &&
      (typeof advanceAfter !== "number" ||
        !Number.isFinite(advanceAfter) ||
        advanceAfter < 0 ||
        advanceAfter > 4_294_967_295)
    ) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${id}]`,
          "invalid-advance-after",
          "Slide.advanceAfter must be a finite millisecond value between 0 and 4294967295",
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
      ...(advanceAfter === undefined ? {} : { advanceAfter }),
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

function addResource(
  resource: CapabilityResource,
  context: Pick<BuildContext, "document" | "slide">,
  objectId?: number,
): string {
  if (
    !(resource.bytes instanceof Uint8Array) ||
    resource.bytes.byteLength === 0
  ) {
    throw new CapabilityError(
      "invalid-resource",
      "Resource must contain non-empty binary content",
    );
  }
  if (
    typeof resource.extension !== "string" ||
    !/^[a-z0-9]+$/iu.test(resource.extension) ||
    typeof resource.contentType !== "string" ||
    resource.contentType.trim() === "" ||
    typeof resource.relationshipType !== "string" ||
    resource.relationshipType.trim() === ""
  ) {
    throw new CapabilityError(
      "invalid-resource",
      "Resource format or relationship type is invalid",
    );
  }
  const extension = resource.extension.toLowerCase();
  if (extension === "xml" || extension === "rels") {
    throw new CapabilityError(
      "invalid-resource",
      `Resource extension .${extension} conflicts with core PPTX parts`,
    );
  }
  const knownType = context.document.resourceTypes.get(extension);
  if (knownType !== undefined && knownType !== resource.contentType) {
    throw new CapabilityError(
      "conflicting-resource-type",
      `Resource extension .${extension} corresponds to multiple content types`,
    );
  }
  context.document.resourceTypes.set(extension, resource.contentType);
  try {
    assertXml10Characters(resource.contentType);
    assertXml10Characters(resource.relationshipType);
  } catch {
    throw new CapabilityError(
      "invalid-resource",
      "Resource relationship contains characters not allowed in XML",
    );
  }

  const digest = createHash("sha256").update(resource.bytes).digest("hex");
  const assetKey = `${extension}\u0000${resource.contentType}\u0000${digest}`;
  let asset = context.document.resources.get(assetKey);
  if (asset === undefined) {
    asset = {
      name: `resource${context.document.resources.size + 1}.${extension}`,
      bytes: new Uint8Array(resource.bytes),
      extension,
      contentType: resource.contentType,
    };
    context.document.resources.set(assetKey, asset);
  }

  if (objectId !== undefined) {
    recordResourceObservation(context.slide, objectId, {
      part: `ppt/media/${asset.name}`,
      contentType: asset.contentType,
      byteLength: asset.bytes.byteLength,
    });
  }

  const relationshipKey = `${assetKey}\u0000${resource.relationshipType}`;
  const existingRelationship =
    context.slide.relationshipIds.get(relationshipKey);
  if (existingRelationship !== undefined) return existingRelationship;
  const id = `rId${context.slide.relationships.length + 2}`;
  context.slide.relationshipIds.set(relationshipKey, id);
  context.slide.relationships.push({
    id,
    type: resource.relationshipType,
    target: `../media/${asset.name}`,
  });
  return id;
}

function relationshipTarget(sourcePath: string, targetPath: string): string {
  const sourceSegments = sourcePath.split("/").slice(0, -1);
  const targetSegments = targetPath.split("/");
  while (
    sourceSegments.length > 0 &&
    targetSegments.length > 0 &&
    sourceSegments[0] === targetSegments[0]
  ) {
    sourceSegments.shift();
    targetSegments.shift();
  }
  return [...sourceSegments.map(() => ".."), ...targetSegments].join("/");
}

function recordResourceObservation(
  slide: SlideBuildState,
  objectId: number,
  resource: ResourceObservation,
): void {
  if (
    slide.resourceObservations.some(
      (existing) =>
        existing.objectId === objectId && existing.part === resource.part,
    )
  ) {
    return;
  }
  slide.resourceObservations.push({ objectId, ...resource });
}

function packagePartPath(template: string, sequence: number): string {
  if (
    (template.match(/\{n\}/gu) ?? []).length !== 1 ||
    !/^(?:[a-z0-9_-]+\/)*[a-z0-9_-]+\.[a-z0-9]+$/iu.test(
      template.replace("{n}", String(sequence)),
    )
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part path must be a safe package-relative path containing a single part number",
    );
  }
  return template.replace("{n}", String(sequence));
}

function addPackagePart(
  packagePart: CapabilityPackagePart,
  context: BuildContext,
  objectId?: number,
): string {
  const { part, relatedParts } = packagePart;
  if (
    typeof part.pathTemplate !== "string" ||
    typeof part.contentType !== "string" ||
    part.contentType.trim() === "" ||
    typeof packagePart.relationshipType !== "string" ||
    packagePart.relationshipType.trim() === "" ||
    typeof packagePart.content !== "function" ||
    !Array.isArray(relatedParts) ||
    relatedParts.some(
      (related) =>
        typeof related.pathTemplate !== "string" ||
        typeof related.contentType !== "string" ||
        related.contentType.trim() === "" ||
        typeof related.relationshipType !== "string" ||
        related.relationshipType.trim() === "" ||
        !(related.bytes instanceof Uint8Array) ||
        related.bytes.byteLength === 0,
    )
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part or its related parts format and relationships are invalid",
    );
  }
  try {
    for (const value of [part.contentType, packagePart.relationshipType]) {
      assertXml10Characters(value);
    }
    for (const related of relatedParts) {
      for (const value of [related.contentType, related.relationshipType]) {
        assertXml10Characters(value);
      }
    }
  } catch {
    throw new CapabilityError(
      "invalid-part",
      "Native part content type or relationship contains characters not allowed in XML 1.0",
    );
  }

  const sequence = context.document.nextPackagePartId++;
  const partPath = packagePartPath(part.pathTemplate, sequence);
  const related = relatedParts.map((item, index) => ({
    path: packagePartPath(item.pathTemplate, sequence),
    relationship: {
      id: `rId${index + 1}`,
      type: item.relationshipType,
      target: "",
    },
    descriptor: item,
  }));
  for (const item of related) {
    item.relationship.target = relationshipTarget(partPath, item.path);
  }
  const relatedIds = related.map(({ relationship }) => relationship.id);
  const content = packagePart.content(relatedIds);
  if (
    !(typeof content === "string"
      ? content.length > 0
      : content instanceof Uint8Array && content.byteLength > 0)
  ) {
    throw new CapabilityError(
      "invalid-part",
      "Native part must contain content",
    );
  }
  const relatedPaths = new Set(related.map(({ path }) => path));
  if (
    context.document.packageParts.has(partPath) ||
    relatedPaths.size !== related.length ||
    relatedPaths.has(partPath) ||
    related.some(({ path }) => context.document.packageParts.has(path))
  ) {
    throw new CapabilityError("duplicate-part", "Duplicate PPTX part path");
  }
  if (objectId !== undefined) {
    recordResourceObservation(context.slide, objectId, {
      part: partPath,
      contentType: part.contentType,
      byteLength:
        typeof content === "string"
          ? strToU8(content).byteLength
          : content.byteLength,
    });
    for (const item of related) {
      recordResourceObservation(context.slide, objectId, {
        part: item.path,
        contentType: item.descriptor.contentType,
        byteLength: item.descriptor.bytes.byteLength,
      });
    }
  }
  context.document.packageParts.set(partPath, {
    path: partPath,
    content,
    contentType: part.contentType,
  });
  for (const item of related) {
    context.document.packageParts.set(item.path, {
      path: item.path,
      content: new Uint8Array(item.descriptor.bytes),
      contentType: item.descriptor.contentType,
    });
  }
  if (related.length > 0) {
    context.document.packageRelationships.set(
      partPath,
      related.map(({ relationship }) => relationship),
    );
  }

  const relationshipId = `rId${context.slide.relationships.length + 2}`;
  context.slide.relationships.push({
    id: relationshipId,
    type: packagePart.relationshipType,
    target: relationshipTarget("ppt/slides/slide.xml", partPath),
  });
  return relationshipId;
}

function xmlName(value: string): boolean {
  return /^(?:[A-Za-z_][A-Za-z0-9_.-]*:)?[A-Za-z_][A-Za-z0-9_.-]*$/u.test(
    value,
  );
}

function addPresentationContribution(
  contribution: CapabilityPresentationContribution,
  context: BuildContext,
  path: string,
  objectId: number,
): void {
  const { collection, record, item } = contribution;
  if (
    !xmlName(collection) ||
    typeof record.key !== "string" ||
    record.key === "" ||
    typeof record.identity !== "string" ||
    typeof record.tagName !== "string" ||
    !xmlName(record.tagName) ||
    typeof record.content !== "string" ||
    typeof item.key !== "string" ||
    item.key === "" ||
    typeof item.identity !== "string" ||
    !Number.isInteger(item.order) ||
    item.order < 0 ||
    typeof item.relationshipType !== "string" ||
    item.relationshipType.trim() === "" ||
    typeof item.pathTemplate !== "string" ||
    typeof item.contentType !== "string" ||
    item.contentType.trim() === "" ||
    (item.bytes !== undefined &&
      (!(item.bytes instanceof Uint8Array) || item.bytes.byteLength === 0)) ||
    (item.bytes === undefined && item.resolveBytes === undefined) ||
    (item.resourceClaim === undefined) !== (item.resolveBytes === undefined) ||
    (item.resourceClaim !== undefined &&
      (typeof item.resourceClaim.kind !== "string" ||
        item.resourceClaim.kind === "" ||
        typeof item.resourceClaim.key !== "string" ||
        item.resourceClaim.key === "")) ||
    (item.resolveBytes !== undefined &&
      typeof item.resolveBytes !== "function") ||
    typeof item.content !== "function"
  ) {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution record, relationship, or resource is invalid",
    );
  }
  try {
    assertXml10Characters(collection);
    assertXml10Characters(record.content);
    assertXml10Characters(item.relationshipType);
    assertXml10Characters(item.contentType);
    for (const [name, value] of Object.entries(
      contribution.presentationAttributes ?? {},
    )) {
      if (!xmlName(name) || typeof value !== "string") {
        throw new Error("invalid presentation attribute");
      }
      assertXml10Characters(value);
    }
  } catch {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution contains invalid XML name or characters",
    );
  }

  for (const [name, value] of Object.entries(
    contribution.presentationAttributes ?? {},
  )) {
    const existing = context.document.presentationAttributes.get(name);
    if (existing !== undefined && existing !== value) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "conflicting-presentation-attribute",
          `Conflicting contribution value for presentation attribute ${name}`,
        ),
      );
    } else {
      context.document.presentationAttributes.set(name, value);
    }
  }

  let targetCollection =
    context.document.presentationCollections.get(collection);
  if (targetCollection === undefined) {
    targetCollection = { tagName: collection, records: new Map() };
    context.document.presentationCollections.set(collection, targetCollection);
  }
  let targetRecord = targetCollection.records.get(record.key);
  if (targetRecord === undefined) {
    targetRecord = {
      identity: record.identity,
      tagName: record.tagName,
      content: record.content,
      items: new Map(),
    };
    targetCollection.records.set(record.key, targetRecord);
  } else if (
    targetRecord.identity !== record.identity ||
    targetRecord.tagName !== record.tagName ||
    targetRecord.content !== record.content
  ) {
    context.slide.diagnostics.push(
      diagnostic(
        path,
        "conflicting-presentation-record",
        `Conflicting declaration for presentation-level record ${record.key}`,
      ),
    );
    return;
  }

  let preview: string;
  try {
    preview = item.content("rId0");
    assertXml10Characters(preview);
  } catch {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution failed to generate XML content",
    );
  }
  if (preview.trim() === "") {
    throw new CapabilityError(
      "invalid-presentation-contribution",
      "Presentation-level contribution must generate content",
    );
  }
  const existingItem = targetRecord.items.get(item.key);
  if (existingItem !== undefined) {
    const conflicts =
      existingItem.identity !== item.identity ||
      existingItem.order !== item.order ||
      existingItem.preview !== preview ||
      existingItem.resourceClaim?.kind !== item.resourceClaim?.kind ||
      existingItem.resourceClaim?.key !== item.resourceClaim?.key;
    if (conflicts) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "conflicting-presentation-item",
          `Presentation-level record ${record.key} item ${item.key} uses conflicting resources`,
        ),
      );
    } else {
      const objectIds =
        existingItem.resourceOwners.get(context.slide) ?? new Set<number>();
      objectIds.add(objectId);
      existingItem.resourceOwners.set(context.slide, objectIds);
    }
    return;
  }

  const sequence = context.document.nextPackagePartId++;
  const partPath = packagePartPath(item.pathTemplate, sequence);
  if (
    context.document.packageParts.has(partPath) ||
    context.document.presentationPartPaths.has(partPath)
  ) {
    throw new CapabilityError("duplicate-part", "Duplicate PPTX part path");
  }
  context.document.presentationPartPaths.add(partPath);
  const relationshipId = `rId${context.document.nextPresentationRelationshipId++}`;
  const content = item.content(relationshipId);
  targetRecord.items.set(item.key, {
    identity: item.identity,
    order: item.order,
    preview,
    content,
    partPath,
    relationshipId,
    relationshipType: item.relationshipType,
    contentType: item.contentType,
    path,
    resourceOwners: new Map([[context.slide, new Set([objectId])]]),
    ...(item.bytes === undefined ? {} : { bytes: item.bytes }),
    ...(item.resourceClaim === undefined
      ? {}
      : { resourceClaim: item.resourceClaim }),
    ...(item.resolveBytes === undefined
      ? {}
      : { resolveBytes: item.resolveBytes }),
  });
}

function finalizePresentationContributions(document: DocumentBuildState): void {
  for (const collection of document.presentationCollections.values()) {
    for (const record of collection.records.values()) {
      for (const item of record.items.values()) {
        let bytes = item.bytes;
        if (item.resourceClaim !== undefined) {
          const claimKey = `${item.resourceClaim.kind}\u0000${item.resourceClaim.key}`;
          const claim = document.resourceClaims.get(claimKey);
          if (claim === undefined || claim.identity !== item.identity) {
            throw new BuildError([
              diagnostic(
                item.path,
                "missing-resource-claim",
                `Presentation resource ${item.resourceClaim.key} is missing matching character claims`,
              ),
            ]);
          }
          try {
            bytes = item.resolveBytes!(claim.characters);
          } catch (error) {
            if (error instanceof CapabilityError) {
              throw new BuildError([
                diagnostic(error.path ?? item.path, error.code, error.message),
              ]);
            }
            throw new BuildError([
              diagnostic(
                item.path,
                "resource-contribution-failed",
                error instanceof Error ? error.message : String(error),
              ),
            ]);
          }
        }
        if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) {
          throw new BuildError([
            diagnostic(
              item.path,
              "invalid-presentation-contribution",
              "Presentation-level contribution did not produce valid resource bytes",
            ),
          ]);
        }
        document.packageParts.set(item.partPath, {
          path: item.partPath,
          content: new Uint8Array(bytes),
          contentType: item.contentType,
        });
        const relationships =
          document.packageRelationships.get("ppt/presentation.xml") ?? [];
        relationships.push({
          id: item.relationshipId,
          type: item.relationshipType,
          target: relationshipTarget("ppt/presentation.xml", item.partPath),
        });
        document.packageRelationships.set(
          "ppt/presentation.xml",
          relationships,
        );
        for (const [slide, objectIds] of item.resourceOwners) {
          for (const objectId of objectIds) {
            recordResourceObservation(slide, objectId, {
              part: item.partPath,
              contentType: item.contentType,
              byteLength: bytes.byteLength,
            });
          }
        }
      }
    }
  }
}

function presentationContent(document: DocumentBuildState): string {
  return [...document.presentationCollections.values()]
    .map(
      ({ tagName, records }) =>
        `<${tagName}>${[...records.values()]
          .map(
            (record) =>
              `<${record.tagName}>${record.content}${[...record.items.values()]
                .sort((left, right) => left.order - right.order)
                .map(({ content }) => content)
                .join("")}</${record.tagName}>`,
          )
          .join("")}</${tagName}>`,
    )
    .join("");
}

function capabilityContext(
  context: BuildContext,
  child: CapabilityElement,
  id: number,
  path: string,
  layoutBox: CapabilityContext["layoutBox"],
  theme: ThemeContext,
  insideGroup: boolean,
  morphName?: string,
  nativeLink?: NativeLink,
): CapabilityContext {
  return {
    id,
    path,
    options: context.options,
    slideNumber: context.slideNumber,
    createTextFieldId() {
      const serial = context.document.nextTextFieldId++;
      const suffix = serial.toString(16).padStart(12, "0").toUpperCase();
      return `{00000000-0000-4000-8000-${suffix}}`;
    },
    ...(morphName === undefined ? {} : { morphName }),
    ...(nativeLink === undefined ? {} : { nativeLink }),
    theme,
    ...(layoutBox === undefined ? {} : { layoutBox }),
    compileChildren(nested) {
      const compiled = compileChildren(
        nested,
        context,
        path,
        insideGroup || child.type === "rp:group",
      );
      if (child.type === "rp:group") {
        context.slide.contentChildrenByObjectId.set(id, compiled.content);
      }
      return { xml: compiled.xml, objects: compiled.objects };
    },
    registerResource(resource) {
      return addResource(resource, context, id);
    },
    registerPackagePart(part) {
      return addPackagePart(part, context, id);
    },
  };
}

function registerLink(
  value: unknown,
  context: BuildContext,
  path: string,
): CompiledLink | undefined {
  if (value === undefined) return undefined;
  const reportInvalid = (message: string) => {
    context.slide.diagnostics.push(diagnostic(path, "invalid-link", message));
  };
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    reportInvalid("link must specify slideId, url, or home action");
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
  if (kind === "slideId" && typeof target === "string") {
    const targetIndex = context.document.slideIndexById.get(target);
    if (targetIndex === undefined) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-link-target",
          `No target slide with slideId ${JSON.stringify(target)}`,
        ),
      );
      return undefined;
    }
    const relationshipKey = `link:slide:${target}`;
    let relationshipId = context.slide.relationshipIds.get(relationshipKey);
    if (relationshipId === undefined) {
      relationshipId = `rId${context.slide.relationships.length + 2}`;
      context.slide.relationshipIds.set(relationshipKey, relationshipId);
      context.slide.relationships.push({
        id: relationshipId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
        target: `slide${targetIndex}.xml`,
      });
    }
    if (targetIndex !== context.slide.index)
      context.slide.linkedSlideIndexes.add(targetIndex);
    return {
      nativeLink: {
        relationshipId,
        action: "ppaction://hlinksldjump",
      },
      summary: `slide:${target}`,
    };
  }
  if (kind === "url" && typeof target === "string") {
    let url: URL;
    try {
      url = new URL(target);
    } catch {
      reportInvalid("link.url must be an absolute URL");
      return undefined;
    }
    const relationshipKey = `link:url:${url.href}`;
    let relationshipId = context.slide.relationshipIds.get(relationshipKey);
    if (relationshipId === undefined) {
      relationshipId = `rId${context.slide.relationships.length + 2}`;
      context.slide.relationshipIds.set(relationshipKey, relationshipId);
      context.slide.relationships.push({
        id: relationshipId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        target: url.href,
        targetMode: "External",
      });
    }
    return { nativeLink: { relationshipId }, summary: `url:${url.href}` };
  }
  if (kind === "action" && target === "home") {
    return {
      nativeLink: {
        relationshipId: "",
        action: "ppaction://hlinkshowjump?jump=firstslide",
      },
      summary: "home",
    };
  }
  reportInvalid("link supports { slideId }, { url }, or { action: 'home' }");
  return undefined;
}

function compileChildren(
  children: readonly CapabilityChild[],
  context: BuildContext,
  parentPath: string,
  insideGroup = false,
): CompiledChildren {
  const painted: PaintLayer[] = [];
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
    const requestedMorphKey = child.props["morphKey"];
    if (child.type === "rp:box") {
      const box = context.layout.boxes.get(child);
      if (box === undefined) {
        context.slide.diagnostics.push(
          diagnostic(
            path,
            "layout-unavailable",
            "Box is missing layout result",
          ),
        );
        return;
      }
      const id = context.slide.nextId++;
      const module = context.modules.get("rp:box");
      const parsedResult: CapabilityParseResult | undefined =
        context.layout.parsed?.get(child);
      const parsedBox = parsedResult?.parsed as
        | { hasVisibleStyle?: boolean }
        | undefined;
      const hasNativeBox = Boolean(parsedBox?.hasVisibleStyle);
      const rawLink = child.props["link"];
      const boxLink =
        rawLink === undefined
          ? undefined
          : hasNativeBox
            ? registerLink(rawLink, context, path)
            : (context.slide.diagnostics.push(
                diagnostic(
                  path,
                  "link-needs-object",
                  "Transparent Box does not produce clickable objects; attach link to a visible Box or Group",
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
        } else if (!hasNativeBox) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "morph-needs-object",
              "Only visible Box generating native background shapes can participate in Morph",
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
      let boxXml = "";
      let boxObservation: ObjectObservation | undefined;
      if (module !== undefined && isObjectCapability(module)) {
        const objectId = context.slide.nextId++;
        try {
          const compileInput = parsedResult!.parsed;
          const output = module.compile(
            compileInput,
            capabilityContext(
              context,
              child,
              objectId,
              path,
              box,
              context.layout.themes.get(child) ?? context.theme,
              insideGroup,
              morphName,
              boxLink?.nativeLink,
            ),
          );
          boxXml = output.xml;
          if (output.xml !== "") {
            boxObservation = {
              ...output.observation,
              id: objectId,
              ...(boxLink === undefined ? {} : { link: boxLink.summary }),
            };
            if (morphKey !== undefined && morphName !== undefined) {
              const nativeKind = parsedResult?.contract?.nativeKind ?? "sp";
              if (output.observation.name !== morphName) {
                context.slide.diagnostics.push(
                  diagnostic(
                    path,
                    "unsupported-morph-object",
                    "Box did not write morphKey to a supported native background shape name",
                  ),
                );
              } else if (!context.slide.morphObjects.has(morphKey)) {
                context.slide.morphObjects.set(morphKey, {
                  kind: nativeKind,
                  name: morphName,
                  path,
                });
              }
            }
          } else if (morphKey !== undefined) {
            context.slide.diagnostics.push(
              diagnostic(
                path,
                "morph-needs-object",
                "Only visible Box generating native background shapes can participate in Morph",
              ),
            );
          }
          if (boxObservation !== undefined) {
            for (const annotation of output.annotations ?? []) {
              context.slide.contributions.push({
                ...annotation,
                path,
                targetId: boxObservation.id,
                order,
              });
            }
          }
          for (const warning of output.warnings ?? []) {
            context.document.warnings.push({
              severity: "warning",
              path:
                warning.path === undefined ? path : `${path}/${warning.path}`,
              code: warning.code,
              message: warning.message,
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
      } else if (hasNativeBox) {
        context.slide.diagnostics.push(
          diagnostic(
            path,
            "missing-capability",
            "Visible Box requires registered rp:box shape capability",
          ),
        );
      }
      const compiled = compileChildren(
        child.children,
        context,
        path,
        insideGroup,
      );
      const testId = child.props["testId"];
      context.slide.layouts.push({
        id,
        ...(authorId === undefined ? {} : { authorId }),
        ...(typeof testId === "string" ? { testId } : {}),
        ...(typeof child.props["part"] === "string"
          ? { part: child.props["part"] }
          : {}),
        box,
        children: [
          ...(boxObservation === undefined
            ? []
            : [{ kind: "object" as const, id: boxObservation.id }]),
          ...compiled.references,
        ],
      });
      content.push({
        ...(authorId === undefined ? {} : { authorId }),
        ...(typeof testId === "string" ? { testId } : {}),
        ...(typeof child.props["part"] === "string"
          ? { part: child.props["part"] }
          : {}),
        reference: {
          kind: "layout",
          id,
          objectIds: boxObservation === undefined ? [] : [boxObservation.id],
        },
        children: compiled.content,
      });
      painted.push({
        sourceOrder: [index],
        zIndex: paintOrder(parsedResult?.contract),
        xml: boxXml + compiled.xml,
        objects: [
          ...(boxObservation === undefined ? [] : [boxObservation]),
          ...compiled.objects,
        ],
        reference: { kind: "layout", id },
      });
      return;
    }
    const module = context.modules.get(child.type);
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
      } else if (
        "compileDeclaration" in module ||
        !isObjectCapability(module)
      ) {
        context.slide.diagnostics.push(
          diagnostic(
            path,
            "morph-needs-object",
            "morphKey can only be used with capabilities that produce native objects",
          ),
        );
      } else if (child.props["name"] !== undefined) {
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
    if ("compileDeclaration" in module) {
      try {
        for (const annotation of module.compileDeclaration(child, { path })) {
          context.slide.contributions.push({ ...annotation, path, order });
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
      const compiled = compileChildren(
        child.children,
        context,
        path,
        insideGroup,
      );
      content.push(...compiled.content);
      // Transparent declarations do not occupy paint order; children carry their own order and sort with siblings.
      painted.push(
        ...compiled.paintLayers.map((layer) => ({
          ...layer,
          sourceOrder: [index, ...layer.sourceOrder],
        })),
      );
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
    const id = context.slide.nextId++;
    const input: CapabilityInput = child;
    const layoutBox = context.layout.boxes.get(child);
    const theme = context.layout.themes.get(child) ?? context.theme;
    const compiledLink = registerLink(child.props["link"], context, path);
    try {
      const parsedResult: CapabilityParseResult =
        context.layout.parsed?.get(child) ??
        module.parse(input, {
          path,
          theme,
          options: context.options,
          slideNumber: context.slideNumber,
        });
      const compileInput = parsedResult.parsed;
      const output = module.compile(
        compileInput,
        capabilityContext(
          context,
          child,
          id,
          path,
          layoutBox,
          theme,
          insideGroup,
          morphName,
          compiledLink?.nativeLink,
        ),
      );
      if (morphKey !== undefined && morphName !== undefined) {
        const nativeKind = parsedResult.contract.nativeKind;
        if (output.observation.name !== morphName) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "unsupported-morph-object",
              `${child.type} did not write morphKey to native object name, cannot participate in Morph matching`,
            ),
          );
        } else if (!context.slide.morphObjects.has(morphKey)) {
          context.slide.morphObjects.set(morphKey, {
            kind: nativeKind,
            name: morphName,
            path,
          });
        }
      }
      for (const extension of output.extensions) {
        if (!context.target.extensions.includes(extension)) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "forbidden-mechanism",
              `Target forbids ${extension}`,
            ),
          );
        } else {
          context.document.usedExtensions.add(extension);
        }
      }
      for (const claim of output.resourceClaims ?? []) {
        const claimKey = `${claim.kind}\u0000${claim.key}`;
        const existing = context.document.resourceClaims.get(claimKey);
        if (existing !== undefined && existing.identity !== claim.identity) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "conflicting-resource",
              `Resource ${claim.key} has conflicting definition with ${existing.path}`,
            ),
          );
        } else {
          context.document.resourceClaims.set(claimKey, {
            identity: claim.identity,
            path: existing?.path ?? path,
            characters: (existing?.characters ?? "") + (claim.characters ?? ""),
          });
        }
      }
      for (const contribution of output.presentationContributions ?? []) {
        addPresentationContribution(contribution, context, path, id);
      }
      for (const warning of output.warnings ?? []) {
        context.document.warnings.push({
          severity: "warning",
          path: warning.path === undefined ? path : `${path}/${warning.path}`,
          code: warning.code,
          message: warning.message,
        });
      }
      for (const annotation of output.annotations ?? []) {
        context.slide.contributions.push({
          ...annotation,
          path,
          targetId: id,
          order,
        });
      }
      const observation = {
        ...output.observation,
        id,
        ...(authorId === undefined ? {} : { authorId }),
        ...(compiledLink === undefined ? {} : { link: compiledLink.summary }),
      };
      content.push({
        ...(authorId === undefined ? {} : { authorId }),
        ...(typeof child.props["testId"] === "string"
          ? { testId: child.props["testId"] }
          : {}),
        ...(typeof observation.part === "string"
          ? { part: observation.part }
          : {}),
        reference: { kind: "object", id: observation.id },
        children: [
          ...(context.slide.contentChildrenByObjectId.get(id) ?? []),
          ...tableContentNodes(observation),
        ],
      });
      if (observation.box !== undefined) {
        const box = observation.box;
        if (
          box.x < 0 ||
          box.y < 0 ||
          box.x + box.width > context.dimensions.width ||
          box.y + box.height > context.dimensions.height
        ) {
          context.document.warnings.push({
            severity: "warning",
            path,
            code: "overflow",
            message:
              "Object overflows slide canvas; build retains original declaration",
          });
        }
      }
      painted.push({
        sourceOrder: [index],
        zIndex: paintOrder(parsedResult.contract),
        xml: output.xml,
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
  painted.sort(
    (left, right) =>
      left.zIndex - right.zIndex ||
      compareSourceOrder(left.sourceOrder, right.sourceOrder),
  );
  return {
    xml: painted.map(({ xml }) => xml).join(""),
    objects: painted.flatMap(({ objects }) => objects),
    references: painted.map(({ reference }) => reference),
    content,
    paintLayers: painted,
  };
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

function tableContentNodes(object: ObjectObservation): QueryContentNode[] {
  const parts = object.table?.parts;
  if (parts === undefined) return [];
  const nodeAt = (partIndex: number): QueryContentNode => {
    const part = parts[partIndex];
    if (part === undefined)
      throw new Error("Table query part index does not exist");
    return {
      ...(part.authorId === undefined ? {} : { authorId: part.authorId }),
      ...(part.testId === undefined ? {} : { testId: part.testId }),
      ...(part.part === undefined ? {} : { part: part.part }),
      reference: { kind: "table-part", tableId: object.id, partIndex },
      children: [],
    };
  };
  return parts.flatMap((part, partIndex) =>
    part.kind === "row"
      ? [
          {
            ...nodeAt(partIndex),
            children: parts.flatMap((cell, cellIndex) =>
              cell.kind === "cell" && cell.row === part.row
                ? [nodeAt(cellIndex)]
                : [],
            ),
          },
        ]
      : [],
  );
}

function compileSlide(
  resolved: ResolvedSlide,
  modules: ReadonlyMap<string, RegisteredCapabilityModule>,
  target: Target,
  dimensions: Dimensions,
  document: DocumentBuildState,
  options: BuildOptions,
  documentTheme: ThemeContext,
): SlideBuild {
  const {
    node,
    id,
    authorId,
    index,
    section,
    hidden = false,
    advanceAfter,
  } = resolved;
  const slide: SlideBuildState = {
    index,
    nextId: 2,
    nextNodeOrder: 0,
    diagnostics: [],
    relationships: [],
    relationshipIds: new Map(),
    linkedSlideIndexes: new Set(),
    layouts: [],
    resourceObservations: [],
    contentChildrenByObjectId: new Map(),
    contributions: [],
    morphObjects: new Map(),
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
  const layout = new LayoutPass(modules, dimensions, index, options).run(
    children,
    path,
    theme,
    node.props["style"],
  );
  document.warnings.push(...layout.warnings);
  const compiled = compileChildren(
    children,
    {
      document,
      slide,
      modules,
      target,
      dimensions,
      options,
      theme,
      layout,
      slideNumber: index,
    },
    path,
  );
  const slideXml: {
    readonly slot: SlideCapabilitySlot;
    readonly xml: string;
  }[] = [];
  const capabilities: Record<string, unknown> = {};
  let backgroundXml = "";
  const input: CapabilityInput = {
    type: node.type,
    props: node.props,
    children: node.children.map(capabilityChild),
  };
  for (const module of modules.values()) {
    if (!isSlideCapability(module)) continue;
    const slideProps = module.slideProps ?? [module.slideProp];
    if (
      !module.always &&
      !slideProps.some((property) => Object.hasOwn(node.props, property))
    )
      continue;
    try {
      const output = module.compileSlide(input, {
        path,
        dimensions,
        contributions: slide.contributions,
        theme,
        registerResource(resource) {
          return addResource(resource, { document, slide });
        },
      });
      if (module.slot === "background") backgroundXml += output.xml;
      else slideXml.push({ slot: module.slot, xml: output.xml });
      if (output.observation !== undefined) {
        capabilities[module.declaration] = output.observation;
      }
      for (const extension of output.extensions) {
        if (!target.extensions.includes(extension)) {
          slide.diagnostics.push(
            diagnostic(
              path,
              "forbidden-mechanism",
              `Target forbids ${extension}`,
            ),
          );
        } else {
          document.usedExtensions.add(extension);
        }
      }
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
  for (const contribution of slide.contributions) {
    const consumer = modules.get(contribution.capability);
    if (
      consumer === undefined ||
      !isSlideCapability(consumer) ||
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
  return {
    id,
    ...(authorId === undefined ? {} : { authorId }),
    index,
    ...(section === undefined ? {} : { section }),
    hidden,
    ...(advanceAfter === undefined ? {} : { advanceAfter }),
    objects: compiled.objects,
    layouts: [...slide.layouts].sort((left, right) => left.id - right.id),
    content: compiled.content,
    resources: slide.resourceObservations,
    shapeXml: compiled.xml,
    backgroundXml,
    slideXml: slideXml
      .map((item, index) => ({ item, index }))
      .sort(
        (left, right) =>
          slideSlotOrder(left.item.slot) - slideSlotOrder(right.item.slot) ||
          left.index - right.index,
      )
      .map(({ item }) => item.xml)
      .join(""),
    capabilities,
    relationships: slide.relationships,
    morphObjects: slide.morphObjects,
    linkedSlideIndexes: slide.linkedSlideIndexes,
  };
}

function slideSlotOrder(slot: SlideCapabilitySlot): number {
  return slot === "transition" ? 0 : 1;
}

function compileDocument(
  container: DeclarationContainer,
  modules: ReadonlyMap<string, RegisteredCapabilityModule>,
  target: Target,
  options: BuildOptions,
): BuildResult {
  const resolved = resolveSlides(container, modules);
  const { slides, sections } = resolved;
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
  if (sections.length > 0 && !target.extensions.includes("p14")) {
    throw new BuildError([
      diagnostic(
        "deck/sections",
        "forbidden-mechanism",
        "Current target forbids p14 section mechanism",
      ),
    ]);
  }
  const document: DocumentBuildState = {
    usedExtensions: new Set(),
    warnings: [],
    resourceClaims: new Map(),
    resources: new Map(),
    resourceTypes: new Map(),
    packageParts: new Map(),
    packageRelationships: new Map(),
    presentationAttributes: new Map(),
    presentationPartPaths: new Set(),
    presentationCollections: new Map(),
    nextPackagePartId: 1,
    nextPresentationRelationshipId: slides.length + 2,
    morphNames: new Map(),
    nextMorphName: 1,
    slideIndexById,
    nextTextFieldId: 1,
  };
  const builtSlides = slides.map((resolvedSlide) =>
    compileSlide(
      resolvedSlide,
      modules,
      target,
      size,
      document,
      options,
      theme,
    ),
  );
  if (sections.length > 0) document.usedExtensions.add("p14");
  const morphDiagnostics: Diagnostic[] = [];
  for (let index = 0; index < slides.length; index += 1) {
    const current = slides[index];
    const currentBuild = builtSlides[index];
    if (
      current === undefined ||
      currentBuild === undefined ||
      !usesMorphTransition(current.node)
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
      sourceIndex < builtSlides.length;
      sourceIndex += 1
    ) {
      if (
        sourceIndex !== index &&
        builtSlides[sourceIndex]?.linkedSlideIndexes.has(current.index)
      ) {
        sourceIndexes.add(sourceIndex);
      }
    }
    for (const [key, object] of currentBuild.morphObjects) {
      for (const sourceIndex of sourceIndexes) {
        const source = builtSlides[sourceIndex]?.morphObjects.get(key);
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
  finalizePresentationContributions(document);
  const parts = pptxParts(
    size,
    builtSlides.map(
      ({ id, hidden, shapeXml, backgroundXml, slideXml, relationships }) => ({
        id,
        hidden,
        shapeXml,
        backgroundXml,
        slideXml,
        relationships,
      }),
    ),
    [...document.resources.values()],
    [...document.packageParts.values()],
    [...document.packageRelationships].map(([sourcePath, relationships]) => ({
      sourcePath,
      relationships,
    })) satisfies PptxPackageRelationshipSet[],
    presentationContent(document),
    [...document.presentationAttributes].map(([name, value]) => ({
      name,
      value,
    })),
    theme.theme,
    title,
    sections,
  );
  const pptx = zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([path, content]) => [
        path,
        typeof content === "string" ? strToU8(content) : content,
      ]),
    ),
    { level: 6, mtime: new Date("2020-01-01T00:00:00Z") },
  );
  const slideObservations: SlideObservation[] = builtSlides.map((slide) => ({
    id: slide.id,
    ...(slide.authorId === undefined ? {} : { authorId: slide.authorId }),
    index: slide.index,
    ...(slide.section === undefined ? {} : { section: slide.section }),
    ...(slide.hidden ? { hidden: true } : {}),
    ...(slide.advanceAfter === undefined
      ? {}
      : { advanceAfter: Math.round(slide.advanceAfter) }),
    objects: slide.objects,
    ...(slide.layouts.length > 0 && { layouts: slide.layouts }),
    ...(Object.keys(slide.capabilities).length > 0 && {
      capabilities: slide.capabilities,
    }),
  }));
  const report: BuildReport = deepFreeze({
    evidence: "generation",
    ...(typeof title === "string" ? { title } : {}),
    target: {
      name: target.name,
      format: target.format,
      extensions: [...target.extensions],
    },
    extensions: [...document.usedExtensions],
    slides: slideObservations,
    diagnostics: document.warnings,
  });
  const query = createBuildQuery(
    builtSlides.map((slide, index) => ({
      observation: report.slides[index]!,
      content: slide.content,
      resources: slide.resources,
    })),
  );
  return Object.freeze({ pptx, report, query });
}

/**
 * Creates the synchronous formal build over a plain declaration tree.
 *
 * The formal build performs validation, layout, choreography, resource coordination, and packaging
 * without executing React lifecycle or importing React at runtime. React execution hands off its
 * declaration tree before this function runs, and the same implementation serves later authoring
 * entries. The shared editable document should expose this declaration view directly; do not add a
 * document-to-host reverse tree or a second compiler.
 */
export function createFormalBuild(modules: readonly CapabilityModule[]) {
  const bound = bindCapabilityModules(modules);

  return function buildDeclaration(
    declaration: DeclarationContainer,
    options: BuildOptions = {},
  ): BuildResult {
    const target = validateOptions(options);
    if (options.signal?.aborted) throw options.signal.reason;
    return compileDocument(declaration, bound, target, options);
  };
}

/**
 * Internal React composition entry. React expansion has already locked its completion when the
 * declaration is handed off, so a cleanup abort must not be treated as cancelling it. The original
 * options, including the caller's signal, still reach capability parse and compile contexts.
 */
export function createCompletedDeclarationBuild(
  modules: readonly CapabilityModule[],
) {
  const bound = bindCapabilityModules(modules);

  return function buildCompletedDeclaration(
    declaration: DeclarationContainer,
    options: BuildOptions = {},
  ): BuildResult {
    const target = validateOptions(options);
    return compileDocument(declaration, bound, target, options);
  };
}

function bindCapabilityModules(
  modules: readonly CapabilityModule[],
): ReadonlyMap<string, RegisteredCapabilityModule> {
  const bound = new Map<string, RegisteredCapabilityModule>();
  for (const module of modules) {
    const declarations = [
      module,
      ...("compileSlide" in module
        ? (module.transparentDeclarations ?? [])
        : []),
    ];
    for (const declaration of declarations) {
      if (bound.has(declaration.declaration)) {
        throw new BuildError([
          diagnostic(
            "build.modules",
            "duplicate-capability",
            "Duplicate capability module declaration",
          ),
        ]);
      }
      bound.set(declaration.declaration, declaration);
    }
  }
  return bound;
}
