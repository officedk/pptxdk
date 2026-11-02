import { createHash } from "node:crypto";

import { strToU8, zipSync } from "fflate";
import type { ReactNode } from "react";

import { CapabilityError, BuildError } from "#/diagnostics.ts";
import { HostNode, type HostChild, type HostContainer } from "#/host/nodes.ts";
import { withHostTree } from "#/host/reconciler.ts";
import {
  assertXml10Characters,
  pptxParts,
  pxToEmu,
  type PptxPackageResource,
  type PptxRelationship,
} from "#/ooxml.ts";
import {
  formatBaseline,
  targets,
  type BuildOptions,
  type BuildReport,
  type BuildResult,
  type CapabilityChild,
  type CapabilityChildren,
  type CapabilityInput,
  type CapabilityContribution,
  type CapabilityModule,
  type CapabilityResource,
  type Diagnostic,
  type Extension,
  type ObjectCapabilityModule,
  type ObjectObservation,
  type SlideCapabilityModule,
  type SlideObservation,
  type SlideCapabilitySlot,
  type Target,
} from "#/types.ts";

type SlideBuild = {
  readonly id: string;
  readonly objects: readonly ObjectObservation[];
  readonly shapeXml: string;
  readonly slideXml: string;
  readonly capabilities: Readonly<Record<string, unknown>>;
  readonly relationships: readonly PptxRelationship[];
};

type DocumentBuildState = {
  readonly usedExtensions: Set<Extension>;
  readonly warnings: Diagnostic[];
  readonly resourceClaims: Map<string, { identity: string; path: string }>;
  readonly resources: Map<string, PptxPackageResource>;
  readonly resourceTypes: Map<string, string>;
};

type SlideBuildState = {
  nextId: number;
  readonly diagnostics: Diagnostic[];
  readonly relationships: PptxRelationship[];
  readonly relationshipIds: Map<string, string>;
  readonly contributions: CapabilityContribution[];
};

type BuildContext = {
  readonly document: DocumentBuildState;
  readonly slide: SlideBuildState;
  readonly modules: ReadonlyMap<string, CapabilityModule>;
  readonly target: Target;
  readonly dimensions: Dimensions;
};

type Dimensions = { readonly width: number; readonly height: number };
const minimumSlideSizeEmu = 914_400;
const maximumSlideSizeEmu = 51_206_400;

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

function deepFreeze<T>(value: T): T {
  if (typeof value !== "object" || value === null || Object.isFrozen(value))
    return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}

function ownKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
): string | undefined {
  return Object.keys(value).find((key) => !allowed.includes(key));
}

function isHostNode(child: HostChild): child is HostNode {
  return child instanceof HostNode;
}

function isSlideCapability(
  module: CapabilityModule,
): module is SlideCapabilityModule {
  return "compileSlide" in module;
}

function isObjectCapability(
  module: CapabilityModule,
): module is ObjectCapabilityModule {
  return "compile" in module;
}

function nonWhitespaceText(children: readonly HostChild[]): boolean {
  return children.some(
    (child) => !isHostNode(child) && child.text.trim() !== "",
  );
}

function childNodes(children: readonly HostChild[]): HostNode[] {
  return children.filter(isHostNode);
}

function componentLabel(node: HostNode, index: number): string {
  const type = node.type.startsWith("rp:") ? node.type.slice(3) : node.type;
  const identity = node.props["testId"];
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${type}[${label}]`;
}

function capabilityChild(child: HostChild): CapabilityChild {
  if (!isHostNode(child)) return { type: "#text", text: child.text };
  return {
    type: child.type,
    props: child.props,
    children: child.children.map(capabilityChild),
  };
}

function capabilityLabel(child: CapabilityChild, index: number): string {
  if (!("props" in child)) return `text[${index + 1}]`;
  const type = child.type.startsWith("rp:") ? child.type.slice(3) : child.type;
  const identity = child.props["testId"];
  const label =
    typeof identity === "string" && identity.trim() !== ""
      ? identity
      : String(index + 1);
  return `${type}[${label}]`;
}

function validateTarget(target: Target | undefined): Target {
  const value = target ?? targets.transitional;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new BuildError([
      diagnostic("build.target", "invalid-target", "输出目标必须是对象"),
    ]);
  }
  const record = value as unknown as Record<string, unknown>;
  const unknown = ownKeys(record, ["name", "format", "extensions"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "unknown-parameter",
        `输出目标含未知参数 ${unknown}`,
      ),
    ]);
  }
  if (typeof record["name"] !== "string" || record["name"].trim() === "") {
    throw new BuildError([
      diagnostic("build.target", "invalid-target", "输出目标需要非空名称"),
    ]);
  }
  if (record["format"] !== formatBaseline) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "unsupported-format",
        `本构建仅写出 ${formatBaseline}`,
      ),
    ]);
  }
  if (!Array.isArray(record["extensions"])) {
    throw new BuildError([
      diagnostic(
        "build.target",
        "invalid-target",
        "输出目标 extensions 必须是数组",
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
          `不支持输出目标扩展 ${String(extension)}`,
        ),
      ]);
    }
    if (extensions.includes(extension)) {
      throw new BuildError([
        diagnostic(
          "build.target",
          "duplicate-extension",
          `输出目标重复声明 ${extension}`,
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

function validateOptions(options: BuildOptions): Target {
  if (
    typeof options !== "object" ||
    options === null ||
    Array.isArray(options)
  ) {
    throw new BuildError([
      diagnostic("build", "invalid-options", "构建选项必须是对象"),
    ]);
  }
  const unknown = ownKeys(options, ["target", "signal"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic("build", "unknown-parameter", `构建选项含未知参数 ${unknown}`),
    ]);
  }
  return validateTarget(options.target);
}

function positiveDimension(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new BuildError([
      diagnostic("deck", "invalid-dimension", `${label} 必须是有限正像素数值`),
    ]);
  }
  const sizeEmu = pxToEmu(value);
  if (sizeEmu < minimumSlideSizeEmu || sizeEmu > maximumSlideSizeEmu) {
    throw new BuildError([
      diagnostic(
        "deck",
        "invalid-dimension",
        `${label} 换算后的页面尺寸必须在 ${minimumSlideSizeEmu} 到 ${maximumSlideSizeEmu} EMU 的闭区间内`,
      ),
    ]);
  }
  return value;
}

function deckDimensions(node: HostNode): Dimensions {
  const unknown = ownKeys(node.props, ["width", "height"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic("deck", "unknown-parameter", `Deck 含未知参数 ${unknown}`),
    ]);
  }
  return {
    width: positiveDimension(node.props["width"] ?? 1280, "Deck.width"),
    height: positiveDimension(node.props["height"] ?? 720, "Deck.height"),
  };
}

function resolveSlides(
  container: HostContainer,
  modules: ReadonlyMap<string, CapabilityModule>,
): { node: HostNode; id: string }[] {
  const decks = childNodes(container.children);
  if (decks.length !== 1 || decks[0]?.type !== "rp:deck") {
    throw new BuildError([
      diagnostic("deck", "document-root", "构建内容必须只有一个 Deck 根组件"),
    ]);
  }
  const deck = decks[0];
  if (deck === undefined)
    throw new BuildError([
      diagnostic("deck", "document-root", "缺少 Deck 根组件"),
    ]);
  if (nonWhitespaceText(container.children)) {
    throw new BuildError([
      diagnostic("deck", "unexpected-text", "Deck 根组件外不能有文本"),
    ]);
  }
  const slideNodes = childNodes(deck.children);
  if (nonWhitespaceText(deck.children)) {
    throw new BuildError([
      diagnostic("deck", "unexpected-text", "Deck 只能包含 Slide 页面"),
    ]);
  }
  if (
    slideNodes.length === 0 ||
    slideNodes.length !== deck.children.filter(isHostNode).length
  ) {
    throw new BuildError([
      diagnostic("deck", "missing-slide", "Deck 至少需要一个 Slide 页面"),
    ]);
  }
  const ids = new Set<string>();
  return slideNodes.map((node, index) => {
    if (node.type !== "rp:slide") {
      throw new BuildError([
        diagnostic(
          `deck/${componentLabel(node, index)}`,
          "unexpected-declaration",
          "Deck 的直接子项必须是 Slide",
        ),
      ]);
    }
    const slideProperties = [...modules.values()]
      .filter(isSlideCapability)
      .map(({ slideProp }) => slideProp);
    const unknown = ownKeys(node.props, ["slideId", ...slideProperties]);
    if (unknown !== undefined) {
      const rawId = node.props["slideId"];
      const label = typeof rawId === "string" ? rawId : index + 1;
      throw new BuildError([
        diagnostic(
          `deck/slide[${label}]`,
          "missing-capability",
          `没有注册 Slide.${unknown} 的能力模块`,
        ),
      ]);
    }
    const id = node.props["slideId"];
    if (typeof id !== "string" || id.trim() === "") {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-slide-id",
          "Slide.slideId 必须是非空字符串",
        ),
      ]);
    }
    try {
      assertXml10Characters(id);
    } catch {
      throw new BuildError([
        diagnostic(
          `deck/slide[${index + 1}]`,
          "invalid-xml-character",
          "Slide.slideId 含 XML 1.0 不允许的字符",
        ),
      ]);
    }
    if (ids.has(id)) {
      throw new BuildError([
        diagnostic(
          `deck/slide[${id}]`,
          "duplicate-slide-id",
          `页面标识 ${id} 重复`,
        ),
      ]);
    }
    ids.add(id);
    return { node, id };
  });
}

function addResource(
  resource: CapabilityResource,
  context: BuildContext,
): string {
  if (
    !(resource.bytes instanceof Uint8Array) ||
    resource.bytes.byteLength === 0
  ) {
    throw new CapabilityError("invalid-resource", "资源必须包含非空二进制内容");
  }
  if (
    typeof resource.extension !== "string" ||
    !/^[a-z0-9]+$/iu.test(resource.extension) ||
    typeof resource.contentType !== "string" ||
    resource.contentType.trim() === "" ||
    typeof resource.relationshipType !== "string" ||
    resource.relationshipType.trim() === ""
  ) {
    throw new CapabilityError("invalid-resource", "资源的格式或关系类型无效");
  }
  const extension = resource.extension.toLowerCase();
  if (extension === "xml" || extension === "rels") {
    throw new CapabilityError(
      "invalid-resource",
      `资源扩展名 .${extension} 与 PPTX 核心部件冲突`,
    );
  }
  const knownType = context.document.resourceTypes.get(extension);
  if (knownType !== undefined && knownType !== resource.contentType) {
    throw new CapabilityError(
      "conflicting-resource-type",
      `资源扩展名 .${extension} 对应多个内容类型`,
    );
  }
  context.document.resourceTypes.set(extension, resource.contentType);
  try {
    assertXml10Characters(resource.contentType);
    assertXml10Characters(resource.relationshipType);
  } catch {
    throw new CapabilityError(
      "invalid-resource",
      "资源关系含 XML 不允许的字符",
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

function compileChildren(
  children: readonly CapabilityChild[],
  context: BuildContext,
  parentPath: string,
): CapabilityChildren {
  const xml: string[] = [];
  const objects: ObjectObservation[] = [];
  children.forEach((child, index) => {
    if (!("props" in child)) {
      if (child.text.trim() !== "") {
        context.slide.diagnostics.push(
          diagnostic(
            `${parentPath}/${capabilityLabel(child, index)}`,
            "unexpected-text",
            "原生对象容器不能直接包含文本",
          ),
        );
      }
      return;
    }

    const path = `${parentPath}/${capabilityLabel(child, index)}`;
    const module = context.modules.get(child.type);
    if (module === undefined || !isObjectCapability(module)) {
      context.slide.diagnostics.push(
        diagnostic(
          path,
          "missing-capability",
          `没有注册 ${child.type} 的能力模块`,
        ),
      );
      return;
    }
    const id = context.slide.nextId++;
    const input: CapabilityInput = child;
    try {
      const output = module.compile(input, {
        id,
        path,
        compileChildren(nested) {
          return compileChildren(nested, context, path);
        },
        registerResource(resource) {
          return addResource(resource, context);
        },
      });
      for (const extension of output.extensions) {
        if (!context.target.extensions.includes(extension)) {
          context.slide.diagnostics.push(
            diagnostic(
              path,
              "forbidden-mechanism",
              `输出目标禁止 ${extension}`,
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
              `资源 ${claim.key} 与 ${existing.path} 使用了冲突定义`,
            ),
          );
        } else {
          context.document.resourceClaims.set(claimKey, {
            identity: claim.identity,
            path,
          });
        }
      }
      for (const warning of output.warnings ?? []) {
        context.document.warnings.push({
          severity: "warning",
          path,
          code: warning.code,
          message: warning.message,
        });
      }
      for (const annotation of output.annotations ?? []) {
        context.slide.contributions.push({
          ...annotation,
          path,
          targetId: id,
        });
      }
      xml.push(output.xml);
      const observation = { ...output.observation, id };
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
            message: "对象超出页面画布，构建保留原始声明",
          });
        }
      }
      objects.push(observation);
    } catch (error) {
      context.slide.diagnostics.push(
        error instanceof CapabilityError
          ? diagnostic(path, error.code, error.message)
          : diagnostic(
              path,
              "capability-failure",
              error instanceof Error ? error.message : String(error),
            ),
      );
    }
  });
  return { xml: xml.join(""), objects };
}

function compileSlide(
  node: HostNode,
  id: string,
  modules: ReadonlyMap<string, CapabilityModule>,
  target: Target,
  dimensions: Dimensions,
  document: DocumentBuildState,
): SlideBuild {
  const slide: SlideBuildState = {
    nextId: 2,
    diagnostics: [],
    relationships: [],
    relationshipIds: new Map(),
    contributions: [],
  };
  const path = `deck/slide[${id}]`;
  if (nonWhitespaceText(node.children)) {
    slide.diagnostics.push(
      diagnostic(path, "unexpected-text", "Slide 页面不能直接包含文本"),
    );
  }
  const compiled = compileChildren(
    node.children.map(capabilityChild),
    { document, slide, modules, target, dimensions },
    path,
  );
  const slideXml: {
    readonly slot: SlideCapabilitySlot;
    readonly xml: string;
  }[] = [];
  const capabilities: Record<string, unknown> = {};
  const input: CapabilityInput = {
    type: node.type,
    props: node.props,
    children: node.children.map(capabilityChild),
  };
  for (const module of modules.values()) {
    if (!isSlideCapability(module)) continue;
    if (!module.always && !Object.hasOwn(node.props, module.slideProp))
      continue;
    try {
      const output = module.compileSlide(input, {
        path,
        contributions: slide.contributions,
      });
      slideXml.push({ slot: module.slot, xml: output.xml });
      if (output.observation !== undefined) {
        capabilities[module.declaration] = output.observation;
      }
      for (const extension of output.extensions) {
        if (!target.extensions.includes(extension)) {
          slide.diagnostics.push(
            diagnostic(
              path,
              "forbidden-mechanism",
              `输出目标禁止 ${extension}`,
            ),
          );
        } else {
          document.usedExtensions.add(extension);
        }
      }
    } catch (error) {
      slide.diagnostics.push(
        error instanceof CapabilityError
          ? diagnostic(path, error.code, error.message)
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
      (!consumer.always && !Object.hasOwn(node.props, consumer.slideProp))
    ) {
      slide.diagnostics.push(
        diagnostic(
          contribution.path,
          "missing-capability",
          `没有注册 ${contribution.capability} 的能力模块`,
        ),
      );
    }
  }
  if (slide.diagnostics.length > 0) throw new BuildError(slide.diagnostics);
  return {
    id,
    objects: compiled.objects,
    shapeXml: compiled.xml,
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
  };
}

function slideSlotOrder(slot: SlideCapabilitySlot): number {
  return slot === "transition" ? 0 : 1;
}

function compileDocument(
  container: HostContainer,
  modules: ReadonlyMap<string, CapabilityModule>,
  target: Target,
): BuildResult {
  const slides = resolveSlides(container, modules);
  const deck = childNodes(container.children)[0];
  if (deck === undefined)
    throw new BuildError([
      diagnostic("deck", "document-root", "缺少 Deck 根组件"),
    ]);
  const size = deckDimensions(deck);
  const document: DocumentBuildState = {
    usedExtensions: new Set(),
    warnings: [],
    resourceClaims: new Map(),
    resources: new Map(),
    resourceTypes: new Map(),
  };
  const builtSlides = slides.map(({ node, id }) =>
    compileSlide(node, id, modules, target, size, document),
  );
  const parts = pptxParts(
    size,
    builtSlides.map(({ id, shapeXml, slideXml, relationships }) => ({
      id,
      shapeXml,
      slideXml,
      relationships,
    })),
    [...document.resources.values()],
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
  const slideObservations: SlideObservation[] = builtSlides.map(
    (slide, index) => ({
      id: slide.id,
      index: index + 1,
      objects: slide.objects,
      ...(Object.keys(slide.capabilities).length > 0 && {
        capabilities: slide.capabilities,
      }),
    }),
  );
  const report: BuildReport = deepFreeze({
    evidence: "generation",
    target: {
      name: target.name,
      format: target.format,
      extensions: [...target.extensions],
    },
    extensions: [...document.usedExtensions],
    slides: slideObservations,
    diagnostics: document.warnings,
  });
  return Object.freeze({ pptx, report });
}

export function createBuilder(modules: readonly CapabilityModule[]) {
  const bound = new Map(modules.map((module) => [module.declaration, module]));
  if (bound.size !== modules.length) {
    throw new BuildError([
      diagnostic("build.modules", "duplicate-capability", "能力模块声明重复"),
    ]);
  }

  return async function build(
    element: ReactNode,
    options: BuildOptions = {},
  ): Promise<BuildResult> {
    const target = validateOptions(options);
    try {
      return await withHostTree(
        element,
        (container) => compileDocument(container, bound, target),
        options.signal,
      );
    } catch (error) {
      if (options.signal?.aborted && error === options.signal.reason) {
        throw error;
      }
      if (error instanceof BuildError) throw error;
      throw new BuildError([
        diagnostic(
          "deck",
          "react-render",
          error instanceof Error ? error.message : String(error),
        ),
      ]);
    }
  };
}
