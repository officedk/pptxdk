import { strToU8, zipSync } from "fflate";
import type { ReactNode } from "react";

import { CapabilityError, BuildError } from "#/diagnostics.ts";
import { HostNode, type HostChild, type HostContainer } from "#/host/nodes.ts";
import { withHostTree } from "#/host/reconciler.ts";
import { pptxParts } from "#/ooxml.ts";
import {
  formatBaseline,
  targets,
  type BuildOptions,
  type BuildReport,
  type BuildResult,
  type CapabilityInput,
  type CapabilityModule,
  type Diagnostic,
  type Extension,
  type ObjectObservation,
  type SlideObservation,
  type Target,
} from "#/types.ts";

type SlideBuild = {
  readonly id: string;
  readonly objects: readonly ObjectObservation[];
  readonly shapeXml: string;
};

type Dimensions = { readonly width: number; readonly height: number };

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
  const unknown = ownKeys(options, ["target"]);
  if (unknown !== undefined) {
    throw new BuildError([
      diagnostic("build", "unknown-parameter", `构建选项含未知参数 ${unknown}`),
    ]);
  }
  return validateTarget(options.target);
}

function positiveDimension(value: unknown, label: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value <= 0 ||
    value > 1_000_000
  ) {
    throw new BuildError([
      diagnostic(
        "deck",
        "invalid-dimension",
        `${label} 必须是 0 到 1000000 之间的有限像素数值`,
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
    const unknown = ownKeys(node.props, ["slideId"]);
    if (unknown !== undefined) {
      const rawId = node.props["slideId"];
      const label = typeof rawId === "string" ? rawId : index + 1;
      throw new BuildError([
        diagnostic(
          `deck/slide[${label}]`,
          "unknown-parameter",
          `Slide 含未知参数 ${unknown}`,
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

function compileSlide(
  node: HostNode,
  id: string,
  index: number,
  modules: ReadonlyMap<string, CapabilityModule>,
  target: Target,
  dimensions: Dimensions,
  usedExtensions: Set<Extension>,
  warnings: Diagnostic[],
): SlideBuild {
  const diagnostics: Diagnostic[] = [];
  const objects: ObjectObservation[] = [];
  const shapeXml: string[] = [];
  if (nonWhitespaceText(node.children)) {
    diagnostics.push(
      diagnostic(
        `deck/slide[${id}]`,
        "unexpected-text",
        "Slide 页面不能直接包含文本",
      ),
    );
  }
  const declarations = childNodes(node.children);
  declarations.forEach((declaration, declarationIndex) => {
    const path = `deck/slide[${id}]/${componentLabel(declaration, declarationIndex)}`;
    const module = modules.get(declaration.type);
    if (module === undefined) {
      diagnostics.push(
        diagnostic(
          path,
          "missing-capability",
          `没有注册 ${declaration.type} 的能力模块`,
        ),
      );
      return;
    }
    const input: CapabilityInput = {
      type: declaration.type,
      props: declaration.props,
      children: declaration.children,
    };
    try {
      const output = module.compile(input, { id: objects.length + 2, path });
      for (const extension of output.extensions) {
        if (!target.extensions.includes(extension)) {
          diagnostics.push(
            diagnostic(
              path,
              "forbidden-mechanism",
              `输出目标禁止 ${extension}`,
            ),
          );
        } else {
          usedExtensions.add(extension);
        }
      }
      shapeXml.push(output.xml);
      const observation = { ...output.observation, id: objects.length + 2 };
      if (observation.box !== undefined) {
        const box = observation.box;
        if (
          box.x < 0 ||
          box.y < 0 ||
          box.x + box.width > dimensions.width ||
          box.y + box.height > dimensions.height
        ) {
          warnings.push({
            severity: "warning",
            path,
            code: "overflow",
            message: "对象超出页面画布，构建保留原始声明",
          });
        }
      }
      objects.push(observation);
    } catch (error) {
      diagnostics.push(
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
  if (diagnostics.length > 0) throw new BuildError(diagnostics);
  return { id, objects, shapeXml: shapeXml.join("") };
}

function compileDocument(
  container: HostContainer,
  modules: ReadonlyMap<string, CapabilityModule>,
  target: Target,
): BuildResult {
  const slides = resolveSlides(container);
  const deck = childNodes(container.children)[0];
  if (deck === undefined)
    throw new BuildError([
      diagnostic("deck", "document-root", "缺少 Deck 根组件"),
    ]);
  const size = deckDimensions(deck);
  const usedExtensions = new Set<Extension>();
  const warnings: Diagnostic[] = [];
  const builtSlides = slides.map(({ node, id }, index) =>
    compileSlide(
      node,
      id,
      index,
      modules,
      target,
      size,
      usedExtensions,
      warnings,
    ),
  );
  const parts = pptxParts(
    size,
    builtSlides.map(({ id, shapeXml }) => ({ id, shapeXml })),
  );
  const pptx = zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([path, xml]) => [path, strToU8(xml)]),
    ),
    { level: 6, mtime: new Date("2020-01-01T00:00:00Z") },
  );
  const slideObservations: SlideObservation[] = builtSlides.map(
    (slide, index) => ({
      id: slide.id,
      index: index + 1,
      objects: slide.objects,
    }),
  );
  const report: BuildReport = deepFreeze({
    evidence: "generation",
    target: {
      name: target.name,
      format: target.format,
      extensions: [...target.extensions],
    },
    extensions: [...usedExtensions],
    slides: slideObservations,
    diagnostics: warnings,
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
      return withHostTree(element, (container) =>
        compileDocument(container, bound, target),
      );
    } catch (error) {
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
