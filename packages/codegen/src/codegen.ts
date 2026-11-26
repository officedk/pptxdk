import {
  isSlideLayoutDefinition,
  type LayoutPlaceholderSpec,
  type PlaceholderMap,
} from "@pptxdk/core";
import type { DeclarationChild, DeclarationText } from "@pptxdk/core/build";
import {
  defaultTheme,
  getMorphKeyLabel,
  isFont,
  isMorphKey,
  isTheme,
  type Font,
  type MorphKey,
  type RegisteredCapabilityModule,
  type Theme,
} from "@pptxdk/core/capability";
import {
  CommonNode,
  SlideLayoutContainer,
  SlideMasterContainer,
} from "@pptxdk/core/document";

import { AssetRegistry } from "./assets.ts";
import {
  IdentifierAllocator,
  toCamelCaseIdentifier,
  toPascalCaseIdentifier,
} from "./identifiers.ts";
import {
  formatJsxElement,
  formatJsxTextChild,
  formatPropsObject,
  formatTsValue,
  formatVanillaCall,
  indentBlock,
  type ValueHooks,
} from "./printer.ts";
import {
  CodegenError,
  codegenOptionsSchema,
  expr,
  type CodegenContext,
  type CodegenDiagnostic,
  type CodegenDiagnosticCode,
  type CodegenExpr,
  type CodegenFileContent,
  type CodegenOptions,
  type CodegenPropSerializer,
  type CodegenResult,
  type CodegenTarget,
  type CodegenValueSerializer,
  type DeclarationCodegenHandler,
  type ObjectCodegenContribution,
  type SlideCodegenContribution,
} from "./types.ts";

const coreDeckPropKeys = new Set([
  "width",
  "height",
  "theme",
  "style",
  "title",
]);

const coreMasterPropKeys = new Set([
  "id",
  "name",
  "preserve",
  "theme",
  "colorMap",
  "placeholders",
  "textStyles",
  "style",
]);

const coreLayoutPropKeys = new Set([
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
  "style",
]);

const coreSlidePropKeys = new Set([
  "id",
  "name",
  "layout",
  "placeholders",
  "theme",
  "colorMapOverride",
  "themeOverride",
  "style",
  "hidden",
  "showMasterPlaceholderAnimations",
  "showMasterShapes",
]);

type CollectedContributions = {
  readonly declarations: ReadonlyMap<string, DeclarationCodegenHandler>;
  readonly objectProps: ReadonlyMap<string, CodegenPropSerializer>;
  readonly ownerProps: ReadonlyMap<string, CodegenPropSerializer>;
  readonly slideProps: ReadonlyMap<string, CodegenPropSerializer>;
  readonly valueSerializers: readonly CodegenValueSerializer[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isObjectCodegenContribution(
  value: unknown,
): value is ObjectCodegenContribution {
  return isRecord(value) && isRecord(value["declarations"]);
}

function isSlideCodegenContribution(
  value: unknown,
): value is SlideCodegenContribution {
  return isRecord(value) && isRecord(value["slideProps"]);
}

function isDeclarationTextLeaf(
  child: DeclarationChild,
): child is DeclarationText {
  return !(child instanceof CommonNode) && child.type === "#text";
}

function collectContributions(
  capabilities: readonly RegisteredCapabilityModule[],
): CollectedContributions {
  const declarations = new Map<string, DeclarationCodegenHandler>();
  const objectProps = new Map<string, CodegenPropSerializer>();
  const ownerProps = new Map<string, CodegenPropSerializer>();
  const slideProps = new Map<string, CodegenPropSerializer>();
  const valueSerializers: CodegenValueSerializer[] = [];

  for (const mod of capabilities) {
    const rawCodegen = (mod as { readonly codegen?: unknown }).codegen;
    if (rawCodegen === undefined) continue;
    if (isObjectCodegenContribution(rawCodegen)) {
      for (const [kind, handler] of Object.entries(rawCodegen.declarations)) {
        declarations.set(kind, handler);
      }
      if (rawCodegen.objectProps !== undefined) {
        for (const [prop, serializer] of Object.entries(
          rawCodegen.objectProps,
        )) {
          objectProps.set(prop, serializer);
        }
      }
      if (rawCodegen.ownerProps !== undefined) {
        for (const [prop, serializer] of Object.entries(
          rawCodegen.ownerProps,
        )) {
          ownerProps.set(prop, serializer);
        }
      }
      if (rawCodegen.valueSerializers !== undefined) {
        valueSerializers.push(...rawCodegen.valueSerializers);
      }
    } else if (isSlideCodegenContribution(rawCodegen)) {
      for (const [prop, serializer] of Object.entries(rawCodegen.slideProps)) {
        slideProps.set(prop, serializer);
      }
      if (rawCodegen.valueSerializers !== undefined) {
        valueSerializers.push(...rawCodegen.valueSerializers);
      }
    }
  }

  return {
    declarations,
    objectProps,
    ownerProps,
    slideProps,
    valueSerializers,
  };
}

function extractFieldFromCodegenPath(path: string): string | undefined {
  let bracketDepth = 0;
  for (let index = 0; index < path.length; index += 1) {
    const ch = path[index];
    if (ch === "[") {
      bracketDepth += 1;
    } else if (ch === "]" && bracketDepth > 0) {
      bracketDepth -= 1;
    } else if (ch === "@" && bracketDepth === 0) {
      const field = path.slice(index + 1);
      return field.length > 0 ? field : undefined;
    }
  }
  return undefined;
}

function extractDeclarationFromCodegenPath(path: string): string | undefined {
  const withoutField = path.split("@")[0] ?? path;
  const lastSegment = withoutField.split("/").at(-1);
  if (lastSegment === undefined || lastSegment === "") return undefined;
  const bracketIdx = lastSegment.indexOf("[");
  return bracketIdx === -1 ? lastSegment : lastSegment.slice(0, bracketIdx);
}

function resolveFailSubPath(basePath: string, subPath?: string): string {
  if (subPath === undefined || subPath === "") return basePath;
  if (
    subPath === "deck" ||
    subPath.startsWith("deck/") ||
    subPath.startsWith("deck@")
  ) {
    return subPath;
  }
  const hasField = extractFieldFromCodegenPath(basePath) !== undefined;
  if (subPath.startsWith("@")) {
    return hasField
      ? `${basePath}.${subPath.slice(1)}`
      : `${basePath}${subPath}`;
  }
  if (subPath.startsWith("[")) {
    return `${basePath}${subPath}`;
  }
  return hasField ? `${basePath}.${subPath}` : `${basePath}@${subPath}`;
}

function failAt(
  path: string,
  code: CodegenDiagnosticCode,
  message: string,
  context?: {
    readonly field?: string;
    readonly declaration?: string;
  },
): never {
  const resolvedField = context?.field ?? extractFieldFromCodegenPath(path);
  let diagnostic: CodegenDiagnostic;
  switch (code) {
    case "unsupported-declaration":
      diagnostic = {
        code,
        path,
        message,
        declaration:
          context?.declaration ??
          extractDeclarationFromCodegenPath(path) ??
          "unknown",
      };
      break;
    case "unsupported-slide-property":
      diagnostic = {
        code,
        path,
        message,
        field: resolvedField ?? "unknown",
      };
      break;
    case "invalid-document":
    case "invalid-options":
    case "unexpressible-value":
      diagnostic = {
        code,
        path,
        message,
        ...(resolvedField === undefined ? {} : { field: resolvedField }),
      };
      break;
  }
  throw new CodegenError([diagnostic]);
}

function childPathLabel(node: CommonNode, index: number): string {
  return node.id === undefined
    ? `${node.type}[${index + 1}]`
    : `${node.type}[${node.id}]`;
}

function collectMorphKeysFromValue(
  value: unknown,
  out: MorphKey[],
  seen: Set<MorphKey>,
  visitedObjects: Set<object>,
): void {
  if (typeof value !== "object" || value === null) return;
  if (isMorphKey(value)) {
    if (!seen.has(value)) {
      seen.add(value);
      out.push(value);
    }
    return;
  }
  if (
    value instanceof CommonNode ||
    value instanceof Uint8Array ||
    value instanceof URL ||
    isTheme(value) ||
    isSlideLayoutDefinition(value) ||
    isFont(value)
  ) {
    return;
  }
  if (visitedObjects.has(value)) return;
  visitedObjects.add(value);
  if (Array.isArray(value)) {
    for (const item of value) {
      collectMorphKeysFromValue(item, out, seen, visitedObjects);
    }
    return;
  }
  for (const entry of Object.values(value)) {
    collectMorphKeysFromValue(entry, out, seen, visitedObjects);
  }
}

function collectMorphKeysFromTree(root: CommonNode): readonly MorphKey[] {
  const out: MorphKey[] = [];
  const seen = new Set<MorphKey>();
  const visitedObjects = new Set<object>();
  const visit = (node: CommonNode): void => {
    collectMorphKeysFromValue(node.props, out, seen, visitedObjects);
    for (const child of node.children) {
      if (child instanceof CommonNode) {
        visit(child);
      }
    }
  };
  visit(root);
  return out;
}

function fontDedupKey(fontValue: Font): string {
  const faces = (fontValue.faces ?? []).map((face) => ({
    weight: face.weight ?? null,
    style: face.style ?? null,
    dataLen: face.data?.byteLength ?? null,
    // Actual asset dedup uses SHA-256 inside AssetRegistry; include first/last bytes + length as fast key,
    // or hash inside registerFont. We compute a structural key once asset varName is resolved.
  }));
  return JSON.stringify({
    family: fontValue.family,
    substitute: fontValue.substitute ?? false,
    embeddable: fontValue.embeddable ?? true,
    faces,
  });
}

function isDefined(value: unknown): boolean {
  return value !== undefined;
}

const knownNativePlaceholderKeys = new Set([
  "type",
  "explicitType",
  "idx",
  "size",
  "orient",
  "hasCustomPrompt",
]);

function isNativePlaceholderRedundant(
  nativeRef: unknown,
  layoutSpec: LayoutPlaceholderSpec,
): boolean {
  if (!isRecord(nativeRef)) return false;
  for (const key of Object.keys(nativeRef)) {
    if (nativeRef[key] !== undefined && !knownNativePlaceholderKeys.has(key)) {
      return false;
    }
  }
  return (
    nativeRef["type"] === layoutSpec.type &&
    (nativeRef["explicitType"] !== false) ===
      (layoutSpec.explicitType !== false) &&
    nativeRef["idx"] === layoutSpec.idx &&
    nativeRef["size"] === layoutSpec.size &&
    nativeRef["orient"] === layoutSpec.orient &&
    (nativeRef["hasCustomPrompt"] === true) ===
      (layoutSpec.hasCustomPrompt === true)
  );
}

function stripRedundantNativePlaceholder(
  entries: readonly (readonly [string, unknown])[],
  layoutPlaceholders: PlaceholderMap | undefined,
): readonly (readonly [string, unknown])[] {
  if (layoutPlaceholders === undefined) return entries;
  let placeholderKey: string | undefined;
  let nativePlaceholderValue: unknown;
  for (const [key, value] of entries) {
    if (key === "placeholder" && typeof value === "string") {
      placeholderKey = value;
    } else if (key === "nativePlaceholder" && value !== undefined) {
      nativePlaceholderValue = value;
    }
  }
  if (placeholderKey === undefined || nativePlaceholderValue === undefined) {
    return entries;
  }
  const layoutSpec = layoutPlaceholders[placeholderKey];
  if (
    layoutSpec === undefined ||
    !isNativePlaceholderRedundant(nativePlaceholderValue, layoutSpec)
  ) {
    return entries;
  }
  return entries.filter(([key]) => key !== "nativePlaceholder");
}

/**
 * Creates a synchronous source code generator wired with the provided capability modules.
 */
export function createCodegen(
  capabilities: readonly RegisteredCapabilityModule[],
) {
  const registry = collectContributions(capabilities);

  return function codegen<T extends CodegenTarget>(
    document: CommonNode,
    options: CodegenOptions<T>,
  ): CodegenResult<T> {
    if (!(document instanceof CommonNode) || document.type !== "deck") {
      failAt("deck", "invalid-document", "Expected a Deck root CommonNode");
    }
    const parsedOptions = codegenOptionsSchema.safeParse(options);
    if (!parsedOptions.success) {
      const issueField = parsedOptions.error.issues[0]?.path
        .map(String)
        .join(".");
      failAt(
        "deck",
        "invalid-options",
        'Codegen options.target must be "vanilla" or "react"',
        {
          field:
            issueField !== undefined && issueField.length > 0
              ? issueField
              : "target",
        },
      );
    }
    const target = parsedOptions.data.target as T;
    const entry = (
      target === "vanilla" ? "deck.ts" : "deck.tsx"
    ) as CodegenResult<T>["entry"];

    const allocator = new IdentifierAllocator();
    const assets = new AssetRegistry(allocator);
    const sdkImports = new Set<string>();
    const nodeFsImports = new Set<string>();

    const useImport = (
      symbol: string,
      from: "sdk" | "node:fs" = "sdk",
    ): string => {
      if (from === "node:fs") {
        nodeFsImports.add(symbol);
      } else {
        sdkImports.add(symbol);
      }
      return symbol;
    };

    // 1. Greedy MorphScope allocation across the entire document
    const morphKeys = collectMorphKeysFromTree(document);
    const morphBuckets: {
      readonly varName: string;
      readonly labels: Set<string>;
    }[] = [];
    const morphScopeVarByKey = new Map<MorphKey, string>();
    for (const key of morphKeys) {
      const label = getMorphKeyLabel(key);
      let bucket = morphBuckets.find(
        (candidate) => !candidate.labels.has(label),
      );
      if (bucket === undefined) {
        const baseName =
          morphBuckets.length === 0
            ? "morphScope"
            : `morphScope${morphBuckets.length + 1}`;
        bucket = {
          varName: allocator.allocate(baseName),
          labels: new Set<string>(),
        };
        morphBuckets.push(bucket);
      }
      bucket.labels.add(label);
      morphScopeVarByKey.set(key, bucket.varName);
    }
    if (morphBuckets.length > 0) {
      useImport("createMorphScope");
    }

    // 2. Shared Fonts & Themes registries (emitted in Section 1)
    const fontVarByIdentity = new Map<string, string>();
    const fontDeclarations: string[] = [];
    const themeVarByInstance = new Map<Theme, string>();
    const themeDeclarations: string[] = [];

    // 3. Layout lookup maps (populated in Section 2, referenced in Section 3)
    const layoutIdentByContainer = new Map<SlideLayoutContainer, string>();
    const layoutIdentByDefinition = new Map<object, string>();
    const layoutDefinitionCounts = new Map<object, number>();

    for (const child of document.children) {
      if (!(child instanceof SlideMasterContainer)) continue;
      for (const layout of child.layouts) {
        const def = layout.props["layout"];
        if (isSlideLayoutDefinition(def)) {
          layoutDefinitionCounts.set(
            def,
            (layoutDefinitionCounts.get(def) ?? 0) + 1,
          );
        }
      }
    }

    let currentScopeAllocator: IdentifierAllocator | undefined;
    let currentMotionScopeVars: Map<string, string> | undefined;
    let currentSlideLayoutPlaceholders: PlaceholderMap | undefined;

    const serializeFont = (
      fontValue: Font,
      ctx: CodegenContext,
      hooks: ValueHooks,
    ): string => {
      const resolvedFaces = (fontValue.faces ?? []).map((face) => {
        const assetVar =
          face.data === undefined
            ? undefined
            : ctx.registerFontAsset(face.data, fontValue.family);
        return {
          ...(face.weight === undefined ? {} : { weight: face.weight }),
          ...(face.style === undefined ? {} : { style: face.style }),
          ...(assetVar === undefined ? {} : { data: expr(assetVar) }),
        };
      });
      const dedupKey = `${fontDedupKey(fontValue)}:${resolvedFaces
        .map((face) =>
          "data" in face && face.data !== undefined ? face.data.code : "",
        )
        .join(",")}`;
      const existing = fontVarByIdentity.get(dedupKey);
      if (existing !== undefined) return existing;

      const varName = allocator.allocate(
        toCamelCaseIdentifier(fontValue.family, "font", "Font"),
      );
      fontVarByIdentity.set(dedupKey, varName);

      const fontFactory = useImport("font");
      const optionsObj: Record<string, unknown> = {
        ...(fontValue.substitute === true ? { substitute: true } : {}),
        ...(fontValue.embeddable === false ? { embeddable: false } : {}),
      };
      const args: string[] = [JSON.stringify(fontValue.family)];
      const hasOptions = Object.keys(optionsObj).length > 0;
      if (resolvedFaces.length > 0 || hasOptions) {
        args.push(formatTsValue(resolvedFaces, ctx, hooks));
      }
      if (hasOptions) {
        args.push(formatTsValue(optionsObj, ctx, hooks));
      }
      fontDeclarations.push(
        `const ${varName} = ${fontFactory}(${args.join(", ")});`,
      );
      return varName;
    };

    const serializeTheme = (
      themeValue: Theme,
      ctx: CodegenContext,
      hooks: ValueHooks,
    ): string => {
      if (themeValue === defaultTheme) {
        return useImport("defaultTheme");
      }
      const existing = themeVarByInstance.get(themeValue);
      if (existing !== undefined) return existing;

      const varName = allocator.allocate(
        toCamelCaseIdentifier(themeValue.name, "theme", "Theme"),
      );
      themeVarByInstance.set(themeValue, varName);

      const themePayload: Record<string, unknown> = {
        name: themeValue.name,
        colors: themeValue.colors,
        fonts: themeValue.fonts,
        ...(themeValue.formatScheme === undefined
          ? {}
          : { formatScheme: themeValue.formatScheme }),
      };
      const formattedPayload = formatTsValue(themePayload, ctx, hooks);
      const defineThemeFn = useImport("defineTheme");
      themeDeclarations.push(
        `const ${varName} = ${defineThemeFn}(${formattedPayload});`,
      );
      return varName;
    };

    const createValueHooks = (ctx: CodegenContext): ValueHooks => {
      const hooks: ValueHooks = {
        serializeSpecial: (value, propPath) => {
          for (const serializer of registry.valueSerializers) {
            const serialized = serializer(value, ctx);
            if (serialized !== undefined) {
              return serialized.code;
            }
          }
          if (value === defaultTheme) {
            return useImport("defaultTheme");
          }
          if (isTheme(value)) {
            return serializeTheme(value, ctx, hooks);
          }
          if (isFont(value)) {
            return serializeFont(value, ctx, hooks);
          }
          if (isMorphKey(value)) {
            const scopeVar = morphScopeVarByKey.get(value);
            if (scopeVar === undefined) {
              ctx.fail(
                "unexpressible-value",
                "MorphKey handle was not registered in deck scope",
                propPath,
              );
            }
            return `${scopeVar}.handle(${JSON.stringify(getMorphKeyLabel(value))})`;
          }
          if (isSlideLayoutDefinition(value)) {
            const ident = layoutIdentByDefinition.get(value);
            if (ident === undefined) {
              ctx.fail(
                "unexpressible-value",
                "SlideLayoutDefinition is missing or mounted multiple times in this deck",
                propPath,
              );
            }
            return ident;
          }
          if (value instanceof SlideLayoutContainer) {
            const ident = layoutIdentByContainer.get(value);
            if (ident === undefined) {
              ctx.fail(
                "unexpressible-value",
                "SlideLayout is not owned by any SlideMaster in this deck",
                propPath,
              );
            }
            return ident;
          }
          return undefined;
        },
      };
      return hooks;
    };

    const makeContext = (path: string): CodegenContext => {
      let localHooks: ValueHooks | undefined;
      const getHooks = (): ValueHooks => {
        localHooks ??= createValueHooks(ctx);
        return localHooks;
      };
      const ctx: CodegenContext = {
        target,
        path,
        useImport,
        registerImageAsset: (bytes, ext) => assets.registerImage(bytes, ext),
        registerFontAsset: (bytes, family, ext) => {
          useImport("readFileSync", "node:fs");
          return assets.registerFont(bytes, family, ext);
        },
        allocateMotionScope: (identity) => {
          if (currentMotionScopeVars === undefined) {
            ctx.fail(
              "invalid-document",
              "MotionScope can only be used inside a Slide",
            );
          }
          const existing = currentMotionScopeVars.get(identity);
          if (existing !== undefined) {
            return existing;
          }
          const scopeAlloc = currentScopeAllocator ?? allocator;
          const varName = scopeAlloc.allocate("motionScopeHandle");
          currentMotionScopeVars.set(identity, varName);
          return varName;
        },
        emitValueExpr: (value, propPath) =>
          formatTsValue(value, ctx, getHooks(), propPath),
        emitChildNode: (node, index) => emitNode(node, path, index),
        emitRichTextChildren: (children) => emitChildrenList(children, path),
        emitObjectExtraProps: (props) => {
          const result: Record<string, CodegenExpr> = {};
          for (const [key, serializer] of registry.objectProps) {
            if (!Object.hasOwn(props, key) || props[key] === undefined) {
              continue;
            }
            const subCtx = makeContext(`${path}@${key}`);
            result[key] = serializer(props[key], subCtx);
          }
          return result;
        },
        formatProps: (entries) =>
          formatPropsObject(
            stripRedundantNativePlaceholder(
              entries,
              currentSlideLayoutPlaceholders,
            ),
            ctx,
            getHooks(),
          ),
        formatJsxElement: (tag, entries, children = []) =>
          formatJsxElement(
            tag,
            stripRedundantNativePlaceholder(
              entries,
              currentSlideLayoutPlaceholders,
            ),
            children,
            ctx,
            getHooks(),
          ),
        formatVanillaCall: (callee, args, children = []) =>
          formatVanillaCall(callee, args, children),
        fail: (code, message, subPath) => {
          failAt(resolveFailSubPath(path, subPath), code, message);
        },
      };
      return ctx;
    };

    const emitNode = (
      node: CommonNode,
      parentPath: string,
      index: number,
    ): string => {
      const nodePath = `${parentPath}/${childPathLabel(node, index)}`;
      const handler = registry.declarations.get(node.type);
      if (handler === undefined) {
        failAt(
          nodePath,
          "unsupported-declaration",
          `Declaration type ${JSON.stringify(node.type)} is not supported by registered codegen capabilities`,
          { declaration: node.type },
        );
      }
      const ctx = makeContext(nodePath);
      return target === "vanilla"
        ? handler.emitVanilla(node, ctx)
        : handler.emitReact(node, ctx);
    };

    const emitChildrenList = (
      children: readonly DeclarationChild[],
      parentPath: string,
    ): readonly string[] => {
      const emitted: string[] = [];
      let childNodeIndex = 0;
      for (const child of children) {
        if (child instanceof CommonNode) {
          emitted.push(emitNode(child, parentPath, childNodeIndex));
          childNodeIndex += 1;
        } else if (isDeclarationTextLeaf(child)) {
          if (target === "vanilla") {
            emitted.push(JSON.stringify(child.text));
          } else {
            // Multiple children are formatted on separate lines, so adjacent bare JSX text
            // would normalize newlines into spaces; wrap leaves in expression containers.
            emitted.push(
              children.length === 1
                ? formatJsxTextChild(child.text)
                : `{${JSON.stringify(child.text)}}`,
            );
          }
        }
      }
      return emitted;
    };

    const resolveOwnerSlotProp = (
      ownerPath: string,
      propName: string,
      propValue: unknown,
    ): CodegenExpr => {
      const handler =
        registry.slideProps.get(propName) ?? registry.ownerProps.get(propName);
      if (handler === undefined) {
        failAt(
          `${ownerPath}@${propName}`,
          "unsupported-slide-property",
          `Property ${JSON.stringify(propName)} on ${ownerPath} is not supported by registered codegen capabilities`,
          { field: propName },
        );
      }
      const ctx = makeContext(`${ownerPath}@${propName}`);
      return handler(propValue, ctx);
    };

    // Partition Deck children into SlideMasters and Section/Slide bodies
    const masterNodes: CommonNode[] = [];
    const bodyNodes: CommonNode[] = [];
    for (const child of document.children) {
      if (!(child instanceof CommonNode)) {
        if (isDeclarationTextLeaf(child) && child.text.trim() !== "") {
          failAt(
            "deck",
            "invalid-document",
            "Deck cannot directly contain text content",
          );
        }
        continue;
      }
      if (child.type === "slide-master") {
        masterNodes.push(child);
      } else {
        bodyNodes.push(child);
      }
    }

    // Validate Deck props
    for (const [key, value] of Object.entries(document.props)) {
      if (!isDefined(value)) continue;
      if (!coreDeckPropKeys.has(key)) {
        failAt(
          `deck@${key}`,
          "invalid-document",
          `Deck does not support property ${JSON.stringify(key)}`,
        );
      }
    }

    const usesSections = bodyNodes.some((node) => node.type === "section");
    if (usesSections) {
      for (const [index, node] of bodyNodes.entries()) {
        if (node.type !== "section") {
          failAt(
            `deck/${childPathLabel(node, index)}`,
            "invalid-document",
            "When using Section, every slide must belong to a Section",
          );
        }
      }
    } else {
      for (const [index, node] of bodyNodes.entries()) {
        if (node.type !== "slide") {
          failAt(
            `deck/${childPathLabel(node, index)}`,
            "invalid-document",
            "Direct children of Deck must be SlideMaster, Slide, or Section",
          );
        }
      }
    }

    // Section 2: SlideMasters & SlideLayouts
    const masterGroupSections: string[] = [];
    const masterRefsForDeck: string[] = [];

    for (const [masterIndex, masterNode] of masterNodes.entries()) {
      const masterPath = `deck/${childPathLabel(masterNode, masterIndex)}`;
      const masterRawName =
        masterNode.id ?? (masterNode.props["name"] as string | undefined);
      const masterVar =
        target === "vanilla"
          ? allocator.allocate(
              toCamelCaseIdentifier(masterRawName, `master${masterIndex + 1}`),
            )
          : allocator.allocate(
              toPascalCaseIdentifier(masterRawName, `Master${masterIndex + 1}`),
            );
      const masterPrefix =
        target === "vanilla"
          ? masterVar
          : `${masterVar.charAt(0).toLowerCase()}${masterVar.slice(1)}`;

      const masterPropsEntries: [string, unknown][] = [];
      for (const [key, value] of Object.entries(masterNode.props)) {
        if (!isDefined(value)) continue;
        if (coreMasterPropKeys.has(key)) {
          masterPropsEntries.push([key, value]);
        } else {
          masterPropsEntries.push([
            key,
            resolveOwnerSlotProp(masterPath, key, value),
          ]);
        }
      }

      const masterSharedCanvasChildren: CommonNode[] = [];
      const masterLayouts: SlideLayoutContainer[] = [];
      for (const child of masterNode.children) {
        if (!(child instanceof CommonNode)) {
          if (isDeclarationTextLeaf(child) && child.text.trim() !== "") {
            failAt(
              masterPath,
              "invalid-document",
              "SlideMaster cannot directly contain text content",
            );
          }
          continue;
        }
        if (child instanceof SlideLayoutContainer) {
          masterLayouts.push(child);
        } else {
          masterSharedCanvasChildren.push(child);
        }
      }

      const layoutDecls: string[] = [];
      const layoutExprsForMaster: string[] = [];
      for (const [layoutIndex, layoutNode] of masterLayouts.entries()) {
        const layoutPath = `${masterPath}/${childPathLabel(layoutNode, layoutIndex)}`;
        const layoutRawName =
          layoutNode.id ?? (layoutNode.props["name"] as string | undefined);
        const layoutIdent = allocator.allocate(
          toCamelCaseIdentifier(
            layoutRawName === undefined
              ? undefined
              : `${masterPrefix}_${layoutRawName}`,
            `${masterPrefix}Layout${layoutIndex + 1}`,
            "Layout",
          ),
        );
        layoutIdentByContainer.set(layoutNode, layoutIdent);
        const def = layoutNode.props["layout"];
        if (
          isSlideLayoutDefinition(def) &&
          layoutDefinitionCounts.get(def) === 1
        ) {
          layoutIdentByDefinition.set(def, layoutIdent);
        }

        const layoutSpecEntries: [string, unknown][] = [];
        for (const [key, value] of Object.entries(layoutNode.props)) {
          if (!isDefined(value) || key === "layout") continue;
          if (coreLayoutPropKeys.has(key)) {
            layoutSpecEntries.push([key, value]);
          } else {
            layoutSpecEntries.push([
              key,
              resolveOwnerSlotProp(layoutPath, key, value),
            ]);
          }
        }

        const layoutCtx = makeContext(layoutPath);
        const specCode =
          layoutSpecEntries.length === 0
            ? "{}"
            : layoutCtx.formatProps(layoutSpecEntries);
        const defineLayoutFn = useImport("defineSlideLayout");
        layoutDecls.push(
          `const ${layoutIdent} = ${defineLayoutFn}(${specCode});`,
        );

        const layoutChildren: string[] = [];
        let layoutChildIdx = 0;
        for (const child of layoutNode.children) {
          if (!(child instanceof CommonNode)) {
            if (isDeclarationTextLeaf(child) && child.text.trim() !== "") {
              failAt(
                layoutPath,
                "invalid-document",
                "SlideLayout cannot directly contain text content",
              );
            }
            continue;
          }
          layoutChildren.push(emitNode(child, layoutPath, layoutChildIdx));
          layoutChildIdx += 1;
        }

        if (target === "vanilla") {
          const slideLayoutFn = useImport("slideLayout");
          layoutExprsForMaster.push(
            layoutCtx.formatVanillaCall(
              slideLayoutFn,
              [layoutIdent],
              layoutChildren,
            ),
          );
        } else {
          const slideLayoutComp = useImport("SlideLayout");
          layoutExprsForMaster.push(
            layoutCtx.formatJsxElement(
              slideLayoutComp,
              [["layout", expr(layoutIdent)]],
              layoutChildren,
            ),
          );
        }
      }

      const masterSharedExprs = masterSharedCanvasChildren.map((child, idx) =>
        emitNode(child, masterPath, idx),
      );
      const allMasterChildren = [...masterSharedExprs, ...layoutExprsForMaster];

      const masterCtx = makeContext(masterPath);
      let masterDeclCode: string;
      if (target === "vanilla") {
        const slideMasterFn = useImport("slideMaster");
        const propsArg =
          masterPropsEntries.length === 0
            ? []
            : [masterCtx.formatProps(masterPropsEntries)];
        masterDeclCode = `const ${masterVar} = ${masterCtx.formatVanillaCall(
          slideMasterFn,
          propsArg,
          allMasterChildren,
        )};`;
        masterRefsForDeck.push(masterVar);
      } else {
        const slideMasterComp = useImport("SlideMaster");
        const jsxBody = masterCtx.formatJsxElement(
          slideMasterComp,
          masterPropsEntries,
          allMasterChildren,
        );
        masterDeclCode = `function ${masterVar}() {\n  return (\n${indentBlock(jsxBody, 4)}\n  );\n}`;
        masterRefsForDeck.push(`<${masterVar} />`);
      }

      masterGroupSections.push([...layoutDecls, masterDeclCode].join("\n\n"));
    }

    // Section 3: Slide declarations
    const slideDeclarations: string[] = [];

    const emitSlideDeclaration = (
      slideNode: CommonNode,
      slideIndex: number,
    ): string => {
      const slideLabel = slideNode.id ?? String(slideIndex + 1);
      const slidePath = `deck/slide[${slideLabel}]`;
      const slidePropsEntries: [string, unknown][] = [];
      for (const [key, value] of Object.entries(slideNode.props)) {
        if (!isDefined(value)) continue;
        if (key === "layout") {
          if (isSlideLayoutDefinition(value)) {
            if (!layoutIdentByDefinition.has(value)) {
              failAt(
                `${slidePath}@layout`,
                "unexpressible-value",
                "Slide references a SlideLayoutDefinition not uniquely owned by this deck",
              );
            }
          } else if (value instanceof SlideLayoutContainer) {
            if (!layoutIdentByContainer.has(value)) {
              failAt(
                `${slidePath}@layout`,
                "unexpressible-value",
                "Slide references a SlideLayout not owned by this deck",
              );
            }
          } else {
            failAt(
              `${slidePath}@layout`,
              "unexpressible-value",
              "Slide.layout must be a SlideLayout or SlideLayoutDefinition",
            );
          }
          slidePropsEntries.push([key, value]);
        } else if (coreSlidePropKeys.has(key)) {
          slidePropsEntries.push([key, value]);
        } else {
          slidePropsEntries.push([
            key,
            resolveOwnerSlotProp(slidePath, key, value),
          ]);
        }
      }

      const slideMotionVars = new Map<string, string>();
      currentMotionScopeVars = slideMotionVars;
      currentScopeAllocator = target === "react" ? allocator.fork() : allocator;
      const rawSlideLayout = slideNode.props["layout"];
      currentSlideLayoutPlaceholders = isSlideLayoutDefinition(rawSlideLayout)
        ? rawSlideLayout.placeholders
        : rawSlideLayout instanceof SlideLayoutContainer
          ? (rawSlideLayout.props["placeholders"] as PlaceholderMap | undefined)
          : undefined;

      const childExprs: string[] = [];
      try {
        let childIdx = 0;
        for (const child of slideNode.children) {
          if (!(child instanceof CommonNode)) {
            if (isDeclarationTextLeaf(child) && child.text.trim() !== "") {
              failAt(
                slidePath,
                "invalid-document",
                "Slide cannot directly contain text content",
              );
            }
            continue;
          }
          childExprs.push(emitNode(child, slidePath, childIdx));
          childIdx += 1;
        }
      } finally {
        currentMotionScopeVars = undefined;
        currentScopeAllocator = undefined;
        currentSlideLayoutPlaceholders = undefined;
      }

      const slideCtx = makeContext(slidePath);

      if (target === "vanilla") {
        const slideVar = allocator.allocate(
          toCamelCaseIdentifier(
            slideNode.id,
            `slide${slideIndex + 1}`,
            "Slide",
          ),
        );
        const motionDecls = [...slideMotionVars.values()].map(
          (handleVar) =>
            `const ${handleVar} = ${useImport("createMotionScopeHandle")}();`,
        );
        const slideFn = useImport("slide");
        const propsObj = slideCtx.formatProps(slidePropsEntries);
        const slideDecl = `const ${slideVar} = ${slideCtx.formatVanillaCall(
          slideFn,
          [propsObj],
          childExprs,
        )};`;
        slideDeclarations.push([...motionDecls, slideDecl].join("\n"));
        return slideVar;
      }

      const slideComponent = allocator.allocate(
        toPascalCaseIdentifier(slideNode.id, `Slide${slideIndex + 1}`, "Slide"),
      );
      const motionDecls = [...slideMotionVars.values()].map(
        (handleVar) => `const ${handleVar} = ${useImport("useMotionScope")}();`,
      );
      const slideComp = useImport("Slide");
      const slideJsx = slideCtx.formatJsxElement(
        slideComp,
        slidePropsEntries,
        childExprs,
      );
      const bodyLines = [
        ...motionDecls,
        "return (",
        indentBlock(slideJsx, 2),
        ");",
      ];
      slideDeclarations.push(
        `function ${slideComponent}() {\n${indentBlock(bodyLines.join("\n"), 2)}\n}`,
      );
      return `<${slideComponent} />`;
    };

    const bodyRefsForDeck: string[] = [];
    let slideCounter = 0;

    if (usesSections) {
      for (const [secIndex, sectionNode] of bodyNodes.entries()) {
        const rawSectionName = sectionNode.props["name"];
        const sectionPath =
          typeof rawSectionName === "string" && rawSectionName.trim() !== ""
            ? `deck/section[${rawSectionName}]`
            : `deck/section[${secIndex + 1}]`;
        for (const [key, value] of Object.entries(sectionNode.props)) {
          if (!isDefined(value)) continue;
          if (key !== "name") {
            failAt(
              `${sectionPath}@${key}`,
              "invalid-document",
              `Section does not support property ${JSON.stringify(key)}`,
            );
          }
        }
        if (
          typeof rawSectionName !== "string" ||
          rawSectionName.trim() === ""
        ) {
          failAt(
            sectionPath,
            "invalid-document",
            "Section.name must be a non-empty string",
            { field: "name" },
          );
        }
        const sectionSlideRefs: string[] = [];
        for (const [memberIdx, member] of sectionNode.children.entries()) {
          if (!(member instanceof CommonNode)) {
            if (isDeclarationTextLeaf(member) && member.text.trim() !== "") {
              failAt(
                sectionPath,
                "invalid-document",
                "Section can only contain Slide declarations",
              );
            }
            continue;
          }
          if (member.type !== "slide") {
            failAt(
              `${sectionPath}/${childPathLabel(member, memberIdx)}`,
              "invalid-document",
              "Section can only contain Slide declarations",
            );
          }
          sectionSlideRefs.push(emitSlideDeclaration(member, slideCounter));
          slideCounter += 1;
        }
        const secCtx = makeContext(sectionPath);
        if (target === "vanilla") {
          const sectionFn = useImport("section");
          const arg = secCtx.formatProps([["name", rawSectionName]]);
          bodyRefsForDeck.push(
            secCtx.formatVanillaCall(sectionFn, [arg], sectionSlideRefs),
          );
        } else {
          const sectionComp = useImport("Section");
          bodyRefsForDeck.push(
            secCtx.formatJsxElement(
              sectionComp,
              [["name", rawSectionName]],
              sectionSlideRefs,
            ),
          );
        }
      }
    } else {
      for (const slideNode of bodyNodes) {
        bodyRefsForDeck.push(emitSlideDeclaration(slideNode, slideCounter));
        slideCounter += 1;
      }
    }

    // Section 4: Default export Deck entry
    const deckPropsEntries: [string, unknown][] = [];
    for (const [key, value] of Object.entries(document.props)) {
      if (!isDefined(value)) continue;
      deckPropsEntries.push([key, value]);
    }
    const deckCtx = makeContext("deck");
    const deckChildren = [...masterRefsForDeck, ...bodyRefsForDeck];

    let section4Code: string;
    if (target === "vanilla") {
      const deckFn = useImport("deck");
      const deckArgs =
        deckPropsEntries.length === 0
          ? []
          : [deckCtx.formatProps(deckPropsEntries)];
      section4Code = `export default ${deckCtx.formatVanillaCall(
        deckFn,
        deckArgs,
        deckChildren,
      )};`;
    } else {
      const deckComp = useImport("Deck");
      const deckJsx = deckCtx.formatJsxElement(
        deckComp,
        deckPropsEntries,
        deckChildren,
      );
      section4Code = `export default function GeneratedDeck() {\n  return (\n${indentBlock(deckJsx, 4)}\n  );\n}`;
    }

    // Section 1: Imports, asset constants, fonts, themes, morphScopes
    const importLines: string[] = [];
    if (nodeFsImports.size > 0) {
      const sortedFs = [...nodeFsImports].toSorted();
      importLines.push(`import { ${sortedFs.join(", ")} } from "node:fs";`);
    }
    if (sdkImports.size > 0) {
      const sortedSdk = [...sdkImports].toSorted();
      const moduleSpecifier = `@officedk/pptxdk/${target}`;
      const singleLine = `import { ${sortedSdk.join(", ")} } from "${moduleSpecifier}";`;
      importLines.push(
        singleLine.length <= 80
          ? singleLine
          : `import {\n${indentBlock(sortedSdk.join(",\n"))},\n} from "${moduleSpecifier}";`,
      );
    }

    const assetDeclarations = assets.entries.map((asset) => {
      const urlExpr = `new URL(${JSON.stringify(`./${asset.relPath}`)}, import.meta.url)`;
      if (asset.kind === "image") {
        return `const ${asset.varName} = ${urlExpr};`;
      }
      return `const ${asset.varName} = new Uint8Array(\n  readFileSync(${urlExpr}),\n);`;
    });

    const morphDeclarations = morphBuckets.map(
      (bucket) => `const ${bucket.varName} = createMorphScope();`,
    );

    const section1Blocks = [
      importLines.join("\n"),
      assetDeclarations.join("\n"),
      fontDeclarations.join("\n"),
      themeDeclarations.join("\n"),
      morphDeclarations.join("\n"),
    ].filter((block) => block !== "");

    const fileSections = [
      section1Blocks.join("\n\n"),
      masterGroupSections.join("\n\n"),
      slideDeclarations.join("\n\n"),
      section4Code,
    ].filter((section) => section !== "");

    const entrySource = `${fileSections.join("\n\n")}\n`;

    const files: Record<string, CodegenFileContent> = {
      [entry]: entrySource,
    };
    for (const asset of assets.entries) {
      files[asset.relPath] = new Uint8Array(asset.bytes);
    }

    return Object.freeze({
      target,
      entry,
      files: Object.freeze(files),
    });
  };
}
