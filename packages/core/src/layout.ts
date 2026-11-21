import {
  blockFlexLayoutProperties,
  createBlockFlexLayoutSolver,
  isStyleRecord,
  paddingStyleKeys,
  validateBlockFlexContainerContext,
  validateBlockFlexContainerProperties,
  validateBlockFlexDimensionsAndOffsets,
  validateBlockFlexEnumValues,
  validateBlockFlexSpacingAndBounds,
} from "#/block-flex-layout.ts";
import { BuildError, CapabilityError } from "#/diagnostics.ts";
import { gridLayoutMode } from "#/grid-layout.ts";
import {
  expandedSpacing,
  type CustomContainerMode,
  type LayoutSolver,
  type LayoutSolverNode,
} from "#/layout-protocol.ts";
import { textStyleKeys } from "#/theme.ts";
import type {
  BoundContainerCollaboration,
  BoundObjectCapability,
  BoundObjectCapabilityModule,
  CapabilityChild,
  CapabilityElement,
  CapabilityLayoutContract,
  CapabilityLayoutMeasurement,
  Diagnostic,
  GeometryBox,
  LayoutStyle,
  NativeElementKind,
  RegisteredCapabilityModule,
  ResolveOptions,
  ThemeContext,
} from "#/types.ts";

export { paddingStyleKeys };

type Dimensions = { readonly width: number; readonly height: number };

export type LayoutPassResult = {
  readonly boxes: WeakMap<CapabilityElement, GeometryBox>;
  readonly themes: WeakMap<CapabilityElement, ThemeContext>;
  readonly hidden: WeakSet<CapabilityElement>;
  readonly warnings: readonly Diagnostic[];
  readonly bound: WeakMap<CapabilityElement, BoundObjectCapability>;
};

const customContainerModes: readonly CustomContainerMode[] = [gridLayoutMode];

const layoutProperties = [
  ...blockFlexLayoutProperties,
  ...customContainerModes.flatMap((mode) => [
    ...mode.containerStyleKeys,
    ...mode.itemStyleKeys,
  ]),
] as const;

const layoutPropertySet = new Set<string>(layoutProperties);

export const layoutStyleKeys = layoutProperties.filter(
  (key) => !(paddingStyleKeys as readonly string[]).includes(key),
);

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

function fail(path: string, code: string, message: string): never {
  throw new BuildError([diagnostic(path, code, message)]);
}

function element(child: CapabilityChild): child is CapabilityElement {
  return "props" in child;
}

function identity(element: CapabilityElement): string | undefined {
  const id = element.props["id"];
  return typeof id === "string" && id.trim() !== "" ? id : undefined;
}

function label(element: CapabilityElement, index: number): string {
  return `${element.type}[${identity(element) ?? String(index + 1)}]`;
}

function readStyle(
  value: unknown,
  path: string,
  isContainer: boolean,
  allowsPadding = isContainer,
  allowedPresentation: ReadonlySet<string> = new Set(),
): LayoutStyle {
  if (value === undefined) return {};
  if (!isStyleRecord(value)) {
    fail(path, "invalid-layout-style", "style must be an object");
  }
  const unknown = Object.keys(value).find(
    (key) => !layoutPropertySet.has(key) && !allowedPresentation.has(key),
  );
  if (unknown !== undefined) {
    fail(
      path,
      "unsupported-layout-style",
      `style contains unsupported property ${unknown}`,
    );
  }
  const resolvedDisplay = validateBlockFlexContainerContext(
    value,
    path,
    isContainer,
    customContainerModes,
  );
  validateBlockFlexContainerProperties(
    value,
    resolvedDisplay,
    path,
    customContainerModes,
  );
  for (const mode of customContainerModes) {
    mode.validateContainerStyle(value, resolvedDisplay, path);
  }
  validateBlockFlexEnumValues(value, path);
  for (const mode of customContainerModes) {
    mode.validateStyleValues(value, path);
  }
  validateBlockFlexDimensionsAndOffsets(value, allowsPadding, path);
  validateBlockFlexSpacingAndBounds(value, allowsPadding, path);
  return value;
}

/**
 * Parses and validates host props into a CapabilityLayoutContract for object capabilities.
 */
export function parseCapabilityLayoutContract(options: {
  readonly nativeKind: NativeElementKind;
  readonly props: Readonly<Record<string, unknown>>;
  readonly allowedStyleKeys: ReadonlySet<string>;
  readonly allowPadding?: boolean;
  readonly isContainer?: boolean;
  readonly path?: string;
}): CapabilityLayoutContract {
  const {
    nativeKind,
    props,
    allowedStyleKeys,
    allowPadding = false,
    isContainer = false,
  } = options;
  const path = options.path ?? "style";
  const rawStyle = props["style"];
  if (rawStyle !== undefined) {
    if (!isStyleRecord(rawStyle)) {
      throw new CapabilityError(
        "invalid-style",
        "style must be an object",
        options.path,
      );
    }
    if (!allowPadding) {
      const paddingKey = (paddingStyleKeys as readonly string[]).find(
        (key) => rawStyle[key] !== undefined,
      );
      if (paddingKey !== undefined) {
        throw new CapabilityError(
          "invalid-layout-context",
          "padding is only applicable to layout containers",
          options.path,
        );
      }
    }
    const unknown = Object.keys(rawStyle).find(
      (key) => rawStyle[key] !== undefined && !allowedStyleKeys.has(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-style",
        `style contains unsupported property ${unknown}`,
        options.path,
      );
    }
  }

  const layoutStyle = readStyle(
    rawStyle,
    path,
    isContainer,
    allowPadding,
    allowedStyleKeys,
  );

  let padding:
    | {
        readonly top: number;
        readonly right: number;
        readonly bottom: number;
        readonly left: number;
      }
    | undefined;
  if (allowPadding && rawStyle !== undefined) {
    const hasPadding = (paddingStyleKeys as readonly string[]).some(
      (key) => rawStyle[key] !== undefined,
    );
    if (hasPadding) {
      const edges = expandedSpacing(layoutStyle, "padding");
      padding = {
        top: edges.top,
        right: edges.right,
        bottom: edges.bottom,
        left: edges.left,
      };
    }
  }

  const x = typeof props["x"] === "number" ? props["x"] : undefined;
  const y = typeof props["y"] === "number" ? props["y"] : undefined;
  const styleWidth = rawStyle !== undefined ? rawStyle["width"] : undefined;
  const styleHeight = rawStyle !== undefined ? rawStyle["height"] : undefined;
  const width = (styleWidth ?? props["width"]) as
    | number
    | `${number}%`
    | "auto"
    | undefined;
  const height = (styleHeight ?? props["height"]) as
    | number
    | `${number}%`
    | "auto"
    | undefined;

  return {
    nativeKind,
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
    layoutStyle,
    ...(padding === undefined ? {} : { padding }),
  };
}

const nativeElementKinds = new Set<string>([
  "sp",
  "pic",
  "grpSp",
  "graphicFrame",
  "cxnSp",
]);

export function validateBoundObjectCapability(
  bound: unknown,
): BoundObjectCapability {
  if (typeof bound !== "object" || bound === null) {
    throw new CapabilityError(
      "invalid-capability-binding",
      "Object capability bind() must return a bound capability object",
    );
  }
  const candidate = bound as {
    readonly contract?: Partial<CapabilityLayoutContract>;
    readonly resolve?: unknown;
  };
  if (
    typeof candidate.contract !== "object" ||
    candidate.contract === null ||
    candidate.contract.nativeKind === undefined ||
    !nativeElementKinds.has(candidate.contract.nativeKind) ||
    typeof candidate.contract.layoutStyle !== "object" ||
    candidate.contract.layoutStyle === null ||
    Array.isArray(candidate.contract.layoutStyle) ||
    typeof candidate.resolve !== "function"
  ) {
    throw new CapabilityError(
      "invalid-capability-binding",
      "Object capability bind() must return a valid contract and resolve handler",
    );
  }
  return bound as BoundObjectCapability;
}

function objectModule(
  module: RegisteredCapabilityModule | undefined,
): module is BoundObjectCapabilityModule {
  return module !== undefined && "bind" in module;
}

type CommittedNodeMeta = {
  readonly element: CapabilityElement;
  readonly checksCanvasOverflow: boolean;
};

/** Coordinates layout tree construction, capability binding, solver collaboration, and box commitment. */
export class LayoutPass {
  readonly #modules: ReadonlyMap<string, RegisteredCapabilityModule>;
  readonly #dimensions: Dimensions;
  readonly #slideNumber: number;
  readonly #options: ResolveOptions;
  readonly #boxes = new WeakMap<CapabilityElement, GeometryBox>();
  readonly #themes = new WeakMap<CapabilityElement, ThemeContext>();
  readonly #hidden = new WeakSet<CapabilityElement>();
  readonly #boundObjects = new WeakMap<
    CapabilityElement,
    BoundObjectCapability
  >();
  readonly #nodeMeta = new WeakMap<LayoutSolverNode, CommittedNodeMeta>();
  readonly #warnings: Diagnostic[] = [];
  readonly #warningKeys = new Set<string>();
  readonly #solver: LayoutSolver;

  constructor(
    modules: ReadonlyMap<string, RegisteredCapabilityModule>,
    dimensions: Dimensions,
    slideNumber = 1,
    options: ResolveOptions = {},
  ) {
    this.#modules = modules;
    this.#dimensions = dimensions;
    this.#slideNumber = slideNumber;
    this.#options = options;
    this.#solver = createBlockFlexLayoutSolver({
      dimensions,
      customContainers: customContainerModes.map((mode) =>
        mode.createSolver((path, code, message) => {
          this.#warnOnce(path, code, message);
        }),
      ),
    });
  }

  run(
    children: readonly CapabilityChild[],
    parentPath: string,
    theme: ThemeContext,
    slideStyle?: unknown,
  ): LayoutPassResult {
    const rootStyle = readStyle(
      slideStyle,
      parentPath,
      true,
      true,
      textStyleKeys,
    );
    const root = this.#solver.createRoot({
      path: parentPath,
      style: rootStyle,
    });
    this.#appendChildren(
      children,
      parentPath,
      root,
      slideStyle !== undefined,
      theme,
    );
    this.#solver.solve((node, box) => {
      const meta = this.#nodeMeta.get(node);
      if (meta === undefined) return;
      this.#boxes.set(meta.element, box);
      if (
        meta.checksCanvasOverflow &&
        (box.x < 0 ||
          box.y < 0 ||
          box.x + box.width > this.#dimensions.width ||
          box.y + box.height > this.#dimensions.height)
      ) {
        this.#warnings.push({
          severity: "warning",
          path: node.path,
          code: "overflow",
          message:
            "Layout container overflows slide canvas; build retains original declaration",
        });
      }
    });
    return {
      boxes: this.#boxes,
      themes: this.#themes,
      hidden: this.#hidden,
      warnings: this.#warnings,
      bound: this.#boundObjects,
    };
  }

  #bindCapability(
    child: CapabilityElement,
    childPath: string,
    themeContext: ThemeContext,
    module: BoundObjectCapabilityModule,
  ): BoundObjectCapability {
    let bound = this.#boundObjects.get(child);
    if (bound === undefined) {
      try {
        bound = validateBoundObjectCapability(
          module.bind(child, {
            path: childPath,
            theme: themeContext,
            options: this.#options,
            slideNumber: this.#slideNumber,
          }),
        );
        this.#boundObjects.set(child, bound);
      } catch (error) {
        if (error instanceof BuildError) throw error;
        const code =
          error instanceof CapabilityError
            ? error.code
            : "invalid-capability-input";
        const message = error instanceof Error ? error.message : String(error);
        const targetPath =
          error instanceof CapabilityError && error.path !== undefined
            ? error.path.startsWith(childPath)
              ? error.path
              : `${childPath}/${error.path}`
            : childPath;
        return fail(targetPath, code, message);
      }
    }
    return bound;
  }

  // Transparent or loose containers are not in the layout tree, so theme context must be passed separately along the declaration tree.
  #appendChildren(
    children: readonly CapabilityChild[],
    parentPath: string,
    parent: LayoutSolverNode,
    includeObjects: boolean,
    themeContext: ThemeContext,
  ): void {
    children.forEach((child, index) => {
      if (!element(child)) return;
      const childPath = `${parentPath}/${label(child, index)}`;
      if (parent.style.display === "none") {
        this.#markHidden(child);
        return;
      }
      const module = this.#modules.get(child.type);
      if (module !== undefined && "compileDeclaration" in module) {
        this.#themes.set(child, themeContext);
        // An owned subtree belongs to its capability and never becomes layout content.
        if (module.children === "owned") return;
        this.#appendChildren(
          child.children,
          childPath,
          parent,
          includeObjects,
          themeContext,
        );
        return;
      }
      if (objectModule(module)) {
        const bound = this.#bindCapability(
          child,
          childPath,
          themeContext,
          module,
        );
        const contract = bound.contract;
        if (contract.layoutStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        for (const mode of customContainerModes) {
          mode.validateItemStyle(contract.layoutStyle, parent.style, childPath);
        }
        if (bound.container !== undefined) {
          if (
            !includeObjects &&
            bound.container.layoutMode === "when-styled" &&
            !Object.keys(contract.layoutStyle).some(
              (key) =>
                key !== "zIndex" &&
                layoutPropertySet.has(key) &&
                contract.layoutStyle[key as keyof LayoutStyle] !== undefined,
            )
          ) {
            const theme = bound.container.theme ?? themeContext;
            this.#themes.set(child, theme);
            this.#appendChildren(
              bound.container.children,
              childPath,
              parent,
              false,
              theme,
            );
            return;
          }
          const built = this.#buildContainer(
            child,
            childPath,
            parent,
            themeContext,
            contract,
            bound.container,
          );
          this.#solver.appendChild(parent, built);
          return;
        }
        this.#themes.set(child, themeContext);
        if (
          !includeObjects &&
          !this.#solver.requiresLooseLayout(contract.layoutStyle)
        ) {
          return;
        }
        const built = this.#buildLeafObject(
          child,
          childPath,
          parent,
          themeContext,
          bound,
          contract,
        );
        if (!includeObjects) {
          this.#solver.positionLooseGeometry(built);
        }
        this.#solver.appendChild(parent, built);
      }
    });
  }

  #buildContainer(
    input: CapabilityElement,
    path: string,
    parent: LayoutSolverNode,
    themeContext: ThemeContext,
    contract: CapabilityLayoutContract,
    container: BoundContainerCollaboration,
  ): LayoutSolverNode {
    const style = contract.layoutStyle;
    const theme = container.theme ?? themeContext;
    this.#themes.set(input, theme);
    const node = this.#solver.createNode(
      {
        path,
        style,
        isContainer: true,
        propagatesBaseline: container.propagatesBaseline !== false,
        ...(contract.width === undefined
          ? {}
          : { contractWidth: contract.width }),
        ...(contract.height === undefined
          ? {}
          : { contractHeight: contract.height }),
        ...(contract.x === undefined ? {} : { contractX: contract.x }),
        ...(contract.y === undefined ? {} : { contractY: contract.y }),
        ...(contract.padding === undefined
          ? {}
          : { padding: contract.padding }),
      },
      parent,
    );
    this.#nodeMeta.set(node, {
      element: input,
      checksCanvasOverflow: container.checksCanvasOverflow !== false,
    });
    this.#appendChildren(container.children, path, node, true, theme);
    return node;
  }

  #buildLeafObject(
    input: CapabilityElement,
    path: string,
    parent: LayoutSolverNode,
    themeContext: ThemeContext,
    bound: BoundObjectCapability,
    contract: CapabilityLayoutContract,
  ): LayoutSolverNode {
    const node = this.#solver.createNode(
      {
        path,
        style: contract.layoutStyle,
        isContainer: false,
        propagatesBaseline: false,
        ...(contract.width === undefined
          ? {}
          : { contractWidth: contract.width }),
        ...(contract.height === undefined
          ? {}
          : { contractHeight: contract.height }),
        ...(contract.x === undefined ? {} : { contractX: contract.x }),
        ...(contract.y === undefined ? {} : { contractY: contract.y }),
        ...(contract.padding === undefined
          ? {}
          : { padding: contract.padding }),
        ...(bound.measureLayout === undefined
          ? {}
          : {
              measure: (request: {
                readonly availableWidth?: number;
                readonly widthMode?: "at-most" | "exactly";
              }) => this.#measure(path, themeContext, bound, request),
            }),
      },
      parent,
    );
    this.#nodeMeta.set(node, {
      element: input,
      checksCanvasOverflow: false,
    });
    return node;
  }

  #markHidden(child: CapabilityElement): void {
    this.#hidden.add(child);
    for (const descendant of child.children) {
      if (element(descendant)) this.#markHidden(descendant);
    }
  }

  #warnOnce(path: string, code: string, message: string): void {
    const key = `${path}\u0000${code}\u0000${message}`;
    if (this.#warningKeys.has(key)) return;
    this.#warningKeys.add(key);
    this.#warnings.push({ severity: "warning", path, code, message });
  }

  #measure(
    path: string,
    theme: ThemeContext,
    bound: BoundObjectCapability,
    request: {
      readonly availableWidth?: number;
      readonly widthMode?: "at-most" | "exactly";
    },
  ): CapabilityLayoutMeasurement | undefined {
    if (bound.measureLayout === undefined) {
      return undefined;
    }
    try {
      return bound.measureLayout({
        ...request,
        theme,
        slideNumber: this.#slideNumber,
      });
    } catch (error) {
      if (error instanceof BuildError) throw error;
      const code =
        error instanceof CapabilityError
          ? error.code
          : "layout-measurement-failed";
      const message = error instanceof Error ? error.message : String(error);
      const targetPath =
        error instanceof CapabilityError && error.path !== undefined
          ? error.path.startsWith(path)
            ? error.path
            : `${path}/${error.path}`
          : path;
      return fail(targetPath, code, message);
    }
  }
}
