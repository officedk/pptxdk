import {
  blockFlexLayoutProperties,
  createBlockFlexLayoutSolver,
  isStyleRecord,
  paddingStyleKeys,
  validateBlockFlexContainerContext,
  validateBlockFlexContainerProperties,
  validateBlockFlexLayoutContextBounds,
} from "#/block-flex-layout.ts";
import {
  BuildError,
  CapabilityError,
  capabilityErrorToDiagnostic,
  toBuildDiagnostic,
} from "#/diagnostics.ts";
import { gridLayoutMode } from "#/grid-layout.ts";
import {
  expandedSpacing,
  type CustomContainerMode,
  type LayoutGridTracks,
  type LayoutSolver,
  type LayoutSolverNode,
} from "#/layout-protocol.ts";
import { validateLayoutStyleValues } from "#/layout-schemas.ts";
import { textStyleKeys } from "#/theme.ts";
import type {
  BoundContainerCollaboration,
  BoundObjectCapability,
  BoundObjectCapabilityModule,
  BuildDiagnostic,
  CapabilityChild,
  CapabilityElement,
  CapabilityExplicitPixelSize,
  CapabilityLayoutAxes,
  CapabilityLayoutContract,
  CapabilityLayoutMeasurement,
  LayoutStyle,
  NativeElementKind,
  PlaceholderResolutionContext,
  Rect,
  RegisteredCapabilityModule,
  ResolveOptions,
  ThemeContext,
} from "#/types.ts";

export { paddingStyleKeys };

type Dimensions = { readonly width: number; readonly height: number };

function roundLayoutMetric(value: number): number {
  const rounded = Math.round(value * 100) / 100;
  return Object.is(rounded, -0) ? 0 : rounded;
}

function roundRect(box: Rect): Rect {
  return {
    x: roundLayoutMetric(box.x),
    y: roundLayoutMetric(box.y),
    width: roundLayoutMetric(box.width),
    height: roundLayoutMetric(box.height),
  };
}

export type LayoutPassResult = {
  readonly boxes: WeakMap<CapabilityElement, Rect>;
  readonly sizedAxes: WeakMap<CapabilityElement, CapabilityLayoutAxes>;
  readonly gridTracks: WeakMap<CapabilityElement, LayoutGridTracks>;
  readonly gridAreas: WeakMap<CapabilityElement, Rect>;
  readonly themes: WeakMap<CapabilityElement, ThemeContext>;
  readonly hidden: WeakSet<CapabilityElement>;
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

function diagnostic(
  path: string,
  code: string,
  message: string,
  details?: Readonly<Record<string, unknown>>,
): BuildDiagnostic {
  return toBuildDiagnostic(path, code, message, details);
}

function fail(
  path: string,
  code: string,
  message: string,
  details?: Readonly<Record<string, unknown>>,
): never {
  throw new BuildError([diagnostic(path, code, message, details)]);
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
    fail(path, "invalid-style", "style must be an object", { field: "style" });
  }
  const unknown = Object.keys(value).find(
    (key) => !layoutPropertySet.has(key) && !allowedPresentation.has(key),
  );
  if (unknown !== undefined) {
    fail(
      path,
      "unknown-style",
      `style contains unsupported property ${unknown}`,
      { prop: unknown, field: `style.${unknown}`, scope: "style" },
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
  validateLayoutStyleValues(value, path);
  for (const mode of customContainerModes) {
    mode.validateStyleValues(value, path);
  }
  validateBlockFlexLayoutContextBounds(value, allowsPadding, path);
  return value;
}

function hasOppositeOffsets(
  style: LayoutStyle,
  startKey: "left" | "top",
  endKey: "right" | "bottom",
): boolean {
  if (style.position !== "absolute") return false;
  const start = style[startKey] ?? style.inset;
  const end = style[endKey] ?? style.inset;
  return (
    start !== undefined &&
    start !== "auto" &&
    end !== undefined &&
    end !== "auto"
  );
}

function computeFixedAxes(style: LayoutStyle): CapabilityLayoutAxes {
  return {
    width:
      (style.width !== undefined && style.width !== "auto") ||
      hasOppositeOffsets(style, "left", "right"),
    height:
      (style.height !== undefined && style.height !== "auto") ||
      hasOppositeOffsets(style, "top", "bottom"),
  };
}

function computeExplicitPixelSize(
  style: LayoutStyle,
): CapabilityExplicitPixelSize {
  const width =
    typeof style.width === "number" && Number.isFinite(style.width)
      ? style.width
      : undefined;
  const height =
    typeof style.height === "number" && Number.isFinite(style.height)
      ? style.height
      : undefined;
  return {
    ...(width === undefined ? {} : { width }),
    ...(height === undefined ? {} : { height }),
  };
}

function computeLayoutSizedAxes(
  contract: CapabilityLayoutContract,
  parentStyle: LayoutStyle,
): CapabilityLayoutAxes {
  const style = contract.layoutStyle;
  const fixed = contract.fixedAxes ?? computeFixedAxes(style);
  let width = fixed.width;
  let height = fixed.height;

  if (style.position !== "absolute") {
    if (parentStyle.display === "flex") {
      const direction = parentStyle.flexDirection ?? "row";
      const isRow = direction === "row" || direction === "row-reverse";
      const isMainFlexSized =
        style.flex !== undefined ||
        (style.flexGrow !== undefined && style.flexGrow > 0) ||
        (style.flexBasis !== undefined && style.flexBasis !== "auto");
      const effectiveAlign =
        style.alignSelf === undefined || style.alignSelf === "auto"
          ? (parentStyle.alignItems ?? "stretch")
          : style.alignSelf;
      const isCrossStretched = effectiveAlign === "stretch";
      if (isRow) {
        width ||= isMainFlexSized;
        height ||= isCrossStretched;
      } else {
        height ||= isMainFlexSized;
        width ||= isCrossStretched;
      }
    } else if (parentStyle.display === "grid") {
      const effectiveJustify =
        style.justifySelf === undefined || style.justifySelf === "auto"
          ? (parentStyle.justifyItems ?? "stretch")
          : style.justifySelf;
      const effectiveAlign =
        style.alignSelf === undefined || style.alignSelf === "auto"
          ? (parentStyle.alignItems ?? "stretch")
          : style.alignSelf;
      if (effectiveJustify === "stretch") width = true;
      if (effectiveAlign === "stretch" && style.aspectRatio === undefined) {
        height = true;
      }
    } else if (!contract.preservesIntrinsicSize) {
      width = true;
    }
  }

  if (style.aspectRatio !== undefined) {
    if (width && !height) height = true;
    else if (height && !width) width = true;
  }

  return { width, height };
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
  readonly excludedFromLayout?: boolean;
  readonly path?: string;
}): CapabilityLayoutContract {
  const {
    nativeKind,
    props,
    allowedStyleKeys,
    allowPadding = false,
    isContainer = false,
    excludedFromLayout = false,
  } = options;
  const path = options.path ?? "style";
  const rawStyle = props["style"];
  if (rawStyle !== undefined) {
    if (!isStyleRecord(rawStyle)) {
      throw new CapabilityError(
        "invalid-style",
        "style must be an object",
        options.path,
        { field: "style" },
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
          { field: `style.${paddingKey}` },
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
        { prop: unknown, field: `style.${unknown}`, scope: "style" },
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

  const explicitPixelSize = computeExplicitPixelSize(layoutStyle);
  const fixedAxes = computeFixedAxes(layoutStyle);
  const positioned =
    layoutStyle.position === "relative" || layoutStyle.position === "absolute";

  const isIntrinsicWidthObject =
    nativeKind === "pic" ||
    (nativeKind === "grpSp" &&
      !positioned &&
      (layoutStyle.width === undefined || layoutStyle.width === "auto"));

  return {
    nativeKind,
    layoutStyle,
    explicitPixelSize,
    fixedAxes,
    positioned,
    ...(padding === undefined ? {} : { padding }),
    ...(isIntrinsicWidthObject ? { preservesIntrinsicSize: true } : {}),
    ...(excludedFromLayout ? { excludedFromLayout: true } : {}),
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
};

/** Coordinates layout tree construction, capability binding, solver collaboration, and box commitment. */
export class LayoutPass {
  readonly #modules: ReadonlyMap<string, RegisteredCapabilityModule>;
  readonly #slideNumber: number;
  readonly #options: ResolveOptions;
  readonly #boxes = new WeakMap<CapabilityElement, Rect>();
  readonly #sizedAxes = new WeakMap<CapabilityElement, CapabilityLayoutAxes>();
  readonly #gridTracks = new WeakMap<CapabilityElement, LayoutGridTracks>();
  readonly #gridAreas = new WeakMap<CapabilityElement, Rect>();
  readonly #themes = new WeakMap<CapabilityElement, ThemeContext>();
  readonly #hidden = new WeakSet<CapabilityElement>();
  readonly #boundObjects = new WeakMap<
    CapabilityElement,
    BoundObjectCapability
  >();
  readonly #nodeMeta = new WeakMap<LayoutSolverNode, CommittedNodeMeta>();
  readonly #solver: LayoutSolver;
  #placeholderFor: (
    element: CapabilityElement,
  ) => PlaceholderResolutionContext | undefined = () => undefined;

  constructor(
    modules: ReadonlyMap<string, RegisteredCapabilityModule>,
    dimensions: Dimensions,
    slideNumber = 1,
    options: ResolveOptions = {},
  ) {
    this.#modules = modules;
    this.#slideNumber = slideNumber;
    this.#options = options;
    this.#solver = createBlockFlexLayoutSolver({
      dimensions,
      customContainers: customContainerModes.map((mode) => mode.createSolver()),
    });
  }

  run(
    children: readonly CapabilityChild[],
    parentPath: string,
    theme: ThemeContext,
    slideStyle?: unknown,
    placeholderFor?: (
      element: CapabilityElement,
    ) => PlaceholderResolutionContext | undefined,
  ): LayoutPassResult {
    this.#placeholderFor = placeholderFor ?? (() => undefined);
    const rootStyle = readStyle(
      slideStyle,
      parentPath,
      true,
      true,
      new Set([...textStyleKeys, "background"]),
    );
    const root = this.#solver.createRoot({
      path: parentPath,
      style: rootStyle,
    });
    this.#appendChildren(children, parentPath, root, theme);
    this.#solver.solve((node, box, meta) => {
      const elementMeta = this.#nodeMeta.get(node);
      if (elementMeta === undefined) return;
      this.#boxes.set(elementMeta.element, box);
      if (meta?.gridTracks !== undefined) {
        this.#gridTracks.set(elementMeta.element, {
          columnWidths: meta.gridTracks.columnWidths.map(roundLayoutMetric),
          rowHeights: meta.gridTracks.rowHeights.map(roundLayoutMetric),
        });
      }
      if (meta?.gridArea !== undefined) {
        this.#gridAreas.set(elementMeta.element, roundRect(meta.gridArea));
      }
    });
    return {
      boxes: this.#boxes,
      sizedAxes: this.#sizedAxes,
      gridTracks: this.#gridTracks,
      gridAreas: this.#gridAreas,
      themes: this.#themes,
      hidden: this.#hidden,
      bound: this.#boundObjects,
    };
  }

  #bindCapability(
    child: CapabilityElement,
    childPath: string,
    themeContext: ThemeContext,
    module: BoundObjectCapabilityModule,
    placeholder?: PlaceholderResolutionContext,
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
            ...(placeholder === undefined ? {} : { placeholder }),
          }),
        );
        this.#boundObjects.set(child, bound);
      } catch (error) {
        if (error instanceof BuildError) throw error;
        if (error instanceof CapabilityError) {
          throw new BuildError([
            capabilityErrorToDiagnostic(error, childPath, true),
          ]);
        }
        return fail(
          childPath,
          "capability-failure",
          error instanceof Error ? error.message : String(error),
        );
      }
    }
    return bound;
  }

  // Transparent containers are not in the layout tree, so theme context must be passed separately along the declaration tree.
  #appendChildren(
    children: readonly CapabilityChild[],
    parentPath: string,
    parent: LayoutSolverNode,
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
        this.#appendChildren(child.children, childPath, parent, themeContext);
        return;
      }
      if (objectModule(module)) {
        const placeholder = this.#placeholderFor(child);
        const elementTheme =
          placeholder === undefined ? themeContext : placeholder.masterTheme;
        const bound = this.#bindCapability(
          child,
          childPath,
          elementTheme,
          module,
          placeholder,
        );
        const contract = bound.contract;
        if (contract.layoutStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        if (placeholder !== undefined) {
          // Bound native placeholders resolve position/style from master+layout+local overrides
          // outside Yoga flow containers.
          this.#themes.set(child, elementTheme);
          return;
        }
        if (contract.excludedFromLayout === true) {
          this.#themes.set(child, themeContext);
          return;
        }
        for (const mode of customContainerModes) {
          mode.validateItemStyle(contract.layoutStyle, parent.style, childPath);
        }
        if (bound.container !== undefined) {
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
        const built = this.#buildLeafObject(
          child,
          childPath,
          parent,
          themeContext,
          bound,
          contract,
        );
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
    this.#sizedAxes.set(input, computeLayoutSizedAxes(contract, parent.style));
    const node = this.#solver.createNode(
      {
        path,
        style,
        isContainer: true,
        propagatesBaseline: container.propagatesBaseline !== false,
        ...(contract.preservesIntrinsicSize
          ? { preservesIntrinsicSize: true }
          : {}),
        ...(contract.padding === undefined
          ? {}
          : { padding: contract.padding }),
      },
      parent,
    );
    this.#nodeMeta.set(node, {
      element: input,
    });
    this.#appendChildren(container.children, path, node, theme);
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
    this.#sizedAxes.set(input, computeLayoutSizedAxes(contract, parent.style));
    const node = this.#solver.createNode(
      {
        path,
        style: contract.layoutStyle,
        isContainer: false,
        propagatesBaseline: false,
        ...(contract.preservesIntrinsicSize
          ? { preservesIntrinsicSize: true }
          : {}),
        ...(contract.contentMinimumsEnforced
          ? { contentMinimumsEnforced: true }
          : {}),
        ...(contract.padding === undefined
          ? {}
          : { padding: contract.padding }),
        ...(bound.measureLayout === undefined
          ? {}
          : {
              measure: (request: {
                readonly availableWidth?: number;
                readonly widthMode?: "at-most" | "exactly";
                readonly availableHeight?: number;
                readonly heightMode?: "at-most" | "exactly";
                readonly purpose?: "baseline" | "layout";
              }) => this.#measure(path, themeContext, bound, request),
            }),
      },
      parent,
    );
    this.#nodeMeta.set(node, {
      element: input,
    });
    return node;
  }

  #markHidden(child: CapabilityElement): void {
    this.#hidden.add(child);
    for (const descendant of child.children) {
      if (element(descendant)) this.#markHidden(descendant);
    }
  }

  #measure(
    path: string,
    theme: ThemeContext,
    bound: BoundObjectCapability,
    request: {
      readonly availableWidth?: number;
      readonly widthMode?: "at-most" | "exactly";
      readonly availableHeight?: number;
      readonly heightMode?: "at-most" | "exactly";
      readonly purpose?: "baseline" | "layout";
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
      if (error instanceof CapabilityError) {
        throw new BuildError([capabilityErrorToDiagnostic(error, path, true)]);
      }
      return fail(
        path,
        "capability-failure",
        error instanceof Error ? error.message : String(error),
      );
    }
  }
}
