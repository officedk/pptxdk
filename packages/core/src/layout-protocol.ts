import { BuildError, toBuildDiagnostic } from "#/diagnostics.ts";
import type {
  CapabilityLayoutMeasurement,
  LayoutLength,
  LayoutStyle,
  Rect,
} from "#/types.ts";

export type LayoutGridTracks = {
  readonly columnWidths: readonly number[];
  readonly rowHeights: readonly number[];
};

export type LayoutCommitMetadata = {
  readonly gridTracks?: LayoutGridTracks;
  readonly gridArea?: Rect;
};

export type LayoutEdges = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};

export type LayoutMeasurementRequest = {
  readonly availableWidth?: number;
  readonly widthMode?: "at-most" | "exactly";
  readonly availableHeight?: number;
  readonly heightMode?: "at-most" | "exactly";
  readonly purpose?: "baseline" | "layout";
};

export type LayoutNodeMeasure = (
  request: LayoutMeasurementRequest,
) => CapabilityLayoutMeasurement | undefined;

export type LayoutNodeSpec = {
  readonly path: string;
  readonly style: LayoutStyle;
  readonly isContainer: boolean;
  readonly propagatesBaseline?: boolean;
  readonly preservesIntrinsicSize?: boolean;
  readonly contentMinimumsEnforced?: boolean;
  readonly contractWidth?: LayoutLength;
  readonly contractHeight?: LayoutLength;
  readonly contractX?: number;
  readonly contractY?: number;
  readonly padding?: LayoutEdges;
  readonly measure?: LayoutNodeMeasure;
};

export type LayoutSolverNode = {
  readonly path: string;
  readonly style: LayoutStyle;
  readonly isRoot: boolean;
  readonly isContainer: boolean;
  readonly contentMinimumsEnforced?: boolean;
  readonly contractWidth?: LayoutLength;
  readonly contractHeight?: LayoutLength;
  readonly padding?: LayoutEdges;
  readonly children: readonly LayoutSolverNode[];
};

export type LayoutChildConstraints = {
  readonly width?: number;
  readonly height?: number | "auto";
  readonly minWidth?: number;
  readonly maxWidth?: number;
  readonly minHeight?: number;
  readonly maxHeight?: number;
};

/**
 * Collaboration port exposed by the layout solver so delegated container modes
 * can inspect computed sizes, share allocated area heights across passes, request
 * natural and min-content child measurements, apply resolved constraints, and position
 * children without depending directly on the underlying solver.
 */
export type LayoutSolverPort = {
  getComputedSize(node: LayoutSolverNode): {
    readonly width: number;
    readonly height: number;
  };
  getAllocatedHeight(node: LayoutSolverNode): number | undefined;
  setAllocatedHeight(node: LayoutSolverNode, height: number): void;
  setContainerSize(
    node: LayoutSolverNode,
    size: { readonly width?: number; readonly height?: number },
  ): void;
  measureNaturalSize(
    node: LayoutSolverNode,
    availableWidth?: number,
  ): { readonly width: number; readonly height: number };
  measureMinContentContribution(
    node: LayoutSolverNode,
    axis: "width" | "height",
  ): number;
  setChildConstraints(
    node: LayoutSolverNode,
    constraints: LayoutChildConstraints,
  ): void;
  placeChild(
    node: LayoutSolverNode,
    position: { readonly x: number; readonly y: number },
  ): void;
  setGridTracks?(node: LayoutSolverNode, tracks: LayoutGridTracks): void;
  setGridCellArea?(node: LayoutSolverNode, area: Rect): void;
};

/**
 * Collaboration seam for container display modes delegated out of the Block/Flex solver.
 */
export type CustomContainerSolver = {
  handlesDisplay(display: LayoutStyle["display"]): boolean;
  measureMinContent(
    node: LayoutSolverNode,
    axis: "width" | "height",
    port: LayoutSolverPort,
  ): number;
  placeContainer(node: LayoutSolverNode, port: LayoutSolverPort): void;
};

/**
 * Ownership descriptor for a delegated container layout mode, including style
 * vocabulary, context/value validation, and solver creation.
 */
export type CustomContainerMode = {
  readonly display: NonNullable<LayoutStyle["display"]>;
  readonly containerStyleKeys: readonly string[];
  readonly itemStyleKeys: readonly string[];
  readonly sharedContainerProperties: readonly string[];
  validateContainerStyle(
    style: Readonly<Record<string, unknown>>,
    resolvedDisplay: NonNullable<LayoutStyle["display"]>,
    path: string,
  ): void;
  validateStyleValues(
    style: Readonly<Record<string, unknown>>,
    path: string,
  ): void;
  validateItemStyle(
    style: LayoutStyle,
    parentStyle: LayoutStyle,
    path: string,
  ): void;
  createSolver(
    warnOnce: (path: string, code: string, message: string) => void,
  ): CustomContainerSolver;
};

export type LayoutSolver = {
  createRoot(spec: {
    readonly path: string;
    readonly style: LayoutStyle;
  }): LayoutSolverNode;
  createNode(spec: LayoutNodeSpec, parent: LayoutSolverNode): LayoutSolverNode;
  appendChild(parent: LayoutSolverNode, child: LayoutSolverNode): void;
  solve(
    onCommit: (
      node: LayoutSolverNode,
      box: Rect,
      meta?: LayoutCommitMetadata,
    ) => void,
  ): void;
};

function failLayoutProtocol(
  path: string,
  code: string,
  message: string,
  field?: string,
): never {
  throw new BuildError([
    toBuildDiagnostic(
      path,
      code,
      message,
      field === undefined ? undefined : { field },
    ),
  ]);
}

export function parseLayoutNumber(
  value: unknown,
  name: string,
  path: string,
): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    failLayoutProtocol(
      path,
      "invalid-layout-style",
      `${name} must be a finite pixel number`,
      `style.${name}`,
    );
  }
  return value;
}

export function parseLayoutSpacing(
  value: unknown,
  name: string,
  path: string,
): LayoutEdges {
  if (value === undefined) return { top: 0, right: 0, bottom: 0, left: 0 };
  if (typeof value === "number") {
    const amount = parseLayoutNumber(value, name, path);
    return { top: amount, right: amount, bottom: amount, left: amount };
  }
  if (!Array.isArray(value) || (value.length !== 2 && value.length !== 4)) {
    failLayoutProtocol(
      path,
      "invalid-layout-style",
      `${name} must be a number, 2-tuple, or 4-tuple`,
    );
  }
  const values = value.map((item) => parseLayoutNumber(item, name, path));
  if (values.length === 2) {
    return {
      top: values[0]!,
      right: values[1]!,
      bottom: values[0]!,
      left: values[1]!,
    };
  }
  return {
    top: values[0]!,
    right: values[1]!,
    bottom: values[2]!,
    left: values[3]!,
  };
}

export function expandedSpacing(
  style: LayoutStyle,
  kind: "margin" | "padding",
): LayoutEdges {
  const base = parseLayoutSpacing(style[kind], kind, "layout");
  if (kind === "margin") {
    return {
      top: style.marginTop ?? base.top,
      right: style.marginRight ?? base.right,
      bottom: style.marginBottom ?? base.bottom,
      left: style.marginLeft ?? base.left,
    };
  }
  return {
    top: style.paddingTop ?? base.top,
    right: style.paddingRight ?? base.right,
    bottom: style.paddingBottom ?? base.bottom,
    left: style.paddingLeft ?? base.left,
  };
}

export function resolvePercent(
  value: LayoutLength | undefined,
  basis: number,
): number | undefined {
  if (typeof value !== "string") return undefined;
  const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+))%$/u.exec(value);
  return match === null ? undefined : (Number(match[1]) * basis) / 100;
}
