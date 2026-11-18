import type {
  CapabilityLayoutMeasurement,
  GeometryBox,
  LayoutLength,
  LayoutStyle,
} from "#/types.ts";

export type LayoutEdges = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};

export type LayoutMeasurementRequest = {
  readonly availableWidth?: number;
  readonly widthMode?: "at-most" | "exactly";
};

export type LayoutNodeMeasure = (
  request: LayoutMeasurementRequest,
) => CapabilityLayoutMeasurement | undefined;

export type LayoutNodeSpec = {
  readonly path: string;
  readonly style: LayoutStyle;
  readonly isContainer: boolean;
  readonly propagatesBaseline?: boolean;
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
 * Collaboration port exposed by the layout solver so delegated container modes (such as Grid)
 * can inspect computed sizes, request natural and min-content child measurements, apply resolved
 * constraints, and position children without depending directly on Yoga.
 */
export type LayoutSolverPort = {
  getComputedSize(node: LayoutSolverNode): {
    readonly width: number;
    readonly height: number;
  };
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

export type LayoutSolver = {
  createRoot(spec: {
    readonly path: string;
    readonly style: LayoutStyle;
  }): LayoutSolverNode;
  createNode(spec: LayoutNodeSpec, parent: LayoutSolverNode): LayoutSolverNode;
  appendChild(parent: LayoutSolverNode, child: LayoutSolverNode): void;
  requiresLooseLayout(style: LayoutStyle): boolean;
  positionLooseGeometry(node: LayoutSolverNode): void;
  solve(onCommit: (node: LayoutSolverNode, box: GeometryBox) => void): void;
};
