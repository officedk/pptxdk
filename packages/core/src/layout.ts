import {
  blockFlexLayoutProperties,
  createBlockFlexLayoutSolver,
  describeStyleValue,
  expandedSpacing,
  isStyleRecord,
  paddingStyleKeys,
  resolvePercent,
  validateBlockFlexContainerContext,
  validateBlockFlexContainerProperties,
  validateBlockFlexDimensionsAndOffsets,
  validateBlockFlexEnumValues,
  validateBlockFlexSpacingAndBounds,
} from "#/block-flex-layout.ts";
import { BuildError, CapabilityError } from "#/diagnostics.ts";
import type {
  CustomContainerSolver,
  LayoutSolver,
  LayoutSolverNode,
  LayoutSolverPort,
} from "#/layout-protocol.ts";
import { textStyleFrom, textStyleKeys } from "#/theme.ts";
import type {
  BoundContainerCollaboration,
  BoundObjectCapability,
  BoundObjectCapabilityModule,
  BuildOptions,
  CapabilityChild,
  CapabilityElement,
  CapabilityLayoutContract,
  CapabilityLayoutMeasurement,
  Diagnostic,
  GeometryBox,
  LayoutLength,
  LayoutStyle,
  NativeElementKind,
  RegisteredCapabilityModule,
  ThemeContext,
  ThemeScope,
} from "#/types.ts";

export { paddingStyleKeys };

type Dimensions = { readonly width: number; readonly height: number };

// Grid-owned layout vocabulary and solver helpers stay isolated here until Ticket 07 extracts Grid.
type GridTrack =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "auto"; readonly min: number }
  | {
      readonly kind: "fr";
      readonly value: number;
      readonly min: number;
      readonly contentMinimum?: boolean;
    };
type GridPlacement = {
  readonly node: LayoutSolverNode;
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly columnSpan: number;
};
type ContentGrowableGridTrack = Extract<
  GridTrack,
  { readonly kind: "auto" | "fr" }
>;
type GridModel = {
  readonly placements: readonly GridPlacement[];
  readonly widths: readonly number[];
  readonly heights: readonly number[];
  readonly columnGap: number;
  readonly rowGap: number;
  readonly spanWidth: (start: number, span: number) => number;
  readonly spanHeight: (start: number, span: number) => number;
  readonly width: number;
  readonly height: number;
};

export type LayoutPassResult = {
  readonly boxes: WeakMap<CapabilityElement, GeometryBox>;
  readonly themes: WeakMap<CapabilityElement, ThemeContext>;
  readonly hidden: WeakSet<CapabilityElement>;
  readonly warnings: readonly Diagnostic[];
  readonly bound: WeakMap<CapabilityElement, BoundObjectCapability>;
};

const gridContainerProperties = [
  "gridTemplateColumns",
  "gridTemplateRows",
  "gridAutoRows",
  "justifyItems",
] as const;
const gridItemProperties = ["gridColumn", "gridRow", "justifySelf"] as const;
const gridLayoutProperties = [
  ...gridContainerProperties,
  ...gridItemProperties,
] as const;

const layoutProperties = [
  ...blockFlexLayoutProperties,
  ...gridLayoutProperties,
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

function deriveTheme(
  parent: ThemeContext,
  scope: ThemeScope | undefined,
  style: unknown,
  path: string,
): ThemeContext {
  try {
    const textStyle = style === undefined ? {} : textStyleFrom(style);
    return parent.with(scope, textStyle);
  } catch (error) {
    return fail(
      path,
      error instanceof CapabilityError ? error.code : "invalid-theme-scope",
      error instanceof Error ? error.message : String(error),
    );
  }
}

function element(child: CapabilityChild): child is CapabilityElement {
  return "props" in child;
}

function identity(element: CapabilityElement): string | undefined {
  const id = element.props["id"];
  return typeof id === "string" && id.trim() !== "" ? id : undefined;
}

function label(element: CapabilityElement, index: number): string {
  const type = element.type.startsWith("rp:")
    ? element.type.slice(3)
    : element.type;
  return `${type}[${identity(element) ?? String(index + 1)}]`;
}

function gridTokens(input: string, path: string): string[] {
  const tokens: string[] = [];
  let depth = 0;
  let current = "";
  for (const character of input.trim()) {
    if (character === "(") depth += 1;
    if (character === ")") depth -= 1;
    if (depth < 0)
      fail(
        path,
        "invalid-grid-tracks",
        "Mismatched parentheses in grid tracks",
      );
    if (/\s/u.test(character) && depth === 0) {
      if (current !== "") tokens.push(current);
      current = "";
    } else current += character;
  }
  if (depth !== 0)
    fail(path, "invalid-grid-tracks", "Mismatched parentheses in grid tracks");
  if (current !== "") tokens.push(current);
  return tokens;
}

function parseGridPixels(token: string, path: string): number {
  const match = /^(\d+(?:\.\d+)?)(?:px)?$/u.exec(token);
  if (match === null)
    fail(path, "invalid-grid-tracks", `Unsupported grid track ${token}`);
  return Number(match[1]);
}

function parseGridBasicTrack(token: string, path: string): GridTrack {
  if (token === "auto") return { kind: "auto", min: 0 };
  const fraction = /^(\d+(?:\.\d+)?)fr$/u.exec(token);
  if (fraction !== null) {
    const value = Number(fraction[1]);
    if (value <= 0)
      fail(
        path,
        "invalid-grid-tracks",
        "fr track fraction must be greater than 0",
      );
    return { kind: "fr", value, min: 0 };
  }
  return { kind: "px", value: parseGridPixels(token, path) };
}

function gridTracks(
  input: string | undefined,
  path: string,
  defaultFr: boolean,
): GridTrack[] {
  if (input === undefined || input.trim() === "") {
    return defaultFr ? [{ kind: "fr", value: 1, min: 0 }] : [];
  }
  const parseOne = (token: string): GridTrack[] => {
    const repeat = /^repeat\(\s*(\d+)\s*,\s*(.+)\)$/u.exec(token);
    if (repeat !== null) {
      const count = Number(repeat[1]);
      if (!Number.isInteger(count) || count < 1 || count > 128) {
        fail(
          path,
          "invalid-grid-tracks",
          "repeat count must be an integer between 1 and 128",
        );
      }
      const repeated = gridTokens(repeat[2]!, path).flatMap(parseOne);
      if (repeated.length * count > 128) {
        fail(
          path,
          "invalid-grid-tracks",
          "Expanded grid tracks cannot exceed 128 tracks",
        );
      }
      return Array.from({ length: count }, () => repeated).flat();
    }
    const minmax = /^minmax\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)$/u.exec(token);
    if (minmax !== null) {
      const lower = minmax[1]!.trim();
      const minValue = lower === "auto" ? 0 : parseGridPixels(lower, path);
      const maximum = parseGridBasicTrack(minmax[2]!.trim(), path);
      if (maximum.kind === "fr") {
        return [
          {
            ...maximum,
            min: Math.max(minValue, maximum.min),
            contentMinimum: lower === "auto",
          },
        ];
      }
      if (maximum.kind === "auto") return [{ kind: "auto", min: minValue }];
      return [{ kind: "px", value: Math.max(minValue, maximum.value) }];
    }
    return [parseGridBasicTrack(token, path)];
  };
  const tracks = gridTokens(input, path).flatMap(parseOne);
  if (tracks.length === 0 || tracks.length > 128) {
    fail(
      path,
      "invalid-grid-tracks",
      "Grid tracks must contain between 1 and 128 tracks",
    );
  }
  return tracks;
}

function gridAutoRowTracks(
  input: string | undefined,
  path: string,
): readonly GridTrack[] {
  if (input === undefined) return [{ kind: "auto", min: 0 }];
  const tracks = gridTracks(input, path, false);
  if (tracks.length === 0)
    fail(
      path,
      "invalid-grid-tracks",
      "gridAutoRows must contain at least 1 track",
    );
  return tracks;
}

// Only auto rows and fr rows with an auto minimum accept intrinsic content growth. Fixed rows and
// zero-minimum fr rows keep the size assigned by the grid, including when spanning content overlaps
// them and is reported as overflow instead of rewriting track sizes.
function rowTrackGrowsWithContent(
  track: GridTrack,
): track is ContentGrowableGridTrack {
  return (
    track.kind === "auto" ||
    (track.kind === "fr" && track.contentMinimum === true)
  );
}

function distributeFrTracks(
  sizes: number[],
  tracks: readonly GridTrack[],
  availableSize: number,
  gap: number,
): void {
  const trackAt = (index: number): GridTrack =>
    tracks[index] ?? { kind: "auto", min: 0 };
  const flexible = new Set(
    sizes.flatMap((_, index) => (trackAt(index).kind === "fr" ? [index] : [])),
  );
  if (flexible.size === 0) return;

  const gapSize = gap * Math.max(0, sizes.length - 1);
  const fixedSize = sizes.reduce(
    (sum, size, index) => sum + (trackAt(index).kind === "fr" ? 0 : size),
    0,
  );

  // When a lower bound exceeds the initial share, freeze that track first, then allocate remaining space to unfrozen tracks.
  while (flexible.size > 0) {
    const factorTotal = [...flexible].reduce((sum, index) => {
      const track = trackAt(index);
      return sum + (track.kind === "fr" ? track.value : 0);
    }, 0);
    const frozenSize = sizes.reduce(
      (sum, size, index) =>
        sum +
        (flexible.has(index) ? 0 : trackAt(index).kind === "fr" ? size : 0),
      0,
    );
    const freeSize = Math.max(
      0,
      availableSize - fixedSize - gapSize - frozenSize,
    );
    const undersized = [...flexible].filter((index) => {
      const track = trackAt(index);
      return (
        track.kind === "fr" &&
        freeSize * (track.value / factorTotal) < (sizes[index] ?? track.min)
      );
    });
    if (undersized.length > 0) {
      for (const index of undersized) flexible.delete(index);
      continue;
    }
    for (const index of flexible) {
      const track = trackAt(index);
      if (track.kind === "fr")
        sizes[index] = freeSize * (track.value / factorTotal);
    }
    return;
  }
}

function gridPlacement(
  value: number | string | undefined,
  path: string,
): { readonly start?: number; readonly span: number } {
  if (value === undefined) return { span: 1 };
  const positiveIndex = (raw: string): number => {
    const result = Number(raw);
    if (!Number.isInteger(result) || result < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid line numbers must be positive integers",
      );
    }
    return result - 1;
  };
  if (typeof value === "number") {
    if (!Number.isInteger(value) || value < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid line numbers must be positive integers",
      );
    }
    return { start: value - 1, span: 1 };
  }
  const text = value.trim();
  let match = /^span\s+(\d+)$/u.exec(text);
  if (match !== null) {
    const span = Number(match[1]);
    if (span < 1)
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    return { span };
  }
  match = /^(\d+)\s*\/\s*span\s+(\d+)$/u.exec(text);
  if (match !== null) {
    const span = Number(match[2]);
    if (span < 1)
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    return { start: positiveIndex(match[1]!), span };
  }
  match = /^(\d+)\s*\/\s*(\d+)$/u.exec(text);
  if (match !== null) {
    const start = positiveIndex(match[1]!);
    const end = positiveIndex(match[2]!);
    if (end <= start)
      fail(
        path,
        "invalid-grid-placement",
        "Grid end line number must be greater than start line number",
      );
    return { start, span: end - start };
  }
  if (/^\d+$/u.test(text)) return { start: positiveIndex(text), span: 1 };
  return fail(
    path,
    "invalid-grid-placement",
    `Unsupported grid placement ${value}`,
  );
}

function gridAlignment(
  value:
    | LayoutStyle["alignItems"]
    | LayoutStyle["justifyItems"]
    | LayoutStyle["justifySelf"],
): "start" | "end" | "center" | "stretch" {
  if (value === "center") return "center";
  if (value === "end" || value === "flex-end") return "end";
  if (value === "stretch") return "stretch";
  return "start";
}

function validateGridEnumValues(
  value: Record<string, unknown>,
  path: string,
): void {
  const allowedValue = (
    name: string,
    candidate: unknown,
    values: readonly string[],
  ) => {
    if (
      candidate !== undefined &&
      !values.includes(typeof candidate === "string" ? candidate : "")
    ) {
      fail(
        path,
        "unsupported-layout-value",
        `${name} does not support value ${describeStyleValue(candidate)}`,
      );
    }
  };
  allowedValue("justifyItems", value["justifyItems"], [
    "start",
    "end",
    "center",
    "stretch",
    "flex-start",
    "flex-end",
  ]);
  allowedValue("justifySelf", value["justifySelf"], [
    "auto",
    "start",
    "end",
    "center",
    "stretch",
    "flex-start",
    "flex-end",
  ]);
}

function validateGridValueTypes(
  value: Record<string, unknown>,
  path: string,
): void {
  for (const key of [
    "gridTemplateColumns",
    "gridTemplateRows",
    "gridAutoRows",
  ] as const) {
    if (value[key] !== undefined && typeof value[key] !== "string") {
      fail(path, "invalid-layout-style", `${key} must be a grid track string`);
    }
  }
  for (const key of gridItemProperties) {
    const candidate = value[key];
    if (
      candidate !== undefined &&
      key !== "justifySelf" &&
      typeof candidate !== "number" &&
      typeof candidate !== "string"
    ) {
      fail(
        path,
        "invalid-layout-style",
        `${key} must be a line number or span declaration`,
      );
    }
  }
}

function validateGridItemStyle(
  style: LayoutStyle,
  parentStyle: LayoutStyle,
  path: string,
): void {
  if (
    parentStyle.display === "grid" &&
    style.position === "absolute" &&
    (style.alignSelf !== undefined ||
      style.justifySelf !== undefined ||
      style.gridColumn !== undefined ||
      style.gridRow !== undefined)
  ) {
    fail(
      path,
      "unsupported-grid-positioning",
      "Absolutely positioned grid items currently do not support grid placement or self alignment",
    );
  }
  if (
    parentStyle.display === "grid" &&
    style.position !== "absolute" &&
    style.alignSelf === "baseline"
  ) {
    fail(
      path,
      "unsupported-grid-alignment",
      "Grid items do not support baseline alignment",
    );
  }
}

function createGridContainerSolver(
  warnOnce: (path: string, code: string, message: string) => void,
): CustomContainerSolver {
  // Retain allocated box height when a nested grid parent has not yet recalculated Yoga child nodes.
  const gridAreaHeights = new WeakMap<LayoutSolverNode, number>();

  const computePlacements = (
    node: LayoutSolverNode,
    explicitColumns: readonly GridTrack[],
  ): {
    readonly placements: readonly GridPlacement[];
    readonly columns: GridTrack[];
  } => {
    const columns = [...explicitColumns];
    const children = node.children.filter(
      (child) => child.style.position !== "absolute",
    );
    const occupied = new Set<string>();
    const placements: GridPlacement[] = [];
    let cursorRow = 0;
    let cursorColumn = 0;
    const ensureColumns = (count: number) => {
      if (count > 128)
        fail(
          node.path,
          "invalid-grid-placement",
          "Implicit grid columns cannot exceed 128 columns",
        );
      while (columns.length < count) columns.push({ kind: "auto", min: 0 });
    };
    for (const child of children) {
      const column = gridPlacement(child.style.gridColumn, child.path);
      const row = gridPlacement(child.style.gridRow, child.path);
      const columnSpan = column.span;
      if (column.start !== undefined) ensureColumns(column.start + columnSpan);
      else if (columnSpan > columns.length) ensureColumns(columnSpan);
      let currentRow = row.start ?? cursorRow;
      let currentColumn = column.start ?? cursorColumn;
      if (
        row.start === undefined &&
        column.start !== undefined &&
        column.start < cursorColumn
      )
        currentRow += 1;
      const fits = (rowIndex: number, columnIndex: number) => {
        if (row.start === undefined && column.start === undefined) {
          if (columnIndex + columnSpan > columns.length) return false;
        } else {
          ensureColumns(columnIndex + columnSpan);
        }
        for (let rowOffset = 0; rowOffset < row.span; rowOffset += 1) {
          for (
            let columnOffset = 0;
            columnOffset < columnSpan;
            columnOffset += 1
          ) {
            if (
              occupied.has(
                `${rowIndex + rowOffset}:${columnIndex + columnOffset}`,
              )
            )
              return false;
          }
        }
        return true;
      };
      const fixed = row.start !== undefined && column.start !== undefined;
      if (!fixed) {
        while (!fits(currentRow, currentColumn)) {
          if (column.start !== undefined) currentRow += 1;
          else {
            currentColumn += 1;
            if (currentColumn + columnSpan > columns.length) {
              currentColumn = 0;
              currentRow += 1;
            }
          }
        }
      }
      if (currentRow + row.span > 128) {
        fail(
          child.path,
          "invalid-grid-placement",
          "Implicit grid rows cannot exceed 128 rows",
        );
      }
      for (let rowOffset = 0; rowOffset < row.span; rowOffset += 1) {
        for (
          let columnOffset = 0;
          columnOffset < columnSpan;
          columnOffset += 1
        ) {
          occupied.add(
            `${currentRow + rowOffset}:${currentColumn + columnOffset}`,
          );
        }
      }
      placements.push({
        node: child,
        row: currentRow,
        column: currentColumn,
        rowSpan: row.span,
        columnSpan,
      });
      if (row.start === undefined) {
        cursorRow = currentRow;
        cursorColumn = currentColumn + columnSpan;
      }
    }
    return { placements, columns };
  };

  const buildGridModel = (
    node: LayoutSolverNode,
    innerWidth: number | undefined,
    innerHeight: number | undefined,
    port: LayoutSolverPort,
  ): GridModel => {
    const columnTracks = gridTracks(
      node.style.gridTemplateColumns,
      node.path,
      true,
    );
    const rowTracks = gridTracks(node.style.gridTemplateRows, node.path, false);
    const autoRowTracks = gridAutoRowTracks(node.style.gridAutoRows, node.path);
    const { placements, columns } = computePlacements(node, columnTracks);
    const columnGap = node.style.columnGap ?? node.style.gap ?? 0;
    const rowGap = node.style.rowGap ?? node.style.gap ?? 0;
    const rowCount = Math.max(
      rowTracks.length,
      ...placements.map((placement) => placement.row + placement.rowSpan),
      0,
    );
    // Explicit rows keep their template; rows beyond it cycle through gridAutoRows.
    const rowTrackAt = (index: number): GridTrack =>
      rowTracks[index] ??
      autoRowTracks[(index - rowTracks.length) % autoRowTracks.length]!;
    const minContentWidth = (item: LayoutSolverNode): number =>
      port.measureMinContentContribution(item, "width");
    const columnWidths = columns.map((track) =>
      track.kind === "px" ? track.value : track.min,
    );
    for (const placement of placements) {
      if (placement.columnSpan !== 1) continue;
      const track = columns[placement.column];
      if (
        track === undefined ||
        track.kind === "px" ||
        (track.kind === "fr" && track.contentMinimum !== true)
      )
        continue;
      columnWidths[placement.column] = Math.max(
        columnWidths[placement.column] ?? 0,
        Math.max(track.min, minContentWidth(placement.node)),
      );
    }
    const columnSpanWidth = (start: number, span: number) =>
      columnWidths
        .slice(start, start + span)
        .reduce((sum, width) => sum + width, 0) +
      columnGap * Math.max(0, span - 1);
    for (const placement of placements) {
      if (placement.columnSpan === 1) continue;
      const required = minContentWidth(placement.node);
      const missing =
        required - columnSpanWidth(placement.column, placement.columnSpan);
      if (missing <= 0) continue;
      const growable = Array.from(
        { length: placement.columnSpan },
        (_, index) => placement.column + index,
      ).filter(
        (index) =>
          columns[index]?.kind === "auto" ||
          (columns[index]?.kind === "fr" &&
            columns[index]?.contentMinimum === true),
      );
      if (growable.length === 0) continue;
      for (const index of growable) {
        columnWidths[index] =
          (columnWidths[index] ?? 0) + missing / growable.length;
      }
    }
    if (innerWidth !== undefined) {
      distributeFrTracks(columnWidths, columns, innerWidth, columnGap);
    }
    const spanWidth = (start: number, span: number) =>
      columnWidths
        .slice(start, start + span)
        .reduce((sum, width) => sum + width, 0) +
      columnGap * Math.max(0, span - 1);
    const rowHeights = Array.from({ length: rowCount }, (_, index) => {
      const track = rowTrackAt(index);
      return track.kind === "px" ? track.value : track.min;
    });
    for (const placement of placements) {
      if (placement.rowSpan !== 1) continue;
      const track = rowTrackAt(placement.row);
      if (!rowTrackGrowsWithContent(track)) continue;
      const margin = expandedSpacing(placement.node.style, "margin");
      const size = port.measureNaturalSize(
        placement.node,
        spanWidth(placement.column, placement.columnSpan),
      );
      rowHeights[placement.row] = Math.max(
        rowHeights[placement.row] ?? 0,
        track.min,
        size.height + margin.top + margin.bottom,
      );
    }
    const spanHeight = (start: number, span: number) =>
      rowHeights
        .slice(start, start + span)
        .reduce((sum, height) => sum + height, 0) +
      rowGap * Math.max(0, span - 1);
    for (const placement of placements) {
      if (placement.rowSpan === 1) continue;
      const natural = port.measureNaturalSize(
        placement.node,
        spanWidth(placement.column, placement.columnSpan),
      );
      const margin = expandedSpacing(placement.node.style, "margin");
      const missing =
        natural.height +
        margin.top +
        margin.bottom -
        spanHeight(placement.row, placement.rowSpan);
      if (missing <= 0) continue;
      const growable = Array.from(
        { length: placement.rowSpan },
        (_, index) => placement.row + index,
      ).filter((index) => rowTrackGrowsWithContent(rowTrackAt(index)));
      if (growable.length === 0) continue;
      for (const index of growable) {
        rowHeights[index] =
          (rowHeights[index] ?? 0) + missing / growable.length;
      }
    }
    if (innerHeight !== undefined) {
      distributeFrTracks(
        rowHeights,
        Array.from({ length: rowCount }, (_, index) => rowTrackAt(index)),
        innerHeight,
        rowGap,
      );
    }
    const spanHeightFinal = (start: number, span: number) =>
      rowHeights
        .slice(start, start + span)
        .reduce((sum, height) => sum + height, 0) +
      rowGap * Math.max(0, span - 1);
    return {
      placements,
      widths: columnWidths,
      heights: rowHeights,
      columnGap,
      rowGap,
      spanWidth,
      spanHeight: spanHeightFinal,
      width: spanWidth(0, columnWidths.length),
      height: rowCount === 0 ? 0 : spanHeightFinal(0, rowCount),
    };
  };

  return {
    handlesDisplay(display) {
      return display === "grid";
    },

    measureMinContent(node, axis, port) {
      const model = buildGridModel(node, undefined, undefined, port);
      const padding = expandedSpacing(node.style, "padding");
      return (
        (axis === "width" ? model.width : model.height) +
        (axis === "width"
          ? padding.left + padding.right
          : padding.top + padding.bottom)
      );
    },

    placeContainer(node, port) {
      if (node.style.display !== "grid") return;
      const padding = expandedSpacing(node.style, "padding");
      const computed = port.getComputedSize(node);
      const outerWidth = computed.width;
      const outerHeight = computed.height || gridAreaHeights.get(node) || 0;
      const hasWidth =
        node.isRoot ||
        (node.style.width !== undefined && node.style.width !== "auto");
      const hasResolvedHeight =
        node.isRoot ||
        (node.style.height !== undefined && node.style.height !== "auto") ||
        outerHeight > 0;
      const innerWidth =
        outerWidth > 0
          ? Math.max(0, outerWidth - padding.left - padding.right)
          : undefined;
      const innerHeight = hasResolvedHeight
        ? Math.max(0, outerHeight - padding.top - padding.bottom)
        : undefined;
      const model = buildGridModel(node, innerWidth, innerHeight, port);
      if (innerWidth !== undefined && model.width > innerWidth + 0.01) {
        warnOnce(
          node.path,
          "grid-overflow",
          `Grid columns minimum size total ${model.width}px exceeds available width ${innerWidth}px; build retains track dimensions`,
        );
      }
      if (innerHeight !== undefined && model.height > innerHeight + 0.01) {
        warnOnce(
          node.path,
          "grid-overflow",
          `Grid rows minimum size total ${model.height}px exceeds available height ${innerHeight}px; build retains track dimensions`,
        );
      }
      port.setContainerSize(node, {
        ...(!hasWidth && outerWidth <= 0
          ? { width: model.width + padding.left + padding.right }
          : {}),
        ...(!hasResolvedHeight
          ? { height: model.height + padding.top + padding.bottom }
          : {}),
      });
      for (const placement of model.placements) {
        const child = placement.node;
        const childStyle = child.style;
        const cellWidth = model.spanWidth(
          placement.column,
          placement.columnSpan,
        );
        const cellHeight = model.spanHeight(placement.row, placement.rowSpan);
        const cellX =
          padding.left +
          model.spanWidth(0, placement.column) +
          (placement.column > 0 ? model.columnGap : 0);
        const cellY =
          padding.top +
          model.spanHeight(0, placement.row) +
          (placement.row > 0 ? model.rowGap : 0);
        const margin = expandedSpacing(childStyle, "margin");
        const availableWidth = Math.max(
          0,
          cellWidth - margin.left - margin.right,
        );
        const availableHeight = Math.max(
          0,
          cellHeight - margin.top - margin.bottom,
        );
        const justify = gridAlignment(
          childStyle.justifySelf === undefined ||
            childStyle.justifySelf === "auto"
            ? (node.style.justifyItems ?? "stretch")
            : childStyle.justifySelf,
        );
        const alignValue =
          childStyle.alignSelf === undefined || childStyle.alignSelf === "auto"
            ? (node.style.alignItems ?? "stretch")
            : childStyle.alignSelf;
        if (alignValue === "baseline") {
          fail(
            child.path,
            "unsupported-grid-alignment",
            "Text baseline alignment in grid does not yet have a reliable row baseline",
          );
        }
        const align = gridAlignment(alignValue);
        const declaredWidth = childStyle.width ?? child.contractWidth;
        const declaredHeight = childStyle.height ?? child.contractHeight;
        const widthPercent = resolvePercent(declaredWidth, cellWidth);
        const heightPercent = resolvePercent(declaredHeight, cellHeight);
        const widthSpecified =
          declaredWidth !== undefined && declaredWidth !== "auto";
        let heightSpecified =
          declaredHeight !== undefined && declaredHeight !== "auto";
        const minWidth = resolvePercent(childStyle.minWidth, cellWidth);
        const maxWidth = resolvePercent(childStyle.maxWidth, cellWidth);
        const minHeight = resolvePercent(childStyle.minHeight, cellHeight);
        const maxHeight = resolvePercent(childStyle.maxHeight, cellHeight);
        let initialHeightConstraint: number | "auto" | undefined;
        if (heightPercent !== undefined && availableHeight > 0) {
          initialHeightConstraint = heightPercent;
        } else if (heightPercent !== undefined) {
          initialHeightConstraint = "auto";
          heightSpecified = false;
        }
        port.setChildConstraints(child, {
          ...(widthPercent === undefined ? {} : { width: widthPercent }),
          ...(initialHeightConstraint === undefined
            ? {}
            : { height: initialHeightConstraint }),
          ...(minWidth === undefined ? {} : { minWidth }),
          ...(maxWidth === undefined ? {} : { maxWidth }),
          ...(minHeight === undefined ? {} : { minHeight }),
          ...(maxHeight === undefined ? {} : { maxHeight }),
        });
        const intrinsic = port.measureNaturalSize(
          child,
          typeof declaredWidth === "string" && declaredWidth.endsWith("%")
            ? cellWidth
            : undefined,
        );
        let width = widthSpecified
          ? intrinsic.width
          : justify === "stretch"
            ? availableWidth
            : intrinsic.width;
        if (
          childStyle.aspectRatio !== undefined &&
          heightSpecified &&
          !widthSpecified
        ) {
          width = intrinsic.height * childStyle.aspectRatio;
        }
        if (!widthSpecified || widthPercent !== undefined) {
          port.setChildConstraints(child, { width: Math.max(0, width) });
        }
        const naturalAtWidth = port.measureNaturalSize(child, width);
        let height = heightSpecified
          ? naturalAtWidth.height
          : align === "stretch" && childStyle.aspectRatio === undefined
            ? availableHeight
            : naturalAtWidth.height;
        if (!heightSpecified && childStyle.aspectRatio === undefined) {
          port.setChildConstraints(child, { height: Math.max(0, height) });
        }
        // Recalculate after assigning the stretched size; the previous natural-size pass may
        // still hold the content height and would report a false track overflow.
        const finalSize = port.measureNaturalSize(child, width);
        const usedWidth = finalSize.width || width;
        const usedHeight = finalSize.height || height;
        if (child.style.display === "grid" && !heightSpecified) {
          gridAreaHeights.set(child, height);
        }
        if (
          usedWidth > availableWidth + 0.01 ||
          usedHeight > availableHeight + 0.01
        ) {
          warnOnce(
            child.path,
            "grid-item-overflow",
            "Grid item overflows its track area; build retains declared dimensions",
          );
        }
        const xAlign =
          justify === "center"
            ? Math.max(0, (availableWidth - usedWidth) / 2)
            : justify === "end"
              ? Math.max(0, availableWidth - usedWidth)
              : 0;
        const yAlign =
          align === "center"
            ? Math.max(0, (availableHeight - usedHeight) / 2)
            : align === "end"
              ? Math.max(0, availableHeight - usedHeight)
              : 0;
        const relative = (
          start: LayoutLength | undefined,
          end: LayoutLength | undefined,
          basis: number,
        ) => {
          const resolve = (value: LayoutLength | undefined) =>
            typeof value === "number"
              ? value
              : typeof value === "string" && value.endsWith("%")
                ? (Number(value.slice(0, -1)) * basis) / 100
                : undefined;
          const startValue = resolve(start);
          return startValue ?? -(resolve(end) ?? 0);
        };
        const dx =
          childStyle.position === "relative"
            ? relative(
                childStyle.left ?? childStyle.inset,
                childStyle.right ?? childStyle.inset,
                cellWidth,
              )
            : 0;
        const dy =
          childStyle.position === "relative"
            ? relative(
                childStyle.top ?? childStyle.inset,
                childStyle.bottom ?? childStyle.inset,
                cellHeight,
              )
            : 0;
        port.placeChild(child, {
          x: cellX + xAlign + dx,
          y: cellY + yAlign + dy,
        });
      }
    },
  };
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
  );
  if (resolvedDisplay === "grid" && value["alignItems"] === "baseline") {
    fail(
      path,
      "unsupported-grid-alignment",
      "Grid containers do not support baseline alignment",
    );
  }
  validateBlockFlexContainerProperties(value, resolvedDisplay, path);
  for (const key of gridContainerProperties) {
    if (value[key] !== undefined && resolvedDisplay !== "grid") {
      fail(
        path,
        "invalid-layout-context",
        `${key} is only applicable to display: "grid" containers`,
      );
    }
  }
  validateBlockFlexEnumValues(value, path);
  validateGridEnumValues(value, path);
  validateBlockFlexDimensionsAndOffsets(value, allowsPadding, path);
  validateGridValueTypes(value, path);
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
  readonly #options: BuildOptions;
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
    options: BuildOptions = {},
  ) {
    this.#modules = modules;
    this.#dimensions = dimensions;
    this.#slideNumber = slideNumber;
    this.#options = options;
    this.#solver = createBlockFlexLayoutSolver({
      dimensions,
      customContainers: [
        createGridContainerSolver((path, code, message) => {
          this.#warnOnce(path, code, message);
        }),
      ],
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
        bound = module.bind(child, {
          path: childPath,
          theme: themeContext,
          options: this.#options,
          slideNumber: this.#slideNumber,
        });
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

  // Transparent Groups are not in the layout tree, so theme context must be passed separately along the declaration tree.
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
      if (child.type === "rp:group") {
        const module = this.#modules.get(child.type);
        if (!objectModule(module)) {
          fail(
            childPath,
            "missing-capability",
            "rp:group capability is required",
          );
        }
        const bound = this.#bindCapability(
          child,
          childPath,
          themeContext,
          module,
        );
        const contract = bound.contract;
        const childStyle = contract.layoutStyle;
        if (childStyle.display === "none") {
          this.#markHidden(child);
          return;
        }
        validateGridItemStyle(childStyle, parent.style, childPath);
        if (!includeObjects) {
          // Group text defaults apply only to descendants, not indicating it should become a layout item.
          const optsIntoLayout = Object.keys(childStyle).some(
            (key) => key !== "zIndex" && !textStyleKeys.has(key),
          );
          if (!optsIntoLayout) {
            const groupTheme = deriveTheme(
              themeContext,
              child.props["theme"] as ThemeScope | undefined,
              child.props["style"],
              childPath,
            );
            this.#themes.set(child, groupTheme);
            this.#appendChildren(
              child.children,
              childPath,
              parent,
              false,
              groupTheme,
            );
            return;
          }
        }
        const built = this.#buildGroup(
          child,
          childPath,
          parent,
          themeContext,
          contract,
        );
        this.#solver.appendChild(parent, built);
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
        validateGridItemStyle(contract.layoutStyle, parent.style, childPath);
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

  #buildGroup(
    input: CapabilityElement,
    path: string,
    parent: LayoutSolverNode,
    themeContext: ThemeContext,
    contract: CapabilityLayoutContract,
  ): LayoutSolverNode {
    const rawStyle = input.props["style"];
    const style = contract.layoutStyle;
    const theme = deriveTheme(
      themeContext,
      input.props["theme"] as ThemeScope | undefined,
      rawStyle,
      path,
    );
    this.#themes.set(input, theme);
    const node = this.#solver.createNode(
      {
        path,
        style,
        isContainer: true,
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
      },
      parent,
    );
    this.#nodeMeta.set(node, {
      element: input,
      checksCanvasOverflow: false,
    });
    this.#appendChildren(input.children, path, node, true, theme);
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
