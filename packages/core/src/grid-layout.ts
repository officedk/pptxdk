import { BuildError, toBuildDiagnostic } from "#/diagnostics.ts";
import {
  expandedSpacing,
  resolvePercent,
  type CustomContainerMode,
  type CustomContainerSolver,
  type LayoutSolverNode,
  type LayoutSolverPort,
} from "#/layout-protocol.ts";
import type { LayoutLength, LayoutStyle } from "#/types.ts";

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

const gridContainerProperties = [
  "gridTemplateColumns",
  "gridTemplateRows",
  "gridAutoRows",
  "justifyItems",
] as const;

const gridItemProperties = ["gridColumn", "gridRow", "justifySelf"] as const;

function fail(
  path: string,
  code: string,
  message: string,
  details?: Readonly<Record<string, unknown>>,
): never {
  throw new BuildError([toBuildDiagnostic(path, code, message, details)]);
}

function gridTokens(input: string, path: string): string[] {
  const tokens: string[] = [];
  let depth = 0;
  let current = "";
  for (const character of input.trim()) {
    if (character === "(") depth += 1;
    if (character === ")") depth -= 1;
    if (depth < 0) {
      fail(
        path,
        "invalid-grid-tracks",
        "Mismatched parentheses in grid tracks",
      );
    }
    if (/\s/u.test(character) && depth === 0) {
      if (current !== "") tokens.push(current);
      current = "";
    } else {
      current += character;
    }
  }
  if (depth !== 0) {
    fail(path, "invalid-grid-tracks", "Mismatched parentheses in grid tracks");
  }
  if (current !== "") tokens.push(current);
  return tokens;
}

function parseGridPixels(token: string, path: string): number {
  const match = /^(\d+(?:\.\d+)?)(?:px)?$/u.exec(token);
  if (match === null) {
    fail(path, "invalid-grid-tracks", `Unsupported grid track ${token}`);
  }
  return Number(match[1]);
}

function parseGridBasicTrack(token: string, path: string): GridTrack {
  if (token === "auto") return { kind: "auto", min: 0 };
  const fraction = /^(\d+(?:\.\d+)?)fr$/u.exec(token);
  if (fraction !== null) {
    const value = Number(fraction[1]);
    if (value <= 0) {
      fail(
        path,
        "invalid-grid-tracks",
        "fr track fraction must be greater than 0",
      );
    }
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
  if (tracks.length === 0) {
    fail(
      path,
      "invalid-grid-tracks",
      "gridAutoRows must contain at least 1 track",
    );
  }
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
      if (track.kind === "fr") {
        sizes[index] = freeSize * (track.value / factorTotal);
      }
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
    if (span < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    }
    return { span };
  }
  match = /^(\d+)\s*\/\s*span\s+(\d+)$/u.exec(text);
  if (match !== null) {
    const span = Number(match[2]);
    if (span < 1) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid span must be a positive integer",
      );
    }
    return { start: positiveIndex(match[1]!), span };
  }
  match = /^(\d+)\s*\/\s*(\d+)$/u.exec(text);
  if (match !== null) {
    const start = positiveIndex(match[1]!);
    const end = positiveIndex(match[2]!);
    if (end <= start) {
      fail(
        path,
        "invalid-grid-placement",
        "Grid end line number must be greater than start line number",
      );
    }
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

function validateGridContainerStyle(
  value: Readonly<Record<string, unknown>>,
  resolvedDisplay: NonNullable<LayoutStyle["display"]>,
  path: string,
): void {
  if (resolvedDisplay === "grid" && value["alignItems"] === "baseline") {
    fail(
      path,
      "unsupported-grid-alignment",
      "Grid containers do not support baseline alignment",
    );
  }
  for (const key of gridContainerProperties) {
    if (value[key] !== undefined && resolvedDisplay !== "grid") {
      fail(
        path,
        "invalid-layout-context",
        `${key} is only applicable to display: "grid" containers`,
      );
    }
  }
}

function validateGridStyleValues(): void {
  // Pure grid property shapes are validated upfront by layoutStyleSchema in layout.ts.
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

function computeGridPlacements(
  node: LayoutSolverNode,
  explicitColumns: readonly GridTrack[],
): {
  readonly placements: readonly GridPlacement[];
  readonly columns: GridTrack[];
} {
  const columns = [...explicitColumns];
  const children = node.children.filter(
    (child) => child.style.position !== "absolute",
  );
  const occupied = new Set<string>();
  const placements: GridPlacement[] = [];
  let cursorRow = 0;
  let cursorColumn = 0;
  const ensureColumns = (count: number) => {
    if (count > 128) {
      fail(
        node.path,
        "invalid-grid-placement",
        "Implicit grid columns cannot exceed 128 columns",
      );
    }
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
    ) {
      currentRow += 1;
    }
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
          ) {
            return false;
          }
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
      for (let columnOffset = 0; columnOffset < columnSpan; columnOffset += 1) {
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
}

function buildGridModel(
  node: LayoutSolverNode,
  innerWidth: number | undefined,
  innerHeight: number | undefined,
  port: LayoutSolverPort,
): GridModel {
  const columnTracks = gridTracks(
    node.style.gridTemplateColumns,
    node.path,
    true,
  );
  const rowTracks = gridTracks(node.style.gridTemplateRows, node.path, false);
  const autoRowTracks = gridAutoRowTracks(node.style.gridAutoRows, node.path);
  const { placements, columns } = computeGridPlacements(node, columnTracks);
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
    ) {
      continue;
    }
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
      rowHeights[index] = (rowHeights[index] ?? 0) + missing / growable.length;
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
}

/**
 * Creates the Grid container solver collaborating strictly through `LayoutSolverPort`.
 */
export function createGridContainerSolver(): CustomContainerSolver {
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
      const outerHeight = computed.height || port.getAllocatedHeight(node) || 0;
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
      port.setContainerSize(node, {
        ...(!hasWidth && outerWidth <= 0
          ? { width: model.width + padding.left + padding.right }
          : {}),
        ...(!hasResolvedHeight
          ? { height: model.height + padding.top + padding.bottom }
          : {}),
      });
      port.setGridTracks?.(node, {
        columnWidths: [...model.widths],
        rowHeights: [...model.heights],
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
        port.setGridCellArea?.(child, {
          x: cellX,
          y: cellY,
          width: cellWidth,
          height: cellHeight,
        });
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
        const declaredWidth = childStyle.width;
        const declaredHeight = childStyle.height;
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
        if (
          !widthSpecified ||
          widthPercent !== undefined ||
          child.contentMinimumsEnforced
        ) {
          port.setChildConstraints(child, { width: Math.max(0, width) });
        }
        const naturalAtWidth = port.measureNaturalSize(child, width);
        let height = heightSpecified
          ? naturalAtWidth.height
          : align === "stretch" && childStyle.aspectRatio === undefined
            ? availableHeight
            : naturalAtWidth.height;
        if (
          (!heightSpecified && childStyle.aspectRatio === undefined) ||
          child.contentMinimumsEnforced
        ) {
          port.setChildConstraints(child, { height: Math.max(0, height) });
        }
        // Recalculate after assigning the stretched size; the previous natural-size pass may
        // still hold the content height and would report a false track overflow.
        const finalSize = port.measureNaturalSize(child, width);
        const usedWidth = finalSize.width || width;
        const usedHeight = finalSize.height || height;
        if (child.contentMinimumsEnforced) {
          port.setChildConstraints(child, {
            width: Math.max(0, usedWidth),
            height: Math.max(0, usedHeight),
          });
        }
        if (child.isContainer && !heightSpecified) {
          port.setAllocatedHeight(child, height);
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
            typeof value === "number" ? value : resolvePercent(value, basis);
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

/**
 * Ownership descriptor for Grid layout mode, registered with the core layout coordinator.
 */
export const gridLayoutMode: CustomContainerMode = {
  display: "grid",
  containerStyleKeys: gridContainerProperties,
  itemStyleKeys: gridItemProperties,
  sharedContainerProperties: ["alignItems", "gap", "rowGap", "columnGap"],
  validateContainerStyle: validateGridContainerStyle,
  validateStyleValues: validateGridStyleValues,
  validateItemStyle: validateGridItemStyle,
  createSolver: createGridContainerSolver,
};
