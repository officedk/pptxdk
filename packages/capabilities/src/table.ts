import {
  assertValidLang,
  CapabilityError,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleKeys,
  xmlEscape,
  type CapabilityContext,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityOutput,
  type CapabilityParseContext,
  type CapabilityParseResult,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type NativeTableObservation,
  type ObjectCapabilityModule,
  type ResolvedColor,
  type TextStyle,
  type ThemeContext,
} from "@react-pptx/core/capability";
import { CommonContainer, RichTextContainer } from "@react-pptx/core/document";

import {
  fontPresentationContributions,
  type FontEmbeddingMode,
  type FontEmbeddingUse,
} from "./font-embedding.ts";
import { nativeLinkXml } from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";
import {
  emptyTableCellXml,
  tableCellPropertiesXml,
  tableFrameXml,
  type TableCellBorder,
  type TableCellPadding,
} from "./table-xml.ts";
import {
  parseTextParagraphs,
  textExtent,
  writeTextBodyFromParagraphs,
  type FontReference,
  type FontSize,
  type Paragraph,
  type ResolvedTextStyle,
} from "./text-service.ts";

/** Column width in px; Nfr allocates from remaining width, auto uses shared font measurement. */
export type TableColumn = number | `${number}fr` | "auto";

export type TableVerticalAlign = "top" | "middle" | "bottom";
export type TableSpacing =
  | number
  | readonly [vertical: number, horizontal: number]
  | readonly [top: number, right: number, bottom: number, left: number];

export type TableTextStyle = {
  readonly fontFamily?: FontReference;
  readonly fontSize?: FontSize;
  readonly fontWeight?: number | "normal" | "bold";
  readonly fontStyle?: "normal" | "italic";
  readonly color?: ColorValue;
  readonly lineHeight?: FontSize;
  readonly verticalAlign?: TableVerticalAlign;
  readonly letterSpacing?: TextStyle["letterSpacing"];
  readonly textTransform?: TextStyle["textTransform"];
  readonly textDecoration?: TextStyle["textDecoration"];
  readonly textShadow?: TextStyle["textShadow"];
  readonly lang?: string;
};

export type TableStyle = TableTextStyle & LayoutStyle;

export type TableRowStyle = TableTextStyle & {
  /** When unspecified, row height is determined from text measurement of unmerged cells in this row. */
  readonly height?: number;
  readonly backgroundColor?: ColorValue;
};

export type TableCellStyle = TableTextStyle & {
  readonly backgroundColor?: ColorValue;
  readonly borderColor?: ColorValue;
  readonly borderWidth?: number;
  readonly borderStyle?: "solid" | "dashed" | "dotted";
  readonly padding?: TableSpacing;
  readonly paddingTop?: number;
  readonly paddingRight?: number;
  readonly paddingBottom?: number;
  readonly paddingLeft?: number;
};

export type TableDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly testId?: string;
    readonly part?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    /** Provides allocatable width for fr columns; when omitted, all columns are sized to natural content width. */
    readonly width?: number;
    readonly columns: readonly TableColumn[];
    readonly style?: TableStyle;
  };

type TablePartDeclarationProps = {
  readonly id?: string;
  readonly testId?: string;
  readonly part?: string;
};

export type RowDeclarationProps = TablePartDeclarationProps & {
  readonly style?: TableRowStyle;
};

export type CellDeclarationProps = TablePartDeclarationProps & {
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly style?: TableCellStyle;
};

type Track =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "fr"; readonly value: number }
  | { readonly kind: "auto" };

type Edges = TableCellPadding;

type ParsedCell = {
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly colSpan: number;
  readonly path: string;
  readonly authorId?: string;
  readonly testId?: string;
  readonly part?: string;
  readonly paragraphs: readonly Paragraph[];
  readonly textStyle: ResolvedTextStyle;
  readonly verticalAlign: TableVerticalAlign;
  readonly padding: Edges;
  readonly fill?: ResolvedColor;
  readonly border?: TableCellBorder;
};

type ParsedRow = {
  readonly path: string;
  readonly authorId?: string;
  readonly testId?: string;
  readonly part?: string;
  readonly height?: number;
  readonly textStyle: Record<string, unknown>;
  readonly verticalAlign?: TableVerticalAlign;
  readonly fill?: ResolvedColor;
};
const rowStyleKeys = ["height", "backgroundColor"] as const;
const cellStyleKeys = [
  "backgroundColor",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
] as const;
const tableParameters = [
  "id",
  "testId",
  "part",
  "name",
  "morphKey",
  "link",
  "x",
  "y",
  "width",
  "columns",
  "style",
  "children",
  "enter",
  "emphasis",
  "motion",
] as const;
const tablePartParameters = ["id", "testId", "part", "style"] as const;
const cellParameters = [...tablePartParameters, "colSpan", "rowSpan"] as const;
const verticalAlignments = new Set(["top", "middle", "bottom"]);
const maximumExtentEmu = 2_147_483_647;

function fail(code: string, message: string, path?: string): never {
  throw new CapabilityError(code, message, path);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function record(
  value: unknown,
  label: string,
  path?: string,
): Record<string, unknown> {
  if (!isRecord(value))
    fail("invalid-parameter", `${label} must be an object`, path);
  return value;
}

function onlyKeys(
  value: Record<string, unknown>,
  allowed: readonly string[],
  label: string,
  path?: string,
): void {
  const unknown = Object.keys(value).find((key) => !allowed.includes(key));
  if (unknown !== undefined) {
    fail(
      "unknown-parameter",
      `${label} contains unsupported parameter ${unknown}`,
      path,
    );
  }
}

function optionalName(
  value: unknown,
  label: "id" | "testId" | "name" | "part",
  path?: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    fail("invalid-parameter", `${label} must be a non-empty string`, path);
  }
  return value;
}

function pathLabel(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() !== "" ? value : fallback;
}

function coordinate(value: unknown, label: "x" | "y"): number {
  if (value === undefined && (label === "x" || label === "y")) {
    return 0;
  }
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < -1_000_000 ||
    value > 1_000_000
  ) {
    fail(
      "invalid-table-geometry",
      `${label} must be a finite pixel number between -1000000 and 1000000`,
    );
  }
  return value;
}

function positiveSize(value: unknown, label: string, path?: string): number {
  if (typeof value !== "number") {
    fail(
      "invalid-table-geometry",
      `${label} must be a finite positive pixel number`,
      path,
    );
  }
  try {
    textExtent(value, label);
  } catch (error) {
    if (error instanceof CapabilityError) {
      fail(error.code, error.message, path);
    }
    throw error;
  }
  return value;
}

function color(
  value: unknown,
  label: string,
  theme: ThemeContext,
  path?: string,
): ResolvedColor | undefined {
  if (value === undefined) return undefined;
  try {
    return theme.resolveColor(value as ColorValue);
  } catch (error) {
    if (error instanceof CapabilityError) {
      fail(error.code, `${label}: ${error.message}`, path);
    }
    throw error;
  }
}

function verticalAlign(
  value: unknown,
  path?: string,
): TableVerticalAlign | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || !verticalAlignments.has(value)) {
    fail(
      "invalid-vertical-align",
      "verticalAlign must be top, middle, or bottom",
      path,
    );
  }
  return value as TableVerticalAlign;
}

function textStyle(
  styleValue: unknown,
  label: string,
  path?: string,
  additional: readonly string[] = [],
): {
  readonly textStyle: Record<string, unknown>;
  readonly style: Record<string, unknown>;
} {
  const style = record(styleValue ?? {}, `${label}.style`, path);
  onlyKeys(style, [...textStyleKeys, ...additional], `${label}.style`, path);
  if (style["lang"] !== undefined) {
    try {
      assertValidLang(style["lang"]);
    } catch (error) {
      if (error instanceof CapabilityError) {
        fail(error.code, error.message, path);
      }
      throw error;
    }
  }
  const textStyle: Record<string, unknown> = {};
  for (const key of textStyleKeys) {
    const value = style[key];
    if (value !== undefined) textStyle[key] = value;
  }
  verticalAlign(style["verticalAlign"], path);
  return { textStyle, style };
}

function mergeTextStyle(
  parent: Record<string, unknown>,
  child: Record<string, unknown>,
): Record<string, unknown> {
  return { ...parent, ...child };
}

function positiveTrack(value: number, label: string): number {
  if (!Number.isFinite(value) || value <= 0) {
    fail("invalid-table-column", `${label} must be greater than 0`);
  }
  return value;
}

function parseTracks(value: unknown): Track[] {
  if (!Array.isArray(value) || value.length === 0) {
    fail("invalid-table-column", "Table.columns must be a non-empty array");
  }
  return value.map((column, index) => {
    if (typeof column === "number") {
      return {
        kind: "px",
        value: positiveTrack(column, `column ${index + 1} width`),
      };
    }
    if (column === "auto") return { kind: "auto" };
    if (typeof column === "string") {
      const fraction = /^(\d+(?:\.\d+)?)fr$/u.exec(column);
      if (fraction !== null) {
        return {
          kind: "fr",
          value: positiveTrack(
            Number(fraction[1]),
            `column ${index + 1} fraction`,
          ),
        };
      }
    }
    fail(
      "invalid-table-column",
      `Invalid width for column ${index + 1}; use positive px number, positive "Nfr", or "auto"`,
    );
  });
}

function span(
  value: unknown,
  label: "colSpan" | "rowSpan",
  path: string,
): number {
  if (value === undefined) return 1;
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
    fail("invalid-table-merge", `${label} must be a positive integer`, path);
  }
  return value;
}

function trackStyle(
  input: CapabilityInput,
  type: "table" | "row" | "cell",
  path?: string,
): {
  readonly textStyle: Record<string, unknown>;
  readonly style: Record<string, unknown>;
} {
  const additional =
    type === "table"
      ? layoutStyleKeys
      : type === "row"
        ? rowStyleKeys
        : cellStyleKeys;
  return textStyle(
    input.props["style"],
    type === "table" ? "Table" : type === "row" ? "Row" : "Cell",
    path,
    additional,
  );
}

function spacing(value: unknown, path: string): Edges {
  if (value === undefined) return { top: 0, right: 0, bottom: 0, left: 0 };
  const values: readonly unknown[] =
    typeof value === "number" ? [value] : Array.isArray(value) ? value : [];
  if (values.length !== 1 && values.length !== 2 && values.length !== 4) {
    fail(
      "invalid-table-padding",
      "padding must be a number, a 2-tuple, or a 4-tuple",
      path,
    );
  }
  const number = (entry: unknown): number => {
    if (typeof entry !== "number" || !Number.isFinite(entry) || entry < 0) {
      fail(
        "invalid-table-padding",
        "Cell padding must be a finite non-negative pixel number",
        path,
      );
    }
    return entry;
  };
  const [a, b, c, d] = values.map(number);
  if (values.length === 1) return { top: a!, right: a!, bottom: a!, left: a! };
  if (values.length === 2) return { top: a!, right: b!, bottom: a!, left: b! };
  return { top: a!, right: b!, bottom: c!, left: d! };
}

function cellPadding(style: Record<string, unknown>, path: string): Edges {
  const base = spacing(style["padding"], path);
  const edge = (
    name: "paddingTop" | "paddingRight" | "paddingBottom" | "paddingLeft",
    fallback: number,
  ) => {
    const value = style[name];
    if (value === undefined) return fallback;
    if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
      fail(
        "invalid-table-padding",
        `${name} must be a finite non-negative pixel number`,
        path,
      );
    }
    return value;
  };
  return {
    top: edge("paddingTop", base.top),
    right: edge("paddingRight", base.right),
    bottom: edge("paddingBottom", base.bottom),
    left: edge("paddingLeft", base.left),
  };
}

function mergeBorder(
  style: Record<string, unknown>,
  theme: ThemeContext,
  path: string,
): ParsedCell["border"] {
  const borderColor = color(style["borderColor"], "borderColor", theme, path);
  const widthValue = style["borderWidth"];
  const lineStyle = style["borderStyle"];
  if (
    widthValue !== undefined &&
    (typeof widthValue !== "number" ||
      !Number.isFinite(widthValue) ||
      widthValue < 0)
  ) {
    fail(
      "invalid-table-border",
      "borderWidth must be a finite non-negative pixel number",
      path,
    );
  }
  if (
    lineStyle !== undefined &&
    lineStyle !== "solid" &&
    lineStyle !== "dashed" &&
    lineStyle !== "dotted"
  ) {
    fail(
      "invalid-table-border",
      "borderStyle must be solid, dashed, or dotted",
      path,
    );
  }
  const width = widthValue ?? (borderColor === undefined ? 0 : 1);
  if (width === 0) return undefined;
  pxExtent(width, "borderWidth", path);
  return {
    color: borderColor ?? theme.resolveColor("dk1"),
    width,
    dash:
      lineStyle === "dashed"
        ? "dash"
        : lineStyle === "dotted"
          ? "dot"
          : "solid",
  };
}

function pxExtent(value: number, label: string, path?: string): number {
  const emu = pxToEmu(value);
  if (!Number.isSafeInteger(emu) || emu < 1 || emu > maximumExtentEmu) {
    fail(
      "invalid-table-geometry",
      `${label} exceeds writable table dimension range after conversion`,
      path,
    );
  }
  return emu;
}

function textBody(
  cell: ParsedCell,
  theme: ThemeContext,
  width: number | "max-content",
  height?: number,
  contentSized = false,
  fields?: {
    readonly slideNumber: number;
    readonly createId?: () => string;
  },
  embedFonts?: FontEmbeddingMode,
): ReturnType<typeof writeTextBodyFromParagraphs> {
  try {
    return writeTextBodyFromParagraphs(
      cell.paragraphs,
      cell.textStyle,
      theme,
      cell.verticalAlign,
      {
        width,
        ...(height === undefined ? {} : { height }),
        ...(contentSized ? { contentSized: true } : {}),
      },
      "table-cell",
      embedFonts,
      fields?.createId,
    );
  } catch (error) {
    if (error instanceof CapabilityError) {
      fail(error.code, error.message, cell.path);
    }
    throw error;
  }
}

function identityPath(
  type: "row" | "cell",
  props: Readonly<Record<string, unknown>>,
  index: number,
): string {
  const identity = props["testId"] ?? props["part"] ?? props["id"];
  return `${type}[${pathLabel(identity, String(index + 1))}]`;
}

function prepareRows(
  input: CapabilityInput,
  columns: readonly Track[],
  tableTextStyle: Record<string, unknown>,
  tableVerticalAlign: TableVerticalAlign,
  theme: ThemeContext,
  slideNumber: number,
): {
  readonly rows: readonly ParsedRow[];
  readonly cells: readonly ParsedCell[];
} {
  const rows: ParsedRow[] = [];
  const cells: ParsedCell[] = [];
  const owners = new Map<string, ParsedCell>();

  for (const child of input.children) {
    if (!("props" in child)) {
      if (child.text.trim() !== "") {
        fail(
          "invalid-table-structure",
          "Text in Table must be placed inside Cell",
        );
      }
      continue;
    }
    if (child.type !== "rp:row") {
      fail("invalid-table-structure", "Table can only contain Row directly");
    }
    const rowIndex = rows.length;
    const rowPath = identityPath("row", child.props, rowIndex);
    onlyKeys(child.props, tablePartParameters, "Row", rowPath);
    const rowAuthorId = optionalName(child.props["id"], "id", rowPath);
    optionalName(child.props["testId"], "testId", rowPath);
    optionalName(child.props["part"], "part", rowPath);
    const rowStyle = trackStyle(child, "row", rowPath);
    const heightValue = rowStyle.style["height"];
    const height =
      heightValue === undefined
        ? undefined
        : positiveSize(heightValue, "Row.style.height", rowPath);
    const rowFill = color(
      rowStyle.style["backgroundColor"],
      "Row.style.backgroundColor",
      theme,
      rowPath,
    );
    const rowVerticalAlign =
      verticalAlign(rowStyle.style["verticalAlign"], rowPath) ??
      tableVerticalAlign;
    const row: ParsedRow = {
      path: rowPath,
      ...(rowAuthorId === undefined ? {} : { authorId: rowAuthorId }),
      ...(typeof child.props["testId"] === "string"
        ? { testId: child.props["testId"] }
        : {}),
      ...(typeof child.props["part"] === "string"
        ? { part: child.props["part"] }
        : {}),
      ...(height === undefined ? {} : { height }),
      textStyle: mergeTextStyle(tableTextStyle, rowStyle.textStyle),
      verticalAlign: rowVerticalAlign,
      ...(rowFill === undefined ? {} : { fill: rowFill }),
    };
    rows.push(row);
    let column = 0;
    let cellIndex = 0;

    for (const cellChild of child.children) {
      if (!("props" in cellChild)) {
        if (cellChild.text.trim() !== "") {
          fail(
            "invalid-table-structure",
            "Text in Row must be placed inside Cell",
            rowPath,
          );
        }
        continue;
      }
      if (cellChild.type !== "rp:cell") {
        fail(
          "invalid-table-structure",
          "Row can only contain Cell directly",
          rowPath,
        );
      }
      const cellPath = `${rowPath}/${identityPath("cell", cellChild.props, cellIndex)}`;
      cellIndex += 1;
      onlyKeys(cellChild.props, cellParameters, "Cell", cellPath);
      const cellAuthorId = optionalName(cellChild.props["id"], "id", cellPath);
      const cellTestId = optionalName(
        cellChild.props["testId"],
        "testId",
        cellPath,
      );
      const cellPart = optionalName(cellChild.props["part"], "part", cellPath);
      while (owners.has(`${rowIndex}:${column}`)) column += 1;
      const colSpan = span(cellChild.props["colSpan"], "colSpan", cellPath);
      const rowSpan = span(cellChild.props["rowSpan"], "rowSpan", cellPath);
      if (column + colSpan > columns.length) {
        fail(
          "invalid-table-merge",
          `Cells in row ${rowIndex + 1} exceed ${columns.length} column count`,
          cellPath,
        );
      }
      if (
        rowIndex + rowSpan >
        input.children.filter(
          (item) => "props" in item && item.type === "rp:row",
        ).length
      ) {
        fail(
          "invalid-table-merge",
          `rowSpan ${rowSpan} extends past the last row of the table`,
          cellPath,
        );
      }
      for (let row = rowIndex; row < rowIndex + rowSpan; row += 1) {
        for (let col = column; col < column + colSpan; col += 1) {
          if (owners.has(`${row}:${col}`)) {
            fail(
              "conflicting-table-merge",
              "Cell merge region overlaps with an existing cell",
              cellPath,
            );
          }
        }
      }
      const cellStyle = trackStyle(cellChild, "cell", cellPath);
      const fill =
        color(
          cellStyle.style["backgroundColor"],
          "Cell.style.backgroundColor",
          theme,
          cellPath,
        ) ?? row.fill;
      const cellVerticalAlign =
        verticalAlign(cellStyle.style["verticalAlign"], cellPath) ??
        rowVerticalAlign;
      const border = mergeBorder(cellStyle.style, theme, cellPath);
      const cellTextStyle = {
        ...mergeTextStyle(row.textStyle, cellStyle.textStyle),
        textWrap: "normal",
      };
      let parsedParagraphs: readonly Paragraph[];
      let parsedTextStyle: ResolvedTextStyle;
      try {
        const textResult = parseTextParagraphs(
          cellChild.children,
          cellTextStyle,
          theme,
          { slideNumber },
          cellPath,
        );
        parsedParagraphs = textResult.paragraphs;
        parsedTextStyle = textResult.style;
      } catch (error) {
        if (error instanceof CapabilityError) {
          fail(error.code, error.message, cellPath);
        }
        throw error;
      }
      const cell: ParsedCell = {
        row: rowIndex,
        column,
        rowSpan,
        colSpan,
        path: cellPath,
        ...(cellAuthorId === undefined ? {} : { authorId: cellAuthorId }),
        ...(cellTestId === undefined ? {} : { testId: cellTestId }),
        ...(cellPart === undefined ? {} : { part: cellPart }),
        paragraphs: parsedParagraphs,
        textStyle: parsedTextStyle,
        verticalAlign: cellVerticalAlign,
        padding: cellPadding(cellStyle.style, cellPath),
        ...(fill === undefined ? {} : { fill }),
        ...(border === undefined ? {} : { border }),
      };
      cells.push(cell);
      for (let row = rowIndex; row < rowIndex + rowSpan; row += 1) {
        for (let col = column; col < column + colSpan; col += 1) {
          owners.set(`${row}:${col}`, cell);
        }
      }
      column += colSpan;
    }
  }

  if (rows.length === 0)
    fail("invalid-table-structure", "Table requires at least one row");
  for (let row = 0; row < rows.length; row += 1) {
    const missing = columns.findIndex(
      (_, column) => !owners.has(`${row}:${column}`),
    );
    if (missing >= 0) {
      fail(
        "invalid-table-structure",
        `Row ${row + 1}, column ${missing + 1} is not covered by a Cell or merged cell`,
        rows[row]?.path,
      );
    }
  }
  return { rows, cells };
}

export type ParsedTable = {
  readonly testId?: string;
  readonly part?: string;
  readonly name?: string;
  readonly morphKey?: string;
  readonly link?: LinkTarget;
  readonly columns: readonly Track[];
  readonly tableStyle: ReturnType<typeof trackStyle>;
  readonly tableVerticalAlign: TableVerticalAlign;
  readonly rows: readonly ParsedRow[];
  readonly cells: readonly ParsedCell[];
  readonly contract: CapabilityLayoutContract;
  readonly rawProps: Readonly<Record<string, unknown>>;
};

type TableLayoutModel = {
  readonly columns: readonly Track[];
  readonly tableStyle: ReturnType<typeof trackStyle>;
  readonly rows: readonly ParsedRow[];
  readonly cells: readonly ParsedCell[];
  readonly widths: readonly number[];
  readonly heights: readonly number[];
  readonly warnings: { code: string; message: string; path?: string }[];
  readonly warningKeys: Set<string>;
  readonly claims: { kind: string; key: string; identity: string }[];
};

/** Shared calculation of table tracks and text dimensions used by LayoutPass and XML emission. */
function tableLayoutModel(
  parsed: ParsedTable,
  requestedWidth: number | undefined,
  theme: ThemeContext,
  slideNumber: number,
  embedFonts?: FontEmbeddingMode,
): TableLayoutModel {
  const { columns, tableStyle, rows, cells } = parsed;
  const warnings: { code: string; message: string; path?: string }[] = [];
  const claims: { kind: string; key: string; identity: string }[] = [];
  const warningKeys = new Set<string>();
  const addTextResult = (
    result: ReturnType<typeof writeTextBodyFromParagraphs>,
    path: string,
  ) => {
    for (const warning of result.warnings) {
      const key = `${path}\u0000${warning.code}\u0000${warning.message}`;
      if (warningKeys.has(key)) continue;
      warningKeys.add(key);
      warnings.push({ ...warning, path });
    }
    claims.push(...result.resourceClaims);
  };
  const naturalWidths = new Map<ParsedCell, number>();
  const needNaturalWidth = columns.some(
    (track) =>
      track.kind === "auto" ||
      (track.kind === "fr" && requestedWidth === undefined),
  );
  if (needNaturalWidth) {
    for (const cell of cells) {
      if (cell.colSpan !== 1) continue;
      const track = columns[cell.column];
      const needsTrackWidth =
        track?.kind === "auto" ||
        (track?.kind === "fr" && requestedWidth === undefined);
      if (!needsTrackWidth) continue;
      const measured = textBody(
        cell,
        theme,
        "max-content",
        undefined,
        false,
        {
          slideNumber,
        },
        embedFonts,
      );
      addTextResult(measured, cell.path);
      if (measured.measurement === undefined && measured.text.trim() !== "") {
        fail(
          "font-measurement-unavailable",
          "auto columns require available font metrics; provide measurement fonts or use explicit column widths",
          cell.path,
        );
      }
      naturalWidths.set(cell, measured.measurement?.width ?? 0);
    }
  }
  const widths = columns.map((track, index) => {
    if (track.kind === "px") return track.value;
    const intrinsic = Math.max(
      0,
      ...cells
        .filter((cell) => cell.column === index && cell.colSpan === 1)
        .map(
          (cell) =>
            (naturalWidths.get(cell) ?? 0) +
            cell.padding.left +
            cell.padding.right,
        ),
    );
    return Math.max(1, intrinsic);
  });
  if (requestedWidth !== undefined) {
    const fractionTotal = columns.reduce(
      (sum, track) => sum + (track.kind === "fr" ? track.value : 0),
      0,
    );
    const fixedWidth = widths.reduce(
      (sum, value, index) => sum + (columns[index]?.kind === "fr" ? 0 : value),
      0,
    );
    if (fractionTotal > 0) {
      const free = requestedWidth - fixedWidth;
      const fractionalColumns = columns.filter(
        (track) => track.kind === "fr",
      ).length;
      if (free >= fractionalColumns) {
        columns.forEach((track, index) => {
          if (track.kind === "fr")
            widths[index] = (free * track.value) / fractionTotal;
        });
      } else {
        columns.forEach((track, index) => {
          if (track.kind === "fr") widths[index] = 1;
        });
      }
    } else {
      const baseWidth = widths.reduce((sum, value) => sum + value, 0);
      if (requestedWidth > baseWidth && baseWidth > 0) {
        widths.forEach((value, index) => {
          widths[index] = (value * requestedWidth) / baseWidth;
        });
      }
    }
    const actualWidth = widths.reduce((sum, value) => sum + value, 0);
    if (actualWidth > requestedWidth + 0.01) {
      warnings.push({
        code: "table-width-overflow",
        message: `Sum of column minimum widths ${actualWidth}px exceeds declared width ${requestedWidth}px; build preserves column widths`,
      });
    }
  }
  const widthsEmu = widths.map((width, index) =>
    pxExtent(width, `column ${index + 1} width`),
  );
  if (widthsEmu.reduce((sum, value) => sum + value, 0) > maximumExtentEmu) {
    fail(
      "invalid-table-geometry",
      "Total table column width exceeds writable range",
    );
  }
  const heights = rows.map((row) => row.height ?? 0);
  const widthOf = (cell: ParsedCell) =>
    widths
      .slice(cell.column, cell.column + cell.colSpan)
      .reduce((sum, width) => sum + width, 0);
  const requiredHeight = (cell: ParsedCell) => {
    const innerWidth = Math.max(
      0,
      widthOf(cell) - cell.padding.left - cell.padding.right,
    );
    const result = textBody(
      cell,
      theme,
      innerWidth,
      undefined,
      false,
      {
        slideNumber,
      },
      embedFonts,
    );
    addTextResult(result, cell.path);
    if (result.measurement === undefined && result.text.trim() !== "") {
      fail(
        "font-measurement-unavailable",
        "Automatic row height requires available font metrics; provide measurement fonts or specify height on Row",
        cell.path,
      );
    }
    return (
      (result.measurement?.height ?? 0) + cell.padding.top + cell.padding.bottom
    );
  };
  for (const cell of cells) {
    if (cell.rowSpan !== 1 || rows[cell.row]?.height !== undefined) continue;
    heights[cell.row] = Math.max(heights[cell.row] ?? 0, requiredHeight(cell));
  }
  for (const cell of cells) {
    if (cell.rowSpan === 1) continue;
    const covered = Array.from(
      { length: cell.rowSpan },
      (_, index) => cell.row + index,
    );
    const growable = covered.filter((row) => rows[row]?.height === undefined);
    if (growable.length === 0) continue;
    const currentHeight = covered.reduce(
      (sum, row) => sum + (heights[row] ?? 0),
      0,
    );
    const missing = requiredHeight(cell) - currentHeight;
    if (missing <= 0) continue;
    const targetRow = growable.at(-1);
    if (targetRow !== undefined)
      heights[targetRow] = (heights[targetRow] ?? 0) + missing;
  }
  heights.forEach((height, index) => {
    if (height <= 0) heights[index] = 1;
  });
  const heightsEmu = heights.map((height, index) =>
    pxExtent(height, `row ${index + 1} height`, rows[index]?.path),
  );
  if (heightsEmu.reduce((sum, value) => sum + value, 0) > maximumExtentEmu) {
    fail(
      "invalid-table-geometry",
      "Total table row height exceeds writable range",
    );
  }
  return {
    columns,
    tableStyle,
    rows,
    cells,
    widths,
    heights,
    warnings,
    warningKeys,
    claims,
  };
}

/** Table module independently validates row and column structure, emitting a single object with native graphicFrame and a:tbl. */
export const tableCapability: ObjectCapabilityModule<ParsedTable> = {
  declaration: "rp:table",

  parse(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): CapabilityParseResult<ParsedTable> {
    onlyKeys(input.props, tableParameters, "Table", context.path);
    const testId = optionalName(input.props["testId"], "testId", context.path);
    const part = optionalName(input.props["part"], "part", context.path);
    const requestedName = optionalName(
      input.props["name"],
      "name",
      context.path,
    );

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...textStyleKeys,
    ]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "graphicFrame",
      props: input.props,
      allowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const columns = parseTracks(input.props["columns"]);
    const tableStyle = trackStyle(input, "table", context.path);
    const tableVerticalAlign =
      verticalAlign(tableStyle.style["verticalAlign"], context.path) ??
      "middle";
    const { rows, cells } = prepareRows(
      input,
      columns,
      tableStyle.textStyle,
      tableVerticalAlign,
      context.theme,
      context.slideNumber,
    );

    const parsed: ParsedTable = {
      ...(testId !== undefined ? { testId } : {}),
      ...(part !== undefined ? { part } : {}),
      ...(requestedName !== undefined ? { name: requestedName } : {}),
      ...(input.props["morphKey"] !== undefined
        ? { morphKey: input.props["morphKey"] as string }
        : {}),
      ...(input.props["link"] !== undefined
        ? { link: input.props["link"] as LinkTarget }
        : {}),
      columns,
      tableStyle,
      tableVerticalAlign,
      rows,
      cells,
      contract,
      rawProps: input.props,
    };

    return { contract, parsed };
  },

  measureLayout(
    parsed: ParsedTable,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement {
    const declaredWidth =
      parsed.contract.layoutStyle.width ?? parsed.rawProps["width"];
    const declaredSize =
      typeof declaredWidth === "number"
        ? positiveSize(declaredWidth, "Table.width")
        : undefined;
    const natural = tableLayoutModel(
      parsed,
      undefined,
      request.theme,
      request.slideNumber,
    );
    const naturalWidth = natural.widths.reduce((sum, width) => sum + width, 0);
    const availableWidth =
      request.availableWidth !== undefined && request.availableWidth > 0
        ? request.availableWidth
        : undefined;
    const requestedWidth =
      declaredSize ??
      (availableWidth !== undefined &&
      (request.widthMode === "exactly" ||
        request.widthMode === undefined ||
        naturalWidth > availableWidth)
        ? availableWidth
        : undefined);
    const model =
      requestedWidth === undefined
        ? natural
        : tableLayoutModel(
            parsed,
            requestedWidth,
            request.theme,
            request.slideNumber,
          );
    const width = model.widths.reduce((sum, value) => sum + value, 0);
    const height = model.heights.reduce((sum, value) => sum + value, 0);
    return {
      width,
      height,
      minContentWidth: natural.widths.reduce((sum, value) => sum + value, 0),
      minContentHeight: natural.heights.reduce((sum, value) => sum + value, 0),
    };
  },

  compile(parsed: ParsedTable, context: CapabilityContext): CapabilityOutput {
    const testId = parsed.testId;
    const part = parsed.part;
    const requestedName = parsed.name;
    const declaredWidth =
      parsed.contract.layoutStyle.width ?? parsed.rawProps["width"];
    const requestedWidth =
      context.layoutBox?.width ??
      (typeof declaredWidth === "number"
        ? positiveSize(declaredWidth, "Table.width")
        : undefined);
    const model = tableLayoutModel(
      parsed,
      requestedWidth,
      context.theme,
      context.slideNumber,
      context.options.embedFonts,
    );
    const { columns, rows, cells, widths } = model;
    const heights = [...model.heights];
    const warnings = [...model.warnings];
    const claims = [...model.claims];
    const fontUses: FontEmbeddingUse[] = [];
    const warningKeys = new Set(model.warningKeys);
    const addTextResult = (
      result: ReturnType<typeof writeTextBodyFromParagraphs>,
      path: string,
    ) => {
      for (const warning of result.warnings) {
        const key = `${path}\u0000${warning.code}\u0000${warning.message}`;
        if (warningKeys.has(key)) continue;
        warningKeys.add(key);
        warnings.push({ ...warning, path });
      }
      claims.push(...result.resourceClaims);
      fontUses.push(...result.fontUses.map((use) => ({ ...use, path })));
    };
    const targetHeight = context.layoutBox?.height;
    if (targetHeight !== undefined) {
      const naturalHeight = heights.reduce((sum, value) => sum + value, 0);
      if (targetHeight > naturalHeight + 0.01) {
        const growable = rows.flatMap((row, index) =>
          row.height === undefined ? [index] : [],
        );
        if (growable.length > 0) {
          const extra = (targetHeight - naturalHeight) / growable.length;
          for (const row of growable)
            heights[row] = (heights[row] ?? 0) + extra;
        } else {
          warnings.push({
            code: "table-height-underflow",
            message: `Sum of fixed row heights ${naturalHeight}px is less than layout height ${targetHeight}px; build preserves row heights`,
          });
        }
      } else if (targetHeight < naturalHeight - 0.01) {
        warnings.push({
          code: "table-height-overflow",
          message: `Sum of row heights ${naturalHeight}px exceeds layout height ${targetHeight}px; build preserves row heights`,
        });
      }
    }
    const x = coordinate(context.layoutBox?.x ?? parsed.rawProps["x"], "x");
    const y = coordinate(context.layoutBox?.y ?? parsed.rawProps["y"], "y");
    const widthOf = (cell: ParsedCell) =>
      widths
        .slice(cell.column, cell.column + cell.colSpan)
        .reduce((sum, width) => sum + width, 0);
    const innerWidthOf = (cell: ParsedCell) =>
      Math.max(0, widthOf(cell) - cell.padding.left - cell.padding.right);

    const anchors = new Map(
      cells.map((cell) => [`${cell.row}:${cell.column}`, cell]),
    );
    const owners = new Map<string, ParsedCell>();
    for (const cell of cells) {
      for (let row = cell.row; row < cell.row + cell.rowSpan; row += 1) {
        for (
          let column = cell.column;
          column < cell.column + cell.colSpan;
          column += 1
        ) {
          owners.set(`${row}:${column}`, cell);
        }
      }
    }
    const matrix = rows.map(() =>
      Array.from({ length: columns.length }, () => ""),
    );
    const rowsXml = rows.map((_, rowIndex) => {
      const rowCells: string[] = [];
      for (let column = 0; column < columns.length; column += 1) {
        const cell = anchors.get(`${rowIndex}:${column}`);
        if (cell === undefined) {
          const owner = owners.get(`${rowIndex}:${column}`);
          if (owner === undefined) {
            fail(
              "invalid-table-structure",
              "Table row contains undeclared cells",
            );
          }
          const hMerge = column > owner.column ? ' hMerge="1"' : "";
          const vMerge = rowIndex > owner.row ? ' vMerge="1"' : "";
          const spanAttributes = [
            ...(rowIndex === owner.row && owner.rowSpan > 1
              ? [`rowSpan="${owner.rowSpan}"`]
              : []),
            ...(column === owner.column && owner.colSpan > 1
              ? [`gridSpan="${owner.colSpan}"`]
              : []),
          ];
          rowCells.push(
            `<a:tc${hMerge}${vMerge}${spanAttributes.length === 0 ? "" : ` ${spanAttributes.join(" ")}`}>${emptyTableCellXml(owner?.textStyle?.["lang"])}</a:tc>`,
          );
          continue;
        }
        const innerHeight = Math.max(
          0,
          heights
            .slice(cell.row, cell.row + cell.rowSpan)
            .reduce((sum, value) => sum + value, 0) -
            cell.padding.top -
            cell.padding.bottom,
        );
        const contentSized = rows
          .slice(cell.row, cell.row + cell.rowSpan)
          .some((row) => row.height === undefined);
        const body = textBody(
          cell,
          context.theme,
          innerWidthOf(cell),
          innerHeight,
          contentSized,
          {
            slideNumber: context.slideNumber,
            createId: context.createTextFieldId,
          },
          context.options.embedFonts,
        );
        addTextResult(body, cell.path);
        matrix[rowIndex]![column] = body.text;
        const spanAttributes = [
          ...(cell.colSpan > 1 ? [`gridSpan="${cell.colSpan}"`] : []),
          ...(cell.rowSpan > 1 ? [`rowSpan="${cell.rowSpan}"`] : []),
        ];
        const attributes =
          spanAttributes.length === 0 ? "" : ` ${spanAttributes.join(" ")}`;
        const rowFill = rows[cell.row]?.fill;
        const fill = cell.fill ?? rowFill;
        const properties = tableCellPropertiesXml(
          cell.padding,
          cell.verticalAlign,
          cell.border,
          fill,
        );
        rowCells.push(`<a:tc${attributes}>${body.xml}${properties}</a:tc>`);
      }
      return rowCells.join("");
    });
    for (const warning of warnings) {
      if (warning.code === "overset") {
        warning.message = `${warning.message} (table cell preserves original content and declared dimensions)`;
      }
    }
    const name =
      context.morphName ?? requestedName ?? testId ?? `Table ${context.id - 1}`;
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      fail(
        "invalid-xml-character",
        "Table.name contains characters not allowed in XML 1.0",
      );
    }
    const width = widths.reduce((sum, value) => sum + value, 0);
    const height = heights.reduce((sum, value) => sum + value, 0);
    const box = { x, y, width, height };
    const tableParts = [
      ...rows.map((row, rowIndex) => ({
        kind: "row" as const,
        ...(row.authorId === undefined ? {} : { authorId: row.authorId }),
        ...(row.testId === undefined ? {} : { testId: row.testId }),
        ...(row.part === undefined ? {} : { part: row.part }),
        row: rowIndex,
        box: {
          x,
          y:
            y +
            heights.slice(0, rowIndex).reduce((sum, value) => sum + value, 0),
          width,
          height: heights[rowIndex] ?? 0,
        },
        text: matrix[rowIndex]?.join("\t") ?? "",
      })),
      ...cells.map((cell) => ({
        kind: "cell" as const,
        ...(cell.authorId === undefined ? {} : { authorId: cell.authorId }),
        ...(cell.testId === undefined ? {} : { testId: cell.testId }),
        ...(cell.part === undefined ? {} : { part: cell.part }),
        row: cell.row,
        column: cell.column,
        rowSpan: cell.rowSpan,
        colSpan: cell.colSpan,
        box: {
          x:
            x +
            widths.slice(0, cell.column).reduce((sum, value) => sum + value, 0),
          y:
            y +
            heights.slice(0, cell.row).reduce((sum, value) => sum + value, 0),
          width: widths
            .slice(cell.column, cell.column + cell.colSpan)
            .reduce((sum, value) => sum + value, 0),
          height: heights
            .slice(cell.row, cell.row + cell.rowSpan)
            .reduce((sum, value) => sum + value, 0),
        },
        text: matrix[cell.row]?.[cell.column] ?? "",
      })),
    ];
    const observation: NativeTableObservation = {
      columnWidths: [...widths],
      rowHeights: [...heights],
      cells: matrix,
      merges: cells
        .filter((cell) => cell.colSpan > 1 || cell.rowSpan > 1)
        .map(({ row, column, rowSpan, colSpan }) => ({
          row,
          column,
          rowSpan,
          colSpan,
        })),
      ...(tableParts.some(
        ({ authorId, testId, part }) =>
          authorId !== undefined || testId !== undefined || part !== undefined,
      )
        ? { parts: tableParts }
        : {}),
    };
    const annotations = motionAnnotations(parsed.rawProps, {
      kind: "table",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: true,
      filled: cells.some(
        (cell) => cell.fill !== undefined || rows[cell.row]?.fill !== undefined,
      ),
      box,
      rotationDegrees: 0,
      hasBackground: cells.some(
        (cell) => cell.fill !== undefined || rows[cell.row]?.fill !== undefined,
      ),
    });
    return {
      xml: tableFrameXml(
        context.id,
        escapedName,
        x,
        y,
        widths,
        heights,
        rowsXml,
        nativeLinkXml(context.nativeLink),
      ),
      observation: {
        id: context.id,
        kind: "table",
        name,
        ...(testId === undefined ? {} : { testId }),
        ...(part === undefined ? {} : { part }),
        box,
        table: observation,
      },
      extensions: [],
      ...(annotations.length === 0 ? {} : { annotations }),
      resourceClaims: claims,
      presentationContributions: fontPresentationContributions(
        fontUses,
        context.options.embedFonts,
      ),
      warnings,
    };
  },
};

export type TableNode = CommonContainer<
  "table",
  TableDeclarationProps,
  NonNullable<TableDeclarationProps["style"]>
>;
export type RowNode = CommonContainer<
  "row",
  RowDeclarationProps,
  NonNullable<RowDeclarationProps["style"]>
>;
export type CellNode = RichTextContainer<
  "cell",
  CellDeclarationProps,
  NonNullable<CellDeclarationProps["style"]>
>;

/** Creates the editable native table shared by the command factory and the Table JSX element. */
export function createTableNode(props: TableDeclarationProps): TableNode {
  return new CommonContainer<
    "table",
    TableDeclarationProps,
    NonNullable<TableDeclarationProps["style"]>
  >("table", "rp:table", props);
}

/** Creates the editable row shared by command authoring and the Row JSX element. */
export function createRowNode(props: RowDeclarationProps): RowNode {
  return new CommonContainer<
    "row",
    RowDeclarationProps,
    NonNullable<RowDeclarationProps["style"]>
  >("row", "rp:row", props);
}

/** Creates the editable rich-text cell shared by command authoring and the Cell JSX element. */
export function createCellNode(props: CellDeclarationProps): CellNode {
  return new RichTextContainer<
    "cell",
    CellDeclarationProps,
    NonNullable<CellDeclarationProps["style"]>
  >("cell", "rp:cell", props);
}
