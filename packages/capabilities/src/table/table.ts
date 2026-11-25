import {
  assertValidLang,
  CapabilityError,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  type BoundObjectCapability,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityLayoutContract,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type CapabilitySubPartObservation,
  type ColorValue,
  type EmitContext,
  type EmitOutput,
  type GeometryBox,
  type NativePlaceholderRef,
  type ObjectCapabilityModule,
  type ObjectResolution,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type ResolvedColor,
  type ResolvedTable,
  type ResolvedTableCell as ResolvedTableCellObservation,
  type ResolvedTableRow,
  type ThemeContext,
} from "@pptxdk/core/capability";
import { assertXml10Characters, xmlEscape } from "@pptxdk/emit";
import { defineCanvasMechanism } from "@pptxdk/import";

import {
  fontPresentationContributions,
  type FontEmbeddingUse,
} from "../font/font-embedding.ts";
import { motionAnnotations } from "../motion/motion-api.ts";
import {
  nativeLinkXml,
  parseLinkTarget,
  parseMorphKey,
  parseOptionalName,
} from "../shared/identity.ts";
import { resolvePlaceholderGeometry } from "../shared/style-inheritance.ts";
import {
  emitTextBody,
  parseTextParagraphs,
  renderTextOutlinesSvg,
  resolveTextBody,
  type Paragraph,
  type ResolvedTextBody,
  type ResolvedTextStyle,
} from "../shared/text-service.ts";
import { tableCodegen } from "./table-codegen.ts";
import type { TableVerticalAlign } from "./table-declaration.ts";
import {
  tableImportHandler,
  tablePlaceholderImportHandler,
} from "./table-import.ts";
import {
  assertAllowedTableKeys,
  assertTableEmuExtent,
  cellAllowedParameters,
  cellStyleKeys,
  parseOptionalTableId,
  parseTableCellBorder,
  parseTableCellPadding,
  parseTableColumns,
  parseTableCoordinate,
  parseTableMergeSpan,
  parseTablePositiveSize,
  parseTableVerticalAlign,
  rowAllowedParameters,
  rowStyleKeys,
  tableAllowedParameters,
  tableStyleKeys,
  tableTextStyleKeys,
  type TableTrack,
} from "./table-schemas.ts";
import {
  emptyTableCellXml,
  tableCellPropertiesXml,
  tableFrameXml,
  type TableCellBorder,
  type TableCellPadding,
} from "./table-xml.ts";

export {
  createCellNode,
  createRowNode,
  createTableNode,
  type CellDeclarationProps,
  type CellNode,
  type RowDeclarationProps,
  type RowNode,
  type TableCellStyle,
  type TableColumn,
  type TableDeclarationProps,
  type TableNode,
  type TableRowStyle,
  type TableSpacing,
  type TableStyle,
  type TableTextStyle,
  type TableVerticalAlign,
} from "./table-declaration.ts";

type Track = TableTrack;
type Edges = TableCellPadding;

type ParsedCell = {
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly colSpan: number;
  readonly path: string;
  readonly authorId?: string;
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
  readonly height?: number;
  readonly textStyle: Record<string, unknown>;
  readonly verticalAlign?: TableVerticalAlign;
  readonly fill?: ResolvedColor;
};
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

function pathLabel(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim() !== "" ? value : fallback;
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
  assertAllowedTableKeys(
    style,
    [...tableTextStyleKeys, ...additional],
    `${label}.style`,
    path,
  );
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
  for (const key of tableTextStyleKeys) {
    const value = style[key];
    if (value !== undefined) textStyle[key] = value;
  }
  parseTableVerticalAlign(style["verticalAlign"], path);
  return { textStyle, style };
}

function mergeTextStyle(
  parent: Record<string, unknown>,
  child: Record<string, unknown>,
): Record<string, unknown> {
  return { ...parent, ...child };
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

function resolveCellBody(
  cell: ParsedCell,
  theme: ThemeContext,
  width: number | "max-content",
  height?: number,
  contentSized = false,
): ResolvedTextBody {
  try {
    return resolveTextBody(
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
  return `${type}[${pathLabel(props["id"], String(index + 1))}]`;
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
    if (child.type !== "row") {
      fail("invalid-table-structure", "Table can only contain Row directly");
    }
    const rowIndex = rows.length;
    const rowPath = identityPath("row", child.props, rowIndex);
    assertAllowedTableKeys(child.props, rowAllowedParameters, "Row", rowPath);
    const rowAuthorId = parseOptionalTableId(child.props["id"], rowPath);
    const rowStyle = trackStyle(child, "row", rowPath);
    const heightValue = rowStyle.style["height"];
    const height =
      heightValue === undefined
        ? undefined
        : parseTablePositiveSize(heightValue, "Row.style.height", rowPath);
    const rowFill = color(
      rowStyle.style["backgroundColor"],
      "Row.style.backgroundColor",
      theme,
      rowPath,
    );
    const rowVerticalAlign =
      parseTableVerticalAlign(rowStyle.style["verticalAlign"], rowPath) ??
      tableVerticalAlign;
    const row: ParsedRow = {
      path: rowPath,
      ...(rowAuthorId === undefined ? {} : { authorId: rowAuthorId }),
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
      if (cellChild.type !== "cell") {
        fail(
          "invalid-table-structure",
          "Row can only contain Cell directly",
          rowPath,
        );
      }
      const cellPath = `${rowPath}/${identityPath("cell", cellChild.props, cellIndex)}`;
      cellIndex += 1;
      assertAllowedTableKeys(
        cellChild.props,
        cellAllowedParameters,
        "Cell",
        cellPath,
      );
      const cellAuthorId = parseOptionalTableId(
        cellChild.props["id"],
        cellPath,
      );
      while (owners.has(`${rowIndex}:${column}`)) column += 1;
      const colSpan = parseTableMergeSpan(
        cellChild.props["colSpan"],
        "colSpan",
        cellPath,
      );
      const rowSpan = parseTableMergeSpan(
        cellChild.props["rowSpan"],
        "rowSpan",
        cellPath,
      );
      if (column + colSpan > columns.length) {
        fail(
          "invalid-table-merge",
          `Cells in row ${rowIndex + 1} exceed ${columns.length} column count`,
          cellPath,
        );
      }
      if (
        rowIndex + rowSpan >
        input.children.filter((item) => "props" in item && item.type === "row")
          .length
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
        parseTableVerticalAlign(cellStyle.style["verticalAlign"], cellPath) ??
        rowVerticalAlign;
      const border = parseTableCellBorder(
        cellStyle.style,
        (value, label, itemPath) => color(value, label, theme, itemPath),
        theme,
        cellPath,
      );
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
        paragraphs: parsedParagraphs,
        textStyle: parsedTextStyle,
        verticalAlign: cellVerticalAlign,
        padding: parseTableCellPadding(cellStyle.style, cellPath),
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

type ResolvedTableCell = {
  readonly cell: ParsedCell;
  readonly box: GeometryBox;
  readonly fill?: ResolvedColor;
  readonly body: ResolvedTextBody;
  readonly hasText: boolean;
  readonly hasOverflow: boolean;
};

/** Closed resolved table model produced in resolve() and exhaustively consumed by emit and paint. */
type ResolvedTableModel = {
  readonly x: number;
  readonly y: number;
  readonly widths: readonly number[];
  readonly heights: readonly number[];
  readonly rows: readonly ParsedRow[];
  readonly columns: readonly Track[];
  readonly resolvedCells: readonly ResolvedTableCell[];
  readonly resolvedAnchors: ReadonlyMap<string, ResolvedTableCell>;
  readonly owners: ReadonlyMap<string, ParsedCell>;
  readonly name: string;
  readonly fontUses: readonly FontEmbeddingUse[];
  readonly placeholder?: NativePlaceholderRef;
};

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function tableBorderAttributes(border: TableCellBorder): string {
  const w = Math.max(1, border.width);
  let dashAttr = "";
  switch (border.dash) {
    case "solid":
      dashAttr = "";
      break;
    case "dash":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w * 4)} ${roundSvgNumber(w * 3)}"`;
      break;
    case "dot":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w)} ${roundSvgNumber(w * 2)}"`;
      break;
    default: {
      const exhaustive: never = border.dash;
      return exhaustive;
    }
  }
  return `fill="none" stroke="#${border.color.hex}" stroke-opacity="${roundSvgNumber(border.color.alpha ?? 1)}" stroke-width="${roundSvgNumber(border.width)}"${dashAttr}`;
}

function emitTableModel(
  model: ResolvedTableModel,
  emitContext: EmitContext,
): EmitOutput {
  const {
    x,
    y,
    widths,
    heights,
    rows,
    columns,
    resolvedAnchors,
    owners,
    name,
    fontUses,
  } = model;
  const embedFonts = emitContext.options.embedFonts;
  const rowsXml = rows.map((_, rowIndex) => {
    const rowCells: string[] = [];
    for (let column = 0; column < columns.length; column += 1) {
      const entry = resolvedAnchors.get(`${rowIndex}:${column}`);
      if (entry === undefined) {
        const owner = owners.get(`${rowIndex}:${column}`)!;
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
          `<a:tc${hMerge}${vMerge}${spanAttributes.length === 0 ? "" : ` ${spanAttributes.join(" ")}`}>${emptyTableCellXml(owner.textStyle["lang"])}</a:tc>`,
        );
        continue;
      }
      const { cell, body, fill } = entry;
      if (
        (embedFonts === "full" || embedFonts === "subset") &&
        body.missingGlyphErrors.length > 0
      ) {
        fail("font-missing-glyph", body.missingGlyphErrors[0]!, cell.path);
      }
      let emittedXml: string;
      try {
        emittedXml = emitTextBody(body, emitContext.theme, {
          createFieldId: emitContext.createTextFieldId,
        }).xml;
      } catch (error) {
        if (error instanceof CapabilityError) {
          fail(error.code, error.message, error.path ?? cell.path);
        }
        throw error;
      }
      const spanAttributes = [
        ...(cell.colSpan > 1 ? [`gridSpan="${cell.colSpan}"`] : []),
        ...(cell.rowSpan > 1 ? [`rowSpan="${cell.rowSpan}"`] : []),
      ];
      const attributes =
        spanAttributes.length === 0 ? "" : ` ${spanAttributes.join(" ")}`;
      const properties = tableCellPropertiesXml(
        cell.padding,
        cell.verticalAlign,
        cell.border,
        fill,
      );
      rowCells.push(`<a:tc${attributes}>${emittedXml}${properties}</a:tc>`);
    }
    return rowCells.join("");
  });
  const presentationContributions = fontPresentationContributions(
    fontUses,
    embedFonts,
  );
  return {
    xml: tableFrameXml(
      emitContext.id,
      xmlEscape(name),
      x,
      y,
      widths,
      heights,
      rowsXml,
      nativeLinkXml(emitContext.nativeLink),
      model.placeholder,
    ),
    ...(presentationContributions.length === 0
      ? {}
      : { presentationContributions }),
  };
}

function paintTableModel(
  model: ResolvedTableModel,
  paintContext: PaintContext,
): PaintOutput {
  const defs: string[] = [];
  const fills: string[] = [];
  const texts: string[] = [];
  const borders: string[] = [];
  for (const entry of model.resolvedCells) {
    const { cell, box: cellBox, fill, body, hasText, hasOverflow } = entry;
    if (fill !== undefined) {
      fills.push(
        `<rect x="${roundSvgNumber(cellBox.x)}" y="${roundSvgNumber(cellBox.y)}" width="${roundSvgNumber(cellBox.width)}" height="${roundSvgNumber(cellBox.height)}" fill="#${fill.hex}" fill-opacity="${roundSvgNumber(fill.alpha ?? 1)}" stroke="none"/>`,
      );
    }
    if (hasText) {
      const idPrefix = `tbl-s${paintContext.slideNumber}-o${paintContext.objectId}-r${cell.row}-c${cell.column}`;
      try {
        const glyphs = renderTextOutlinesSvg({
          paragraphs: body.paragraphs,
          box: cellBox,
          insets: cell.padding,
          verticalAlign: cell.verticalAlign,
          theme: paintContext.theme,
          measuredLayout: body.measuredLayout,
          idPrefix,
        });
        if (glyphs !== "") {
          if (hasOverflow) {
            const clipId = `${idPrefix}-clip`;
            defs.push(
              `<clipPath id="${clipId}" clipPathUnits="userSpaceOnUse"><rect x="${roundSvgNumber(cellBox.x)}" y="${roundSvgNumber(cellBox.y)}" width="${roundSvgNumber(cellBox.width)}" height="${roundSvgNumber(cellBox.height)}"/></clipPath>`,
            );
            texts.push(`<g clip-path="url(#${clipId})">${glyphs}</g>`);
          } else {
            texts.push(glyphs);
          }
        }
      } catch (error) {
        if (error instanceof CapabilityError) {
          throw new CapabilityError(
            error.code,
            error.message,
            error.path ?? cell.path,
          );
        }
        throw error;
      }
    }
    if (cell.border !== undefined) {
      borders.push(
        `<rect x="${roundSvgNumber(cellBox.x)}" y="${roundSvgNumber(cellBox.y)}" width="${roundSvgNumber(cellBox.width)}" height="${roundSvgNumber(cellBox.height)}" ${tableBorderAttributes(cell.border)}/>`,
      );
    }
  }
  const defsXml = defs.length > 0 ? `<defs>${defs.join("")}</defs>` : "";
  return {
    svg: `${defsXml}<g>${fills.join("")}${texts.join("")}${borders.join("")}</g>`,
  };
}

type ParsedTable = {
  readonly columns: readonly Track[];
  readonly tableStyle: ReturnType<typeof trackStyle>;
  readonly tableVerticalAlign: TableVerticalAlign;
  readonly rows: readonly ParsedRow[];
  readonly cells: readonly ParsedCell[];
  readonly contract: CapabilityLayoutContract;
  readonly declaredX: unknown;
  readonly declaredY: unknown;
  readonly declaredWidth: unknown;
  readonly declaredHeight: unknown;
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

/** Shared calculation of table tracks and text dimensions used by layout measurement and table resolution. */
function tableLayoutModel(
  parsed: ParsedTable,
  requestedWidth: number | undefined,
  theme: ThemeContext,
): TableLayoutModel {
  const { columns, tableStyle, rows, cells } = parsed;
  const warnings: { code: string; message: string; path?: string }[] = [];
  const claims: { kind: string; key: string; identity: string }[] = [];
  const warningKeys = new Set<string>();
  const addTextResult = (result: ResolvedTextBody, path: string) => {
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
      const measured = resolveCellBody(
        cell,
        theme,
        "max-content",
        undefined,
        false,
      );
      addTextResult(measured, cell.path);
      if (
        measured.textMetrics?.status === "unmeasured" &&
        measured.textMetrics.reason === "missing-glyphs"
      ) {
        fail(
          "font-missing-glyph",
          measured.missingGlyphErrors[0] ??
            "auto columns require glyphs for all visible characters",
          cell.path,
        );
      }
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
    assertTableEmuExtent(width, `column ${index + 1} width`),
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
    const result = resolveCellBody(cell, theme, innerWidth, undefined, false);
    addTextResult(result, cell.path);
    if (
      result.textMetrics?.status === "unmeasured" &&
      result.textMetrics.reason === "missing-glyphs"
    ) {
      fail(
        "font-missing-glyph",
        result.missingGlyphErrors[0] ??
          "Automatic row height requires glyphs for all visible characters",
        cell.path,
      );
    }
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
    assertTableEmuExtent(height, `row ${index + 1} height`, rows[index]?.path),
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

const tableCompatiblePlaceholderRoles = new Set<string>(["tbl", "obj"]);

/** Typed native canvas mechanism connecting emit, paint, and import for `table`. */
export const tableMechanism = defineCanvasMechanism({
  key: "table",
  role: "canvas",
  emit: emitTableModel,
  paint: paintTableModel,
  import: tableImportHandler,
});

/** Typed native canvas mechanism connecting emit, paint, and import for `table-placeholder`. */
export const tablePlaceholderMechanism = defineCanvasMechanism({
  key: "table-placeholder",
  role: "canvas",
  emit: emitTableModel,
  paint: paintTableModel,
  import: tablePlaceholderImportHandler,
});

/** Table module independently validates row and column structure, resolving layout/cells once for emit and paint. */
export const tableCapability = {
  declaration: "table",
  mechanisms: [tableMechanism, tablePlaceholderMechanism] as const,
  codegen: tableCodegen,

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): BoundObjectCapability {
    const placeholder = context.placeholder;
    if (placeholder !== undefined) {
      const role =
        placeholder.layoutPlaceholder?.type ?? placeholder.nativeRef.type;
      if (!tableCompatiblePlaceholderRoles.has(role)) {
        fail(
          "incompatible-placeholder-content",
          `Table cannot be bound to ${role} placeholder`,
        );
      }
    }
    assertAllowedTableKeys(input.props, tableAllowedParameters, "Table");
    const requestedName = parseOptionalName(input.props["name"]);
    const morphKey = parseMorphKey(input.props["morphKey"]);
    const link = parseLinkTarget(input.props["link"]);

    const allowedStyleKeys = new Set<string>(tableStyleKeys);

    const contract: CapabilityLayoutContract = {
      ...parseCapabilityLayoutContract({
        nativeKind: "graphicFrame",
        props: input.props,
        allowedStyleKeys,
        allowPadding: false,
        path: context.path,
      }),
      contentMinimumsEnforced: true,
    };

    const columns = parseTableColumns(input.props["columns"]);
    const tableStyle = trackStyle(input, "table");
    const tableVerticalAlign =
      parseTableVerticalAlign(tableStyle.style["verticalAlign"]) ?? "middle";
    const { rows, cells } = prepareRows(
      input,
      columns,
      tableStyle.textStyle,
      tableVerticalAlign,
      context.theme,
      context.slideNumber,
    );

    const placeholderGeometry =
      placeholder === undefined
        ? undefined
        : resolvePlaceholderGeometry(placeholder, input.props);
    const parsed: ParsedTable = {
      columns,
      tableStyle,
      tableVerticalAlign,
      rows,
      cells,
      contract,
      declaredX: placeholderGeometry?.x ?? input.props["x"],
      declaredY: placeholderGeometry?.y ?? input.props["y"],
      declaredWidth:
        placeholderGeometry?.width ??
        contract.layoutStyle.width ??
        input.props["width"],
      declaredHeight:
        placeholderGeometry?.height ??
        contract.layoutStyle.height ??
        input.props["height"],
    };

    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["emphasis"] !== undefined
        ? { emphasis: input.props["emphasis"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(requestedName === undefined ? {} : { hasExplicitName: true }),
      ...(link === undefined ? {} : { link }),
    };

    return {
      contract,
      ...(morphKey !== undefined ||
      requestedName !== undefined ||
      link !== undefined
        ? { native }
        : {}),

      measureLayout(
        request: CapabilityLayoutRequest,
      ): CapabilityLayoutMeasurement {
        const declaredSize =
          typeof parsed.declaredWidth === "number"
            ? parseTablePositiveSize(parsed.declaredWidth, "Table.width")
            : undefined;
        const declaredHeight =
          typeof parsed.declaredHeight === "number"
            ? parseTablePositiveSize(parsed.declaredHeight, "Table.height")
            : undefined;
        const availableWidth =
          request.availableWidth !== undefined && request.availableWidth > 0
            ? request.availableWidth
            : undefined;
        const canSkipNatural =
          (declaredSize !== undefined || request.widthMode === "exactly") &&
          !columns.some((track) => track.kind === "auto");
        const natural = canSkipNatural
          ? undefined
          : tableLayoutModel(parsed, undefined, request.theme);
        const naturalWidth =
          natural?.widths.reduce((sum, width) => sum + width, 0) ?? 0;
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
            ? (natural ?? tableLayoutModel(parsed, undefined, request.theme))
            : tableLayoutModel(parsed, requestedWidth, request.theme);
        const width = model.widths.reduce((sum, value) => sum + value, 0);
        const baseHeight = model.heights.reduce((sum, value) => sum + value, 0);
        const targetHeight =
          declaredHeight ??
          (request.heightMode === "exactly"
            ? request.availableHeight
            : undefined);
        const hasAutoRow = rows.some((row) => row.height === undefined);
        const height =
          targetHeight !== undefined &&
          targetHeight > baseHeight + 0.01 &&
          hasAutoRow
            ? targetHeight
            : baseHeight;
        return {
          width,
          height,
          minContentWidth: natural === undefined ? width : naturalWidth,
          minContentHeight: baseHeight,
        };
      },

      resolve(resolveContext: ObjectResolveContext): ObjectResolution {
        const declaredRequestedWidth =
          typeof parsed.declaredWidth === "number"
            ? parseTablePositiveSize(parsed.declaredWidth, "Table.width")
            : undefined;
        const requestedWidth =
          declaredRequestedWidth ?? resolveContext.layoutBox?.width;
        const model = tableLayoutModel(
          parsed,
          requestedWidth,
          resolveContext.theme,
        );
        const { widths } = model;
        const heights = [...model.heights];
        const warnings = [...model.warnings];
        const claims = [...model.claims];
        const fontUses: FontEmbeddingUse[] = [];
        const warningKeys = new Set(model.warningKeys);
        const addTextResult = (result: ResolvedTextBody, path: string) => {
          for (const warning of result.warnings) {
            const key = `${path}\u0000${warning.code}\u0000${warning.message}`;
            if (warningKeys.has(key)) continue;
            warningKeys.add(key);
            warnings.push({ ...warning, path });
          }
          claims.push(...result.resourceClaims);
          fontUses.push(...result.fontUses.map((use) => ({ ...use, path })));
        };
        const declaredTargetHeight =
          typeof parsed.declaredHeight === "number"
            ? parseTablePositiveSize(parsed.declaredHeight, "Table.height")
            : undefined;
        const targetHeight =
          declaredTargetHeight ?? resolveContext.layoutBox?.height;
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
        const x = parseTableCoordinate(
          resolveContext.layoutBox?.x ?? parsed.declaredX,
          "x",
        );
        const y = parseTableCoordinate(
          resolveContext.layoutBox?.y ?? parsed.declaredY,
          "y",
        );
        const width = widths.reduce((sum, value) => sum + value, 0);
        const height = heights.reduce((sum, value) => sum + value, 0);
        const box: GeometryBox = { x, y, width, height };
        const columnOffsets: number[] = [];
        let offsetX = x;
        for (const w of widths) {
          columnOffsets.push(offsetX);
          offsetX += w;
        }
        const rowOffsets: number[] = [];
        let offsetY = y;
        for (const h of heights) {
          rowOffsets.push(offsetY);
          offsetY += h;
        }
        const cellBoxes = new Map<ParsedCell, GeometryBox>();
        for (const cell of cells) {
          cellBoxes.set(cell, {
            x: columnOffsets[cell.column] ?? x,
            y: rowOffsets[cell.row] ?? y,
            width: widths
              .slice(cell.column, cell.column + cell.colSpan)
              .reduce((sum, value) => sum + value, 0),
            height: heights
              .slice(cell.row, cell.row + cell.rowSpan)
              .reduce((sum, value) => sum + value, 0),
          });
        }
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
        const resolvedCells: ResolvedTableCell[] = [];
        const resolvedAnchors = new Map<string, ResolvedTableCell>();
        const fontUsages: {
          readonly family: string;
          readonly weight: number;
          readonly style: "normal" | "italic";
          readonly hasResource: boolean;
          readonly substitute?: boolean;
          readonly embeddable?: boolean;
          readonly missingGlyphs?: readonly string[];
          readonly path: string;
        }[] = [];
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
          for (let column = 0; column < columns.length; column += 1) {
            const cell = anchors.get(`${rowIndex}:${column}`);
            if (cell === undefined) {
              if (owners.get(`${rowIndex}:${column}`) === undefined) {
                fail(
                  "invalid-table-structure",
                  "Table row contains undeclared cells",
                );
              }
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
            const body = resolveCellBody(
              cell,
              resolveContext.theme,
              innerWidthOf(cell),
              innerHeight,
              contentSized,
            );
            addTextResult(body, cell.path);
            for (const usage of body.fontUsages) {
              fontUsages.push({ ...usage, path: cell.path });
            }
            matrix[rowIndex]![column] = body.text;
            const rowFill = rows[cell.row]?.fill;
            const fill = cell.fill ?? rowFill;
            const entry: ResolvedTableCell = {
              cell,
              box: cellBoxes.get(cell)!,
              ...(fill === undefined ? {} : { fill }),
              body,
              hasText: cell.paragraphs.some(
                (paragraph) =>
                  paragraph.style.list !== undefined ||
                  paragraph.runs.some((run) => run.text.length > 0),
              ),
              hasOverflow:
                body.textMetrics?.status === "measured" &&
                !body.textMetrics.fits,
            };
            resolvedCells.push(entry);
            resolvedAnchors.set(`${rowIndex}:${column}`, entry);
          }
        }
        for (const warning of warnings) {
          if (warning.code === "overset") {
            warning.message = `${warning.message} (table cell preserves original content and declared dimensions)`;
          }
        }
        const name =
          resolveContext.morphName ??
          requestedName ??
          placeholder?.layoutPlaceholder?.name ??
          `Table ${resolveContext.id - 1}`;
        try {
          assertXml10Characters(name);
        } catch {
          fail(
            "invalid-xml-character",
            "Table.name contains characters not allowed in XML 1.0",
          );
        }
        const resolvedRowObservations: ResolvedTableRow[] = rows.map(
          (row, rowIndex) => {
            const rowY = rowOffsets[rowIndex] ?? y;
            const rowHeight = heights[rowIndex] ?? 0;
            const rowPath = `${resolveContext.path}/row[${row.authorId ?? rowIndex}]`;
            return {
              index: rowIndex,
              ...(row.authorId === undefined ? {} : { id: row.authorId }),
              path: rowPath,
              y: rowY,
              height: rowHeight,
              box: {
                x,
                y: rowY,
                width,
                height: rowHeight,
              },
            };
          },
        );
        const resolvedCellObservations: ResolvedTableCellObservation[] =
          cells.map((cell) => {
            const anchor = resolvedAnchors.get(`${cell.row}:${cell.column}`);
            const textMetrics = anchor?.body.textMetrics;
            if (textMetrics === undefined) {
              throw new Error(
                `Missing textMetrics for table cell (${cell.row}, ${cell.column})`,
              );
            }
            const rowPath =
              resolvedRowObservations[cell.row]?.path ??
              `${resolveContext.path}/row[${cell.row}]`;
            const cellPath = `${rowPath}/cell[${cell.authorId ?? cell.column}]`;
            return {
              row: cell.row,
              column: cell.column,
              rowSpan: cell.rowSpan,
              colSpan: cell.colSpan,
              ...(cell.authorId === undefined ? {} : { id: cell.authorId }),
              path: cellPath,
              box: cellBoxes.get(cell)!,
              text: matrix[cell.row]?.[cell.column] ?? "",
              textMetrics,
            };
          });
        const tableParts = [
          ...resolvedRowObservations.map((rowObs) => ({
            kind: "row" as const,
            ...(rowObs.id === undefined ? {} : { authorId: rowObs.id }),
            row: rowObs.index,
            box: rowObs.box,
            text: matrix[rowObs.index]?.join("\t") ?? "",
          })),
          ...resolvedCellObservations.map((cellObs) => ({
            kind: "cell" as const,
            ...(cellObs.id === undefined ? {} : { authorId: cellObs.id }),
            row: cellObs.row,
            column: cellObs.column,
            rowSpan: cellObs.rowSpan,
            colSpan: cellObs.colSpan,
            box: cellObs.box,
            text: cellObs.text,
            textMetrics: cellObs.textMetrics,
          })),
        ];
        const subParts: CapabilitySubPartObservation[] =
          resolvedRowObservations.map((rowObs) => ({
            kind: "row",
            ...(rowObs.id === undefined ? {} : { authorId: rowObs.id }),
            path: rowObs.path,
            box: rowObs.box,
            text: matrix[rowObs.index]?.join("\t") ?? "",
            parts: resolvedCellObservations
              .filter((cellObs) => cellObs.row === rowObs.index)
              .map((cellObs) => ({
                kind: "cell" as const,
                ...(cellObs.id === undefined ? {} : { authorId: cellObs.id }),
                path: cellObs.path,
                box: cellObs.box,
                text: cellObs.text,
                textMetrics: cellObs.textMetrics,
              })),
          }));
        const observation: ResolvedTable = {
          columns: [...widths],
          rows: resolvedRowObservations,
          cells: resolvedCellObservations,
          merges: cells
            .filter((cell) => cell.colSpan > 1 || cell.rowSpan > 1)
            .map(({ row, column, rowSpan, colSpan }) => ({
              row,
              column,
              rowSpan,
              colSpan,
            })),
          columnWidths: [...widths],
          rowHeights: [...heights],
          parts: tableParts,
        };
        const annotations = motionAnnotations(motionProps, {
          kind: "table",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: true,
          filled: cells.some(
            (cell) =>
              cell.fill !== undefined || rows[cell.row]?.fill !== undefined,
          ),
          box,
          rotationDegrees: 0,
          hasBackground: cells.some(
            (cell) =>
              cell.fill !== undefined || rows[cell.row]?.fill !== undefined,
          ),
        });
        const resolvedModel: ResolvedTableModel = {
          x,
          y,
          widths,
          heights,
          rows,
          columns,
          resolvedCells,
          resolvedAnchors,
          owners,
          name,
          fontUses,
          ...(placeholder === undefined
            ? {}
            : { placeholder: placeholder.nativeRef }),
        };
        const mechanism =
          placeholder === undefined
            ? tableMechanism
            : tablePlaceholderMechanism;
        return {
          observation: {
            id: resolveContext.id,
            kind: "table",
            name,
            box,
            table: observation,
          },
          parts: subParts,
          extensions: [],
          ...(annotations.length === 0 ? {} : { annotations }),
          resourceClaims: claims,
          ...(fontUsages.length === 0 ? {} : { fontUsages }),
          ...(warnings.length === 0 ? {} : { warnings }),

          emit(emitContext: EmitContext): EmitOutput {
            return mechanism.emit(resolvedModel, emitContext);
          },

          paint(paintContext: PaintContext): PaintOutput {
            return mechanism.paint(resolvedModel, paintContext);
          },
        };
      },
    };
  },
} satisfies ObjectCapabilityModule;
