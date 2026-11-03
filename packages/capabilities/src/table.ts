import {
  CapabilityError,
  pxToEmu,
  resolvedColorXml,
  xmlEscape,
  type CapabilityInput,
  type CapabilityModule,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type LayoutStyle,
  type LinkTarget,
  type NativeTableObservation,
  type ColorValue,
  type ResolvedColor,
  type ThemeContext,
  type TextStyle,
} from "@react-pptx/core";
import { createElement, type ReactNode } from "react";

import {
  fontPresentationContributions,
  type FontEmbeddingUse,
} from "./font-embedding.ts";
import { nativeLinkXml } from "./identity.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import {
  textExtent,
  writeTextBody,
  type FontReference,
  type FontSize,
} from "./text-service.ts";

/** 列宽以 px 表示；Nfr 按剩余宽度分配，auto 使用共享字体度量。 */
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
};

export type TableStyle = TableTextStyle & LayoutStyle;

export type TableRowStyle = TableTextStyle & {
  /** 未指定时按本行未合并单元格的文字度量确定行高。 */
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

export type TableProps = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly x?: number;
    readonly y?: number;
    /** 为 fr 列提供可分配宽度；省略时所有列按内容自然宽度确定。 */
    readonly width?: number;
    readonly columns: readonly TableColumn[];
    readonly style?: TableStyle;
    readonly children?: ReactNode;
  };

type TablePartProps = {
  readonly testId?: string;
  readonly part?: string;
  readonly children?: ReactNode;
};

export type RowProps = TablePartProps & {
  readonly style?: TableRowStyle;
};

export type CellProps = TablePartProps & {
  readonly colSpan?: number;
  readonly rowSpan?: number;
  readonly style?: TableCellStyle;
};

type Track =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "fr"; readonly value: number }
  | { readonly kind: "auto" };

type Edges = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};

type ParsedCell = {
  readonly node: Extract<
    CapabilityInput["children"][number],
    { props: unknown }
  >;
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly colSpan: number;
  readonly path: string;
  readonly textStyle: Record<string, unknown>;
  readonly verticalAlign: TableVerticalAlign;
  readonly padding: Edges;
  readonly fill?: ResolvedColor;
  readonly border?: {
    readonly color: ResolvedColor;
    readonly width: number;
    readonly dash: "solid" | "dash" | "dot";
  };
};

type ParsedRow = {
  readonly path: string;
  readonly height?: number;
  readonly textStyle: Record<string, unknown>;
  readonly verticalAlign?: TableVerticalAlign;
  readonly fill?: ResolvedColor;
};

const textStyleKeys = [
  "fontFamily",
  "fontSize",
  "fontWeight",
  "fontStyle",
  "color",
  "lineHeight",
  "verticalAlign",
  "align",
  "spacingAfter",
  "letterSpacing",
  "textTransform",
  "textDecoration",
  "textShadow",
] as const;
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
const layoutStyleKeys = [
  "display",
  "position",
  "inset",
  "left",
  "top",
  "right",
  "bottom",
  "width",
  "height",
  "minWidth",
  "minHeight",
  "maxWidth",
  "maxHeight",
  "margin",
  "marginTop",
  "marginRight",
  "marginBottom",
  "marginLeft",
  "padding",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
  "gap",
  "rowGap",
  "columnGap",
  "flexDirection",
  "flexWrap",
  "justifyContent",
  "alignItems",
  "alignSelf",
  "alignContent",
  "flex",
  "flexGrow",
  "flexShrink",
  "flexBasis",
  "aspectRatio",
  "gridTemplateColumns",
  "gridTemplateRows",
  "gridColumn",
  "gridRow",
  "justifyItems",
  "justifySelf",
  "zIndex",
] as const;
const tableParameters = [
  "testId",
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
  "__motionContext",
] as const;
const tablePartParameters = ["testId", "part", "style"] as const;
const cellParameters = [...tablePartParameters, "colSpan", "rowSpan"] as const;
const verticalAlignments = new Set(["top", "middle", "bottom"]);
const maximumExtentEmu = 2_147_483_647;
const tableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";

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
  if (!isRecord(value)) fail("invalid-parameter", `${label} 必须是对象`, path);
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
    fail("unknown-parameter", `${label} 含未实现参数 ${unknown}`, path);
  }
}

function optionalName(
  value: unknown,
  label: "testId" | "name" | "part",
  path?: string,
): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    fail("invalid-parameter", `${label} 必须是非空字符串`, path);
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
      `${label} 必须是 -1000000 到 1000000 间的有限像素数值`,
    );
  }
  return value;
}

function positiveSize(value: unknown, label: string, path?: string): number {
  if (typeof value !== "number") {
    fail("invalid-table-geometry", `${label} 必须是有限正像素数值`, path);
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
      fail(error.code, `${label}：${error.message}`, path);
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
      "verticalAlign 只能是 top、middle 或 bottom",
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
    fail("invalid-table-column", `${label} 必须大于 0`);
  }
  return value;
}

function parseTracks(value: unknown): Track[] {
  if (!Array.isArray(value) || value.length === 0) {
    fail("invalid-table-column", "Table.columns 必须是非空数组");
  }
  return value.map((column, index) => {
    if (typeof column === "number") {
      return {
        kind: "px",
        value: positiveTrack(column, `第 ${index + 1} 列宽`),
      };
    }
    if (column === "auto") return { kind: "auto" };
    if (typeof column === "string") {
      const fraction = /^(\d+(?:\.\d+)?)fr$/u.exec(column);
      if (fraction !== null) {
        return {
          kind: "fr",
          value: positiveTrack(Number(fraction[1]), `第 ${index + 1} 列比例`),
        };
      }
    }
    fail(
      "invalid-table-column",
      `第 ${index + 1} 列宽无效；请使用正 px 数值、正 "Nfr" 或 "auto"`,
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
    fail("invalid-table-merge", `${label} 必须是正整数`, path);
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
    fail("invalid-table-padding", "padding 必须是数字、二元组或四元组", path);
  }
  const number = (entry: unknown): number => {
    if (typeof entry !== "number" || !Number.isFinite(entry) || entry < 0) {
      fail("invalid-table-padding", "单元格内边距必须是有限非负像素数值", path);
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
      fail("invalid-table-padding", `${name} 必须是有限非负像素数值`, path);
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
    fail("invalid-table-border", "borderWidth 必须是有限非负像素数值", path);
  }
  if (
    lineStyle !== undefined &&
    lineStyle !== "solid" &&
    lineStyle !== "dashed" &&
    lineStyle !== "dotted"
  ) {
    fail(
      "invalid-table-border",
      "borderStyle 必须是 solid、dashed 或 dotted",
      path,
    );
  }
  const width =
    (widthValue as number | undefined) ?? (borderColor === undefined ? 0 : 1);
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
      `${label} 换算后超出表格可写入的尺寸范围`,
      path,
    );
  }
  return emu;
}

function fillXml(value: ResolvedColor | undefined): string {
  return value === undefined
    ? "<a:noFill/>"
    : `<a:solidFill>${resolvedColorXml(value)}</a:solidFill>`;
}

function lineXml(
  side: "L" | "R" | "T" | "B",
  border: ParsedCell["border"],
): string {
  if (border === undefined) return `<a:ln${side}><a:noFill/></a:ln${side}>`;
  const width = Math.max(1, pxToEmu(border.width));
  const dash =
    border.dash === "solid" ? "" : `<a:prstDash val="${border.dash}"/>`;
  return `<a:ln${side} w="${width}"><a:solidFill>${resolvedColorXml(border.color)}</a:solidFill>${dash}</a:ln${side}>`;
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
): ReturnType<typeof writeTextBody> {
  try {
    return writeTextBody(
      cell.node.children,
      cell.textStyle,
      theme,
      cell.verticalAlign,
      {
        width,
        ...(height === undefined ? {} : { height }),
        ...(contentSized ? { contentSized: true } : {}),
      },
      "table-cell",
      fields,
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
  const identity = props["testId"] ?? props["part"];
  return `${type}[${pathLabel(identity, String(index + 1))}]`;
}

function prepareRows(
  input: CapabilityInput,
  columns: readonly Track[],
  tableTextStyle: Record<string, unknown>,
  tableVerticalAlign: TableVerticalAlign,
  theme: ThemeContext,
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
        fail("invalid-table-structure", "Table 中的文字必须放在 Cell 内");
      }
      continue;
    }
    if (child.type !== "rp:row") {
      fail("invalid-table-structure", "Table 只能直接包含 Row");
    }
    const rowIndex = rows.length;
    const rowPath = identityPath("row", child.props, rowIndex);
    onlyKeys(child.props, tablePartParameters, "Row", rowPath);
    optionalName(child.props["testId"], "testId", rowPath);
    optionalName(child.props["part"], "part", rowPath);
    const rowStyle = trackStyle(child as CapabilityInput, "row", rowPath);
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
            "Row 中的文字必须放在 Cell 内",
            rowPath,
          );
        }
        continue;
      }
      if (cellChild.type !== "rp:cell") {
        fail("invalid-table-structure", "Row 只能直接包含 Cell", rowPath);
      }
      const cellPath = `${rowPath}/${identityPath("cell", cellChild.props, cellIndex)}`;
      cellIndex += 1;
      onlyKeys(cellChild.props, cellParameters, "Cell", cellPath);
      optionalName(cellChild.props["testId"], "testId", cellPath);
      optionalName(cellChild.props["part"], "part", cellPath);
      while (owners.has(`${rowIndex}:${column}`)) column += 1;
      const colSpan = span(cellChild.props["colSpan"], "colSpan", cellPath);
      const rowSpan = span(cellChild.props["rowSpan"], "rowSpan", cellPath);
      if (column + colSpan > columns.length) {
        fail(
          "invalid-table-merge",
          `第 ${rowIndex + 1} 行的单元格超出 ${columns.length} 列范围`,
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
          `rowSpan ${rowSpan} 越过表格最后一行`,
          cellPath,
        );
      }
      for (let row = rowIndex; row < rowIndex + rowSpan; row += 1) {
        for (let col = column; col < column + colSpan; col += 1) {
          if (owners.has(`${row}:${col}`)) {
            fail(
              "conflicting-table-merge",
              "单元格合并区域与已有单元格重叠",
              cellPath,
            );
          }
        }
      }
      const cellStyle = trackStyle(
        cellChild as CapabilityInput,
        "cell",
        cellPath,
      );
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
      const cell: ParsedCell = {
        node: cellChild as ParsedCell["node"],
        row: rowIndex,
        column,
        rowSpan,
        colSpan,
        path: cellPath,
        textStyle: mergeTextStyle(row.textStyle, cellStyle.textStyle),
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

  if (rows.length === 0) fail("invalid-table-structure", "Table 至少需要一行");
  for (let row = 0; row < rows.length; row += 1) {
    const missing = columns.findIndex(
      (_, column) => !owners.has(`${row}:${column}`),
    );
    if (missing >= 0) {
      fail(
        "invalid-table-structure",
        `第 ${row + 1} 行第 ${missing + 1} 列没有 Cell 或合并单元格覆盖`,
        rows[row]?.path,
      );
    }
  }
  return { rows, cells };
}

function mergeCellProperties(
  padding: Edges,
  vertical: TableVerticalAlign,
  border: ParsedCell["border"],
  fill: ResolvedColor | undefined,
): string {
  const lines = (["L", "R", "T", "B"] as const)
    .map((side) => lineXml(side, border))
    .join("");
  return `<a:tcPr marL="${pxToEmu(padding.left)}" marR="${pxToEmu(padding.right)}" marT="${pxToEmu(padding.top)}" marB="${pxToEmu(padding.bottom)}" anchor="${vertical === "middle" ? "ctr" : vertical === "bottom" ? "b" : "t"}">${lines}${fillXml(fill)}</a:tcPr>`;
}

function emptyCell(): string {
  return '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/>';
}

function frameXml(
  id: number,
  escapedName: string,
  x: number,
  y: number,
  widths: readonly number[],
  heights: readonly number[],
  rows: readonly string[],
  linkXml: string,
): string {
  const columnXml = widths
    .map((width) => `<a:gridCol w="${pxToEmu(width)}"/>`)
    .join("");
  const rowXml = rows
    .map((cells, row) => `<a:tr h="${pxToEmu(heights[row]!)}">${cells}</a:tr>`)
    .join("");
  const width = widths.reduce((sum, value) => sum + pxToEmu(value), 0);
  const height = heights.reduce((sum, value) => sum + pxToEmu(value), 0);
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pxToEmu(x)}" y="${pxToEmu(y)}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm><a:graphic><a:graphicData uri="${tableUri}"><a:tbl><a:tblPr/><a:tblGrid>${columnXml}</a:tblGrid>${rowXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

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

/** Shared table-owned track and text sizing used by both LayoutPass and XML emission. */
function tableLayoutModel(
  input: CapabilityInput,
  requestedWidth: number | undefined,
  theme: ThemeContext,
  slideNumber: number,
): TableLayoutModel {
  const columns = parseTracks(input.props["columns"]);
  const tableStyle = trackStyle(input, "table");
  const tableVerticalAlign =
    verticalAlign(tableStyle.style["verticalAlign"]) ?? "middle";
  const { rows, cells } = prepareRows(
    input,
    columns,
    tableStyle.textStyle,
    tableVerticalAlign,
    theme,
  );
  const warnings: { code: string; message: string; path?: string }[] = [];
  const claims: { kind: string; key: string; identity: string }[] = [];
  const warningKeys = new Set<string>();
  const addTextResult = (
    result: ReturnType<typeof writeTextBody>,
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
      const measured = textBody(cell, theme, "max-content", undefined, false, {
        slideNumber,
      });
      addTextResult(measured, cell.path);
      if (measured.measurement === undefined && measured.text.trim() !== "") {
        fail(
          "font-measurement-unavailable",
          "auto 列需要可用的字体度量；请提供度量字体或改用明确列宽",
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
        message: `列的最小宽度合计 ${actualWidth}px 超出声明宽度 ${requestedWidth}px；构建保留列宽`,
      });
    }
  }
  const widthsEmu = widths.map((width, index) =>
    pxExtent(width, `第 ${index + 1} 列宽`),
  );
  if (widthsEmu.reduce((sum, value) => sum + value, 0) > maximumExtentEmu) {
    fail("invalid-table-geometry", "表格列宽合计超出可写入范围");
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
    const result = textBody(cell, theme, innerWidth, undefined, false, {
      slideNumber,
    });
    addTextResult(result, cell.path);
    if (result.measurement === undefined && result.text.trim() !== "") {
      fail(
        "font-measurement-unavailable",
        "自动行高需要可用的字体度量；请提供度量字体或为 Row 指定 height",
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
    pxExtent(height, `第 ${index + 1} 行高`, rows[index]?.path),
  );
  if (heightsEmu.reduce((sum, value) => sum + value, 0) > maximumExtentEmu) {
    fail("invalid-table-geometry", "表格行高合计超出可写入范围");
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

/** 表格模块独立校验行列结构，以原生 graphicFrame 和 a:tbl 写入单一对象。 */
export const tableCapability: CapabilityModule = {
  declaration: "rp:table",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement {
    onlyKeys(input.props, tableParameters, "Table");
    const style = trackStyle(input, "table").style;
    const declaredWidth = style["width"] ?? input.props["width"];
    const declaredSize =
      typeof declaredWidth === "number"
        ? positiveSize(declaredWidth, "Table.width")
        : undefined;
    const natural = tableLayoutModel(
      input,
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
            input,
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
  compile(input, context) {
    onlyKeys(input.props, tableParameters, "Table");
    const testId = optionalName(input.props["testId"], "testId");
    const requestedName = optionalName(input.props["name"], "name");
    const style = trackStyle(input, "table").style;
    const declaredWidth = style["width"] ?? input.props["width"];
    const requestedWidth =
      context.layoutBox?.width ??
      (typeof declaredWidth === "number"
        ? positiveSize(declaredWidth, "Table.width")
        : undefined);
    const model = tableLayoutModel(
      input,
      requestedWidth,
      context.theme,
      context.slideNumber,
    );
    const { columns, rows, cells, widths } = model;
    const heights = [...model.heights];
    const warnings = [...model.warnings];
    const claims = [...model.claims];
    const fontUses: FontEmbeddingUse[] = [];
    const warningKeys = new Set(model.warningKeys);
    const addTextResult = (
      result: ReturnType<typeof writeTextBody>,
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
            message: `固定行高合计 ${naturalHeight}px 小于布局高度 ${targetHeight}px；构建保留行高`,
          });
        }
      } else if (targetHeight < naturalHeight - 0.01) {
        warnings.push({
          code: "table-height-overflow",
          message: `行高合计 ${naturalHeight}px 超出布局高度 ${targetHeight}px；构建保留行高`,
        });
      }
    }
    const x = coordinate(context.layoutBox?.x ?? input.props["x"], "x");
    const y = coordinate(context.layoutBox?.y ?? input.props["y"], "y");
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
            fail("invalid-table-structure", "表格行存在未声明的单元格");
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
            `<a:tc${hMerge}${vMerge}${spanAttributes.length === 0 ? "" : ` ${spanAttributes.join(" ")}`}>${emptyCell()}</a:tc>`,
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
        const properties = mergeCellProperties(
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
        warning.message = `${warning.message}（表格单元格保留原始内容和声明尺寸）`;
      }
    }
    const name =
      context.morphName ?? requestedName ?? testId ?? `Table ${context.id - 1}`;
    let escapedName: string;
    try {
      escapedName = xmlEscape(name);
    } catch {
      fail("invalid-xml-character", "Table.name 含 XML 1.0 不允许的字符");
    }
    const width = widths.reduce((sum, value) => sum + value, 0);
    const height = heights.reduce((sum, value) => sum + value, 0);
    const box = { x, y, width, height };
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
    };
    const annotations = motionAnnotations(input.props, {
      kind: "table",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: true,
      filled: cells.some(
        (cell) => cell.fill !== undefined || rows[cell.row]?.fill !== undefined,
      ),
    });
    return {
      xml: frameXml(
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

/** 在 React root 中声明由表格模块处理的可编辑原生表格。 */
export function Table(props: TableProps): ReactNode {
  return motionElement("rp:table", props);
}

/** 声明原生表格的一行；未指定高度时按该行内容的真实度量计算。 */
export function Row(props: RowProps): ReactNode {
  return createElement("rp:row", props, props.children);
}

/** 声明表格单元格正文与横向、纵向合并范围。 */
export function Cell(props: CellProps): ReactNode {
  return createElement("rp:cell", props, props.children);
}
