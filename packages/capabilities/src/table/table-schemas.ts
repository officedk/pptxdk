import {
  CapabilityError,
  backgroundValueSchema,
  borderValueSchema,
  layoutStyleShape,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  textStyleShape,
  validateTextStyle,
  type BackgroundValue,
  type BorderValue,
  type ColorValue,
  type LayoutStyle,
  type LinkTarget,
  type NativePlaceholderRef,
  type NonNegativeLayoutSpacing,
  type ResolvedColor,
  type TextStyle,
  type ThemeContext,
} from "@pptxdk/core/capability";
import { pxToEmu } from "@pptxdk/emit";
import {
  ooxmlBoolSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
} from "@pptxdk/import";
import { z } from "zod";

import {
  motionPropsShape,
  type MotionProps,
} from "../motion/motion-schemas.ts";
import { optionalNameSchema } from "../shared/identity.ts";
import { morphKeySchema, type MorphKey } from "../shared/morph-api.ts";
import {
  parseTextInsets,
  textCoordinateSchema,
  textPaddingValueSchema,
  textPositiveSizeSchema,
} from "../shared/text-schemas.ts";
import type {
  TableCellBorder,
  TableCellBorders,
  TableCellPadding,
} from "./table-xml.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

const frColumnPattern = /^\d+(?:\.\d+)?fr$/u;
const maximumExtentEmu = 2_147_483_647;

export type TableTrack =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "fr"; readonly value: number }
  | { readonly kind: "auto" };

export type TableVerticalAlign = "top" | "middle" | "bottom";
const rawTableVerticalAlignSchema = z.enum(["top", "middle", "bottom"]);
export const tableVerticalAlignSchema: z.ZodType<TableVerticalAlign> =
  rawTableVerticalAlignSchema satisfies z.ZodType<
    AssertExact<
      TableVerticalAlign,
      z.infer<typeof rawTableVerticalAlignSchema>
    > extends true
      ? TableVerticalAlign
      : never
  >;

const tableColumnPxSchema = z.number().finite().positive();
const tableColumnFrSchema = z
  .string()
  .regex(frColumnPattern)
  .refine(
    (value) => Number(value.slice(0, -2)) > 0,
  ) as unknown as z.ZodType<`${number}fr`>;

export type TableColumn = number | `${number}fr` | "auto";
const rawTableColumnSchema = z.union([
  tableColumnPxSchema,
  tableColumnFrSchema,
  z.literal("auto"),
]);
/** Column width in px; Nfr allocates from remaining width, auto uses shared font measurement. */
export const tableColumnSchema: z.ZodType<TableColumn> =
  rawTableColumnSchema satisfies z.ZodType<
    AssertExact<TableColumn, z.infer<typeof rawTableColumnSchema>> extends true
      ? TableColumn
      : never
  >;

export type TableSpacing = NonNegativeLayoutSpacing;
export const tableSpacingSchema: z.ZodType<TableSpacing> =
  textPaddingValueSchema;

export type TableTextStyle = Pick<
  TextStyle,
  | "fontFamily"
  | "fontSize"
  | "fontWeight"
  | "fontStyle"
  | "color"
  | "lineHeight"
  | "align"
  | "verticalAlign"
  | "letterSpacing"
  | "textTransform"
  | "textDecoration"
  | "textShadow"
  | "lang"
>;

const tableTextStyleObjectSchema = z.strictObject({
  fontFamily: textStyleShape.fontFamily,
  fontSize: textStyleShape.fontSize,
  fontWeight: textStyleShape.fontWeight,
  fontStyle: textStyleShape.fontStyle,
  color: textStyleShape.color,
  lineHeight: textStyleShape.lineHeight,
  align: textStyleShape.align,
  verticalAlign: textStyleShape.verticalAlign,
  letterSpacing: textStyleShape.letterSpacing,
  textTransform: textStyleShape.textTransform,
  textDecoration: textStyleShape.textDecoration,
  textShadow: textStyleShape.textShadow,
  lang: textStyleShape.lang,
});

const rawTableTextStyleSchema = tableTextStyleObjectSchema.readonly();
export const tableTextStyleSchema: z.ZodType<TableTextStyle> =
  rawTableTextStyleSchema satisfies z.ZodType<
    AssertExact<
      TableTextStyle,
      z.infer<typeof rawTableTextStyleSchema>
    > extends true
      ? TableTextStyle
      : never
  >;

export type TableStyle = TableTextStyle &
  LayoutStyle & {
    readonly background?: BackgroundValue | undefined;
    readonly border?: BorderValue | undefined;
    readonly borderTop?: BorderValue | undefined;
    readonly borderRight?: BorderValue | undefined;
    readonly borderBottom?: BorderValue | undefined;
    readonly borderLeft?: BorderValue | undefined;
  };

const tableStyleObjectSchema = z.strictObject({
  ...tableTextStyleObjectSchema.shape,
  ...layoutStyleShape,
  background: backgroundValueSchema.optional(),
  border: borderValueSchema.optional(),
  borderTop: borderValueSchema.optional(),
  borderRight: borderValueSchema.optional(),
  borderBottom: borderValueSchema.optional(),
  borderLeft: borderValueSchema.optional(),
});
const rawTableStyleSchema = tableStyleObjectSchema.readonly();
export const tableStyleSchema: z.ZodType<TableStyle> =
  rawTableStyleSchema satisfies z.ZodType<
    AssertExact<TableStyle, z.infer<typeof rawTableStyleSchema>> extends true
      ? TableStyle
      : never
  >;

export type TableRowStyle = TableTextStyle & {
  /** When unspecified, row height is determined from text measurement of unmerged cells in this row. */
  readonly height?: number | undefined;
  readonly background?: BackgroundValue | undefined;
  readonly border?: BorderValue | undefined;
  readonly borderTop?: BorderValue | undefined;
  readonly borderRight?: BorderValue | undefined;
  readonly borderBottom?: BorderValue | undefined;
  readonly borderLeft?: BorderValue | undefined;
};

const tableRowOwnStyleObjectSchema = z.strictObject({
  /** When unspecified, row height is determined from text measurement of unmerged cells in this row. */
  height: textPositiveSizeSchema.optional(),
  background: backgroundValueSchema.optional(),
  border: borderValueSchema.optional(),
  borderTop: borderValueSchema.optional(),
  borderRight: borderValueSchema.optional(),
  borderBottom: borderValueSchema.optional(),
  borderLeft: borderValueSchema.optional(),
});

const rawTableRowStyleSchema = z
  .strictObject({
    ...tableTextStyleObjectSchema.shape,
    ...tableRowOwnStyleObjectSchema.shape,
  })
  .readonly();
export const tableRowStyleSchema: z.ZodType<TableRowStyle> =
  rawTableRowStyleSchema satisfies z.ZodType<
    AssertExact<
      TableRowStyle,
      z.infer<typeof rawTableRowStyleSchema>
    > extends true
      ? TableRowStyle
      : never
  >;

export type TableCellStyle = TableTextStyle & {
  readonly background?: BackgroundValue | undefined;
  readonly border?: BorderValue | undefined;
  readonly borderTop?: BorderValue | undefined;
  readonly borderRight?: BorderValue | undefined;
  readonly borderBottom?: BorderValue | undefined;
  readonly borderLeft?: BorderValue | undefined;
  readonly padding?: NonNegativeLayoutSpacing | undefined;
  readonly paddingTop?: number | undefined;
  readonly paddingRight?: number | undefined;
  readonly paddingBottom?: number | undefined;
  readonly paddingLeft?: number | undefined;
};

const tableCellOwnStyleObjectSchema = z.strictObject({
  background: backgroundValueSchema.optional(),
  border: borderValueSchema.optional(),
  borderTop: borderValueSchema.optional(),
  borderRight: borderValueSchema.optional(),
  borderBottom: borderValueSchema.optional(),
  borderLeft: borderValueSchema.optional(),
  padding: textPaddingValueSchema.optional(),
  paddingTop: z.number().finite().nonnegative().optional(),
  paddingRight: z.number().finite().nonnegative().optional(),
  paddingBottom: z.number().finite().nonnegative().optional(),
  paddingLeft: z.number().finite().nonnegative().optional(),
});

const tableCellStyleObjectSchema = z.strictObject({
  ...tableTextStyleObjectSchema.shape,
  ...tableCellOwnStyleObjectSchema.shape,
});
const rawTableCellStyleSchema = tableCellStyleObjectSchema.readonly();
export const tableCellStyleSchema: z.ZodType<TableCellStyle> =
  rawTableCellStyleSchema satisfies z.ZodType<
    AssertExact<
      TableCellStyle,
      z.infer<typeof rawTableCellStyleSchema>
    > extends true
      ? TableCellStyle
      : never
  >;

export const tableMergeSpanSchema = z.number().int().min(1);

export type TableDeclarationProps = MotionProps & {
  readonly id?: string | undefined;
  readonly name?: string | undefined;
  readonly placeholder?: string | undefined;
  readonly nativePlaceholder?: NativePlaceholderRef | undefined;
  readonly morphKey?: MorphKey | undefined;
  readonly link?: LinkTarget | undefined;
  readonly columns: readonly TableColumn[];
  readonly style?: TableStyle | undefined;
};

const tableDeclarationPropsObjectSchema = z.strictObject({
  id: optionalNameSchema.optional(),
  name: optionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: morphKeySchema.optional(),
  link: linkTargetSchema.optional(),
  columns: z.array(rawTableColumnSchema).nonempty().readonly(),
  style: rawTableStyleSchema.optional(),
  enter: motionPropsShape.enter,
  emphasis: motionPropsShape.emphasis,
  motion: motionPropsShape.motion,
});

const rawTableDeclarationPropsSchema =
  tableDeclarationPropsObjectSchema.readonly();
export const tableDeclarationPropsSchema: z.ZodType<TableDeclarationProps> =
  rawTableDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      TableDeclarationProps,
      z.infer<typeof rawTableDeclarationPropsSchema>
    > extends true
      ? TableDeclarationProps
      : never
  >;

export type RowDeclarationProps = {
  readonly id?: string | undefined;
  readonly style?: TableRowStyle | undefined;
};

const rowDeclarationPropsObjectSchema = z.strictObject({
  id: optionalNameSchema.optional(),
  style: rawTableRowStyleSchema.optional(),
});
const rawRowDeclarationPropsSchema = rowDeclarationPropsObjectSchema.readonly();
export const rowDeclarationPropsSchema: z.ZodType<RowDeclarationProps> =
  rawRowDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      RowDeclarationProps,
      z.infer<typeof rawRowDeclarationPropsSchema>
    > extends true
      ? RowDeclarationProps
      : never
  >;

export type CellDeclarationProps = {
  readonly id?: string | undefined;
  readonly colSpan?: number | undefined;
  readonly rowSpan?: number | undefined;
  readonly style?: TableCellStyle | undefined;
};

const cellDeclarationPropsObjectSchema = z.strictObject({
  id: optionalNameSchema.optional(),
  colSpan: tableMergeSpanSchema.optional(),
  rowSpan: tableMergeSpanSchema.optional(),
  style: rawTableCellStyleSchema.optional(),
});
const rawCellDeclarationPropsSchema =
  cellDeclarationPropsObjectSchema.readonly();
export const cellDeclarationPropsSchema: z.ZodType<CellDeclarationProps> =
  rawCellDeclarationPropsSchema satisfies z.ZodType<
    AssertExact<
      CellDeclarationProps,
      z.infer<typeof rawCellDeclarationPropsSchema>
    > extends true
      ? CellDeclarationProps
      : never
  >;

export const tableDeclarationAllowedPropKeys: readonly string[] = Object.keys(
  tableDeclarationPropsObjectSchema.shape,
);
export const tableAllowedParameters: readonly string[] = [
  ...tableDeclarationAllowedPropKeys,
  "children",
];
export const rowAllowedParameters: readonly string[] = Object.keys(
  rowDeclarationPropsObjectSchema.shape,
);
export const cellAllowedParameters: readonly string[] = Object.keys(
  cellDeclarationPropsObjectSchema.shape,
);
export const tableTextStyleKeys: readonly string[] = Object.keys(
  tableTextStyleObjectSchema.shape,
);
export const tableStyleKeys: readonly string[] = Object.keys(
  tableStyleObjectSchema.shape,
);
export const rowStyleKeys: readonly string[] = Object.keys(
  tableRowOwnStyleObjectSchema.shape,
);
export const cellStyleKeys: readonly string[] = Object.keys(
  tableCellOwnStyleObjectSchema.shape,
);
export const tableCellAllStyleKeys: readonly string[] = Object.keys(
  tableCellStyleObjectSchema.shape,
);

/** OOXML `<p:cNvPr>` attribute fragment schema inside `<p:graphicFrame>`. */
export const nonVisualGraphicFrameIdSchema = z.object({
  id: ooxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

/** OOXML `<a:graphicFrameLocks>` attribute fragment schema. */
export const graphicFrameLocksXmlSchema = z.object({
  noGrp: ooxmlBoolSchema.optional(),
  noDrilldown: ooxmlBoolSchema.optional(),
  noSelect: ooxmlBoolSchema.optional(),
  noChangeAspect: ooxmlBoolSchema.optional(),
  noMove: ooxmlBoolSchema.optional(),
  noResize: ooxmlBoolSchema.optional(),
});

/** OOXML `<a:tblPr>` attribute fragment schema. */
export const tblPrXmlSchema = z.object({
  firstRow: ooxmlBoolSchema.optional(),
  firstCol: ooxmlBoolSchema.optional(),
  lastRow: ooxmlBoolSchema.optional(),
  lastCol: ooxmlBoolSchema.optional(),
  bandRow: ooxmlBoolSchema.optional(),
  bandCol: ooxmlBoolSchema.optional(),
  rtl: ooxmlBoolSchema.optional(),
});

/** OOXML `<a:gridCol w="..."/>` fragment schema. */
export const tblGridColXmlSchema = z.object({
  w: ooxmlPositiveIntSchema,
});

/** OOXML `<a:tr h="..."/>` fragment schema. */
export const tblRowXmlSchema = z.object({
  h: ooxmlPositiveIntSchema,
});

/** OOXML `<a:tc gridSpan="..." rowSpan="..." hMerge="..." vMerge="..."/>` fragment schema. */
export const tblCellXmlSchema = z.object({
  gridSpan: ooxmlPositiveIntSchema.optional(),
  rowSpan: ooxmlPositiveIntSchema.optional(),
  hMerge: ooxmlBoolSchema.optional(),
  vMerge: ooxmlBoolSchema.optional(),
});

/** OOXML `<a:tcPr .../>` attribute fragment schema. */
export const tblCellPrXmlSchema = z.object({
  marL: ooxmlNonNegativeIntSchema.optional(),
  marR: ooxmlNonNegativeIntSchema.optional(),
  marT: ooxmlNonNegativeIntSchema.optional(),
  marB: ooxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
  vert: z.string().optional(),
  horzOverflow: z.string().optional(),
  anchorCtr: ooxmlBoolSchema.optional(),
});

/** Validates a plain object parameter map against allowed parameter keys. */
export function assertAllowedTableKeys(
  value: Readonly<Record<string, unknown>>,
  allowed: readonly string[],
  label: string,
  path?: string,
): void {
  const unknown = Object.keys(value).find((key) => !allowed.includes(key));
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `${label} contains unsupported parameter ${unknown}`,
      path,
    );
  }
}

/** Validates an optional non-empty author ID on Row or Cell. */
export function parseOptionalTableId(
  value: unknown,
  path?: string,
): string | undefined {
  if (value === undefined) return undefined;
  const parsed = optionalNameSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-parameter",
      "id must be a non-empty string",
      path,
    );
  }
  return parsed.data;
}

/** Validates table coordinate (`x` or `y`) via shared coordinate schema, defaulting omitted to `0`. */
export function parseTableCoordinate(value: unknown, label: "x" | "y"): number {
  if (value === undefined) return 0;
  const parsed = textCoordinateSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-table-geometry",
      `${label} must be a finite pixel number between -1000000 and 1000000`,
    );
  }
  return parsed.data;
}

/** Validates positive pixel extent (`width` or `Row.style.height`) via shared positive size schema. */
export function parseTablePositiveSize(
  value: unknown,
  label: string,
  path?: string,
): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    if (value === 0 || pxToEmu(value) === 0) {
      throw new CapabilityError(
        "invalid-table-geometry",
        `Computed ${label} is 0px; ${label} must be greater than 0px`,
        path,
      );
    }
    if (value < 0 || pxToEmu(value) < 1) {
      throw new CapabilityError(
        "invalid-table-geometry",
        `Computed ${label} is ${value}px; ${label} must be greater than 0px`,
        path,
      );
    }
  } else {
    throw new CapabilityError(
      "invalid-table-geometry",
      `${label} must be a finite positive pixel number`,
      path,
    );
  }
  const parsed = textPositiveSizeSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-table-geometry",
      `${label} exceeds writable table dimension range after conversion`,
      path,
    );
  }
  return parsed.data;
}

/** Validates EMU extent bounds for resolved table tracks and borders. */
export function assertTableEmuExtent(
  value: number,
  label: string,
  path?: string,
): number {
  const emu = pxToEmu(value);
  if (Number.isFinite(value) && (value === 0 || emu === 0)) {
    throw new CapabilityError(
      "invalid-table-geometry",
      `Computed ${label} is 0px; ${label} must be greater than 0px`,
      path,
    );
  }
  if (Number.isFinite(value) && (value < 0 || emu < 1)) {
    throw new CapabilityError(
      "invalid-table-geometry",
      `Computed ${label} is ${value}px; ${label} must be greater than 0px`,
      path,
    );
  }
  if (!Number.isSafeInteger(emu) || emu > maximumExtentEmu) {
    throw new CapabilityError(
      "invalid-table-geometry",
      `${label} exceeds writable table dimension range after conversion`,
      path,
    );
  }
  return emu;
}

/** Validates vertical alignment via Zod schema. */
export function parseTableVerticalAlign(
  value: unknown,
  path?: string,
): TableVerticalAlign | undefined {
  if (value === undefined) return undefined;
  const parsed = tableVerticalAlignSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-vertical-align",
      "verticalAlign must be top, middle, or bottom",
      path,
    );
  }
  return parsed.data;
}

/** Validates non-empty `Table.columns` tracks via Zod schema. */
export function parseTableColumns(value: unknown): TableTrack[] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new CapabilityError(
      "invalid-table-column",
      "Table.columns must be a non-empty array",
    );
  }
  return value.map((column, index) => {
    const parsed = tableColumnSchema.safeParse(column);
    if (!parsed.success) {
      if (
        (typeof column === "number" &&
          (!Number.isFinite(column) || column <= 0)) ||
        (typeof column === "string" &&
          frColumnPattern.test(column) &&
          Number(column.slice(0, -2)) <= 0)
      ) {
        const kindLabel = typeof column === "number" ? "width" : "fraction";
        throw new CapabilityError(
          "invalid-table-column",
          `column ${index + 1} ${kindLabel} must be greater than 0`,
        );
      }
      throw new CapabilityError(
        "invalid-table-column",
        `Invalid width for column ${index + 1}; use positive px number, positive "Nfr", or "auto"`,
      );
    }
    if (typeof parsed.data === "number") {
      return { kind: "px", value: parsed.data };
    }
    if (parsed.data === "auto") {
      return { kind: "auto" };
    }
    return { kind: "fr", value: Number(parsed.data.slice(0, -2)) };
  });
}

/** Validates `colSpan` or `rowSpan` via Zod schema. */
export function parseTableMergeSpan(
  value: unknown,
  label: "colSpan" | "rowSpan",
  path: string,
): number {
  if (value === undefined) return 1;
  const parsed = tableMergeSpanSchema.safeParse(value);
  if (!parsed.success) {
    throw new CapabilityError(
      "invalid-table-merge",
      `${label} must be a positive integer`,
      path,
    );
  }
  return parsed.data;
}

/** Validates cell padding declarations by reusing shared `parseTextInsets`. */
export function parseTableCellPadding(
  style: Readonly<Record<string, unknown>>,
  path: string,
): TableCellPadding {
  try {
    return parseTextInsets(style);
  } catch (error) {
    if (
      error instanceof CapabilityError &&
      error.code === "invalid-text-insets"
    ) {
      throw new CapabilityError("invalid-table-padding", error.message, path);
    }
    throw error;
  }
}

/** Validates a single unified `BorderValue` (`"none"` or `{ color, width?, style? }`) on Table, Row, or Cell. */
export function parseTableBorderValue(
  value: unknown,
  label: string,
  theme: ThemeContext,
  path?: string,
): TableCellBorder | undefined {
  if (value === "none") return undefined;
  const parsed = borderValueSchema.safeParse(value);
  if (!parsed.success || parsed.data === "none") {
    const issue = parsed.success ? undefined : parsed.error.issues[0];
    throw new CapabilityError(
      "invalid-table-border",
      issue?.message ?? `${label} must be "none" or { color, width?, style? }`,
      path,
    );
  }
  const width = parsed.data.width ?? 1;
  if (width === 0) return undefined;
  assertTableEmuExtent(width, `${label}.width`, path);
  let resolvedColor: ResolvedColor;
  try {
    resolvedColor = theme.resolveColor(parsed.data.color as ColorValue);
  } catch (error) {
    if (error instanceof CapabilityError) {
      throw new CapabilityError(
        error.code,
        `${label}.color: ${error.message}`,
        path,
      );
    }
    throw error;
  }
  return {
    color: resolvedColor,
    width,
    dash:
      parsed.data.style === "dashed"
        ? "dash"
        : parsed.data.style === "dotted"
          ? "dot"
          : "solid",
  };
}

/** Validates optional `align` on Table, Row, or Cell using shared core text style validation. */
export function assertTableTextAlignment(align: unknown, path?: string): void {
  if (align === undefined) return;
  try {
    validateTextStyle({ align });
  } catch (error) {
    if (error instanceof CapabilityError) {
      throw new CapabilityError(error.code, error.message, path);
    }
    throw error;
  }
}

const tableBorderSideEntries = [
  ["borderTop", "top"],
  ["borderRight", "right"],
  ["borderBottom", "bottom"],
  ["borderLeft", "left"],
] as const;

/**
 * Resolves 4-side borders for a Table, Row, or Cell layer by applying `border` shorthand
 * then individual `borderTop`/`borderRight`/`borderBottom`/`borderLeft` overrides over inherited borders.
 */
export function resolveTableLayerBorders(
  style: Readonly<Record<string, unknown>>,
  inherited: TableCellBorders,
  label: "Table" | "Row" | "Cell",
  theme: ThemeContext,
  path?: string,
): TableCellBorders {
  let top = inherited.top;
  let right = inherited.right;
  let bottom = inherited.bottom;
  let left = inherited.left;

  if (Object.hasOwn(style, "border") && style["border"] !== undefined) {
    const baseBorder = parseTableBorderValue(
      style["border"],
      `${label}.style.border`,
      theme,
      path,
    );
    top = baseBorder;
    right = baseBorder;
    bottom = baseBorder;
    left = baseBorder;
  }

  for (const [sideKey, sideName] of tableBorderSideEntries) {
    if (!Object.hasOwn(style, sideKey) || style[sideKey] === undefined) {
      continue;
    }
    const sideValue = parseTableBorderValue(
      style[sideKey],
      `${label}.style.${sideKey}`,
      theme,
      path,
    );
    if (sideName === "top") top = sideValue;
    else if (sideName === "right") right = sideValue;
    else if (sideName === "bottom") bottom = sideValue;
    else left = sideValue;
  }

  return {
    ...(top === undefined ? {} : { top }),
    ...(right === undefined ? {} : { right }),
    ...(bottom === undefined ? {} : { bottom }),
    ...(left === undefined ? {} : { left }),
  };
}

/** Resolves final 4-side borders for a Cell by cascading Table -> Row -> Cell.border -> Cell.border{Side}. */
export function resolveTableCellBorders(
  style: Readonly<Record<string, unknown>>,
  inheritedBorders: TableCellBorders,
  theme: ThemeContext,
  path: string,
): TableCellBorders {
  return resolveTableLayerBorders(style, inheritedBorders, "Cell", theme, path);
}
