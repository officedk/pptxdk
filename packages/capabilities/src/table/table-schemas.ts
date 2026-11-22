import {
  CapabilityError,
  colorValueSchema,
  isMorphKey,
  layoutStyleSchema,
  linkTargetSchema,
  nativePlaceholderRefSchema,
  textStyleSchema,
  type MorphKey,
  type ResolvedColor,
  type ThemeContext,
} from "@pptxdk/core/capability";
import { pxToEmu } from "@pptxdk/emit";
import {
  OoxmlBoolSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
} from "@pptxdk/import";
import { z } from "zod";

import { OptionalNameSchema } from "../shared/identity.ts";
import {
  parseTextInsets,
  textCoordinateSchema,
  textPaddingValueSchema,
  textPositiveSizeSchema,
} from "../shared/text-schemas.ts";
import type { TableCellBorder, TableCellPadding } from "./table-xml.ts";

const frColumnPattern = /^\d+(?:\.\d+)?fr$/u;
const maximumExtentEmu = 2_147_483_647;

export type TableTrack =
  | { readonly kind: "px"; readonly value: number }
  | { readonly kind: "fr"; readonly value: number }
  | { readonly kind: "auto" };

export const tableVerticalAlignSchema = z.enum(["top", "middle", "bottom"]);
export type TableVerticalAlign = z.infer<typeof tableVerticalAlignSchema>;

const tableColumnPxSchema = z.number().finite().positive();
const tableColumnFrSchema = z
  .string()
  .regex(frColumnPattern)
  .refine(
    (value) => Number(value.slice(0, -2)) > 0,
  ) as unknown as z.ZodType<`${number}fr`>;

/** Column width in px; Nfr allocates from remaining width, auto uses shared font measurement. */
export const tableColumnSchema = z.union([
  tableColumnPxSchema,
  tableColumnFrSchema,
  z.literal("auto"),
]);
export type TableColumn = z.infer<typeof tableColumnSchema>;

const nonNegativeTableInsetSchema = z
  .number()
  .finite()
  .nonnegative()
  .refine((value) => Number.isSafeInteger(pxToEmu(value)));

export const tableSpacingSchema = z.union([
  nonNegativeTableInsetSchema,
  z
    .tuple([nonNegativeTableInsetSchema, nonNegativeTableInsetSchema])
    .readonly(),
  z
    .tuple([
      nonNegativeTableInsetSchema,
      nonNegativeTableInsetSchema,
      nonNegativeTableInsetSchema,
      nonNegativeTableInsetSchema,
    ])
    .readonly(),
]);
export type TableSpacing = z.infer<typeof tableSpacingSchema>;

const coreTextStyleShape = textStyleSchema.unwrap().shape;

const tableTextStyleObjectSchema = z.strictObject({
  fontFamily: coreTextStyleShape.fontFamily,
  fontSize: coreTextStyleShape.fontSize,
  fontWeight: coreTextStyleShape.fontWeight,
  fontStyle: coreTextStyleShape.fontStyle,
  color: coreTextStyleShape.color,
  lineHeight: coreTextStyleShape.lineHeight,
  verticalAlign: coreTextStyleShape.verticalAlign,
  letterSpacing: coreTextStyleShape.letterSpacing,
  textTransform: coreTextStyleShape.textTransform,
  textDecoration: coreTextStyleShape.textDecoration,
  textShadow: coreTextStyleShape.textShadow,
  lang: coreTextStyleShape.lang,
});

export const tableTextStyleSchema = tableTextStyleObjectSchema.readonly();
export type TableTextStyle = z.infer<typeof tableTextStyleSchema>;

export const tableStyleSchema = z
  .strictObject({
    ...tableTextStyleObjectSchema.shape,
    ...layoutStyleSchema.unwrap().shape,
  })
  .readonly();
export type TableStyle = z.infer<typeof tableStyleSchema>;

const tableRowOwnStyleObjectSchema = z.strictObject({
  /** When unspecified, row height is determined from text measurement of unmerged cells in this row. */
  height: textPositiveSizeSchema.optional(),
  backgroundColor: colorValueSchema.optional(),
});

export const tableRowStyleSchema = z
  .strictObject({
    ...tableTextStyleObjectSchema.shape,
    ...tableRowOwnStyleObjectSchema.shape,
  })
  .readonly();
export type TableRowStyle = z.infer<typeof tableRowStyleSchema>;

export const tableBorderStyleSchema = z.enum(["solid", "dashed", "dotted"]);
export type TableBorderStyle = z.infer<typeof tableBorderStyleSchema>;

const tableCellOwnStyleObjectSchema = z.strictObject({
  backgroundColor: colorValueSchema.optional(),
  borderColor: colorValueSchema.optional(),
  borderWidth: z.number().finite().nonnegative().optional(),
  borderStyle: tableBorderStyleSchema.optional(),
  padding: textPaddingValueSchema.optional(),
  paddingTop: z.number().finite().nonnegative().optional(),
  paddingRight: z.number().finite().nonnegative().optional(),
  paddingBottom: z.number().finite().nonnegative().optional(),
  paddingLeft: z.number().finite().nonnegative().optional(),
});

export const tableCellStyleSchema = z
  .strictObject({
    ...tableTextStyleObjectSchema.shape,
    ...tableCellOwnStyleObjectSchema.shape,
  })
  .readonly();
export type TableCellStyle = z.infer<typeof tableCellStyleSchema>;

export const tableMergeSpanSchema = z.number().int().min(1);

const tableDeclarationPropsObjectSchema = z.strictObject({
  id: OptionalNameSchema.optional(),
  name: OptionalNameSchema.optional(),
  placeholder: z.string().optional(),
  nativePlaceholder: nativePlaceholderRefSchema.optional(),
  morphKey: z.custom<MorphKey>(isMorphKey).optional(),
  link: linkTargetSchema.optional(),
  x: textCoordinateSchema.optional(),
  y: textCoordinateSchema.optional(),
  /** Provides allocatable width for fr columns; when omitted, all columns are sized to natural content width. */
  width: textPositiveSizeSchema.optional(),
  columns: z.array(tableColumnSchema).nonempty().readonly(),
  style: tableStyleSchema.optional(),
  enter: z.unknown().optional(),
  emphasis: z.unknown().optional(),
  motion: z.unknown().optional(),
});

export const tableDeclarationPropsSchema =
  tableDeclarationPropsObjectSchema.readonly();

const rowDeclarationPropsObjectSchema = z.strictObject({
  id: OptionalNameSchema.optional(),
  style: tableRowStyleSchema.optional(),
});
export const rowDeclarationPropsSchema =
  rowDeclarationPropsObjectSchema.readonly();
export type RowDeclarationProps = z.infer<typeof rowDeclarationPropsSchema>;

const cellDeclarationPropsObjectSchema = z.strictObject({
  id: OptionalNameSchema.optional(),
  colSpan: tableMergeSpanSchema.optional(),
  rowSpan: tableMergeSpanSchema.optional(),
  style: tableCellStyleSchema.optional(),
});
export const cellDeclarationPropsSchema =
  cellDeclarationPropsObjectSchema.readonly();
export type CellDeclarationProps = z.infer<typeof cellDeclarationPropsSchema>;

export const tableAllowedParameters: readonly string[] = [
  ...Object.keys(tableDeclarationPropsObjectSchema.shape),
  "children",
];
export const rowAllowedParameters: readonly string[] = Object.keys(
  rowDeclarationPropsObjectSchema.shape,
);
export const cellAllowedParameters: readonly string[] = Object.keys(
  cellDeclarationPropsObjectSchema.shape,
);
export const rowStyleKeys: readonly string[] = Object.keys(
  tableRowOwnStyleObjectSchema.shape,
);
export const cellStyleKeys: readonly string[] = Object.keys(
  tableCellOwnStyleObjectSchema.shape,
);

/** OOXML `<p:cNvPr>` attribute fragment schema inside `<p:graphicFrame>`. */
export const nonVisualGraphicFrameIdSchema = z.object({
  id: OoxmlNonNegativeIntSchema,
  name: z.string().optional(),
});

/** OOXML `<a:graphicFrameLocks>` attribute fragment schema. */
export const graphicFrameLocksXmlSchema = z.object({
  noGrp: OoxmlBoolSchema.optional(),
  noDrilldown: OoxmlBoolSchema.optional(),
  noSelect: OoxmlBoolSchema.optional(),
  noChangeAspect: OoxmlBoolSchema.optional(),
  noMove: OoxmlBoolSchema.optional(),
  noResize: OoxmlBoolSchema.optional(),
});

/** OOXML `<a:tblPr>` attribute fragment schema. */
export const tblPrXmlSchema = z.object({
  firstRow: OoxmlBoolSchema.optional(),
  firstCol: OoxmlBoolSchema.optional(),
  lastRow: OoxmlBoolSchema.optional(),
  lastCol: OoxmlBoolSchema.optional(),
  bandRow: OoxmlBoolSchema.optional(),
  bandCol: OoxmlBoolSchema.optional(),
  rtl: OoxmlBoolSchema.optional(),
});

/** OOXML `<a:gridCol w="..."/>` fragment schema. */
export const tblGridColXmlSchema = z.object({
  w: OoxmlPositiveIntSchema,
});

/** OOXML `<a:tr h="..."/>` fragment schema. */
export const tblRowXmlSchema = z.object({
  h: OoxmlPositiveIntSchema,
});

/** OOXML `<a:tc gridSpan="..." rowSpan="..." hMerge="..." vMerge="..."/>` fragment schema. */
export const tblCellXmlSchema = z.object({
  gridSpan: OoxmlPositiveIntSchema.optional(),
  rowSpan: OoxmlPositiveIntSchema.optional(),
  hMerge: OoxmlBoolSchema.optional(),
  vMerge: OoxmlBoolSchema.optional(),
});

/** OOXML `<a:tcPr .../>` attribute fragment schema. */
export const tblCellPrXmlSchema = z.object({
  marL: OoxmlNonNegativeIntSchema.optional(),
  marR: OoxmlNonNegativeIntSchema.optional(),
  marT: OoxmlNonNegativeIntSchema.optional(),
  marB: OoxmlNonNegativeIntSchema.optional(),
  anchor: z.enum(["t", "ctr", "b"]).optional(),
  vert: z.string().optional(),
  horzOverflow: z.string().optional(),
  anchorCtr: OoxmlBoolSchema.optional(),
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
  const parsed = OptionalNameSchema.safeParse(value);
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
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
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
  if (!Number.isSafeInteger(emu) || emu < 1 || emu > maximumExtentEmu) {
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

/** Validates cell border width, style, and color via Zod schemas and shared theme color resolver. */
export function parseTableCellBorder(
  style: Readonly<Record<string, unknown>>,
  resolveColor: (
    value: unknown,
    label: string,
    path?: string,
  ) => ResolvedColor | undefined,
  theme: ThemeContext,
  path: string,
): TableCellBorder | undefined {
  const borderColor = resolveColor(style["borderColor"], "borderColor", path);
  const widthValue = style["borderWidth"];
  const styleValue = style["borderStyle"];

  if (widthValue !== undefined) {
    const parsedWidth = z.number().finite().nonnegative().safeParse(widthValue);
    if (!parsedWidth.success) {
      throw new CapabilityError(
        "invalid-table-border",
        "borderWidth must be a finite non-negative pixel number",
        path,
      );
    }
  }
  if (styleValue !== undefined) {
    const parsedStyle = tableBorderStyleSchema.safeParse(styleValue);
    if (!parsedStyle.success) {
      throw new CapabilityError(
        "invalid-table-border",
        "borderStyle must be solid, dashed, or dotted",
        path,
      );
    }
  }

  const width =
    typeof widthValue === "number"
      ? widthValue
      : borderColor === undefined
        ? 0
        : 1;
  if (width === 0) return undefined;
  assertTableEmuExtent(width, "borderWidth", path);
  return {
    color: borderColor ?? theme.resolveColor("dk1"),
    width,
    dash:
      styleValue === "dashed"
        ? "dash"
        : styleValue === "dotted"
          ? "dot"
          : "solid",
  };
}
