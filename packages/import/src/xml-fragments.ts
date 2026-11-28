import { z } from "zod";

const integerStringPattern = /^-?\d+$/u;
const hexRgbPattern = /^[0-9a-fA-F]{6}$/u;

/** Parses an OOXML signed integer attribute string into a safe integer. */
export const ooxmlIntSchema = z
  .string()
  .regex(integerStringPattern)
  .transform((value) => Number(value))
  .refine((value) => Number.isSafeInteger(value));

/** Parses an OOXML non-negative integer attribute string into a safe integer. */
export const ooxmlNonNegativeIntSchema = ooxmlIntSchema.refine(
  (value) => value >= 0,
);

/** Parses an OOXML positive integer attribute string into a safe integer. */
export const ooxmlPositiveIntSchema = ooxmlIntSchema.refine(
  (value) => value > 0,
);

/** Parses OOXML xsd:boolean attribute values ("1" | "0" | "true" | "false"). */
export const ooxmlBoolSchema = z
  .enum(["1", "0", "true", "false"])
  .transform((value) => value === "1" || value === "true");

/** Six-digit hex RGB value stored in `<a:srgbClr val="RRGGBB"/>`. */
export const hexRgb6Schema = z.string().regex(hexRgbPattern);

/** `<a:off x="..." y="..."/>` fragment schema. */
export const offsetEmuFragmentSchema = z.object({
  x: ooxmlIntSchema,
  y: ooxmlIntSchema,
});
export type OffsetEmuFragment = z.infer<typeof offsetEmuFragmentSchema>;

/** `<a:ext cx="..." cy="..."/>` fragment schema. */
export const extentEmuFragmentSchema = z.object({
  cx: ooxmlPositiveIntSchema,
  cy: ooxmlPositiveIntSchema,
});
export type ExtentEmuFragment = z.infer<typeof extentEmuFragmentSchema>;

/** `<a:ext cx="..." cy="..."/>` fragment schema allowing zero on one axis for 1D connectors. */
export const nonNegativeExtentEmuFragmentSchema = z.object({
  cx: ooxmlNonNegativeIntSchema,
  cy: ooxmlNonNegativeIntSchema,
});
export type NonNegativeExtentEmuFragment = z.infer<
  typeof nonNegativeExtentEmuFragmentSchema
>;

/** `<a:xfrm rot="..." flipH="..." flipV="..."/>` attribute fragment schema. */
export const transformAttributesFragmentSchema = z.object({
  rot: ooxmlIntSchema.optional(),
  flipH: ooxmlBoolSchema.optional(),
  flipV: ooxmlBoolSchema.optional(),
});
export type TransformAttributesFragment = z.infer<
  typeof transformAttributesFragmentSchema
>;

function isValidXml10CodePoint(codePoint: number): boolean {
  return (
    codePoint === 0x9 ||
    codePoint === 0xa ||
    codePoint === 0xd ||
    (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
    (codePoint >= 0x10000 && codePoint <= 0x10ffff)
  );
}

function isBulletSkipCodePoint(codePoint: number): boolean {
  return (
    codePoint <= 0x20 ||
    (codePoint >= 0xfe00 && codePoint <= 0xfe0f) ||
    (codePoint >= 0xe0100 && codePoint <= 0xe01ef)
  );
}

/**
 * Normalizes `<a:buChar char="..."/>` to a single valid XML 1.0 Unicode code point,
 * falling back to `"•"` when empty or malformed.
 */
export function normalizeBulletCharacter(raw: string | undefined): {
  readonly character: string;
  readonly recovered: boolean;
} {
  if (raw !== undefined && raw.length > 0) {
    const first = raw.codePointAt(0);
    if (
      first !== undefined &&
      String.fromCodePoint(first) === raw &&
      isValidXml10CodePoint(first)
    ) {
      return { character: raw, recovered: false };
    }
    for (const unit of raw) {
      const cp = unit.codePointAt(0);
      if (
        cp !== undefined &&
        isValidXml10CodePoint(cp) &&
        !isBulletSkipCodePoint(cp)
      ) {
        return { character: String.fromCodePoint(cp), recovered: true };
      }
    }
  }
  return { character: "•", recovered: true };
}
