import { z } from "zod";

const integerStringPattern = /^-?\d+$/u;
const hexRgbPattern = /^[0-9a-fA-F]{6}$/u;

/** Parses an OOXML signed integer attribute string into a safe integer. */
export const OoxmlIntSchema = z
  .string()
  .regex(integerStringPattern)
  .transform((value) => Number(value))
  .refine((value) => Number.isSafeInteger(value));

/** Parses an OOXML non-negative integer attribute string into a safe integer. */
export const OoxmlNonNegativeIntSchema = OoxmlIntSchema.refine(
  (value) => value >= 0,
);

/** Parses an OOXML positive integer attribute string into a safe integer. */
export const OoxmlPositiveIntSchema = OoxmlIntSchema.refine(
  (value) => value > 0,
);

/** Parses OOXML xsd:boolean attribute values ("1" | "0" | "true" | "false"). */
export const OoxmlBoolSchema = z
  .enum(["1", "0", "true", "false"])
  .transform((value) => value === "1" || value === "true");

/** Six-digit hex RGB value stored in `<a:srgbClr val="RRGGBB"/>`. */
export const HexRgb6Schema = z.string().regex(hexRgbPattern);

/** `<a:off x="..." y="..."/>` fragment schema. */
export const OffsetEmuFragmentSchema = z.object({
  x: OoxmlIntSchema,
  y: OoxmlIntSchema,
});
export type OffsetEmuFragment = z.infer<typeof OffsetEmuFragmentSchema>;

/** `<a:ext cx="..." cy="..."/>` fragment schema. */
export const ExtentEmuFragmentSchema = z.object({
  cx: OoxmlPositiveIntSchema,
  cy: OoxmlPositiveIntSchema,
});
export type ExtentEmuFragment = z.infer<typeof ExtentEmuFragmentSchema>;

/** `<a:xfrm rot="..." flipH="..." flipV="..."/>` attribute fragment schema. */
export const TransformAttributesFragmentSchema = z.object({
  rot: OoxmlIntSchema.optional(),
  flipH: OoxmlBoolSchema.optional(),
  flipV: OoxmlBoolSchema.optional(),
});
export type TransformAttributesFragment = z.infer<
  typeof TransformAttributesFragmentSchema
>;
