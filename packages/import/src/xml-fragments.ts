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

/** `<a:xfrm rot="..." flipH="..." flipV="..."/>` attribute fragment schema. */
export const transformAttributesFragmentSchema = z.object({
  rot: ooxmlIntSchema.optional(),
  flipH: ooxmlBoolSchema.optional(),
  flipV: ooxmlBoolSchema.optional(),
});
export type TransformAttributesFragment = z.infer<
  typeof transformAttributesFragmentSchema
>;
