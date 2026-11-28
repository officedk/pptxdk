import { assertValidLang, type TextStyle } from "@pptxdk/core";
import { z } from "zod";

import type { ImportDiagnostic, ImportLocation } from "#/diagnostics.ts";
import { EMU_PER_PX, emuToPx } from "#/namespaces.ts";
import type { XmlElementCursor } from "#/xml-reader.ts";

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

/** DrawingML `<a:rPr sz="...">` valid integer centipoint domain (`1pt .. 4000pt`). */
export const ooxmlFontSizeCentipointsSchema = ooxmlIntSchema.refine(
  (value) => value >= 100 && value <= 400_000,
);

/** DrawingML `<a:rPr spc="...">` valid integer centipoint domain (`-4000pt .. 4000pt`). */
export const ooxmlLetterSpacingCentipointsSchema = ooxmlIntSchema.refine(
  (value) => value >= -400_000 && value <= 400_000,
);

/** DrawingML `<a:rPr baseline="...">` valid 32-bit signed thousandths of a percent (`ST_Percentage`). */
export const ooxmlBaselineShiftThousandthsSchema = ooxmlIntSchema.refine(
  (value) => value >= -2_147_483_648 && value <= 2_147_483_647,
);

/** DrawingML `<a:lnSpc><a:spcPts val="...">` valid centipoint domain (`>0 .. 1584pt`). */
export const ooxmlLineSpacingPointsSchema = ooxmlIntSchema.refine(
  (value) => value >= 1 && value <= 158_400,
);

/** DrawingML `<a:lnSpc><a:spcPct val="...">` valid thousandth-percent domain (`>0% .. 13200%`). */
export const ooxmlLineSpacingPercentSchema = ooxmlIntSchema.refine(
  (value) => value >= 1 && value <= 13_200_000,
);

/** DrawingML `<a:spcBef|spcAft><a:spcPts val="...">` valid non-negative centipoint domain (`0 .. 1584pt`). */
export const ooxmlParagraphSpacingPointsSchema = ooxmlIntSchema.refine(
  (value) => value >= 0 && value <= 158_400,
);

/** DrawingML `<a:buAutoNum startAt="...">` valid integer domain (`1 .. 32767`). */
export const ooxmlAutoNumStartAtSchema = ooxmlIntSchema.refine(
  (value) => value >= 1 && value <= 32_767,
);

/** DrawingML `<a:pPr lvl="...">` valid paragraph level (`0 .. 8`). */
export const ooxmlListLevelSchema = ooxmlIntSchema.refine(
  (value) => value >= 0 && value <= 8,
);

/** DrawingML `ST_PositiveFixedAngle` domain (`0 <= value < 21_600_000`, 60,000ths of a degree). */
export const ooxmlPositiveFixedAngleSchema = ooxmlIntSchema.refine(
  (value) => value >= 0 && value < 21_600_000,
);

export type OoxmlRunAttrStyleResult = {
  readonly fontSize?: NonNullable<TextStyle["fontSize"]>;
  readonly sizePt?: number;
  readonly fontWeight?: NonNullable<TextStyle["fontWeight"]>;
  readonly fontStyle?: NonNullable<TextStyle["fontStyle"]>;
  readonly textDecoration?: NonNullable<TextStyle["textDecoration"]>;
  readonly textTransform?: NonNullable<TextStyle["textTransform"]>;
  readonly letterSpacing?: number;
  readonly lang?: string;
  readonly baselineShift?: `${number}%`;
  readonly inlineKind?: "sup" | "sub";
};

function pushInvalidXmlAttrDiagnostic(
  diagnostics: ImportDiagnostic[],
  location: ImportLocation,
  elementXpath: string,
  attrName: string,
  rawValue: string,
): void {
  diagnostics.push({
    code: "invalid-xml-value",
    category: "corrupt-input",
    disposition: "omitted",
    location: { ...location, xpath: `${elementXpath}/@${attrName}` },
    attributeOrElement: `@${attrName}`,
    rawValue,
  });
}

/**
 * Parses `<a:rPr>`, `<a:defRPr>`, or `<a:endParaRPr>` scalar attributes independently so an
 * out-of-range or malformed attribute strips only itself (recording `invalid-xml-value`) while
 * preserving valid sibling attributes.
 */
export function readOoxmlRunAttrStyles(args: {
  readonly rPrCursor: XmlElementCursor;
  readonly location: ImportLocation;
  readonly diagnostics: ImportDiagnostic[];
  readonly allowBaseline?: boolean;
  readonly allowNoneTextDecoration?: boolean;
  readonly diagnoseUnsupportedEnums?: boolean;
}): OoxmlRunAttrStyleResult {
  const { rPrCursor, location, diagnostics } = args;
  const raw = rPrCursor.attrsObject([
    "lang",
    "sz",
    "b",
    "i",
    "u",
    "strike",
    "cap",
    "spc",
    "kern",
    "dirty",
    "smtClean",
    "baseline",
  ]);

  let lang: string | undefined;
  if (raw["lang"] !== undefined) {
    try {
      lang = assertValidLang(raw["lang"]);
    } catch {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "lang",
        raw["lang"],
      );
    }
  }

  let fontSize: NonNullable<TextStyle["fontSize"]> | undefined;
  let sizePt: number | undefined;
  if (raw["sz"] !== undefined) {
    const parsed = ooxmlFontSizeCentipointsSchema.safeParse(raw["sz"]);
    if (parsed.success) {
      sizePt = parsed.data / 100;
      fontSize = `${sizePt}pt`;
    } else {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "sz",
        raw["sz"],
      );
    }
  }

  let fontWeight: NonNullable<TextStyle["fontWeight"]> | undefined;
  if (raw["b"] !== undefined) {
    const parsed = ooxmlBoolSchema.safeParse(raw["b"]);
    if (parsed.success) {
      fontWeight = parsed.data ? "bold" : "normal";
    } else {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "b",
        raw["b"],
      );
    }
  }

  let fontStyle: NonNullable<TextStyle["fontStyle"]> | undefined;
  if (raw["i"] !== undefined) {
    const parsed = ooxmlBoolSchema.safeParse(raw["i"]);
    if (parsed.success) {
      fontStyle = parsed.data ? "italic" : "normal";
    } else {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "i",
        raw["i"],
      );
    }
  }

  const rawU = raw["u"];
  if (
    rawU !== undefined &&
    rawU !== "none" &&
    rawU !== "sng" &&
    args.diagnoseUnsupportedEnums === true
  ) {
    pushInvalidXmlAttrDiagnostic(
      diagnostics,
      location,
      rPrCursor.xpath,
      "u",
      rawU,
    );
  }
  const rawStrike = raw["strike"];
  if (
    rawStrike !== undefined &&
    rawStrike !== "noStrike" &&
    rawStrike !== "sngStrike" &&
    args.diagnoseUnsupportedEnums === true
  ) {
    pushInvalidXmlAttrDiagnostic(
      diagnostics,
      location,
      rPrCursor.xpath,
      "strike",
      rawStrike,
    );
  }
  const textDecoration: NonNullable<TextStyle["textDecoration"]> | undefined =
    rawU === "sng"
      ? "underline"
      : rawStrike === "sngStrike"
        ? "line-through"
        : args.allowNoneTextDecoration === true &&
            rawU === "none" &&
            rawStrike === "noStrike"
          ? "none"
          : undefined;

  const rawCap = raw["cap"];
  if (
    rawCap !== undefined &&
    rawCap !== "none" &&
    rawCap !== "all" &&
    args.diagnoseUnsupportedEnums === true
  ) {
    pushInvalidXmlAttrDiagnostic(
      diagnostics,
      location,
      rPrCursor.xpath,
      "cap",
      rawCap,
    );
  }
  const textTransform: NonNullable<TextStyle["textTransform"]> | undefined =
    rawCap === "all" ? "uppercase" : rawCap === "none" ? "none" : undefined;

  let letterSpacing: number | undefined;
  if (raw["spc"] !== undefined) {
    const parsed = ooxmlLetterSpacingCentipointsSchema.safeParse(raw["spc"]);
    if (parsed.success) {
      letterSpacing = parsed.data === 0 ? 0 : parsed.data / 75;
    } else {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "spc",
        raw["spc"],
      );
    }
  }

  let baselineShift: `${number}%` | undefined;
  let inlineKind: "sup" | "sub" | undefined;
  if (args.allowBaseline === true && raw["baseline"] !== undefined) {
    const parsed = ooxmlBaselineShiftThousandthsSchema.safeParse(
      raw["baseline"],
    );
    if (!parsed.success) {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        rPrCursor.xpath,
        "baseline",
        raw["baseline"],
      );
    } else if (parsed.data === 30_000) {
      inlineKind = "sup";
    } else if (parsed.data === -25_000) {
      inlineKind = "sub";
    } else if (parsed.data !== 0) {
      baselineShift = `${parsed.data / 1000}%`;
    }
  }

  return {
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(sizePt === undefined ? {} : { sizePt }),
    ...(fontWeight === undefined ? {} : { fontWeight }),
    ...(fontStyle === undefined ? {} : { fontStyle }),
    ...(textDecoration === undefined ? {} : { textDecoration }),
    ...(textTransform === undefined ? {} : { textTransform }),
    ...(letterSpacing === undefined ? {} : { letterSpacing }),
    ...(lang === undefined ? {} : { lang }),
    ...(baselineShift === undefined ? {} : { baselineShift }),
    ...(inlineKind === undefined ? {} : { inlineKind }),
  };
}

/** Parses OOXML xsd:boolean attribute values ("1" | "0" | "true" | "false"). */
export const ooxmlBoolSchema = z
  .enum(["1", "0", "true", "false"])
  .transform((value) => value === "1" || value === "true");

/** Six-digit hex RGB value stored in `<a:srgbClr val="RRGGBB"/>`. */
export const hexRgb6Schema = z.string().regex(hexRgbPattern);

/** Maximum logical pixel coordinate allowed by downstream builders (`[-1_000_000, 1_000_000]`). */
export const MAXIMUM_COORDINATE_PX = 1_000_000;
/** Maximum positive signed 32-bit DrawingML EMU value (`ST_PositiveCoordinate32`). */
export const MAXIMUM_WRITABLE_EMU = 2_147_483_647;
/** Minimum signed 32-bit DrawingML EMU value. */
export const MINIMUM_WRITABLE_INT32_EMU = -2_147_483_648;
/** Maximum DrawingML stroke line width in EMU (`ST_LineWidth`). */
export const MAXIMUM_LINE_WIDTH_EMU = 20_116_800;

function isWritablePixelCoordinateEmu(emu: number): boolean {
  const px = emuToPx(emu);
  const roundtrip = Math.round(px * EMU_PER_PX);
  return (
    Number.isFinite(px) &&
    Math.abs(px) <= MAXIMUM_COORDINATE_PX &&
    Number.isSafeInteger(roundtrip)
  );
}

function isWritableNonNegativeEmu(
  emu: number,
  maxEmu = MAXIMUM_WRITABLE_EMU,
): boolean {
  const roundtrip = Math.round(emuToPx(emu) * EMU_PER_PX);
  return (
    emu >= 0 &&
    emu <= maxEmu &&
    Number.isSafeInteger(roundtrip) &&
    roundtrip >= 0 &&
    roundtrip <= maxEmu
  );
}

function isWritablePositiveEmu(
  emu: number,
  maxEmu = MAXIMUM_WRITABLE_EMU,
): boolean {
  const roundtrip = Math.round(emuToPx(emu) * EMU_PER_PX);
  return (
    emu >= 1 &&
    emu <= maxEmu &&
    Number.isSafeInteger(roundtrip) &&
    roundtrip >= 1 &&
    roundtrip <= maxEmu
  );
}

/** DrawingML non-negative writable EMU extent (`0 .. 2_147_483_647`). */
export const ooxmlWritableExtentEmuSchema = ooxmlNonNegativeIntSchema.refine(
  (value) => isWritableNonNegativeEmu(value),
);

/** DrawingML positive writable EMU extent (`1 .. 2_147_483_647`). */
export const ooxmlWritablePositiveExtentEmuSchema =
  ooxmlPositiveIntSchema.refine((value) => isWritablePositiveEmu(value));

/** DrawingML signed 32-bit EMU coordinate within `[-1_000_000, 1_000_000]` px after conversion. */
export const ooxmlWritableInt32CoordEmuSchema = ooxmlIntSchema.refine(
  (value) =>
    value >= MINIMUM_WRITABLE_INT32_EMU &&
    value <= MAXIMUM_WRITABLE_EMU &&
    isWritablePixelCoordinateEmu(value),
);

/** DrawingML `<a:ln w="...">` valid line width domain (`0 .. 20_116_800` EMU). */
export const ooxmlLineWidthEmuSchema = ooxmlNonNegativeIntSchema.refine(
  (value) => isWritableNonNegativeEmu(value, MAXIMUM_LINE_WIDTH_EMU),
);

/** DrawingML `<a:ln w="...">` strictly positive line width domain (`1 .. 20_116_800` EMU). */
export const ooxmlPositiveLineWidthEmuSchema = ooxmlPositiveIntSchema.refine(
  (value) => isWritablePositiveEmu(value, MAXIMUM_LINE_WIDTH_EMU),
);

/** `<a:off x="..." y="..."/>` fragment schema bounded to `[-1_000_000, 1_000_000]` px. */
export const offsetEmuFragmentSchema = z.object({
  x: ooxmlIntSchema.refine(isWritablePixelCoordinateEmu),
  y: ooxmlIntSchema.refine(isWritablePixelCoordinateEmu),
});
export type OffsetEmuFragment = z.infer<typeof offsetEmuFragmentSchema>;

/** `<a:ext cx="..." cy="..."/>` fragment schema bounded to `1 .. 2_147_483_647` EMU. */
export const extentEmuFragmentSchema = z.object({
  cx: ooxmlWritablePositiveExtentEmuSchema,
  cy: ooxmlWritablePositiveExtentEmuSchema,
});
export type ExtentEmuFragment = z.infer<typeof extentEmuFragmentSchema>;

/** `<a:ext cx="..." cy="..."/>` fragment schema allowing zero on one axis for 1D connectors, bounded to `0 .. 2_147_483_647` EMU. */
export const nonNegativeExtentEmuFragmentSchema = z.object({
  cx: ooxmlWritableExtentEmuSchema,
  cy: ooxmlWritableExtentEmuSchema,
});
export type NonNegativeExtentEmuFragment = z.infer<
  typeof nonNegativeExtentEmuFragmentSchema
>;

export type OoxmlBoxInsetsResult = {
  readonly paddingTop?: number;
  readonly paddingRight?: number;
  readonly paddingBottom?: number;
  readonly paddingLeft?: number;
};

/**
 * Parses box inset/margin attributes (`lIns/tIns/rIns/bIns` or `marL/marT/marR/marB`) independently
 * within DrawingML `0 .. 2_147_483_647` EMU bounds and enforces opposing-pair sums `<= 2_147_483_647` EMU.
 */
export function readOoxmlBoxInsets(args: {
  readonly rawAttrs: Readonly<Record<string, string | undefined>>;
  readonly attrNames: {
    readonly left: "lIns" | "marL";
    readonly top: "tIns" | "marT";
    readonly right: "rIns" | "marR";
    readonly bottom: "bIns" | "marB";
  };
  readonly elementXpath: string;
  readonly location: ImportLocation;
  readonly diagnostics: ImportDiagnostic[];
}): OoxmlBoxInsetsResult {
  const { rawAttrs, attrNames, elementXpath, location, diagnostics } = args;
  const parseEdge = (attrName: string): number | undefined => {
    const raw = rawAttrs[attrName];
    if (raw === undefined) return undefined;
    const parsed = ooxmlWritableExtentEmuSchema.safeParse(raw);
    if (!parsed.success) {
      pushInvalidXmlAttrDiagnostic(
        diagnostics,
        location,
        elementXpath,
        attrName,
        raw,
      );
      return undefined;
    }
    return parsed.data;
  };

  let leftEmu = parseEdge(attrNames.left);
  let rightEmu = parseEdge(attrNames.right);
  let topEmu = parseEdge(attrNames.top);
  let bottomEmu = parseEdge(attrNames.bottom);

  if (
    leftEmu !== undefined &&
    rightEmu !== undefined &&
    leftEmu + rightEmu > MAXIMUM_WRITABLE_EMU
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { ...location, xpath: elementXpath },
      attributeOrElement: `${attrNames.left}+${attrNames.right}`,
      rawValue: `${leftEmu}+${rightEmu}`,
    });
    leftEmu = undefined;
    rightEmu = undefined;
  }
  if (
    topEmu !== undefined &&
    bottomEmu !== undefined &&
    topEmu + bottomEmu > MAXIMUM_WRITABLE_EMU
  ) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { ...location, xpath: elementXpath },
      attributeOrElement: `${attrNames.top}+${attrNames.bottom}`,
      rawValue: `${topEmu}+${bottomEmu}`,
    });
    topEmu = undefined;
    bottomEmu = undefined;
  }

  return {
    ...(topEmu !== undefined && topEmu !== 0
      ? { paddingTop: emuToPx(topEmu) }
      : {}),
    ...(rightEmu !== undefined && rightEmu !== 0
      ? { paddingRight: emuToPx(rightEmu) }
      : {}),
    ...(bottomEmu !== undefined && bottomEmu !== 0
      ? { paddingBottom: emuToPx(bottomEmu) }
      : {}),
    ...(leftEmu !== undefined && leftEmu !== 0
      ? { paddingLeft: emuToPx(leftEmu) }
      : {}),
  };
}

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
