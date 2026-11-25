import {
  colorMapSchema,
  defaultColorMap,
  defineTheme,
  nativeColorSlots,
  schemeColorNames,
  schemeColorSlots,
  type ColorMap,
  type ColorMapOverride,
  type FormatColorTransform,
  type FormatEffectStyle,
  type FormatFillStyle,
  type FormatLineStyle,
  type FormatScheme,
  type NativeColorSlot,
  type Shadow,
  type Theme,
  type ThemeOverride,
} from "@pptxdk/core";
import { z } from "zod";

import type { ImportDiagnostic } from "#/diagnostics.ts";
import { A_NS, P_NS } from "#/namespaces.ts";
import type { PptxPackage } from "#/opc.ts";
import {
  hexRgb6Schema,
  ooxmlBoolSchema,
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
} from "#/xml-fragments.ts";
import { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";

const percentValAttrSchema = z.object({
  val: ooxmlNonNegativeIntSchema,
});

const percentPosAttrSchema = z.object({
  pos: ooxmlNonNegativeIntSchema,
});

const sysClrAttrSchema = z.object({
  val: z.string().min(1),
  lastClr: hexRgb6Schema.optional(),
});

const colorMapXmlAttrSchema = colorMapSchema;

const lineAttrSchema = z.object({
  w: ooxmlNonNegativeIntSchema.optional(),
  cap: z.enum(["rnd", "sq", "flat"]).optional(),
  cmpd: z.enum(["sng", "dbl", "thickThin", "thinThick", "tri"]).optional(),
  algn: z.enum(["ctr", "in"]).optional(),
});

const linearGradAttrSchema = z.object({
  ang: ooxmlIntSchema.optional(),
  scaled: ooxmlBoolSchema.optional(),
});

const pathGradAttrSchema = z.object({
  path: z.enum(["circle", "rect", "shape"]).optional(),
});

const fillToRectAttrSchema = z.object({
  l: ooxmlIntSchema.optional(),
  t: ooxmlIntSchema.optional(),
  r: ooxmlIntSchema.optional(),
  b: ooxmlIntSchema.optional(),
});

const outerShdwAttrSchema = z.object({
  blurRad: ooxmlNonNegativeIntSchema.optional(),
  dist: ooxmlNonNegativeIntSchema.optional(),
  dir: ooxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: ooxmlBoolSchema.optional(),
});

const officeFallbackColors: Record<NativeColorSlot, string> = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#1F497D",
  lt2: "#EEECE1",
  accent1: "#4F81BD",
  accent2: "#C0504D",
  accent3: "#9BBB59",
  accent4: "#8064A2",
  accent5: "#4BACC6",
  accent6: "#F79646",
  hlink: "#0000FF",
  folHlink: "#800080",
};

function colorMapsEqual(left: ColorMap, right: ColorMap): boolean {
  return schemeColorSlots.every((slot) => left[slot] === right[slot]);
}

/** Reads `<p:clrMap>` or `<a:overrideClrMapping>` attributes into a validated 12-slot `ColorMap`. */
export function readColorMapElement(
  cursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): ColorMap {
  const raw = cursor.attrsObject(schemeColorSlots);
  const parsed = colorMapXmlAttrSchema.safeParse(raw);
  if (parsed.success) {
    return { ...parsed.data };
  }
  for (const issue of parsed.error.issues) {
    const slot = String(issue.path[0] ?? "clrMap");
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: {
        part: partPath,
        xpath: `${cursor.xpath}/@${slot}`,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
      attributeOrElement: `@${slot}`,
      rawValue: raw[slot] ?? "",
    });
  }
  const recovered = { ...defaultColorMap } as Record<string, NativeColorSlot>;
  for (const slot of schemeColorSlots) {
    const candidate = raw[slot];
    if (
      candidate !== undefined &&
      (nativeColorSlots as readonly string[]).includes(candidate)
    ) {
      recovered[slot] = candidate as NativeColorSlot;
    }
  }
  return colorMapSchema.parse(recovered);
}

/** Reads `<p:clrMapOvr>` into `"master"` or an explicit 12-slot `ColorMap`. */
export function readColorMapOverride(
  clrMapOvrCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): ColorMapOverride | undefined {
  if (clrMapOvrCursor.ignoreEmptyChild(A_NS, "masterClrMapping")) {
    return "master";
  }
  const overrideCursor = clrMapOvrCursor.child(A_NS, "overrideClrMapping");
  if (overrideCursor !== undefined) {
    return readColorMapElement(
      overrideCursor,
      partPath,
      slideIndex,
      diagnostics,
    );
  }
  return undefined;
}

function resolveSimpleThemeHex(
  slotCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): string | undefined {
  let result: string | undefined;
  for (const child of slotCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    if (child.localName === "srgbClr" && result === undefined) {
      const raw = child.attrsObject(["val"]);
      const parsed = hexRgb6Schema.safeParse(raw["val"]);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: {
            part: partPath,
            xpath: `${child.xpath}/@val`,
            ...(slideIndex === undefined ? {} : { slideIndex }),
          },
          attributeOrElement: "a:srgbClr/@val",
          rawValue: raw["val"] ?? "",
        });
      } else {
        result = `#${parsed.data.toUpperCase()}`;
      }
      child.suppressUnconsumed();
    } else if (child.localName === "sysClr" && result === undefined) {
      const raw = child.attrsObject(["val", "lastClr"]);
      const parsed = sysClrAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "recovered",
          location: {
            part: partPath,
            xpath: child.xpath,
            ...(slideIndex === undefined ? {} : { slideIndex }),
          },
          attributeOrElement: "a:sysClr",
          rawValue: JSON.stringify(raw),
        });
      } else {
        const fallback =
          parsed.data.val === "windowText"
            ? "000000"
            : parsed.data.val === "window"
              ? "FFFFFF"
              : undefined;
        const hex = parsed.data.lastClr ?? fallback;
        if (hex !== undefined) {
          result = `#${hex.toUpperCase()}`;
        }
      }
      child.suppressUnconsumed();
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `theme-color:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }
  return result;
}

function readClrScheme(
  clrSchemeCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): Record<NativeColorSlot, string> {
  clrSchemeCursor.consumeAttr("name");
  const colors = { ...officeFallbackColors };
  for (const child of clrSchemeCursor.remainingElementChildren()) {
    if (
      child.namespaceUri === A_NS &&
      (nativeColorSlots as readonly string[]).includes(child.localName)
    ) {
      const slot = child.localName as NativeColorSlot;
      const hex = resolveSimpleThemeHex(
        child,
        partPath,
        slideIndex,
        diagnostics,
      );
      if (hex !== undefined) {
        colors[slot] = hex;
      }
    } else {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
    }
  }
  return colors;
}

function readFontCollection(
  fontCollectionCursor: XmlElementCursor,
): string | undefined {
  let latinTypeface: string | undefined;
  for (const child of fontCollectionCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) continue;
    if (child.localName === "latin") {
      const typeface = child.attr("typeface");
      if (typeface !== undefined && typeface.trim() !== "") {
        latinTypeface = typeface;
      }
      child.consumeAttr("panose");
      child.consumeAttr("pitchFamily");
      child.consumeAttr("charset");
    } else if (child.localName === "ea" || child.localName === "cs") {
      child.consumeAttr("typeface");
      child.consumeAttr("panose");
      child.consumeAttr("pitchFamily");
      child.consumeAttr("charset");
    } else if (child.localName === "font") {
      child.consumeAttr("script");
      child.consumeAttr("typeface");
    }
  }
  return latinTypeface;
}

function readFontScheme(fontSchemeCursor: XmlElementCursor): {
  readonly major?: string;
  readonly minor?: string;
} {
  fontSchemeCursor.consumeAttr("name");
  const majorCursor = fontSchemeCursor.child(A_NS, "majorFont");
  const minorCursor = fontSchemeCursor.child(A_NS, "minorFont");
  const major =
    majorCursor === undefined ? undefined : readFontCollection(majorCursor);
  const minor =
    minorCursor === undefined ? undefined : readFontCollection(minorCursor);
  return {
    ...(major === undefined ? {} : { major }),
    ...(minor === undefined ? {} : { minor }),
  };
}

function readFormatColorTransform(
  parentCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): FormatColorTransform | undefined {
  let colorCursor: XmlElementCursor | undefined;
  let colorValue: FormatColorTransform["color"] | undefined;
  for (const child of parentCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) continue;
    if (child.localName === "schemeClr" && colorCursor === undefined) {
      colorCursor = child;
      const val = child.attr("val");
      if (
        val === "phClr" ||
        (val !== undefined &&
          (schemeColorNames as readonly string[]).includes(val))
      ) {
        colorValue = val;
      }
    } else if (child.localName === "srgbClr" && colorCursor === undefined) {
      colorCursor = child;
      const val = child.attr("val");
      const parsed = hexRgb6Schema.safeParse(val);
      if (parsed.success) {
        colorValue = `#${parsed.data.toUpperCase()}`;
      }
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `format-color:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }
  if (colorCursor === undefined || colorValue === undefined) return undefined;

  let alpha: number | undefined;
  let lumMod: number | undefined;
  let lumOff: number | undefined;
  let tint: number | undefined;
  let shade: number | undefined;
  let satMod: number | undefined;
  for (const mod of colorCursor.remainingElementChildren()) {
    if (mod.namespaceUri !== A_NS) continue;
    const parsed = percentValAttrSchema.safeParse(mod.attrsObject(["val"]));
    if (!parsed.success) continue;
    const fraction = parsed.data.val / 100_000;
    switch (mod.localName) {
      case "alpha":
        alpha = Math.min(1, Math.max(0, fraction));
        break;
      case "lumMod":
        lumMod = Math.min(1, Math.max(0, fraction));
        break;
      case "lumOff":
        lumOff = Math.min(1, Math.max(0, fraction));
        break;
      case "tint":
        tint = Math.min(1, Math.max(0, fraction));
        break;
      case "shade":
        shade = Math.min(1, Math.max(0, fraction));
        break;
      case "satMod":
        satMod = Math.max(0, fraction);
        break;
      default:
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: {
            part: partPath,
            xpath: mod.xpath,
            ...(slideIndex === undefined ? {} : { slideIndex }),
          },
          namespaceUri: mod.namespaceUri,
          localName: mod.localName,
          mechanismHint: `format-color-transform:${mod.localName}`,
        });
        break;
    }
  }

  return {
    color: colorValue,
    ...(alpha === undefined ? {} : { alpha }),
    ...(lumMod === undefined ? {} : { lumMod }),
    ...(lumOff === undefined ? {} : { lumOff }),
    ...(tint === undefined ? {} : { tint }),
    ...(shade === undefined ? {} : { shade }),
    ...(satMod === undefined ? {} : { satMod }),
  };
}

function readFormatFillStyle(
  cursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): FormatFillStyle | undefined {
  if (cursor.namespaceUri !== A_NS) return undefined;
  if (cursor.localName === "noFill") {
    return { kind: "none" };
  }
  if (cursor.localName === "solidFill") {
    const color = readFormatColorTransform(
      cursor,
      partPath,
      slideIndex,
      diagnostics,
    );
    return color === undefined ? undefined : { kind: "solid", color };
  }
  if (cursor.localName === "gradFill") {
    const gradAttrs = linearGradAttrSchema
      .extend({
        rotWithShape: ooxmlBoolSchema.optional(),
      })
      .safeParse(cursor.attrsObject(["rotWithShape"]));
    const rotWithShape = gradAttrs.success
      ? gradAttrs.data.rotWithShape
      : undefined;
    cursor.consumeAttr("flip");
    const stops: {
      readonly position: number;
      readonly color: FormatColorTransform;
    }[] = [];
    const gsLst = cursor.child(A_NS, "gsLst");
    if (gsLst !== undefined) {
      for (const gs of gsLst.children(A_NS, "gs")) {
        const posParsed = percentPosAttrSchema.safeParse(
          gs.attrsObject(["pos"]),
        );
        const color = readFormatColorTransform(
          gs,
          partPath,
          slideIndex,
          diagnostics,
        );
        if (posParsed.success && color !== undefined) {
          stops.push({
            position: Math.min(1, Math.max(0, posParsed.data.pos / 100_000)),
            color,
          });
        }
      }
    }
    if (stops.length === 0) {
      return undefined;
    }
    let linear:
      | { readonly angle: number; readonly scaled?: boolean }
      | undefined;
    const lin = cursor.child(A_NS, "lin");
    if (lin !== undefined) {
      const parsed = linearGradAttrSchema.safeParse(
        lin.attrsObject(["ang", "scaled"]),
      );
      if (parsed.success) {
        linear = {
          angle: (parsed.data.ang ?? 0) / 60_000,
          ...(parsed.data.scaled === undefined
            ? {}
            : { scaled: parsed.data.scaled }),
        };
      }
    }
    let pathSpec:
      | {
          readonly shape: "circle" | "rect" | "shape";
          readonly fillToRect?: {
            readonly l: number;
            readonly t: number;
            readonly r: number;
            readonly b: number;
          };
        }
      | undefined;
    const pathCursor = cursor.child(A_NS, "path");
    if (pathCursor !== undefined) {
      const parsedPath = pathGradAttrSchema.safeParse(
        pathCursor.attrsObject(["path"]),
      );
      const rectCursor = pathCursor.child(A_NS, "fillToRect");
      let fillToRect:
        | {
            readonly l: number;
            readonly t: number;
            readonly r: number;
            readonly b: number;
          }
        | undefined;
      if (rectCursor !== undefined) {
        const parsedRect = fillToRectAttrSchema.safeParse(
          rectCursor.attrsObject(["l", "t", "r", "b"]),
        );
        if (parsedRect.success) {
          fillToRect = {
            l: (parsedRect.data.l ?? 0) / 100_000,
            t: (parsedRect.data.t ?? 0) / 100_000,
            r: (parsedRect.data.r ?? 0) / 100_000,
            b: (parsedRect.data.b ?? 0) / 100_000,
          };
        }
      }
      pathSpec = {
        shape: parsedPath.success ? (parsedPath.data.path ?? "rect") : "rect",
        ...(fillToRect === undefined ? {} : { fillToRect }),
      };
    }
    cursor.ignoreEmptyChild(A_NS, "tileRect");
    return {
      kind: "gradient",
      ...(rotWithShape === undefined ? {} : { rotWithShape }),
      stops,
      ...(linear === undefined ? {} : { linear }),
      ...(pathSpec === undefined ? {} : { path: pathSpec }),
    };
  }
  diagnostics.push({
    code: "unsupported-mechanism",
    category: "unsupported",
    disposition: "omitted",
    location: {
      part: partPath,
      xpath: cursor.xpath,
      ...(slideIndex === undefined ? {} : { slideIndex }),
    },
    namespaceUri: cursor.namespaceUri,
    localName: cursor.localName,
    mechanismHint: `format-fill:${cursor.localName}`,
  });
  cursor.suppressUnconsumed();
  return undefined;
}

function readFormatLineStyle(
  lnCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): FormatLineStyle {
  const parsed = lineAttrSchema.safeParse(
    lnCursor.attrsObject(["w", "cap", "cmpd", "algn"]),
  );
  let fill: FormatFillStyle | undefined;
  let dash: string | undefined;
  for (const child of lnCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) continue;
    if (
      (child.localName === "solidFill" ||
        child.localName === "gradFill" ||
        child.localName === "noFill") &&
      fill === undefined
    ) {
      fill = readFormatFillStyle(child, partPath, slideIndex, diagnostics);
    } else if (child.localName === "prstDash") {
      const val = child.attr("val");
      if (val !== undefined) dash = val;
    } else if (
      child.localName === "miter" ||
      child.localName === "round" ||
      child.localName === "bevel"
    ) {
      child.consumeAttr("lim");
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `format-line:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }
  return {
    ...(parsed.success && parsed.data.w !== undefined
      ? { widthEmu: parsed.data.w }
      : {}),
    ...(parsed.success && parsed.data.cap !== undefined
      ? { cap: parsed.data.cap }
      : {}),
    ...(parsed.success && parsed.data.cmpd !== undefined
      ? { compound: parsed.data.cmpd }
      : {}),
    ...(parsed.success && parsed.data.algn !== undefined
      ? { align: parsed.data.algn }
      : {}),
    ...(dash === undefined ? {} : { dash }),
    ...(fill === undefined ? {} : { fill }),
  };
}

function readFormatEffectStyle(
  effectCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): FormatEffectStyle {
  let shadow: Shadow | undefined;
  for (const child of effectCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) continue;
    if (child.localName === "effectLst") {
      const outerShdw = child.child(A_NS, "outerShdw");
      if (outerShdw !== undefined) {
        const parsed = outerShdwAttrSchema.safeParse(
          outerShdw.attrsObject([
            "blurRad",
            "dist",
            "dir",
            "algn",
            "rotWithShape",
          ]),
        );
        const colorHex = resolveSimpleThemeHex(
          outerShdw,
          partPath,
          slideIndex,
          diagnostics,
        );
        shadow = {
          ...(colorHex === undefined ? {} : { color: colorHex }),
          ...(parsed.success && parsed.data.blurRad !== undefined
            ? { blur: parsed.data.blurRad / 12_700 }
            : {}),
          ...(parsed.success && parsed.data.dist !== undefined
            ? { distance: parsed.data.dist / 12_700 }
            : {}),
          ...(parsed.success && parsed.data.dir !== undefined
            ? { angle: parsed.data.dir / 60_000 }
            : {}),
        };
      }
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: {
          part: partPath,
          xpath: child.xpath,
          ...(slideIndex === undefined ? {} : { slideIndex }),
        },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `format-effect:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }
  return shadow === undefined ? {} : { shadow };
}

function readFmtScheme(
  fmtSchemeCursor: XmlElementCursor,
  partPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): FormatScheme {
  const name = fmtSchemeCursor.attr("name");
  const fillStyles: FormatFillStyle[] = [];
  const lineStyles: FormatLineStyle[] = [];
  const effectStyles: FormatEffectStyle[] = [];
  const bgFillStyles: FormatFillStyle[] = [];

  const fillStyleLst = fmtSchemeCursor.child(A_NS, "fillStyleLst");
  if (fillStyleLst !== undefined) {
    for (const child of fillStyleLst.remainingElementChildren()) {
      const parsed = readFormatFillStyle(
        child,
        partPath,
        slideIndex,
        diagnostics,
      );
      if (parsed !== undefined) fillStyles.push(parsed);
    }
  }
  const lnStyleLst = fmtSchemeCursor.child(A_NS, "lnStyleLst");
  if (lnStyleLst !== undefined) {
    for (const child of lnStyleLst.children(A_NS, "ln")) {
      lineStyles.push(
        readFormatLineStyle(child, partPath, slideIndex, diagnostics),
      );
    }
  }
  const effectStyleLst = fmtSchemeCursor.child(A_NS, "effectStyleLst");
  if (effectStyleLst !== undefined) {
    for (const child of effectStyleLst.children(A_NS, "effectStyle")) {
      effectStyles.push(
        readFormatEffectStyle(child, partPath, slideIndex, diagnostics),
      );
    }
  }
  const bgFillStyleLst = fmtSchemeCursor.child(A_NS, "bgFillStyleLst");
  if (bgFillStyleLst !== undefined) {
    for (const child of bgFillStyleLst.remainingElementChildren()) {
      const parsed = readFormatFillStyle(
        child,
        partPath,
        slideIndex,
        diagnostics,
      );
      if (parsed !== undefined) bgFillStyles.push(parsed);
    }
  }

  return {
    ...(name === undefined ? {} : { name }),
    ...(fillStyles.length === 0 ? {} : { fillStyles }),
    ...(lineStyles.length === 0 ? {} : { lineStyles }),
    ...(effectStyles.length === 0 ? {} : { effectStyles }),
    ...(bgFillStyles.length === 0 ? {} : { bgFillStyles }),
  };
}

/**
 * Reads a native `ppt/theme/theme{n}.xml` part into a branded `Theme` instance,
 * deduplicating shared OPC theme parts by normalized target path so masters sharing
 * `ppt/theme/theme1.xml` share the exact same `Theme` object identity.
 */
export function readThemePart(
  pkg: PptxPackage,
  themePartPath: string,
  themesByPartPath: Map<string, Theme>,
  diagnostics: ImportDiagnostic[],
): Theme | undefined {
  const cached = themesByPartPath.get(themePartPath);
  if (cached !== undefined) return cached;

  const xmlText = pkg.readTextPart(themePartPath);
  if (xmlText === undefined) return undefined;

  let rootEl: ReturnType<typeof parseXmlPart>;
  try {
    rootEl = parseXmlPart(themePartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: themePartPath },
    });
    return undefined;
  }

  if (
    (rootEl.namespaceURI ?? "") !== A_NS ||
    (rootEl.localName ?? rootEl.tagName) !== "theme"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: themePartPath },
    });
    return undefined;
  }

  const themeCursor = XmlElementCursor.createRoot(
    rootEl,
    themePartPath,
    "/a:theme",
  );
  const rawName = themeCursor.attr("name");
  const themeName =
    rawName !== undefined && rawName.trim() !== "" ? rawName : "Office Theme";

  const themeElements = themeCursor.child(A_NS, "themeElements");
  let colors: Record<NativeColorSlot, string> = { ...officeFallbackColors };
  let fonts: { readonly major?: string; readonly minor?: string } = {};
  let formatScheme: FormatScheme | undefined;

  if (themeElements !== undefined) {
    const clrScheme = themeElements.child(A_NS, "clrScheme");
    if (clrScheme !== undefined) {
      colors = readClrScheme(clrScheme, themePartPath, undefined, diagnostics);
    }
    const fontScheme = themeElements.child(A_NS, "fontScheme");
    if (fontScheme !== undefined) {
      fonts = readFontScheme(fontScheme);
    }
    const fmtScheme = themeElements.child(A_NS, "fmtScheme");
    if (fmtScheme !== undefined) {
      formatScheme = readFmtScheme(
        fmtScheme,
        themePartPath,
        undefined,
        diagnostics,
      );
    }
  }

  const objectDefaults = themeCursor.child(A_NS, "objectDefaults");
  if (objectDefaults !== undefined) {
    for (const child of objectDefaults.remainingElementChildren()) {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { part: themePartPath, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `theme-object-defaults:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }
  const extraClrSchemeLst = themeCursor.child(A_NS, "extraClrSchemeLst");
  if (extraClrSchemeLst !== undefined) {
    for (const child of extraClrSchemeLst.remainingElementChildren()) {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { part: themePartPath, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `theme-extra-color-scheme:${child.localName}`,
      });
      child.suppressUnconsumed();
    }
  }

  for (const item of themeCursor.collectUnconsumed()) {
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: item.kind,
      location: item.location,
      namespaceUri: item.namespaceUri,
      localName: item.localName,
    });
  }

  const theme = defineTheme({
    name: themeName,
    colors,
    fonts,
    ...(formatScheme === undefined ? {} : { formatScheme }),
  });
  themesByPartPath.set(themePartPath, theme);
  return theme;
}

/** Reads an optional `<a:themeOverride>` part referenced by a SlideLayout or Slide. */
export function readThemeOverridePart(
  pkg: PptxPackage,
  overridePartPath: string,
  slideIndex: number | undefined,
  diagnostics: ImportDiagnostic[],
): ThemeOverride | undefined {
  const xmlText = pkg.readTextPart(overridePartPath);
  if (xmlText === undefined) return undefined;

  let rootEl: ReturnType<typeof parseXmlPart>;
  try {
    rootEl = parseXmlPart(overridePartPath, xmlText);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: overridePartPath,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
    });
    return undefined;
  }

  if (
    (rootEl.namespaceURI ?? "") !== A_NS ||
    (rootEl.localName ?? rootEl.tagName) !== "themeOverride"
  ) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: {
        part: overridePartPath,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
    });
    return undefined;
  }

  const cursor = XmlElementCursor.createRoot(
    rootEl,
    overridePartPath,
    "/a:themeOverride",
  );
  const clrScheme = cursor.child(A_NS, "clrScheme");
  const colors =
    clrScheme === undefined
      ? undefined
      : readClrScheme(clrScheme, overridePartPath, slideIndex, diagnostics);
  const fontScheme = cursor.child(A_NS, "fontScheme");
  const fonts =
    fontScheme === undefined ? undefined : readFontScheme(fontScheme);
  const fmtScheme = cursor.child(A_NS, "fmtScheme");
  const formatScheme =
    fmtScheme === undefined
      ? undefined
      : readFmtScheme(fmtScheme, overridePartPath, slideIndex, diagnostics);

  for (const item of cursor.collectUnconsumed()) {
    diagnostics.push({
      code: "unconsumed-xml",
      category: "unsupported",
      disposition: "omitted",
      kind: item.kind,
      location: {
        ...item.location,
        ...(slideIndex === undefined ? {} : { slideIndex }),
      },
      namespaceUri: item.namespaceUri,
      localName: item.localName,
    });
  }

  if (
    colors === undefined &&
    fonts === undefined &&
    formatScheme === undefined
  ) {
    return undefined;
  }
  return {
    ...(colors === undefined ? {} : { colors }),
    ...(fonts === undefined ? {} : { fonts }),
    ...(formatScheme === undefined ? {} : { formatScheme }),
  };
}

export function isNonDefaultColorMap(map: ColorMap): boolean {
  return !colorMapsEqual(map, defaultColorMap);
}

export { P_NS };
