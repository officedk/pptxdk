import {
  CapabilityError,
  linearGradientSchema,
  radialGradientSchema,
  shadowSchema,
  visualStyleSchema,
  type CapabilityResource,
  type ColorValue,
  type GradientStop,
  type LinearGradient,
  type NativeColorSlot,
  type PaintOutput,
  type RadialGradient,
  type ResolvedColor,
  type ThemeContext,
  type VisualStyle as CoreVisualStyle,
} from "@pptxdk/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@pptxdk/emit";
import {
  A_NS,
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  OffsetEmuFragmentSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  P_NS,
  R_NS,
  TransformAttributesFragmentSchema,
  defineCanvasMechanism,
  emuToPx,
  resolvePartPath,
  type ImportDiagnostic,
  type ImportLocation,
  type PptxPackage,
  type ShapeTreeImportContext,
  type ShapeTreeImportOutcome,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  decodeImageToPngBytes,
  readImage,
  type LoadedImage,
} from "../image/image-decode.ts";
import {
  evaluateGeometry,
  readCustomGeometryXml,
  readPresetGeometryXml,
  resolveGeometry,
  resolvedGeometryXml,
  shadeHex,
  type Geometry,
  type PathCommand,
  type PresetGeometryName,
  type PresetPathFillMode,
  type ResolvedGeometry,
} from "./geometry.ts";

const imageRelationship =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

export function linearGradient(
  angle: number,
  ...stops: readonly (GradientStop | ColorValue)[]
): LinearGradient {
  return {
    type: "linear-gradient",
    angle,
    stops: stops.map((stop, index) =>
      typeof stop === "object" && stop !== null && "position" in stop
        ? stop
        : {
            color: stop,
            position: stops.length === 1 ? 0 : index / (stops.length - 1),
          },
    ),
  };
}

export function radialGradient(
  ...stops: readonly (GradientStop | ColorValue)[]
): RadialGradient {
  return {
    type: "radial-gradient",
    stops: stops.map((stop, index) =>
      typeof stop === "object" && stop !== null && "position" in stop
        ? stop
        : {
            color: stop,
            position: stops.length === 1 ? 0 : index / (stops.length - 1),
          },
    ),
  };
}

export type ResolvedGradientStop = {
  readonly position: number;
  readonly color: ResolvedColor;
};

export type ResolvedGradientSpec =
  | {
      readonly kind: "linear-gradient";
      readonly angle: number;
      readonly stops: readonly ResolvedGradientStop[];
    }
  | {
      readonly kind: "radial-gradient";
      readonly stops: readonly ResolvedGradientStop[];
    };

export type ResolvedFill =
  | { readonly kind: "none" }
  | { readonly kind: "solid"; readonly color: ResolvedColor }
  | {
      readonly kind: "linear-gradient";
      readonly angle: number;
      readonly stops: readonly ResolvedGradientStop[];
    }
  | {
      readonly kind: "radial-gradient";
      readonly stops: readonly ResolvedGradientStop[];
    }
  | {
      readonly kind: "image";
      readonly image: LoadedImage;
      readonly opacity?: number;
    };

export type ResolvedBorder = {
  readonly color: ResolvedColor;
  readonly width: number;
  readonly style: "solid" | "dashed" | "dotted";
  /** Undefined when author omitted startArrow (omits `<a:headEnd>` in OOXML), otherwise explicit ArrowType. */
  readonly startArrow?: ArrowType;
  /** Undefined when author omitted endArrow (omits `<a:tailEnd>` in OOXML), otherwise explicit ArrowType. */
  readonly endArrow?: ArrowType;
};

export type ResolvedShadow = {
  readonly color: ResolvedColor;
  readonly blur: number;
  readonly distance: number;
  readonly angle: number;
};

export type VisualWarning = { readonly code: string; readonly message: string };
export type ShapeBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};
export type VisualStyleRecord = Readonly<Record<string, unknown>>;

/** Closed resolved visual model produced in resolve() and exhaustively consumed by emit() and paint(). */
export type ResolvedVisualModel = {
  readonly box: ShapeBox;
  readonly geometry: ResolvedGeometry;
  readonly geometryName: PresetGeometryName | "custom";
  readonly fill: ResolvedFill;
  readonly border?: ResolvedBorder;
  readonly opacity?: number;
  readonly boxShadow?: ResolvedShadow;
  readonly softEdge?: number;
  readonly rotate: number;
  readonly rotationUnits?: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly hasVisibleStyle: boolean;
  readonly layoutOnlyWhenInvisible: boolean;
  readonly warnings: readonly VisualWarning[];
};

export const arrowTypes = [
  "none",
  "triangle",
  "stealth",
  "arrow",
  "diamond",
  "oval",
] as const;

export const arrowTypeSchema = z.enum(arrowTypes);
export type ArrowType = z.infer<typeof arrowTypeSchema>;

export const lineEndStyleSchema = z
  .strictObject({
    startArrow: arrowTypeSchema.optional(),
    endArrow: arrowTypeSchema.optional(),
  })
  .readonly();
export type LineEndStyle = z.infer<typeof lineEndStyleSchema>;

export const lineEndStyleKeys = ["startArrow", "endArrow"] as const;

export type ShapeVisualStyle = CoreVisualStyle & LineEndStyle;

export const visualStyleKeys = [
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "borderRadius",
  "opacity",
  "boxShadow",
  "softEdge",
  "rotate",
] as const;

export type ShapePropertiesOptions = {
  readonly theme: ThemeContext;
  readonly style: VisualStyleRecord;
  readonly box: ShapeBox;
  readonly geometry?: Geometry;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** Legacy top-level fill property of shape capability. */
  readonly fill?: ColorValue;
  readonly image?: boolean;
  /** Layout containers without visible styling do not require emitting native transforms. */
  readonly layoutOnlyWhenInvisible?: boolean;
};

export type ShapePropertiesResult = {
  readonly xml: string;
  readonly opacity?: number;
  readonly warnings: readonly VisualWarning[];
  readonly hasVisibleStyle: boolean;
  readonly geometryName: PresetGeometryName | "custom";
  readonly geometry: ResolvedGeometry;
};

function invalid(code: string, message: string): never {
  throw new CapabilityError(code, message);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function finite(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    invalid("invalid-visual-style", `${label} must be a finite number`);
  }
  return value;
}

function emu(
  value: number,
  label: string,
  allowZero = true,
  signed = false,
): number {
  const converted = pxToEmu(value);
  if (
    !Number.isSafeInteger(converted) ||
    (!signed && converted < (allowZero ? 0 : 1)) ||
    (!signed && converted > 2_147_483_647)
  ) {
    invalid(
      "invalid-visual-style",
      `${label} must be within writable DrawingML range after conversion`,
    );
  }
  return converted;
}

function lineWidthEmu(value: number): number {
  const converted = pxToEmu(value);
  if (
    !Number.isSafeInteger(converted) ||
    converted < 1 ||
    converted > 20_116_800
  ) {
    invalid(
      "invalid-visual-style",
      "borderWidth must be within the DrawingML line width range of 1 to 20116800 EMU after conversion",
    );
  }
  return converted;
}

function withOpacity(
  color: ResolvedColor,
  opacity: number | undefined,
): ResolvedColor {
  return opacity === undefined
    ? color
    : { ...color, alpha: (color.alpha ?? 1) * opacity };
}

function roundSvgNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

export function resolveShadow(
  theme: ThemeContext,
  value: unknown,
  label: string,
): ResolvedShadow {
  if (!isRecord(value))
    invalid("invalid-shadow", `${label} must be a shadow object`);
  const parsed = shadowSchema.safeParse(value);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.code === "unrecognized_keys" && issue.keys[0] !== undefined) {
      invalid(
        "invalid-shadow",
        `${label} contains unsupported parameter ${issue.keys[0]}`,
      );
    }
    const key = String(issue?.path[0] ?? "value");
    if (key === "blur" || key === "distance") {
      if (typeof value[key] !== "number" || !Number.isFinite(value[key])) {
        finite(value[key], `${label}.${key}`);
      }
      invalid(
        "invalid-shadow",
        `${label}.blur and distance must be non-negative lengths`,
      );
    }
    if (key === "angle") {
      finite(value[key], `${label}.angle`);
    }
    invalid("invalid-shadow", issue?.message ?? `${label} is invalid`);
  }
  const blur = parsed.data.blur ?? 4;
  const distance = parsed.data.distance ?? 0;
  const angle = parsed.data.angle ?? 90;
  emu(blur, `${label}.blur`);
  emu(distance, `${label}.distance`);
  const colorValue: ColorValue = parsed.data.color ?? {
    color: "#000000",
    alpha: 0.35,
  };
  return { color: theme.resolveColor(colorValue), blur, distance, angle };
}

export function textShadowXml(theme: ThemeContext, value: unknown): string {
  if (value === undefined || value === "none") return "";
  const shadow = resolveShadow(theme, value, "textShadow");
  const direction = Math.round((((shadow.angle % 360) + 360) % 360) * 60000);
  return `<a:effectLst><a:outerShdw blurRad="${emu(shadow.blur, "textShadow.blur")}" dist="${emu(shadow.distance, "textShadow.distance")}" dir="${direction}" algn="ctr" rotWithShape="0">${resolvedColorXml(shadow.color)}</a:outerShdw></a:effectLst>`;
}

const gradientSchema = z
  .union([linearGradientSchema, radialGradientSchema])
  .refine((value) => value.stops.length >= 2);

/** Resolves and validates a linear or radial gradient declaration against a ThemeContext. */
export function resolveGradientSpec(
  theme: ThemeContext,
  image: unknown,
  opacity: number | undefined,
): ResolvedGradientSpec {
  const parsed = gradientSchema.safeParse(image);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.path[0] === "angle" && isRecord(image)) {
      finite(image["angle"], "backgroundImage.angle");
    }
    if (
      issue?.path[0] === "stops" &&
      isRecord(image) &&
      Array.isArray(image["stops"])
    ) {
      const idx = typeof issue.path[1] === "number" ? issue.path[1] : undefined;
      if (idx !== undefined && issue.path[2] === "position") {
        const stop = image["stops"][idx];
        if (isRecord(stop)) {
          finite(stop["position"], `backgroundImage.stops[${idx}].position`);
        }
      }
    }
    invalid(
      "invalid-gradient",
      "backgroundImage must be a linear or radial gradient with at least two color stops",
    );
  }
  const stops: ResolvedGradientStop[] = parsed.data.stops.map((stop, index) => {
    if (stop.position < 0 || stop.position > 1) {
      invalid(
        "invalid-gradient",
        `position for gradient stop ${index + 1} must be between 0 and 1`,
      );
    }
    const color = withOpacity(theme.resolveColor(stop.color), opacity);
    return { position: stop.position, color };
  });
  if (parsed.data.type === "radial-gradient") {
    return { kind: "radial-gradient", stops };
  }
  return { kind: "linear-gradient", angle: parsed.data.angle, stops };
}

/** Emits DrawingML `<a:gradFill>` from a closed resolved gradient specification. */
export function gradientFillXml(gradient: ResolvedGradientSpec): string {
  const stops = gradient.stops
    .map(
      (stop) =>
        `<a:gs pos="${Math.round(stop.position * 100000)}">${resolvedColorXml(stop.color)}</a:gs>`,
    )
    .join("");
  switch (gradient.kind) {
    case "radial-gradient":
      return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
    case "linear-gradient": {
      const drawingAngle = Math.round(
        ((((gradient.angle - 90) % 360) + 360) % 360) * 60000,
      );
      return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${drawingAngle}" scaled="0"/></a:gradFill>`;
    }
    default: {
      const exhaustive: never = gradient;
      return exhaustive;
    }
  }
}

function gradientStopsSvg(stops: readonly ResolvedGradientStop[]): string {
  return [...stops]
    .sort((a, b) => a.position - b.position)
    .map(
      (stop) =>
        `<stop offset="${roundSvgNumber(stop.position * 100)}%" stop-color="#${stop.color.hex}" stop-opacity="${roundSvgNumber(stop.color.alpha ?? 1)}"/>`,
    )
    .join("");
}

/** Generates deterministic SVG `<linearGradient>` or `<radialGradient>` definition markup. */
export function gradientDefSvg(
  id: string,
  gradient: ResolvedGradientSpec,
): string {
  const stops = gradientStopsSvg(gradient.stops);
  switch (gradient.kind) {
    case "linear-gradient": {
      const rad = (gradient.angle * Math.PI) / 180;
      const dx = Math.sin(rad);
      const dy = -Math.cos(rad);
      const halfSpan = (Math.abs(dx) + Math.abs(dy)) / 2;
      const x1 = roundSvgNumber(0.5 - dx * halfSpan);
      const y1 = roundSvgNumber(0.5 - dy * halfSpan);
      const x2 = roundSvgNumber(0.5 + dx * halfSpan);
      const y2 = roundSvgNumber(0.5 + dy * halfSpan);
      return `<linearGradient id="${id}" gradientUnits="objectBoundingBox" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient>`;
    }
    case "radial-gradient":
      return `<radialGradient id="${id}" gradientUnits="objectBoundingBox" cx="0.5" cy="0.5" r="0.707107">${stops}</radialGradient>`;
    default: {
      const exhaustive: never = gradient;
      return exhaustive;
    }
  }
}

function rotationUnits(style: VisualStyleRecord): {
  readonly degrees: number;
  readonly units: number | undefined;
} {
  if (style["rotate"] === undefined) {
    return { degrees: 0, units: undefined };
  }
  const degrees = finite(style["rotate"], "rotate");
  const units = Math.round((((degrees % 360) + 360) % 360) * 60000);
  if (!Number.isSafeInteger(units) || units < 0 || units > 21_600_000) {
    invalid(
      "invalid-visual-style",
      "rotate exceeds writable DrawingML range after conversion",
    );
  }
  return { degrees, units };
}

function resolveVisualFill(
  theme: ThemeContext,
  style: VisualStyleRecord,
  explicitFill: ColorValue | undefined,
  opacity: number | undefined,
  isImageObject: boolean,
): ResolvedFill {
  if (isImageObject) {
    return { kind: "none" };
  }
  if (
    explicitFill !== undefined &&
    (style["backgroundColor"] !== undefined ||
      style["backgroundImage"] !== undefined)
  ) {
    invalid(
      "unsupported-style-combination",
      "Shape.fill cannot be set together with style.backgroundColor or backgroundImage",
    );
  }
  if (explicitFill !== undefined) {
    return {
      kind: "solid",
      color: withOpacity(theme.resolveColor(explicitFill), opacity),
    };
  }
  const background = style["backgroundColor"];
  const image = style["backgroundImage"];
  if (background !== undefined && image !== undefined) {
    invalid(
      "unsupported-style-combination",
      "backgroundColor and backgroundImage cannot be set simultaneously",
    );
  }
  if (image !== undefined) {
    if (isRecord(image) && Object.hasOwn(image, "image")) {
      const unknown = Object.keys(image).find((key) => key !== "image");
      if (unknown !== undefined) {
        invalid(
          "invalid-visual-style",
          `backgroundImage contains unsupported parameter ${unknown}`,
        );
      }
      const loaded = readImage(image["image"]);
      return {
        kind: "image",
        image: loaded,
        ...(opacity === undefined ? {} : { opacity }),
      };
    }
    return resolveGradientSpec(theme, image, opacity);
  }
  if (background === undefined) {
    return { kind: "none" };
  }
  return {
    kind: "solid",
    color: withOpacity(theme.resolveColor(background as ColorValue), opacity),
  };
}

function resolveVisualBorder(
  theme: ThemeContext,
  style: VisualStyleRecord,
  opacity: number | undefined,
): ResolvedBorder | undefined {
  const color = style["borderColor"];
  const width =
    style["borderWidth"] === undefined
      ? 1
      : finite(style["borderWidth"], "borderWidth");
  if (width < 0) {
    invalid(
      "invalid-visual-style",
      "borderWidth must be a non-negative number",
    );
  }
  const parsedBorder = visualStyleSchema.safeParse(
    style["borderStyle"] === undefined
      ? {}
      : { borderStyle: style["borderStyle"] },
  );
  if (!parsedBorder.success) {
    invalid(
      "invalid-visual-style",
      "borderStyle must be solid, dashed, or dotted",
    );
  }
  const parsedLineEnds = lineEndStyleSchema.safeParse({
    ...(style["startArrow"] === undefined
      ? {}
      : { startArrow: style["startArrow"] }),
    ...(style["endArrow"] === undefined ? {} : { endArrow: style["endArrow"] }),
  });
  if (!parsedLineEnds.success) {
    const field = String(
      parsedLineEnds.error.issues[0]?.path[0] ?? "startArrow",
    );
    invalid(
      "invalid-visual-style",
      `${field} must be none, triangle, stealth, arrow, diamond, or oval`,
    );
  }
  if (color === undefined || width === 0) {
    return undefined;
  }
  lineWidthEmu(width);
  return {
    color: withOpacity(theme.resolveColor(color as ColorValue), opacity),
    width,
    style: parsedBorder.data.borderStyle ?? "solid",
    ...(parsedLineEnds.data.startArrow === undefined
      ? {}
      : { startArrow: parsedLineEnds.data.startArrow }),
    ...(parsedLineEnds.data.endArrow === undefined
      ? {}
      : { endArrow: parsedLineEnds.data.endArrow }),
  };
}

/** Determines whether a Box style specifies visible fill, border, or shadow. */
export function hasVisibleBoxStyle(
  style: Readonly<Record<string, unknown>> | undefined,
): boolean {
  if (style === undefined || typeof style !== "object" || style === null)
    return false;
  const borderWidth =
    style["borderWidth"] === undefined ? 1 : (style["borderWidth"] as number);
  return (
    style["backgroundColor"] !== undefined ||
    style["backgroundImage"] !== undefined ||
    (style["borderColor"] !== undefined && borderWidth > 0) ||
    style["boxShadow"] !== undefined
  );
}

/**
 * Validates and resolves visual styles, geometry, fill, border, shadow, softEdge, and transforms
 * into a closed strongly-typed model consumed exhaustively by emitVisualModel and paintVisualModel.
 */
export function resolveVisualModel(
  options: ShapePropertiesOptions,
): ResolvedVisualModel {
  const { theme, style, box } = options;
  const warnings: VisualWarning[] = [];

  for (const [name, value] of [
    ["flipH", options.flipH],
    ["flipV", options.flipV],
  ] as const) {
    if (value !== undefined && typeof value !== "boolean") {
      invalid("invalid-visual-style", `${name} must be a boolean`);
    }
  }

  const opacity =
    style["opacity"] === undefined
      ? undefined
      : finite(style["opacity"], "opacity");
  if (opacity !== undefined && (opacity < 0 || opacity > 1)) {
    invalid("invalid-visual-style", "opacity must be a number between 0 and 1");
  }
  if (
    options.image &&
    (style["backgroundColor"] !== undefined ||
      style["backgroundImage"] !== undefined ||
      style["borderRadius"] !== undefined)
  ) {
    invalid(
      "unsupported-style-combination",
      "Image does not support background fill or borderRadius",
    );
  }
  const radius =
    style["borderRadius"] === undefined
      ? undefined
      : finite(style["borderRadius"], "borderRadius");
  if (radius !== undefined && radius < 0) {
    invalid(
      "invalid-visual-style",
      "borderRadius must be a non-negative length",
    );
  }

  const fill = resolveVisualFill(
    theme,
    style,
    options.fill,
    opacity,
    Boolean(options.image),
  );
  const border = resolveVisualBorder(theme, style, opacity);
  const boxShadow =
    style["boxShadow"] === undefined
      ? undefined
      : resolveShadow(theme, style["boxShadow"], "boxShadow");

  let softEdge: number | undefined;
  if (style["softEdge"] !== undefined) {
    const rawSoft = finite(style["softEdge"], "softEdge");
    if (rawSoft < 0) {
      invalid("invalid-visual-style", "softEdge must be a non-negative length");
    }
    if (rawSoft > 0) {
      emu(rawSoft, "softEdge", false);
      softEdge = rawSoft;
    }
  }

  const hasVisibleStyle =
    (options.fill !== undefined && !options.image) || hasVisibleBoxStyle(style);
  const layoutOnlyWhenInvisible = Boolean(options.layoutOnlyWhenInvisible);
  const resolvedGeometry = resolveGeometry({
    geometry: options.geometry,
    borderRadius:
      layoutOnlyWhenInvisible && !hasVisibleStyle ? undefined : radius,
    box,
  });
  warnings.push(...resolvedGeometry.warnings);
  const rotation = rotationUnits(style);

  return {
    box,
    geometry: resolvedGeometry.geometry,
    geometryName: resolvedGeometry.name,
    fill,
    ...(border === undefined ? {} : { border }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(boxShadow === undefined ? {} : { boxShadow }),
    ...(softEdge === undefined ? {} : { softEdge }),
    rotate: rotation.degrees,
    ...(rotation.units === undefined ? {} : { rotationUnits: rotation.units }),
    flipH: Boolean(options.flipH),
    flipV: Boolean(options.flipV),
    hasVisibleStyle,
    layoutOnlyWhenInvisible,
    warnings,
  };
}

function emitFillXml(
  fill: ResolvedFill,
  registerResource: ((resource: CapabilityResource) => string) | undefined,
): string {
  switch (fill.kind) {
    case "none":
      return "<a:noFill/>";
    case "solid":
      return solidFillMechanism.emit(fill, undefined).xml;
    case "linear-gradient":
    case "radial-gradient":
      return gradientFillMechanism.emit(fill, undefined).xml;
    case "image":
      return imageFillMechanism.emit(fill, registerResource).xml;
    default: {
      const exhaustive: never = fill;
      return exhaustive;
    }
  }
}

function emitBorderXml(border: ResolvedBorder | undefined): string {
  if (border === undefined) {
    return "<a:ln><a:noFill/></a:ln>";
  }
  let dash: string | undefined;
  switch (border.style) {
    case "solid":
      dash = undefined;
      break;
    case "dashed":
      dash = "dash";
      break;
    case "dotted":
      dash = "sysDot";
      break;
    default: {
      const exhaustive: never = border.style;
      return exhaustive;
    }
  }
  const headEnd =
    border.startArrow === undefined
      ? ""
      : `<a:headEnd type="${border.startArrow}"/>`;
  const tailEnd =
    border.endArrow === undefined
      ? ""
      : `<a:tailEnd type="${border.endArrow}"/>`;
  return `<a:ln w="${lineWidthEmu(border.width)}"><a:solidFill>${resolvedColorXml(border.color)}</a:solidFill>${dash === undefined ? "" : `<a:prstDash val="${dash}"/>`}${headEnd}${tailEnd}</a:ln>`;
}

function emitEffectsXml(model: ResolvedVisualModel): string {
  const items: string[] = [];
  if (model.boxShadow !== undefined) {
    const shadow = model.boxShadow;
    const dir = Math.round((((shadow.angle % 360) + 360) % 360) * 60000);
    items.push(
      `<a:outerShdw blurRad="${emu(shadow.blur, "boxShadow.blur")}" dist="${emu(shadow.distance, "boxShadow.distance")}" dir="${dir}" algn="ctr" rotWithShape="0">${resolvedColorXml(shadow.color)}</a:outerShdw>`,
    );
  }
  if (model.softEdge !== undefined && model.softEdge > 0) {
    items.push(`<a:softEdge rad="${emu(model.softEdge, "softEdge", false)}"/>`);
  }
  return items.length === 0
    ? ""
    : `<a:effectLst>${items.join("")}</a:effectLst>`;
}

/** Serializes a ResolvedVisualModel into PresentationML `<p:spPr>` XML. */
export function emitVisualModel(
  model: ResolvedVisualModel,
  registerResource?: (resource: CapabilityResource) => string,
  options?: { readonly omitFill?: boolean },
): ShapePropertiesResult {
  const { box } = model;
  const transform =
    model.layoutOnlyWhenInvisible && !model.hasVisibleStyle
      ? ""
      : `<a:xfrm${model.rotationUnits === undefined || model.rotationUnits === 0 ? "" : ` rot="${model.rotationUnits}"`}${model.flipH ? ' flipH="1"' : ""}${model.flipV ? ' flipV="1"' : ""}><a:off x="${emu(box.x, "x", true, true)}" y="${emu(box.y, "y", true, true)}"/><a:ext cx="${emu(box.width, "width", false)}" cy="${emu(box.height, "height", false)}"/></a:xfrm>`;
  const geometryXmlText = resolvedGeometryXml(model.geometry);
  const visualFill = options?.omitFill
    ? ""
    : emitFillXml(model.fill, registerResource);
  const line = emitBorderXml(model.border);
  const effects = emitEffectsXml(model);
  return {
    xml: `<p:spPr>${transform}${geometryXmlText}${visualFill}${line}${effects}</p:spPr>`,
    ...(model.opacity === undefined ? {} : { opacity: model.opacity }),
    warnings: model.warnings,
    hasVisibleStyle: model.hasVisibleStyle,
    geometryName: model.geometryName,
    geometry: model.geometry,
  };
}

/** Returns the SVG transform string (flip around box center, then rotate around box center). */
export function visualSvgTransform(model: ResolvedVisualModel): string {
  const cx = model.box.x + model.box.width / 2;
  const cy = model.box.y + model.box.height / 2;
  return `translate(${roundSvgNumber(cx)} ${roundSvgNumber(cy)}) rotate(${roundSvgNumber(model.rotate)}) scale(${model.flipH ? -1 : 1} ${model.flipV ? -1 : 1}) translate(${roundSvgNumber(-cx)} ${roundSvgNumber(-cy)})`;
}

type Point2 = { readonly x: number; readonly y: number };

function normalizeVec(dx: number, dy: number): Point2 | undefined {
  const len = Math.hypot(dx, dy);
  if (len <= 1e-9) return undefined;
  return { x: dx / len, y: dy / len };
}

function splitSubpaths(
  commands: readonly PathCommand[],
): readonly (readonly PathCommand[])[] {
  const result: PathCommand[][] = [];
  let current: PathCommand[] | undefined;
  let start: Point2 = { x: 0, y: 0 };
  for (const command of commands) {
    if (command.op === "M") {
      current = [command];
      result.push(current);
      start = { x: command.x, y: command.y };
      continue;
    }
    if (current === undefined) {
      current = [{ op: "M", x: start.x, y: start.y }];
      result.push(current);
    }
    current.push(command);
    if (command.op === "Z") {
      current = undefined;
    }
  }
  return result;
}

function arrowShortenLength(arrow: ArrowType, strokeWidth: number): number {
  const w = Math.max(1, strokeWidth);
  const length = Math.max(6, w * 3.5);
  switch (arrow) {
    case "none":
      return 0;
    case "triangle":
      return 0.85 * length;
    case "stealth":
      return 0.65 * length;
    case "arrow":
      return 0.5 * w;
    case "diamond":
    case "oval":
      return 0;
    default: {
      const exhaustive: never = arrow;
      return exhaustive;
    }
  }
}

function arrowHeadSvg(
  arrow: ArrowType,
  tip: Point2,
  dir: Point2,
  strokeHex: string,
  strokeAlpha: number,
  strokeWidth: number,
): string {
  if (arrow === "none") return "";
  const w = Math.max(1, strokeWidth);
  const length = Math.max(6, w * 3.5);
  const halfWidth = Math.max(3, w * 1.75);
  const deg = roundSvgNumber((Math.atan2(dir.y, dir.x) * 180) / Math.PI);
  const tx = roundSvgNumber(tip.x);
  const ty = roundSvgNumber(tip.y);
  const alpha = roundSvgNumber(strokeAlpha);
  const transform = `translate(${tx} ${ty}) rotate(${deg})`;

  switch (arrow) {
    case "triangle":
      return `<path d="M 0 0 L ${roundSvgNumber(-length)} ${roundSvgNumber(-halfWidth)} L ${roundSvgNumber(-length)} ${roundSvgNumber(halfWidth)} Z" fill="#${strokeHex}" fill-opacity="${alpha}" stroke="none" transform="${transform}"/>`;
    case "stealth":
      return `<path d="M 0 0 L ${roundSvgNumber(-length)} ${roundSvgNumber(-halfWidth)} L ${roundSvgNumber(-0.65 * length)} 0 L ${roundSvgNumber(-length)} ${roundSvgNumber(halfWidth)} Z" fill="#${strokeHex}" fill-opacity="${alpha}" stroke="none" transform="${transform}"/>`;
    case "arrow":
      return `<path d="M ${roundSvgNumber(-length)} ${roundSvgNumber(-halfWidth)} L 0 0 L ${roundSvgNumber(-length)} ${roundSvgNumber(halfWidth)}" fill="none" stroke="#${strokeHex}" stroke-opacity="${alpha}" stroke-width="${roundSvgNumber(w)}" stroke-linecap="round" stroke-linejoin="round" transform="${transform}"/>`;
    case "diamond":
      return `<path d="M ${roundSvgNumber(length * 0.5)} 0 L 0 ${roundSvgNumber(-halfWidth)} L ${roundSvgNumber(-length * 0.5)} 0 L 0 ${roundSvgNumber(halfWidth)} Z" fill="#${strokeHex}" fill-opacity="${alpha}" stroke="none" transform="${transform}"/>`;
    case "oval":
      return `<ellipse cx="0" cy="0" rx="${roundSvgNumber(length * 0.45)}" ry="${roundSvgNumber(halfWidth * 0.8)}" fill="#${strokeHex}" fill-opacity="${alpha}" stroke="none" transform="${transform}"/>`;
    default: {
      const exhaustive: never = arrow;
      return exhaustive;
    }
  }
}

function commandsToSvgPath(commands: readonly PathCommand[]): string {
  return commands
    .map((command) => {
      switch (command.op) {
        case "Z":
          return "Z";
        case "M":
        case "L":
          return `${command.op}${roundSvgNumber(command.x)} ${roundSvgNumber(command.y)}`;
        case "Q":
          return `Q${roundSvgNumber(command.x1)} ${roundSvgNumber(command.y1)} ${roundSvgNumber(command.x)} ${roundSvgNumber(command.y)}`;
        case "C":
          return `C${roundSvgNumber(command.x1)} ${roundSvgNumber(command.y1)} ${roundSvgNumber(command.x2)} ${roundSvgNumber(command.y2)} ${roundSvgNumber(command.x)} ${roundSvgNumber(command.y)}`;
        default: {
          const exhaustive: never = command;
          return exhaustive;
        }
      }
    })
    .join(" ");
}

function extractOpenSubpathEnds(subpath: readonly PathCommand[]):
  | {
      readonly startPoint: Point2;
      readonly startTangent: Point2;
      readonly firstSegLen: number;
      readonly endPoint: Point2;
      readonly endTangent: Point2;
      readonly lastSegLen: number;
      readonly segCount: number;
    }
  | undefined {
  const first = subpath[0];
  const last = subpath.at(-1);
  if (first === undefined || first.op !== "M" || last === undefined) {
    return undefined;
  }
  if (last.op === "Z" || subpath.length < 2) {
    return undefined;
  }

  const startPoint: Point2 = { x: first.x, y: first.y };
  let startTangent: Point2 | undefined;
  let firstSegLen = 0;

  for (let i = 1; i < subpath.length; i += 1) {
    const cmd = subpath[i]!;
    if (cmd.op === "L") {
      if (i === 1)
        firstSegLen = Math.hypot(cmd.x - startPoint.x, cmd.y - startPoint.y);
      startTangent = normalizeVec(cmd.x - startPoint.x, cmd.y - startPoint.y);
      if (startTangent !== undefined) break;
    } else if (cmd.op === "Q") {
      if (i === 1)
        firstSegLen = Math.hypot(cmd.x - startPoint.x, cmd.y - startPoint.y);
      startTangent =
        normalizeVec(cmd.x1 - startPoint.x, cmd.y1 - startPoint.y) ??
        normalizeVec(cmd.x - startPoint.x, cmd.y - startPoint.y);
      if (startTangent !== undefined) break;
    } else if (cmd.op === "C") {
      if (i === 1)
        firstSegLen = Math.hypot(cmd.x - startPoint.x, cmd.y - startPoint.y);
      startTangent =
        normalizeVec(cmd.x1 - startPoint.x, cmd.y1 - startPoint.y) ??
        normalizeVec(cmd.x2 - startPoint.x, cmd.y2 - startPoint.y) ??
        normalizeVec(cmd.x - startPoint.x, cmd.y - startPoint.y);
      if (startTangent !== undefined) break;
    }
  }

  if (last.op === "M") return undefined;
  const endPoint: Point2 = { x: last.x, y: last.y };
  let endTangent: Point2 | undefined;
  let lastSegLen = 0;

  for (let i = subpath.length - 1; i >= 1; i -= 1) {
    const cmd = subpath[i]!;
    const prevCmd = subpath[i - 1]!;
    if (prevCmd.op === "Z") continue;
    const prevPoint: Point2 = { x: prevCmd.x, y: prevCmd.y };
    if (cmd.op === "L") {
      if (i === subpath.length - 1) {
        lastSegLen = Math.hypot(
          endPoint.x - prevPoint.x,
          endPoint.y - prevPoint.y,
        );
      }
      endTangent = normalizeVec(
        endPoint.x - prevPoint.x,
        endPoint.y - prevPoint.y,
      );
      if (endTangent !== undefined) break;
    } else if (cmd.op === "Q") {
      if (i === subpath.length - 1) {
        lastSegLen = Math.hypot(
          endPoint.x - prevPoint.x,
          endPoint.y - prevPoint.y,
        );
      }
      endTangent =
        normalizeVec(endPoint.x - cmd.x1, endPoint.y - cmd.y1) ??
        normalizeVec(endPoint.x - prevPoint.x, endPoint.y - prevPoint.y);
      if (endTangent !== undefined) break;
    } else if (cmd.op === "C") {
      if (i === subpath.length - 1) {
        lastSegLen = Math.hypot(
          endPoint.x - prevPoint.x,
          endPoint.y - prevPoint.y,
        );
      }
      endTangent =
        normalizeVec(endPoint.x - cmd.x2, endPoint.y - cmd.y2) ??
        normalizeVec(endPoint.x - cmd.x1, endPoint.y - cmd.y1) ??
        normalizeVec(endPoint.x - prevPoint.x, endPoint.y - prevPoint.y);
      if (endTangent !== undefined) break;
    }
  }

  if (startTangent === undefined || endTangent === undefined) {
    return undefined;
  }

  return {
    startPoint,
    startTangent,
    firstSegLen,
    endPoint,
    endTangent,
    lastSegLen,
    segCount: subpath.length - 1,
  };
}

function shortenOpenSubpath(
  subpath: readonly PathCommand[],
  startShorten: number,
  endShorten: number,
  ends: NonNullable<ReturnType<typeof extractOpenSubpathEnds>>,
): readonly PathCommand[] {
  const canShortenStart =
    startShorten > 0 &&
    (ends.segCount === 1
      ? ends.firstSegLen > (startShorten + endShorten) * 1.5
      : ends.firstSegLen > startShorten * 1.5);
  const canShortenEnd =
    endShorten > 0 &&
    (ends.segCount === 1
      ? ends.lastSegLen > (startShorten + endShorten) * 1.5
      : ends.lastSegLen > endShorten * 1.5);

  if (!canShortenStart && !canShortenEnd) {
    return subpath;
  }

  const next = [...subpath];
  if (canShortenStart) {
    const first = next[0]!;
    if (first.op === "M") {
      next[0] = {
        op: "M",
        x: first.x + ends.startTangent.x * startShorten,
        y: first.y + ends.startTangent.y * startShorten,
      };
    }
  }
  if (canShortenEnd) {
    const lastIdx = next.length - 1;
    const last = next[lastIdx]!;
    if (last.op === "L" || last.op === "Q" || last.op === "C") {
      next[lastIdx] = {
        ...last,
        x: last.x - ends.endTangent.x * endShorten,
        y: last.y - ends.endTangent.y * endShorten,
      };
    }
  }
  return next;
}

function borderAttributes(border: ResolvedBorder | undefined): string {
  if (border === undefined) {
    return 'stroke="none" stroke-width="0"';
  }
  const w = Math.max(1, border.width);
  let dashAttr = "";
  switch (border.style) {
    case "solid":
      dashAttr = "";
      break;
    case "dashed":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w * 4)} ${roundSvgNumber(w * 3)}"`;
      break;
    case "dotted":
      dashAttr = ` stroke-dasharray="${roundSvgNumber(w)} ${roundSvgNumber(w * 2)}"`;
      break;
    default: {
      const exhaustive: never = border.style;
      return exhaustive;
    }
  }
  return `stroke="#${border.color.hex}" stroke-opacity="${roundSvgNumber(border.color.alpha ?? 1)}" stroke-width="${roundSvgNumber(border.width)}"${dashAttr}`;
}

function subpathFillHex(
  baseHex: string,
  fillMode: Exclude<"norm" | PresetPathFillMode, "none">,
): string {
  switch (fillMode) {
    case "norm":
      return baseHex;
    case "lighten":
      return shadeHex(baseHex, 0.6, 0.4);
    case "lightenLess":
      return shadeHex(baseHex, 0.8, 0.2);
    case "darken":
      return shadeHex(baseHex, 0.6, 0);
    case "darkenLess":
      return shadeHex(baseHex, 0.8, 0);
    default: {
      const exhaustive: never = fillMode;
      return exhaustive;
    }
  }
}

function shadeOverlayAttrs(
  fillMode: Exclude<"norm" | PresetPathFillMode, "none">,
): string | undefined {
  switch (fillMode) {
    case "norm":
      return undefined;
    case "lighten":
      return 'fill="#FFFFFF" fill-opacity="0.35"';
    case "lightenLess":
      return 'fill="#FFFFFF" fill-opacity="0.2"';
    case "darken":
      return 'fill="#000000" fill-opacity="0.3"';
    case "darkenLess":
      return 'fill="#000000" fill-opacity="0.15"';
    default: {
      const exhaustive: never = fillMode;
      return exhaustive;
    }
  }
}

function resolveSubpathFill(
  fill: ResolvedFill,
  fillDefId: string | undefined,
  fillMode: "norm" | PresetPathFillMode,
): {
  readonly fillAttr: string;
  readonly overlayAttr: string | undefined;
  readonly hasFill: boolean;
} {
  if (fillMode === "none") {
    return { fillAttr: 'fill="none"', overlayAttr: undefined, hasFill: false };
  }
  switch (fill.kind) {
    case "none":
      return {
        fillAttr: 'fill="none"',
        overlayAttr: undefined,
        hasFill: false,
      };
    case "solid":
      return {
        fillAttr: solidFillMechanism.paint(fill, { fillMode }).svg,
        overlayAttr: undefined,
        hasFill: true,
      };
    case "linear-gradient":
    case "radial-gradient":
    case "image":
      return {
        fillAttr: `fill="url(#${fillDefId})" fill-opacity="1"`,
        overlayAttr: shadeOverlayAttrs(fillMode),
        hasFill: true,
      };
    default: {
      const exhaustive: never = fill;
      return exhaustive;
    }
  }
}

function renderEvaluatedGeometryElements(
  model: ResolvedVisualModel,
  fillDefId: string | undefined,
  overrideAttrs?: string,
): string {
  const evaluated = evaluateGeometry(model.geometry, model.box);
  const { border } = model;
  const strokeAttr = borderAttributes(border);
  const startArrow = border?.startArrow ?? "none";
  const endArrow = border?.endArrow ?? "none";

  return evaluated.paths
    .map((pathItem) => {
      if (overrideAttrs !== undefined) {
        if (
          pathItem.fillMode === "none" &&
          (!pathItem.stroke || border === undefined)
        ) {
          return "";
        }
        return `<path d="${pathItem.d}" ${overrideAttrs}/>`;
      }

      const { fillAttr, overlayAttr, hasFill } = resolveSubpathFill(
        model.fill,
        fillDefId,
        pathItem.fillMode,
      );
      const allowStroke = pathItem.stroke && border !== undefined;
      const subpaths = splitSubpaths(pathItem.commands);
      const hasOpenSubpath = subpaths.some((sp) => sp.at(-1)?.op !== "Z");
      const hasArrows =
        allowStroke &&
        hasOpenSubpath &&
        (startArrow !== "none" || endArrow !== "none");

      if (!hasArrows) {
        const base = `<path d="${pathItem.d}" ${fillAttr} ${allowStroke ? strokeAttr : 'stroke="none" stroke-width="0"'}/>`;
        return overlayAttr === undefined
          ? base
          : `${base}<path d="${pathItem.d}" ${overlayAttr} stroke="none" stroke-width="0"/>`;
      }

      const fragments: string[] = [];
      if (hasFill) {
        fragments.push(`<path d="${pathItem.d}" ${fillAttr} stroke="none"/>`);
        if (overlayAttr !== undefined) {
          fragments.push(
            `<path d="${pathItem.d}" ${overlayAttr} stroke="none"/>`,
          );
        }
      }

      const startShorten = arrowShortenLength(startArrow, border.width);
      const endShorten = arrowShortenLength(endArrow, border.width);
      const strokeAlpha = border.color.alpha ?? 1;

      for (const subpath of subpaths) {
        const ends = extractOpenSubpathEnds(subpath);
        if (ends === undefined) {
          fragments.push(
            `<path d="${commandsToSvgPath(subpath)}" fill="none" ${strokeAttr}/>`,
          );
          continue;
        }
        const shortened = shortenOpenSubpath(
          subpath,
          startShorten,
          endShorten,
          ends,
        );
        fragments.push(
          `<path d="${commandsToSvgPath(shortened)}" fill="none" ${strokeAttr}/>`,
        );
        if (startArrow !== "none") {
          fragments.push(
            arrowHeadSvg(
              startArrow,
              ends.startPoint,
              { x: -ends.startTangent.x, y: -ends.startTangent.y },
              border.color.hex,
              strokeAlpha,
              border.width,
            ),
          );
        }
        if (endArrow !== "none") {
          fragments.push(
            arrowHeadSvg(
              endArrow,
              ends.endPoint,
              ends.endTangent,
              border.color.hex,
              strokeAlpha,
              border.width,
            ),
          );
        }
      }
      return fragments.join("");
    })
    .join("");
}

/**
 * Renders a closed ResolvedVisualModel into deterministic SVG markup with scoped `<defs>`,
 * linear/radial gradients, image `<pattern>`, rounded corners, dashed/dotted strokes,
 * open-path tangent arrows, screen-space `boxShadow`, and feathered `softEdge` masks.
 */
export function paintVisualModel(options: {
  readonly model: ResolvedVisualModel;
  readonly slideNumber: number;
  readonly objectId: number;
}): PaintOutput {
  const { model, slideNumber, objectId } = options;
  const { box } = model;
  const prefix = `obj-s${slideNumber}-o${objectId}`;
  const defs: string[] = [];

  let fillDefId: string | undefined;
  switch (model.fill.kind) {
    case "none":
    case "solid":
      break;
    case "linear-gradient":
    case "radial-gradient": {
      fillDefId = `${prefix}-fill`;
      defs.push(gradientFillMechanism.paint(model.fill, { id: fillDefId }).svg);
      break;
    }
    case "image": {
      fillDefId = `${prefix}-fill`;
      defs.push(
        imageFillMechanism.paint(model.fill, { id: fillDefId, box }).svg,
      );
      break;
    }
    default: {
      const exhaustive: never = model.fill;
      return exhaustive;
    }
  }

  let content = renderEvaluatedGeometryElements(model, fillDefId);

  if (model.softEdge !== undefined && model.softEdge > 0) {
    const r = model.softEdge;
    const minSide = Math.min(box.width, box.height);
    const borderPad = model.border === undefined ? 0 : model.border.width;
    const pad = Math.ceil(r * 2 + borderPad + 8);
    const maskX = roundSvgNumber(box.x - pad);
    const maskY = roundSvgNumber(box.y - pad);
    const maskW = roundSvgNumber(box.width + pad * 2);
    const maskH = roundSvgNumber(box.height + pad * 2);
    const blurId = `${prefix}-soft-blur`;
    const maskId = `${prefix}-soft-mask`;
    const innerStrokeWidth = r >= minSide / 2 ? 0 : Math.min(r, minSide * 0.45);
    const maskAttrs =
      innerStrokeWidth > 0
        ? `fill="white" stroke="black" stroke-width="${roundSvgNumber(innerStrokeWidth)}" filter="url(#${blurId})"`
        : `fill="white" stroke="none" filter="url(#${blurId})"`;
    const maskShapes = renderEvaluatedGeometryElements(
      model,
      fillDefId,
      maskAttrs,
    );
    defs.push(
      `<filter id="${blurId}" filterUnits="userSpaceOnUse" x="${maskX}" y="${maskY}" width="${maskW}" height="${maskH}"><feGaussianBlur stdDeviation="${roundSvgNumber(r / 2)}"/></filter>`,
    );
    defs.push(
      `<mask id="${maskId}" maskUnits="userSpaceOnUse" x="${maskX}" y="${maskY}" width="${maskW}" height="${maskH}">${maskShapes}</mask>`,
    );
    content = `<g mask="url(#${maskId})">${content}</g>`;
  }

  let layer = `<g transform="${visualSvgTransform(model)}">${content}</g>`;

  if (model.boxShadow !== undefined) {
    const shadow = model.boxShadow;
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const halfDiag = Math.hypot(box.width, box.height) / 2;
    const rad = (shadow.angle * Math.PI) / 180;
    const dx = shadow.distance * Math.cos(rad);
    const dy = shadow.distance * Math.sin(rad);
    const sigma = shadow.blur / 2;
    const borderPad = model.border === undefined ? 0 : model.border.width;
    const pad = shadow.blur * 3 + Math.abs(dx) + Math.abs(dy) + borderPad + 16;
    const minX = roundSvgNumber(cx - halfDiag - pad);
    const minY = roundSvgNumber(cy - halfDiag - pad);
    const filterSize = roundSvgNumber((halfDiag + pad) * 2);
    const shadowId = `${prefix}-shadow`;
    defs.push(
      `<filter id="${shadowId}" filterUnits="userSpaceOnUse" x="${minX}" y="${minY}" width="${filterSize}" height="${filterSize}"><feDropShadow dx="${roundSvgNumber(dx)}" dy="${roundSvgNumber(dy)}" stdDeviation="${roundSvgNumber(sigma)}" flood-color="#${shadow.color.hex}" flood-opacity="${roundSvgNumber(shadow.color.alpha ?? 1)}"/></filter>`,
    );
    layer = `<g filter="url(#${shadowId})">${layer}</g>`;
  }

  const defsXml = defs.length > 0 ? `<defs>${defs.join("")}</defs>` : "";
  return {
    svg: `${defsXml}${layer}`,
  };
}

// ---------------------------------------------------------------------------
// DrawingML XML Import Readers & Production Fill Mechanisms
// ---------------------------------------------------------------------------

const PercentValAttrSchema = z.object({
  val: OoxmlNonNegativeIntSchema,
});

const AlphaModFixAttrSchema = z.object({
  amt: OoxmlNonNegativeIntSchema,
});

const SrgbColorAttrSchema = z.object({
  val: HexRgb6Schema,
});

const SchemeColorAttrSchema = z.object({
  val: z.enum([
    "tx1",
    "bg1",
    "tx2",
    "bg2",
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ]),
});

const SysColorAttrSchema = z.object({
  val: z.string().min(1),
  lastClr: HexRgb6Schema.optional(),
});

const schemeToNativeSlot: Readonly<
  Record<z.infer<typeof SchemeColorAttrSchema>["val"], NativeColorSlot>
> = {
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2",
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const GradStopAttrSchema = z.object({
  pos: OoxmlNonNegativeIntSchema,
});

const GradLinAttrSchema = z.object({
  ang: OoxmlIntSchema.optional(),
  scaled: z.enum(["0", "1", "true", "false"]).optional(),
});

const LineAttrSchema = z.object({
  w: OoxmlNonNegativeIntSchema.optional(),
  cap: z.enum(["rnd", "sq", "flat"]).optional(),
  cmpd: z.enum(["sng", "dbl", "thickThin", "thinThick", "tri"]).optional(),
  algn: z.enum(["ctr", "in"]).optional(),
});

const DashPresetAttrSchema = z.object({
  val: z.string().min(1),
});

const LineEndAttrSchema = z.object({
  type: arrowTypeSchema.optional(),
  w: z.enum(["sm", "med", "lg"]).optional(),
  len: z.enum(["sm", "med", "lg"]).optional(),
});

const OuterShadowAttrSchema = z.object({
  blurRad: OoxmlNonNegativeIntSchema.optional(),
  dist: OoxmlNonNegativeIntSchema.optional(),
  dir: OoxmlIntSchema.optional(),
  algn: z.string().optional(),
  rotWithShape: z.enum(["0", "1", "true", "false"]).optional(),
});

const SoftEdgeAttrSchema = z.object({
  rad: OoxmlNonNegativeIntSchema,
});

/** Reads a DrawingML color child (`<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>`) with optional alpha and luminance transforms. */
export function readDrawingColor(
  parentCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ColorValue | undefined {
  let colorCursor: XmlElementCursor | undefined;
  let baseHex: string | undefined;
  let slot: NativeColorSlot | undefined;

  for (const child of parentCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    if (child.localName === "srgbClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val"]);
      const parsed = SrgbColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:srgbClr/@val",
          rawValue: raw["val"] ?? "",
        });
        return undefined;
      }
      baseHex = `#${parsed.data.val.toUpperCase()}`;
    } else if (child.localName === "schemeClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val"]);
      const parsed = SchemeColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:schemeClr/@val",
          rawValue: raw["val"] ?? "",
        });
        return undefined;
      }
      slot = schemeToNativeSlot[parsed.data.val];
    } else if (child.localName === "sysClr" && colorCursor === undefined) {
      colorCursor = child;
      const raw = child.attrsObject(["val", "lastClr"]);
      const parsed = SysColorAttrSchema.safeParse(raw);
      if (!parsed.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          attributeOrElement: "a:sysClr",
          rawValue: JSON.stringify(raw),
        });
        return undefined;
      }
      const fallback =
        parsed.data.val === "windowText"
          ? "000000"
          : parsed.data.val === "window"
            ? "FFFFFF"
            : undefined;
      const hex = parsed.data.lastClr ?? fallback;
      if (hex === undefined) return undefined;
      baseHex = `#${hex.toUpperCase()}`;
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `color:${child.localName}`,
      });
    }
  }

  if (colorCursor === undefined) return undefined;

  let alpha: number | undefined;
  let lumMod: number | undefined;
  let lumOff: number | undefined;

  for (const mod of colorCursor.remainingElementChildren()) {
    if (mod.namespaceUri === A_NS && mod.localName === "alpha") {
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        alpha = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumMod") {
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        lumMod = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else if (mod.namespaceUri === A_NS && mod.localName === "lumOff") {
      const parsed = PercentValAttrSchema.safeParse(mod.attrsObject(["val"]));
      if (parsed.success) {
        lumOff = Math.min(1, Math.max(0, parsed.data.val / 100_000));
      }
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: mod.xpath },
        namespaceUri: mod.namespaceUri,
        localName: mod.localName,
        mechanismHint: `color-transform:${mod.localName}`,
      });
    }
  }

  const baseColor: ColorValue | undefined =
    slot !== undefined
      ? lumMod !== undefined
        ? {
            slot,
            lumMod,
            ...(lumOff === undefined ? {} : { lumOff }),
          }
        : slot
      : baseHex;

  if (baseColor === undefined) return undefined;
  if (alpha !== undefined && alpha < 1) {
    return { color: baseColor, alpha };
  }
  return baseColor;
}

export type ImportedTransform = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly xEmu: number;
  readonly yEmu: number;
  readonly cxEmu: number;
  readonly cyEmu: number;
  readonly rotate?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
};

/** Parses `<a:xfrm>` offset, extent, rotation, and flips into pixels and raw EMU. */
export function readTransformXml(
  xfrmCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedTransform | undefined {
  const rawXfrm = xfrmCursor.attrsObject(["rot", "flipH", "flipV"]);
  const parsedXfrm = TransformAttributesFragmentSchema.safeParse(rawXfrm);
  const off = xfrmCursor.child(A_NS, "off");
  const ext = xfrmCursor.child(A_NS, "ext");
  if (off === undefined || ext === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: xfrmCursor.xpath },
    });
    return undefined;
  }
  const parsedOff = OffsetEmuFragmentSchema.safeParse(
    off.attrsObject(["x", "y"]),
  );
  const parsedExt = ExtentEmuFragmentSchema.safeParse(
    ext.attrsObject(["cx", "cy"]),
  );
  if (!parsedOff.success || !parsedExt.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: xfrmCursor.xpath },
      attributeOrElement: "a:xfrm",
    });
    return undefined;
  }
  const rotate =
    parsedXfrm.success &&
    parsedXfrm.data.rot !== undefined &&
    parsedXfrm.data.rot !== 0
      ? parsedXfrm.data.rot / 60_000
      : undefined;
  const flipH = parsedXfrm.success && parsedXfrm.data.flipH === true;
  const flipV = parsedXfrm.success && parsedXfrm.data.flipV === true;

  return {
    x: emuToPx(parsedOff.data.x),
    y: emuToPx(parsedOff.data.y),
    width: emuToPx(parsedExt.data.cx),
    height: emuToPx(parsedExt.data.cy),
    xEmu: parsedOff.data.x,
    yEmu: parsedOff.data.y,
    cxEmu: parsedExt.data.cx,
    cyEmu: parsedExt.data.cy,
    ...(rotate === undefined ? {} : { rotate }),
    ...(flipH ? { flipH: true } : {}),
    ...(flipV ? { flipV: true } : {}),
  };
}

/** Reads DrawingML `<a:gradFill>` into a `LinearGradient` or `RadialGradient`. */
export function readGradientFill(
  gradFillCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): LinearGradient | RadialGradient | undefined {
  gradFillCursor.consumeAttr("rotWithShape");
  gradFillCursor.consumeAttr("flip");
  const gsLst = gradFillCursor.child(A_NS, "gsLst");
  const stops: GradientStop[] = [];
  if (gsLst !== undefined) {
    for (const gs of gsLst.children(A_NS, "gs")) {
      const rawPos = gs.attrsObject(["pos"]);
      const parsedPos = GradStopAttrSchema.safeParse(rawPos);
      const color = readDrawingColor(gs, location, diagnostics);
      if (!parsedPos.success || color === undefined) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: gs.xpath },
          attributeOrElement: "a:gs",
          rawValue: JSON.stringify(rawPos),
        });
        continue;
      }
      stops.push({
        position: Math.min(1, Math.max(0, parsedPos.data.pos / 100_000)),
        color,
      });
    }
  }
  if (stops.length < 2) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: gradFillCursor.xpath },
      attributeOrElement: "a:gradFill/a:gsLst",
    });
    return undefined;
  }

  let result: LinearGradient | RadialGradient | undefined;
  const lin = gradFillCursor.child(A_NS, "lin");
  if (lin !== undefined) {
    const rawLin = lin.attrsObject(["ang", "scaled"]);
    const parsedLin = GradLinAttrSchema.safeParse(rawLin);
    const drawingAngle = parsedLin.success ? (parsedLin.data.ang ?? 0) : 0;
    const authorAngle = (((drawingAngle / 60_000 + 90) % 360) + 360) % 360;
    result = { type: "linear-gradient", angle: authorAngle, stops };
  } else {
    const pathEl = gradFillCursor.child(A_NS, "path");
    if (pathEl !== undefined) {
      pathEl.consumeAttr("path");
      const fillToRect = pathEl.child(A_NS, "fillToRect");
      if (fillToRect !== undefined) {
        fillToRect.consumeAttr("l");
        fillToRect.consumeAttr("t");
        fillToRect.consumeAttr("r");
        fillToRect.consumeAttr("b");
      }
      result = { type: "radial-gradient", stops };
    }
  }
  const tileRect = gradFillCursor.child(A_NS, "tileRect");
  if (tileRect !== undefined) {
    tileRect.consumeAttr("l");
    tileRect.consumeAttr("t");
    tileRect.consumeAttr("r");
    tileRect.consumeAttr("b");
  }
  return result;
}

/** Reads DrawingML `<a:blipFill>` embedded image bytes and optional `<a:alphaModFix>` opacity. */
export function readBlipFill(
  blipFillCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): { readonly image: Uint8Array; readonly opacity?: number } | undefined {
  blipFillCursor.consumeAttr("rotWithShape");
  blipFillCursor.consumeAttr("dpi");
  const srcRect = blipFillCursor.child(A_NS, "srcRect");
  if (srcRect !== undefined) {
    srcRect.consumeAttr("l");
    srcRect.consumeAttr("t");
    srcRect.consumeAttr("r");
    srcRect.consumeAttr("b");
  }
  const stretch = blipFillCursor.child(A_NS, "stretch");
  if (stretch !== undefined) {
    stretch.ignoreEmptyChild(A_NS, "fillRect");
  }
  const blip = blipFillCursor.child(A_NS, "blip");
  if (blip === undefined) {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blipFillCursor.xpath },
    });
    return undefined;
  }
  blip.consumeAttr("cstate");
  const embedId = blip.attr("embed", R_NS);
  const linkId = blip.attr("link", R_NS);
  let opacity: number | undefined;
  const alphaModFix = blip.child(A_NS, "alphaModFix");
  if (alphaModFix !== undefined) {
    const parsedAmt = AlphaModFixAttrSchema.safeParse(
      alphaModFix.attrsObject(["amt"]),
    );
    if (parsedAmt.success) {
      opacity = Math.min(1, Math.max(0, parsedAmt.data.amt / 100_000));
    }
  }
  if (embedId === undefined) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      namespaceUri: A_NS,
      localName: "blip",
      mechanismHint: linkId !== undefined ? "linked-image-fill" : "blip-fill",
    });
    return undefined;
  }
  const rels = pkg.readRelationships(location.part);
  diagnostics.push(...rels.diagnostics);
  const rel = rels.entries.find((entry) => entry.Id === embedId);
  if (rel === undefined || rel.TargetMode === "External") {
    diagnostics.push({
      code: "broken-relationship",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      relationshipId: embedId,
    });
    return undefined;
  }
  const targetPart = resolvePartPath(location.part, rel.Target);
  if (targetPart === undefined || !pkg.hasPart(targetPart)) {
    diagnostics.push({
      code: "missing-target-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: blip.xpath },
      relationshipId: embedId,
      targetPart: targetPart ?? rel.Target,
    });
    return undefined;
  }
  const detachedBytes = pkg.readDetachedBytes(targetPart);
  if (detachedBytes === undefined) return undefined;
  try {
    readImage(detachedBytes);
  } catch {
    diagnostics.push({
      code: "corrupt-part",
      category: "corrupt-input",
      disposition: "omitted",
      location: { part: targetPart },
    });
    return undefined;
  }
  return {
    image: detachedBytes,
    ...(opacity === undefined ? {} : { opacity }),
  };
}

export type ImportedLineBorder = {
  readonly borderColor?: ColorValue;
  readonly borderWidth?: number;
  readonly borderStyle?: "solid" | "dashed" | "dotted";
  readonly startArrow?: ArrowType;
  readonly endArrow?: ArrowType;
};

/** Reads DrawingML `<a:ln>` stroke width, fill, dash preset, join, and line end arrows. */
export function readLineBorder(
  lnCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedLineBorder {
  const rawLn = lnCursor.attrsObject(["w", "cap", "cmpd", "algn"]);
  const parsedLn = LineAttrSchema.safeParse(rawLn);
  const width =
    parsedLn.success && parsedLn.data.w !== undefined
      ? emuToPx(parsedLn.data.w)
      : undefined;

  let borderColor: ColorValue | undefined;
  let borderStyle: "solid" | "dashed" | "dotted" | undefined;
  let startArrow: ArrowType | undefined;
  let endArrow: ArrowType | undefined;
  let hasNoFill = false;

  for (const child of lnCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    switch (child.localName) {
      case "noFill":
        hasNoFill = true;
        break;
      case "solidFill":
        borderColor = readDrawingColor(child, location, diagnostics);
        break;
      case "prstDash": {
        const rawDash = child.attrsObject(["val"]);
        const parsedDash = DashPresetAttrSchema.safeParse(rawDash);
        if (parsedDash.success) {
          const val = parsedDash.data.val;
          if (val === "solid") borderStyle = "solid";
          else if (val === "dot" || val === "sysDot") borderStyle = "dotted";
          else if (
            val === "dash" ||
            val === "lgDash" ||
            val === "dashDot" ||
            val === "lgDashDot" ||
            val === "lgDashDotDot" ||
            val === "sysDash" ||
            val === "sysDashDot" ||
            val === "sysDashDotDot"
          ) {
            borderStyle = "dashed";
          } else {
            diagnostics.push({
              code: "unsupported-mechanism",
              category: "unsupported",
              disposition: "omitted",
              location: { ...location, xpath: child.xpath },
              namespaceUri: A_NS,
              localName: "prstDash",
              mechanismHint: `dash:${val}`,
            });
          }
        }
        break;
      }
      case "headEnd":
      case "tailEnd": {
        const rawEnd = child.attrsObject(["type", "w", "len"]);
        const parsedEnd = LineEndAttrSchema.safeParse(rawEnd);
        if (!parsedEnd.success) {
          diagnostics.push({
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: { ...location, xpath: child.xpath },
            attributeOrElement: `a:${child.localName}`,
            rawValue: JSON.stringify(rawEnd),
          });
          break;
        }
        if (parsedEnd.data.type !== undefined) {
          if (child.localName === "headEnd") startArrow = parsedEnd.data.type;
          else endArrow = parsedEnd.data.type;
        }
        break;
      }
      case "round":
      case "miter":
      case "bevel":
        if (child.localName === "miter") child.consumeAttr("lim");
        break;
      default:
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          namespaceUri: child.namespaceUri,
          localName: child.localName,
          mechanismHint: `line:${child.localName}`,
        });
        break;
    }
  }

  if (hasNoFill || borderColor === undefined) {
    return {};
  }

  return {
    borderColor,
    borderWidth: width ?? 1,
    ...(borderStyle === undefined ? {} : { borderStyle }),
    ...(startArrow === undefined ? {} : { startArrow }),
    ...(endArrow === undefined ? {} : { endArrow }),
  };
}

/** Reads DrawingML `<a:effectLst>` outer shadow and soft edge while reporting unsupported effects cleanly. */
export function readEffectList(
  effectLstCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): {
  readonly boxShadow?: {
    readonly color?: ColorValue;
    readonly blur?: number;
    readonly distance?: number;
    readonly angle?: number;
  };
  readonly softEdge?: number;
} {
  let boxShadow:
    | {
        readonly color?: ColorValue;
        readonly blur?: number;
        readonly distance?: number;
        readonly angle?: number;
      }
    | undefined;
  let softEdge: number | undefined;

  for (const child of effectLstCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    if (child.localName === "outerShdw") {
      const raw = child.attrsObject([
        "blurRad",
        "dist",
        "dir",
        "algn",
        "rotWithShape",
      ]);
      const parsed = OuterShadowAttrSchema.safeParse(raw);
      const color = readDrawingColor(child, location, diagnostics);
      if (parsed.success) {
        boxShadow = {
          ...(color === undefined ? {} : { color }),
          blur:
            parsed.data.blurRad === undefined
              ? 0
              : emuToPx(parsed.data.blurRad),
          distance:
            parsed.data.dist === undefined ? 0 : emuToPx(parsed.data.dist),
          angle: parsed.data.dir === undefined ? 0 : parsed.data.dir / 60_000,
        };
      }
    } else if (child.localName === "softEdge") {
      const raw = child.attrsObject(["rad"]);
      const parsed = SoftEdgeAttrSchema.safeParse(raw);
      if (parsed.success && parsed.data.rad > 0) {
        softEdge = emuToPx(parsed.data.rad);
      }
    } else {
      diagnostics.push({
        code: "unsupported-mechanism",
        category: "unsupported",
        disposition: "omitted",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
        mechanismHint: `effect:${child.localName}`,
      });
    }
  }

  return {
    ...(boxShadow === undefined ? {} : { boxShadow }),
    ...(softEdge === undefined ? {} : { softEdge }),
  };
}

export type ImportedSpPrVisuals = {
  readonly transform: ImportedTransform;
  readonly geometry?: Geometry;
  readonly style: ShapeVisualStyle;
};

/**
 * Reads PresentationML `<p:spPr>` geometry, transform, fills, borders, and effects.
 * Isolates unsupported visual elements/fills by emitting a localized ImportDiagnostic
 * and omitting only the unsupported visual attribute while keeping geometry and text intact.
 */
export function readSpPrVisuals(
  spPrCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): ImportedSpPrVisuals | undefined {
  let transform: ImportedTransform | undefined;
  let geometry: Geometry | undefined;
  let backgroundColor: ColorValue | undefined;
  let backgroundImage: ShapeVisualStyle["backgroundImage"];
  let opacity: number | undefined;
  let border: ImportedLineBorder = {};
  let effects: ReturnType<typeof readEffectList> = {};

  // Read xfrm first so custom geometry fallback extent is available regardless of child element order.
  const xfrmCursor = spPrCursor.child(A_NS, "xfrm");
  if (xfrmCursor !== undefined) {
    transform = readTransformXml(xfrmCursor, location, diagnostics);
  }
  if (transform === undefined) {
    if (xfrmCursor === undefined) {
      diagnostics.push({
        code: "corrupt-part",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: spPrCursor.xpath },
      });
    }
    return undefined;
  }

  const fallbackUnits = {
    w: transform.cxEmu,
    h: transform.cyEmu,
  };

  for (const child of spPrCursor.remainingElementChildren()) {
    if (child.namespaceUri !== A_NS) {
      diagnostics.push({
        code: "unconsumed-xml",
        category: "unsupported",
        disposition: "omitted",
        kind: "element",
        location: { ...location, xpath: child.xpath },
        namespaceUri: child.namespaceUri,
        localName: child.localName,
      });
      continue;
    }
    switch (child.localName) {
      case "prstGeom":
        geometry = readPresetGeometryXml(child, location, diagnostics);
        break;
      case "custGeom":
        geometry = readCustomGeometryXml(
          child,
          fallbackUnits,
          location,
          diagnostics,
        );
        break;
      case "noFill":
        break;
      case "solidFill":
        backgroundColor = readDrawingColor(child, location, diagnostics);
        break;
      case "gradFill":
        backgroundImage = readGradientFill(child, location, diagnostics);
        break;
      case "blipFill": {
        const blip = readBlipFill(child, pkg, location, diagnostics);
        if (blip !== undefined) {
          backgroundImage = { image: blip.image };
          if (blip.opacity !== undefined) opacity = blip.opacity;
        }
        break;
      }
      case "ln":
        border = readLineBorder(child, location, diagnostics);
        break;
      case "effectLst":
        effects = readEffectList(child, location, diagnostics);
        break;
      default:
        diagnostics.push({
          code: "unsupported-mechanism",
          category: "unsupported",
          disposition: "omitted",
          location: { ...location, xpath: child.xpath },
          namespaceUri: child.namespaceUri,
          localName: child.localName,
          mechanismHint: child.localName,
        });
        break;
    }
  }

  const style: ShapeVisualStyle = {
    ...(backgroundColor === undefined ? {} : { backgroundColor }),
    ...(backgroundImage === undefined ? {} : { backgroundImage }),
    ...(border.borderColor === undefined
      ? {}
      : { borderColor: border.borderColor }),
    ...(border.borderWidth === undefined
      ? {}
      : { borderWidth: border.borderWidth }),
    ...(border.borderStyle === undefined
      ? {}
      : { borderStyle: border.borderStyle }),
    ...(border.startArrow === undefined
      ? {}
      : { startArrow: border.startArrow }),
    ...(border.endArrow === undefined ? {} : { endArrow: border.endArrow }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(effects.boxShadow === undefined
      ? {}
      : { boxShadow: effects.boxShadow }),
    ...(effects.softEdge === undefined ? {} : { softEdge: effects.softEdge }),
    ...(transform.rotate === undefined ? {} : { rotate: transform.rotate }),
  };

  return {
    transform,
    ...(geometry === undefined ? {} : { geometry }),
    style,
  };
}

type SolidFillModel = Extract<ResolvedFill, { readonly kind: "solid" }>;
type GradientFillModel = Extract<
  ResolvedFill,
  { readonly kind: "linear-gradient" | "radial-gradient" }
>;
type ImageFillModel = Extract<ResolvedFill, { readonly kind: "image" }>;

function readFillSubElement(
  context: ShapeTreeImportContext,
): ShapeTreeImportOutcome {
  const diagnostics: ImportDiagnostic[] = [];
  const spPrCursor = context.cursor.child(P_NS, "spPr");
  if (spPrCursor === undefined) {
    return { diagnostics };
  }
  readSpPrVisuals(
    spPrCursor,
    context.package,
    { part: context.part, slideIndex: context.slideIndex },
    diagnostics,
  );
  return { diagnostics };
}

/** Native canvas mechanism for DrawingML `<a:solidFill>`. */
export const solidFillMechanism = defineCanvasMechanism({
  key: "solid-fill",
  role: "canvas",
  emit(fill: SolidFillModel, _context: undefined): { readonly xml: string } {
    return {
      xml: `<a:solidFill>${resolvedColorXml(fill.color)}</a:solidFill>`,
    };
  },
  paint(
    fill: SolidFillModel,
    context: {
      readonly fillMode: Exclude<"norm" | PresetPathFillMode, "none">;
    },
  ): PaintOutput {
    const hex = subpathFillHex(fill.color.hex, context.fillMode);
    return {
      svg: `fill="#${hex}" fill-opacity="${roundSvgNumber(fill.color.alpha ?? 1)}"`,
    };
  },
  import: {
    target: "shape-tree",
    localName: "sp",
    matches: () => false,
    read: readFillSubElement,
  },
});

/** Native canvas mechanism for DrawingML `<a:gradFill>` (linear and radial gradients). */
export const gradientFillMechanism = defineCanvasMechanism({
  key: "gradient-fill",
  role: "canvas",
  emit(fill: GradientFillModel, _context: undefined): { readonly xml: string } {
    return { xml: gradientFillXml(fill) };
  },
  paint(
    fill: GradientFillModel,
    context: { readonly id: string },
  ): PaintOutput {
    return { svg: gradientDefSvg(context.id, fill) };
  },
  import: {
    target: "shape-tree",
    localName: "sp",
    matches: () => false,
    read: readFillSubElement,
  },
});

/** Native canvas mechanism for DrawingML `<a:blipFill>` shape image fills. */
export const imageFillMechanism = defineCanvasMechanism({
  key: "image-fill",
  role: "canvas",
  emit(
    fill: ImageFillModel,
    registerResource: ((resource: CapabilityResource) => string) | undefined,
  ): { readonly xml: string } {
    if (registerResource === undefined) {
      invalid(
        "image-resource-failure",
        "Cannot emit image background fill without resource registrar",
      );
    }
    const relationshipId = registerResource({
      bytes: fill.image.bytes,
      extension: fill.image.extension,
      contentType: fill.image.contentType,
      relationshipType: imageRelationship,
    });
    const alpha =
      fill.opacity === undefined
        ? ""
        : `<a:alphaModFix amt="${Math.round(fill.opacity * 100000)}"/>`;
    return {
      xml: `<a:blipFill rotWithShape="1"><a:blip r:embed="${xmlEscape(relationshipId)}">${alpha}</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>`,
    };
  },
  paint(
    fill: ImageFillModel,
    context: { readonly id: string; readonly box: ShapeBox },
  ): PaintOutput {
    const pngBytes = decodeImageToPngBytes(fill.image);
    const base64 = Buffer.from(pngBytes).toString("base64");
    const imgOpacity = roundSvgNumber(fill.opacity ?? 1);
    return {
      svg: `<pattern id="${context.id}" patternUnits="userSpaceOnUse" x="${roundSvgNumber(context.box.x)}" y="${roundSvgNumber(context.box.y)}" width="${roundSvgNumber(context.box.width)}" height="${roundSvgNumber(context.box.height)}"><image x="0" y="0" width="${roundSvgNumber(context.box.width)}" height="${roundSvgNumber(context.box.height)}" preserveAspectRatio="none" opacity="${imgOpacity}" href="data:image/png;base64,${base64}"/></pattern>`,
    };
  },
  import: {
    target: "shape-tree",
    localName: "sp",
    matches: () => false,
    read: readFillSubElement,
  },
});
