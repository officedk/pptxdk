import {
  CapabilityError,
  borderObjectSchema,
  borderValueSchema,
  imageBackgroundSchema,
  linearGradientSchema,
  radialGradientSchema,
  shadowSchema,
  visualStyleSchema,
  type CapabilityResource,
  type ColorValue,
  type GradientStop,
  type ImageBackgroundFit,
  type ImageBackgroundPosition,
  type LinearGradient,
  type PaintOutput,
  type RadialGradient,
  type ResolvedColor,
  type ThemeContext,
  type VisualStyle as CoreVisualStyle,
} from "@pptxdk/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@pptxdk/emit";
import {
  A_NS,
  P_NS,
  arrowTypeSchema,
  arrowTypes,
  defineCanvasMechanism,
  readBlipFill as importReadBlipFill,
  readDrawingColor,
  readEffectList,
  readGradientFill,
  readLineBorder,
  readTransformXml,
  type ArrowType,
  type ImportDiagnostic,
  type ImportLocation,
  type ImportedLineBorder,
  type ImportedTransform,
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

export type ResolvedImageFillSource =
  | ({ readonly kind: "embedded" } & LoadedImage)
  | { readonly kind: "external"; readonly externalUrl: string };

export type ResolvedImageFill = {
  readonly kind: "image";
  readonly source: ResolvedImageFillSource;
  readonly fit: ImageBackgroundFit;
  readonly position: { readonly x: number; readonly y: number };
  readonly opacity?: number;
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
  | ResolvedImageFill;

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
  readonly explicitNoShadow?: boolean;
  readonly softEdge?: number;
  readonly rotate: number;
  readonly rotationUnits?: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly hasVisibleStyle: boolean;
  readonly layoutOnlyWhenInvisible: boolean;
};

export { arrowTypeSchema, arrowTypes };
export type { ArrowType };

export const lineEndStyleSchema = z
  .strictObject({
    startArrow: arrowTypeSchema.optional(),
    endArrow: arrowTypeSchema.optional(),
  })
  .readonly();
export type LineEndStyle = z.infer<typeof lineEndStyleSchema>;

export const shapeBorderObjectSchema = borderObjectSchema
  .unwrap()
  .extend({
    startArrow: arrowTypeSchema.optional(),
    endArrow: arrowTypeSchema.optional(),
  })
  .strict()
  .readonly();

export const shapeBorderValueSchema = z.union([
  z.literal("none"),
  shapeBorderObjectSchema,
]);
export type ShapeBorderValue = z.infer<typeof shapeBorderValueSchema>;

export const lineEndStyleKeys = Object.keys(
  lineEndStyleSchema.unwrap().shape,
) as readonly (keyof LineEndStyle)[];

export type ShapeVisualStyle = Omit<CoreVisualStyle, "border"> & {
  readonly border?: ShapeBorderValue;
};

export const visualStyleKeys = Object.keys(
  visualStyleSchema.unwrap().shape,
) as readonly (keyof CoreVisualStyle)[];

export type ShapePropertiesOptions = {
  readonly theme: ThemeContext;
  readonly style: VisualStyleRecord;
  readonly box: ShapeBox;
  readonly geometry?: Geometry;
  readonly image?: boolean;
  /** Allows `startArrow` / `endArrow` inside `style.border` (Shape only). */
  readonly allowBorderArrows?: boolean;
  /** Layout containers without visible styling do not require emitting native transforms. */
  readonly layoutOnlyWhenInvisible?: boolean;
};

export type ShapePropertiesResult = {
  readonly xml: string;
  readonly opacity?: number;
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

function resolveImageFillSource(source: unknown): ResolvedImageFillSource {
  if (isRecord(source) && typeof source["externalUrl"] === "string") {
    return { kind: "external", externalUrl: source["externalUrl"] };
  }
  if (
    source instanceof URL &&
    (source.protocol === "http:" || source.protocol === "https:")
  ) {
    return { kind: "external", externalUrl: source.toString() };
  }
  if (
    typeof source === "string" &&
    (source.startsWith("http://") || source.startsWith("https://"))
  ) {
    return { kind: "external", externalUrl: source };
  }
  return { kind: "embedded", ...readImage(source) };
}

function resolvePositionCoordinate(
  value: number | `${number}%` | undefined,
): number {
  if (value === undefined) return 0.5;
  if (typeof value === "number") {
    return Math.max(0, Math.min(1, value));
  }
  const parsed = Number.parseFloat(value.slice(0, -1));
  return Number.isFinite(parsed) ? Math.max(0, Math.min(1, parsed / 100)) : 0.5;
}

export function resolveImageBackgroundPosition(
  position: ImageBackgroundPosition | undefined,
): { readonly x: number; readonly y: number } {
  if (position === undefined) return { x: 0.5, y: 0.5 };
  if (typeof position === "string") {
    switch (position) {
      case "center":
        return { x: 0.5, y: 0.5 };
      case "top":
        return { x: 0.5, y: 0 };
      case "bottom":
        return { x: 0.5, y: 1 };
      case "left":
        return { x: 0, y: 0.5 };
      case "right":
        return { x: 1, y: 0.5 };
      case "top left":
        return { x: 0, y: 0 };
      case "top right":
        return { x: 1, y: 0 };
      case "bottom left":
        return { x: 0, y: 1 };
      case "bottom right":
        return { x: 1, y: 1 };
      default: {
        const exhaustive: never = position;
        return exhaustive;
      }
    }
  }
  return {
    x: resolvePositionCoordinate(position.x),
    y: resolvePositionCoordinate(position.y),
  };
}

/** Validates and resolves `{ image, fit?, position?, opacity? }` background declarations. */
export function resolveImageBackgroundSpec(
  spec: unknown,
  opacity: number | undefined,
): ResolvedImageFill {
  const parsed = imageBackgroundSchema.safeParse(spec);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.code === "unrecognized_keys" && issue.keys[0] !== undefined) {
      invalid(
        "invalid-visual-style",
        `background image contains unsupported parameter ${issue.keys[0]}`,
      );
    }
    invalid(
      "invalid-visual-style",
      issue?.message ?? "Invalid image background declaration",
    );
  }
  const source = resolveImageFillSource(parsed.data.image);
  const combinedOpacity = (parsed.data.opacity ?? 1) * (opacity ?? 1);
  return {
    kind: "image",
    source,
    fit: parsed.data.fit ?? "fill",
    position: resolveImageBackgroundPosition(parsed.data.position),
    ...(combinedOpacity < 1 ? { opacity: combinedOpacity } : {}),
  };
}

/** Resolves a unified `BackgroundValue` (`"none"`, color, gradient, or `{ image, fit?, position?, opacity? }`). */
export function resolveBackgroundValue(
  theme: ThemeContext,
  value: unknown,
  opacity: number | undefined,
): ResolvedFill {
  if (value === undefined || value === "none") {
    return { kind: "none" };
  }
  if (isRecord(value) && Object.hasOwn(value, "image")) {
    return resolveImageBackgroundSpec(value, opacity);
  }
  if (
    isRecord(value) &&
    (value["type"] === "linear-gradient" || value["type"] === "radial-gradient")
  ) {
    return resolveGradientSpec(theme, value, opacity);
  }
  return {
    kind: "solid",
    color: withOpacity(theme.resolveColor(value as ColorValue), opacity),
  };
}

function resolveVisualFill(
  theme: ThemeContext,
  style: VisualStyleRecord,
  opacity: number | undefined,
): ResolvedFill {
  try {
    return resolveBackgroundValue(theme, style["background"], opacity);
  } catch (error) {
    if (
      error instanceof CapabilityError &&
      error.details?.["field"] === undefined &&
      Object.hasOwn(style, "background") &&
      style["background"] !== undefined
    ) {
      throw new CapabilityError(error.code, error.message, error.path, {
        ...error.details,
        field: "style.background",
      });
    }
    throw error;
  }
}

function resolveVisualBorder(
  theme: ThemeContext,
  style: VisualStyleRecord,
  opacity: number | undefined,
  allowBorderArrows: boolean,
): ResolvedBorder | undefined {
  if (!Object.hasOwn(style, "border") || style["border"] === undefined) {
    return undefined;
  }
  const rawBorder = style["border"];
  if (rawBorder === "none") {
    return undefined;
  }
  const parsed = (
    allowBorderArrows ? shapeBorderValueSchema : borderValueSchema
  ).safeParse(rawBorder);
  if (!parsed.success || parsed.data === "none") {
    const issue = parsed.success ? undefined : parsed.error.issues[0];
    if (issue?.code === "unrecognized_keys" && issue.keys[0] !== undefined) {
      invalid(
        "invalid-visual-style",
        `border contains unsupported property ${issue.keys[0]}`,
      );
    }
    invalid(
      "invalid-visual-style",
      issue?.message ?? "Invalid border declaration",
    );
  }
  const width = parsed.data.width ?? 1;
  if (width === 0) {
    return undefined;
  }
  lineWidthEmu(width);
  const startArrow = (
    parsed.data as { readonly startArrow?: ArrowType | undefined }
  ).startArrow;
  const endArrow = (
    parsed.data as { readonly endArrow?: ArrowType | undefined }
  ).endArrow;
  return {
    color: withOpacity(theme.resolveColor(parsed.data.color), opacity),
    width,
    style: parsed.data.style ?? "solid",
    ...(startArrow === undefined ? {} : { startArrow }),
    ...(endArrow === undefined ? {} : { endArrow }),
  };
}

/** Determines whether a Box style specifies visible fill, border, or shadow. */
export function hasVisibleBoxStyle(
  style: Readonly<Record<string, unknown>> | undefined,
): boolean {
  if (style === undefined || typeof style !== "object" || style === null)
    return false;
  const hasNewBorder =
    style["border"] !== undefined &&
    style["border"] !== "none" &&
    !(
      typeof style["border"] === "object" &&
      style["border"] !== null &&
      (style["border"] as Record<string, unknown>)["width"] === 0
    );
  return (
    (style["background"] !== undefined && style["background"] !== "none") ||
    hasNewBorder ||
    (style["boxShadow"] !== undefined && style["boxShadow"] !== "none")
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

  for (const [name, value] of [
    ["flipH", style["flipH"]],
    ["flipV", style["flipV"]],
  ] as const) {
    if (value !== undefined && typeof value !== "boolean") {
      invalid("invalid-visual-style", `${name} must be a boolean`);
    }
  }
  const flipH = Boolean(style["flipH"]);
  const flipV = Boolean(style["flipV"]);

  const opacity =
    style["opacity"] === undefined
      ? undefined
      : finite(style["opacity"], "opacity");
  if (opacity !== undefined && (opacity < 0 || opacity > 1)) {
    invalid("invalid-visual-style", "opacity must be a number between 0 and 1");
  }
  if (options.image && style["borderRadius"] !== undefined) {
    invalid(
      "unsupported-style-combination",
      "Image does not support borderRadius",
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

  const fill = resolveVisualFill(theme, style, opacity);
  const border = resolveVisualBorder(
    theme,
    style,
    opacity,
    Boolean(options.allowBorderArrows),
  );
  const explicitNoShadow = style["boxShadow"] === "none";
  const boxShadow =
    style["boxShadow"] === undefined || explicitNoShadow
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

  const hasVisibleStyle = hasVisibleBoxStyle(style);
  const layoutOnlyWhenInvisible = Boolean(options.layoutOnlyWhenInvisible);
  const resolvedGeometry = resolveGeometry({
    geometry: options.geometry,
    borderRadius:
      layoutOnlyWhenInvisible && !hasVisibleStyle ? undefined : radius,
    box,
  });
  const rotation = rotationUnits(style);

  return {
    box,
    geometry: resolvedGeometry.geometry,
    geometryName: resolvedGeometry.name,
    fill,
    ...(border === undefined ? {} : { border }),
    ...(opacity === undefined ? {} : { opacity }),
    ...(boxShadow === undefined ? {} : { boxShadow }),
    ...(explicitNoShadow ? { explicitNoShadow: true } : {}),
    ...(softEdge === undefined ? {} : { softEdge }),
    rotate: rotation.degrees,
    ...(rotation.units === undefined ? {} : { rotationUnits: rotation.units }),
    flipH,
    flipV,
    hasVisibleStyle,
    layoutOnlyWhenInvisible,
  };
}

export function emitFillXml(
  fill: ResolvedFill,
  registerResource: ((resource: CapabilityResource) => string) | undefined,
  box?: ShapeBox,
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
      return emitImageFillXml(fill, registerResource, box);
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

function emitEffectsXml(
  model: ResolvedVisualModel,
  options?: { readonly explicitEmptyWhenCleared?: boolean },
): string {
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
  if (items.length > 0) {
    return `<a:effectLst>${items.join("")}</a:effectLst>`;
  }
  return options?.explicitEmptyWhenCleared && model.explicitNoShadow
    ? "<a:effectLst/>"
    : "";
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
    : emitFillXml(model.fill, registerResource, box);
  const line = emitBorderXml(model.border);
  const effects = emitEffectsXml(model);
  return {
    xml: `<p:spPr>${transform}${geometryXmlText}${visualFill}${line}${effects}</p:spPr>`,
    ...(model.opacity === undefined ? {} : { opacity: model.opacity }),
    hasVisibleStyle: model.hasVisibleStyle,
    geometryName: model.geometryName,
    geometry: model.geometry,
  };
}

export type PlaceholderSpPrFlags = {
  readonly hasBox: boolean;
  readonly hasRotate: boolean;
  readonly hasFlipH: boolean;
  readonly hasFlipV: boolean;
  readonly hasGeometryOrRadius: boolean;
  readonly hasFill: boolean;
  readonly hasBorder: boolean;
  readonly hasEffects: boolean;
  readonly includeDefaultRectWhenHasBox?: boolean;
};

/**
 * Serializes `<p:spPr>` for a master, layout, or bound slide placeholder, emitting only
 * transform, geometry, fill, border, and effect elements declared on that layer.
 */
export function emitPlaceholderVisualSpPr(
  model: ResolvedVisualModel,
  flags: PlaceholderSpPrFlags,
  registerResource?: (resource: CapabilityResource) => string,
): string {
  let xfrmXml = "";
  if (flags.hasBox || flags.hasRotate || flags.hasFlipH || flags.hasFlipV) {
    const rotAttr = flags.hasRotate ? ` rot="${model.rotationUnits ?? 0}"` : "";
    const flipHAttr = flags.hasFlipH
      ? model.flipH
        ? ' flipH="1"'
        : ' flipH="0"'
      : "";
    const flipVAttr = flags.hasFlipV
      ? model.flipV
        ? ' flipV="1"'
        : ' flipV="0"'
      : "";
    if (flags.hasBox) {
      const { box } = model;
      xfrmXml = `<a:xfrm${rotAttr}${flipHAttr}${flipVAttr}><a:off x="${emu(box.x, "x", true, true)}" y="${emu(box.y, "y", true, true)}"/><a:ext cx="${emu(box.width, "width", false)}" cy="${emu(box.height, "height", false)}"/></a:xfrm>`;
    } else {
      xfrmXml = `<a:xfrm${rotAttr}${flipHAttr}${flipVAttr}/>`;
    }
  }
  const geomXml =
    flags.hasGeometryOrRadius ||
    (flags.hasBox && flags.includeDefaultRectWhenHasBox === true)
      ? resolvedGeometryXml(model.geometry)
      : "";
  const fillXml = flags.hasFill
    ? emitFillXml(model.fill, registerResource, model.box)
    : "";
  const borderXml = flags.hasBorder ? emitBorderXml(model.border) : "";
  const effectsXml = flags.hasEffects
    ? emitEffectsXml(model, { explicitEmptyWhenCleared: true })
    : "";
  const inner = `${xfrmXml}${geomXml}${fillXml}${borderXml}${effectsXml}`;
  return inner === "" ? "<p:spPr/>" : `<p:spPr>${inner}</p:spPr>`;
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
    ...(model.fill.kind === "image" && model.fill.source.kind === "external"
      ? { placeholders: [{ url: model.fill.source.externalUrl }] }
      : {}),
  };
}

// ---------------------------------------------------------------------------
// DrawingML XML Import Readers & Production Fill Mechanisms
// ---------------------------------------------------------------------------

export {
  readDrawingColor,
  readEffectList,
  readGradientFill,
  readLineBorder,
  readTransformXml,
};
export type { ImportedLineBorder, ImportedTransform };

/** Reads DrawingML `<a:blipFill>` embedded image bytes and validates them against supported decoders. */
export function readBlipFill(
  blipFillCursor: XmlElementCursor,
  pkg: PptxPackage,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): { readonly image: Uint8Array; readonly opacity?: number } | undefined {
  return importReadBlipFill(
    blipFillCursor,
    pkg,
    location,
    diagnostics,
    (bytes) => {
      readImage(bytes);
    },
  );
}

export type ImportedSpPrVisuals = {
  readonly transform: ImportedTransform;
  readonly geometry?: Geometry;
  readonly style: ShapeVisualStyle;
};

export type ImportedSpPrAppearance = {
  readonly geometry?: Geometry;
  readonly style: ShapeVisualStyle;
};

/**
 * Reads PresentationML `<p:spPr>` geometry, fills, borders, and effects (without requiring `<a:xfrm>`).
 * Used both by full-shape import and placeholder import where transform overrides may be absent.
 */
export function readSpPrAppearance(
  spPrCursor: XmlElementCursor,
  pkg: PptxPackage,
  fallbackUnits: { readonly w: number; readonly h: number },
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
  options?: { readonly preservePlaceholderClears?: boolean },
): ImportedSpPrAppearance {
  let geometry: Geometry | undefined;
  let background: CoreVisualStyle["background"];
  let hasNoFill = false;
  let hasEmptyEffectLst = false;
  let border: ImportedLineBorder = {};
  let effects: ReturnType<typeof readEffectList> = {};

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
        hasNoFill = true;
        break;
      case "solidFill":
        background = readDrawingColor(child, location, diagnostics);
        break;
      case "gradFill":
        background = readGradientFill(child, location, diagnostics);
        break;
      case "blipFill": {
        const blip = readBlipFill(child, pkg, location, diagnostics);
        if (blip !== undefined) {
          background = {
            image: blip.image,
            ...(blip.opacity === undefined ? {} : { opacity: blip.opacity }),
          };
        }
        break;
      }
      case "ln":
        border = readLineBorder(child, location, diagnostics);
        break;
      case "effectLst":
        hasEmptyEffectLst = child.element.children.length === 0;
        effects = readEffectList(child, location, diagnostics);
        break;
      case "extLst":
        child.ignoreEmptyChild(A_NS, "ext");
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

  const resolvedBackground = background ?? (hasNoFill ? "none" : undefined);
  const resolvedBorder: ShapeBorderValue | undefined =
    border.borderColor !== undefined
      ? {
          color: border.borderColor,
          ...(border.borderWidth === undefined
            ? {}
            : { width: border.borderWidth }),
          ...(border.borderStyle === undefined
            ? {}
            : { style: border.borderStyle }),
          ...(border.startArrow === undefined
            ? {}
            : { startArrow: border.startArrow }),
          ...(border.endArrow === undefined
            ? {}
            : { endArrow: border.endArrow }),
        }
      : border.noFill === true
        ? "none"
        : undefined;
  const resolvedBoxShadow: ShapeVisualStyle["boxShadow"] =
    effects.boxShadow ??
    (options?.preservePlaceholderClears === true &&
    hasEmptyEffectLst &&
    effects.softEdge === undefined
      ? "none"
      : undefined);

  const style: ShapeVisualStyle = {
    ...(resolvedBackground === undefined
      ? {}
      : { background: resolvedBackground }),
    ...(resolvedBorder === undefined ? {} : { border: resolvedBorder }),
    ...(resolvedBoxShadow === undefined
      ? {}
      : { boxShadow: resolvedBoxShadow }),
    ...(effects.softEdge === undefined ? {} : { softEdge: effects.softEdge }),
  };

  return {
    ...(geometry === undefined ? {} : { geometry }),
    style,
  };
}

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

  const appearance = readSpPrAppearance(
    spPrCursor,
    pkg,
    { w: transform.cxEmu, h: transform.cyEmu },
    location,
    diagnostics,
  );

  const style: ShapeVisualStyle = {
    ...appearance.style,
    ...(transform.rotate === undefined ? {} : { rotate: transform.rotate }),
    ...(transform.flipH === true ? { flipH: true } : {}),
    ...(transform.flipV === true ? { flipV: true } : {}),
  };

  return {
    transform,
    ...(appearance.geometry === undefined
      ? {}
      : { geometry: appearance.geometry }),
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

function formatOoxmlRectAttrs(rect: {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}): string {
  const attrs: string[] = [];
  if (rect.l !== 0) attrs.push(`l="${rect.l}"`);
  if (rect.t !== 0) attrs.push(`t="${rect.t}"`);
  if (rect.r !== 0) attrs.push(`r="${rect.r}"`);
  if (rect.b !== 0) attrs.push(`b="${rect.b}"`);
  return attrs.length === 0 ? "" : ` ${attrs.join(" ")}`;
}

/** Emits DrawingML `<a:blipFill>` XML for embedded or external image background fills. */
export function emitImageFillXml(
  fill: ImageFillModel,
  registerResource: ((resource: CapabilityResource) => string) | undefined,
  box?: ShapeBox,
): string {
  if (registerResource === undefined) {
    invalid(
      "image-resource-failure",
      "Cannot emit image background fill without resource registrar",
    );
  }
  const relationshipId =
    fill.source.kind === "external"
      ? registerResource({
          externalTarget: fill.source.externalUrl,
          relationshipType: imageRelationship,
        })
      : registerResource({
          bytes: fill.source.bytes,
          extension: fill.source.extension,
          contentType: fill.source.contentType,
          relationshipType: imageRelationship,
        });
  const relAttr =
    fill.source.kind === "external"
      ? `r:link="${xmlEscape(relationshipId)}"`
      : `r:embed="${xmlEscape(relationshipId)}"`;
  const alpha =
    fill.opacity === undefined
      ? ""
      : `<a:alphaModFix amt="${Math.round(fill.opacity * 100000)}"/>`;
  const blipXml =
    alpha === ""
      ? `<a:blip ${relAttr}/>`
      : `<a:blip ${relAttr}>${alpha}</a:blip>`;

  let srcRectXml = "";
  let fillRectXml = "<a:fillRect/>";
  if (
    fill.source.kind === "embedded" &&
    fill.source.width > 0 &&
    fill.source.height > 0 &&
    box !== undefined &&
    box.width > 0 &&
    box.height > 0
  ) {
    const imgRatio = fill.source.width / fill.source.height;
    const boxRatio = box.width / box.height;
    if (fill.fit === "cover") {
      let l = 0;
      let t = 0;
      let r = 0;
      let b = 0;
      if (imgRatio > boxRatio) {
        const cropX = 1 - boxRatio / imgRatio;
        l = Math.round(cropX * fill.position.x * 100000);
        r = Math.round(cropX * (1 - fill.position.x) * 100000);
      } else if (imgRatio < boxRatio) {
        const cropY = 1 - imgRatio / boxRatio;
        t = Math.round(cropY * fill.position.y * 100000);
        b = Math.round(cropY * (1 - fill.position.y) * 100000);
      }
      const attrs = formatOoxmlRectAttrs({ l, t, r, b });
      if (attrs !== "") {
        srcRectXml = `<a:srcRect${attrs}/>`;
      }
    } else if (fill.fit === "contain") {
      let l = 0;
      let t = 0;
      let r = 0;
      let b = 0;
      if (imgRatio > boxRatio) {
        const insY = 1 - boxRatio / imgRatio;
        t = Math.round(insY * fill.position.y * 100000);
        b = Math.round(insY * (1 - fill.position.y) * 100000);
      } else if (imgRatio < boxRatio) {
        const insX = 1 - imgRatio / boxRatio;
        l = Math.round(insX * fill.position.x * 100000);
        r = Math.round(insX * (1 - fill.position.x) * 100000);
      }
      const attrs = formatOoxmlRectAttrs({ l, t, r, b });
      if (attrs !== "") {
        fillRectXml = `<a:fillRect${attrs}/>`;
      }
    }
  }

  return `<a:blipFill rotWithShape="1">${blipXml}${srcRectXml}<a:stretch>${fillRectXml}</a:stretch></a:blipFill>`;
}

/** Renders deterministic SVG `<pattern>` markup for embedded or external image background fills. */
export function paintImageFillSvg(
  fill: ImageFillModel,
  id: string,
  box: ShapeBox,
): string {
  const bx = roundSvgNumber(box.x);
  const by = roundSvgNumber(box.y);
  const bw = roundSvgNumber(box.width);
  const bh = roundSvgNumber(box.height);
  const imgOpacity = roundSvgNumber(fill.opacity ?? 1);

  if (fill.source.kind === "external") {
    return `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${bx}" y="${by}" width="${bw}" height="${bh}"><rect x="0" y="0" width="${bw}" height="${bh}" fill="#F1F5F9" fill-opacity="${imgOpacity}"/></pattern>`;
  }

  let pngBytes: Uint8Array;
  try {
    pngBytes = decodeImageToPngBytes(fill.source);
  } catch (error) {
    if (
      error instanceof CapabilityError &&
      error.details?.["field"] === undefined
    ) {
      throw new CapabilityError(error.code, error.message, error.path, {
        ...error.details,
        field: "style.background",
      });
    }
    throw error;
  }
  const base64 = Buffer.from(pngBytes).toString("base64");
  let drawX = 0;
  let drawY = 0;
  let drawW = box.width;
  let drawH = box.height;

  if (
    fill.fit !== "fill" &&
    fill.source.width > 0 &&
    fill.source.height > 0 &&
    box.width > 0 &&
    box.height > 0
  ) {
    const imgRatio = fill.source.width / fill.source.height;
    const boxRatio = box.width / box.height;
    if (fill.fit === "cover") {
      if (imgRatio > boxRatio) {
        drawH = box.height;
        drawW = box.height * imgRatio;
      } else {
        drawW = box.width;
        drawH = box.width / imgRatio;
      }
    } else if (fill.fit === "contain") {
      if (imgRatio > boxRatio) {
        drawW = box.width;
        drawH = box.width / imgRatio;
      } else {
        drawH = box.height;
        drawW = box.height * imgRatio;
      }
    }
    drawX = (box.width - drawW) * fill.position.x;
    drawY = (box.height - drawH) * fill.position.y;
  }

  return `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${bx}" y="${by}" width="${bw}" height="${bh}"><image x="${roundSvgNumber(drawX)}" y="${roundSvgNumber(drawY)}" width="${roundSvgNumber(drawW)}" height="${roundSvgNumber(drawH)}" preserveAspectRatio="none" opacity="${imgOpacity}" href="data:image/png;base64,${base64}"/></pattern>`;
}

/** Native canvas mechanism for DrawingML `<a:blipFill>` shape image fills. */
export const imageFillMechanism = defineCanvasMechanism({
  key: "image-fill",
  role: "canvas",
  emit(
    fill: ImageFillModel,
    registerResource: ((resource: CapabilityResource) => string) | undefined,
  ): { readonly xml: string } {
    return {
      xml: emitImageFillXml(fill, registerResource),
    };
  },
  paint(
    fill: ImageFillModel,
    context: { readonly id: string; readonly box: ShapeBox },
  ): PaintOutput {
    return {
      svg: paintImageFillSvg(fill, context.id, context.box),
    };
  },
  import: {
    target: "shape-tree",
    localName: "sp",
    matches: () => false,
    read: readFillSubElement,
  },
});
