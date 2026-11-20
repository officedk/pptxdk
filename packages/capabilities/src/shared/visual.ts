import {
  CapabilityError,
  type CapabilityResource,
  type ColorValue,
  type GradientStop,
  type LinearGradient,
  type PaintOutput,
  type RadialGradient,
  type ResolvedColor,
  type ThemeContext,
  type VisualStyle as CoreVisualStyle,
} from "@react-pptx/core/capability";
import { pxToEmu, resolvedColorXml, xmlEscape } from "@react-pptx/emit";

import {
  decodeImageToPngBytes,
  readImage,
  type LoadedImage,
} from "../image/image-decode.ts";
import {
  evaluateGeometry,
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

export type ArrowType =
  | "none"
  | "triangle"
  | "stealth"
  | "arrow"
  | "diamond"
  | "oval";

export const arrowTypes = [
  "none",
  "triangle",
  "stealth",
  "arrow",
  "diamond",
  "oval",
] as const;

export type LineEndStyle = {
  readonly startArrow?: ArrowType;
  readonly endArrow?: ArrowType;
};

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
  /** Optional resource registrar used by convenience `shapePropertiesXml` wrapper when filling with an image. */
  readonly registerResource?: (resource: CapabilityResource) => string;
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

function shadowValue(
  theme: ThemeContext,
  value: unknown,
  label: string,
): ResolvedShadow {
  if (!isRecord(value))
    invalid("invalid-shadow", `${label} must be a shadow object`);
  const unknown = Object.keys(value).find(
    (key) => !["color", "blur", "distance", "angle"].includes(key),
  );
  if (unknown !== undefined)
    invalid(
      "invalid-shadow",
      `${label} contains unsupported parameter ${unknown}`,
    );
  const blur =
    value["blur"] === undefined ? 4 : finite(value["blur"], `${label}.blur`);
  const distance =
    value["distance"] === undefined
      ? 0
      : finite(value["distance"], `${label}.distance`);
  const angle =
    value["angle"] === undefined
      ? 90
      : finite(value["angle"], `${label}.angle`);
  if (blur < 0 || distance < 0)
    invalid(
      "invalid-shadow",
      `${label}.blur and distance must be non-negative lengths`,
    );
  emu(blur, `${label}.blur`);
  emu(distance, `${label}.distance`);
  const colorValue = (value["color"] ?? {
    color: "#000000",
    alpha: 0.35,
  }) as ColorValue;
  return { color: theme.resolveColor(colorValue), blur, distance, angle };
}

export function textShadowXml(theme: ThemeContext, value: unknown): string {
  if (value === undefined || value === "none") return "";
  const shadow = shadowValue(theme, value, "textShadow");
  const direction = Math.round((((shadow.angle % 360) + 360) % 360) * 60000);
  return `<a:effectLst><a:outerShdw blurRad="${emu(shadow.blur, "textShadow.blur")}" dist="${emu(shadow.distance, "textShadow.distance")}" dir="${direction}" algn="ctr" rotWithShape="0">${resolvedColorXml(shadow.color)}</a:outerShdw></a:effectLst>`;
}

/** Resolves and validates a linear or radial gradient declaration against a ThemeContext. */
export function resolveGradientSpec(
  theme: ThemeContext,
  image: unknown,
  opacity: number | undefined,
): ResolvedGradientSpec {
  if (
    !isRecord(image) ||
    (image["type"] !== "linear-gradient" &&
      image["type"] !== "radial-gradient") ||
    !Array.isArray(image["stops"]) ||
    image["stops"].length < 2
  ) {
    invalid(
      "invalid-gradient",
      "backgroundImage must be a linear or radial gradient with at least two color stops",
    );
  }
  const stops: ResolvedGradientStop[] = image["stops"].map(
    (stop: unknown, index: number) => {
      if (!isRecord(stop)) {
        invalid("invalid-gradient", `Gradient stop ${index + 1} is invalid`);
      }
      const position = finite(
        stop["position"],
        `backgroundImage.stops[${index}].position`,
      );
      if (position < 0 || position > 1) {
        invalid(
          "invalid-gradient",
          `position for gradient stop ${index + 1} must be between 0 and 1`,
        );
      }
      const color = withOpacity(
        theme.resolveColor(stop["color"] as ColorValue),
        opacity,
      );
      return { position, color };
    },
  );
  if (image["type"] === "radial-gradient") {
    return { kind: "radial-gradient", stops };
  }
  const angle = finite(image["angle"], "backgroundImage.angle");
  return { kind: "linear-gradient", angle, stops };
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
  const styleName = style["borderStyle"];
  if (
    styleName !== undefined &&
    (typeof styleName !== "string" ||
      !["solid", "dashed", "dotted"].includes(styleName))
  ) {
    invalid(
      "invalid-visual-style",
      "borderStyle must be solid, dashed, or dotted",
    );
  }
  const startArrow = style["startArrow"];
  if (
    startArrow !== undefined &&
    (typeof startArrow !== "string" ||
      !arrowTypes.includes(startArrow as ArrowType))
  ) {
    invalid(
      "invalid-visual-style",
      "startArrow must be none, triangle, stealth, arrow, diamond, or oval",
    );
  }
  const endArrow = style["endArrow"];
  if (
    endArrow !== undefined &&
    (typeof endArrow !== "string" ||
      !arrowTypes.includes(endArrow as ArrowType))
  ) {
    invalid(
      "invalid-visual-style",
      "endArrow must be none, triangle, stealth, arrow, diamond, or oval",
    );
  }
  if (color === undefined || width === 0) {
    return undefined;
  }
  lineWidthEmu(width);
  return {
    color: withOpacity(theme.resolveColor(color as ColorValue), opacity),
    width,
    style: (styleName as "solid" | "dashed" | "dotted" | undefined) ?? "solid",
    ...(startArrow === undefined
      ? {}
      : { startArrow: startArrow as ArrowType }),
    ...(endArrow === undefined ? {} : { endArrow: endArrow as ArrowType }),
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
      : shadowValue(theme, style["boxShadow"], "boxShadow");

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
      return `<a:solidFill>${resolvedColorXml(fill.color)}</a:solidFill>`;
    case "linear-gradient":
    case "radial-gradient":
      return gradientFillXml(fill);
    case "image": {
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
      return `<a:blipFill rotWithShape="1"><a:blip r:embed="${xmlEscape(relationshipId)}">${alpha}</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
    }
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

/** Compatibility wrapper resolving and emitting `<p:spPr>` in a single step. */
export function shapePropertiesXml(
  options: ShapePropertiesOptions,
): ShapePropertiesResult {
  const model = resolveVisualModel(options);
  return emitVisualModel(model, options.registerResource, {
    omitFill: Boolean(options.image),
  });
}

/** Legacy helper preserved for slide-background callers that pass raw fill style records. */
export function fillXml(
  theme: ThemeContext,
  style: VisualStyleRecord,
  explicitFill: ColorValue | undefined,
  opacity: number | undefined,
  registerResource?: (resource: CapabilityResource) => string,
): string {
  return emitFillXml(
    resolveVisualFill(theme, style, explicitFill, opacity, false),
    registerResource,
  );
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
    case "solid": {
      const hex = subpathFillHex(fill.color.hex, fillMode);
      return {
        fillAttr: `fill="#${hex}" fill-opacity="${roundSvgNumber(fill.color.alpha ?? 1)}"`,
        overlayAttr: undefined,
        hasFill: true,
      };
    }
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
      defs.push(gradientDefSvg(fillDefId, model.fill));
      break;
    }
    case "image": {
      const pngBytes = decodeImageToPngBytes(model.fill.image);
      const base64 = Buffer.from(pngBytes).toString("base64");
      fillDefId = `${prefix}-fill`;
      const imgOpacity = roundSvgNumber(model.fill.opacity ?? 1);
      defs.push(
        `<pattern id="${fillDefId}" patternUnits="userSpaceOnUse" x="${roundSvgNumber(box.x)}" y="${roundSvgNumber(box.y)}" width="${roundSvgNumber(box.width)}" height="${roundSvgNumber(box.height)}"><image x="0" y="0" width="${roundSvgNumber(box.width)}" height="${roundSvgNumber(box.height)}" preserveAspectRatio="none" opacity="${imgOpacity}" href="data:image/png;base64,${base64}"/></pattern>`,
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
    const innerStrokeWidth =
      r >= minSide / 2 ? 0 : Math.min(r, minSide * 0.45);
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
