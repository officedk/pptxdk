import {
  CapabilityError,
  pxToEmu,
  resolvedColorXml,
  type ColorValue,
  type GeometryBox,
  type GradientStop,
  type LinearGradient,
  type PreviewTransform,
  type RadialGradient,
  type ResolvedColor,
  type ThemeContext,
  type VisualStyle as CoreVisualStyle,
} from "@react-pptx/core/capability";

import {
  presetGeometryNames,
  type CustomGeometry,
  type Geometry,
  type PathCommand,
} from "./geometry.ts";

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

export type VisualWarning = { readonly code: string; readonly message: string };
export type ShapeBox = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
};
export type VisualStyleRecord = Readonly<Record<string, unknown>>;

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
};

export type ShapePropertiesResult = {
  readonly xml: string;
  readonly opacity?: number;
  readonly warnings: readonly VisualWarning[];
  readonly hasVisibleStyle: boolean;
  readonly geometryName: string;
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

function shadowValue(
  theme: ThemeContext,
  value: unknown,
  label: string,
): {
  readonly color: ResolvedColor;
  readonly blur: number;
  readonly distance: number;
  readonly angle: number;
} {
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

function transformXml(
  box: ShapeBox,
  rotation: number | undefined,
  flipH: boolean | undefined,
  flipV: boolean | undefined,
): string {
  for (const [name, value] of [
    ["flipH", flipH],
    ["flipV", flipV],
  ] as const) {
    if (value !== undefined && typeof value !== "boolean")
      invalid("invalid-visual-style", `${name} must be a boolean`);
  }
  return `<a:xfrm${rotation === undefined || rotation === 0 ? "" : ` rot="${rotation}"`}${flipH ? ' flipH="1"' : ""}${flipV ? ' flipV="1"' : ""}><a:off x="${emu(box.x, "x", true, true)}" y="${emu(box.y, "y", true, true)}"/><a:ext cx="${emu(box.width, "width", false)}" cy="${emu(box.height, "height", false)}"/></a:xfrm>`;
}

function rotationUnits(style: VisualStyleRecord): number | undefined {
  if (style["rotate"] === undefined) return undefined;
  const degrees = finite(style["rotate"], "rotate");
  const units = Math.round((((degrees % 360) + 360) % 360) * 60000);
  if (!Number.isSafeInteger(units) || units < 0 || units > 21_600_000)
    invalid(
      "invalid-visual-style",
      "rotate exceeds writable DrawingML range after conversion",
    );
  return units;
}

function point(value: number, label: string): string {
  const coordinate = Math.round(finite(value, label) * 1000);
  if (
    !Number.isSafeInteger(coordinate) ||
    Math.abs(coordinate) > 2_147_483_647
  ) {
    invalid(
      "invalid-geometry",
      `${label} exceeds DrawingML path coordinate range after conversion`,
    );
  }
  return String(coordinate);
}

function pathCommandXml(command: PathCommand): string {
  switch (command.op) {
    case "M":
      return `<a:moveTo><a:pt x="${point(command.x, "path.x")}" y="${point(command.y, "path.y")}"/></a:moveTo>`;
    case "L":
      return `<a:lnTo><a:pt x="${point(command.x, "path.x")}" y="${point(command.y, "path.y")}"/></a:lnTo>`;
    case "C":
      return `<a:cubicBezTo><a:pt x="${point(command.x1, "path.x1")}" y="${point(command.y1, "path.y1")}"/><a:pt x="${point(command.x2, "path.x2")}" y="${point(command.y2, "path.y2")}"/><a:pt x="${point(command.x, "path.x")}" y="${point(command.y, "path.y")}"/></a:cubicBezTo>`;
    case "Q":
      return `<a:quadBezTo><a:pt x="${point(command.x1, "path.x1")}" y="${point(command.y1, "path.y1")}"/><a:pt x="${point(command.x, "path.x")}" y="${point(command.y, "path.y")}"/></a:quadBezTo>`;
    case "Z":
      return "<a:close/>";
    default:
      return invalid(
        "invalid-geometry",
        "Custom path contains unsupported command",
      );
  }
}

function customGeometryXml(geometry: CustomGeometry): string {
  if (!isRecord(geometry) || !Array.isArray(geometry.paths))
    invalid("invalid-geometry", "Custom geometry must contain a paths array");
  const width = finite(geometry.width, "geometry.width");
  const height = finite(geometry.height, "geometry.height");
  if (width <= 0 || height <= 0 || geometry.paths.length === 0) {
    invalid(
      "invalid-geometry",
      "Custom geometry must have a positive viewBox and at least one path",
    );
  }
  const widthUnits = point(width, "geometry.width");
  const heightUnits = point(height, "geometry.height");
  if (Number(widthUnits) === 0 || Number(heightUnits) === 0) {
    invalid(
      "invalid-geometry",
      "Custom geometry viewBox must retain positive width and height in path coordinate space after conversion",
    );
  }
  const pathList = geometry.paths
    .map((path, index) => {
      if (
        !isRecord(path) ||
        !Array.isArray(path.commands) ||
        path.commands.length === 0
      ) {
        invalid(
          "invalid-geometry",
          `Custom path ${index + 1} must contain commands`,
        );
      }
      if (
        (path.fill !== undefined && typeof path.fill !== "boolean") ||
        (path.stroke !== undefined && typeof path.stroke !== "boolean")
      )
        invalid(
          "invalid-geometry",
          `fill/stroke for custom path ${index + 1} must be boolean`,
        );
      const commands = path.commands.map(pathCommandXml).join("");
      return `<a:path w="${widthUnits}" h="${heightUnits}"${path.fill === false ? ' fill="none"' : ""}${path.stroke === false ? ' stroke="0"' : ""}>${commands}</a:path>`;
    })
    .join("");
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${pathList}</a:pathLst></a:custGeom>`;
}

function geometryXml(
  value: Geometry | undefined,
  radius: number | undefined,
  box: ShapeBox,
  warnings: VisualWarning[],
): { readonly xml: string; readonly name: string } {
  if (isRecord(value) && value["type"] === "custom") {
    if (radius !== undefined && radius > 0)
      invalid(
        "unsupported-style-combination",
        "borderRadius cannot be used together with custom geometry",
      );
    return { xml: customGeometryXml(value), name: "custom" };
  }
  let name: string;
  let adjust: Record<string, number> = {};
  if (value === undefined)
    name = radius !== undefined && radius > 0 ? "roundRect" : "rect";
  else if (typeof value === "string") name = value;
  else if (
    isRecord(value) &&
    value["type"] === "preset" &&
    typeof value["name"] === "string"
  ) {
    name = value["name"];
    if (value["adjust"] !== undefined) {
      if (!isRecord(value["adjust"]))
        invalid("invalid-geometry", "preset.adjust must be an object");
      adjust = { ...value["adjust"] };
    }
  } else
    invalid(
      "invalid-geometry",
      "geometry must be a preset name, preset(...), or custom path",
    );
  if (!presetGeometryNames.has(name))
    invalid("unsupported-geometry", `Unsupported preset geometry ${name}`);
  if (radius !== undefined && radius > 0) {
    if (name !== "rect" && name !== "roundRect")
      invalid(
        "unsupported-style-combination",
        "borderRadius is only applicable to rect or roundRect geometry",
      );
    name = "roundRect";
    if (adjust["adj"] === undefined) {
      const raw = Math.round(
        (radius / Math.min(box.width, box.height)) * 100000,
      );
      adjust["adj"] = raw;
      if (raw > 50000)
        warnings.push({
          code: "geometry-adjustment-range",
          message:
            "borderRadius exceeds half of the shorter side; preserved original radius and wrote native adjust value",
        });
    }
  }
  for (const [key, amount] of Object.entries(adjust)) {
    if (!/^adj(?:[1-9]\d*)?$/u.test(key) || !Number.isSafeInteger(amount))
      invalid(
        "invalid-geometry",
        `preset.adjust.${key} must be a finite integer`,
      );
    if (
      name === "blockArc" &&
      (((key === "adj1" || key === "adj2") &&
        (amount < 0 || amount > 21_600_000)) ||
        (key === "adj3" && (amount < 0 || amount > 50_000)))
    ) {
      warnings.push({
        code: "geometry-adjustment-range",
        message: `blockArc.${key} exceeds common preset range; preserved original adjust value`,
      });
    }
  }
  const values = Object.entries(adjust)
    .map(([key, amount]) => `<a:gd name="${key}" fmla="val ${amount}"/>`)
    .join("");
  return {
    xml: `<a:prstGeom prst="${name}"><a:avLst>${values}</a:avLst></a:prstGeom>`,
    name,
  };
}

export function fillXml(
  theme: ThemeContext,
  style: VisualStyleRecord,
  explicitFill: ColorValue | undefined,
  opacity: number | undefined,
): string {
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
  if (explicitFill !== undefined)
    return `<a:solidFill>${resolvedColorXml(withOpacity(theme.resolveColor(explicitFill), opacity))}</a:solidFill>`;
  const background = style["backgroundColor"];
  const image = style["backgroundImage"];
  if (background !== undefined && image !== undefined)
    invalid(
      "unsupported-style-combination",
      "backgroundColor and backgroundImage cannot be set simultaneously",
    );
  if (image !== undefined) {
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
    const stops = image["stops"]
      .map((stop: unknown, index: number) => {
        if (!isRecord(stop))
          invalid("invalid-gradient", `Gradient stop ${index + 1} is invalid`);
        const position = finite(
          stop["position"],
          `backgroundImage.stops[${index}].position`,
        );
        if (position < 0 || position > 1)
          invalid(
            "invalid-gradient",
            `position for gradient stop ${index + 1} must be between 0 and 1`,
          );
        const color = withOpacity(
          theme.resolveColor(stop["color"] as ColorValue),
          opacity,
        );
        return `<a:gs pos="${Math.round(position * 100000)}">${resolvedColorXml(color)}</a:gs>`;
      })
      .join("");
    if (image["type"] === "radial-gradient")
      return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
    const angle = finite(image["angle"], "backgroundImage.angle");
    const drawingAngle = Math.round(
      ((((angle - 90) % 360) + 360) % 360) * 60000,
    );
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${drawingAngle}" scaled="0"/></a:gradFill>`;
  }
  return background === undefined
    ? "<a:noFill/>"
    : `<a:solidFill>${resolvedColorXml(withOpacity(theme.resolveColor(background as ColorValue), opacity))}</a:solidFill>`;
}

function lineXml(
  theme: ThemeContext,
  style: VisualStyleRecord,
  opacity: number | undefined,
): string {
  const color = style["borderColor"];
  const width =
    style["borderWidth"] === undefined
      ? 1
      : finite(style["borderWidth"], "borderWidth");
  if (width < 0)
    invalid(
      "invalid-visual-style",
      "borderWidth must be a non-negative number",
    );
  const styleName = style["borderStyle"];
  if (
    styleName !== undefined &&
    (typeof styleName !== "string" ||
      !["solid", "dashed", "dotted"].includes(styleName))
  )
    invalid(
      "invalid-visual-style",
      "borderStyle must be solid, dashed, or dotted",
    );
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
  if (color === undefined || width === 0) return "<a:ln><a:noFill/></a:ln>";
  const dash =
    styleName === "dashed"
      ? "dash"
      : styleName === "dotted"
        ? "sysDot"
        : undefined;
  const headEnd =
    startArrow === undefined ? "" : `<a:headEnd type="${startArrow}"/>`;
  const tailEnd =
    endArrow === undefined ? "" : `<a:tailEnd type="${endArrow}"/>`;
  return `<a:ln w="${lineWidthEmu(width)}"><a:solidFill>${resolvedColorXml(withOpacity(theme.resolveColor(color as ColorValue), opacity))}</a:solidFill>${dash === undefined ? "" : `<a:prstDash val="${dash}"/>`}${headEnd}${tailEnd}</a:ln>`;
}

function effectsXml(theme: ThemeContext, style: VisualStyleRecord): string {
  const items: string[] = [];
  if (style["boxShadow"] !== undefined) {
    const shadow = shadowValue(theme, style["boxShadow"], "boxShadow");
    const dir = Math.round((((shadow.angle % 360) + 360) % 360) * 60000);
    items.push(
      `<a:outerShdw blurRad="${emu(shadow.blur, "boxShadow.blur")}" dist="${emu(shadow.distance, "boxShadow.distance")}" dir="${dir}" algn="ctr" rotWithShape="0">${resolvedColorXml(shadow.color)}</a:outerShdw>`,
    );
  }
  if (style["softEdge"] !== undefined) {
    const softEdge = finite(style["softEdge"], "softEdge");
    if (softEdge < 0)
      invalid("invalid-visual-style", "softEdge must be a non-negative length");
    if (softEdge > 0)
      items.push(`<a:softEdge rad="${emu(softEdge, "softEdge", false)}"/>`);
  }
  return items.length === 0
    ? ""
    : `<a:effectLst>${items.join("")}</a:effectLst>`;
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

/** Shared serializer for shape properties of Shape, Image, Text, and visible Box. */
export function shapePropertiesXml(
  options: ShapePropertiesOptions,
): ShapePropertiesResult {
  const { theme, style, box } = options;
  const warnings: VisualWarning[] = [];
  const opacity =
    style["opacity"] === undefined
      ? undefined
      : finite(style["opacity"], "opacity");
  if (opacity !== undefined && (opacity < 0 || opacity > 1))
    invalid("invalid-visual-style", "opacity must be a number between 0 and 1");
  if (
    options.image &&
    (style["backgroundColor"] !== undefined ||
      style["backgroundImage"] !== undefined ||
      style["borderRadius"] !== undefined)
  )
    invalid(
      "unsupported-style-combination",
      "Image does not support background fill or borderRadius",
    );
  const radius =
    style["borderRadius"] === undefined
      ? undefined
      : finite(style["borderRadius"], "borderRadius");
  if (radius !== undefined && radius < 0)
    invalid(
      "invalid-visual-style",
      "borderRadius must be a non-negative length",
    );
  const visualFill = options.image
    ? ""
    : fillXml(theme, style, options.fill, opacity);
  const line = lineXml(theme, style, opacity);
  const effects = effectsXml(theme, style);
  const hasVisibleStyle =
    (options.fill !== undefined && !options.image) || hasVisibleBoxStyle(style);
  const geometry = geometryXml(
    options.geometry,
    options.layoutOnlyWhenInvisible && !hasVisibleStyle ? undefined : radius,
    box,
    warnings,
  );
  // Transparent layout Box does not write native transforms, but invalid rotation remains an unwritable visual declaration.
  const rotation = rotationUnits(style);
  const transform =
    options.layoutOnlyWhenInvisible && !hasVisibleStyle
      ? ""
      : transformXml(box, rotation, options.flipH, options.flipV);
  return {
    xml: `<p:spPr>${transform}${geometry.xml}${visualFill}${line}${effects}</p:spPr>`,
    ...(opacity === undefined ? {} : { opacity }),
    warnings,
    hasVisibleStyle,
    geometryName: geometry.name,
  };
}

const unsupportedBoxPreviewStyleKeys = [
  "backgroundImage",
  "boxShadow",
  "softEdge",
  "borderRadius",
  "startArrow",
  "endArrow",
] as const;

export type BoxPreviewVisuals = {
  readonly attributes: string;
  readonly hasFillOrStroke: boolean;
  readonly svgTransform: string;
  readonly transform: PreviewTransform;
};

/**
 * Validates and resolves visual styles for static SVG preview across rect-backed objects
 * (Shape, Text, and visible Box). Throws located CapabilityError codes for unsupported effects.
 */
export function resolveBoxPreviewVisuals(options: {
  readonly label: string;
  readonly visualStyle: Readonly<Record<string, unknown>>;
  readonly box: GeometryBox;
  readonly theme: ThemeContext;
  readonly fill?: ColorValue | undefined;
  readonly flipH?: boolean | undefined;
  readonly flipV?: boolean | undefined;
}): BoxPreviewVisuals {
  const { label, visualStyle, box, theme, fill, flipH, flipV } = options;
  for (const key of unsupportedBoxPreviewStyleKeys) {
    if (visualStyle[key] !== undefined) {
      throw new CapabilityError(
        "preview-unsupported-style",
        `Preview does not support ${label} ${key}`,
      );
    }
  }
  const borderStyle = visualStyle["borderStyle"];
  if (borderStyle !== undefined && borderStyle !== "solid") {
    throw new CapabilityError(
      "preview-unsupported-border-style",
      "Preview does not yet support non-solid borders",
    );
  }
  const fillValue =
    fill ?? (visualStyle["backgroundColor"] as ColorValue | undefined);
  const resolvedFill =
    fillValue === undefined ? undefined : theme.resolveColor(fillValue);
  const strokeValue = visualStyle["borderColor"] as ColorValue | undefined;
  const borderWidth =
    visualStyle["borderWidth"] === undefined
      ? 1
      : Number(visualStyle["borderWidth"]);
  const resolvedStroke =
    strokeValue === undefined || borderWidth === 0
      ? undefined
      : theme.resolveColor(strokeValue);
  const opacity = Number(visualStyle["opacity"] ?? 1);
  const attributes = `fill="${resolvedFill === undefined ? "none" : `#${resolvedFill.hex}`}" fill-opacity="${(resolvedFill?.alpha ?? 1) * opacity}" stroke="${resolvedStroke === undefined ? "none" : `#${resolvedStroke.hex}`}" stroke-opacity="${(resolvedStroke?.alpha ?? 1) * opacity}" stroke-width="${resolvedStroke === undefined ? 0 : borderWidth}"`;

  const angle =
    typeof visualStyle["rotate"] === "number" ? visualStyle["rotate"] : 0;
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const flipX = Boolean(flipH);
  const flipY = Boolean(flipV);
  const svgTransform = `translate(${cx} ${cy}) rotate(${angle}) scale(${flipX ? -1 : 1} ${flipY ? -1 : 1}) translate(${-cx} ${-cy})`;

  return {
    attributes,
    hasFillOrStroke: resolvedFill !== undefined || resolvedStroke !== undefined,
    svgTransform,
    transform: {
      box,
      angle,
      flipH: flipX,
      flipV: flipY,
    },
  };
}
