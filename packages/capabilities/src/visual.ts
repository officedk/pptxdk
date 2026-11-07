import {
  CapabilityError,
  pxToEmu,
  resolvedColorXml,
  type ColorValue,
  type GradientStop,
  type LinearGradient,
  type RadialGradient,
  type ResolvedColor,
  type ThemeContext,
} from "@react-pptx/core";

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

export type ShapePropertiesOptions = {
  readonly theme: ThemeContext;
  readonly style: VisualStyleRecord;
  readonly box: ShapeBox;
  readonly geometry?: Geometry;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** 形状能力旧有的顶层填充属性。 */
  readonly fill?: ColorValue;
  readonly image?: boolean;
  /** 无可见视觉的布局容器不要求写出原生变换。 */
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
    invalid("invalid-visual-style", `${label} 必须是有限数值`);
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
      `${label} 换算后必须处于 DrawingML 可写范围`,
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
      "borderWidth 换算后必须在 1 到 20116800 EMU 的 DrawingML 线宽范围内",
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
  if (!isRecord(value)) invalid("invalid-shadow", `${label} 必须是阴影对象`);
  const unknown = Object.keys(value).find(
    (key) => !["color", "blur", "distance", "angle"].includes(key),
  );
  if (unknown !== undefined)
    invalid("invalid-shadow", `${label} 含未实现参数 ${unknown}`);
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
    invalid("invalid-shadow", `${label}.blur 与 distance 必须是非负长度`);
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
      invalid("invalid-visual-style", `${name} 必须是布尔值`);
  }
  return `<a:xfrm${rotation === undefined || rotation === 0 ? "" : ` rot="${rotation}"`}${flipH ? ' flipH="1"' : ""}${flipV ? ' flipV="1"' : ""}><a:off x="${emu(box.x, "x", true, true)}" y="${emu(box.y, "y", true, true)}"/><a:ext cx="${emu(box.width, "width", false)}" cy="${emu(box.height, "height", false)}"/></a:xfrm>`;
}

function rotationUnits(style: VisualStyleRecord): number | undefined {
  if (style["rotate"] === undefined) return undefined;
  const degrees = finite(style["rotate"], "rotate");
  const units = Math.round((((degrees % 360) + 360) % 360) * 60000);
  if (!Number.isSafeInteger(units) || units < 0 || units > 21_600_000)
    invalid("invalid-visual-style", "rotate 换算后超出 DrawingML 可写范围");
  return units;
}

function point(value: number, label: string): string {
  const coordinate = Math.round(finite(value, label) * 1000);
  if (
    !Number.isSafeInteger(coordinate) ||
    Math.abs(coordinate) > 2_147_483_647
  ) {
    invalid("invalid-geometry", `${label} 换算后超出 DrawingML 路径坐标范围`);
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
      return invalid("invalid-geometry", "自定义路径含不支持的命令");
  }
}

function customGeometryXml(geometry: CustomGeometry): string {
  if (!isRecord(geometry) || !Array.isArray(geometry.paths))
    invalid("invalid-geometry", "自定义 geometry 必须包含路径数组");
  const width = finite(geometry.width, "geometry.width");
  const height = finite(geometry.height, "geometry.height");
  if (width <= 0 || height <= 0 || geometry.paths.length === 0) {
    invalid("invalid-geometry", "自定义几何必须有正 viewBox 和至少一个路径");
  }
  const widthUnits = point(width, "geometry.width");
  const heightUnits = point(height, "geometry.height");
  if (Number(widthUnits) === 0 || Number(heightUnits) === 0) {
    invalid(
      "invalid-geometry",
      "自定义几何 viewBox 换算后必须在路径坐标空间内保留正宽和正高",
    );
  }
  const pathList = geometry.paths
    .map((path, index) => {
      if (
        !isRecord(path) ||
        !Array.isArray(path.commands) ||
        path.commands.length === 0
      ) {
        invalid("invalid-geometry", `自定义路径 ${index + 1} 必须包含命令`);
      }
      if (
        (path.fill !== undefined && typeof path.fill !== "boolean") ||
        (path.stroke !== undefined && typeof path.stroke !== "boolean")
      )
        invalid(
          "invalid-geometry",
          `自定义路径 ${index + 1} 的 fill/stroke 必须是布尔值`,
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
        "borderRadius 不能与自定义 geometry 同时使用",
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
        invalid("invalid-geometry", "preset.adjust 必须是对象");
      adjust = { ...value["adjust"] };
    }
  } else
    invalid(
      "invalid-geometry",
      "geometry 必须是预设名称、preset(...) 或自定义路径",
    );
  if (!presetGeometryNames.has(name))
    invalid("unsupported-geometry", `不支持预设几何 ${name}`);
  if (radius !== undefined && radius > 0) {
    if (name !== "rect" && name !== "roundRect")
      invalid(
        "unsupported-style-combination",
        "borderRadius 只适用于 rect 或 roundRect geometry",
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
            "borderRadius 超过短边的一半；已保留原始半径并写入原生调整值",
        });
    }
  }
  for (const [key, amount] of Object.entries(adjust)) {
    if (!/^adj(?:[1-9]\d*)?$/u.test(key) || !Number.isSafeInteger(amount))
      invalid("invalid-geometry", `preset.adjust.${key} 必须是有限整数`);
    if (
      name === "blockArc" &&
      (((key === "adj1" || key === "adj2") &&
        (amount < 0 || amount > 21_600_000)) ||
        (key === "adj3" && (amount < 0 || amount > 50_000)))
    ) {
      warnings.push({
        code: "geometry-adjustment-range",
        message: `blockArc.${key} 超出常用预设范围；已保留原始调整值`,
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
      "Shape.fill 不能与 style.backgroundColor 或 backgroundImage 同时设置",
    );
  }
  if (explicitFill !== undefined)
    return `<a:solidFill>${resolvedColorXml(withOpacity(theme.resolveColor(explicitFill), opacity))}</a:solidFill>`;
  const background = style["backgroundColor"];
  const image = style["backgroundImage"];
  if (background !== undefined && image !== undefined)
    invalid(
      "unsupported-style-combination",
      "backgroundColor 与 backgroundImage 不能同时设置",
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
        "backgroundImage 必须是含至少两个色标的线性或径向渐变",
      );
    }
    const stops = image["stops"]
      .map((stop: unknown, index: number) => {
        if (!isRecord(stop))
          invalid("invalid-gradient", `渐变色标 ${index + 1} 无效`);
        const position = finite(
          stop["position"],
          `backgroundImage.stops[${index}].position`,
        );
        if (position < 0 || position > 1)
          invalid(
            "invalid-gradient",
            `渐变色标 ${index + 1} 的 position 必须位于 0 到 1`,
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
  if (width < 0) invalid("invalid-visual-style", "borderWidth 必须是非负数");
  const styleName = style["borderStyle"];
  if (
    styleName !== undefined &&
    (typeof styleName !== "string" ||
      !["solid", "dashed", "dotted"].includes(styleName))
  )
    invalid(
      "invalid-visual-style",
      "borderStyle 只能是 solid、dashed 或 dotted",
    );
  if (color === undefined || width === 0) return "<a:ln><a:noFill/></a:ln>";
  const dash =
    styleName === "dashed"
      ? "dash"
      : styleName === "dotted"
        ? "sysDot"
        : undefined;
  return `<a:ln w="${lineWidthEmu(width)}"><a:solidFill>${resolvedColorXml(withOpacity(theme.resolveColor(color as ColorValue), opacity))}</a:solidFill>${dash === undefined ? "" : `<a:prstDash val="${dash}"/>`}</a:ln>`;
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
      invalid("invalid-visual-style", "softEdge 必须是非负长度");
    if (softEdge > 0)
      items.push(`<a:softEdge rad="${emu(softEdge, "softEdge", false)}"/>`);
  }
  return items.length === 0
    ? ""
    : `<a:effectLst>${items.join("")}</a:effectLst>`;
}

/** Shape、Image、Text 与可见 Box 形状属性共用的序列化器。 */
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
    invalid("invalid-visual-style", "opacity 必须是 0 到 1 之间的数值");
  if (
    options.image &&
    (style["backgroundColor"] !== undefined ||
      style["backgroundImage"] !== undefined ||
      style["borderRadius"] !== undefined)
  )
    invalid(
      "unsupported-style-combination",
      "Image 不支持背景填充或 borderRadius",
    );
  const radius =
    style["borderRadius"] === undefined
      ? undefined
      : finite(style["borderRadius"], "borderRadius");
  if (radius !== undefined && radius < 0)
    invalid("invalid-visual-style", "borderRadius 必须是非负长度");
  const visualFill = options.image
    ? ""
    : fillXml(theme, style, options.fill, opacity);
  const line = lineXml(theme, style, opacity);
  const effects = effectsXml(theme, style);
  const borderWidth =
    style["borderWidth"] === undefined ? 1 : (style["borderWidth"] as number);
  const hasVisibleStyle =
    (options.fill !== undefined && !options.image) ||
    style["backgroundColor"] !== undefined ||
    style["backgroundImage"] !== undefined ||
    (style["borderColor"] !== undefined && borderWidth > 0) ||
    style["boxShadow"] !== undefined;
  const geometry = geometryXml(
    options.geometry,
    options.layoutOnlyWhenInvisible && !hasVisibleStyle ? undefined : radius,
    box,
    warnings,
  );
  // 透明布局 Box 不写原生变换，但非法旋转仍是不可写的视觉声明。
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
