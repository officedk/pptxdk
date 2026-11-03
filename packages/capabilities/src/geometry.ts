import { CapabilityError } from "@react-pptx/core";

export type PresetGeometry = {
  readonly type: "preset";
  readonly name: string;
  /** 按引导名称（adj、adj1、adj2 等）索引的原始 DrawingML 调整值。 */
  readonly adjust?: Readonly<Record<string, number>>;
};

export type PathCommand =
  | { readonly op: "M"; readonly x: number; readonly y: number }
  | { readonly op: "L"; readonly x: number; readonly y: number }
  | {
      readonly op: "C";
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly op: "Q";
      readonly x1: number;
      readonly y1: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: "Z" };

export type GeometryPath = {
  readonly commands: readonly PathCommand[];
  readonly fill?: boolean;
  readonly stroke?: boolean;
};

/** 自定义几何采用局部坐标，并缩放到布局分配的边界框。 */
export type CustomGeometry = {
  readonly type: "custom";
  readonly width: number;
  readonly height: number;
  readonly paths: readonly GeometryPath[];
};

export type Geometry = string | PresetGeometry | CustomGeometry;

/** ECMA-376 ST_ShapeType 的完整 DrawingML 预设几何枚举。 */
export const presetGeometryNames = new Set(
  [
    "line lineInv triangle rtTriangle rect diamond parallelogram trapezoid nonIsoscelesTrapezoid pentagon hexagon heptagon octagon decagon dodecagon star4 star5 star6 star7 star8 star10 star12 star16 star24 star32 roundRect round1Rect round2SameRect round2DiagRect snipRoundRect snip1Rect snip2SameRect snip2DiagRect plaque ellipse teardrop homePlate chevron pieWedge pie blockArc donut noSmoking rightArrow leftArrow upArrow downArrow stripedRightArrow notchedRightArrow bentUpArrow leftRightArrow upDownArrow leftUpArrow leftRightUpArrow quadArrow leftArrowCallout rightArrowCallout upArrowCallout downArrowCallout leftRightArrowCallout upDownArrowCallout quadArrowCallout bentArrow uturnArrow circularArrow leftCircularArrow leftRightCircularArrow curvedRightArrow curvedLeftArrow curvedUpArrow curvedDownArrow swooshArrow cube can lightningBolt heart sun moon smileyFace irregularSeal1 irregularSeal2 foldedCorner bevel frame halfFrame corner diagStripe chord arc leftBracket rightBracket leftBrace rightBrace bracketPair bracePair straightConnector1 bentConnector2 bentConnector3 bentConnector4 bentConnector5 curvedConnector2 curvedConnector3 curvedConnector4 curvedConnector5 callout1 callout2 callout3 accentCallout1 accentCallout2 accentCallout3 borderCallout1 borderCallout2 borderCallout3 accentBorderCallout1 accentBorderCallout2 accentBorderCallout3 wedgeRectCallout wedgeRoundRectCallout wedgeEllipseCallout cloudCallout cloud ribbon ribbon2 ellipseRibbon ellipseRibbon2 leftRightRibbon verticalScroll horizontalScroll wave doubleWave plus flowChartProcess flowChartDecision flowChartInputOutput flowChartPredefinedProcess flowChartInternalStorage flowChartDocument flowChartMultidocument flowChartTerminator flowChartPreparation flowChartManualInput flowChartManualOperation flowChartConnector flowChartPunchedCard flowChartPunchedTape flowChartSummingJunction flowChartOr flowChartCollate flowChartSort flowChartExtract flowChartMerge flowChartOfflineStorage flowChartOnlineStorage flowChartMagneticTape flowChartMagneticDisk flowChartMagneticDrum flowChartDisplay flowChartDelay flowChartAlternateProcess flowChartOffpageConnector actionButtonBlank actionButtonHome actionButtonHelp actionButtonInformation actionButtonForwardNext actionButtonBackPrevious actionButtonEnd actionButtonBeginning actionButtonReturn actionButtonDocument actionButtonSound actionButtonMovie gear6 gear9 funnel mathPlus mathMinus mathMultiply mathDivide mathEqual mathNotEqual cornerTabs squareTabs plaqueTabs chartX chartStar chartPlus",
  ].flatMap((names) => names.split(" ")),
);

function invalidGeometry(message: string): never {
  throw new CapabilityError("invalid-geometry", message);
}

export function preset(
  name: string,
  adjust?: Readonly<Record<string, number>>,
): PresetGeometry {
  if (!presetGeometryNames.has(name)) {
    throw new CapabilityError("unsupported-geometry", `不支持预设几何 ${name}`);
  }
  if (adjust !== undefined) {
    for (const [key, value] of Object.entries(adjust)) {
      if (!/^adj(?:[1-9]\d*)?$/u.test(key) || !Number.isSafeInteger(value)) {
        invalidGeometry(`几何调整 ${key} 必须是 DrawingML 整数参数`);
      }
    }
  }
  return { type: "preset", name, ...(adjust !== undefined && { adjust }) };
}

/**
 * 原生 blockArc（环形弧段）。角度单位为度，从三点钟方向起按顺时针计算，与 DrawingML 一致；
 * 厚度表示环宽占图形较短边的比例。
 */
export function blockArc(options: {
  readonly startAngle: number;
  readonly endAngle: number;
  readonly thickness: number;
}): PresetGeometry {
  for (const key of ["startAngle", "endAngle", "thickness"] as const) {
    if (!Number.isFinite(options[key])) {
      invalidGeometry(`blockArc.${key} 必须是有限数值`);
    }
  }
  const angle = (degrees: number) =>
    Math.round((((degrees % 360) + 360) % 360) * 60000);
  return preset("blockArc", {
    adj1: angle(options.startAngle),
    adj2: angle(options.endAngle),
    // 即使数值超出预设的常见引导范围，也保留合法声明；形状编译器会报告范围警告，不会静默截断。
    adj3: Math.round(options.thickness * 100000),
  });
}

export type PathOptions = {
  /** 局部坐标空间的尺寸，类似从 0,0 开始的 SVG viewBox。 */
  readonly viewBox: readonly [width: number, height: number];
  readonly fill?: boolean;
  readonly stroke?: boolean;
  /** SVG 路径数据的填充规则，见 `path`。 */
  readonly fillRule?: "nonzero" | "evenodd";
};

/**
 * 根据 SVG 路径数据创建自定义几何。一个字符串中的多个子路径构成同一原生路径；分别绘制的路径请传入数组。
 *
 * PPTX 没有填充规则，因此不同应用程序会各自解释。默认 `fillRule: "nonzero"` 时按原样写入路径；
 * 在所有应用程序中，方向与外轮廓相反的内层子路径都会形成孔洞。设置为 `"evenodd"` 时，子路径会按嵌套
 * 深度交替方向，因此各应用程序显示偶奇填充结果。彼此重叠但未嵌套，或自相交的子路径无法通过这种方式保持一致。
 */
export function path(
  data: string | readonly string[],
  options: PathOptions,
): CustomGeometry {
  if (options.fill !== undefined && typeof options.fill !== "boolean") {
    invalidGeometry("path.fill 必须是布尔值");
  }
  if (options.stroke !== undefined && typeof options.stroke !== "boolean") {
    invalidGeometry("path.stroke 必须是布尔值");
  }
  if (
    !Array.isArray(options.viewBox) ||
    options.viewBox.length !== 2 ||
    !options.viewBox.every((value) => Number.isFinite(value) && value > 0)
  ) {
    invalidGeometry("path.viewBox 必须由两个有限正数构成");
  }
  if (
    options.fillRule !== undefined &&
    !["nonzero", "evenodd"].includes(options.fillRule)
  ) {
    invalidGeometry("path.fillRule 只能是 nonzero 或 evenodd");
  }
  if (typeof data !== "string" && !Array.isArray(data)) {
    invalidGeometry("path data 必须是 SVG 字符串或字符串数组");
  }
  const list = typeof data === "string" ? [data] : data;
  if (
    list.length === 0 ||
    list.some((item) => typeof item !== "string" || item.trim() === "")
  ) {
    invalidGeometry("path 必须至少含一个非空 SVG 路径");
  }
  return {
    type: "custom",
    width: options.viewBox[0],
    height: options.viewBox[1],
    paths: list.map((item) => {
      const commands = parseSvgPath(item);
      return {
        commands:
          options.fillRule === "evenodd"
            ? alternateByNesting(commands)
            : commands,
        ...(options.fill !== undefined && { fill: options.fill }),
        ...(options.stroke !== undefined && { stroke: options.stroke }),
      };
    }),
  };
}

type Point = { readonly x: number; readonly y: number };
type Segment = Extract<PathCommand, { op: "L" | "C" | "Q" }>;

/** 按 Z 命令拆分子路径；Z 后继续绘制时，从上一个子路径的起点开始。 */
function subpaths(commands: readonly PathCommand[]): PathCommand[][] {
  const result: PathCommand[][] = [];
  let current: PathCommand[] | undefined;
  let start: Point = { x: 0, y: 0 };
  for (const command of commands) {
    if (command.op === "M") {
      current = [command];
      result.push(current);
      start = command;
      continue;
    }
    if (current === undefined) {
      current = [{ op: "M", x: start.x, y: start.y }];
      result.push(current);
    }
    current.push(command);
    if (command.op === "Z") current = undefined;
  }
  return result;
}

/** 将子路径表示为多边形，并对曲线采样。 */
function outline(subpath: readonly PathCommand[]): Point[] {
  const points: Point[] = [];
  let x = 0;
  let y = 0;
  for (const command of subpath) {
    if (command.op === "Z") continue;
    if (command.op === "C" || command.op === "Q")
      for (let step = 1; step < 8; step += 1) {
        const t = step / 8;
        const u = 1 - t;
        points.push(
          command.op === "C"
            ? {
                x:
                  u * u * u * x +
                  3 * u * u * t * command.x1 +
                  3 * u * t * t * command.x2 +
                  t * t * t * command.x,
                y:
                  u * u * u * y +
                  3 * u * u * t * command.y1 +
                  3 * u * t * t * command.y2 +
                  t * t * t * command.y,
              }
            : {
                x: u * u * x + 2 * u * t * command.x1 + t * t * command.x,
                y: u * u * y + 2 * u * t * command.y1 + t * t * command.y,
              },
        );
      }
    points.push({ x: command.x, y: command.y });
    x = command.x;
    y = command.y;
  }
  return points;
}

function signedArea(points: readonly Point[]): number {
  let area = 0;
  points.forEach((point, index) => {
    const next = points[(index + 1) % points.length] ?? point;
    area += point.x * next.y - next.x * point.y;
  });
  return area / 2;
}

function contains(polygon: readonly Point[], point: Point): boolean {
  let inside = false;
  polygon.forEach((a, index) => {
    const b = polygon[(index + polygon.length - 1) % polygon.length] ?? a;
    if (
      a.y > point.y !== b.y > point.y &&
      point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x
    )
      inside = !inside;
  });
  return inside;
}

/** 反向绘制同一子路径。 */
function reversed(subpath: readonly PathCommand[]): PathCommand[] {
  const first = subpath[0];
  if (first === undefined || first.op !== "M") return [...subpath];
  const closed = subpath.at(-1)?.op === "Z";
  const segments = subpath.filter(
    (command): command is Segment =>
      command.op === "L" || command.op === "C" || command.op === "Q",
  );
  const starts: Point[] = [];
  let end: Point = first;
  for (const segment of segments) {
    starts.push(end);
    end = segment;
  }
  if (closed && (end.x !== first.x || end.y !== first.y)) {
    segments.push({ op: "L", x: first.x, y: first.y });
    starts.push(end);
    end = first;
  }
  const result: PathCommand[] = [{ op: "M", x: end.x, y: end.y }];
  for (let index = segments.length - 1; index >= 0; index -= 1) {
    const segment = segments[index];
    const to = starts[index];
    if (segment === undefined || to === undefined) continue;
    if (segment.op === "C")
      result.push({
        op: "C",
        x1: segment.x2,
        y1: segment.y2,
        x2: segment.x1,
        y2: segment.y1,
        x: to.x,
        y: to.y,
      });
    else if (segment.op === "Q")
      result.push({
        op: "Q",
        x1: segment.x1,
        y1: segment.y1,
        x: to.x,
        y: to.y,
      });
    else result.push({ op: "L", x: to.x, y: to.y });
  }
  if (closed) result.push({ op: "Z" });
  return result;
}

/** 外轮廓按一种方向绘制，孔洞反向绘制，孔洞中的岛屿再恢复为外轮廓方向。 */
function alternateByNesting(commands: readonly PathCommand[]): PathCommand[] {
  const parts = subpaths(commands);
  const outlines = parts.map(outline);
  return parts.flatMap((part, index) => {
    const shape = outlines[index] ?? [];
    const area = signedArea(shape);
    const probe = shape[0];
    if (area === 0 || probe === undefined) return part;
    const depth = outlines.filter(
      (other, otherIndex) =>
        otherIndex !== index && other.length > 2 && contains(other, probe),
    ).length;
    return Math.sign(area) === (depth % 2 === 0 ? 1 : -1)
      ? part
      : reversed(part);
  });
}

const tokenPattern =
  /[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/gu;

export function parseSvgPath(data: string): PathCommand[] {
  if (typeof data !== "string" || data.trim() === "") {
    invalidGeometry("SVG 路径必须是非空字符串");
  }
  const tokens: string[] = [];
  let cursor = 0;
  for (const match of data.matchAll(tokenPattern)) {
    const index = match.index;
    if (!/^[\s,]*$/u.test(data.slice(cursor, index))) {
      invalidGeometry(`SVG 路径含不支持的字符：${data.slice(cursor, index)}`);
    }
    tokens.push(match[0]);
    cursor = index + match[0].length;
  }
  if (!/^[\s,]*$/u.test(data.slice(cursor)) || tokens.length === 0) {
    invalidGeometry(`SVG 路径含不支持或缺失的参数：${data.slice(cursor)}`);
  }
  const commands: PathCommand[] = [];
  let index = 0;
  let command = "";
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let lastControl: { x: number; y: number; kind: "C" | "Q" } | undefined;
  const isCommand = (token: string | undefined) =>
    token !== undefined && /^[A-Za-z]$/u.test(token);
  const number = () => {
    const token = tokens[index];
    index += 1;
    if (token === undefined || isCommand(token))
      throw new Error(`Invalid SVG path data: ${data}`);
    const value = Number(token);
    if (!Number.isFinite(value)) invalidGeometry("SVG 路径坐标必须是有限数值");
    return value;
  };
  const flag = (name: string) => {
    const token = tokens[index];
    if (token !== "0" && token !== "1") {
      invalidGeometry(`SVG 弧线 ${name} 必须是单个 0 或 1`);
    }
    index += 1;
    return token === "1";
  };
  while (index < tokens.length) {
    if (isCommand(tokens[index])) {
      command = tokens[index] ?? "";
      index += 1;
    } else if (command === "") {
      invalidGeometry("SVG close 命令 Z 后不能跟随隐式参数");
    }
    const relative = command === command.toLowerCase();
    const upper = command.toUpperCase();
    const ox = relative ? x : 0;
    const oy = relative ? y : 0;
    switch (upper) {
      case "M": {
        x = ox + number();
        y = oy + number();
        startX = x;
        startY = y;
        commands.push({ op: "M", x, y });
        command = relative ? "l" : "L";
        lastControl = undefined;
        break;
      }
      case "L":
        x = ox + number();
        y = oy + number();
        commands.push({ op: "L", x, y });
        lastControl = undefined;
        break;
      case "H":
        x = ox + number();
        commands.push({ op: "L", x, y });
        lastControl = undefined;
        break;
      case "V":
        y = oy + number();
        commands.push({ op: "L", x, y });
        lastControl = undefined;
        break;
      case "C": {
        const x1 = ox + number();
        const y1 = oy + number();
        const x2 = ox + number();
        const y2 = oy + number();
        x = ox + number();
        y = oy + number();
        commands.push({ op: "C", x1, y1, x2, y2, x, y });
        lastControl = { x: x2, y: y2, kind: "C" };
        break;
      }
      case "S": {
        const x1 = lastControl?.kind === "C" ? 2 * x - lastControl.x : x;
        const y1 = lastControl?.kind === "C" ? 2 * y - lastControl.y : y;
        const x2 = ox + number();
        const y2 = oy + number();
        x = ox + number();
        y = oy + number();
        commands.push({ op: "C", x1, y1, x2, y2, x, y });
        lastControl = { x: x2, y: y2, kind: "C" };
        break;
      }
      case "Q": {
        const x1 = ox + number();
        const y1 = oy + number();
        x = ox + number();
        y = oy + number();
        commands.push({ op: "Q", x1, y1, x, y });
        lastControl = { x: x1, y: y1, kind: "Q" };
        break;
      }
      case "T": {
        const x1 = lastControl?.kind === "Q" ? 2 * x - lastControl.x : x;
        const y1 = lastControl?.kind === "Q" ? 2 * y - lastControl.y : y;
        x = ox + number();
        y = oy + number();
        commands.push({ op: "Q", x1, y1, x, y });
        lastControl = { x: x1, y: y1, kind: "Q" };
        break;
      }
      case "A": {
        const rx = number();
        const ry = number();
        const rotation = number();
        const largeArc = flag("large-arc-flag");
        const sweep = flag("sweep-flag");
        const endX = ox + number();
        const endY = oy + number();
        commands.push(
          ...arcToCubics(x, y, rx, ry, rotation, largeArc, sweep, endX, endY),
        );
        x = endX;
        y = endY;
        lastControl = undefined;
        break;
      }
      case "Z":
        commands.push({ op: "Z" });
        x = startX;
        y = startY;
        lastControl = undefined;
        command = "";
        break;
      default:
        invalidGeometry(`SVG 路径命令 ${command} 不受支持`);
    }
  }
  if (commands[0]?.op !== "M") invalidGeometry("SVG 路径必须以移动命令 M 开始");
  return commands;
}

/** 按 SVG 实现说明 F.6 将椭圆弧转换为三次贝塞尔曲线。 */
function arcToCubics(
  x1: number,
  y1: number,
  rxInput: number,
  ryInput: number,
  rotationDegrees: number,
  largeArc: boolean,
  sweep: boolean,
  x2: number,
  y2: number,
): PathCommand[] {
  if (x1 === x2 && y1 === y2) return [];
  if (rxInput === 0 || ryInput === 0) return [{ op: "L", x: x2, y: y2 }];
  const phi = (rotationDegrees * Math.PI) / 180;
  const cos = Math.cos(phi);
  const sin = Math.sin(phi);
  const dx = (x1 - x2) / 2;
  const dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  let rx = Math.abs(rxInput);
  let ry = Math.abs(ryInput);
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const numerator =
    rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const denominator = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  const factor =
    (largeArc === sweep ? -1 : 1) *
    Math.sqrt(Math.max(0, numerator / denominator));
  const cxp = (factor * rx * y1p) / ry;
  const cyp = (-factor * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number) => {
    const sign = ux * vy - uy * vx < 0 ? -1 : 1;
    const dot = (ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy));
    return sign * Math.acos(Math.min(1, Math.max(-1, dot)));
  };
  const theta1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let delta = angle(
    (x1p - cxp) / rx,
    (y1p - cyp) / ry,
    (-x1p - cxp) / rx,
    (-y1p - cyp) / ry,
  );
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.ceil(Math.abs(delta) / (Math.PI / 2));
  const step = delta / segments;
  const kappa = (4 / 3) * Math.tan(step / 4);
  const point = (theta: number) => ({
    x: cx + rx * Math.cos(theta) * cos - ry * Math.sin(theta) * sin,
    y: cy + rx * Math.cos(theta) * sin + ry * Math.sin(theta) * cos,
  });
  const derivative = (theta: number) => ({
    x: -rx * Math.sin(theta) * cos - ry * Math.cos(theta) * sin,
    y: -rx * Math.sin(theta) * sin + ry * Math.cos(theta) * cos,
  });
  const result: PathCommand[] = [];
  for (let segment = 0; segment < segments; segment += 1) {
    const a = theta1 + segment * step;
    const b = a + step;
    const start = point(a);
    const end = segment === segments - 1 ? { x: x2, y: y2 } : point(b);
    const startDerivative = derivative(a);
    const endDerivative = derivative(b);
    result.push({
      op: "C",
      x1: start.x + kappa * startDerivative.x,
      y1: start.y + kappa * startDerivative.y,
      x2: end.x - kappa * endDerivative.x,
      y2: end.y - kappa * endDerivative.y,
      x: end.x,
      y: end.y,
    });
  }
  return result;
}
