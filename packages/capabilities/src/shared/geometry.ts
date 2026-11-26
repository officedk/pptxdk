import {
  CapabilityError,
  customGeometrySchema,
  geometryAdjustKeySchema,
  geometryAdjustSchema,
  geometryPathSchema,
  geometrySchema,
  pathCommandSchema,
  presetGeometryNameSchema,
  presetGeometrySchema,
  presetGeometryTuple,
  type CustomGeometry,
  type Geometry,
  type Rect,
  type GeometryPath,
  type PathCommand,
  type PresetGeometry,
  type PresetGeometryName,
} from "@pptxdk/core/capability";
import { pxToEmu } from "@pptxdk/emit";
import {
  A_NS,
  ooxmlBoolSchema,
  ooxmlIntSchema,
  ooxmlPositiveIntSchema,
  type ImportDiagnostic,
  type ImportLocation,
  type XmlElementCursor,
} from "@pptxdk/import";
import { z } from "zod";

import {
  presetDefinitions,
  type PresetPathFillMode,
  type PresetShapeDefinition,
} from "./preset-definitions.ts";

export {
  customGeometrySchema,
  geometryAdjustKeySchema,
  geometryAdjustSchema,
  geometryPathSchema,
  geometrySchema,
  pathCommandSchema,
  presetGeometryNameSchema,
  presetGeometrySchema,
  presetGeometryTuple,
};
export type {
  CustomGeometry,
  Geometry,
  GeometryPath,
  PathCommand,
  PresetGeometry,
  PresetGeometryName,
  PresetPathFillMode,
};

/** Closed resolved geometry model consumed exhaustively by both emit and paint. */
export type ResolvedGeometry =
  | {
      readonly kind: "preset";
      readonly name: PresetGeometryName;
      /** Ordered adjust overrides already merged with borderRadius -> roundRect.adj when applicable. */
      readonly adjust: Readonly<Record<string, number>>;
    }
  | {
      readonly kind: "custom";
      readonly width: number;
      readonly height: number;
      readonly paths: readonly {
        readonly commands: readonly PathCommand[];
        readonly fill: boolean;
        readonly stroke: boolean;
      }[];
    };

export type GeometryWarning = {
  readonly code: string;
  readonly message: string;
};

export type EvaluatedGeometryPath = {
  readonly d: string;
  readonly commands: readonly PathCommand[];
  readonly fillMode: "norm" | PresetPathFillMode;
  readonly stroke: boolean;
};

export type EvaluatedGeometry = {
  readonly paths: readonly EvaluatedGeometryPath[];
  readonly textRect: Rect;
};

/** Complete DrawingML preset geometry enumeration from ECMA-376 ST_ShapeType. */
export const presetGeometryNames: ReadonlySet<string> = new Set(
  presetGeometryTuple,
);

function invalidGeometry(message: string): never {
  throw new CapabilityError("invalid-geometry", message);
}

export function preset(
  name: PresetGeometryName,
  adjust?: Readonly<Record<string, number>>,
): PresetGeometry {
  if (!presetGeometryNameSchema.safeParse(name).success) {
    throw new CapabilityError(
      "unsupported-geometry",
      `Unsupported preset geometry ${String(name)}`,
    );
  }
  if (adjust !== undefined) {
    const parsedAdjust = geometryAdjustSchema.safeParse(adjust);
    if (!parsedAdjust.success) {
      const key = String(parsedAdjust.error.issues[0]?.path[0] ?? "adjust");
      invalidGeometry(
        `Geometry adjust ${key} must be an integer DrawingML parameter`,
      );
    }
  }
  return { type: "preset", name, ...(adjust !== undefined && { adjust }) };
}

const blockArcOptionsSchema = z
  .strictObject({
    startAngle: z.number().finite(),
    endAngle: z.number().finite(),
    thickness: z.number().finite(),
  })
  .readonly();

/**
 * Native blockArc (annular arc segment). Angles are in degrees, measured clockwise from 3 o'clock, consistent with DrawingML;
 * thickness represents the ring width as a fraction of the shape's shorter side.
 */
export function blockArc(
  options: z.infer<typeof blockArcOptionsSchema>,
): PresetGeometry {
  const parsed = blockArcOptionsSchema.safeParse(options);
  if (!parsed.success) {
    const key = String(parsed.error.issues[0]?.path[0] ?? "options");
    invalidGeometry(`blockArc.${key} must be a finite number`);
  }
  const angle = (degrees: number) =>
    Math.round((((degrees % 360) + 360) % 360) * 60000);
  return preset("blockArc", {
    adj1: angle(parsed.data.startAngle),
    adj2: angle(parsed.data.endAngle),
    // Clamped to the OOXML [0, 50000] adjust range during serialization without emitting warnings.
    adj3: Math.round(parsed.data.thickness * 100000),
  });
}

export const pathOptionsSchema = z
  .strictObject({
    /** Dimensions of the local coordinate space, similar to an SVG viewBox starting from 0,0. */
    viewBox: z
      .tuple([z.number().finite().positive(), z.number().finite().positive()])
      .readonly(),
    fill: z.boolean().optional(),
    stroke: z.boolean().optional(),
    /** Fill rule for SVG path data; see `path`. */
    fillRule: z.enum(["nonzero", "evenodd"]).optional(),
  })
  .readonly();
export type PathOptions = z.infer<typeof pathOptionsSchema>;

const pathDataSchema = z.union([
  z.string().refine((item) => item.trim() !== ""),
  z
    .array(z.string().refine((item) => item.trim() !== ""))
    .min(1)
    .readonly(),
]);

/**
 * Creates custom geometry from SVG path data. Multiple subpaths in a single string constitute the same native path; pass an array for separately stroked paths.
 *
 * PPTX does not have fill rules, so different applications interpret them individually. The default `fillRule: "nonzero"` writes paths as-is;
 * in all applications, inner subpaths whose direction opposes the outer contour form holes. When set to `"evenodd"`, subpaths alternate directions according to nesting
 * depth, allowing applications to display even-odd fill results. Subpaths that overlap without nesting, or self-intersecting subpaths, cannot achieve consistency this way.
 */
export function path(
  data: string | readonly string[],
  options: PathOptions,
): CustomGeometry {
  const parsedOptions = pathOptionsSchema.safeParse(options);
  if (!parsedOptions.success) {
    const issue = parsedOptions.error.issues[0];
    const key = issue?.path[0];
    if (key === "fill") invalidGeometry("path.fill must be a boolean");
    if (key === "stroke") invalidGeometry("path.stroke must be a boolean");
    if (key === "fillRule") {
      invalidGeometry("path.fillRule must be nonzero or evenodd");
    }
    invalidGeometry("path.viewBox must consist of two finite positive numbers");
  }
  const parsedData = pathDataSchema.safeParse(data);
  if (!parsedData.success) {
    if (typeof data !== "string" && !Array.isArray(data)) {
      invalidGeometry("path data must be an SVG string or an array of strings");
    }
    invalidGeometry("path must contain at least one non-empty SVG path");
  }
  const validOptions = parsedOptions.data;
  const list =
    typeof parsedData.data === "string" ? [parsedData.data] : parsedData.data;
  return {
    type: "custom",
    width: validOptions.viewBox[0],
    height: validOptions.viewBox[1],
    paths: list.map((item) => {
      const commands = parseSvgPath(item);
      return {
        commands:
          validOptions.fillRule === "evenodd"
            ? alternateByNesting(commands)
            : commands,
        ...(validOptions.fill !== undefined && { fill: validOptions.fill }),
        ...(validOptions.stroke !== undefined && {
          stroke: validOptions.stroke,
        }),
      };
    }),
  };
}

type Point = { readonly x: number; readonly y: number };
type Segment = Extract<PathCommand, { op: "L" | "C" | "Q" }>;

/** Splits subpaths by Z commands; when drawing continues after Z, it starts from the previous subpath start point. */
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

/** Represents a subpath as a polygon and samples curves. */
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

/** Draws the same subpath in reverse order. */
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

/** Outer contours are drawn in one direction, holes are reversed, and islands inside holes restore outer contour direction. */
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
    invalidGeometry("SVG path must be a non-empty string");
  }
  const tokens: string[] = [];
  let cursor = 0;
  for (const match of data.matchAll(tokenPattern)) {
    const index = match.index;
    if (!/^[\s,]*$/u.test(data.slice(cursor, index))) {
      invalidGeometry(
        `SVG path contains unsupported character: ${data.slice(cursor, index)}`,
      );
    }
    tokens.push(match[0]);
    cursor = index + match[0].length;
  }
  if (!/^[\s,]*$/u.test(data.slice(cursor)) || tokens.length === 0) {
    invalidGeometry(
      `SVG path contains unsupported or missing parameter: ${data.slice(cursor)}`,
    );
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
    if (!Number.isFinite(value))
      invalidGeometry("SVG path coordinate must be a finite number");
    return value;
  };
  const flag = (name: string) => {
    const token = tokens[index];
    if (token !== "0" && token !== "1") {
      invalidGeometry(`SVG arc ${name} must be a single 0 or 1`);
    }
    index += 1;
    return token === "1";
  };
  while (index < tokens.length) {
    if (isCommand(tokens[index])) {
      command = tokens[index] ?? "";
      index += 1;
    } else if (command === "") {
      invalidGeometry(
        "SVG close command Z cannot be followed by implicit parameters",
      );
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
        invalidGeometry(`SVG path command ${command} is not supported`);
    }
  }
  if (commands[0]?.op !== "M")
    invalidGeometry("SVG path must start with move command M");
  return commands;
}

/** Converts elliptical arcs to cubic Bezier curves per SVG implementation note F.6. */
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function finiteGeometryNumber(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new CapabilityError(
      "invalid-visual-style",
      `${label} must be a finite number`,
    );
  }
  return value;
}

function pointUnit(value: number, label: string): string {
  const coordinate = Math.round(finiteGeometryNumber(value, label) * 1000);
  if (
    !Number.isSafeInteger(coordinate) ||
    Math.abs(coordinate) > 2_147_483_647
  ) {
    invalidGeometry(
      `${label} exceeds DrawingML path coordinate range after conversion`,
    );
  }
  return String(coordinate);
}

function validatePathCommand(command: PathCommand): void {
  switch (command.op) {
    case "M":
    case "L":
      pointUnit(command.x, "path.x");
      pointUnit(command.y, "path.y");
      return;
    case "C":
      pointUnit(command.x1, "path.x1");
      pointUnit(command.y1, "path.y1");
      pointUnit(command.x2, "path.x2");
      pointUnit(command.y2, "path.y2");
      pointUnit(command.x, "path.x");
      pointUnit(command.y, "path.y");
      return;
    case "Q":
      pointUnit(command.x1, "path.x1");
      pointUnit(command.y1, "path.y1");
      pointUnit(command.x, "path.x");
      pointUnit(command.y, "path.y");
      return;
    case "Z":
      return;
    default:
      invalidGeometry("Custom path contains unsupported command");
  }
}

/** Validates and resolves author Geometry (+ optional borderRadius) into a closed ResolvedGeometry model. */
export function resolveGeometry(options: {
  readonly geometry: Geometry | undefined;
  readonly borderRadius?: number | undefined;
  readonly box: { readonly width: number; readonly height: number };
}): {
  readonly geometry: ResolvedGeometry;
  readonly name: PresetGeometryName | "custom";
  readonly warnings: readonly GeometryWarning[];
} {
  const { geometry: value, borderRadius: radius, box } = options;
  const warnings: GeometryWarning[] = [];

  if (isRecord(value) && value["type"] === "custom") {
    if (radius !== undefined && radius > 0) {
      throw new CapabilityError(
        "unsupported-style-combination",
        "borderRadius cannot be used together with custom geometry",
      );
    }
    const parsedCustom = customGeometrySchema.safeParse(value);
    if (!parsedCustom.success) {
      const issue = parsedCustom.error.issues[0];
      if (issue?.path[0] === "width" || issue?.path[0] === "height") {
        if (
          typeof value[issue.path[0]] !== "number" ||
          !Number.isFinite(value[issue.path[0]])
        ) {
          finiteGeometryNumber(
            value[issue.path[0]],
            `geometry.${issue.path[0]}`,
          );
        }
        invalidGeometry(
          "Custom geometry must have a positive viewBox and at least one path",
        );
      }
      invalidGeometry(
        issue?.message ??
          "Custom geometry must have a positive viewBox and at least one path",
      );
    }
    const custom = parsedCustom.data;
    const widthUnits = pointUnit(custom.width, "geometry.width");
    const heightUnits = pointUnit(custom.height, "geometry.height");
    if (Number(widthUnits) === 0 || Number(heightUnits) === 0) {
      invalidGeometry(
        "Custom geometry viewBox must retain positive width and height in path coordinate space after conversion",
      );
    }
    const resolvedPaths = custom.paths.map((item) => {
      for (const command of item.commands) {
        validatePathCommand(command);
      }
      return {
        commands: item.commands,
        fill: item.fill ?? true,
        stroke: item.stroke ?? true,
      };
    });
    return {
      geometry: {
        kind: "custom",
        width: custom.width,
        height: custom.height,
        paths: resolvedPaths,
      },
      name: "custom",
      warnings,
    };
  }

  let rawName: string;
  const adjust: Record<string, number> = {};
  if (value === undefined) {
    rawName = radius !== undefined && radius > 0 ? "roundRect" : "rect";
  } else if (typeof value === "string") {
    rawName = value;
  } else if (
    isRecord(value) &&
    value["type"] === "preset" &&
    typeof value["name"] === "string"
  ) {
    rawName = value["name"];
    if (value["adjust"] !== undefined) {
      if (!isRecord(value["adjust"])) {
        invalidGeometry("preset.adjust must be an object");
      }
      const parsedAdjust = geometryAdjustSchema.safeParse(value["adjust"]);
      if (!parsedAdjust.success) {
        const key = String(parsedAdjust.error.issues[0]?.path[0] ?? "adjust");
        invalidGeometry(`preset.adjust.${key} must be a finite integer`);
      }
      Object.assign(adjust, parsedAdjust.data);
    }
  } else {
    invalidGeometry(
      "geometry must be a preset name, preset(...), or custom path",
    );
  }

  const parsedPresetName = presetGeometryNameSchema.safeParse(rawName);
  if (!parsedPresetName.success) {
    throw new CapabilityError(
      "unsupported-geometry",
      `Unsupported preset geometry ${rawName}`,
    );
  }
  let name = parsedPresetName.data;

  if (radius !== undefined && radius > 0) {
    if (name !== "rect" && name !== "roundRect") {
      throw new CapabilityError(
        "unsupported-style-combination",
        "borderRadius is only applicable to rect or roundRect geometry",
      );
    }
    name = "roundRect";
    if (adjust["adj"] === undefined) {
      const raw = Math.round(
        (radius / Math.min(box.width, box.height)) * 100000,
      );
      adjust["adj"] = Math.max(0, Math.min(50000, raw));
    }
  }

  return {
    geometry: {
      kind: "preset",
      name,
      adjust,
    },
    name,
    warnings,
  };
}

function pathCommandXml(command: PathCommand): string {
  switch (command.op) {
    case "M":
      return `<a:moveTo><a:pt x="${pointUnit(command.x, "path.x")}" y="${pointUnit(command.y, "path.y")}"/></a:moveTo>`;
    case "L":
      return `<a:lnTo><a:pt x="${pointUnit(command.x, "path.x")}" y="${pointUnit(command.y, "path.y")}"/></a:lnTo>`;
    case "C":
      return `<a:cubicBezTo><a:pt x="${pointUnit(command.x1, "path.x1")}" y="${pointUnit(command.y1, "path.y1")}"/><a:pt x="${pointUnit(command.x2, "path.x2")}" y="${pointUnit(command.y2, "path.y2")}"/><a:pt x="${pointUnit(command.x, "path.x")}" y="${pointUnit(command.y, "path.y")}"/></a:cubicBezTo>`;
    case "Q":
      return `<a:quadBezTo><a:pt x="${pointUnit(command.x1, "path.x1")}" y="${pointUnit(command.y1, "path.y1")}"/><a:pt x="${pointUnit(command.x, "path.x")}" y="${pointUnit(command.y, "path.y")}"/></a:quadBezTo>`;
    case "Z":
      return "<a:close/>";
    default: {
      const exhaustive: never = command;
      return exhaustive;
    }
  }
}

function clampPresetAdjustValue(
  presetName: string,
  key: string,
  amount: number,
): number {
  if (presetName === "roundRect" && key === "adj") {
    return Math.max(0, Math.min(50_000, amount));
  }
  if (presetName === "blockArc") {
    if (key === "adj1" || key === "adj2") {
      return Math.max(0, Math.min(21_600_000, amount));
    }
    if (key === "adj3") {
      return Math.max(0, Math.min(50_000, amount));
    }
  }
  return amount;
}

/** Exhaustively serializes a ResolvedGeometry to DrawingML <a:prstGeom> or <a:custGeom>. */
export function resolvedGeometryXml(geometry: ResolvedGeometry): string {
  switch (geometry.kind) {
    case "preset": {
      const values = Object.entries(geometry.adjust)
        .map(
          ([key, amount]) =>
            `<a:gd name="${key}" fmla="val ${clampPresetAdjustValue(geometry.name, key, amount)}"/>`,
        )
        .join("");
      return `<a:prstGeom prst="${geometry.name}"><a:avLst>${values}</a:avLst></a:prstGeom>`;
    }
    case "custom": {
      const widthUnits = pointUnit(geometry.width, "geometry.width");
      const heightUnits = pointUnit(geometry.height, "geometry.height");
      const pathList = geometry.paths
        .map((pathItem) => {
          const commands = pathItem.commands.map(pathCommandXml).join("");
          return `<a:path w="${widthUnits}" h="${heightUnits}"${pathItem.fill ? "" : ' fill="none"'}${pathItem.stroke ? "" : ' stroke="0"'}>${commands}</a:path>`;
        })
        .join("");
      return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${pathList}</a:pathLst></a:custGeom>`;
    }
    default: {
      const exhaustive: never = geometry;
      return exhaustive;
    }
  }
}

const drawingAngleToRad = (ang: number): number =>
  (ang / 60_000) * (Math.PI / 180);
const radToDrawingAngle = (rad: number): number =>
  rad * (180 / Math.PI) * 60_000;

function formatSvgCoord(value: number): number {
  if (Object.is(value, -0) || value === 0) return 0;
  return Math.round(value * 1_000_000) / 1_000_000;
}

function commandsToSvgD(commands: readonly PathCommand[]): string {
  return commands
    .map((command) => {
      switch (command.op) {
        case "Z":
          return "Z";
        case "M":
        case "L":
          return `${command.op}${formatSvgCoord(command.x)} ${formatSvgCoord(command.y)}`;
        case "Q":
          return `Q${formatSvgCoord(command.x1)} ${formatSvgCoord(command.y1)} ${formatSvgCoord(command.x)} ${formatSvgCoord(command.y)}`;
        case "C":
          return `C${formatSvgCoord(command.x1)} ${formatSvgCoord(command.y1)} ${formatSvgCoord(command.x2)} ${formatSvgCoord(command.y2)} ${formatSvgCoord(command.x)} ${formatSvgCoord(command.y)}`;
        default: {
          const exhaustive: never = command;
          return exhaustive;
        }
      }
    })
    .join(" ");
}

/**
 * Unwraps the DrawingML polar angle onto the continuous eccentric parameter angle of an ellipse
 * with semi-axes (rx, ry), preserving exact winding across arbitrary sweep angles.
 */
function eccentricAngle(theta: number, rx: number, ry: number): number {
  if (rx === ry) return theta;
  const base = Math.atan2(rx * Math.sin(theta), ry * Math.cos(theta));
  return base + 2 * Math.PI * Math.round((theta - base) / (2 * Math.PI));
}

function drawingMlArcToLocalCubics(
  curX: number,
  curY: number,
  wR: number,
  hR: number,
  stAng: number,
  swAng: number,
): {
  readonly commands: readonly PathCommand[];
  readonly endX: number;
  readonly endY: number;
} {
  if (swAng === 0) {
    return { commands: [], endX: curX, endY: curY };
  }
  const rx = Math.max(0, wR);
  const ry = Math.max(0, hR);
  if (rx === 0 && ry === 0) {
    return { commands: [], endX: curX, endY: curY };
  }

  const theta1 = drawingAngleToRad(stAng);
  const deltaTheta = drawingAngleToRad(swAng);
  const theta2 = theta1 + deltaTheta;

  if (rx === 0 || ry === 0) {
    // Degenerate 1D ellipse: project along the non-zero axis via eccentric parameterization.
    const phi1 =
      rx === 0
        ? Math.sin(theta1) >= 0
          ? Math.PI / 2
          : -Math.PI / 2
        : Math.cos(theta1) >= 0
          ? 0
          : Math.PI;
    const phi2 =
      rx === 0
        ? Math.sin(theta2) >= 0
          ? Math.PI / 2
          : -Math.PI / 2
        : Math.cos(theta2) >= 0
          ? 0
          : Math.PI;
    const cx = curX - rx * Math.cos(phi1);
    const cy = curY - ry * Math.sin(phi1);
    const endX = cx + rx * Math.cos(phi2);
    const endY = cy + ry * Math.sin(phi2);
    return {
      commands: [{ op: "L", x: endX, y: endY }],
      endX,
      endY,
    };
  }

  const phi1 = eccentricAngle(theta1, rx, ry);
  const phi2 = eccentricAngle(theta2, rx, ry);
  const deltaPhi = phi2 - phi1;
  const cx = curX - rx * Math.cos(phi1);
  const cy = curY - ry * Math.sin(phi1);

  const segments = Math.max(1, Math.ceil(Math.abs(deltaPhi) / (Math.PI / 2)));
  const step = deltaPhi / segments;
  const kappa = (4 / 3) * Math.tan(step / 4);
  const result: PathCommand[] = [];

  for (let segment = 0; segment < segments; segment += 1) {
    const a = phi1 + segment * step;
    const b = a + step;
    const paX = cx + rx * Math.cos(a);
    const paY = cy + ry * Math.sin(a);
    const daX = -rx * Math.sin(a);
    const daY = ry * Math.cos(a);
    const pbX = cx + rx * Math.cos(b);
    const pbY = cy + ry * Math.sin(b);
    const dbX = -rx * Math.sin(b);
    const dbY = ry * Math.cos(b);
    result.push({
      op: "C",
      x1: paX + kappa * daX,
      y1: paY + kappa * daY,
      x2: pbX - kappa * dbX,
      y2: pbY - kappa * dbY,
      x: pbX,
      y: pbY,
    });
  }

  const endX = cx + rx * Math.cos(phi2);
  const endY = cy + ry * Math.sin(phi2);
  return { commands: result, endX, endY };
}

function evaluatePresetGeometry(
  name: PresetGeometryName,
  adjust: Readonly<Record<string, number>>,
  box: Rect,
): EvaluatedGeometry {
  const def: PresetShapeDefinition = presetDefinitions[name];
  // Evaluate guides in EMU space to preserve ECMA-376 precision on small pixel boxes.
  const W = Math.max(1, pxToEmu(box.width));
  const H = Math.max(1, pxToEmu(box.height));
  const ss = Math.min(W, H);
  const ls = Math.max(W, H);

  const builtins: Record<string, number> = {
    w: W,
    h: H,
    l: 0,
    t: 0,
    r: W,
    b: H,
    hc: W / 2,
    vc: H / 2,
    ss,
    ls,
    cd2: 10_800_000,
    cd4: 5_400_000,
    cd8: 2_700_000,
    "3cd4": 16_200_000,
    "3cd8": 8_100_000,
    "5cd8": 13_500_000,
    "7cd8": 18_900_000,
  };

  const env = new Map<string, number>();

  const arg = (token: string): number => {
    const bound = env.get(token);
    if (bound !== undefined) return bound;
    const builtin = builtins[token];
    if (builtin !== undefined) return builtin;
    const denomMatch = /^(wd|hd|ssd)(\d+)$/u.exec(token);
    if (denomMatch !== null) {
      const kind = denomMatch[1];
      const divisor = Number(denomMatch[2]);
      return kind === "wd"
        ? W / divisor
        : kind === "hd"
          ? H / divisor
          : ss / divisor;
    }
    const numeric = Number(token);
    if (!Number.isFinite(numeric)) {
      invalidGeometry(`Preset ${name} references unknown guide token ${token}`);
    }
    return numeric;
  };

  const evalFormula = (formula: string): number => {
    const parts = formula.trim().split(/\s+/u);
    const op = parts[0];
    const x = parts[1] === undefined ? 0 : arg(parts[1]);
    const y = parts[2] === undefined ? 0 : arg(parts[2]);
    const z = parts[3] === undefined ? 0 : arg(parts[3]);
    switch (op) {
      case "val":
        return x;
      case "*/":
        return z === 0 ? 0 : (x * y) / z;
      case "+-":
        return x + y - z;
      case "+/":
        return z === 0 ? 0 : (x + y) / z;
      case "?:":
      case "ifelse":
        return x > 0 ? y : z;
      case "abs":
        return Math.abs(x);
      case "at2":
        return radToDrawingAngle(Math.atan2(y, x));
      case "cat2":
        return x * Math.cos(Math.atan2(z, y));
      case "sat2":
        return x * Math.sin(Math.atan2(z, y));
      case "cos":
        return x * Math.cos(drawingAngleToRad(y));
      case "sin":
        return x * Math.sin(drawingAngleToRad(y));
      case "tan":
        return x * Math.tan(drawingAngleToRad(y));
      case "max":
        return Math.max(x, y);
      case "min":
        return Math.min(x, y);
      case "mod":
        return Math.sqrt(x * x + y * y + z * z);
      case "pin":
        return y < x ? x : y > z ? z : y;
      case "sqrt":
        return x <= 0 ? 0 : Math.sqrt(x);
      default:
        invalidGeometry(`Preset ${name} uses unsupported guide operator ${op}`);
    }
  };

  if (def.av !== undefined) {
    for (const [key, formula] of def.av) {
      env.set(key, evalFormula(formula));
    }
  }
  for (const [key, value] of Object.entries(adjust)) {
    env.set(key, value);
  }
  if (def.gd !== undefined) {
    for (const [key, formula] of def.gd) {
      env.set(key, evalFormula(formula));
    }
  }

  const rectTokens = def.rect ?? (["l", "t", "r", "b"] as const);
  const rectL = arg(rectTokens[0]);
  const rectT = arg(rectTokens[1]);
  const rectR = arg(rectTokens[2]);
  const rectB = arg(rectTokens[3]);
  const textRect: Rect = {
    x: box.x + (rectL / W) * box.width,
    y: box.y + (rectT / H) * box.height,
    width: Math.max(0, ((rectR - rectL) / W) * box.width),
    height: Math.max(0, ((rectB - rectT) / H) * box.height),
  };

  const paths: EvaluatedGeometryPath[] = def.paths.map((pathDef) => {
    const pathW = pathDef.w ?? W;
    const pathH = pathDef.h ?? H;
    const scaleX = box.width / pathW;
    const scaleY = box.height / pathH;
    const toPageX = (lx: number): number => box.x + lx * scaleX;
    const toPageY = (ly: number): number => box.y + ly * scaleY;
    const mapCmdToPage = (cmd: PathCommand): PathCommand => {
      switch (cmd.op) {
        case "Z":
          return cmd;
        case "M":
        case "L":
          return { op: cmd.op, x: toPageX(cmd.x), y: toPageY(cmd.y) };
        case "Q":
          return {
            op: "Q",
            x1: toPageX(cmd.x1),
            y1: toPageY(cmd.y1),
            x: toPageX(cmd.x),
            y: toPageY(cmd.y),
          };
        case "C":
          return {
            op: "C",
            x1: toPageX(cmd.x1),
            y1: toPageY(cmd.y1),
            x2: toPageX(cmd.x2),
            y2: toPageY(cmd.y2),
            x: toPageX(cmd.x),
            y: toPageY(cmd.y),
          };
      }
    };

    const commands: PathCommand[] = [];
    let curX = 0;
    let curY = 0;
    let subStartX = 0;
    let subStartY = 0;

    for (const cmd of pathDef.cmds) {
      switch (cmd[0]) {
        case "M": {
          curX = arg(cmd[1]);
          curY = arg(cmd[2]);
          subStartX = curX;
          subStartY = curY;
          commands.push({ op: "M", x: toPageX(curX), y: toPageY(curY) });
          break;
        }
        case "L": {
          curX = arg(cmd[1]);
          curY = arg(cmd[2]);
          commands.push({ op: "L", x: toPageX(curX), y: toPageY(curY) });
          break;
        }
        case "Q": {
          const x1 = arg(cmd[1]);
          const y1 = arg(cmd[2]);
          curX = arg(cmd[3]);
          curY = arg(cmd[4]);
          commands.push({
            op: "Q",
            x1: toPageX(x1),
            y1: toPageY(y1),
            x: toPageX(curX),
            y: toPageY(curY),
          });
          break;
        }
        case "C": {
          const x1 = arg(cmd[1]);
          const y1 = arg(cmd[2]);
          const x2 = arg(cmd[3]);
          const y2 = arg(cmd[4]);
          curX = arg(cmd[5]);
          curY = arg(cmd[6]);
          commands.push({
            op: "C",
            x1: toPageX(x1),
            y1: toPageY(y1),
            x2: toPageX(x2),
            y2: toPageY(y2),
            x: toPageX(curX),
            y: toPageY(curY),
          });
          break;
        }
        case "A": {
          const wR = arg(cmd[1]);
          const hR = arg(cmd[2]);
          const stAng = arg(cmd[3]);
          const swAng = arg(cmd[4]);
          const arc = drawingMlArcToLocalCubics(
            curX,
            curY,
            wR,
            hR,
            stAng,
            swAng,
          );
          for (const cubic of arc.commands) {
            commands.push(mapCmdToPage(cubic));
          }
          curX = arc.endX;
          curY = arc.endY;
          break;
        }
        case "Z": {
          curX = subStartX;
          curY = subStartY;
          commands.push({ op: "Z" });
          break;
        }
      }
    }

    return {
      d: commandsToSvgD(commands),
      commands,
      fillMode: pathDef.fill ?? "norm",
      stroke: pathDef.stroke ?? true,
    };
  });

  return { paths, textRect };
}

function evaluateCustomGeometry(
  geometry: Extract<ResolvedGeometry, { readonly kind: "custom" }>,
  box: Rect,
): EvaluatedGeometry {
  const mapX = (value: number) => box.x + (value * box.width) / geometry.width;
  const mapY = (value: number) =>
    box.y + (value * box.height) / geometry.height;

  const paths: EvaluatedGeometryPath[] = geometry.paths.map((pathItem) => {
    const commands = pathItem.commands.map((command): PathCommand => {
      switch (command.op) {
        case "Z":
          return command;
        case "M":
        case "L":
          return {
            op: command.op,
            x: mapX(command.x),
            y: mapY(command.y),
          };
        case "Q":
          return {
            op: "Q",
            x1: mapX(command.x1),
            y1: mapY(command.y1),
            x: mapX(command.x),
            y: mapY(command.y),
          };
        case "C":
          return {
            op: "C",
            x1: mapX(command.x1),
            y1: mapY(command.y1),
            x2: mapX(command.x2),
            y2: mapY(command.y2),
            x: mapX(command.x),
            y: mapY(command.y),
          };
        default: {
          const exhaustive: never = command;
          return exhaustive;
        }
      }
    });
    return {
      d: commandsToSvgD(commands),
      commands,
      fillMode: pathItem.fill ? "norm" : "none",
      stroke: pathItem.stroke,
    };
  });

  return { paths, textRect: box };
}

/** Exhaustively evaluates any ResolvedGeometry into page-space SVG paths and embedded textRect. */
export function evaluateGeometry(
  geometry: ResolvedGeometry,
  box: Rect,
): EvaluatedGeometry {
  switch (geometry.kind) {
    case "preset":
      return evaluatePresetGeometry(geometry.name, geometry.adjust, box);
    case "custom":
      return evaluateCustomGeometry(geometry, box);
    default: {
      const exhaustive: never = geometry;
      return exhaustive;
    }
  }
}

/** Modulates RGB hex color lightness using DrawingML HSL lumMod/lumOff rules. */
export function shadeHex(hex: string, lumMod: number, lumOff: number): string {
  const [r, g, b] = [0, 2, 4].map(
    (start) => Number.parseInt(hex.slice(start, start + 2), 16) / 255,
  ) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation =
    delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  const hue =
    delta === 0
      ? 0
      : max === r
        ? ((g - b) / delta + 6) % 6
        : max === g
          ? (b - r) / delta + 2
          : (r - g) / delta + 4;
  const shaded = Math.min(1, Math.max(0, lightness * lumMod + lumOff));
  const chroma = (1 - Math.abs(2 * shaded - 1)) * saturation;
  const x = chroma * (1 - Math.abs((hue % 2) - 1));
  const m = shaded - chroma / 2;
  const [r1, g1, b1] =
    hue < 1
      ? [chroma, x, 0]
      : hue < 2
        ? [x, chroma, 0]
        : hue < 3
          ? [0, chroma, x]
          : hue < 4
            ? [0, x, chroma]
            : hue < 5
              ? [x, 0, chroma]
              : [chroma, 0, x];
  return [r1, g1, b1]
    .map((channel) =>
      Math.round((channel + m) * 255)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

const PresetGeomAttrSchema = z.object({
  prst: z.string().min(1),
});

const AdjustGuideAttrSchema = z.object({
  name: geometryAdjustKeySchema,
  fmla: z.string().regex(/^val\s+(-?\d+)$/u),
});

const CustomPathAttrSchema = z.object({
  w: ooxmlPositiveIntSchema.optional(),
  h: ooxmlPositiveIntSchema.optional(),
  fill: z
    .enum(["none", "norm", "lighten", "lightenLess", "darken", "darkenLess"])
    .optional(),
  stroke: ooxmlBoolSchema.optional(),
});

const CustomPtAttrSchema = z.object({
  x: ooxmlIntSchema,
  y: ooxmlIntSchema,
});

const CustomArcAttrSchema = z.object({
  wR: ooxmlIntSchema,
  hR: ooxmlIntSchema,
  stAng: ooxmlIntSchema,
  swAng: ooxmlIntSchema,
});

/** Reads DrawingML `<a:prstGeom>` (and `<a:avLst>`) into author `Geometry`, emitting localized diagnostics on invalid/unsupported values. */
export function readPresetGeometryXml(
  prstGeomCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): Geometry {
  const raw = prstGeomCursor.attrsObject(["prst"]);
  const parsedAttr = PresetGeomAttrSchema.safeParse(raw);
  if (!parsedAttr.success) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "recovered",
      location: { ...location, xpath: prstGeomCursor.xpath },
      attributeOrElement: "a:prstGeom/@prst",
      rawValue: raw["prst"] ?? "",
    });
    prstGeomCursor.ignoreEmptyChild(A_NS, "avLst");
    return "rect";
  }

  const parsedName = presetGeometryNameSchema.safeParse(parsedAttr.data.prst);
  if (!parsedName.success) {
    diagnostics.push({
      code: "unsupported-mechanism",
      category: "unsupported",
      disposition: "omitted",
      location: { ...location, xpath: prstGeomCursor.xpath },
      namespaceUri: A_NS,
      localName: "prstGeom",
      mechanismHint: `preset:${parsedAttr.data.prst}`,
    });
    prstGeomCursor.ignoreEmptyChild(A_NS, "avLst");
    return "rect";
  }

  const adjust: Record<string, number> = {};
  const avLst = prstGeomCursor.child(A_NS, "avLst");
  if (avLst !== undefined) {
    for (const gd of avLst.children(A_NS, "gd")) {
      const rawGd = gd.attrsObject(["name", "fmla"]);
      const parsedGd = AdjustGuideAttrSchema.safeParse(rawGd);
      if (!parsedGd.success) {
        diagnostics.push({
          code: "invalid-xml-value",
          category: "corrupt-input",
          disposition: "omitted",
          location: { ...location, xpath: gd.xpath },
          attributeOrElement: "a:gd",
          rawValue: JSON.stringify(rawGd),
        });
        continue;
      }
      const match = /^val\s+(-?\d+)$/u.exec(parsedGd.data.fmla);
      if (match !== null) {
        const value = Number(match[1]);
        if (Number.isSafeInteger(value)) {
          adjust[parsedGd.data.name] = value;
        }
      }
    }
  }

  return Object.keys(adjust).length === 0
    ? parsedName.data
    : { type: "preset", name: parsedName.data, adjust };
}

function readCustomPoint(
  ptCursor: XmlElementCursor,
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): { readonly x: number; readonly y: number } | undefined {
  const raw = ptCursor.attrsObject(["x", "y"]);
  const parsed = CustomPtAttrSchema.safeParse(raw);
  if (!parsed.success) {
    const isFormulaRef =
      (raw["x"] !== undefined && !/^-?\d+$/u.test(raw["x"])) ||
      (raw["y"] !== undefined && !/^-?\d+$/u.test(raw["y"]));
    diagnostics.push(
      isFormulaRef
        ? {
            code: "unsupported-mechanism",
            category: "unsupported",
            disposition: "omitted",
            location: { ...location, xpath: ptCursor.xpath },
            namespaceUri: A_NS,
            localName: "pt",
            mechanismHint: "custom-geometry-guide-reference",
          }
        : {
            code: "invalid-xml-value",
            category: "corrupt-input",
            disposition: "omitted",
            location: { ...location, xpath: ptCursor.xpath },
            attributeOrElement: "a:pt",
            rawValue: JSON.stringify(raw),
          },
    );
    return undefined;
  }
  return { x: parsed.data.x / 1000, y: parsed.data.y / 1000 };
}

/** Reads DrawingML `<a:custGeom>` into `CustomGeometry`, reporting localized diagnostics on unsupported guide formulas or invalid paths. */
export function readCustomGeometryXml(
  custGeomCursor: XmlElementCursor,
  fallbackExtentUnits: { readonly w: number; readonly h: number },
  location: ImportLocation,
  diagnostics: ImportDiagnostic[],
): CustomGeometry | undefined {
  custGeomCursor.ignoreEmptyChild(A_NS, "avLst");
  const gdLst = custGeomCursor.child(A_NS, "gdLst");
  if (gdLst !== undefined) {
    for (const gd of gdLst.children(A_NS, "gd")) {
      gd.consumeAttr("name");
      gd.consumeAttr("fmla");
    }
  }
  custGeomCursor.ignoreEmptyChild(A_NS, "ahLst");
  custGeomCursor.ignoreEmptyChild(A_NS, "cxnLst");
  const rect = custGeomCursor.child(A_NS, "rect");
  if (rect !== undefined) {
    rect.consumeAttr("l");
    rect.consumeAttr("t");
    rect.consumeAttr("r");
    rect.consumeAttr("b");
  }

  const pathLst = custGeomCursor.child(A_NS, "pathLst");
  if (pathLst === undefined) {
    diagnostics.push({
      code: "invalid-xml-value",
      category: "corrupt-input",
      disposition: "omitted",
      location: { ...location, xpath: custGeomCursor.xpath },
      attributeOrElement: "a:custGeom/a:pathLst",
    });
    return undefined;
  }

  const paths: GeometryPath[] = [];
  let viewWidth: number | undefined;
  let viewHeight: number | undefined;

  for (const pathCursor of pathLst.children(A_NS, "path")) {
    const rawPath = pathCursor.attrsObject(["w", "h", "fill", "stroke"]);
    const parsedPath = CustomPathAttrSchema.safeParse(rawPath);
    if (!parsedPath.success) {
      diagnostics.push({
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        location: { ...location, xpath: pathCursor.xpath },
        attributeOrElement: "a:path",
        rawValue: JSON.stringify(rawPath),
      });
      continue;
    }
    const pathWUnits = parsedPath.data.w ?? fallbackExtentUnits.w;
    const pathHUnits = parsedPath.data.h ?? fallbackExtentUnits.h;
    const pathWidth = pathWUnits / 1000;
    const pathHeight = pathHUnits / 1000;
    viewWidth ??= pathWidth;
    viewHeight ??= pathHeight;
    const scaleX = pathWidth === viewWidth ? 1 : viewWidth / pathWidth;
    const scaleY = pathHeight === viewHeight ? 1 : viewHeight / pathHeight;

    const commands: PathCommand[] = [];
    let curX = 0;
    let curY = 0;
    let subStartX = 0;
    let subStartY = 0;
    let pathFailed = false;

    for (const cmdCursor of pathCursor.remainingElementChildren()) {
      if (cmdCursor.namespaceUri !== A_NS) {
        diagnostics.push({
          code: "unconsumed-xml",
          category: "unsupported",
          disposition: "omitted",
          kind: "element",
          location: { ...location, xpath: cmdCursor.xpath },
          namespaceUri: cmdCursor.namespaceUri,
          localName: cmdCursor.localName,
        });
        continue;
      }
      switch (cmdCursor.localName) {
        case "moveTo": {
          const pt = cmdCursor.child(A_NS, "pt");
          const coords =
            pt === undefined
              ? undefined
              : readCustomPoint(pt, location, diagnostics);
          if (coords === undefined) {
            pathFailed = true;
            break;
          }
          curX = coords.x * scaleX;
          curY = coords.y * scaleY;
          subStartX = curX;
          subStartY = curY;
          commands.push({ op: "M", x: curX, y: curY });
          break;
        }
        case "lnTo": {
          const pt = cmdCursor.child(A_NS, "pt");
          const coords =
            pt === undefined
              ? undefined
              : readCustomPoint(pt, location, diagnostics);
          if (coords === undefined) {
            pathFailed = true;
            break;
          }
          curX = coords.x * scaleX;
          curY = coords.y * scaleY;
          commands.push({ op: "L", x: curX, y: curY });
          break;
        }
        case "cubicBezTo": {
          const pts = cmdCursor.children(A_NS, "pt");
          if (pts.length !== 3) {
            pathFailed = true;
            break;
          }
          const p1 = readCustomPoint(pts[0]!, location, diagnostics);
          const p2 = readCustomPoint(pts[1]!, location, diagnostics);
          const p3 = readCustomPoint(pts[2]!, location, diagnostics);
          if (p1 === undefined || p2 === undefined || p3 === undefined) {
            pathFailed = true;
            break;
          }
          curX = p3.x * scaleX;
          curY = p3.y * scaleY;
          commands.push({
            op: "C",
            x1: p1.x * scaleX,
            y1: p1.y * scaleY,
            x2: p2.x * scaleX,
            y2: p2.y * scaleY,
            x: curX,
            y: curY,
          });
          break;
        }
        case "quadBezTo": {
          const pts = cmdCursor.children(A_NS, "pt");
          if (pts.length !== 2) {
            pathFailed = true;
            break;
          }
          const p1 = readCustomPoint(pts[0]!, location, diagnostics);
          const p2 = readCustomPoint(pts[1]!, location, diagnostics);
          if (p1 === undefined || p2 === undefined) {
            pathFailed = true;
            break;
          }
          curX = p2.x * scaleX;
          curY = p2.y * scaleY;
          commands.push({
            op: "Q",
            x1: p1.x * scaleX,
            y1: p1.y * scaleY,
            x: curX,
            y: curY,
          });
          break;
        }
        case "arcTo": {
          const rawArc = cmdCursor.attrsObject(["wR", "hR", "stAng", "swAng"]);
          const parsedArc = CustomArcAttrSchema.safeParse(rawArc);
          if (!parsedArc.success) {
            diagnostics.push({
              code: "unsupported-mechanism",
              category: "unsupported",
              disposition: "omitted",
              location: { ...location, xpath: cmdCursor.xpath },
              namespaceUri: A_NS,
              localName: "arcTo",
              mechanismHint: "custom-geometry-arc-guide",
            });
            pathFailed = true;
            break;
          }
          const arc = drawingMlArcToLocalCubics(
            curX / scaleX,
            curY / scaleY,
            parsedArc.data.wR / 1000,
            parsedArc.data.hR / 1000,
            parsedArc.data.stAng,
            parsedArc.data.swAng,
          );
          for (const cubic of arc.commands) {
            if (cubic.op === "C") {
              commands.push({
                op: "C",
                x1: cubic.x1 * scaleX,
                y1: cubic.y1 * scaleY,
                x2: cubic.x2 * scaleX,
                y2: cubic.y2 * scaleY,
                x: cubic.x * scaleX,
                y: cubic.y * scaleY,
              });
            } else if (cubic.op === "L") {
              commands.push({
                op: "L",
                x: cubic.x * scaleX,
                y: cubic.y * scaleY,
              });
            }
          }
          curX = arc.endX * scaleX;
          curY = arc.endY * scaleY;
          break;
        }
        case "close":
          curX = subStartX;
          curY = subStartY;
          commands.push({ op: "Z" });
          break;
        default:
          diagnostics.push({
            code: "unsupported-mechanism",
            category: "unsupported",
            disposition: "omitted",
            location: { ...location, xpath: cmdCursor.xpath },
            namespaceUri: cmdCursor.namespaceUri,
            localName: cmdCursor.localName,
          });
          break;
      }
    }

    if (pathFailed || commands.length === 0) continue;
    paths.push({
      commands,
      ...(parsedPath.data.fill === "none" ? { fill: false } : {}),
      ...(parsedPath.data.stroke === false ? { stroke: false } : {}),
    });
  }

  if (
    viewWidth === undefined ||
    viewHeight === undefined ||
    viewWidth <= 0 ||
    viewHeight <= 0 ||
    paths.length === 0
  ) {
    return undefined;
  }

  return {
    type: "custom",
    width: viewWidth,
    height: viewHeight,
    paths,
  };
}
