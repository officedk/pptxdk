import { bytesSha256, type PaintRasterResource } from "@pptxdk/core/capability";

const EPSILON = 1e-6;
const SAFE_OFFSCREEN_PAD = 512;
const MAX_BLUR_STD_DEVIATION = 256;
const RASTER_TOKEN_PREFIX = "urn:pptxdk-raster:";

type RasterEntry = {
  readonly id: number;
  readonly contentType: PaintRasterResource["contentType"];
  readonly bytes: Uint8Array;
  dataUri?: string;
};

/**
 * Page-scoped registry for raster payloads submitted via `PaintContext.registerRasterResource`.
 *
 * Defers Base64 encoding until after SVG sanitization so offscreen or zero-clip images are never
 * encoded, single-use images remain plain `<image>` elements, and payloads surviving two or more
 * times on the same page are encoded once into `<defs>` + `<use>`.
 */
export type PageRasterRegistry = {
  readonly idPrefix: string;
  readonly size: number;
  register(resource: PaintRasterResource): string;
  getEntry(id: number): RasterEntry | undefined;
  getDataUri(entry: RasterEntry): string;
};

export function createPageRasterRegistry(idPrefix: string): PageRasterRegistry {
  const entries: RasterEntry[] = [];
  const byBytesRef = new Map<
    Uint8Array,
    Map<PaintRasterResource["contentType"], RasterEntry>
  >();
  const byDigest = new Map<string, RasterEntry>();

  return {
    idPrefix,
    get size(): number {
      return entries.length;
    },
    register(resource: PaintRasterResource): string {
      let byType = byBytesRef.get(resource.bytes);
      const refHit = byType?.get(resource.contentType);
      if (refHit !== undefined) {
        return `${RASTER_TOKEN_PREFIX}${refHit.id}`;
      }

      const digestKey = `${resource.contentType}:${bytesSha256(resource.bytes)}`;
      let entry = byDigest.get(digestKey);
      if (entry === undefined) {
        entry = {
          id: entries.length,
          contentType: resource.contentType,
          bytes: resource.bytes,
        };
        entries.push(entry);
        byDigest.set(digestKey, entry);
      }
      if (byType === undefined) {
        byType = new Map();
        byBytesRef.set(resource.bytes, byType);
      }
      byType.set(resource.contentType, entry);
      return `${RASTER_TOKEN_PREFIX}${entry.id}`;
    },
    getEntry(id: number): RasterEntry | undefined {
      return entries[id];
    },
    getDataUri(entry: RasterEntry): string {
      entry.dataUri ??= `data:${entry.contentType};base64,${Buffer.from(entry.bytes).toString("base64")}`;
      return entry.dataUri;
    },
  };
}

type BBox = {
  readonly minX: number;
  readonly minY: number;
  readonly maxX: number;
  readonly maxY: number;
};

type Matrix2D = readonly [number, number, number, number, number, number];

const IDENTITY_MATRIX: Matrix2D = [1, 0, 0, 1, 0, 0];

type SvgNode =
  | { readonly kind: "text"; readonly text: string }
  | {
      readonly kind: "element";
      readonly tag: string;
      readonly attrs: Map<string, string>;
      readonly children: SvgNode[];
      readonly selfClosing: boolean;
    };

type SvgElement = Extract<SvgNode, { readonly kind: "element" }>;

function decodeXmlEntities(value: string): string {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");
}

function encodeXmlAttr(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function parseAttributes(raw: string): Map<string, string> {
  const attrs = new Map<string, string>();
  const attrPattern = /([A-Za-z_:][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/gu;
  for (const match of raw.matchAll(attrPattern)) {
    const key = match[1];
    if (key === undefined) continue;
    const val = match[2] ?? match[3] ?? "";
    attrs.set(key, decodeXmlEntities(val));
  }
  return attrs;
}

function parseSvgTree(svg: string): SvgNode[] {
  const root: SvgNode[] = [];
  const stack: SvgElement[] = [];
  const tagPattern = /<(\/?)([A-Za-z_:][\w:.-]*)([^>]*?)(\/?)>/gu;
  let lastIndex = 0;

  const pushNode = (node: SvgNode): void => {
    const parent = stack.at(-1);
    if (parent !== undefined) {
      parent.children.push(node);
    } else {
      root.push(node);
    }
  };

  for (const match of svg.matchAll(tagPattern)) {
    const index = match.index ?? 0;
    if (index > lastIndex) {
      const text = svg.slice(lastIndex, index);
      if (text !== "") {
        pushNode({ kind: "text", text });
      }
    }
    lastIndex = index + match[0].length;
    const isClosing = match[1] === "/";
    const tag = match[2]!;
    const rawAttrs = match[3] ?? "";
    const selfClosing = match[4] === "/";

    if (isClosing) {
      while (stack.length > 0) {
        const popped = stack.pop()!;
        if (popped.tag === tag) break;
      }
      continue;
    }

    const element: SvgElement = {
      kind: "element",
      tag,
      attrs: parseAttributes(rawAttrs),
      children: [],
      selfClosing,
    };
    pushNode(element);
    if (!selfClosing) {
      stack.push(element);
    }
  }

  if (lastIndex < svg.length) {
    const tail = svg.slice(lastIndex);
    if (tail !== "") {
      pushNode({ kind: "text", text: tail });
    }
  }
  return root;
}

function serializeSvgNode(node: SvgNode): string {
  if (node.kind === "text") return node.text;
  const attrsParts: string[] = [];
  for (const [key, value] of node.attrs.entries()) {
    attrsParts.push(`${key}="${encodeXmlAttr(value)}"`);
  }
  const attrString = attrsParts.length === 0 ? "" : ` ${attrsParts.join(" ")}`;
  if (node.selfClosing && node.children.length === 0) {
    return `<${node.tag}${attrString}/>`;
  }
  const inner = node.children.map(serializeSvgNode).join("");
  return `<${node.tag}${attrString}>${inner}</${node.tag}>`;
}

function formatNumber(value: number): string {
  const rounded = Math.round(value * 1_000_000) / 1_000_000;
  return Object.is(rounded, -0) ? "0" : String(rounded);
}

function parseNumber(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const trimmed = value.trim();
  if (trimmed === "" || trimmed.endsWith("%")) return undefined;
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function parseLengthAttr(
  value: string | undefined,
  referenceSpan: number,
  defaultValue: number,
): number {
  if (value === undefined) return defaultValue;
  const trimmed = value.trim();
  if (trimmed === "") return Number.NaN;
  if (trimmed.endsWith("%")) {
    const pct = Number(trimmed.slice(0, -1));
    return Number.isFinite(pct) ? (pct / 100) * referenceSpan : Number.NaN;
  }
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function hasNonFiniteAttr(
  element: SvgElement,
  keys: readonly string[],
): boolean {
  for (const key of keys) {
    const raw = element.attrs.get(key);
    if (raw === undefined) continue;
    const trimmed = raw.trim();
    if (trimmed === "") return true;
    const numStr = trimmed.endsWith("%") ? trimmed.slice(0, -1) : trimmed;
    if (!Number.isFinite(Number(numStr))) return true;
  }
  return false;
}

function multiplyMatrix(m1: Matrix2D, m2: Matrix2D): Matrix2D {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function parseTransformAttr(transformStr: string | undefined): Matrix2D | null {
  if (transformStr === undefined || transformStr.trim() === "") {
    return IDENTITY_MATRIX;
  }
  let current: Matrix2D = IDENTITY_MATRIX;
  const opPattern = /([A-Za-z]+)\s*\(([^)]*)\)/gu;
  let matchedAny = false;
  for (const match of transformStr.matchAll(opPattern)) {
    matchedAny = true;
    const fn = match[1];
    const rawArgs = match[2] ?? "";
    const nums = rawArgs
      .trim()
      .split(/[\s,]+/u)
      .filter((token) => token !== "")
      .map(Number);
    if (nums.some((n) => !Number.isFinite(n))) {
      return null;
    }
    let next: Matrix2D | undefined;
    switch (fn) {
      case "translate": {
        if (nums.length < 1) return null;
        next = [1, 0, 0, 1, nums[0]!, nums[1] ?? 0];
        break;
      }
      case "scale": {
        if (nums.length < 1) return null;
        const sx = nums[0]!;
        const sy = nums[1] ?? sx;
        next = [sx, 0, 0, sy, 0, 0];
        break;
      }
      case "rotate": {
        if (nums.length < 1) return null;
        const rad = (nums[0]! * Math.PI) / 180;
        const cos = Math.cos(rad);
        const sin = Math.sin(rad);
        const cx = nums[1] ?? 0;
        const cy = nums[2] ?? 0;
        if (cx === 0 && cy === 0) {
          next = [cos, sin, -sin, cos, 0, 0];
        } else {
          const toOrigin: Matrix2D = [1, 0, 0, 1, cx, cy];
          const rot: Matrix2D = [cos, sin, -sin, cos, 0, 0];
          const back: Matrix2D = [1, 0, 0, 1, -cx, -cy];
          next = multiplyMatrix(multiplyMatrix(toOrigin, rot), back);
        }
        break;
      }
      case "skewX": {
        if (nums.length < 1) return null;
        next = [1, 0, Math.tan((nums[0]! * Math.PI) / 180), 1, 0, 0];
        break;
      }
      case "skewY": {
        if (nums.length < 1) return null;
        next = [1, Math.tan((nums[0]! * Math.PI) / 180), 0, 1, 0, 0];
        break;
      }
      case "matrix": {
        if (nums.length !== 6) return null;
        next = [nums[0]!, nums[1]!, nums[2]!, nums[3]!, nums[4]!, nums[5]!];
        break;
      }
      default:
        return null;
    }
    current = multiplyMatrix(current, next);
  }
  if (!matchedAny) return null;
  if (current.some((v) => !Number.isFinite(v))) return null;
  const det = current[0] * current[3] - current[1] * current[2];
  if (!Number.isFinite(det) || Math.abs(det) <= 1e-12) {
    return null;
  }
  return current;
}

function transformBbox(bbox: BBox, m: Matrix2D): BBox {
  const corners = [
    [bbox.minX, bbox.minY],
    [bbox.maxX, bbox.minY],
    [bbox.maxX, bbox.maxY],
    [bbox.minX, bbox.maxY],
  ] as const;
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (const [x, y] of corners) {
    const tx = m[0] * x + m[2] * y + m[4];
    const ty = m[1] * x + m[3] * y + m[5];
    if (!Number.isFinite(tx) || !Number.isFinite(ty)) {
      return {
        minX: Number.NaN,
        minY: Number.NaN,
        maxX: Number.NaN,
        maxY: Number.NaN,
      };
    }
    if (tx < minX) minX = tx;
    if (ty < minY) minY = ty;
    if (tx > maxX) maxX = tx;
    if (ty > maxY) maxY = ty;
  }
  return { minX, minY, maxX, maxY };
}

function boxesIntersect(a: BBox, b: BBox): boolean {
  if (
    !Number.isFinite(a.minX) ||
    !Number.isFinite(a.minY) ||
    !Number.isFinite(a.maxX) ||
    !Number.isFinite(a.maxY)
  ) {
    return false;
  }
  return (
    a.maxX >= b.minX && a.minX <= b.maxX && a.maxY >= b.minY && a.minY <= b.maxY
  );
}

function extractUrlId(attrValue: string | undefined): string | undefined {
  if (attrValue === undefined) return undefined;
  const match = /^url\(#([^)]+)\)$/u.exec(attrValue.trim());
  return match?.[1];
}

function pathBbox(d: string): BBox | null {
  if (d.trim() === "") return null;
  const tokens = d.match(
    /[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?|NaN|Infinity|-Infinity/gu,
  );
  if (tokens === null || tokens.length === 0) return null;

  let index = 0;
  let cmd = "";
  let curX = 0;
  let curY = 0;
  let startX = 0;
  let startY = 0;
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  const recordPoint = (x: number, y: number): boolean => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return false;
    if (x < minX) minX = x;
    if (y < minY) minY = y;
    if (x > maxX) maxX = x;
    if (y > maxY) maxY = y;
    return true;
  };

  const nextNum = (): number | undefined => {
    const token = tokens[index];
    if (token === undefined || /^[A-Za-z]$/u.test(token)) return undefined;
    index += 1;
    const parsed = Number(token);
    return parsed;
  };

  while (index < tokens.length) {
    const tok = tokens[index]!;
    if (/^[A-Za-z]$/u.test(tok) && tok !== "NaN" && tok !== "Infinity") {
      cmd = tok;
      index += 1;
      if (cmd === "Z" || cmd === "z") {
        curX = startX;
        curY = startY;
      }
      continue;
    }
    if (cmd === "") return null;
    const rel = cmd === cmd.toLowerCase();
    const upper = cmd.toUpperCase();
    switch (upper) {
      case "M":
      case "L":
      case "T": {
        const x = nextNum();
        const y = nextNum();
        if (
          x === undefined ||
          y === undefined ||
          !Number.isFinite(x) ||
          !Number.isFinite(y)
        ) {
          return null;
        }
        curX = rel ? curX + x : x;
        curY = rel ? curY + y : y;
        if (!recordPoint(curX, curY)) return null;
        if (upper === "M") {
          startX = curX;
          startY = curY;
          cmd = rel ? "l" : "L";
        }
        break;
      }
      case "H": {
        const x = nextNum();
        if (x === undefined || !Number.isFinite(x)) return null;
        curX = rel ? curX + x : x;
        if (!recordPoint(curX, curY)) return null;
        break;
      }
      case "V": {
        const y = nextNum();
        if (y === undefined || !Number.isFinite(y)) return null;
        curY = rel ? curY + y : y;
        if (!recordPoint(curX, curY)) return null;
        break;
      }
      case "C": {
        const x1 = nextNum();
        const y1 = nextNum();
        const x2 = nextNum();
        const y2 = nextNum();
        const x = nextNum();
        const y = nextNum();
        if (
          x1 === undefined ||
          y1 === undefined ||
          x2 === undefined ||
          y2 === undefined ||
          x === undefined ||
          y === undefined ||
          !Number.isFinite(x1) ||
          !Number.isFinite(y1) ||
          !Number.isFinite(x2) ||
          !Number.isFinite(y2) ||
          !Number.isFinite(x) ||
          !Number.isFinite(y)
        ) {
          return null;
        }
        const ox = rel ? curX : 0;
        const oy = rel ? curY : 0;
        if (
          !recordPoint(ox + x1, oy + y1) ||
          !recordPoint(ox + x2, oy + y2) ||
          !recordPoint(ox + x, oy + y)
        ) {
          return null;
        }
        curX = ox + x;
        curY = oy + y;
        break;
      }
      case "S":
      case "Q": {
        const x1 = nextNum();
        const y1 = nextNum();
        const x = nextNum();
        const y = nextNum();
        if (
          x1 === undefined ||
          y1 === undefined ||
          x === undefined ||
          y === undefined ||
          !Number.isFinite(x1) ||
          !Number.isFinite(y1) ||
          !Number.isFinite(x) ||
          !Number.isFinite(y)
        ) {
          return null;
        }
        const ox = rel ? curX : 0;
        const oy = rel ? curY : 0;
        if (!recordPoint(ox + x1, oy + y1) || !recordPoint(ox + x, oy + y)) {
          return null;
        }
        curX = ox + x;
        curY = oy + y;
        break;
      }
      case "A": {
        const rx = nextNum();
        const ry = nextNum();
        const rot = nextNum();
        const large = nextNum();
        const sweep = nextNum();
        const x = nextNum();
        const y = nextNum();
        if (
          rx === undefined ||
          ry === undefined ||
          rot === undefined ||
          large === undefined ||
          sweep === undefined ||
          x === undefined ||
          y === undefined ||
          !Number.isFinite(rx) ||
          !Number.isFinite(ry) ||
          !Number.isFinite(rot) ||
          !Number.isFinite(large) ||
          !Number.isFinite(sweep) ||
          !Number.isFinite(x) ||
          !Number.isFinite(y)
        ) {
          return null;
        }
        const ox = rel ? curX : 0;
        const oy = rel ? curY : 0;
        if (!recordPoint(curX, curY) || !recordPoint(ox + x, oy + y)) {
          return null;
        }
        curX = ox + x;
        curY = oy + y;
        break;
      }
      default:
        return null;
    }
  }

  if (!Number.isFinite(minX)) return null;
  return { minX, minY, maxX, maxY };
}

function clampRectToSafeViewport(
  x: number,
  y: number,
  width: number,
  height: number,
  safeViewport: BBox,
): {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
} | null {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(y) ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width <= EPSILON ||
    height <= EPSILON
  ) {
    return null;
  }
  const minX = Math.max(x, safeViewport.minX);
  const minY = Math.max(y, safeViewport.minY);
  const maxX = Math.min(x + width, safeViewport.maxX);
  const maxY = Math.min(y + height, safeViewport.maxY);
  const clampedW = maxX - minX;
  const clampedH = maxY - minY;
  if (clampedW <= EPSILON || clampedH <= EPSILON) {
    return null;
  }
  return { x: minX, y: minY, width: clampedW, height: clampedH };
}

function extractPointsBbox(pointsAttr: string): BBox | null {
  const nums = pointsAttr
    .trim()
    .split(/[\s,]+/u)
    .filter((token) => token !== "")
    .map(Number);
  if (
    nums.length < 4 ||
    nums.length % 2 !== 0 ||
    nums.some((n) => !Number.isFinite(n))
  ) {
    return null;
  }
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  for (let i = 0; i < nums.length; i += 2) {
    const x = nums[i]!;
    const y = nums[i + 1]!;
    if (x < minX) minX = x;
    if (y < minY) minY = y;
    if (x > maxX) maxX = x;
    if (y > maxY) maxY = y;
  }
  return { minX, minY, maxX, maxY };
}

function lastGradientStopColor(
  gradient: SvgElement,
): { readonly color: string; readonly opacity?: string } | undefined {
  const stops = gradient.children.filter(
    (child): child is SvgElement =>
      child.kind === "element" && child.tag === "stop",
  );
  const last = stops.at(-1);
  if (last === undefined) return undefined;
  const color = last.attrs.get("stop-color");
  if (color === undefined || color.trim() === "") return undefined;
  const opacity = last.attrs.get("stop-opacity");
  return {
    color,
    ...(opacity === undefined ? {} : { opacity }),
  };
}

function sanitizeGeometryElementBbox(
  element: SvgElement,
  canvas: { readonly width: number; readonly height: number },
  safeViewport: BBox,
  requireArea: boolean,
): BBox | null {
  switch (element.tag) {
    case "rect": {
      if (
        hasNonFiniteAttr(element, ["x", "y", "width", "height", "rx", "ry"])
      ) {
        return null;
      }
      const x = parseLengthAttr(element.attrs.get("x"), canvas.width, 0);
      const y = parseLengthAttr(element.attrs.get("y"), canvas.height, 0);
      const width = parseLengthAttr(
        element.attrs.get("width"),
        canvas.width,
        0,
      );
      const height = parseLengthAttr(
        element.attrs.get("height"),
        canvas.height,
        0,
      );
      if (
        !Number.isFinite(x) ||
        !Number.isFinite(y) ||
        !Number.isFinite(width) ||
        !Number.isFinite(height) ||
        width <= EPSILON ||
        height <= EPSILON
      ) {
        return null;
      }
      if (
        x < safeViewport.minX ||
        y < safeViewport.minY ||
        x + width > safeViewport.maxX ||
        y + height > safeViewport.maxY
      ) {
        const clamped = clampRectToSafeViewport(
          x,
          y,
          width,
          height,
          safeViewport,
        );
        if (clamped === null) return null;
        if (!element.attrs.get("width")?.trim().endsWith("%")) {
          element.attrs.set("x", formatNumber(clamped.x));
          element.attrs.set("y", formatNumber(clamped.y));
          element.attrs.set("width", formatNumber(clamped.width));
          element.attrs.set("height", formatNumber(clamped.height));
        }
        return {
          minX: clamped.x,
          minY: clamped.y,
          maxX: clamped.x + clamped.width,
          maxY: clamped.y + clamped.height,
        };
      }
      return { minX: x, minY: y, maxX: x + width, maxY: y + height };
    }
    case "image": {
      if (hasNonFiniteAttr(element, ["x", "y", "width", "height"])) {
        return null;
      }
      const x = parseNumber(element.attrs.get("x")) ?? 0;
      const y = parseNumber(element.attrs.get("y")) ?? 0;
      const width = parseNumber(element.attrs.get("width")) ?? 0;
      const height = parseNumber(element.attrs.get("height")) ?? 0;
      if (
        !Number.isFinite(x) ||
        !Number.isFinite(y) ||
        !Number.isFinite(width) ||
        !Number.isFinite(height) ||
        width <= EPSILON ||
        height <= EPSILON
      ) {
        return null;
      }
      return { minX: x, minY: y, maxX: x + width, maxY: y + height };
    }
    case "path": {
      const d = element.attrs.get("d");
      if (d === undefined) return null;
      const bbox = pathBbox(d);
      if (bbox === null) return null;
      if (
        requireArea &&
        (bbox.maxX - bbox.minX <= EPSILON || bbox.maxY - bbox.minY <= EPSILON)
      ) {
        return null;
      }
      return bbox;
    }
    case "circle": {
      if (hasNonFiniteAttr(element, ["cx", "cy", "r"])) return null;
      const cx = parseNumber(element.attrs.get("cx")) ?? 0;
      const cy = parseNumber(element.attrs.get("cy")) ?? 0;
      const r = parseNumber(element.attrs.get("r")) ?? 0;
      if (
        !Number.isFinite(cx) ||
        !Number.isFinite(cy) ||
        !Number.isFinite(r) ||
        r <= EPSILON
      ) {
        return null;
      }
      return { minX: cx - r, minY: cy - r, maxX: cx + r, maxY: cy + r };
    }
    case "ellipse": {
      if (hasNonFiniteAttr(element, ["cx", "cy", "rx", "ry"])) return null;
      const cx = parseNumber(element.attrs.get("cx")) ?? 0;
      const cy = parseNumber(element.attrs.get("cy")) ?? 0;
      const rx = parseNumber(element.attrs.get("rx")) ?? 0;
      const ry = parseNumber(element.attrs.get("ry")) ?? 0;
      if (
        !Number.isFinite(cx) ||
        !Number.isFinite(cy) ||
        !Number.isFinite(rx) ||
        !Number.isFinite(ry) ||
        rx <= EPSILON ||
        ry <= EPSILON
      ) {
        return null;
      }
      return { minX: cx - rx, minY: cy - ry, maxX: cx + rx, maxY: cy + ry };
    }
    case "line": {
      if (hasNonFiniteAttr(element, ["x1", "y1", "x2", "y2"])) return null;
      const x1 = parseNumber(element.attrs.get("x1")) ?? 0;
      const y1 = parseNumber(element.attrs.get("y1")) ?? 0;
      const x2 = parseNumber(element.attrs.get("x2")) ?? 0;
      const y2 = parseNumber(element.attrs.get("y2")) ?? 0;
      if (
        !Number.isFinite(x1) ||
        !Number.isFinite(y1) ||
        !Number.isFinite(x2) ||
        !Number.isFinite(y2)
      ) {
        return null;
      }
      return {
        minX: Math.min(x1, x2),
        minY: Math.min(y1, y2),
        maxX: Math.max(x1, x2),
        maxY: Math.max(y1, y2),
      };
    }
    case "polygon":
    case "polyline": {
      const points = element.attrs.get("points");
      if (points === undefined) return null;
      const bbox = extractPointsBbox(points);
      if (bbox === null) return null;
      if (
        requireArea &&
        (bbox.maxX - bbox.minX <= EPSILON || bbox.maxY - bbox.minY <= EPSILON)
      ) {
        return null;
      }
      return bbox;
    }
    case "use": {
      if (hasNonFiniteAttr(element, ["x", "y", "width", "height"])) {
        return null;
      }
      const x = parseNumber(element.attrs.get("x")) ?? 0;
      const y = parseNumber(element.attrs.get("y")) ?? 0;
      const width = parseNumber(element.attrs.get("width")) ?? 1;
      const height = parseNumber(element.attrs.get("height")) ?? 1;
      if (
        !Number.isFinite(x) ||
        !Number.isFinite(y) ||
        !Number.isFinite(width) ||
        !Number.isFinite(height) ||
        width <= EPSILON ||
        height <= EPSILON
      ) {
        return null;
      }
      return { minX: x, minY: y, maxX: x + width, maxY: y + height };
    }
    default:
      return null;
  }
}

type DefState = {
  readonly droppedClipIds: Set<string>;
  readonly droppedMaskIds: Set<string>;
  readonly droppedPatternIds: Set<string>;
  readonly droppedFilterIds: Set<string>;
  readonly droppedGradientIds: Set<string>;
  readonly gradientFallbacks: Map<
    string,
    { readonly color: string; readonly opacity?: string }
  >;
};

function sanitizeDefsBlock(
  defsElement: SvgElement,
  canvas: { readonly width: number; readonly height: number },
  safeViewport: BBox,
  state: DefState,
): void {
  const keptDefs: SvgNode[] = [];

  for (const child of defsElement.children) {
    if (child.kind !== "element") continue;
    const id = child.attrs.get("id");

    switch (child.tag) {
      case "clipPath": {
        const units = child.attrs.get("clipPathUnits") ?? "userSpaceOnUse";
        const keptChildren: SvgNode[] = [];
        for (const item of child.children) {
          if (item.kind !== "element") continue;
          const localMatrix = parseTransformAttr(item.attrs.get("transform"));
          if (localMatrix === null) continue;
          const bbox = sanitizeGeometryElementBbox(
            item,
            canvas,
            safeViewport,
            true,
          );
          if (bbox === null) continue;
          if (units === "userSpaceOnUse") {
            const worldBox = transformBbox(bbox, localMatrix);
            if (!boxesIntersect(worldBox, safeViewport)) continue;
          }
          keptChildren.push(item);
        }
        if (keptChildren.length === 0) {
          if (id !== undefined) state.droppedClipIds.add(id);
          continue;
        }
        child.children.length = 0;
        child.children.push(...keptChildren);
        keptDefs.push(child);
        break;
      }

      case "linearGradient": {
        if (hasNonFiniteAttr(child, ["x1", "y1", "x2", "y2"])) {
          if (id !== undefined) {
            const fallback = lastGradientStopColor(child);
            if (fallback !== undefined)
              state.gradientFallbacks.set(id, fallback);
            else state.droppedGradientIds.add(id);
          }
          continue;
        }
        const x1 = parseLengthAttr(child.attrs.get("x1"), 1, 0);
        const y1 = parseLengthAttr(child.attrs.get("y1"), 1, 0);
        const x2 = parseLengthAttr(child.attrs.get("x2"), 1, 1);
        const y2 = parseLengthAttr(child.attrs.get("y2"), 1, 0);
        if (
          !Number.isFinite(x1) ||
          !Number.isFinite(y1) ||
          !Number.isFinite(x2) ||
          !Number.isFinite(y2) ||
          Math.hypot(x2 - x1, y2 - y1) <= EPSILON
        ) {
          if (id !== undefined) {
            const fallback = lastGradientStopColor(child);
            if (fallback !== undefined)
              state.gradientFallbacks.set(id, fallback);
            else state.droppedGradientIds.add(id);
          }
          continue;
        }
        keptDefs.push(child);
        break;
      }

      case "radialGradient": {
        if (hasNonFiniteAttr(child, ["cx", "cy", "r", "fx", "fy"])) {
          if (id !== undefined) {
            const fallback = lastGradientStopColor(child);
            if (fallback !== undefined)
              state.gradientFallbacks.set(id, fallback);
            else state.droppedGradientIds.add(id);
          }
          continue;
        }
        const r = parseLengthAttr(child.attrs.get("r"), 1, 0.5);
        if (!Number.isFinite(r) || r <= EPSILON) {
          if (id !== undefined) {
            const fallback = lastGradientStopColor(child);
            if (fallback !== undefined)
              state.gradientFallbacks.set(id, fallback);
            else state.droppedGradientIds.add(id);
          }
          continue;
        }
        keptDefs.push(child);
        break;
      }

      case "filter": {
        if (hasNonFiniteAttr(child, ["x", "y", "width", "height"])) {
          if (id !== undefined) state.droppedFilterIds.add(id);
          continue;
        }
        const units = child.attrs.get("filterUnits") ?? "objectBoundingBox";
        if (units === "userSpaceOnUse") {
          const x = parseNumber(child.attrs.get("x")) ?? 0;
          const y = parseNumber(child.attrs.get("y")) ?? 0;
          const width = parseNumber(child.attrs.get("width")) ?? 0;
          const height = parseNumber(child.attrs.get("height")) ?? 0;
          const clamped = clampRectToSafeViewport(
            x,
            y,
            width,
            height,
            safeViewport,
          );
          if (clamped === null) {
            if (id !== undefined) state.droppedFilterIds.add(id);
            continue;
          }
          child.attrs.set("x", formatNumber(clamped.x));
          child.attrs.set("y", formatNumber(clamped.y));
          child.attrs.set("width", formatNumber(clamped.width));
          child.attrs.set("height", formatNumber(clamped.height));
        }
        let validFilter = true;
        for (const prim of child.children) {
          if (prim.kind !== "element") continue;
          if (hasNonFiniteAttr(prim, ["stdDeviation", "dx", "dy"])) {
            validFilter = false;
            break;
          }
          const stdRaw = prim.attrs.get("stdDeviation");
          if (stdRaw !== undefined) {
            const std = parseNumber(stdRaw);
            if (std === undefined || !Number.isFinite(std) || std < 0) {
              validFilter = false;
              break;
            }
            if (std > MAX_BLUR_STD_DEVIATION) {
              prim.attrs.set(
                "stdDeviation",
                formatNumber(MAX_BLUR_STD_DEVIATION),
              );
            }
          }
        }
        if (!validFilter) {
          if (id !== undefined) state.droppedFilterIds.add(id);
          continue;
        }
        keptDefs.push(child);
        break;
      }

      case "mask": {
        if (hasNonFiniteAttr(child, ["x", "y", "width", "height"])) {
          if (id !== undefined) state.droppedMaskIds.add(id);
          continue;
        }
        const units = child.attrs.get("maskUnits") ?? "objectBoundingBox";
        if (units === "userSpaceOnUse") {
          const x = parseNumber(child.attrs.get("x")) ?? 0;
          const y = parseNumber(child.attrs.get("y")) ?? 0;
          const width = parseNumber(child.attrs.get("width")) ?? 0;
          const height = parseNumber(child.attrs.get("height")) ?? 0;
          const clamped = clampRectToSafeViewport(
            x,
            y,
            width,
            height,
            safeViewport,
          );
          if (clamped === null) {
            if (id !== undefined) state.droppedMaskIds.add(id);
            continue;
          }
          child.attrs.set("x", formatNumber(clamped.x));
          child.attrs.set("y", formatNumber(clamped.y));
          child.attrs.set("width", formatNumber(clamped.width));
          child.attrs.set("height", formatNumber(clamped.height));
        }
        keptDefs.push(child);
        break;
      }

      case "pattern": {
        if (hasNonFiniteAttr(child, ["x", "y", "width", "height"])) {
          if (id !== undefined) state.droppedPatternIds.add(id);
          continue;
        }
        const units = child.attrs.get("patternUnits") ?? "objectBoundingBox";
        if (units === "userSpaceOnUse") {
          const x = parseNumber(child.attrs.get("x")) ?? 0;
          const y = parseNumber(child.attrs.get("y")) ?? 0;
          const width = parseNumber(child.attrs.get("width")) ?? 0;
          const height = parseNumber(child.attrs.get("height")) ?? 0;
          const clamped = clampRectToSafeViewport(
            x,
            y,
            width,
            height,
            safeViewport,
          );
          if (clamped === null) {
            if (id !== undefined) state.droppedPatternIds.add(id);
            continue;
          }
          if (
            x < safeViewport.minX ||
            y < safeViewport.minY ||
            x + width > safeViewport.maxX ||
            y + height > safeViewport.maxY
          ) {
            child.attrs.set("x", formatNumber(clamped.x));
            child.attrs.set("y", formatNumber(clamped.y));
            child.attrs.set("width", formatNumber(clamped.width));
            child.attrs.set("height", formatNumber(clamped.height));
            for (const pChild of child.children) {
              if (pChild.kind !== "element") continue;
              if (pChild.tag === "rect" || pChild.tag === "image") {
                const cw = parseNumber(pChild.attrs.get("width")) ?? 0;
                const ch = parseNumber(pChild.attrs.get("height")) ?? 0;
                if (cw > clamped.width) {
                  pChild.attrs.set("width", formatNumber(clamped.width));
                }
                if (ch > clamped.height) {
                  pChild.attrs.set("height", formatNumber(clamped.height));
                }
              }
            }
          }
        }
        keptDefs.push(child);
        break;
      }

      default:
        keptDefs.push(child);
        break;
    }
  }

  defsElement.children.length = 0;
  defsElement.children.push(...keptDefs);
}

function applyPaintRefFallbacks(element: SvgElement, state: DefState): boolean {
  const clipId = extractUrlId(element.attrs.get("clip-path"));
  if (clipId !== undefined && state.droppedClipIds.has(clipId)) {
    return false;
  }
  const maskId = extractUrlId(element.attrs.get("mask"));
  if (maskId !== undefined && state.droppedMaskIds.has(maskId)) {
    return false;
  }
  const filterId = extractUrlId(element.attrs.get("filter"));
  if (filterId !== undefined && state.droppedFilterIds.has(filterId)) {
    element.attrs.delete("filter");
  }

  for (const attr of ["fill", "stroke"] as const) {
    const refId = extractUrlId(element.attrs.get(attr));
    if (refId === undefined) continue;
    const fallback = state.gradientFallbacks.get(refId);
    if (fallback !== undefined) {
      element.attrs.set(attr, fallback.color);
      if (fallback.opacity !== undefined) {
        element.attrs.set(`${attr}-opacity`, fallback.opacity);
      }
      continue;
    }
    if (
      state.droppedGradientIds.has(refId) ||
      state.droppedPatternIds.has(refId)
    ) {
      element.attrs.set(attr, "none");
    }
  }

  return true;
}

function sanitizeBodyNode(
  node: SvgNode,
  ctm: Matrix2D,
  canvas: { readonly width: number; readonly height: number },
  safeViewport: BBox,
  state: DefState,
): SvgNode | null {
  if (node.kind === "text") return node;
  if (node.tag === "defs") {
    sanitizeDefsBlock(node, canvas, safeViewport, state);
    return node.children.length === 0 ? null : node;
  }

  const localMatrix = parseTransformAttr(node.attrs.get("transform"));
  if (localMatrix === null) return null;
  const nextCtm = multiplyMatrix(ctm, localMatrix);
  if (nextCtm.some((v) => !Number.isFinite(v))) return null;

  if (!applyPaintRefFallbacks(node, state)) {
    return null;
  }

  if (node.tag === "g") {
    const keptChildren: SvgNode[] = [];
    for (const child of node.children) {
      const sanitizedChild = sanitizeBodyNode(
        child,
        nextCtm,
        canvas,
        safeViewport,
        state,
      );
      if (sanitizedChild !== null) keptChildren.push(sanitizedChild);
    }
    if (keptChildren.length === 0) return null;
    node.children.length = 0;
    node.children.push(...keptChildren);
    return node;
  }

  const hasStroke =
    node.attrs.get("stroke") !== undefined &&
    node.attrs.get("stroke") !== "none";
  if (
    hasNonFiniteAttr(node, [
      "stroke-width",
      "opacity",
      "fill-opacity",
      "stroke-opacity",
    ])
  ) {
    return null;
  }
  const strokeWidth = parseNumber(node.attrs.get("stroke-width")) ?? 0;
  if (!Number.isFinite(strokeWidth) || strokeWidth < 0) return null;

  const isPercentageRect =
    node.tag === "rect" &&
    (node.attrs.get("width")?.trim().endsWith("%") ||
      node.attrs.get("height")?.trim().endsWith("%"));
  if (isPercentageRect) {
    return node;
  }

  const localBbox = sanitizeGeometryElementBbox(
    node,
    canvas,
    safeViewport,
    !hasStroke,
  );
  if (localBbox === null) return null;

  const pad = hasStroke ? Math.max(2, strokeWidth) : 0;
  const paddedLocal: BBox = {
    minX: localBbox.minX - pad,
    minY: localBbox.minY - pad,
    maxX: localBbox.maxX + pad,
    maxY: localBbox.maxY + pad,
  };
  const worldBbox = transformBbox(paddedLocal, nextCtm);
  if (!boxesIntersect(worldBbox, safeViewport)) {
    return null;
  }

  if (
    !hasStroke &&
    node.attrs.get("fill") === "none" &&
    node.attrs.get("clip-path") === undefined
  ) {
    return null;
  }

  return node;
}

type RasterOccurrence = {
  readonly parent: SvgElement;
  readonly childIndex: number;
  readonly element: SvgElement;
  readonly entry: RasterEntry;
  readonly x: string;
  readonly y: string;
  readonly width: string | undefined;
  readonly height: string | undefined;
  readonly preserveAspectRatio: string | undefined;
  readonly groupKey: string | undefined;
};

function resolveRasterTokensInSvgTree(
  svgRoot: SvgElement,
  registry: PageRasterRegistry,
): void {
  const occurrences: RasterOccurrence[] = [];
  const groups = new Map<string, RasterOccurrence[]>();

  const visitChildren = (parent: SvgElement): void => {
    for (let i = 0; i < parent.children.length; i += 1) {
      const child = parent.children[i];
      if (child === undefined || child.kind !== "element") continue;
      if (child.tag === "image") {
        const href = child.attrs.get("href");
        if (href !== undefined && href.startsWith(RASTER_TOKEN_PREFIX)) {
          const id = Number(href.slice(RASTER_TOKEN_PREFIX.length));
          const entry =
            Number.isInteger(id) && id >= 0 ? registry.getEntry(id) : undefined;
          if (entry !== undefined) {
            const x = child.attrs.get("x") ?? "0";
            const y = child.attrs.get("y") ?? "0";
            const width = child.attrs.get("width");
            const height = child.attrs.get("height");
            const preserveAspectRatio = child.attrs.get("preserveAspectRatio");
            const numX = Number(x);
            const numY = Number(y);
            const numW = Number(width);
            const numH = Number(height);
            const canShare =
              width !== undefined &&
              height !== undefined &&
              Number.isFinite(numX) &&
              Number.isFinite(numY) &&
              Number.isFinite(numW) &&
              Number.isFinite(numH) &&
              numW > EPSILON &&
              numH > EPSILON;
            const groupKey = !canShare
              ? undefined
              : preserveAspectRatio === "none"
                ? `${entry.id}|none`
                : `${entry.id}|${preserveAspectRatio ?? ""}|${width}|${height}`;
            const occurrence: RasterOccurrence = {
              parent,
              childIndex: i,
              element: child,
              entry,
              x,
              y,
              width,
              height,
              preserveAspectRatio,
              groupKey,
            };
            occurrences.push(occurrence);
            if (groupKey !== undefined) {
              let bucket = groups.get(groupKey);
              if (bucket === undefined) {
                bucket = [];
                groups.set(groupKey, bucket);
              }
              bucket.push(occurrence);
            }
          }
        }
      }
      visitChildren(child);
    }
  };

  visitChildren(svgRoot);
  if (occurrences.length === 0) return;

  // Single surviving occurrences keep plain <image href="data:..."/> without <defs> hoisting.
  for (const occurrence of occurrences) {
    if (
      occurrence.groupKey === undefined ||
      (groups.get(occurrence.groupKey)?.length ?? 0) < 2
    ) {
      occurrence.element.attrs.set(
        "href",
        registry.getDataUri(occurrence.entry),
      );
    }
  }

  const sharedDefImages: SvgNode[] = [];
  let sharedSeq = 0;
  for (const members of groups.values()) {
    if (members.length < 2) continue;
    const first = members[0]!;
    sharedSeq += 1;
    const sharedId = `${registry.idPrefix}-${sharedSeq}`;
    const isUnitStretch = first.preserveAspectRatio === "none";
    const defAttrs = new Map<string, string>([
      ["id", sharedId],
      ["x", "0"],
      ["y", "0"],
      ["width", isUnitStretch ? "1" : first.width!],
      ["height", isUnitStretch ? "1" : first.height!],
    ]);
    if (first.preserveAspectRatio !== undefined) {
      defAttrs.set("preserveAspectRatio", first.preserveAspectRatio);
    }
    defAttrs.set("href", registry.getDataUri(first.entry));
    sharedDefImages.push({
      kind: "element",
      tag: "image",
      attrs: defAttrs,
      children: [],
      selfClosing: true,
    });

    for (const member of members) {
      const passthroughAttrs = new Map<string, string>();
      for (const [key, value] of member.element.attrs.entries()) {
        if (
          key === "href" ||
          key === "x" ||
          key === "y" ||
          key === "width" ||
          key === "height" ||
          key === "preserveAspectRatio"
        ) {
          continue;
        }
        passthroughAttrs.set(key, value);
      }
      const useAttrs = new Map<string, string>([["href", `#${sharedId}`]]);
      const transform = isUnitStretch
        ? `translate(${member.x} ${member.y}) scale(${member.width!} ${member.height!})`
        : Number(member.x) === 0 && Number(member.y) === 0
          ? undefined
          : `translate(${member.x} ${member.y})`;
      if (transform !== undefined) {
        useAttrs.set("transform", transform);
      }
      const useNode: SvgElement = {
        kind: "element",
        tag: "use",
        attrs: useAttrs,
        children: [],
        selfClosing: true,
      };
      member.parent.children[member.childIndex] =
        passthroughAttrs.size === 0
          ? useNode
          : {
              kind: "element",
              tag: "g",
              attrs: passthroughAttrs,
              children: [useNode],
              selfClosing: false,
            };
    }
  }

  if (sharedDefImages.length > 0) {
    svgRoot.children.unshift({
      kind: "element",
      tag: "defs",
      attrs: new Map(),
      children: sharedDefImages,
      selfClosing: false,
    });
  }
}

/**
 * Sanitizes an assembled single-page SVG before handing it to `@resvg/resvg-js`,
 * eliminating zero-area clipPaths/masks/filters, degenerate gradients, non-finite numbers,
 * and extreme off-canvas geometry that trigger Rust `geom.rs` panics, then materializes
 * registered page raster tokens into either inline `<image>` or shared `<defs>` + `<use>`.
 */
export function sanitizeSvgForRasterization(
  svg: string,
  dimensions: { readonly width: number; readonly height: number },
  rasterRegistry?: PageRasterRegistry,
): string {
  const width =
    Number.isFinite(dimensions.width) && dimensions.width > EPSILON
      ? dimensions.width
      : 1;
  const height =
    Number.isFinite(dimensions.height) && dimensions.height > EPSILON
      ? dimensions.height
      : 1;
  const safeViewport: BBox = {
    minX: -SAFE_OFFSCREEN_PAD,
    minY: -SAFE_OFFSCREEN_PAD,
    maxX: width + SAFE_OFFSCREEN_PAD,
    maxY: height + SAFE_OFFSCREEN_PAD,
  };
  const state: DefState = {
    droppedClipIds: new Set(),
    droppedMaskIds: new Set(),
    droppedPatternIds: new Set(),
    droppedFilterIds: new Set(),
    droppedGradientIds: new Set(),
    gradientFallbacks: new Map(),
  };

  const nodes = parseSvgTree(svg);
  const svgRoot = nodes.find(
    (node): node is SvgElement => node.kind === "element" && node.tag === "svg",
  );
  if (svgRoot === undefined) return svg;

  // Sanitize all <defs> blocks first so forward/backward url(#id) references see final states.
  const collectDefs = (el: SvgElement): void => {
    for (const child of el.children) {
      if (child.kind !== "element") continue;
      if (child.tag === "defs") {
        sanitizeDefsBlock(child, { width, height }, safeViewport, state);
      } else {
        collectDefs(child);
      }
    }
  };
  collectDefs(svgRoot);

  const keptRootChildren: SvgNode[] = [];
  for (const child of svgRoot.children) {
    const sanitized = sanitizeBodyNode(
      child,
      IDENTITY_MATRIX,
      { width, height },
      safeViewport,
      state,
    );
    if (sanitized !== null) keptRootChildren.push(sanitized);
  }
  svgRoot.children.length = 0;
  svgRoot.children.push(...keptRootChildren);

  if (rasterRegistry !== undefined && rasterRegistry.size > 0) {
    resolveRasterTokensInSvgTree(svgRoot, rasterRegistry);
  }

  return nodes.map(serializeSvgNode).join("");
}
