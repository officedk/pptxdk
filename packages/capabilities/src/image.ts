import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

import {
  assertXml10Characters,
  CapabilityError,
  isMorphKey,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  xmlEscape,
  type BoundObjectCapabilityModule,
  type BoundObjectNativeCollaboration,
  type CapabilityInput,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityParseContext,
  type EmitContext,
  type EmitOutput,
  type GeometryBox,
  type LayoutStyle,
  type LinkTarget,
  type MorphKey,
  type ObjectResolveContext,
  type PaintContext,
  type PaintOutput,
  type ResolvedBoundObjectCapability,
  type ResolvedObject,
  type VisualStyle,
} from "@react-pptx/core/capability";
import { CommonNode } from "@react-pptx/core/document";

import type { Geometry } from "./geometry.ts";
import { nativeLinkXml, parseLinkTarget } from "./identity.ts";
import { imageXml } from "./image-xml.ts";
import type { MorphProps } from "./morph-api.ts";
import { motionAnnotations, type MotionProps } from "./motion-api.ts";
import { shapePropertiesXml } from "./visual.ts";

/** Visual styling supported on Image objects. */
export type ImageVisualStyle = Pick<
  VisualStyle,
  | "borderColor"
  | "borderWidth"
  | "borderStyle"
  | "opacity"
  | "boxShadow"
  | "softEdge"
  | "rotate"
>;

export const imageVisualStyleKeys = new Set([
  "borderColor",
  "borderWidth",
  "borderStyle",
  "opacity",
  "boxShadow",
  "softEdge",
  "rotate",
]);

/** Local file path, file URL, or raw image bytes. */
export type ImageSource = string | URL | Uint8Array;

export type ImageDeclarationProps = MotionProps &
  MorphProps & {
    readonly id?: string;
    readonly name?: string;
    readonly link?: LinkTarget;
    readonly src: ImageSource;
    /** Alternative text for the image. */
    readonly alt?: string;
    /** Position and bounds in px; position defaults to 0, width and height default to original image dimensions. */
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly style?: LayoutStyle & ImageVisualStyle;
    /** `fill` stretches, `cover` crops centered, `contain` preserves aspect ratio; defaults to `fill`. */
    readonly fit?: "fill" | "cover" | "contain";
    /** Explicit crop fraction for each side of the original image; takes precedence over `fit` when provided. */
    readonly crop?: {
      readonly left?: number;
      readonly top?: number;
      readonly right?: number;
      readonly bottom?: number;
    };
  };

/** Creates the editable native image shared by the command factory and the Image JSX element. */
export function createImageNode(
  props: ImageDeclarationProps,
): CommonNode<
  "image",
  ImageDeclarationProps,
  NonNullable<ImageDeclarationProps["style"]>
> {
  return new CommonNode<
    "image",
    ImageDeclarationProps,
    NonNullable<ImageDeclarationProps["style"]>
  >("image", "rp:image", props);
}

type ImageFit = "fill" | "cover" | "contain";

type ImageMetadata = {
  readonly extension: "png" | "jpeg" | "gif";
  readonly contentType: "image/png" | "image/jpeg" | "image/gif";
  readonly width: number;
  readonly height: number;
};

export type LoadedImage = ImageMetadata & { readonly bytes: Uint8Array };

type ImagePlacement = {
  readonly box: GeometryBox;
  readonly sourceCrop: Readonly<Record<string, number>>;
};

interface PngInstance {
  width: number;
  height: number;
  data: Buffer;
}

interface PngJsModule {
  readonly PNG: {
    new (options: { width: number; height: number }): PngInstance;
    readonly sync: {
      read(buffer: Buffer): PngInstance;
      write(png: PngInstance): Buffer;
    };
  };
}

interface JpegJsModule {
  decode(
    jpegData: Uint8Array,
    options?: { readonly tolerantDecoding?: boolean },
  ): {
    readonly width: number;
    readonly height: number;
    readonly data: Uint8Array;
  };
}

let cachedDecoders:
  | { readonly png: PngJsModule; readonly jpeg: JpegJsModule }
  | undefined;

// Load pngjs and jpeg-js lazily only when paint() executes so resolve() and emit() never pull rasterizers.
function imageDecoders(): {
  readonly png: PngJsModule;
  readonly jpeg: JpegJsModule;
} {
  if (cachedDecoders === undefined) {
    const require = createRequire(import.meta.url);
    cachedDecoders = {
      png: require("pngjs") as PngJsModule,
      jpeg: require("jpeg-js") as JpegJsModule,
    };
  }
  return cachedDecoders;
}

const unsupportedPreviewImageStyleKeys = [
  "borderColor",
  "borderWidth",
  "borderStyle",
  "boxShadow",
  "softEdge",
] as const;

function isRectMask(geometry: Geometry | undefined): boolean {
  if (geometry === undefined || geometry === "rect") return true;
  return (
    typeof geometry === "object" &&
    geometry.type === "preset" &&
    geometry.name === "rect" &&
    (geometry.adjust === undefined || Object.keys(geometry.adjust).length === 0)
  );
}

function renderImagePreview(options: {
  readonly image: LoadedImage;
  readonly visualStyle: Readonly<Record<string, unknown>>;
  readonly geometry: Geometry | undefined;
  readonly flipH: boolean | undefined;
  readonly flipV: boolean | undefined;
  readonly placement: ImagePlacement;
  readonly objectId: number;
}): PaintOutput {
  const { image, visualStyle, geometry, flipH, flipV, placement, objectId } =
    options;

  if (!isRectMask(geometry)) {
    throw new CapabilityError(
      "preview-unsupported-geometry",
      "Preview does not yet support shaped images",
    );
  }
  for (const key of unsupportedPreviewImageStyleKeys) {
    if (visualStyle[key] !== undefined) {
      throw new CapabilityError(
        "preview-unsupported-style",
        `Preview does not yet support image ${key}`,
      );
    }
  }
  if (image.extension !== "png" && image.extension !== "jpeg") {
    throw new CapabilityError(
      "preview-unsupported-image",
      "Preview supports only PNG and JPEG images",
    );
  }

  const { png: pngModule, jpeg: jpegModule } = imageDecoders();
  const bytes = Buffer.from(image.bytes);
  let decoded: {
    readonly width: number;
    readonly height: number;
    readonly data: Uint8Array;
  };
  try {
    decoded =
      image.extension === "png"
        ? pngModule.PNG.sync.read(bytes)
        : jpegModule.decode(bytes, { tolerantDecoding: false });
  } catch (error) {
    throw new CapabilityError(
      "preview-corrupt-image",
      error instanceof Error
        ? `Failed to decode ${image.extension.toUpperCase()} image for preview: ${error.message}`
        : `Failed to decode ${image.extension.toUpperCase()} image for preview`,
    );
  }

  const normalized = new pngModule.PNG({
    width: decoded.width,
    height: decoded.height,
  });
  normalized.data = Buffer.from(decoded.data);
  const pngBytes = pngModule.PNG.sync.write(normalized);

  const { box, sourceCrop } = placement;
  const cropSide = (side: "l" | "t" | "r" | "b") =>
    Math.round((sourceCrop[side] ?? 0) * 100000) / 100000;
  const leftCrop = cropSide("l");
  const topCrop = cropSide("t");
  const rightCrop = cropSide("r");
  const bottomCrop = cropSide("b");
  const widthFactor = 1 - leftCrop - rightCrop;
  const heightFactor = 1 - topCrop - bottomCrop;

  const angle =
    typeof visualStyle["rotate"] === "number" ? visualStyle["rotate"] : 0;
  const flipX = Boolean(flipH);
  const flipY = Boolean(flipV);
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const opacity = Number(visualStyle["opacity"] ?? 1);
  const transform = `translate(${cx} ${cy}) rotate(${angle}) scale(${flipX ? -1 : 1} ${flipY ? -1 : 1}) translate(${-cx} ${-cy})`;

  if (widthFactor <= 0 || heightFactor <= 0) {
    return {
      svg: `<g transform="${transform}" opacity="${opacity}"></g>`,
    };
  }

  const width = box.width / widthFactor;
  const height = box.height / heightFactor;
  const clipId = `image-${objectId}`;
  const svg = `<g transform="${transform}" opacity="${opacity}"><defs><clipPath id="${clipId}"><rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"/></clipPath></defs><image clip-path="url(#${clipId})" x="${box.x - leftCrop * width}" y="${box.y - topCrop * height}" width="${width}" height="${height}" preserveAspectRatio="none" href="data:image/png;base64,${pngBytes.toString("base64")}"/></g>`;

  return {
    svg,
  };
}

/** Shared image placement calculation for native OOXML output and static preview. */
function imagePlacement(
  image: { readonly width: number; readonly height: number },
  fit: ImageFit,
  crop: Readonly<Record<string, number>> | undefined,
  allocated: GeometryBox,
): ImagePlacement {
  const { width, height } = allocated;
  const box = { ...allocated };
  const sourceCrop = crop !== undefined ? { ...crop } : {};
  if (crop === undefined && fit === "cover") {
    const imageRatio = image.width / image.height;
    const boxRatio = width / height;
    if (imageRatio > boxRatio) {
      const excess = (1 - boxRatio / imageRatio) / 2;
      sourceCrop["l"] = excess;
      sourceCrop["r"] = excess;
    } else if (imageRatio < boxRatio) {
      const excess = (1 - imageRatio / boxRatio) / 2;
      sourceCrop["t"] = excess;
      sourceCrop["b"] = excess;
    }
  } else if (crop === undefined && fit === "contain") {
    const imageRatio = image.width / image.height;
    const boxRatio = width / height;
    if (imageRatio > boxRatio) {
      const fittedHeight = width / imageRatio;
      box.y += (height - fittedHeight) / 2;
      box.height = fittedHeight;
    } else if (imageRatio < boxRatio) {
      const fittedWidth = height * imageRatio;
      box.x += (width - fittedWidth) / 2;
      box.width = fittedWidth;
    }
  }
  // Dimensions derived by contain must also be representable as non-zero EMU.
  box.width = dimension(box.width, "width");
  box.height = dimension(box.height, "height");
  return { box, sourceCrop };
}

const allowedParameters = [
  "id",
  "name",
  "morphKey",
  "link",
  "src",
  "alt",
  "x",
  "y",
  "width",
  "height",
  "geometry",
  "flipH",
  "flipV",
  "style",
  "fit",
  "crop",
  "enter",
  "emphasis",
  "motion",
] as const;
const maximumShapeExtentEmu = 2_147_483_647;

function uint32(data: Uint8Array, offset: number): number {
  return (
    (((data[offset] ?? 0) << 24) >>> 0) +
    ((data[offset + 1] ?? 0) << 16) +
    ((data[offset + 2] ?? 0) << 8) +
    (data[offset + 3] ?? 0)
  );
}

function uint16(data: Uint8Array, offset: number): number {
  return ((data[offset] ?? 0) << 8) + (data[offset + 1] ?? 0);
}

function pngMetadata(data: Uint8Array): ImageMetadata | undefined {
  if (
    data.length < 45 ||
    data[0] !== 0x89 ||
    data[1] !== 0x50 ||
    data[2] !== 0x4e ||
    data[3] !== 0x47 ||
    data[4] !== 0x0d ||
    data[5] !== 0x0a ||
    data[6] !== 0x1a ||
    data[7] !== 0x0a ||
    uint32(data, 8) !== 13 ||
    data[12] !== 0x49 ||
    data[13] !== 0x48 ||
    data[14] !== 0x44 ||
    data[15] !== 0x52
  ) {
    return undefined;
  }
  let offset = 8;
  let hasHeader = false;
  let hasImageData = false;
  let hasEnd = false;
  while (offset + 12 <= data.length) {
    const size = uint32(data, offset);
    const end = offset + size + 12;
    if (end > data.length) return undefined;
    const isHeader =
      data[offset + 4] === 0x49 &&
      data[offset + 5] === 0x48 &&
      data[offset + 6] === 0x44 &&
      data[offset + 7] === 0x52;
    const isEnd =
      data[offset + 4] === 0x49 &&
      data[offset + 5] === 0x45 &&
      data[offset + 6] === 0x4e &&
      data[offset + 7] === 0x44;
    const isImageData =
      data[offset + 4] === 0x49 &&
      data[offset + 5] === 0x44 &&
      data[offset + 6] === 0x41 &&
      data[offset + 7] === 0x54;
    if (offset === 8 && !isHeader) return undefined;
    if (isHeader) hasHeader = true;
    if (isImageData && size > 0) hasImageData = true;
    if (isEnd) {
      hasEnd = size === 0;
      break;
    }
    offset = end;
  }
  const width = uint32(data, 16);
  const height = uint32(data, 20);
  if (!hasHeader || !hasImageData || !hasEnd || width < 1 || height < 1) {
    return undefined;
  }
  return { extension: "png", contentType: "image/png", width, height };
}

function jpegMetadata(data: Uint8Array): ImageMetadata | undefined {
  if (data.length < 12 || data[0] !== 0xff || data[1] !== 0xd8) {
    return undefined;
  }
  let offset = 2;
  let frame: ImageMetadata | undefined;
  while (offset + 4 <= data.length) {
    if (data[offset] !== 0xff) return undefined;
    while (data[offset] === 0xff) offset += 1;
    const marker = data[offset];
    if (marker === undefined || marker === 0xd9) return undefined;
    if (marker === 0xda) {
      const length = uint16(data, offset + 1);
      if (
        frame === undefined ||
        length < 2 ||
        offset + length + 1 > data.length
      ) {
        return undefined;
      }
      let hasScanData = false;
      for (let index = offset + length + 1; index + 1 < data.length; index++) {
        if (data[index] !== 0xff) {
          hasScanData = true;
          continue;
        }
        let markerIndex = index + 1;
        while (data[markerIndex] === 0xff) markerIndex += 1;
        const scanMarker = data[markerIndex];
        if (scanMarker === 0xd9) return hasScanData ? frame : undefined;
        if (
          scanMarker === 0x00 ||
          (scanMarker !== undefined && scanMarker >= 0xd0 && scanMarker <= 0xd7)
        ) {
          hasScanData = true;
        }
        index = markerIndex;
      }
      return undefined;
    }
    if (
      marker === 0xd8 ||
      marker === 0x01 ||
      (marker >= 0xd0 && marker <= 0xd7)
    ) {
      offset += 1;
      continue;
    }
    const length = uint16(data, offset + 1);
    if (length < 2 || offset + length + 1 > data.length) return undefined;
    const isFrame =
      marker >= 0xc0 &&
      marker <= 0xcf &&
      marker !== 0xc4 &&
      marker !== 0xc8 &&
      marker !== 0xcc;
    if (isFrame) {
      const height = uint16(data, offset + 4);
      const width = uint16(data, offset + 6);
      if (width < 1 || height < 1) return undefined;
      frame = { extension: "jpeg", contentType: "image/jpeg", width, height };
    }
    offset += length + 1;
  }
  return undefined;
}

function gifMetadata(data: Uint8Array): ImageMetadata | undefined {
  const signature = String.fromCharCode(...data.subarray(0, 6));
  if (
    data.length < 14 ||
    (signature !== "GIF87a" && signature !== "GIF89a") ||
    data[data.length - 1] !== 0x3b
  ) {
    return undefined;
  }
  const width = (data[6] ?? 0) + ((data[7] ?? 0) << 8);
  const height = (data[8] ?? 0) + ((data[9] ?? 0) << 8);
  if (width < 1 || height < 1) return undefined;
  let offset = 13;
  const globalTable = data[10] ?? 0;
  if ((globalTable & 0x80) !== 0) {
    offset += 3 * (1 << ((globalTable & 0x07) + 1));
    if (offset > data.length) return undefined;
  }
  let hasImageData = false;
  while (offset < data.length) {
    const block = data[offset];
    if (block === 0x3b) {
      if (offset !== data.length - 1 || !hasImageData) return undefined;
      return { extension: "gif", contentType: "image/gif", width, height };
    }
    if (block === 0x21) {
      if (offset + 1 >= data.length) return undefined;
      offset += 2;
      let terminated = false;
      while (offset < data.length) {
        const size = data[offset] ?? 0;
        offset += 1;
        if (size === 0) {
          terminated = true;
          break;
        }
        offset += size;
        if (offset > data.length) return undefined;
      }
      if (!terminated) return undefined;
      continue;
    }
    if (block !== 0x2c || offset + 10 > data.length) return undefined;
    const imageWidth = (data[offset + 5] ?? 0) + ((data[offset + 6] ?? 0) << 8);
    const imageHeight =
      (data[offset + 7] ?? 0) + ((data[offset + 8] ?? 0) << 8);
    const localTable = data[offset + 9] ?? 0;
    if (imageWidth < 1 || imageHeight < 1) return undefined;
    offset += 10;
    if ((localTable & 0x80) !== 0) {
      offset += 3 * (1 << ((localTable & 0x07) + 1));
      if (offset > data.length) return undefined;
    }
    if (offset >= data.length) return undefined;
    offset += 1;
    let hasImageBlockData = false;
    let terminated = false;
    while (offset < data.length) {
      const size = data[offset] ?? 0;
      offset += 1;
      if (size === 0) {
        terminated = true;
        break;
      }
      offset += size;
      if (offset > data.length) return undefined;
      hasImageBlockData = true;
    }
    if (!terminated || !hasImageBlockData) return undefined;
    hasImageData = true;
  }
  return undefined;
}

export function readImage(source: unknown): LoadedImage {
  if (
    typeof source !== "string" &&
    !(source instanceof URL) &&
    !(source instanceof Uint8Array)
  ) {
    throw new CapabilityError(
      "invalid-image-source",
      "Image.src must be a file path, file URL, or image bytes",
    );
  }
  let bytes: Uint8Array;
  try {
    bytes =
      source instanceof Uint8Array
        ? new Uint8Array(source)
        : new Uint8Array(
            readFileSync(
              source instanceof URL ? fileURLToPath(source) : source,
            ),
          );
  } catch {
    throw new CapabilityError(
      "image-read-failed",
      "Failed to read image resource",
    );
  }
  const metadata =
    pngMetadata(bytes) ?? jpegMetadata(bytes) ?? gifMetadata(bytes);
  if (metadata === undefined) {
    throw new CapabilityError(
      "unsupported-image-format",
      "Invalid image resource content or unsupported format; only PNG, JPEG, and GIF are supported",
    );
  }
  return { ...metadata, bytes };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function optionalText(
  props: Record<string, unknown>,
  name: "name",
  path?: string,
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError(
      "invalid-parameter",
      `${name} must be a non-empty string`,
      path,
    );
  }
  return value;
}

function dimension(
  value: unknown,
  name: "x" | "y" | "width" | "height",
  fallback?: number,
): number {
  const actual = value === undefined ? fallback : value;
  const isExtent = name === "width" || name === "height";
  if (
    typeof actual !== "number" ||
    !Number.isFinite(actual) ||
    (isExtent ? actual <= 0 : actual < -1_000_000 || actual > 1_000_000)
  ) {
    throw new CapabilityError(
      "invalid-coordinate",
      `${name} must be a finite pixel number, width and height must be greater than 0, and coordinates must be between -1000000 and 1000000`,
    );
  }
  if (isExtent) {
    const extentEmu = pxToEmu(actual);
    if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
      throw new CapabilityError(
        "invalid-coordinate",
        `${name} must convert to an EMU value between 1 and ${maximumShapeExtentEmu} inclusive`,
      );
    }
  }
  return actual;
}

function requestedDimension(
  style: Readonly<Record<string, unknown>>,
  props: Readonly<Record<string, unknown>>,
  name: "width" | "height",
): unknown {
  const styled = style[name];
  return styled === undefined
    ? props[name]
    : typeof styled === "number"
      ? styled
      : undefined;
}

function cropValues(
  value: unknown,
  path?: string,
): Record<string, number> | undefined {
  if (value === undefined) return undefined;
  if (!isRecord(value)) {
    throw new CapabilityError(
      "invalid-crop",
      "crop must be a four-sided fraction object",
      path,
    );
  }
  const crop = value;
  const allowed = ["left", "top", "right", "bottom"] as const;
  const unknown = Object.keys(crop).find(
    (key) => !(allowed as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `crop contains unknown parameter ${unknown}`,
      path,
    );
  }
  const result: Record<string, number> = {};
  for (const side of allowed) {
    const fraction = crop[side];
    if (fraction === undefined) continue;
    if (
      typeof fraction !== "number" ||
      !Number.isFinite(fraction) ||
      fraction < 0 ||
      fraction > 1
    ) {
      throw new CapabilityError(
        "invalid-crop",
        `crop.${side} must be a finite fraction between 0 and 1`,
        path,
      );
    }
    result[{ left: "l", top: "t", right: "r", bottom: "b" }[side]] = fraction;
  }
  if ((result["l"] ?? 0) + (result["r"] ?? 0) > 1) {
    throw new CapabilityError(
      "invalid-crop",
      "Sum of crop left and right cannot exceed 1",
      path,
    );
  }
  if ((result["t"] ?? 0) + (result["b"] ?? 0) > 1) {
    throw new CapabilityError(
      "invalid-crop",
      "Sum of crop top and bottom cannot exceed 1",
      path,
    );
  }
  return result;
}

/** Image module reads and validates formal resources, encapsulates bound state, and emits native picture objects. */
export const imageCapability: BoundObjectCapabilityModule = {
  declaration: "rp:image",

  bind(
    input: CapabilityInput,
    context: CapabilityParseContext,
  ): ResolvedBoundObjectCapability {
    const unknown = Object.keys(input.props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Image contains unsupported parameter ${unknown}`,
        context.path,
      );
    }
    if (
      input.children.some(
        (child) => "props" in child || child.text.trim() !== "",
      )
    ) {
      throw new CapabilityError(
        "unsupported-children",
        "Image does not accept children",
        context.path,
      );
    }

    const requestedName = optionalText(input.props, "name", context.path);
    const rawMorphKey = input.props["morphKey"];
    let morphKey: MorphKey | undefined;
    if (rawMorphKey !== undefined) {
      if (!isMorphKey(rawMorphKey)) {
        throw new CapabilityError(
          "invalid-morph-key",
          "morphKey must be created by createMorphScope().handle(name)",
          context.path,
        );
      }
      morphKey = rawMorphKey;
    }
    const link = parseLinkTarget(input.props["link"], context.path);

    const altValue = input.props["alt"];
    if (altValue !== undefined && typeof altValue !== "string") {
      throw new CapabilityError(
        "invalid-parameter",
        "alt must be a string",
        context.path,
      );
    }

    const allowedStyleKeys = new Set<string>([
      ...layoutStyleKeys,
      ...imageVisualStyleKeys,
    ]);

    const contract = parseCapabilityLayoutContract({
      nativeKind: "pic",
      props: input.props,
      allowedStyleKeys,
      allowPadding: false,
      path: context.path,
    });

    const imageStyle = isRecord(input.props["style"])
      ? { ...input.props["style"] }
      : {};

    const requestedFit = input.props["fit"];
    const fit = requestedFit === undefined ? "fill" : requestedFit;
    if (fit !== "fill" && fit !== "cover" && fit !== "contain") {
      throw new CapabilityError(
        "unsupported-fit",
        "fit must be fill, cover, or contain",
        context.path,
      );
    }

    const crop = cropValues(input.props["crop"], context.path);
    const image = readImage(input.props["src"]);
    const geometry = input.props["geometry"] as Geometry | undefined;
    const flipH = input.props["flipH"] as boolean | undefined;
    const flipV = input.props["flipV"] as boolean | undefined;

    // Validate geometry and visual properties during bind before layout runs.
    shapePropertiesXml({
      theme: context.theme,
      style: imageStyle,
      box: { x: 0, y: 0, width: 100, height: 100 },
      ...(geometry !== undefined ? { geometry } : {}),
      ...(flipH !== undefined ? { flipH } : {}),
      ...(flipV !== undefined ? { flipV } : {}),
      image: true,
    });

    const declaredX = input.props["x"];
    const declaredY = input.props["y"];
    const requestedWidth = requestedDimension(imageStyle, input.props, "width");
    const requestedHeight = requestedDimension(
      imageStyle,
      input.props,
      "height",
    );
    const authorId =
      typeof input.props["id"] === "string" ? input.props["id"] : undefined;
    const motionProps: Readonly<Record<string, unknown>> = {
      ...(input.props["enter"] !== undefined
        ? { enter: input.props["enter"] }
        : {}),
      ...(input.props["emphasis"] !== undefined
        ? { emphasis: input.props["emphasis"] }
        : {}),
      ...(input.props["motion"] !== undefined
        ? { motion: input.props["motion"] }
        : {}),
    };

    const native: BoundObjectNativeCollaboration = {
      ...(morphKey === undefined ? {} : { morphKey }),
      ...(requestedName === undefined ? {} : { hasExplicitName: true }),
      ...(link === undefined ? {} : { link }),
    };

    return {
      contract,
      ...(morphKey !== undefined ||
      requestedName !== undefined ||
      link !== undefined
        ? { native }
        : {}),

      measureLayout(
        request: CapabilityLayoutRequest,
      ): CapabilityLayoutMeasurement {
        const declaredWidth =
          typeof requestedWidth === "number" ? requestedWidth : undefined;
        const declaredHeight =
          typeof requestedHeight === "number" ? requestedHeight : undefined;
        const width =
          declaredWidth ??
          Math.min(image.width, request.availableWidth ?? image.width);
        const height = declaredHeight ?? (width * image.height) / image.width;
        return {
          width,
          height,
          minContentWidth: image.width,
          minContentHeight: image.height,
        };
      },

      resolve(resolveContext: ObjectResolveContext): ResolvedObject {
        const x = dimension(resolveContext.layoutBox?.x ?? declaredX, "x", 0);
        const y = dimension(resolveContext.layoutBox?.y ?? declaredY, "y", 0);
        let width: number;
        let height: number;
        if (resolveContext.layoutBox !== undefined) {
          width = dimension(resolveContext.layoutBox.width, "width");
          height = dimension(resolveContext.layoutBox.height, "height");
        } else if (
          requestedWidth === undefined &&
          requestedHeight === undefined
        ) {
          width = image.width;
          height = image.height;
        } else if (requestedWidth === undefined) {
          height = dimension(requestedHeight, "height");
          width = (height * image.width) / image.height;
        } else if (requestedHeight === undefined) {
          width = dimension(requestedWidth, "width");
          height = (width * image.height) / image.width;
        } else {
          width = dimension(requestedWidth, "width");
          height = dimension(requestedHeight, "height");
        }
        width = dimension(width, "width");
        height = dimension(height, "height");

        const placement = imagePlacement(image, fit, crop, {
          x,
          y,
          width,
          height,
        });
        const { box, sourceCrop } = placement;
        const resolvedProperties = shapePropertiesXml({
          theme: resolveContext.theme,
          style: imageStyle,
          box,
          ...(geometry === undefined ? {} : { geometry }),
          ...(flipH === undefined ? {} : { flipH }),
          ...(flipV === undefined ? {} : { flipV }),
          image: true,
        });
        const name =
          resolveContext.morphName ??
          requestedName ??
          `Image ${resolveContext.id - 1}`;
        try {
          assertXml10Characters(name);
          if (altValue !== undefined) assertXml10Characters(altValue);
        } catch {
          throw new CapabilityError(
            "invalid-xml-character",
            "Image.name or alt contains characters not allowed in XML 1.0",
          );
        }
        const angle =
          typeof imageStyle["rotate"] === "number" ? imageStyle["rotate"] : 0;
        const annotations = motionAnnotations(motionProps, {
          kind: "image",
          name,
          ...(authorId === undefined ? {} : { authorId }),
          textBody: false,
          filled: false,
          box,
          rotationDegrees: angle,
          hasBackground:
            imageStyle["borderColor"] !== undefined ||
            (typeof imageStyle["borderWidth"] === "number" &&
              imageStyle["borderWidth"] > 0),
        });

        return {
          observation: {
            id: resolveContext.id,
            kind: "image",
            name,
            box,
          },
          transform: {
            box,
            angle,
            flipH: Boolean(flipH),
            flipV: Boolean(flipV),
          },
          extensions: [],
          ...(resolvedProperties.warnings.length === 0
            ? {}
            : { warnings: resolvedProperties.warnings }),
          ...(annotations.length === 0 ? {} : { annotations }),

          emit(emitContext: EmitContext): EmitOutput {
            const properties = shapePropertiesXml({
              theme: emitContext.theme,
              style: imageStyle,
              box,
              ...(geometry === undefined ? {} : { geometry }),
              ...(flipH === undefined ? {} : { flipH }),
              ...(flipV === undefined ? {} : { flipV }),
              image: true,
            });
            const escapedName = xmlEscape(name);
            const escapedAlt =
              altValue === undefined ? undefined : xmlEscape(altValue);
            let relationshipId: string;
            try {
              relationshipId = emitContext.registerResource({
                bytes: image.bytes,
                extension: image.extension,
                contentType: image.contentType,
                relationshipType:
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
              });
            } catch (error) {
              if (error instanceof CapabilityError) throw error;
              throw new CapabilityError(
                "image-resource-failure",
                error instanceof Error ? error.message : String(error),
              );
            }
            const serializedCrop = Object.fromEntries(
              Object.entries(sourceCrop).filter(
                ([, fraction]) => fraction !== 0,
              ),
            );
            return {
              xml: imageXml(
                emitContext.id,
                escapedName,
                escapedAlt,
                relationshipId,
                box,
                serializedCrop,
                properties.xml,
                properties.opacity,
                nativeLinkXml(emitContext.nativeLink),
              ),
            };
          },

          paint(paintContext: PaintContext): PaintOutput {
            return renderImagePreview({
              image,
              visualStyle: imageStyle,
              geometry,
              flipH,
              flipV,
              placement,
              objectId: paintContext.objectId,
            });
          },
        };
      },
    };
  },
};
