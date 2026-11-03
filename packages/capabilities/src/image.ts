import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  CapabilityError,
  pxToEmu,
  xmlEscape,
  type CapabilityLayoutMeasurement,
  type CapabilityLayoutRequest,
  type CapabilityInput,
  type CapabilityModule,
  type ImageVisualStyle,
  type LayoutStyle,
} from "@react-pptx/core";
import type { ReactNode } from "react";

import type { Geometry } from "./geometry.ts";
import type { MorphProps } from "./morph-api.ts";
import {
  motionAnnotations,
  motionElement,
  type MotionProps,
} from "./motion-api.ts";
import { shapePropertiesXml } from "./visual.ts";

/** 本地文件路径、file URL 或图片原始字节。 */
export type ImageSource = string | URL | Uint8Array;

export type ImageProps = MotionProps &
  MorphProps & {
    readonly testId?: string;
    readonly name?: string;
    readonly src: ImageSource;
    /** 图片的替代说明。 */
    readonly alt?: string;
    /** 位置和范围以 px 表示；位置缺省为 0，宽高缺省按原图尺寸推导。 */
    readonly x?: number;
    readonly y?: number;
    readonly width?: number;
    readonly height?: number;
    readonly geometry?: Geometry;
    readonly flipH?: boolean;
    readonly flipV?: boolean;
    readonly style?: LayoutStyle & ImageVisualStyle;
    /** `fill` 拉伸，`cover` 居中裁切，`contain` 保留比例；缺省为 `fill`。 */
    readonly fit?: "fill" | "cover" | "contain";
    /** 显式裁切每边的原图比例；提供后优先于 `fit`。 */
    readonly crop?: {
      readonly left?: number;
      readonly top?: number;
      readonly right?: number;
      readonly bottom?: number;
    };
    readonly children?: ReactNode;
  };

type ImageMetadata = {
  readonly extension: "png" | "jpeg" | "gif";
  readonly contentType: "image/png" | "image/jpeg" | "image/gif";
  readonly width: number;
  readonly height: number;
};

type LoadedImage = ImageMetadata & { readonly bytes: Uint8Array };

const allowedParameters = [
  "testId",
  "name",
  "morphKey",
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
  "__motionContext",
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

function readImage(source: unknown): LoadedImage {
  if (
    typeof source !== "string" &&
    !(source instanceof URL) &&
    !(source instanceof Uint8Array)
  ) {
    throw new CapabilityError(
      "invalid-image-source",
      "Image.src 必须是文件路径、file URL 或图片字节",
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
    throw new CapabilityError("image-read-failed", "图片资源无法读取");
  }
  const metadata =
    pngMetadata(bytes) ?? jpegMetadata(bytes) ?? gifMetadata(bytes);
  if (metadata === undefined) {
    throw new CapabilityError(
      "unsupported-image-format",
      "图片资源内容无效或格式不受支持；仅支持 PNG、JPEG 和 GIF",
    );
  }
  return { ...metadata, bytes };
}

function optionalText(
  props: Record<string, unknown>,
  name: "testId" | "name",
): string | undefined {
  const value = props[name];
  if (value === undefined) return undefined;
  if (typeof value !== "string" || value.trim() === "") {
    throw new CapabilityError("invalid-parameter", `${name} 必须是非空字符串`);
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
      `${name} 必须是有限像素数值，宽高必须大于 0，坐标范围为 -1000000 到 1000000`,
    );
  }
  if (isExtent) {
    const extentEmu = pxToEmu(actual);
    if (extentEmu < 1 || extentEmu > maximumShapeExtentEmu) {
      throw new CapabilityError(
        "invalid-coordinate",
        `${name} 换算后必须在 1 到 ${maximumShapeExtentEmu} EMU 的闭区间内`,
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

function cropValues(value: unknown): Record<string, number> | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new CapabilityError("invalid-crop", "crop 必须是四边分数对象");
  }
  const crop = value as Record<string, unknown>;
  const allowed = ["left", "top", "right", "bottom"] as const;
  const unknown = Object.keys(crop).find(
    (key) => !(allowed as readonly string[]).includes(key),
  );
  if (unknown !== undefined) {
    throw new CapabilityError(
      "unknown-parameter",
      `crop 含未知参数 ${unknown}`,
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
        `crop.${side} 必须是 0 到 1 之间的有限分数`,
      );
    }
    result[{ left: "l", top: "t", right: "r", bottom: "b" }[side]!] = fraction;
  }
  if ((result["l"] ?? 0) + (result["r"] ?? 0) > 1) {
    throw new CapabilityError("invalid-crop", "crop 左右边之和不能超过 1");
  }
  if ((result["t"] ?? 0) + (result["b"] ?? 0) > 1) {
    throw new CapabilityError("invalid-crop", "crop 上下边之和不能超过 1");
  }
  return result;
}

function imageXml(
  id: number,
  escapedName: string,
  escapedAlt: string | undefined,
  relationshipId: string,
  box: { x: number; y: number; width: number; height: number },
  crop: Record<string, number>,
  properties: string,
  opacity: number | undefined,
): string {
  const description = escapedAlt === undefined ? "" : ` descr="${escapedAlt}"`;
  const cropXml =
    Object.keys(crop).length === 0
      ? ""
      : `<a:srcRect${Object.entries(crop)
          .map(
            ([side, fraction]) => ` ${side}="${Math.round(fraction * 100000)}"`,
          )
          .join("")}/>`;
  const alpha =
    opacity === undefined
      ? ""
      : `<a:alphaModFix amt="${Math.round(opacity * 100000)}"/>`;
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapedName}"${description}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}">${alpha}</a:blip>${cropXml}<a:stretch><a:fillRect/></a:stretch></p:blipFill>${properties}</p:pic>`;
}

/** 图片模块读取并验证正式资源，再输出带媒体关系的原生 picture 对象。 */
export const imageCapability: CapabilityModule = {
  declaration: "rp:image",
  measureLayout(
    input: CapabilityInput,
    request: CapabilityLayoutRequest,
  ): CapabilityLayoutMeasurement {
    const props = input.props;
    const image = readImage(props["src"]);
    const style =
      typeof props["style"] === "object" && props["style"] !== null
        ? (props["style"] as Record<string, unknown>)
        : {};
    const requestedWidth = requestedDimension(style, props, "width");
    const requestedHeight = requestedDimension(style, props, "height");
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
  compile(input: CapabilityInput, context) {
    const props = input.props;
    const unknown = Object.keys(props).find(
      (key) => !(allowedParameters as readonly string[]).includes(key),
    );
    if (unknown !== undefined) {
      throw new CapabilityError(
        "unknown-parameter",
        `Image 含未实现参数 ${unknown}`,
      );
    }
    if (
      input.children.some(
        (child) => "props" in child || child.text.trim() !== "",
      )
    ) {
      throw new CapabilityError("unsupported-children", "Image 不接受子内容");
    }

    const testId = optionalText(props, "testId");
    const requestedName = optionalText(props, "name");
    const altValue = props["alt"];
    if (altValue !== undefined && typeof altValue !== "string") {
      throw new CapabilityError("invalid-parameter", "alt 必须是字符串");
    }
    const image = readImage(props["src"]);
    const imageStyle =
      typeof props["style"] === "object" && props["style"] !== null
        ? (props["style"] as Record<string, unknown>)
        : {};
    const requestedWidth = requestedDimension(imageStyle, props, "width");
    const requestedHeight = requestedDimension(imageStyle, props, "height");
    const x = dimension(context.layoutBox?.x ?? props["x"], "x", 0);
    const y = dimension(context.layoutBox?.y ?? props["y"], "y", 0);
    let width: number;
    let height: number;
    if (context.layoutBox !== undefined) {
      width = dimension(context.layoutBox.width, "width");
      height = dimension(context.layoutBox.height, "height");
    } else if (requestedWidth === undefined && requestedHeight === undefined) {
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

    const requestedFit = props["fit"];
    const fit = requestedFit === undefined ? "fill" : requestedFit;
    if (fit !== "fill" && fit !== "cover" && fit !== "contain") {
      throw new CapabilityError(
        "unsupported-fit",
        "fit 必须是 fill、cover 或 contain",
      );
    }
    const crop = cropValues(props["crop"]);
    const box = { x, y, width, height };
    const sourceCrop = crop ?? {};
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
    // contain 派生的尺寸也必须能写成非零 EMU。
    box.width = dimension(box.width, "width");
    box.height = dimension(box.height, "height");
    const properties = shapePropertiesXml({
      theme: context.theme,
      style: imageStyle,
      box,
      ...(props["geometry"] === undefined
        ? {}
        : { geometry: props["geometry"] as Geometry }),
      ...(props["flipH"] === undefined
        ? {}
        : { flipH: props["flipH"] as boolean }),
      ...(props["flipV"] === undefined
        ? {}
        : { flipV: props["flipV"] as boolean }),
      image: true,
    });
    const name =
      context.morphName ?? requestedName ?? testId ?? `Image ${context.id - 1}`;
    const annotations = motionAnnotations(props, {
      kind: "image",
      name,
      ...(testId === undefined ? {} : { testId }),
      textBody: false,
      filled: false,
    });
    let escapedName: string;
    let escapedAlt: string | undefined;
    try {
      escapedName = xmlEscape(name);
      escapedAlt = altValue === undefined ? undefined : xmlEscape(altValue);
    } catch {
      throw new CapabilityError(
        "invalid-xml-character",
        "Image.name 或 alt 含 XML 1.0 不允许的字符",
      );
    }
    let relationshipId: string;
    try {
      relationshipId = context.registerResource({
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
      Object.entries(sourceCrop).filter(([, fraction]) => fraction !== 0),
    );
    return {
      xml: imageXml(
        context.id,
        escapedName,
        escapedAlt,
        relationshipId,
        box,
        serializedCrop,
        properties.xml,
        properties.opacity,
      ),
      observation: {
        id: context.id,
        kind: "image",
        name,
        ...(testId === undefined ? {} : { testId }),
        box,
      },
      extensions: [],
      ...(properties.warnings.length === 0
        ? {}
        : { warnings: properties.warnings }),
      ...(annotations.length === 0 ? {} : { annotations }),
    };
  },
};

/** 在 React 根树中声明一张由官方图片模块处理的原生图片。 */
export function Image(props: ImageProps): ReactNode {
  return motionElement("rp:image", props);
}
