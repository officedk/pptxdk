import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

import { CapabilityError } from "@pptxdk/core/capability";

import { decodeGifFirstFrame } from "./gif-decoder.ts";

export type ImageFormatExtension = "png" | "jpeg" | "gif";

export type ImageContentType = "image/png" | "image/jpeg" | "image/gif";

export type ImageMetadata = {
  readonly extension: ImageFormatExtension;
  readonly contentType: ImageContentType;
  readonly width: number;
  readonly height: number;
};

export type LoadedImage = ImageMetadata & { readonly bytes: Uint8Array };

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

/** Reads and validates PNG, JPEG, or GIF header metadata from a file path, URL, or byte buffer. */
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

/**
 * Decodes PNG, JPEG, or static first-frame GIF bytes into normalized PNG bytes for SVG embedding.
 * Throws `preview-corrupt-image` if compressed pixel data is invalid.
 */
export function decodeImageToPngBytes(image: LoadedImage): Uint8Array {
  const { png: pngModule, jpeg: jpegModule } = imageDecoders();
  let decoded: {
    readonly width: number;
    readonly height: number;
    readonly data: Uint8Array;
  };
  try {
    switch (image.extension) {
      case "png":
        decoded = pngModule.PNG.sync.read(Buffer.from(image.bytes));
        break;
      case "jpeg":
        decoded = jpegModule.decode(image.bytes, {
          tolerantDecoding: false,
        });
        break;
      case "gif":
        decoded = decodeGifFirstFrame(image.bytes);
        break;
      default: {
        const exhaustive: never = image.extension;
        return exhaustive;
      }
    }
  } catch (error) {
    if (error instanceof CapabilityError) {
      throw error;
    }
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
  return new Uint8Array(pngModule.PNG.sync.write(normalized));
}
