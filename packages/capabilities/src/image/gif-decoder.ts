import { CapabilityError } from "@pptxdk/core/capability";

export type DecodedRgbaImage = {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
};

function corruptGif(message: string): never {
  throw new CapabilityError("preview-corrupt-image", message);
}

function u16LE(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] ?? 0) + ((bytes[offset + 1] ?? 0) << 8);
}

function skipSubBlocks(bytes: Uint8Array, startOffset: number): number {
  let offset = startOffset;
  while (offset < bytes.length) {
    const size = bytes[offset] ?? 0;
    offset += 1;
    if (size === 0) {
      return offset;
    }
    offset += size;
    if (offset > bytes.length) {
      corruptGif("Failed to decode GIF image: truncated extension sub-block");
    }
  }
  corruptGif("Failed to decode GIF image: unterminated extension sub-block");
}

function collectSubBlocks(
  bytes: Uint8Array,
  startOffset: number,
): { readonly data: Uint8Array; readonly nextOffset: number } {
  const chunks: Uint8Array[] = [];
  let totalLength = 0;
  let offset = startOffset;
  while (offset < bytes.length) {
    const size = bytes[offset] ?? 0;
    offset += 1;
    if (size === 0) {
      const merged = new Uint8Array(totalLength);
      let cursor = 0;
      for (const chunk of chunks) {
        merged.set(chunk, cursor);
        cursor += chunk.length;
      }
      return { data: merged, nextOffset: offset };
    }
    if (offset + size > bytes.length) {
      corruptGif("Failed to decode GIF image: truncated image data sub-block");
    }
    chunks.push(bytes.subarray(offset, offset + size));
    totalLength += size;
    offset += size;
  }
  corruptGif("Failed to decode GIF image: unterminated image data sub-block");
}

function decompressLzw(
  compressed: Uint8Array,
  lzwMinCodeSize: number,
  expectedPixels: number,
): Uint8Array {
  if (lzwMinCodeSize < 2 || lzwMinCodeSize > 11) {
    corruptGif(
      `Failed to decode GIF image: invalid LZW minimum code size ${lzwMinCodeSize}`,
    );
  }
  if (compressed.length === 0) {
    corruptGif("Failed to decode GIF image: empty LZW compressed stream");
  }

  const clearCode = 1 << lzwMinCodeSize;
  const eoiCode = clearCode + 1;
  const indices = new Uint8Array(expectedPixels);
  const prefix = new Int32Array(4096);
  const suffix = new Uint8Array(4096);
  const pixelStack = new Uint8Array(4096);

  for (let i = 0; i < clearCode; i += 1) {
    prefix[i] = -1;
    suffix[i] = i;
  }

  let codeSize = lzwMinCodeSize + 1;
  let nextCode = eoiCode + 1;
  let oldCode = -1;
  let firstChar = 0;
  let pixelOut = 0;

  let bytePos = 0;
  let bitBuffer = 0;
  let bitCount = 0;

  const readCode = (): number => {
    while (bitCount < codeSize) {
      if (bytePos >= compressed.length) {
        corruptGif("Failed to decode GIF image: truncated LZW bit stream");
      }
      bitBuffer |= (compressed[bytePos] ?? 0) << bitCount;
      bytePos += 1;
      bitCount += 8;
    }
    const mask = (1 << codeSize) - 1;
    const code = bitBuffer & mask;
    bitBuffer >>>= codeSize;
    bitCount -= codeSize;
    return code;
  };

  while (pixelOut < expectedPixels) {
    const rawCode = readCode();
    if (rawCode === clearCode) {
      codeSize = lzwMinCodeSize + 1;
      nextCode = eoiCode + 1;
      oldCode = -1;
      continue;
    }
    if (rawCode === eoiCode) {
      break;
    }
    if (rawCode > nextCode || (rawCode === nextCode && oldCode === -1)) {
      corruptGif(
        `Failed to decode GIF image: out-of-range LZW code ${rawCode} (nextCode=${nextCode})`,
      );
    }

    let code = rawCode;
    let stackTop = 0;
    if (code === nextCode) {
      pixelStack[stackTop] = firstChar;
      stackTop += 1;
      code = oldCode;
    }
    while (code >= clearCode) {
      if (stackTop >= 4096) {
        corruptGif("Failed to decode GIF image: corrupt LZW dictionary cycle");
      }
      pixelStack[stackTop] = suffix[code] ?? 0;
      stackTop += 1;
      code = prefix[code] ?? -1;
      if (code < 0) {
        corruptGif("Failed to decode GIF image: invalid LZW prefix reference");
      }
    }
    firstChar = suffix[code] ?? 0;
    pixelStack[stackTop] = firstChar;
    stackTop += 1;

    for (let i = stackTop - 1; i >= 0 && pixelOut < expectedPixels; i -= 1) {
      indices[pixelOut] = pixelStack[i] ?? 0;
      pixelOut += 1;
    }

    if (oldCode !== -1 && nextCode < 4096) {
      prefix[nextCode] = oldCode;
      suffix[nextCode] = firstChar;
      nextCode += 1;
      if (nextCode === 1 << codeSize && codeSize < 12) {
        codeSize += 1;
      }
    }
    oldCode = rawCode;
  }

  if (pixelOut < expectedPixels) {
    corruptGif(
      `Failed to decode GIF image: truncated pixel stream (${pixelOut}/${expectedPixels} pixels)`,
    );
  }
  return indices;
}

const interlacePasses: readonly {
  readonly start: number;
  readonly step: number;
}[] = [
  { start: 0, step: 8 },
  { start: 4, step: 8 },
  { start: 2, step: 4 },
  { start: 1, step: 2 },
];

/**
 * Decodes the static first frame of a GIF87a or GIF89a byte stream into RGBA pixels.
 * Supports GCT/LCT palettes, Graphic Control Extension transparency, and 4-pass interlacing.
 */
export function decodeGifFirstFrame(bytes: Uint8Array): DecodedRgbaImage {
  if (bytes.length < 14) {
    corruptGif("Failed to decode GIF image: truncated header");
  }
  const signature = String.fromCharCode(...bytes.subarray(0, 6));
  if (signature !== "GIF87a" && signature !== "GIF89a") {
    corruptGif("Failed to decode GIF image: invalid GIF signature");
  }

  const screenW = u16LE(bytes, 6);
  const screenH = u16LE(bytes, 8);
  if (screenW < 1 || screenH < 1) {
    corruptGif("Failed to decode GIF image: invalid logical screen dimensions");
  }

  const packed = bytes[10] ?? 0;
  const hasGct = (packed & 0x80) !== 0;
  const gctSize = 1 << ((packed & 0x07) + 1);
  let offset = 13;
  let gct: Uint8Array | undefined;
  if (hasGct) {
    const gctEnd = offset + gctSize * 3;
    if (gctEnd > bytes.length) {
      corruptGif("Failed to decode GIF image: truncated global color table");
    }
    gct = bytes.subarray(offset, gctEnd);
    offset = gctEnd;
  }

  let transparentIndex: number | undefined;
  let foundImage = false;
  while (offset < bytes.length) {
    const block = bytes[offset];
    offset += 1;
    if (block === 0x2c) {
      foundImage = true;
      break;
    }
    if (block === 0x3b) {
      corruptGif("Failed to decode GIF image: no image frame before trailer");
    }
    if (block === 0x21) {
      if (offset >= bytes.length) {
        corruptGif("Failed to decode GIF image: truncated extension header");
      }
      const label = bytes[offset] ?? 0;
      offset += 1;
      if (label === 0xf9) {
        const subSize = bytes[offset] ?? 0;
        if (subSize !== 4 || offset + 5 >= bytes.length) {
          corruptGif(
            "Failed to decode GIF image: invalid graphic control extension",
          );
        }
        const gcePacked = bytes[offset + 1] ?? 0;
        const transparentFlag = (gcePacked & 0x01) !== 0;
        const transIdx = bytes[offset + 4] ?? 0;
        transparentIndex = transparentFlag ? transIdx : undefined;
        offset = skipSubBlocks(bytes, offset);
      } else {
        offset = skipSubBlocks(bytes, offset);
      }
      continue;
    }
    corruptGif(
      `Failed to decode GIF image: unexpected block introducer 0x${(block ?? 0).toString(16)}`,
    );
  }

  if (!foundImage || offset + 9 > bytes.length) {
    corruptGif("Failed to decode GIF image: missing image descriptor");
  }

  const left = u16LE(bytes, offset);
  const top = u16LE(bytes, offset + 2);
  const frameW = u16LE(bytes, offset + 4);
  const frameH = u16LE(bytes, offset + 6);
  const imgPacked = bytes[offset + 8] ?? 0;
  offset += 9;

  if (frameW < 1 || frameH < 1) {
    corruptGif("Failed to decode GIF image: invalid frame dimensions");
  }

  const hasLct = (imgPacked & 0x80) !== 0;
  const interlaced = (imgPacked & 0x40) !== 0;
  const lctSize = 1 << ((imgPacked & 0x07) + 1);
  let palette = gct;
  if (hasLct) {
    const lctEnd = offset + lctSize * 3;
    if (lctEnd > bytes.length) {
      corruptGif("Failed to decode GIF image: truncated local color table");
    }
    palette = bytes.subarray(offset, lctEnd);
    offset = lctEnd;
  }
  if (palette === undefined || palette.length < 3) {
    corruptGif("Failed to decode GIF image: missing color table");
  }

  if (offset >= bytes.length) {
    corruptGif("Failed to decode GIF image: missing LZW minimum code size");
  }
  const lzwMinCodeSize = bytes[offset] ?? 0;
  offset += 1;

  const { data: compressed } = collectSubBlocks(bytes, offset);
  const indices = decompressLzw(compressed, lzwMinCodeSize, frameW * frameH);

  const rgba = new Uint8Array(screenW * screenH * 4);
  const rowMap = new Int32Array(frameH);
  if (interlaced) {
    let srcY = 0;
    for (const pass of interlacePasses) {
      for (let destY = pass.start; destY < frameH; destY += pass.step) {
        rowMap[srcY] = destY;
        srcY += 1;
      }
    }
  } else {
    for (let y = 0; y < frameH; y += 1) {
      rowMap[y] = y;
    }
  }

  for (let srcY = 0; srcY < frameH; srcY += 1) {
    const destY = top + (rowMap[srcY] ?? 0);
    for (let x = 0; x < frameW; x += 1) {
      const colorIdx = indices[srcY * frameW + x] ?? 0;
      if (colorIdx === transparentIndex) continue;
      const palOffset = colorIdx * 3;
      if (palOffset + 2 >= palette.length) {
        corruptGif(
          `Failed to decode GIF image: palette index ${colorIdx} out of bounds`,
        );
      }
      const destX = left + x;
      if (destX >= 0 && destX < screenW && destY >= 0 && destY < screenH) {
        const outOffset = (destY * screenW + destX) * 4;
        rgba[outOffset] = palette[palOffset] ?? 0;
        rgba[outOffset + 1] = palette[palOffset + 1] ?? 0;
        rgba[outOffset + 2] = palette[palOffset + 2] ?? 0;
        rgba[outOffset + 3] = 255;
      }
    }
  }

  return { width: screenW, height: screenH, data: rgba };
}
