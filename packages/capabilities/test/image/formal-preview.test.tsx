import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, image, slide } from "@pptxdk/capabilities/document";
import { Image } from "@pptxdk/capabilities/image-elements";
import { preset } from "@pptxdk/capabilities/shape";
import { type CapabilityModule, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

function createTwoToneSource(): PNG {
  // Left half (x < 10) is pure red, right half (x >= 10) is pure blue.
  const source = new PNG({ width: 20, height: 10 });
  for (let y = 0; y < 10; y++) {
    for (let x = 0; x < 20; x++) {
      source.data.set(
        x < 10 ? [255, 0, 0, 255] : [0, 0, 255, 255],
        (y * 20 + x) * 4,
      );
    }
  }
  return source;
}

describe("Image static preview positive behaviors", () => {
  it("renders PNG and JPEG with shared imagePlacement (contain, cover, explicit crop), transforms, opacity, and screenGeometry", async () => {
    const { encode } = await import("jpeg-js");
    const source = createTwoToneSource();
    const pngBytes = new Uint8Array(PNG.sync.write(source));
    const jpegBytes = new Uint8Array(encode(source, 100).data);

    const previewDoc = runDocPreview(documentCapabilities);

    for (const [format, bytes] of [
      ["png", pngBytes],
      ["jpeg", jpegBytes],
    ] as const) {
      const mutable = image({
        id: "cropped",
        src: bytes,
        x: 10,
        y: 10,
        width: 40,
        height: 40,
        crop: { left: 0.5 },
      });
      const doc = deck({ width: 320, height: 180 }).append(
        slide({ id: `img-${format}`, background: "#FFFFFF" }).append(
          mutable,
          image({
            id: "contained",
            src: bytes,
            x: 80,
            y: 10,
            width: 80,
            height: 80,
            fit: "contain",
            style: { rotate: 90 },
          }),
          image({
            id: "covered",
            src: bytes,
            x: 180,
            y: 10,
            width: 40,
            height: 80,
            fit: "cover",
            flipH: true,
            style: { opacity: 0.5 },
          }),
        ),
      );

      const first = await previewDoc(doc);
      const page = first.pages[0]!;
      assert.deepEqual(page.diagnostics, [], format);
      assert.ok(page.png !== undefined, format);
      assert.ok(page.svg !== undefined, format);
      expect(page.svg).toContain("<clipPath");
      expect(page.svg).toContain("<image");

      // Explicit crop: left: 0.5 removes red half, center of (10..50, 10..50) = (30, 30) must be blue.
      const rendered = PNG.sync.read(Buffer.from(page.png));
      const cropCenter = (30 * rendered.width + 30) * 4;
      assert.ok((rendered.data[cropCenter] ?? 255) < 15, format);
      assert.ok((rendered.data[cropCenter + 2] ?? 0) > 240, format);

      // Contain 20x10 inside 80x80 at (80, 10) -> box { x: 80, y: 30, width: 80, height: 40 }.
      // Rotated 90deg -> screenBox width ~ 40, height ~ 80.
      const containedNode = page.nodes.find((n) => n.authorId === "contained")!;
      assert.deepEqual(containedNode.box, {
        x: 80,
        y: 30,
        width: 80,
        height: 40,
      });
      expect(containedNode.screenBox?.width).toBeCloseTo(40, 1);
      expect(containedNode.screenBox?.height).toBeCloseTo(80, 1);

      // Mutate image after preview and verify earlier build & preview remain untouched.
      const prevPng = page.png.slice();
      const prevSvg = page.svg;
      const prevPptx = first.pptx.slice();
      mutable.set({ crop: { right: 0.5 } });
      const second = await previewDoc(doc);
      assert.deepEqual(first.pages[0]?.png, prevPng);
      assert.equal(first.pages[0]?.svg, prevSvg);
      assert.deepEqual(first.pptx, prevPptx);
      assert.notEqual(second.pages[0]?.svg, prevSvg);
      const updatedRendered = PNG.sync.read(Buffer.from(second.pages[0]!.png!));
      assert.ok((updatedRendered.data[cropCenter] ?? 0) > 240, format);
      assert.ok((updatedRendered.data[cropCenter + 2] ?? 255) < 15, format);
    }

    // JSX entry produces valid image preview via same formal pipeline.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "jsx-img" },
          createElement(Image, {
            id: "photo",
            src: pngBytes,
            x: 20,
            y: 20,
            width: 60,
            height: 60,
            fit: "contain",
          }),
        ),
      ),
    );
    const jsxPreview = await previewDoc(jsxDoc);
    assert.deepEqual(jsxPreview.pages[0]?.diagnostics, []);
    assert.deepEqual(
      jsxPreview.pages[0]?.nodes.find((n) => n.authorId === "photo")?.box,
      { x: 20, y: 35, width: 60, height: 30 },
    );
  });

  it("renders GIF87a/89a first frame including GCT, LCT, GCE transparency, and interlacing", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    // Independent known-good 1x1 white GIF89a with GCT.
    const gif1x1White = new Uint8Array(
      Buffer.from(
        "47494638396101000100800000ffffff0000002c00000000010001000002024401003b",
        "hex",
      ),
    );

    // Build an 8x8 interlaced GIF89a with LCT and GCE transparent index 0:
    // Pass 0 (y=0), Pass 1 (y=4), Pass 2 (y=2,6), Pass 3 (y=1,3,5,7).
    // Row 0 = color 1 (red #FF0000), Row 1 = color 0 (transparent), Rows 2..7 = color 2 (blue #0000FF).
    const lzwEncodeMin2 = (pixels: readonly number[]): Uint8Array => {
      const bytes: number[] = [];
      let bitBuf = 0;
      let bitCount = 0;
      const pushCode = (code: number, width: number) => {
        bitBuf |= code << bitCount;
        bitCount += width;
        while (bitCount >= 8) {
          bytes.push(bitBuf & 0xff);
          bitBuf >>>= 8;
          bitCount -= 8;
        }
      };
      // Reset dictionary every few pixels to keep 3-bit codes simple and deterministic.
      pushCode(4, 3); // clear
      for (const p of pixels) {
        pushCode(p, 3);
        pushCode(4, 3); // clear so code size remains 3 bits
      }
      pushCode(5, 3); // EOI
      if (bitCount > 0) bytes.push(bitBuf & 0xff);
      return new Uint8Array(bytes);
    };

    const interlaceRows = [0, 4, 2, 6, 1, 3, 5, 7];
    const rowColor = (y: number) => (y === 0 ? 1 : y === 1 ? 0 : 2);
    const interlacedPixels: number[] = [];
    for (const y of interlaceRows) {
      for (let x = 0; x < 8; x += 1) {
        interlacedPixels.push(rowColor(y));
      }
    }
    const compressedInterlaced = lzwEncodeMin2(interlacedPixels);
    const gifInterlaced = new Uint8Array([
      // Header GIF89a + LSD 8x8 (no GCT: packed = 0x00)
      0x47,
      0x49,
      0x46,
      0x38,
      0x39,
      0x61,
      8,
      0,
      8,
      0,
      0x00,
      0,
      0,
      // GCE: transparent flag=1, transparent index=0
      0x21,
      0xf9,
      0x04,
      0x01,
      0x00,
      0x00,
      0x00,
      0x00,
      // Image descriptor 8x8, LCT=1 (4 colors -> 0x81), Interlace=1 (0x40) => packed = 0xC1
      0x2c,
      0,
      0,
      0,
      0,
      8,
      0,
      8,
      0,
      0xc1,
      // LCT (4 RGB entries)
      0,
      0,
      0, // 0: transparent
      255,
      0,
      0, // 1: red
      0,
      0,
      255, // 2: blue
      0,
      255,
      0, // 3: green
      // LZW min code size = 2
      0x02,
      compressedInterlaced.length,
      ...compressedInterlaced,
      0x00, // sub-block terminator
      0x3b, // trailer
    ]);

    // GIF87a 2x2 solid green (#00AA00) with GCT.
    const compressed87a = lzwEncodeMin2([1, 1, 1, 1]);
    const gif87a = new Uint8Array([
      0x47,
      0x49,
      0x46,
      0x38,
      0x37,
      0x61,
      2,
      0,
      2,
      0,
      0x81,
      0,
      0,
      0,
      0,
      0,
      0,
      170,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0x2c,
      0,
      0,
      0,
      0,
      2,
      0,
      2,
      0,
      0x00,
      0x02,
      compressed87a.length,
      ...compressed87a,
      0x00,
      0x3b,
    ]);

    // GIF89a with a Comment Extension block (0x21 0xFE) and non-transparent GCE (transparentFlag = 0).
    const gifWithCommentAndOpaqueGce = new Uint8Array([
      0x47,
      0x49,
      0x46,
      0x38,
      0x39,
      0x61,
      2,
      0,
      2,
      0,
      0x81,
      0,
      0,
      0,
      0,
      255, // color 0: blue
      255,
      255,
      0, // color 1: yellow
      0,
      0,
      0,
      0,
      0,
      0,
      // Comment extension: 'TEST'
      0x21,
      0xfe,
      0x04,
      0x54,
      0x45,
      0x53,
      0x54,
      0x00,
      // GCE with transparentFlag = 0 (packed = 0x00)
      0x21,
      0xf9,
      0x04,
      0x00,
      0x00,
      0x00,
      0x00,
      0x00,
      // Image descriptor 2x2
      0x2c,
      0,
      0,
      0,
      0,
      2,
      0,
      2,
      0,
      0x00,
      0x02,
      compressed87a.length,
      ...compressed87a,
      0x00,
      0x3b,
    ]);

    const doc = deck({ width: 240, height: 120 }).append(
      slide({ id: "gifs", background: "#10B981" }).append(
        image({
          id: "gif-white",
          src: gif1x1White,
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
        image({
          id: "gif-interlaced",
          src: gifInterlaced,
          x: 60,
          y: 10,
          width: 80,
          height: 80,
        }),
        image({
          id: "gif-87a",
          src: gif87a,
          x: 160,
          y: 10,
          width: 40,
          height: 40,
        }),
        image({
          id: "gif-comment-opaque",
          src: gifWithCommentAndOpaqueGce,
          x: 160,
          y: 60,
          width: 40,
          height: 40,
        }),
      ),
    );
    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    const rendered = PNG.sync.read(Buffer.from(page.png));
    const px = (x: number, y: number) => [
      ...rendered.data.subarray(
        (y * rendered.width + x) * 4,
        (y * rendered.width + x) * 4 + 4,
      ),
    ];
    // 1x1 white GIF at (10..50, 10..50)
    assert.deepEqual(px(30, 30), [255, 255, 255, 255]);
    // Interlaced 8x8 scaled to 80x80 at (60, 10):
    // Row 0 -> y in 10..19 is red (#FF0000)
    const row0 = px(100, 11);
    assert.ok(row0[0]! > 240 && row0[1]! < 20 && row0[2]! < 20);
    // Row 1 -> y in 20..29 is transparent, revealing slide background #10B981
    const row1 = px(100, 25);
    assert.ok(row1[0]! < 80 && row1[1]! > 150 && row1[2]! > 100);
    // Row 2 -> y in 30..39 is blue (#0000FF)
    const row2 = px(100, 35);
    assert.ok(row2[0]! < 40 && row2[1]! < 60 && row2[2]! > 220);
    // GIF87a green (#00AA00) at (160..200, 10..50)
    assert.deepEqual(px(180, 30), [0, 170, 0, 255]);
    // GIF89a with comment extension & opaque GCE yellow (#FFFF00) at (160..200, 60..100)
    assert.deepEqual(px(180, 80), [255, 255, 0, 255]);
  });

  it("renders preset and custom geometry masks, dashed/dotted borders, boxShadow, softEdge, and opacity without double-applying border alpha", async () => {
    const { encode } = await import("jpeg-js");
    const source = createTwoToneSource();
    const pngBytes = new Uint8Array(PNG.sync.write(source));
    const jpegBytes = new Uint8Array(encode(source, 100).data);

    // 1x1 solid red GIF89a.
    const gifRedBytes = new Uint8Array(
      Buffer.from(
        "47494638396101000100800000ff00000000002c00000000010001000002024401003b",
        "hex",
      ),
    );

    const previewDoc = runDocPreview(documentCapabilities);

    // Separate fill=false stroke=true path should be stroked but NOT added to <clipPath>.
    const multiSubpathMask = [
      {
        commands: [
          { op: "M" as const, x: 0, y: 0 },
          { op: "L" as const, x: 10, y: 0 },
          { op: "L" as const, x: 5, y: 10 },
          { op: "Z" as const },
        ],
        fill: true,
        stroke: true,
      },
      {
        commands: [
          { op: "M" as const, x: 1, y: 9 },
          { op: "L" as const, x: 9, y: 9 },
        ],
        fill: false,
        stroke: true,
      },
    ];

    const doc = deck({ width: 480, height: 240 }).append(
      slide({ id: "image-masks-and-effects", background: "#FFFFFF" }).append(
        // 1. Ellipse mask + dashed border + boxShadow on PNG
        image({
          id: "ellipse-dashed-shadow",
          src: pngBytes,
          geometry: "ellipse",
          x: 20,
          y: 20,
          width: 100,
          height: 80,
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "dashed",
            boxShadow: { color: "#000000", blur: 6, distance: 6, angle: 45 },
          },
        }),
        // 2. Adjusted roundRect mask + dotted border on JPEG
        image({
          id: "roundrect-dotted",
          src: jpegBytes,
          geometry: preset("roundRect", { adj: 25000 }),
          x: 150,
          y: 20,
          width: 100,
          height: 80,
          style: {
            borderColor: "#DC2626",
            borderWidth: 2,
            borderStyle: "dotted",
          },
        }),
        // 3. SoftEdge + opacity on GIF
        image({
          id: "gif-softedge-opacity",
          src: gifRedBytes,
          geometry: "diamond",
          x: 280,
          y: 20,
          width: 90,
          height: 80,
          style: {
            softEdge: 8,
            opacity: 0.6,
          },
        }),
        // 4. Custom multi-path mask (one fill=false stroke=true subpath) + flipH + rotate
        image({
          id: "custom-mask-transformed",
          src: pngBytes,
          geometry: {
            type: "custom",
            width: 10,
            height: 10,
            paths: multiSubpathMask,
          },
          x: 20,
          y: 130,
          width: 100,
          height: 80,
          flipH: true,
          style: {
            rotate: 15,
            borderColor: "#0F766E",
            borderWidth: 2,
            borderStyle: "solid",
          },
        }),
        // 5. SmileyFace preset (contains fill="none" mouth arc) with solid border
        image({
          id: "smiley-mask",
          src: pngBytes,
          geometry: "smileyFace",
          x: 160,
          y: 130,
          width: 90,
          height: 80,
          style: {
            borderColor: "#0F172A",
            borderWidth: 2,
          },
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain("<clipPath");
    expect(page.svg).toContain("stroke-dasharray=");
    expect(page.svg).toContain("<feDropShadow");
    expect(page.svg).toContain("<mask");

    const rendered = PNG.sync.read(Buffer.from(page.png));
    const px = (x: number, y: number) => [
      ...rendered.data.subarray(
        (y * rendered.width + x) * 4,
        (y * rendered.width + x) * 4 + 4,
      ),
    ];
    // Corner (22, 22) outside ellipse mask must stay white (not filled by image).
    const outsideEllipse = px(22, 22);
    assert.ok(
      outsideEllipse[0]! > 240 &&
        outsideEllipse[1]! > 240 &&
        outsideEllipse[2]! > 240,
    );
    // Center of left half of ellipse (45, 60) must be red from two-tone PNG.
    const insideEllipseLeft = px(45, 60);
    assert.ok(insideEllipseLeft[0]! > 220 && insideEllipseLeft[2]! < 30);
    // Center of right half of ellipse (95, 60) must be blue.
    const insideEllipseRight = px(95, 60);
    assert.ok(insideEllipseRight[2]! > 220 && insideEllipseRight[0]! < 30);
    // GIF diamond with opacity 0.6 over white at center (325, 60) must be blended pink-red (R=255, G/B around 102).
    const gifCenter = px(325, 60);
    assert.ok(gifCenter[0]! > 240 && gifCenter[1]! > 70 && gifCenter[1]! < 150);
  });

  it("rejects only corrupt PNG/JPEG/GIF pixel streams with preview-corrupt-image while preserving valid sibling pages", async () => {
    const { encode } = await import("jpeg-js");
    const previewDoc = runDocPreview(documentCapabilities);

    const validSource = createTwoToneSource();
    const validPng = new Uint8Array(PNG.sync.write(validSource));
    const validJpeg = new Uint8Array(encode(validSource, 90).data);

    // Valid PNG header/chunks with corrupted IDAT compressed data.
    const corruptPng = new Uint8Array(validPng);
    const idatOffset = Buffer.from(corruptPng).indexOf(Buffer.from("IDAT"));
    corruptPng[idatOffset + 6] = (corruptPng[idatOffset + 6] ?? 0) ^ 0xff;
    corruptPng[idatOffset + 7] = (corruptPng[idatOffset + 7] ?? 0) ^ 0xff;

    // Valid JPEG framing (SOI, SOF, SOS, EOI) with corrupted entropy-coded scan stream.
    const corruptJpeg = new Uint8Array(validJpeg);
    const sosOffset = Buffer.from(corruptJpeg).indexOf(
      Buffer.from([0xff, 0xda]),
    );
    const sosLength =
      ((corruptJpeg[sosOffset + 2] ?? 0) << 8) +
      (corruptJpeg[sosOffset + 3] ?? 0);
    const scanStart = sosOffset + 2 + sosLength;
    for (let i = scanStart; i < corruptJpeg.length - 4; i++) {
      corruptJpeg[i] = 0x55;
    }

    // Valid GIF framing (GIF89a 2x2) with invalid LZW code (> nextCode).
    const corruptGif = new Uint8Array([
      0x47,
      0x49,
      0x46,
      0x38,
      0x39,
      0x61,
      2,
      0,
      2,
      0,
      0x80,
      0,
      0,
      0xff,
      0xff,
      0xff,
      0x00,
      0x00,
      0x00,
      0x2c,
      0,
      0,
      0,
      0,
      2,
      0,
      2,
      0,
      0x00,
      0x02,
      0x02,
      0xbc, // first 3-bit code is 4 (clear), next is 7 (> nextCode=6)!
      0x00,
      0x00,
      0x3b,
    ]);

    const cases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof image>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "corrupt-png",
        node: image({ id: "bad", src: corruptPng, x: 10, y: 10 }),
        expected: /decode PNG/i,
      },
      {
        label: "corrupt-jpeg",
        node: image({ id: "bad", src: corruptJpeg, x: 10, y: 10 }),
        expected: /decode JPEG/i,
      },
      {
        label: "corrupt-gif",
        node: image({ id: "bad", src: corruptGif, x: 10, y: 10 }),
        expected: /decode GIF|GIF/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "ok-sibling" }).append(
            image({
              id: "ok-img",
              src: validPng,
              x: 20,
              y: 20,
              width: 60,
              height: 30,
            }),
          ),
        ),
      );
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, [], item.label);
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.equal(result.pages[0]?.diagnostics[0]?.page, 1, item.label);
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /image\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
    }
  });
});
