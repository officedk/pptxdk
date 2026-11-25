import { readFileSync } from "node:fs";
import { deflateSync } from "node:zlib";

import {
  CodegenError,
  CommonNode,
  build,
  codegen,
  codegenDiagnosticSchema,
  font,
  preset,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolve,
  shapeCapability,
  type CodegenDiagnostic,
  type PreviewError,
} from "@officedk/pptxdk";
import {
  createCodegen,
  createResolver,
  previewErrorSchema as buildPreviewErrorSchema,
  previewPlaceholderSchema as buildPreviewPlaceholderSchema,
} from "@officedk/pptxdk/build";
import {
  previewErrorSchema as nodePreviewErrorSchema,
  previewPlaceholderSchema as nodePreviewPlaceholderSchema,
} from "@officedk/pptxdk/node";
import {
  paint,
  preview,
  previewErrorSchema as previewSubpathErrorSchema,
  previewPlaceholderSchema as previewSubpathPlaceholderSchema,
} from "@officedk/pptxdk/preview";
import { Deck, Image, Shape, Slide, Text } from "@officedk/pptxdk/react";
import {
  box,
  cell,
  chart,
  deck,
  image,
  notes,
  row,
  shape,
  slide,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

const jostRegularBytes = readFileSync(
  new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
);

const jostFont = font("Jost*", [{ weight: 400, data: jostRegularBytes }]);

function withSyntheticFontTable(source: Buffer, tag: string): Uint8Array {
  const numTables = source.readUInt16BE(4);
  const headerSize = 12 + numTables * 16;
  const nextNumTables = numTables + 1;
  const nextHeaderSize = 12 + nextNumTables * 16;
  const records: {
    tag: string;
    checkSum: number;
    offset: number;
    length: number;
  }[] = [];
  for (let i = 0; i < numTables; i += 1) {
    const off = 12 + i * 16;
    records.push({
      tag: source.subarray(off, off + 4).toString("ascii"),
      checkSum: source.readUInt32BE(off + 4),
      offset: source.readUInt32BE(off + 8) + 16,
      length: source.readUInt32BE(off + 12),
    });
  }
  records.push({
    tag,
    checkSum: 0,
    offset: source.length + 16,
    length: 16,
  });
  records.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const out = Buffer.alloc(source.length + 32);
  source.copy(out, 0, 0, 12);
  out.writeUInt16BE(nextNumTables, 4);
  const entrySelector = Math.floor(Math.log2(nextNumTables));
  const searchRange = (1 << entrySelector) * 16;
  out.writeUInt16BE(searchRange, 6);
  out.writeUInt16BE(entrySelector, 8);
  out.writeUInt16BE(nextNumTables * 16 - searchRange, 10);
  for (let i = 0; i < records.length; i += 1) {
    const off = 12 + i * 16;
    out.write(records[i]!.tag, off, 4, "ascii");
    out.writeUInt32BE(records[i]!.checkSum, off + 4);
    out.writeUInt32BE(records[i]!.offset, off + 8);
    out.writeUInt32BE(records[i]!.length, off + 12);
  }
  source.copy(out, nextHeaderSize, headerSize);
  return new Uint8Array(out);
}

const colorFont = font(
  "ColrJost",
  [{ weight: 400, data: withSyntheticFontTable(jostRegularBytes, "COLR") }],
  { substitute: true },
);

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc >>> 1) ^ ((crc & 1) !== 0 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.byteLength, 0);
  const body = Buffer.concat([typeBuf, Buffer.from(data)]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([lenBuf, body, crcBuf]);
}

function createSolidPng(width: number, height: number): Uint8Array {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  const rowSize = 1 + width * 4;
  const raw = Buffer.alloc(height * rowSize);
  for (let y = 0; y < height; y += 1) {
    raw[y * rowSize] = 0;
    for (let x = 0; x < width; x += 1) {
      const idx = y * rowSize + 1 + x * 4;
      raw[idx] = 0x25;
      raw[idx + 1] = 0x63;
      raw[idx + 2] = 0xeb;
      raw[idx + 3] = 0xff;
    }
  }
  const idat = deflateSync(raw);
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      pngChunk("IHDR", ihdr),
      pngChunk("IDAT", idat),
      pngChunk("IEND", new Uint8Array(0)),
    ]),
  );
}

const validPngBytes = createSolidPng(4, 4);
const corruptPngBytes = (() => {
  const copy = new Uint8Array(validPngBytes);
  const idatOffset = Buffer.from(copy).indexOf(Buffer.from("IDAT"));
  copy[idatOffset + 6] = (copy[idatOffset + 6] ?? 0) ^ 0xff;
  copy[idatOffset + 7] = (copy[idatOffset + 7] ?? 0) ^ 0xff;
  return copy;
})();

describe("typed stage errors & schemas — CodegenDiagnostic union", () => {
  it("narrows CodegenDiagnostic by code with path, optional field, and branch payloads without severity", () => {
    const bareCodegen = createCodegen([]);

    const unexpressibleDoc = deck({
      width: Number.NaN,
      height: 540,
    }).append(slide({ id: "s1" }));

    let unexpressibleDiag: CodegenDiagnostic | undefined;
    try {
      codegen(unexpressibleDoc, { target: "vanilla" });
      expect.unreachable("Expected CodegenError on NaN width");
    } catch (error) {
      expect(error).toBeInstanceOf(CodegenError);
      unexpressibleDiag = (error as CodegenError).diagnostics[0];
    }
    expect(unexpressibleDiag).toBeDefined();
    expect("severity" in (unexpressibleDiag as object)).toBe(false);
    expect(codegenDiagnosticSchema.parse(unexpressibleDiag)).toEqual(
      unexpressibleDiag,
    );
    if (unexpressibleDiag?.code === "unexpressible-value") {
      expect(unexpressibleDiag.path).toBe("deck@width");
      expect(unexpressibleDiag.field).toBe("width");
    } else {
      expect.unreachable(
        `Expected unexpressible-value, received ${unexpressibleDiag?.code}`,
      );
    }

    const unsupportedPropDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1", transition: { type: "transition", kind: "fade" } }),
    );
    let unsupportedPropDiag: CodegenDiagnostic | undefined;
    try {
      bareCodegen(unsupportedPropDoc, { target: "react" });
      expect.unreachable("Expected CodegenError on unsupported slide property");
    } catch (error) {
      expect(error).toBeInstanceOf(CodegenError);
      unsupportedPropDiag = (error as CodegenError).diagnostics[0];
    }
    expect(unsupportedPropDiag).toBeDefined();
    expect("severity" in (unsupportedPropDiag as object)).toBe(false);
    expect(codegenDiagnosticSchema.parse(unsupportedPropDiag)).toEqual(
      unsupportedPropDiag,
    );
    if (unsupportedPropDiag?.code === "unsupported-slide-property") {
      expect(unsupportedPropDiag.path).toBe("deck/slide[s1]@transition");
      expect(unsupportedPropDiag.field).toBe("transition");
    } else {
      expect.unreachable(
        `Expected unsupported-slide-property, received ${unsupportedPropDiag?.code}`,
      );
    }

    expect(
      codegenDiagnosticSchema.safeParse({
        code: "unsupported-declaration",
        path: "deck/slide[s1]/custom[c1]",
        message: "Unsupported node",
        declaration: "custom",
      }).success,
    ).toBe(true);
    expect(
      codegenDiagnosticSchema.safeParse({
        code: "unsupported-declaration",
        path: "deck/slide[s1]/custom[c1]",
        message: "Missing declaration field",
      }).success,
    ).toBe(false);
    expect(
      codegenDiagnosticSchema.safeParse({
        severity: "error",
        code: "invalid-options",
        path: "deck",
        message: "Legacy severity field",
      }).success,
    ).toBe(false);
  });
});

describe("typed stage errors & schemas — PreviewError, PreviewPlaceholder, and page isolation", () => {
  it("exports identical previewErrorSchema and previewPlaceholderSchema across ., ./preview, ./node, and ./build", () => {
    expect(previewErrorSchema).toBeDefined();
    expect(previewPlaceholderSchema).toBeDefined();
    expect(previewSubpathErrorSchema).toBe(previewErrorSchema);
    expect(nodePreviewErrorSchema).toBe(previewErrorSchema);
    expect(buildPreviewErrorSchema).toBe(previewErrorSchema);

    expect(previewSubpathPlaceholderSchema).toBe(previewPlaceholderSchema);
    expect(nodePreviewPlaceholderSchema).toBe(previewPlaceholderSchema);
    expect(buildPreviewPlaceholderSchema).toBe(previewPlaceholderSchema);
  });

  it("records typed PreviewError variants on failed pages, omits svg/png on failed pages, and renders healthy sibling pages across JSX and vanilla", async () => {
    const missingFontToken = font("Unbundled Sans");

    const jsxDeck = (
      <Deck width={320} height={180}>
        <Slide id="missing-font-page">
          <Text
            id="t-missing"
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              width: 200,
              height: 48,
              fontFamily: missingFontToken,
              fontWeight: 700,
              fontStyle: "italic",
            }}
          >
            System font text
          </Text>
        </Slide>
        <Slide id="unsupported-font-page">
          <Text
            id="t-colr"
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              width: 200,
              height: 48,
              fontFamily: colorFont,
            }}
          >
            Color font
          </Text>
        </Slide>
        <Slide id="missing-glyph-page">
          <Text
            id="t-glyph"
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              width: 200,
              height: 48,
              fontFamily: jostFont,
            }}
          >
            汉A
          </Text>
        </Slide>
        <Slide id="corrupt-image-page">
          <Image
            id="img-corrupt"
            src={corruptPngBytes}
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              width: 64,
              height: 64,
            }}
          />
        </Slide>
        <Slide id="healthy-page">
          <Shape
            id="ok-box"
            style={{
              position: "absolute",
              left: 20,
              top: 20,
              width: 120,
              height: 80,
              background: "#0D9488",
            }}
          />
        </Slide>
      </Deck>
    );

    const vanillaDeck = deck({ width: 320, height: 180 }).append(
      slide({ id: "missing-font-page" }).append(
        text("System font text", {
          id: "t-missing",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 200,
            height: 48,
            fontFamily: missingFontToken,
            fontWeight: 700,
            fontStyle: "italic",
          },
        }),
      ),
      slide({ id: "unsupported-font-page" }).append(
        text("Color font", {
          id: "t-colr",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 200,
            height: 48,
            fontFamily: colorFont,
          },
        }),
      ),
      slide({ id: "missing-glyph-page" }).append(
        text("汉A", {
          id: "t-glyph",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 200,
            height: 48,
            fontFamily: jostFont,
          },
        }),
      ),
      slide({ id: "corrupt-image-page" }).append(
        image({
          id: "img-corrupt",
          src: corruptPngBytes,
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 64,
            height: 64,
          },
        }),
      ),
      slide({ id: "healthy-page" }).append(
        shape({
          id: "ok-box",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 120,
            height: 80,
            background: "#0D9488",
          },
        }),
      ),
    );

    // PPTX build without embedFonts succeeds even when preview fails on font/pixel limits.
    const builtVanilla = await build(vanillaDeck);
    expect(builtVanilla.pptx.byteLength).toBeGreaterThan(0);

    for (const candidate of [jsxDeck, vanillaDeck]) {
      const result = await preview(candidate);
      expect(result.pages).toHaveLength(5);

      const [p1, p2, p3, p4, p5] = result.pages;

      // Page 1: preview-missing-font (carries family, weight, style)
      expect(p1?.svg).toBeUndefined();
      expect(p1?.png).toBeUndefined();
      expect(p1?.errors).toHaveLength(1);
      const err1: PreviewError | undefined = p1?.errors[0];
      expect(previewErrorSchema.parse(err1)).toEqual(err1);
      expect("severity" in (err1 as object)).toBe(false);
      if (err1?.code === "preview-missing-font") {
        expect(err1.path).toBe("deck/slide[missing-font-page]/text[t-missing]");
        expect(err1.font).toEqual({
          family: "Unbundled Sans",
          weight: 700,
          style: "italic",
        });
      } else {
        expect.unreachable(`Expected preview-missing-font, got ${err1?.code}`);
      }

      // Page 2: preview-unsupported-font
      expect(p2?.svg).toBeUndefined();
      expect(p2?.png).toBeUndefined();
      expect(p2?.errors).toHaveLength(1);
      const err2 = p2?.errors[0];
      expect(previewErrorSchema.parse(err2)).toEqual(err2);
      if (err2?.code === "preview-unsupported-font") {
        expect(err2.path).toBe(
          "deck/slide[unsupported-font-page]/text[t-colr]",
        );
        expect(err2.font.family).toBe("ColrJost");
      } else {
        expect.unreachable(
          `Expected preview-unsupported-font, got ${err2?.code}`,
        );
      }

      // Page 3: preview-missing-glyph
      expect(p3?.svg).toBeUndefined();
      expect(p3?.png).toBeUndefined();
      expect(p3?.errors).toHaveLength(1);
      const err3 = p3?.errors[0];
      expect(previewErrorSchema.parse(err3)).toEqual(err3);
      if (err3?.code === "preview-missing-glyph") {
        expect(err3.path).toBe("deck/slide[missing-glyph-page]/text[t-glyph]");
        expect(err3.font).toEqual({
          family: "Jost*",
          weight: 400,
          style: "normal",
        });
        expect(err3.characters).toEqual(["汉"]);
      } else {
        expect.unreachable(`Expected preview-missing-glyph, got ${err3?.code}`);
      }

      // Page 4: preview-corrupt-image
      expect(p4?.svg).toBeUndefined();
      expect(p4?.png).toBeUndefined();
      expect(p4?.errors).toHaveLength(1);
      const err4 = p4?.errors[0];
      expect(previewErrorSchema.parse(err4)).toEqual(err4);
      if (err4?.code === "preview-corrupt-image") {
        expect(err4.path).toBe(
          "deck/slide[corrupt-image-page]/image[img-corrupt]",
        );
        expect(err4.format).toBe("png");
      } else {
        expect.unreachable(`Expected preview-corrupt-image, got ${err4?.code}`);
      }

      // Page 5: healthy sibling renders SVG & PNG normally with zero errors
      expect(p5?.errors).toEqual([]);
      expect(p5?.placeholders).toEqual([]);
      expect(typeof p5?.svg).toBe("string");
      expect(p5?.png).toBeInstanceOf(Uint8Array);
      expect(p5?.png?.byteLength).toBeGreaterThan(0);
    }
  });

  it("records preview-unsupported-capability when an object capability omits paint while rendering sibling pages", () => {
    const resolveCustom = createResolver([
      shapeCapability,
      {
        declaration: "custom-unpainted",
        bind(input) {
          const authorId =
            typeof input.props["id"] === "string"
              ? input.props["id"]
              : undefined;
          return {
            contract: {
              nativeKind: "sp",
              layoutStyle: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 80,
                height: 40,
              },
            },
            resolve(ctx) {
              return {
                observation: {
                  id: ctx.id,
                  ...(authorId === undefined ? {} : { authorId }),
                  kind: "custom-unpainted",
                  name: "Custom 1",
                  box: ctx.layoutBox ?? { x: 10, y: 10, width: 80, height: 40 },
                },
                extensions: [],
                emit() {
                  return { xml: "<p:sp/>" };
                },
              };
            },
          };
        },
      },
    ]);

    const doc = deck({ width: 240, height: 140 }).append(
      slide({ id: "unpainted-slide" }).append(
        new CommonNode("custom-unpainted", "custom-unpainted", { id: "u1" }),
      ),
      slide({ id: "painted-slide" }).append(
        shape({
          id: "s1",
          style: {
            position: "absolute",
            left: 12,
            top: 12,
            width: 100,
            height: 60,
            background: "#2563EB",
          },
        }),
      ),
    );

    const painted = paint(resolveCustom(doc));
    expect(painted.pages).toHaveLength(2);
    const [failedPage, okPage] = painted.pages;
    expect(failedPage?.svg).toBeUndefined();
    expect(failedPage?.png).toBeUndefined();
    expect(failedPage?.errors).toHaveLength(1);
    const capErr = failedPage?.errors[0];
    expect(previewErrorSchema.parse(capErr)).toEqual(capErr);
    if (capErr?.code === "preview-unsupported-capability") {
      expect(capErr.path).toBe(
        "deck/slide[unpainted-slide]/custom-unpainted[u1]",
      );
      expect(capErr.capability).toBe("custom-unpainted");
    } else {
      expect.unreachable(
        `Expected preview-unsupported-capability, got ${capErr?.code}`,
      );
    }

    expect(okPage?.errors).toEqual([]);
    expect(okPage?.png).toBeInstanceOf(Uint8Array);
  });

  it("renders external images offline as deterministic placeholders with PreviewPageResult.placeholders and ResolvedObject.image without errors or warnings", async () => {
    const heroUrl = "https://cdn.example.com/hero.png";
    const badgeUrl = "https://cdn.example.com/badge.png";

    const doc = deck({ width: 360, height: 200 }).append(
      slide({ id: "external-media" }).append(
        image({
          id: "hero-img",
          src: { externalUrl: heroUrl },
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 140,
            height: 90,
          },
        }),
        shape({
          id: "badge-shape",
          style: {
            position: "absolute",
            left: 180,
            top: 20,
            width: 120,
            height: 90,
            background: { image: { externalUrl: badgeUrl } },
          },
        }),
        image({
          id: "embedded-img",
          src: validPngBytes,
          style: {
            position: "absolute",
            left: 20,
            top: 125,
            width: 48,
            height: 48,
          },
        }),
      ),
    );

    const resolution = resolve(doc);
    const heroLocator = resolution.slide("external-media").get("hero-img");
    expect(heroLocator.image).toEqual({
      kind: "external",
      id: "hero-img",
      path: "deck/slide[external-media]/image[hero-img]",
      url: heroUrl,
      placeholder: {
        id: "hero-img",
        path: "deck/slide[external-media]/image[hero-img]",
        url: heroUrl,
      },
    });

    const embeddedLocator = resolution
      .slide("external-media")
      .get("embedded-img");
    expect(embeddedLocator.image).toEqual({
      kind: "embedded",
      contentType: "image/png",
      width: 4,
      height: 4,
    });

    const previewed = await preview(doc);
    expect(previewed.pages).toHaveLength(1);
    const page = previewed.pages[0]!;
    expect(page.errors).toEqual([]);
    expect(page.diagnostics).toEqual([]);
    expect(typeof page.svg).toBe("string");
    expect(page.png).toBeInstanceOf(Uint8Array);
    expect(page.png!.byteLength).toBeGreaterThan(0);

    expect(page.placeholders).toEqual([
      {
        id: "hero-img",
        path: "deck/slide[external-media]/image[hero-img]",
        url: heroUrl,
      },
      {
        id: "badge-shape",
        path: "deck/slide[external-media]/shape[badge-shape]",
        url: badgeUrl,
      },
    ]);
    for (const item of page.placeholders) {
      expect(previewPlaceholderSchema.parse(item)).toEqual(item);
    }
  });

  it("emits zero runtime warnings across canvas/grid overflow, geometry clamping, table track deltas, overset text, and unbundled fixed-region fonts while preserving OOXML clamping and cell SVG clipping", async () => {
    const systemSans = font("System Sans");
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "painted-former-warnings" }).append(
        shape("Overset line one\nOverset line two\nOverset line three", {
          id: "clamped-overflow-shape",
          geometry: preset("roundRect", { adj: 95000 }),
          style: {
            position: "absolute",
            left: 220,
            top: 120,
            width: 160,
            height: 24,
            borderRadius: 40,
            background: "#DBEAFE",
            fontFamily: jostFont,
            fontSize: 18,
          },
        }),
        box({
          id: "overflowing-grid",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: 60,
            display: "grid",
            gridTemplateColumns: "80px 80px",
            gridTemplateRows: "40px",
          },
        }).append(
          shape({
            id: "grid-item-wide",
            style: {
              width: 120,
              height: 50,
              background: "#FCA5A5",
            },
          }),
        ),
        table({
          id: "clipped-table",
          columns: [90, 90],
          style: {
            position: "absolute",
            left: 10,
            top: 80,
            width: 120,
            height: 36,
            fontFamily: jostFont,
            fontSize: 16,
          },
        }).append(
          row({ style: { height: 22 } }).append(
            cell(
              "Long cell text that wraps onto multiple lines inside a fixed row",
              {
                id: "overset-cell",
              },
            ),
            cell("OK", { id: "ok-cell" }),
          ),
        ),
      ),
      slide({ id: "system-font-fixed", footer: "Confidential" }).append(
        text("Fixed-box system font", {
          id: "sys-text",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 180,
            height: 40,
            fontFamily: systemSans,
          },
        }),
        chart({
          id: "sys-chart",
          title: "Revenue",
          style: {
            position: "absolute",
            left: 16,
            top: 64,
            width: 240,
            height: 100,
            fontFamily: systemSans,
            fontSize: 12,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [{ name: "Rev", values: [12, 18] }],
            },
          ],
        }),
        notes("Speaker notes with unbundled font", {
          style: { fontFamily: systemSans },
        }),
      ),
    );

    const resolution = resolve(doc);
    expect(resolution.report.diagnostics).toEqual([]);

    const clampedShape = resolution
      .slide("painted-former-warnings")
      .get("clamped-overflow-shape");
    expect(clampedShape.textMetrics?.status).toBe("measured");
    if (clampedShape.textMetrics?.status === "measured") {
      expect(clampedShape.textMetrics.fits).toBe(false);
    }

    const sysText = resolution.slide("system-font-fixed").get("sys-text");
    expect(sysText.textMetrics).toEqual({
      status: "unmeasured",
      reason: "missing-font-resource",
      fonts: [{ family: "System Sans", weight: 400, style: "normal" }],
    });

    const built = await build(doc);
    expect(built.report.diagnostics).toEqual([]);

    // Verify OOXML roundRect adjustment clamping (<= 50000) is preserved in slide1.xml.
    const parts = unzipSync(built.pptx);
    const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    expect(slide1Xml).toContain('fmla="val 50000"');

    // Preview Page 1: zero diagnostics/errors, and overset cell retains SVG <clipPath> from textMetrics.
    const paintedPage1 = await preview(doc, { pages: [1] });
    expect(paintedPage1.report.diagnostics).toEqual([]);
    expect(paintedPage1.pages[0]?.errors).toEqual([]);
    expect(paintedPage1.pages[0]?.diagnostics).toEqual([]);
    expect(paintedPage1.pages[0]?.svg).toContain("<clipPath");
    expect(paintedPage1.pages[0]?.png).toBeInstanceOf(Uint8Array);
  });
});
