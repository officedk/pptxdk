import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  cell,
  chart,
  deck,
  font,
  paragraph,
  row,
  shape,
  slideNumber,
  slide,
  span,
  table,
  text,
} from "@pptxdk/capabilities/document";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import {
  P,
  SlideNumber,
  Span,
  Sub,
  Sup,
  Text,
} from "@pptxdk/capabilities/text-elements";
import { type CapabilityModule, CommonNode } from "@pptxdk/core";
import { CapabilityError, createResolver } from "@pptxdk/core/build";
import { getInternalResolution } from "@pptxdk/core/resolution";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { stripVariableFontTables } from "../../../test/sfnt-test-helpers.ts";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      resolution,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Slice 3: Text static preview & shared text-service outlines", () => {
  const jostFont = font("Jost*", [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-400-Book.ttf",
          import.meta.url,
        ),
      ),
    },
    {
      weight: 700,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-700-Bold.ttf",
          import.meta.url,
        ),
      ),
    },
  ]);
  const notoFont = font("Noto Sans CJK SC", [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
          import.meta.url,
        ),
      ),
    },
  ]);
  const notoArabicFont = font("Noto Naskh Arabic", [
    {
      weight: 400,
      data: stripVariableFontTables(
        readFileSync(
          new URL(
            "../../../test/assets/fonts/NotoNaskhArabic.ttf",
            import.meta.url,
          ),
        ),
      ),
    },
  ]);
  const notoHebrewFont = font("Noto Sans Hebrew Thin", [
    {
      weight: 100,
      data: stripVariableFontTables(
        readFileSync(
          new URL(
            "../../../test/assets/fonts/NotoSansHebrew.ttf",
            import.meta.url,
          ),
        ),
      ),
    },
  ]);

  it("renders wrapped CJK/Latin text, alignments, uppercase, slide number, and box visuals across JSX and document entries", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "text-main", background: "#F8FAFC" },
          createElement(
            Text,
            {
              id: "mixed",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 240,
                fontFamily: notoFont,
                fontSize: 18,
                padding: 10,
                background: "#EEF2FF",
                border: { color: "#4F46E5", width: 2 },
              },
            },
            createElement(
              P,
              { style: { lineHeight: 26, spacingAfter: 8 } },
              "共享排版 ",
              createElement(
                Span,
                {
                  style: {
                    fontFamily: jostFont,
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    textTransform: "uppercase",
                  },
                },
                "shared 2026",
              ),
              " 第",
              createElement(SlideNumber, null),
              "页",
            ),
            createElement(
              P,
              { style: { align: "right", lineHeight: 24 } },
              createElement(Span, { style: { fontFamily: jostFont } }, "Right"),
            ),
          ),
          createElement(
            Text,
            {
              id: "centered",
              style: {
                position: "absolute",
                left: 280,
                top: 40,
                width: 160,
                height: 120,
                fontFamily: jostFont,
                fontSize: 16,
                verticalAlign: "bottom",
                align: "center",
                background: "#FEF3C7",
                opacity: 0.9,
                rotate: 90,
              },
            },
            "Bottom center",
          ),
        ),
      ),
    );

    const impDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "text-main", background: "#F8FAFC" }).append(
        text(
          [
            paragraph(
              [
                "共享排版 ",
                span("shared 2026", {
                  style: {
                    fontFamily: jostFont,
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    textTransform: "uppercase",
                  },
                }),
                " 第",
                slideNumber(),
                "页",
              ],
              { style: { lineHeight: 26, spacingAfter: 8 } },
            ),
            paragraph([span("Right", { style: { fontFamily: jostFont } })], {
              style: { align: "right", lineHeight: 24 },
            }),
          ],
          {
            id: "mixed",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 240,
              fontFamily: notoFont,
              fontSize: 18,
              padding: 10,
              background: "#EEF2FF",
              border: { color: "#4F46E5", width: 2 },
            },
          },
        ),
        text("Bottom center", {
          id: "centered",
          style: {
            position: "absolute",
            left: 280,
            top: 40,
            width: 160,
            height: 120,
            fontFamily: jostFont,
            fontSize: 16,
            verticalAlign: "bottom",
            align: "center",
            background: "#FEF3C7",
            opacity: 0.9,
            rotate: 90,
          },
        }),
      ),
    );

    const jsxPreview = await previewDoc(jsxDoc);
    const impPreview = await previewDoc(impDoc);
    assert.deepEqual((await validatePptx(jsxPreview.pptx)).errors, []);
    assert.deepEqual((await validatePptx(impPreview.pptx)).errors, []);

    assert.deepEqual(jsxPreview.pages[0]?.errors, []);
    assert.deepEqual(impPreview.pages[0]?.errors, []);
    assert.ok(jsxPreview.pages[0]?.png !== undefined);
    assert.deepEqual(jsxPreview.pages[0]?.png, impPreview.pages[0]?.png);
    expect(jsxPreview.pages[0]?.svg).toContain("<path");
    expect(jsxPreview.pages[0]?.svg).toContain('fill="#EEF2FF"');
    expect(jsxPreview.pages[0]?.svg).toContain("rotate(90");

    const centeredLocator = jsxPreview.resolution
      .slide("text-main")
      .get("centered");
    expect(centeredLocator.bounds?.width).toBeCloseTo(120);
    expect(centeredLocator.bounds?.height).toBeCloseTo(160);
    assert.equal(
      jsxPreview.resolution.slide("text-main").get("mixed").text,
      "共享排版 shared 2026 第1页\nRight",
    );

    // Box visual effects (boxShadow, dashed border, borderRadius, gradient) render on Text with zero preview errors.
    const styledTextDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "styled-text", background: "#FFFFFF" }).append(
        text("Boxed Text", {
          id: "styled",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 180,
            height: 60,
            fontFamily: jostFont,
            padding: 10,
            background: "#EFF6FF",
            border: { color: "#111827", width: 2, style: "dashed" },
            borderRadius: 10,
            boxShadow: { blur: 4, distance: 4, angle: 90 },
          },
        }),
      ),
    );
    const styledPreview = await previewDoc(styledTextDoc);
    assert.deepEqual((await validatePptx(styledPreview.pptx)).errors, []);
    assert.deepEqual(styledPreview.pages[0]?.errors, []);
    assert.ok(styledPreview.pages[0]?.png !== undefined);
    expect(styledPreview.pages[0]?.svg).toContain("<feDropShadow");
    expect(styledPreview.pages[0]?.svg).toContain("stroke-dasharray");
  });

  it("rejects missing glyphs/fonts and color tables while keeping valid sibling pages", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const jostBytes = readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    );
    const withSyntheticTable = (source: Buffer, tag: string): Uint8Array => {
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
      for (let i = 0; i < numTables; i++) {
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
      for (let i = 0; i < records.length; i++) {
        const off = 12 + i * 16;
        out.write(records[i]!.tag, off, 4, "ascii");
        out.writeUInt32BE(records[i]!.checkSum, off + 4);
        out.writeUInt32BE(records[i]!.offset, off + 8);
        out.writeUInt32BE(records[i]!.length, off + 12);
      }
      source.copy(out, nextHeaderSize, headerSize);
      return new Uint8Array(out);
    };

    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof text>;
      readonly expectedCode:
        | "preview-missing-glyph"
        | "preview-missing-font"
        | "preview-unsupported-font";
      readonly expected: RegExp;
    }[] = [
      {
        label: "missing-glyph",
        node: text("汉", {
          id: "bad",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 160,
            height: 50,
            fontFamily: jostFont,
          },
        }),
        expectedCode: "preview-missing-glyph",
        expected: /glyph/i,
      },
      {
        label: "missing-bullet-glyph",
        node: text(
          [
            paragraph("Bullet", {
              style: { list: { type: "bullet", character: "汉" } },
            }),
          ],
          {
            id: "bad",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 160,
              height: 50,
              fontFamily: jostFont,
            },
          },
        ),
        expectedCode: "preview-missing-glyph",
        expected: /glyph/i,
      },
      {
        label: "missing-font-file",
        node: text("No font file", {
          id: "bad",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 160,
            height: 50,
            fontFamily: font("Missing Preview Face"),
          },
        }),
        expectedCode: "preview-missing-font",
        expected: /font file/i,
      },
      ...(["COLR", "CPAL", "sbix", "CBDT", "SVG "] as const).map((table) => ({
        label: `${table.trim().toLowerCase()}-font`,
        node: text("Color font", {
          id: "bad",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 160,
            height: 50,
            fontFamily: font(
              `${table.trim()}Jost`,
              [{ weight: 400, data: withSyntheticTable(jostBytes, table) }],
              { substitute: true },
            ),
          },
        }),
        expectedCode: "preview-unsupported-font" as const,
        expected: /monochrome outline/i,
      })),
    ];

    for (const item of rejectionCases) {
      const result = await previewDoc(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "good-page" }).append(
            text("Valid", {
              id: "ok",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 120,
                fontFamily: jostFont,
              },
            }),
          ),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.errors[0]?.code,
        item.expectedCode,
        item.label,
      );
      assert.match(
        result.pages[0]?.errors[0]?.path ?? "",
        /text\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.errors[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.deepEqual(result.pages[1]?.errors, [], item.label);
      assert.ok(result.pages[1]?.png !== undefined, item.label);
    }
  });

  it("closes paint() CapabilityError.code strictly to preview-unsupported-font and preview-missing-glyph across Text, Shape, Table, and Chart", () => {
    const resolveDoc = createResolver(documentCapabilities);
    const allowedCodes = new Set([
      "preview-missing-font",
      "preview-unsupported-font",
      "preview-missing-glyph",
      "preview-corrupt-image",
    ]);

    const capturePaintErrorCode = (doc: CommonNode): string => {
      const resolution = resolveDoc(doc);
      // Confirm resolve + emit succeed even when paint rejects underlying font limits.
      const emitted = emit(resolution);
      assert.ok(emitted.pptx.byteLength > 0);
      const internal = getInternalResolution(resolution);
      const slideData = internal.slides[0]!;
      const layer = slideData.layers[0]!;
      try {
        layer.paint?.({
          path: layer.path,
          slideNumber: slideData.index,
          objectId: layer.paintObjectId,
          ...(layer.box === undefined ? {} : { box: layer.box }),
          theme: layer.theme,
        });
        assert.fail("Expected layer.paint to throw a CapabilityError");
      } catch (error) {
        assert.ok(
          error instanceof CapabilityError,
          `Expected CapabilityError, got ${String(error)}`,
        );
        assert.ok(
          allowedCodes.has(error.code),
          `Unexpected CapabilityError.code in paint(): ${error.code}`,
        );
        return error.code;
      }
    };

    const missingFace = font("Missing Preview Face");

    // 1. Fixed-area Text without font files
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "p" }).append(
            text("No font file", {
              id: "t",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 160,
                height: 50,
                fontFamily: missingFace,
              },
            }),
          ),
        ),
      ),
      "preview-missing-font",
    );

    // 2. Fixed-area Text list prefix without font files
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "p" }).append(
            text([paragraph("Item", { style: { list: { type: "bullet" } } })], {
              id: "t-list",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 160,
                height: 50,
                fontFamily: missingFace,
              },
            }),
          ),
        ),
      ),
      "preview-missing-font",
    );

    // 3. Fixed-size Shape inline text without font files
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "p" }).append(
            shape("No font", {
              id: "s",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 160,
                height: 50,
                background: "#1E293B",
                fontFamily: missingFace,
                color: "#FFFFFF",
              },
            }),
          ),
        ),
      ),
      "preview-missing-font",
    );

    // 4. Fixed-row-height Table cell without font files
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "p" }).append(
            table({
              id: "tbl",
              columns: [120],
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                fontFamily: missingFace,
                fontSize: 14,
              },
            }).append(row({ style: { height: 30 } }).append(cell("Cell"))),
          ),
        ),
      ),
      "preview-missing-font",
    );

    // 5. Chart without font files
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 320, height: 200 }).append(
          slide({ id: "p" }).append(
            chart({
              id: "c",
              title: "No font",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 300,
                height: 180,
                fontSize: 12,
              },
              groups: [
                {
                  kind: "column",
                  data: { kind: "shared-categories", categories: ["A", "B"] },
                  series: [{ name: "S", values: [10, 20] }],
                },
              ],
            }),
          ),
        ),
      ),
      "preview-missing-font",
    );

    // 6. Missing glyph in supplied font
    assert.equal(
      capturePaintErrorCode(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "p" }).append(
            text("汉", {
              id: "t-glyph",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 160,
                height: 50,
                fontFamily: jostFont,
              },
            }),
          ),
        ),
      ),
      "preview-missing-glyph",
    );
  });

  it("renders underline, line-through, textShadow (preserving rotWithShape=0 under rotation), and positive/negative letterSpacing across resolve, emit, and paint", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 520, height: 240 }).append(
      slide({ id: "decorations", background: "#FFFFFF" }).append(
        text(
          [
            paragraph([
              span("Underlined ", {
                style: { textDecoration: "underline", color: "#1D4ED8" },
              }),
              span("Struck", {
                style: { textDecoration: "line-through", color: "#DC2626" },
              }),
            ]),
            paragraph(
              [
                span("Shadowed", {
                  style: {
                    textShadow: {
                      color: "#0F172A",
                      blur: 2,
                      distance: 3,
                      angle: 90,
                    },
                  },
                }),
              ],
              { style: { spacingAfter: 4 } },
            ),
          ],
          {
            id: "deco-box",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 240,
              height: 120,
              fontFamily: jostFont,
              fontSize: 18,
            },
          },
        ),
        text("Rotated Shadow", {
          id: "rot-shadow",
          style: {
            position: "absolute",
            left: 280,
            top: 40,
            width: 180,
            height: 50,
            fontFamily: jostFont,
            fontSize: 18,
            rotate: 90,
            textShadow: { color: "#000000", blur: 1, distance: 4, angle: 90 },
          },
        }),
        text("WIDE", {
          id: "wide-track",
          style: {
            position: "absolute",
            left: 20,
            top: 150,
            width: 200,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 18,
            letterSpacing: 6,
          },
        }),
        text("TIGHT", {
          id: "tight-track",
          style: {
            position: "absolute",
            left: 240,
            top: 150,
            width: 200,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 18,
            letterSpacing: -2,
          },
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    expect(slideXml).toContain('u="sng"');
    expect(slideXml).toContain('strike="sngStrike"');
    expect(slideXml).toContain('spc="450"');
    expect(slideXml).toContain('spc="-150"');

    const svg = page.svg!;
    expect(svg).toContain('fill="#1D4ED8"');
    expect(svg).toContain('fill="#DC2626"');
    // Both underline and line-through render rectangular decoration lines.
    expect(svg).toContain("<rect ");
    expect(svg).toContain("<feDropShadow");
    // Under 90-degree container rotation, screen-space 90-deg shadow (dx=0, dy=4) inverse-rotates to (4, 0).
    expect(svg).toMatch(/<feDropShadow dx="4" dy="0"/);
  });

  it("shapes and paints Kana, Hangul, Cyrillic, Greek, math symbols, and Arabic/Hebrew BiDi runs", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 540, height: 300 }).append(
      slide({ id: "multiscript", background: "#FFFFFF" }).append(
        text("あア 한 Д Ω ∑√≠", {
          id: "cjk-sym",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 240,
            height: 50,
            fontFamily: notoFont,
            fontSize: 18,
            color: "#0F172A",
          },
        }),
        text("مرحبا بالعالم 2026!", {
          id: "arabic-bidi",
          style: {
            position: "absolute",
            left: 20,
            top: 80,
            width: 260,
            height: 50,
            fontFamily: notoArabicFont,
            fontSize: 20,
            color: "#0F172A",
          },
        }),
        text(
          [
            paragraph([
              span("שלום ", { style: { color: "#047857" } }),
              span("2026 ABC", {
                style: {
                  fontFamily: jostFont,
                  fontWeight: 400,
                  color: "#7C3AED",
                },
              }),
              span(" עולם", { style: { color: "#B45309" } }),
            ]),
          ],
          {
            id: "hebrew-bidi",
            style: {
              position: "absolute",
              left: 20,
              top: 140,
              width: 300,
              height: 50,
              fontFamily: notoHebrewFont,
              fontWeight: 100,
              fontSize: 20,
            },
          },
        ),
        text(
          [
            paragraph([
              span("مرحبا", { style: { color: "#DC2626" } }),
              span(" بكم", { style: { color: "#2563EB" } }),
            ]),
          ],
          {
            id: "rtl-order",
            style: {
              position: "absolute",
              left: 20,
              top: 200,
              width: 220,
              height: 45,
              fontFamily: notoArabicFont,
              fontSize: 20,
            },
          },
        ),
        text(
          [
            paragraph([
              span("ID 42: ", {
                style: { fontFamily: jostFont, color: "#111827" },
              }),
              span("مرحبا", {
                style: { fontFamily: notoArabicFont, color: "#0284C7" },
              }),
            ]),
          ],
          {
            id: "ltr-with-rtl",
            style: {
              position: "absolute",
              left: 280,
              top: 200,
              width: 220,
              height: 45,
              fontSize: 18,
            },
          },
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const svg = page.svg!;
    const bluePos = svg.indexOf('fill="#2563EB"');
    const redPos = svg.indexOf('fill="#DC2626"');
    assert.ok(bluePos !== -1 && redPos !== -1);
    const blueMatch = /fill="#2563EB"[^>]*transform="translate\(([\d.-]+)/.exec(
      svg,
    );
    const redMatch = /fill="#DC2626"[^>]*transform="translate\(([\d.-]+)/.exec(
      svg,
    );
    assert.ok(blueMatch && redMatch);
    // In an RTL paragraph, the second logical span (" بكم") is visually to the left of the first ("مرحبا").
    assert.ok(Number(blueMatch[1]) < Number(redMatch[1]));

    const ltrMatch = /fill="#111827"[^>]*transform="translate\(([\d.-]+)/.exec(
      svg,
    );
    const rtlMatch = /fill="#0284C7"[^>]*transform="translate\(([\d.-]+)/.exec(
      svg,
    );
    assert.ok(ltrMatch && rtlMatch);
    assert.ok(Number(ltrMatch[1]) < Number(rtlMatch[1]));
  });

  it("justifies non-last soft-wrapped lines (spaces and CJK clusters) while keeping hard breaks and last lines natural", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 520, height: 260 }).append(
      slide({ id: "justify-slide", background: "#FFFFFF" }).append(
        text(
          [
            paragraph("Alpha beta gamma delta epsilon", {
              style: { align: "justify", color: "#1D4ED8" },
            }),
            paragraph("Hard break\nFinal short", {
              style: { align: "justify", color: "#0F172A" },
            }),
          ],
          {
            id: "latin-justify",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 170,
              height: 180,
              fontFamily: jostFont,
              fontSize: 16,
              padding: 10,
            },
          },
        ),
        text("汉字两端对齐自动折行测试末行", {
          id: "cjk-justify",
          style: {
            position: "absolute",
            left: 210,
            top: 20,
            width: 140,
            height: 140,
            fontFamily: notoFont,
            fontSize: 16,
            align: "justify",
            padding: 10,
            color: "#047857",
          },
        }),
        text("One two three four five six", {
          id: "ul-justify",
          style: {
            position: "absolute",
            left: 370,
            top: 20,
            width: 130,
            height: 100,
            fontFamily: jostFont,
            fontSize: 16,
            align: "justify",
            textDecoration: "underline",
            color: "#B45309",
          },
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const svg = page.svg!;
    // First line of latin-justify starts at x=30 (20+10) and innerWidth=150 (right edge 180).
    // Verify last glyph of first justified line reaches close to 180 (> 165).
    const blueGlyphXs = [
      ...svg.matchAll(
        /fill="#1D4ED8"[^>]*transform="translate\(([\d.-]+)\s+([\d.-]+)\)/g,
      ),
    ].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));
    assert.ok(blueGlyphXs.length > 0);
    const firstLineY = blueGlyphXs[0]!.y;
    const firstLineXs = blueGlyphXs
      .filter((g) => Math.abs(g.y - firstLineY) < 1)
      .map((g) => g.x);
    assert.ok(Math.max(...firstLineXs) > 160);

    // Last line of latin-justify ("Final short") should NOT stretch to 160.
    const darkGlyphs = [
      ...svg.matchAll(
        /fill="#0F172A"[^>]*transform="translate\(([\d.-]+)\s+([\d.-]+)\)/g,
      ),
    ].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));
    const maxDarkY = Math.max(...darkGlyphs.map((g) => g.y));
    const lastDarkXs = darkGlyphs
      .filter((g) => Math.abs(g.y - maxDarkY) < 1)
      .map((g) => g.x);
    assert.ok(Math.max(...lastDarkXs) < 120);

    // CJK no-space justify first line (x=220..340) reaches close to right edge (> 315).
    const cjkGlyphs = [
      ...svg.matchAll(
        /fill="#047857"[^>]*transform="translate\(([\d.-]+)\s+([\d.-]+)\)/g,
      ),
    ].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));
    const cjkFirstY = cjkGlyphs[0]!.y;
    const cjkFirstXs = cjkGlyphs
      .filter((g) => Math.abs(g.y - cjkFirstY) < 1)
      .map((g) => g.x);
    assert.ok(Math.max(...cjkFirstXs) > 315);

    // Justified underline rect spans full box width (~130).
    const ulWidths = [
      ...svg.matchAll(
        /<rect x="[\d.-]+" y="[\d.-]+" width="([\d.-]+)" height="[\d.-]+" fill="#B45309"/g,
      ),
    ].map((m) => Number(m[1]));
    assert.ok(ulWidths.some((w) => w > 125));
  });

  it("measures and renders bullet and numbered lists with sequential counters and multi-line hanging indents", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const wrappedItem = "Wrapped list item text across two lines";
    const doc = deck({ width: 540, height: 320 }).append(
      slide({ id: "lists-slide", background: "#FFFFFF" }).append(
        text(
          [
            paragraph(wrappedItem, {
              style: { list: { type: "bullet" }, color: "#1E293B" },
            }),
            paragraph("Nested custom", {
              style: {
                list: { type: "bullet", character: "+", level: 1 },
                color: "#2563EB",
              },
            }),
          ],
          {
            id: "bullet-box",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 190,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 16,
              padding: 8,
            },
          },
        ),
        text(
          [
            paragraph("One", { style: { list: { type: "number" } } }),
            paragraph("Two", { style: { list: { type: "number" } } }),
            paragraph("Five", {
              style: { list: { type: "number", startAt: 5 } },
            }),
            paragraph("Six", {
              style: { list: { type: "number", startAt: 5 } },
            }),
            paragraph("Break", { style: { list: { type: "bullet" } } }),
            paragraph("Restart", { style: { list: { type: "number" } } }),
          ],
          {
            id: "number-box",
            style: {
              position: "absolute",
              left: 240,
              top: 20,
              width: 240,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 16,
              padding: 8,
            },
          },
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const svg = page.svg!;
    // Bullet item has hanging indent: bullet glyph at x~28, text line 1 and line 2 start further right (>= 48).
    const darkGlyphs = [
      ...svg.matchAll(
        /fill="#1E293B"[^>]*transform="translate\(([\d.-]+)\s+([\d.-]+)\)/g,
      ),
    ].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));
    const bulletY = darkGlyphs[0]!.y;
    const line1Xs = darkGlyphs
      .filter((g) => Math.abs(g.y - bulletY) < 1)
      .map((g) => g.x);
    const line2Y = Math.max(...darkGlyphs.map((g) => g.y));
    const line2Xs = darkGlyphs
      .filter((g) => Math.abs(g.y - line2Y) < 1)
      .map((g) => g.x);
    assert.ok(line1Xs[0]! < 35); // bullet prefix near left padding (20+8=28)
    assert.ok(line1Xs[1]! >= 46); // text starts at hanging indent
    assert.ok(Math.abs(line2Xs[0]! - line1Xs[1]!) < 1); // continuation line aligns with text start!

    // Nested level 1 bullet starts further right than level 0.
    const nestedGlyphs = [
      ...svg.matchAll(/fill="#2563EB"[^>]*transform="translate\(([\d.-]+)/g),
    ].map((m) => Number(m[1]));
    assert.ok(nestedGlyphs[0]! > line1Xs[0]! + 15);
  });

  it("applies inter-cluster letterSpacing consistently between line wrapping measurement and painted glyph advances", async () => {
    const resolveWith = createResolver(documentCapabilities);
    const doc = deck({ width: 400, height: 220 }).append(
      slide({ id: "tracking-wrap", background: "#FFFFFF" }).append(
        text("A", {
          id: "single-char",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 18,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
            letterSpacing: -6,
            padding: 0,
            color: "#0F172A",
          },
        }),
        text("AB", {
          id: "negative-wrap",
          style: {
            position: "absolute",
            left: 60,
            top: 20,
            width: 18,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
            letterSpacing: -6,
            padding: 0,
            color: "#111827",
          },
        }),
        text(
          paragraph([
            span("A", { style: { letterSpacing: -6 } }),
            span("B", { style: { letterSpacing: -6 } }),
          ]),
          {
            id: "multi-run-wrap",
            style: {
              position: "absolute",
              left: 100,
              top: 20,
              width: 18,
              height: "auto",
              fontFamily: jostFont,
              fontSize: 20,
              padding: 0,
              color: "#1D4ED8",
            },
          },
        ),
        text("AB", {
          id: "center-neg",
          style: {
            position: "absolute",
            left: 160,
            top: 20,
            width: 80,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
            letterSpacing: -6,
            align: "center",
            padding: 0,
            color: "#047857",
          },
        }),
        text("AB", {
          id: "right-neg",
          style: {
            position: "absolute",
            left: 260,
            top: 20,
            width: 80,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
            letterSpacing: -6,
            align: "right",
            padding: 0,
            color: "#B91C1C",
          },
        }),
      ),
    );

    const resolution = resolveWith(doc);
    const { pptx } = emit(resolution);
    assert.deepEqual((await validatePptx(pptx)).errors, []);
    const page = paint(resolution).pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const trackingSlide = resolution.slide("tracking-wrap");
    const singleBox = trackingSlide.get("single-char").box!;
    const negWrapBox = trackingSlide.get("negative-wrap").box!;
    const multiRunBox = trackingSlide.get("multi-run-wrap").box!;

    // Single-line "A" height vs wrapped 2-line "AB" height (both single-run and multi-run).
    assert.ok(
      negWrapBox.height > singleBox.height * 1.5,
      `Expected "AB" with letterSpacing=-6 in width=20 to wrap onto 2 lines (single=${singleBox.height}, wrapped=${negWrapBox.height})`,
    );
    assert.equal(multiRunBox.height, negWrapBox.height);

    const svg = page.svg!;
    const extractGlyphs = (color: string) =>
      [
        ...svg.matchAll(
          new RegExp(
            `fill="${color}"[^>]*transform="translate\\(([\\d.-]+)\\s+([\\d.-]+)\\)`,
            "g",
          ),
        ),
      ].map((m) => ({ x: Number(m[1]), y: Number(m[2]) }));

    const wrappedGlyphs = extractGlyphs("#111827");
    assert.equal(wrappedGlyphs.length, 2);
    assert.ok(wrappedGlyphs[1]!.y > wrappedGlyphs[0]!.y + 5);

    const multiRunGlyphs = extractGlyphs("#1D4ED8");
    assert.equal(multiRunGlyphs.length, 2);
    assert.ok(multiRunGlyphs[1]!.y > multiRunGlyphs[0]!.y + 5);

    const centerGlyphs = extractGlyphs("#047857");
    const rightGlyphs = extractGlyphs("#B91C1C");
    assert.equal(centerGlyphs.length, 2);
    assert.equal(rightGlyphs.length, 2);
    // Relative x within each 80px box: right-aligned glyphs sit strictly further right than center-aligned glyphs.
    assert.ok(rightGlyphs[0]!.x - 260 > centerGlyphs[0]!.x - 160 + 10);
  });

  it("paints proportional lineHeight, scaled/shifted Sup/Sub/baselineShift, and parity-aligned Shape/Text list indents", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "t03-paint", background: "#FFFFFF" },
          createElement(
            Text,
            {
              id: "lh-100",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 60,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 20,
                lineHeight: "100%",
                padding: 0,
                color: "#111111",
              },
            },
            "Word Wrap",
          ),
          createElement(
            Text,
            {
              id: "lh-150",
              style: {
                position: "absolute",
                left: 150,
                top: 20,
                width: 60,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 20,
                lineHeight: "150%",
                padding: 0,
                color: "#222222",
              },
            },
            "Word Wrap",
          ),
          createElement(
            Text,
            {
              id: "sub-sup",
              style: {
                position: "absolute",
                left: 20,
                top: 140,
                width: 260,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 24,
                padding: 0,
              },
            },
            createElement(Span, { style: { color: "#0F172A" } }, "A"),
            createElement(Sup, { style: { color: "#DC2626" } }, "B"),
            createElement(Sub, { style: { color: "#2563EB" } }, "C"),
            createElement(
              Span,
              { style: { color: "#059669", baselineShift: "40%" } },
              "D",
            ),
          ),
          createElement(
            Shape,
            {
              id: "list-shape",
              style: {
                position: "absolute",
                left: 320,
                top: 20,
                width: 260,
                height: 180,
                fontFamily: jostFont,
                fontSize: 18,
                padding: 10,
                verticalAlign: "top",
              },
            },
            createElement(
              P,
              { style: { list: { type: "bullet", character: "•", level: 0 } } },
              createElement(
                Span,
                { style: { color: "#7C3AED" } },
                "Wrapped list item in shape container",
              ),
            ),
          ),
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);

    const slideRes = result.resolution.slide("t03-paint");
    const box100 = slideRes.get("lh-100").box!;
    const box150 = slideRes.get("lh-150").box!;
    assert.ok(Math.abs(box150.height / box100.height - 1.5) < 0.02);

    const svg = page.svg!;
    const extractPathTransforms = (color: string) =>
      [
        ...svg.matchAll(
          new RegExp(
            `fill="${color}"[^>]*transform="translate\\(([\\d.-]+)\\s+([\\d.-]+)\\)\\s*scale\\(([\\d.-]+)`,
            "g",
          ),
        ),
      ].map((m) => ({
        x: Number(m[1]),
        y: Number(m[2]),
        scale: Number(m[3]),
      }));

    // Proportional lineHeight increases SVG baseline separation by 1.5x.
    const lh100Glyphs = extractPathTransforms("#111111");
    const lh150Glyphs = extractPathTransforms("#222222");
    const sep100 = Math.max(...lh100Glyphs.map((g) => g.y)) - lh100Glyphs[0]!.y;
    const sep150 = Math.max(...lh150Glyphs.map((g) => g.y)) - lh150Glyphs[0]!.y;
    assert.ok(Math.abs(sep150 / sep100 - 1.5) < 0.02);

    // Sup shifts upward & scales ~0.65x; Sub shifts downward & scales ~0.65x; baselineShift="40%" shifts further up.
    const baseGlyph = extractPathTransforms("#0F172A")[0]!;
    const supGlyph = extractPathTransforms("#DC2626")[0]!;
    const subGlyph = extractPathTransforms("#2563EB")[0]!;
    const shiftGlyph = extractPathTransforms("#059669")[0]!;

    assert.ok(supGlyph.y < baseGlyph.y - 4);
    assert.ok(subGlyph.y > baseGlyph.y + 3);
    assert.ok(shiftGlyph.y < supGlyph.y - 1);
    assert.ok(Math.abs(supGlyph.scale / baseGlyph.scale - 0.65) < 0.02);
    assert.ok(Math.abs(subGlyph.scale / baseGlyph.scale - 0.65) < 0.02);

    // Shape bullet prefix & wrapped continuation lines match emitted marL / indent parity.
    const shapeGlyphs = extractPathTransforms("#7C3AED");
    const bulletX = shapeGlyphs[0]!.x;
    const firstLineY = shapeGlyphs[0]!.y;
    const firstTextX = shapeGlyphs.find(
      (g) => Math.abs(g.y - firstLineY) < 1 && g.x > bulletX + 10,
    )!.x;
    const secondLineY = Math.max(...shapeGlyphs.map((g) => g.y));
    const secondTextX = shapeGlyphs.find(
      (g) => Math.abs(g.y - secondLineY) < 1,
    )!.x;

    assert.ok(bulletX >= 329 && bulletX <= 333); // left 320 + padding 10
    assert.ok(firstTextX - bulletX >= 20);
    assert.ok(Math.abs(secondTextX - firstTextX) < 0.5);

    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /<a:spcPct val="150000"\/>/u);
    assert.match(slideXml, /baseline="30000"/u);
    assert.match(slideXml, /baseline="-25000"/u);
    assert.match(slideXml, /baseline="40000"/u);
    assert.match(
      slideXml,
      /<a:pPr[^>]*marL="214313"[^>]*indent="-214313"[^>]*><a:buChar char="•"\/>/u,
    );
  });
});
