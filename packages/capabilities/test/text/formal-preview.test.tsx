import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  documentCapabilities,
  font,
  paragraph,
  slideNumber,
  slide,
  span,
  text,
} from "@react-pptx/capabilities/document";
import {
  P,
  SlideNumber,
  Span,
  Text,
} from "@react-pptx/capabilities/text-elements";
import { type CapabilityModule, CommonNode } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
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

describe("Slice 3: Text static preview & shared text-service outlines", () => {
  const jostFont = font("Jost*", [
    {
      weight: 400,
      data: readFileSync(
        new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    },
    {
      weight: 700,
      data: readFileSync(
        new URL("../assets/fonts/Jost-700-Bold.ttf", import.meta.url),
      ),
    },
  ]);
  const notoFont = font("Noto Sans CJK SC", [
    {
      weight: 400,
      data: readFileSync(
        new URL("../assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    },
  ]);
  const notoArabicFont = font("Noto Naskh Arabic", [
    {
      weight: 400,
      data: readFileSync(
        new URL("../assets/fonts/NotoNaskhArabic.ttf", import.meta.url),
      ),
    },
  ]);
  const notoHebrewFont = font("Noto Sans Hebrew Thin", [
    {
      weight: 100,
      data: readFileSync(
        new URL("../assets/fonts/NotoSansHebrew.ttf", import.meta.url),
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
              x: 20,
              y: 20,
              width: 240,
              style: {
                fontFamily: notoFont,
                fontSize: 18,
                padding: 10,
                backgroundColor: "#EEF2FF",
                borderColor: "#4F46E5",
                borderWidth: 2,
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
              x: 280,
              y: 40,
              width: 160,
              height: 120,
              style: {
                fontFamily: jostFont,
                fontSize: 16,
                verticalAlign: "bottom",
                align: "center",
                backgroundColor: "#FEF3C7",
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
            x: 20,
            y: 20,
            width: 240,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              padding: 10,
              backgroundColor: "#EEF2FF",
              borderColor: "#4F46E5",
              borderWidth: 2,
            },
          },
        ),
        text("Bottom center", {
          id: "centered",
          x: 280,
          y: 40,
          width: 160,
          height: 120,
          style: {
            fontFamily: jostFont,
            fontSize: 16,
            verticalAlign: "bottom",
            align: "center",
            backgroundColor: "#FEF3C7",
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

    assert.deepEqual(jsxPreview.pages[0]?.diagnostics, []);
    assert.deepEqual(impPreview.pages[0]?.diagnostics, []);
    assert.ok(jsxPreview.pages[0]?.png !== undefined);
    assert.deepEqual(jsxPreview.pages[0]?.png, impPreview.pages[0]?.png);
    expect(jsxPreview.pages[0]?.svg).toContain("<path");
    expect(jsxPreview.pages[0]?.svg).toContain('fill="#EEF2FF"');
    expect(jsxPreview.pages[0]?.svg).toContain("rotate(90");

    const centeredNode = jsxPreview.pages[0]?.nodes.find(
      (node) => node.authorId === "centered",
    );
    expect(centeredNode?.screenBox?.width).toBeCloseTo(120);
    expect(centeredNode?.screenBox?.height).toBeCloseTo(160);
    assert.equal(
      jsxPreview.query.getSlide("text-main").get("mixed").observe().text,
      "共享排版 shared 2026 第1页\nRight",
    );

    // Box visual effects (boxShadow, dashed border, borderRadius, gradient) render on Text with zero diagnostics.
    const styledTextDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "styled-text", background: "#FFFFFF" }).append(
        text("Boxed Text", {
          id: "styled",
          x: 20,
          y: 20,
          width: 180,
          height: 60,
          style: {
            fontFamily: jostFont,
            padding: 10,
            backgroundColor: "#EFF6FF",
            borderColor: "#111827",
            borderWidth: 2,
            borderStyle: "dashed",
            borderRadius: 10,
            boxShadow: { blur: 4, distance: 4, angle: 90 },
          },
        }),
      ),
    );
    const styledPreview = await previewDoc(styledTextDoc);
    assert.deepEqual((await validatePptx(styledPreview.pptx)).errors, []);
    assert.deepEqual(styledPreview.pages[0]?.diagnostics, []);
    assert.ok(styledPreview.pages[0]?.png !== undefined);
    expect(styledPreview.pages[0]?.svg).toContain("<feDropShadow");
    expect(styledPreview.pages[0]?.svg).toContain("stroke-dasharray");
  });

  it("rejects missing glyphs/fonts and color tables while keeping valid sibling pages", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const jostBytes = readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
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
      readonly expected: RegExp;
    }[] = [
      {
        label: "missing-glyph",
        node: text("汉", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: jostFont },
        }),
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
            x: 10,
            y: 10,
            width: 160,
            height: 50,
            style: { fontFamily: jostFont },
          },
        ),
        expected: /glyph/i,
      },
      {
        label: "missing-font-file",
        node: text("No font file", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: { fontFamily: font("Missing Preview Face") },
        }),
        expected: /font file/i,
      },
      ...(["COLR", "CPAL", "sbix", "CBDT", "SVG "] as const).map((table) => ({
        label: `${table.trim().toLowerCase()}-font`,
        node: text("Color font", {
          id: "bad",
          x: 10,
          y: 10,
          width: 160,
          height: 50,
          style: {
            fontFamily: font(
              `${table.trim()}Jost`,
              [{ weight: 400, data: withSyntheticTable(jostBytes, table) }],
              { substitute: true },
            ),
          },
        }),
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
              x: 10,
              y: 10,
              width: 120,
              style: { fontFamily: jostFont },
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
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /text\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
    }
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
            x: 20,
            y: 20,
            width: 240,
            height: 120,
            style: { fontFamily: jostFont, fontSize: 18 },
          },
        ),
        text("Rotated Shadow", {
          id: "rot-shadow",
          x: 280,
          y: 40,
          width: 180,
          height: 50,
          style: {
            fontFamily: jostFont,
            fontSize: 18,
            rotate: 90,
            textShadow: { color: "#000000", blur: 1, distance: 4, angle: 90 },
          },
        }),
        text("WIDE", {
          id: "wide-track",
          x: 20,
          y: 150,
          width: 200,
          height: "auto",
          style: { fontFamily: jostFont, fontSize: 18, letterSpacing: 6 },
        }),
        text("TIGHT", {
          id: "tight-track",
          x: 240,
          y: 150,
          width: 200,
          height: "auto",
          style: { fontFamily: jostFont, fontSize: 18, letterSpacing: -2 },
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
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
          x: 20,
          y: 20,
          width: 240,
          height: 50,
          style: { fontFamily: notoFont, fontSize: 18, color: "#0F172A" },
        }),
        text("مرحبا بالعالم 2026!", {
          id: "arabic-bidi",
          x: 20,
          y: 80,
          width: 260,
          height: 50,
          style: { fontFamily: notoArabicFont, fontSize: 20, color: "#0F172A" },
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
            x: 20,
            y: 140,
            width: 300,
            height: 50,
            style: {
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
            x: 20,
            y: 200,
            width: 220,
            height: 45,
            style: { fontFamily: notoArabicFont, fontSize: 20 },
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
            x: 280,
            y: 200,
            width: 220,
            height: 45,
            style: { fontSize: 18 },
          },
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
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
            x: 20,
            y: 20,
            width: 170,
            height: 180,
            style: { fontFamily: jostFont, fontSize: 16, padding: 10 },
          },
        ),
        text("汉字两端对齐自动折行测试末行", {
          id: "cjk-justify",
          x: 210,
          y: 20,
          width: 140,
          height: 140,
          style: {
            fontFamily: notoFont,
            fontSize: 16,
            align: "justify",
            padding: 10,
            color: "#047857",
          },
        }),
        text("One two three four five six", {
          id: "ul-justify",
          x: 370,
          y: 20,
          width: 130,
          height: 100,
          style: {
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
    assert.deepEqual(page.diagnostics, []);
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
            x: 20,
            y: 20,
            width: 190,
            height: "auto",
            style: { fontFamily: jostFont, fontSize: 16, padding: 8 },
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
            x: 240,
            y: 20,
            width: 240,
            height: "auto",
            style: { fontFamily: jostFont, fontSize: 16, padding: 8 },
          },
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
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
});
