import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  font,
  paragraph,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";

const jostRegularBytes = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: jostRegularBytes,
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
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
    data: readFileSync(
      new URL(
        "../../../test/assets/fonts/NotoNaskhArabic.ttf",
        import.meta.url,
      ),
    ),
  },
]);

const notoHebrewFont = font("Noto Sans Hebrew Thin", [
  {
    weight: 100,
    data: readFileSync(
      new URL("../../../test/assets/fonts/NotoSansHebrew.ttf", import.meta.url),
    ),
  },
]);

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
}

const colrJostFont = font(
  "ColrJost",
  [{ weight: 400, data: withSyntheticFontTable(jostRegularBytes, "COLR") }],
  { substitute: true },
);

const sbixJostFont = font(
  "SbixJost",
  [{ weight: 400, data: withSyntheticFontTable(jostRegularBytes, "sbix") }],
  { substitute: true },
);

export const textPositiveScenario: CapabilityScenario = {
  id: "text-typography-and-alignment",
  capability: "text",
  description:
    "Dual-author Text scenario covering multi-script & BiDi shaping, CJK/Latin wrapping, justify, bullet & numbered lists with hanging indent, underline/line-through, textShadow, letterSpacing, padding, vertical alignment, slide numbers, opacity, rotation, and container visual styles.",
  visualCheckpoints: [
    "Page 1 (cjk-latin-wrapping): light slate #F8FAFC canvas; top-left wrapped CJK + bold uppercase Latin; top-right middle/right-aligned card with slide number '1'; bottom-left Kana/Hangul/Cyrillic/Greek/math box; bottom-right Arabic & Hebrew BiDi cards with mixed LTR digits/Latin.",
    "Page 2 (vertical-and-rotation): warm #FFFBEB canvas; top-left top-aligned card; bottom-left bottom-aligned card with slide number '2'; top-right rotated +12 tag; bottom-right rounded dashed-border card showing underline, line-through, textShadow, and positive/negative letterSpacing.",
    "Page 3 (justify-and-lists): crisp white canvas; left card shows Latin and CJK justified paragraphs with flush edges on soft wraps and natural last/hard-break lines; right card shows multi-line bullet hanging indent, nested custom bullet, and sequential/restarted numbered lists.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all Text boxes are editable <p:sp> txBox='1' objects with native <a:fld type='slidenum'>, underline/sngStrike, spc, algn='just', <a:buChar>, <a:buAutoNum>, and rtl='1' attributes.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "cjk-latin-wrapping", background: "#F8FAFC" },
        createElement(
          Text,
          {
            id: "mixed-wrap",
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 280,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 12,
              background: "#FFFFFF",
              border: { color: "#CBD5E1", width: 1 },
            },
          },
          createElement(
            P,
            { style: { lineHeight: 28, spacingAfter: 10 } },
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
              "shared layout 2026",
            ),
            "，检查中英文混排与自动换行。",
          ),
          createElement(
            P,
            { style: { lineHeight: 26, align: "center" } },
            "第二段：居中对齐\n",
            createElement(
              Span,
              { style: { fontFamily: jostFont, color: "#475569" } },
              "Centered line",
            ),
          ),
        ),
        createElement(
          Text,
          {
            id: "fixed-align",
            style: {
              position: "absolute",
              left: 332,
              top: 20,
              width: 280,
              height: 150,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#1E3A8A",
              padding: [10, 14],
              verticalAlign: "middle",
              align: "right",
              lineHeight: 26,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 2 },
            },
          },
          "第",
          createElement(SlideNumber, null),
          "页\n",
          createElement(
            Span,
            { style: { fontFamily: jostFont } },
            "Right middle aligned",
          ),
        ),
        createElement(
          Text,
          {
            id: "multiscript-symbols",
            style: {
              position: "absolute",
              left: 24,
              top: 220,
              width: 280,
              height: 100,
              fontFamily: notoFont,
              fontSize: 17,
              color: "#0F172A",
              padding: 10,
              background: "#FFFFFF",
              border: { color: "#CBD5E1", width: 1 },
              borderRadius: 8,
            },
          },
          "あア 한 Д Ω ∑√≠",
        ),
        createElement(
          Text,
          {
            id: "bidi-mixed",
            style: {
              position: "absolute",
              left: 332,
              top: 190,
              width: 280,
              height: 130,
              fontFamily: notoArabicFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 10,
              background: "#F0FDF4",
              border: { color: "#10B981", width: 1 },
            },
          },
          createElement(
            P,
            { style: { lineHeight: 28, spacingAfter: 6 } },
            "مرحبا بالعالم 2026!",
          ),
          createElement(
            P,
            { style: { lineHeight: 28 } },
            createElement(
              Span,
              { style: { fontFamily: notoHebrewFont, fontWeight: 100 } },
              "שלום ",
            ),
            createElement(
              Span,
              { style: { fontFamily: jostFont, color: "#1D4ED8" } },
              "2026 ABC",
            ),
            createElement(
              Span,
              { style: { fontFamily: notoHebrewFont, fontWeight: 100 } },
              " עולם",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "vertical-and-rotation", background: "#FFFBEB" },
        createElement(
          Text,
          {
            id: "top-aligned",
            style: {
              position: "absolute",
              left: 32,
              top: 24,
              width: 250,
              height: 110,
              fontFamily: jostFont,
              fontSize: 18,
              color: "#78350F",
              padding: 10,
              verticalAlign: "top",
              background: "#FFFFFF",
              border: { color: "#D97706", width: 1 },
            },
          },
          "Top aligned box",
        ),
        createElement(
          Text,
          {
            id: "bottom-aligned",
            style: {
              position: "absolute",
              left: 32,
              top: 156,
              width: 250,
              height: 140,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#92400E",
              padding: 10,
              verticalAlign: "bottom",
              align: "center",
              background: "#FEF3C7",
              border: { color: "#D97706", width: 1 },
            },
          },
          createElement(Span, { style: { fontFamily: jostFont } }, "Bottom "),
          "页码 ",
          createElement(SlideNumber, null),
        ),
        createElement(
          Text,
          {
            id: "rotated-tag",
            style: {
              position: "absolute",
              left: 340,
              top: 34,
              width: 220,
              height: 70,
              fontFamily: jostFont,
              fontSize: 20,
              fontWeight: "bold",
              color: "#FFFFFF",
              verticalAlign: "middle",
              align: "center",
              background: "#1E293B",
              opacity: 0.85,
              rotate: 12,
              textShadow: { color: "#000000", blur: 2, distance: 3, angle: 90 },
            },
          },
          "Rotated 12",
        ),
        createElement(
          Text,
          {
            id: "decorations-card",
            style: {
              position: "absolute",
              left: 320,
              top: 136,
              width: 280,
              height: 160,
              fontFamily: jostFont,
              fontSize: 18,
              padding: 12,
              background: "#FFFFFF",
              border: { color: "#6366F1", width: 2, style: "dashed" },
              borderRadius: 12,
              boxShadow: { color: "#1E293B", blur: 4, distance: 3, angle: 90 },
            },
          },
          createElement(
            P,
            { style: { lineHeight: 28, spacingAfter: 6 } },
            createElement(
              Span,
              { style: { textDecoration: "underline", color: "#1D4ED8" } },
              "Underline ",
            ),
            createElement(
              Span,
              { style: { textDecoration: "line-through", color: "#DC2626" } },
              "Strike",
            ),
          ),
          createElement(
            P,
            { style: { lineHeight: 28, spacingAfter: 6 } },
            createElement(
              Span,
              {
                style: {
                  textShadow: {
                    color: "#0F172A",
                    blur: 2,
                    distance: 2,
                    angle: 45,
                  },
                },
              },
              "Text Shadow",
            ),
          ),
          createElement(
            P,
            { style: { lineHeight: 26 } },
            createElement(
              Span,
              { style: { letterSpacing: 3, color: "#047857" } },
              "WIDE ",
            ),
            createElement(
              Span,
              { style: { letterSpacing: -1.5, color: "#7C3AED" } },
              "TIGHT",
            ),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "justify-and-lists", background: "#FFFFFF" },
        createElement(
          Text,
          {
            id: "justify-box",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 270,
              height: 290,
              fontFamily: notoFont,
              fontSize: 16,
              padding: 12,
              background: "#F8FAFC",
              border: { color: "#94A3B8", width: 1 },
            },
          },
          createElement(
            P,
            {
              style: {
                fontFamily: jostFont,
                align: "justify",
                lineHeight: 24,
                spacingAfter: 10,
              },
            },
            "Alpha beta gamma delta epsilon zeta\nMandatory break line here",
          ),
          createElement(
            P,
            { style: { align: "justify", lineHeight: 24 } },
            "汉字两端对齐自动折行测试第一行与末行",
          ),
        ),
        createElement(
          Text,
          {
            id: "list-box",
            style: {
              position: "absolute",
              left: 324,
              top: 24,
              width: 280,
              height: 290,
              fontFamily: jostFont,
              fontSize: 16,
              padding: 12,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 1 },
            },
          },
          createElement(
            P,
            {
              style: {
                list: { type: "bullet" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            },
            "Wrapped bullet item showing hanging indent alignment",
          ),
          createElement(
            P,
            {
              style: {
                list: { type: "bullet", character: "+", level: 1 },
                lineHeight: 24,
                spacingAfter: 8,
              },
            },
            "Nested custom bullet",
          ),
          createElement(
            P,
            {
              style: {
                list: { type: "number" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            },
            "First auto item",
          ),
          createElement(
            P,
            {
              style: {
                list: { type: "number" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            },
            "Second auto item",
          ),
          createElement(
            P,
            {
              style: {
                list: { type: "number", startAt: 5 },
                lineHeight: 24,
                spacingAfter: 4,
              },
            },
            "Restarted at five",
          ),
          createElement(
            P,
            { style: { list: { type: "number", startAt: 5 }, lineHeight: 24 } },
            "Follows as six",
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "cjk-latin-wrapping", background: "#F8FAFC" }).append(
        text(
          [
            paragraph(
              [
                "共享排版 ",
                span("shared layout 2026", {
                  style: {
                    fontFamily: jostFont,
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    textTransform: "uppercase",
                  },
                }),
                "，检查中英文混排与自动换行。",
              ],
              { style: { lineHeight: 28, spacingAfter: 10 } },
            ),
            paragraph(
              [
                "第二段：居中对齐\n",
                span("Centered line", {
                  style: { fontFamily: jostFont, color: "#475569" },
                }),
              ],
              { style: { lineHeight: 26, align: "center" } },
            ),
          ],
          {
            id: "mixed-wrap",
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 280,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 12,
              background: "#FFFFFF",
              border: { color: "#CBD5E1", width: 1 },
            },
          },
        ),
        text(
          [
            "第",
            slideNumber(),
            "页\n",
            span("Right middle aligned", { style: { fontFamily: jostFont } }),
          ],
          {
            id: "fixed-align",
            style: {
              position: "absolute",
              left: 332,
              top: 20,
              width: 280,
              height: 150,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#1E3A8A",
              padding: [10, 14],
              verticalAlign: "middle",
              align: "right",
              lineHeight: 26,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 2 },
            },
          },
        ),
        text("あア 한 Д Ω ∑√≠", {
          id: "multiscript-symbols",
          style: {
            position: "absolute",
            left: 24,
            top: 220,
            width: 280,
            height: 100,
            fontFamily: notoFont,
            fontSize: 17,
            color: "#0F172A",
            padding: 10,
            background: "#FFFFFF",
            border: { color: "#CBD5E1", width: 1 },
            borderRadius: 8,
          },
        }),
        text(
          [
            paragraph("مرحبا بالعالم 2026!", {
              style: { lineHeight: 28, spacingAfter: 6 },
            }),
            paragraph(
              [
                span("שלום ", {
                  style: { fontFamily: notoHebrewFont, fontWeight: 100 },
                }),
                span("2026 ABC", {
                  style: { fontFamily: jostFont, color: "#1D4ED8" },
                }),
                span(" עולם", {
                  style: { fontFamily: notoHebrewFont, fontWeight: 100 },
                }),
              ],
              { style: { lineHeight: 28 } },
            ),
          ],
          {
            id: "bidi-mixed",
            style: {
              position: "absolute",
              left: 332,
              top: 190,
              width: 280,
              height: 130,
              fontFamily: notoArabicFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 10,
              background: "#F0FDF4",
              border: { color: "#10B981", width: 1 },
            },
          },
        ),
      ),
      slide({ id: "vertical-and-rotation", background: "#FFFBEB" }).append(
        text("Top aligned box", {
          id: "top-aligned",
          style: {
            position: "absolute",
            left: 32,
            top: 24,
            width: 250,
            height: 110,
            fontFamily: jostFont,
            fontSize: 18,
            color: "#78350F",
            padding: 10,
            verticalAlign: "top",
            background: "#FFFFFF",
            border: { color: "#D97706", width: 1 },
          },
        }),
        text(
          [
            span("Bottom ", { style: { fontFamily: jostFont } }),
            "页码 ",
            slideNumber(),
          ],
          {
            id: "bottom-aligned",
            style: {
              position: "absolute",
              left: 32,
              top: 156,
              width: 250,
              height: 140,
              fontFamily: notoFont,
              fontSize: 18,
              color: "#92400E",
              padding: 10,
              verticalAlign: "bottom",
              align: "center",
              background: "#FEF3C7",
              border: { color: "#D97706", width: 1 },
            },
          },
        ),
        text("Rotated 12", {
          id: "rotated-tag",
          style: {
            position: "absolute",
            left: 340,
            top: 34,
            width: 220,
            height: 70,
            fontFamily: jostFont,
            fontSize: 20,
            fontWeight: "bold",
            color: "#FFFFFF",
            verticalAlign: "middle",
            align: "center",
            background: "#1E293B",
            opacity: 0.85,
            rotate: 12,
            textShadow: { color: "#000000", blur: 2, distance: 3, angle: 90 },
          },
        }),
        text(
          [
            paragraph(
              [
                span("Underline ", {
                  style: { textDecoration: "underline", color: "#1D4ED8" },
                }),
                span("Strike", {
                  style: { textDecoration: "line-through", color: "#DC2626" },
                }),
              ],
              { style: { lineHeight: 28, spacingAfter: 6 } },
            ),
            paragraph(
              [
                span("Text Shadow", {
                  style: {
                    textShadow: {
                      color: "#0F172A",
                      blur: 2,
                      distance: 2,
                      angle: 45,
                    },
                  },
                }),
              ],
              { style: { lineHeight: 28, spacingAfter: 6 } },
            ),
            paragraph(
              [
                span("WIDE ", {
                  style: { letterSpacing: 3, color: "#047857" },
                }),
                span("TIGHT", {
                  style: { letterSpacing: -1.5, color: "#7C3AED" },
                }),
              ],
              { style: { lineHeight: 26 } },
            ),
          ],
          {
            id: "decorations-card",
            style: {
              position: "absolute",
              left: 320,
              top: 136,
              width: 280,
              height: 160,
              fontFamily: jostFont,
              fontSize: 18,
              padding: 12,
              background: "#FFFFFF",
              border: { color: "#6366F1", width: 2, style: "dashed" },
              borderRadius: 12,
              boxShadow: { color: "#1E293B", blur: 4, distance: 3, angle: 90 },
            },
          },
        ),
      ),
      slide({ id: "justify-and-lists", background: "#FFFFFF" }).append(
        text(
          [
            paragraph(
              "Alpha beta gamma delta epsilon zeta\nMandatory break line here",
              {
                style: {
                  fontFamily: jostFont,
                  align: "justify",
                  lineHeight: 24,
                  spacingAfter: 10,
                },
              },
            ),
            paragraph("汉字两端对齐自动折行测试第一行与末行", {
              style: { align: "justify", lineHeight: 24 },
            }),
          ],
          {
            id: "justify-box",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 270,
              height: 290,
              fontFamily: notoFont,
              fontSize: 16,
              padding: 12,
              background: "#F8FAFC",
              border: { color: "#94A3B8", width: 1 },
            },
          },
        ),
        text(
          [
            paragraph("Wrapped bullet item showing hanging indent alignment", {
              style: {
                list: { type: "bullet" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            }),
            paragraph("Nested custom bullet", {
              style: {
                list: { type: "bullet", character: "+", level: 1 },
                lineHeight: 24,
                spacingAfter: 8,
              },
            }),
            paragraph("First auto item", {
              style: {
                list: { type: "number" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            }),
            paragraph("Second auto item", {
              style: {
                list: { type: "number" },
                lineHeight: 24,
                spacingAfter: 4,
              },
            }),
            paragraph("Restarted at five", {
              style: {
                list: { type: "number", startAt: 5 },
                lineHeight: 24,
                spacingAfter: 4,
              },
            }),
            paragraph("Follows as six", {
              style: { list: { type: "number", startAt: 5 }, lineHeight: 24 },
            }),
          ],
          {
            id: "list-box",
            style: {
              position: "absolute",
              left: 324,
              top: 24,
              width: 280,
              height: 290,
              fontFamily: jostFont,
              fontSize: 16,
              padding: 12,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 1 },
            },
          },
        ),
      ),
    );
  },
  assert({ author, pptx, resolution, preview, validation }) {
    assert.equal(preview.pages.length, 3, author);
    assert.equal(validation.slides.length, 3, author);

    for (const page of preview.pages) {
      assert.deepEqual(page.errors, [], author);
      assert.ok(page.png !== undefined, author);
      assert.ok(page.svg?.includes("<path"), author);
    }

    const mixedLocator = resolution
      .slide("cjk-latin-wrapping")
      .get("mixed-wrap");
    assert.ok((mixedLocator.box?.height ?? 0) > 100, author);
    assert.equal(
      resolution.slide("cjk-latin-wrapping").get("fixed-align").text,
      "第1页\nRight middle aligned",
      author,
    );
    assert.equal(
      resolution.slide("vertical-and-rotation").get("bottom-aligned").text,
      "Bottom 页码 2",
      author,
    );
    const rotatedLocator = resolution
      .slide("vertical-and-rotation")
      .get("rotated-tag");
    assert.ok((rotatedLocator.bounds?.width ?? 0) > 220, author);
    assert.ok((rotatedLocator.bounds?.height ?? 0) > 70, author);

    const zip = unzipSync(pptx);
    const s2Xml = strFromU8(zip["ppt/slides/slide2.xml"]!);
    assert.ok(s2Xml.includes('u="sng"'), author);
    assert.ok(s2Xml.includes('strike="sngStrike"'), author);
    assert.ok(s2Xml.includes('spc="225"'), author);
    assert.ok(s2Xml.includes('spc="-112"'), author);

    const s3Xml = strFromU8(zip["ppt/slides/slide3.xml"]!);
    assert.ok(s3Xml.includes('algn="just"'), author);
    assert.ok(s3Xml.includes("<a:buChar"), author);
    assert.ok(s3Xml.includes("<a:buAutoNum"), author);

    assert.ok(preview.pages[1]?.svg?.includes("<feDropShadow"), author);
    assert.ok(preview.pages[1]?.svg?.includes("<rect "), author);
  },
};

export const textRejectionScenario: CapabilityScenario = {
  id: "text-preview-rejections",
  capability: "text",
  kind: "rejection",
  description:
    "Dual-author negative scenario verifying Text preview enforces only the two underlying font resource boundaries (color/bitmap font tables and missing glyph) while retaining valid PPTX output and sibling pages.",
  visualCheckpoints: [
    "Page 1 (text-rejections): preview omitted (`png: undefined`) with located error diagnostics for missing glyph (body & list bullet) and COLR/sbix color/bitmap font tables.",
    "Page 2 (valid-sibling): crisp white canvas with blue-bordered valid text box.",
  ],
  nativeCheckpoints: [
    "Verify exported PPTX retains all Page 1 native text objects and passes OOXML validation.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "text-rejections" },
        createElement(
          Text,
          {
            id: "missing-glyph",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 180,
              height: 40,
              fontFamily: jostFont,
            },
          },
          "汉",
        ),
        createElement(
          Text,
          {
            id: "missing-bullet-glyph",
            style: {
              position: "absolute",
              left: 10,
              top: 60,
              width: 180,
              height: 40,
              fontFamily: jostFont,
            },
          },
          createElement(
            P,
            { style: { list: { type: "bullet", character: "汉" } } },
            "Item",
          ),
        ),
        createElement(
          Text,
          {
            id: "colr-font",
            style: {
              position: "absolute",
              left: 220,
              top: 10,
              width: 180,
              height: 40,
              fontFamily: colrJostFont,
            },
          },
          "Color",
        ),
        createElement(
          Text,
          {
            id: "sbix-font",
            style: {
              position: "absolute",
              left: 220,
              top: 60,
              width: 180,
              height: 40,
              fontFamily: sbixJostFont,
            },
          },
          "Bitmap",
        ),
      ),
      createElement(
        Slide,
        { id: "valid-sibling" },
        createElement(
          Text,
          {
            id: "ok-text",
            style: {
              position: "absolute",
              left: 40,
              top: 60,
              width: 240,
              height: 60,
              fontFamily: jostFont,
              fontSize: 22,
              padding: 12,
              verticalAlign: "middle",
              background: "#EFF6FF",
              border: { color: "#2563EB", width: 2 },
            },
          },
          "Valid sibling text",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "text-rejections" }).append(
        text("汉", {
          id: "missing-glyph",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 180,
            height: 40,
            fontFamily: jostFont,
          },
        }),
        text(
          [
            paragraph("Item", {
              style: { list: { type: "bullet", character: "汉" } },
            }),
          ],
          {
            id: "missing-bullet-glyph",
            style: {
              position: "absolute",
              left: 10,
              top: 60,
              width: 180,
              height: 40,
              fontFamily: jostFont,
            },
          },
        ),
        text("Color", {
          id: "colr-font",
          style: {
            position: "absolute",
            left: 220,
            top: 10,
            width: 180,
            height: 40,
            fontFamily: colrJostFont,
          },
        }),
        text("Bitmap", {
          id: "sbix-font",
          style: {
            position: "absolute",
            left: 220,
            top: 60,
            width: 180,
            height: 40,
            fontFamily: sbixJostFont,
          },
        }),
      ),
      slide({ id: "valid-sibling" }).append(
        text("Valid sibling text", {
          id: "ok-text",
          style: {
            position: "absolute",
            left: 40,
            top: 60,
            width: 240,
            height: 60,
            fontFamily: jostFont,
            fontSize: 22,
            padding: 12,
            verticalAlign: "middle",
            background: "#EFF6FF",
            border: { color: "#2563EB", width: 2 },
          },
        }),
      ),
    );
  },
  assert({ author, preview }) {
    const p1 = preview.pages[0]!;
    assert.equal(p1.svg, undefined, author);
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.errors.length, 4, author);
    for (const [id, expectedCode] of [
      ["missing-glyph", "preview-missing-glyph"],
      ["missing-bullet-glyph", "preview-missing-glyph"],
      ["colr-font", "preview-unsupported-font"],
      ["sbix-font", "preview-unsupported-font"],
    ] as const) {
      assert.ok(
        p1.errors.some(
          (err) =>
            err.code === expectedCode && err.path.includes(`text[${id}]`),
        ),
        `${author}: missing ${expectedCode} error for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.errors, [], author);
    assert.ok(p2.png !== undefined, author);
  },
};

const textFillPngBytes = (() => {
  const img = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < img.data.length; i += 4) {
    img.data.set([0x1e, 0x40, 0xaf, 255], i);
  }
  return new Uint8Array(PNG.sync.write(img));
})();

export const textRtlListsAndFillsScenario: CapabilityScenario = {
  id: "text-rtl-lists-image-fill-and-percent-height",
  capability: "text",
  description:
    "Dual-author Text scenario covering RTL bullet/numbered lists with right/center/justify alignment, Arabic numerals and diacritics with lang='ar-SA', image-filled Text box, 100% height inside fixed Box, underline+textShadow on list items, and multi-level numbered list fallback.",
  visualCheckpoints: [
    "Page 1 (rtl-lists-and-image-fill): RTL Arabic/Hebrew bullet and numbered lists aligned right/center/justify, plus an image-filled Text box and multi-level numbered list with shadow+underline.",
  ],
  nativeCheckpoints: [
    "Verify RTL list XML, image relationship on Text shape, and multi-level numbering in PPTX.",
  ],
  jsx() {
    const rtlJustifyCopy =
      "مَرْحَبًا بِكُمْ فِي اخْتِبَارِ ١٢٣ القَائِمَةِ العَرَبِيَّةِ المُنَسَّقَةِ مَعَ الضَّبْطِ";
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "rtl-lists-and-image-fill", background: "#FFFFFF" },
        createElement(
          Text,
          {
            id: "rtl-lists",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 300,
              height: 220,
              fontFamily: notoArabicFont,
              fontSize: 15,
              color: "#0F172A",
              padding: 10,
              background: "#F8FAFC",
              border: { color: "#94A3B8", width: 1 },
            },
          },
          createElement(
            P,
            {
              style: {
                align: "right",
                lang: "ar-SA",
                list: { type: "bullet", character: "-" },
                lineHeight: 24,
              },
            },
            "البند الأول ١٢٣",
          ),
          createElement(
            P,
            {
              style: {
                align: "center",
                lang: "ar-SA",
                list: { type: "number", startAt: 1 },
                lineHeight: 24,
              },
            },
            "البند الثاني",
          ),
          createElement(
            P,
            {
              style: {
                align: "justify",
                lang: "ar-SA",
                list: { type: "bullet", character: "-" },
                lineHeight: 24,
              },
            },
            rtlJustifyCopy,
          ),
          createElement(
            P,
            {
              style: {
                align: "right",
                list: { type: "number", startAt: 1 },
                lineHeight: 24,
              },
            },
            createElement(
              Span,
              { style: { fontFamily: notoHebrewFont, fontWeight: 100 } },
              "שורה ראשונה",
            ),
          ),
        ),
        createElement(
          Text,
          {
            id: "image-filled-text",
            style: {
              position: "absolute",
              left: 340,
              top: 20,
              width: 280,
              height: 70,
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
              padding: 10,
              background: { image: textFillPngBytes },
            },
          },
          "Image Fill Text",
        ),
        createElement(
          Box,
          {
            id: "pct-holder",
            style: {
              position: "absolute",
              left: 340,
              top: 105,
              width: 280,
              height: 150,
              background: "#EFF6FF",
              border: { color: "#3B82F6", width: 1 },
            },
          },
          createElement(
            Text,
            {
              id: "pct-height-text",
              style: {
                width: "100%",
                height: "100%",
                fontFamily: jostFont,
                fontSize: 13,
                color: "#1E293B",
                padding: 8,
              },
            },
            createElement(
              P,
              {
                style: {
                  list: { type: "number", level: 0, startAt: 1 },
                  textDecoration: "underline",
                  textShadow: {
                    color: "#0F172A",
                    blur: 1,
                    distance: 1,
                    angle: 90,
                  },
                },
              },
              "Level 0 item 1",
            ),
            createElement(
              P,
              {
                style: {
                  list: { type: "number", level: 1, startAt: 1 },
                },
              },
              "Level 1 subitem",
            ),
            createElement(
              P,
              {
                style: {
                  list: { type: "number", level: 0, startAt: 1 },
                },
              },
              "Level 0 item 2",
            ),
          ),
        ),
      ),
    );
  },
  document() {
    const rtlJustifyCopy =
      "مَرْحَبًا بِكُمْ فِي اخْتِبَارِ ١٢٣ القَائِمَةِ العَرَبِيَّةِ المُنَسَّقَةِ مَعَ الضَّبْطِ";
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "rtl-lists-and-image-fill", background: "#FFFFFF" }).append(
        text(
          [
            paragraph("البند الأول ١٢٣", {
              style: {
                align: "right",
                lang: "ar-SA",
                list: { type: "bullet", character: "-" },
                lineHeight: 24,
              },
            }),
            paragraph("البند الثاني", {
              style: {
                align: "center",
                lang: "ar-SA",
                list: { type: "number", startAt: 1 },
                lineHeight: 24,
              },
            }),
            paragraph(rtlJustifyCopy, {
              style: {
                align: "justify",
                lang: "ar-SA",
                list: { type: "bullet", character: "-" },
                lineHeight: 24,
              },
            }),
            paragraph(
              [
                span("שורה ראשונה", {
                  style: { fontFamily: notoHebrewFont, fontWeight: 100 },
                }),
              ],
              {
                style: {
                  align: "right",
                  list: { type: "number", startAt: 1 },
                  lineHeight: 24,
                },
              },
            ),
          ],
          {
            id: "rtl-lists",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 300,
              height: 220,
              fontFamily: notoArabicFont,
              fontSize: 15,
              color: "#0F172A",
              padding: 10,
              background: "#F8FAFC",
              border: { color: "#94A3B8", width: 1 },
            },
          },
        ),
        text("Image Fill Text", {
          id: "image-filled-text",
          style: {
            position: "absolute",
            left: 340,
            top: 20,
            width: 280,
            height: 70,
            fontFamily: jostFont,
            fontSize: 16,
            color: "#FFFFFF",
            padding: 10,
            background: { image: textFillPngBytes },
          },
        }),
        box({
          id: "pct-holder",
          style: {
            position: "absolute",
            left: 340,
            top: 105,
            width: 280,
            height: 150,
            background: "#EFF6FF",
            border: { color: "#3B82F6", width: 1 },
          },
        }).append(
          text(
            [
              paragraph("Level 0 item 1", {
                style: {
                  list: { type: "number", level: 0, startAt: 1 },
                  textDecoration: "underline",
                  textShadow: {
                    color: "#0F172A",
                    blur: 1,
                    distance: 1,
                    angle: 90,
                  },
                },
              }),
              paragraph("Level 1 subitem", {
                style: {
                  list: { type: "number", level: 1, startAt: 1 },
                },
              }),
              paragraph("Level 0 item 2", {
                style: {
                  list: { type: "number", level: 0, startAt: 1 },
                },
              }),
            ],
            {
              id: "pct-height-text",
              style: {
                width: "100%",
                height: "100%",
                fontFamily: jostFont,
                fontSize: 13,
                color: "#1E293B",
                padding: 8,
              },
            },
          ),
        ),
      ),
    );
  },
  assert({ author, pptx, preview }) {
    assert.equal(preview.pages.length, 1, author);
    const page = preview.pages[0]!;
    assert.deepEqual(page.errors, [], author);
    assert.ok(page.png !== undefined, author);
    const svg = page.svg ?? "";
    assert.ok(svg.includes("<pattern"), author);
    assert.ok(svg.includes("<feDropShadow"), author);
    const zip = unzipSync(pptx);
    const s1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
    assert.ok(s1Xml.includes("<a:blipFill"), author);
    assert.ok(s1Xml.includes('lang="ar-SA"'), author);
  },
};
