import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Deck, Slide } from "@react-pptx/core";
import { createElement } from "react";

import {
  deck,
  font,
  paragraph,
  slide,
  slideNumber,
  span,
  text,
} from "../../src/document.ts";
import { P, SlideNumber, Span, Text } from "../../src/text-elements.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";

const jostRegularBytes = readFileSync(
  new URL(
    "../../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
    import.meta.url,
  ),
);

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: jostRegularBytes,
  },
  {
    weight: 700,
    data: readFileSync(
      new URL(
        "../../../react-pptx/test/assets/fonts/Jost-700-Bold.ttf",
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
        "../../../react-pptx/test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

function withSyntheticColrTable(source: Buffer): Uint8Array {
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
    tag: "COLR",
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
  [{ weight: 400, data: withSyntheticColrTable(jostRegularBytes) }],
  { substitute: true },
);

export const textPositiveScenario: CapabilityScenario = {
  id: "text-typography-and-alignment",
  capability: "rp:text",
  description:
    "Dual-author Text scenario covering CJK/Latin wrapping, padding, horizontal/vertical alignment, uppercase, slide numbers, content-sized vs fixed height, opacity, and rotation.",
  visualCheckpoints: [
    "Page 1 (cjk-latin-wrapping): light slate #F8FAFC canvas; left white card with 1px #CBD5E1 border contains wrapped CJK + bold uppercase blue Latin 'SHARED LAYOUT 2026' and centered second paragraph; right blue-tinted #EFF6FF card with 2px #3B82F6 border shows middle/right-aligned text including slide number '1'.",
    "Page 2 (vertical-and-rotation): warm #FFFBEB canvas; top-left white bordered card aligns text at top; bottom-left #FEF3C7 bordered card aligns centered text with slide number '2' at bottom; right dark slate #1E293B semi-transparent card rotated +12 degrees displays crisp centered white text.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify all Text boxes are editable <p:sp> txBox='1' objects with native <a:fld type='slidenum'> slide numbers and matching wrapping/alignment.",
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
            x: 24,
            y: 28,
            width: 280,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 12,
              backgroundColor: "#FFFFFF",
              borderColor: "#CBD5E1",
              borderWidth: 1,
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
            x: 332,
            y: 28,
            width: 280,
            height: 180,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#1E3A8A",
              padding: [10, 14],
              verticalAlign: "middle",
              align: "right",
              lineHeight: 26,
              backgroundColor: "#EFF6FF",
              borderColor: "#3B82F6",
              borderWidth: 2,
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
      ),
      createElement(
        Slide,
        { id: "vertical-and-rotation", background: "#FFFBEB" },
        createElement(
          Text,
          {
            id: "top-aligned",
            x: 32,
            y: 24,
            width: 250,
            height: 110,
            style: {
              fontFamily: jostFont,
              fontSize: 18,
              color: "#78350F",
              padding: 10,
              verticalAlign: "top",
              backgroundColor: "#FFFFFF",
              borderColor: "#D97706",
              borderWidth: 1,
            },
          },
          "Top aligned box",
        ),
        createElement(
          Text,
          {
            id: "bottom-aligned",
            x: 32,
            y: 156,
            width: 250,
            height: 140,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#92400E",
              padding: 10,
              verticalAlign: "bottom",
              align: "center",
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 1,
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
            x: 340,
            y: 100,
            width: 220,
            height: 80,
            style: {
              fontFamily: jostFont,
              fontSize: 20,
              fontWeight: "bold",
              color: "#FFFFFF",
              verticalAlign: "middle",
              align: "center",
              backgroundColor: "#1E293B",
              opacity: 0.85,
              rotate: 12,
            },
          },
          "Rotated 12",
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
            x: 24,
            y: 28,
            width: 280,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#0F172A",
              padding: 12,
              backgroundColor: "#FFFFFF",
              borderColor: "#CBD5E1",
              borderWidth: 1,
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
            x: 332,
            y: 28,
            width: 280,
            height: 180,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#1E3A8A",
              padding: [10, 14],
              verticalAlign: "middle",
              align: "right",
              lineHeight: 26,
              backgroundColor: "#EFF6FF",
              borderColor: "#3B82F6",
              borderWidth: 2,
            },
          },
        ),
      ),
      slide({ id: "vertical-and-rotation", background: "#FFFBEB" }).append(
        text("Top aligned box", {
          id: "top-aligned",
          x: 32,
          y: 24,
          width: 250,
          height: 110,
          style: {
            fontFamily: jostFont,
            fontSize: 18,
            color: "#78350F",
            padding: 10,
            verticalAlign: "top",
            backgroundColor: "#FFFFFF",
            borderColor: "#D97706",
            borderWidth: 1,
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
            x: 32,
            y: 156,
            width: 250,
            height: 140,
            style: {
              fontFamily: notoFont,
              fontSize: 18,
              color: "#92400E",
              padding: 10,
              verticalAlign: "bottom",
              align: "center",
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 1,
            },
          },
        ),
        text("Rotated 12", {
          id: "rotated-tag",
          x: 340,
          y: 100,
          width: 220,
          height: 80,
          style: {
            fontFamily: jostFont,
            fontSize: 20,
            fontWeight: "bold",
            color: "#FFFFFF",
            verticalAlign: "middle",
            align: "center",
            backgroundColor: "#1E293B",
            opacity: 0.85,
            rotate: 12,
          },
        }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(preview.pages.length, 2, author);
    assert.equal(validation.slides.length, 2, author);

    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, [], author);
      assert.ok(page.png !== undefined, author);
      assert.ok(page.svg?.includes("<path"), author);
    }

    const mixedNode = preview.pages[0]?.nodes.find(
      (node) => node.authorId === "mixed-wrap",
    );
    assert.ok((mixedNode?.box?.height ?? 0) > 100, author);
    assert.equal(
      preview.build.query
        .getSlide("cjk-latin-wrapping")
        .get("fixed-align")
        .observe().text,
      "第1页\nRight middle aligned",
      author,
    );
    assert.equal(
      preview.build.query
        .getSlide("vertical-and-rotation")
        .get("bottom-aligned")
        .observe().text,
      "Bottom 页码 2",
      author,
    );
    const rotatedNode = preview.pages[1]?.nodes.find(
      (node) => node.authorId === "rotated-tag",
    );
    assert.ok((rotatedNode?.screenBox?.width ?? 0) > 220, author);
    assert.ok((rotatedNode?.screenBox?.height ?? 0) > 80, author);
  },
};

export const textRejectionScenario: CapabilityScenario = {
  id: "text-preview-rejections",
  capability: "rp:text",
  description:
    "Dual-author negative scenario verifying Text preview rejects unsupported script, text styles, missing glyphs, COLR font tables, and unsupported box styles while keeping valid sibling pages.",
  visualCheckpoints: [
    "Page 1 (text-rejections): preview omitted (`png: undefined`) with located error diagnostics for all unsupported text/font declarations.",
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
            id: "arabic",
            x: 10,
            y: 10,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont },
          },
          "مرحبا",
        ),
        createElement(
          Text,
          {
            id: "justify",
            x: 10,
            y: 50,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont, align: "justify" },
          },
          "Justify",
        ),
        createElement(
          Text,
          {
            id: "bullet-list",
            x: 10,
            y: 90,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont },
          },
          createElement(P, { style: { list: { type: "bullet" } } }, "Item"),
        ),
        createElement(
          Text,
          {
            id: "underline",
            x: 10,
            y: 130,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont, textDecoration: "underline" },
          },
          "Underline",
        ),
        createElement(
          Text,
          {
            id: "shadow",
            x: 160,
            y: 10,
            width: 120,
            height: 30,
            style: {
              fontFamily: jostFont,
              textShadow: { blur: 2, distance: 1 },
            },
          },
          "Shadow",
        ),
        createElement(
          Text,
          {
            id: "letter-spacing",
            x: 160,
            y: 50,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont, letterSpacing: 2 },
          },
          "Tracking",
        ),
        createElement(
          Text,
          {
            id: "missing-glyph",
            x: 160,
            y: 90,
            width: 120,
            height: 30,
            style: { fontFamily: jostFont },
          },
          "汉",
        ),
        createElement(
          Text,
          {
            id: "colr-font",
            x: 160,
            y: 130,
            width: 120,
            height: 30,
            style: { fontFamily: colrJostFont },
          },
          "Color",
        ),
        createElement(
          Text,
          {
            id: "box-shadow",
            x: 310,
            y: 10,
            width: 120,
            height: 30,
            style: {
              fontFamily: jostFont,
              boxShadow: { blur: 4, distance: 2 },
            },
          },
          "Box",
        ),
      ),
      createElement(
        Slide,
        { id: "valid-sibling" },
        createElement(
          Text,
          {
            id: "ok-text",
            x: 40,
            y: 60,
            width: 240,
            height: 60,
            style: {
              fontFamily: jostFont,
              fontSize: 22,
              padding: 12,
              verticalAlign: "middle",
              backgroundColor: "#EFF6FF",
              borderColor: "#2563EB",
              borderWidth: 2,
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
        text("مرحبا", {
          id: "arabic",
          x: 10,
          y: 10,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont },
        }),
        text("Justify", {
          id: "justify",
          x: 10,
          y: 50,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont, align: "justify" },
        }),
        text([paragraph("Item", { style: { list: { type: "bullet" } } })], {
          id: "bullet-list",
          x: 10,
          y: 90,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont },
        }),
        text("Underline", {
          id: "underline",
          x: 10,
          y: 130,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont, textDecoration: "underline" },
        }),
        text("Shadow", {
          id: "shadow",
          x: 160,
          y: 10,
          width: 120,
          height: 30,
          style: {
            fontFamily: jostFont,
            textShadow: { blur: 2, distance: 1 },
          },
        }),
        text("Tracking", {
          id: "letter-spacing",
          x: 160,
          y: 50,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont, letterSpacing: 2 },
        }),
        text("汉", {
          id: "missing-glyph",
          x: 160,
          y: 90,
          width: 120,
          height: 30,
          style: { fontFamily: jostFont },
        }),
        text("Color", {
          id: "colr-font",
          x: 160,
          y: 130,
          width: 120,
          height: 30,
          style: { fontFamily: colrJostFont },
        }),
        text("Box", {
          id: "box-shadow",
          x: 310,
          y: 10,
          width: 120,
          height: 30,
          style: {
            fontFamily: jostFont,
            boxShadow: { blur: 4, distance: 2 },
          },
        }),
      ),
      slide({ id: "valid-sibling" }).append(
        text("Valid sibling text", {
          id: "ok-text",
          x: 40,
          y: 60,
          width: 240,
          height: 60,
          style: {
            fontFamily: jostFont,
            fontSize: 22,
            padding: 12,
            verticalAlign: "middle",
            backgroundColor: "#EFF6FF",
            borderColor: "#2563EB",
            borderWidth: 2,
          },
        }),
      ),
    );
  },
  assert({ author, preview }) {
    const p1 = preview.pages[0]!;
    assert.equal(p1.png, undefined, author);
    assert.equal(p1.diagnostics.length, 9, author);
    for (const id of [
      "arabic",
      "justify",
      "bullet-list",
      "underline",
      "shadow",
      "letter-spacing",
      "missing-glyph",
      "colr-font",
      "box-shadow",
    ]) {
      assert.ok(
        p1.diagnostics.some(
          (d) => d.severity === "error" && d.path.includes(`text[${id}]`),
        ),
        `${author}: missing diagnostic for ${id}`,
      );
    }

    const p2 = preview.pages[1]!;
    assert.deepEqual(p2.diagnostics, [], author);
    assert.ok(p2.png !== undefined, author);
  },
};
