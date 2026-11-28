import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  box,
  cell,
  chart,
  deck,
  defineSlideLayout,
  font,
  fromFontCollection,
  notes,
  paragraph,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  span,
  table,
  text,
} from "@pptxdk/capabilities/document";
import { createCodegen } from "@pptxdk/codegen";
import { BuildError, defineTheme, type Theme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { isFont } from "@pptxdk/core/capability";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

import {
  buildTestTtc,
  replaceSfntNameTable,
} from "../../../test/sfnt-test-helpers.ts";

const jostBookBytes = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const notoScWoff2Bytes = new Uint8Array(
  readFileSync(
    new URL(
      "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
      import.meta.url,
    ),
  ),
);

function createSyntheticCjkTtf(
  familyName: string,
  codePoints: readonly number[],
): Uint8Array {
  const named = replaceSfntNameTable(jostBookBytes, [
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 1,
      value: familyName,
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 2,
      value: "Regular",
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 4,
      value: `${familyName} Regular`,
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 5,
      value: "1.0",
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 6,
      value: familyName.replace(/\s+/gu, ""),
    },
  ]);
  const view = new DataView(named.buffer, named.byteOffset, named.byteLength);
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();
  for (let i = 0; i < numTables; i += 1) {
    const rec = 12 + i * 16;
    const tag = String.fromCharCode(...named.subarray(rec, rec + 4));
    const offset = view.getUint32(rec + 8);
    const length = view.getUint32(rec + 12);
    tables.set(tag, named.slice(offset, offset + length));
  }

  // Build a minimal cmap format 4 mapping each requested code point to glyph ID 5 (non-.notdef).
  const sortedCps = [...new Set(codePoints)].toSorted((a, b) => a - b);
  const segCount = sortedCps.length + 1;
  const segCountX2 = segCount * 2;
  const subtableLen = 16 + segCount * 8;
  const cmap = new Uint8Array(12 + subtableLen);
  const cmapView = new DataView(cmap.buffer);
  cmapView.setUint16(0, 0);
  cmapView.setUint16(2, 1);
  cmapView.setUint16(4, 3);
  cmapView.setUint16(6, 1);
  cmapView.setUint32(8, 12);
  cmapView.setUint16(12, 4);
  cmapView.setUint16(14, subtableLen);
  cmapView.setUint16(16, 0);
  cmapView.setUint16(18, segCountX2);
  const entrySelector = Math.floor(Math.log2(segCount));
  const searchRange = (1 << entrySelector) * 2;
  cmapView.setUint16(20, searchRange);
  cmapView.setUint16(22, entrySelector);
  cmapView.setUint16(24, segCountX2 - searchRange);
  for (let i = 0; i < sortedCps.length; i += 1) {
    const cp = sortedCps[i]!;
    cmapView.setUint16(26 + i * 2, cp); // endCode
    cmapView.setUint16(26 + segCount * 2 + 2 + i * 2, cp); // startCode
    cmapView.setInt16(26 + segCount * 4 + 2 + i * 2, (5 - cp) & 0xffff); // idDelta -> gid 5
    cmapView.setUint16(26 + segCount * 6 + 2 + i * 2, 0);
  }
  cmapView.setUint16(26 + sortedCps.length * 2, 0xffff);
  cmapView.setUint16(26 + segCount * 2 + 2 + sortedCps.length * 2, 0xffff);
  cmapView.setInt16(26 + segCount * 4 + 2 + sortedCps.length * 2, 1);
  tables.set("cmap", cmap);

  const sortedTables = [...tables.entries()].toSorted(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  const headerSize = 12 + sortedTables.length * 16;
  let totalSize = headerSize;
  for (const [, data] of sortedTables) {
    totalSize += Math.ceil(data.byteLength / 4) * 4;
  }
  const out = new Uint8Array(totalSize);
  const outView = new DataView(out.buffer);
  outView.setUint32(0, sfntVersion);
  outView.setUint16(4, sortedTables.length);
  const es = Math.floor(Math.log2(sortedTables.length));
  const sr = (1 << es) * 16;
  outView.setUint16(6, sr);
  outView.setUint16(8, es);
  outView.setUint16(10, sortedTables.length * 16 - sr);
  const sum = (buf: Uint8Array): number => {
    const padded = Math.ceil(buf.byteLength / 4) * 4;
    const bv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let acc = 0;
    for (let o = 0; o < padded; o += 4) {
      if (o + 4 <= buf.byteLength) {
        acc = (acc + bv.getUint32(o)) >>> 0;
      } else {
        let w = 0;
        for (let b = 0; b < 4; b += 1) {
          w = (w << 8) | (o + b < buf.byteLength ? buf[o + b]! : 0);
        }
        acc = (acc + (w >>> 0)) >>> 0;
      }
    }
    return acc >>> 0;
  };
  let dataOff = headerSize;
  let headOff = 0;
  for (let i = 0; i < sortedTables.length; i += 1) {
    const [tag, raw] = sortedTables[i]!;
    const data = tag === "head" ? new Uint8Array(raw) : raw;
    if (tag === "head") {
      new DataView(data.buffer).setUint32(8, 0);
      headOff = dataOff;
    }
    out.set(data, dataOff);
    const rec = 12 + i * 16;
    for (let k = 0; k < 4; k += 1) out[rec + k] = tag.charCodeAt(k);
    outView.setUint32(rec + 4, sum(data));
    outView.setUint32(rec + 8, dataOff);
    outView.setUint32(rec + 12, data.byteLength);
    dataOff += Math.ceil(data.byteLength / 4) * 4;
  }
  outView.setUint32(headOff + 8, (0xb1b0afba - sum(out)) >>> 0);
  return out;
}

const resolveText = createResolver(documentCapabilities);

describe("Orthogonal dual-script font resolution, segmentation, and emission", () => {
  it("segments mixed Latin/CJK/Kana/Hangul/fullwidth runs across theme slots and orthogonal Span overrides without false missing-glyph errors", () => {
    const jostRegular = font("Jost*", [{ weight: 400, data: jostBookBytes }]);
    const jostBold = font("Jost*", [{ weight: 700, data: jostBoldBytes }]);
    const jostSpan = font("Jost Span", [{ weight: 400, data: jostBookBytes }], {
      substitute: true,
    });
    const notoEa = font(
      "Noto Sans CJK SC",
      [{ weight: 400, data: notoScWoff2Bytes }],
      { substitute: true },
    );

    const dualTheme = defineTheme({
      name: "Dual",
      fonts: {
        major: jostBold,
        minor: jostRegular,
        majorEastAsian: notoEa,
        minorEastAsian: notoEa,
      },
    });

    const doc = deck({
      width: 640,
      height: 360,
      theme: dualTheme,
    }).append(
      slide({ id: "s1" }).append(
        text(
          [
            paragraph([
              "   Hello 世界 あいう 한국，",
              span("Code 注释", { style: { fontFamily: jostSpan } }),
            ]),
          ],
          {
            id: "mixed",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 500,
              height: "auto",
              fontSize: 20,
            },
          },
        ),
        text("Only ASCII", {
          id: "latin-only",
          style: {
            position: "absolute",
            left: 20,
            top: 100,
            width: 300,
            height: "auto",
            fontFamily: jostRegular,
            eastAsianFontFamily: "Missing EA Family",
          },
        }),
        text("中文日本語", {
          id: "ea-only",
          style: {
            position: "absolute",
            left: 20,
            top: 160,
            width: 300,
            height: "auto",
            fontFamily: "Missing Latin Family",
            eastAsianFontFamily: notoEa,
          },
        }),
      ),
    );

    const resolution = resolveText(doc);
    assert.equal(
      resolution.slide("s1").get("mixed").textMetrics?.status,
      "measured",
    );
    assert.equal(
      resolution.slide("s1").get("latin-only").textMetrics?.status,
      "measured",
    );
    assert.equal(
      resolution.slide("s1").get("ea-only").textMetrics?.status,
      "measured",
    );

    // Verify independent FontUsage: Missing EA/Latin families were never touched when unused
    const usedFamilies = new Set(resolution.report.fonts.map((f) => f.family));
    assert.equal(usedFamilies.has("Missing EA Family"), false);
    assert.equal(usedFamilies.has("Missing Latin Family"), false);
    assert.equal(usedFamilies.has("Jost*"), true);
    assert.equal(usedFamilies.has("Jost Span"), true);
    assert.equal(usedFamilies.has("Noto Sans CJK SC"), true);
  });

  it("preserves inherited eastAsianFontFamily when overriding fontFamily in Box/P/Span and supports subset embedding across both scripts", () => {
    const latinFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);
    const cjkTtfBytes = createSyntheticCjkTtf(
      "Synthetic CJK",
      [0x4e2d, 0x6587],
    );
    const cjkFont = font("Synthetic CJK", [{ weight: 400, data: cjkTtfBytes }]);

    const doc = deck({
      width: 640,
      height: 360,
      theme: defineTheme({
        fonts: {
          major: latinFont,
          minor: latinFont,
          majorEastAsian: cjkFont,
          minorEastAsian: cjkFont,
        },
      }),
    }).append(
      slide({ id: "s1" }).append(
        box({
          id: "container",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 400,
            height: 200,
            eastAsianFontFamily: cjkFont,
          },
        }).append(
          text(
            [
              paragraph([
                span("Hello 中文", { style: { fontFamily: latinFont } }),
              ]),
            ],
            {
              id: "inner",
              style: { width: 300, height: "auto" },
            },
          ),
        ),
        shape("Shape 中文", {
          id: "shp",
          style: {
            position: "absolute",
            left: 20,
            top: 240,
            width: 200,
            height: 60,
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
        }),
        table({
          id: "tbl",
          columns: [120],
          style: {
            position: "absolute",
            left: 240,
            top: 240,
            width: 120,
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
        }).append(row().append(cell("Cell 中文"))),
        notes([paragraph("Notes 中文")], {
          style: {
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
        }),
        chart({
          id: "cht",
          style: {
            position: "absolute",
            left: 380,
            top: 240,
            width: 220,
            height: 100,
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "中"] },
              series: [{ name: "S1", values: [1, 2] }],
            },
          ],
        }),
      ),
    );

    const resolution = resolveText(doc);
    const emitted = emit(resolution, { embedFonts: "subset" });
    const parts = unzipSync(emitted.pptx);
    const presXml = strFromU8(parts["ppt/presentation.xml"]!);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);

    assert.match(presXml, /<p:font typeface="Jost\*"/u);
    assert.match(presXml, /<p:font typeface="Synthetic CJK"/u);
    assert.match(
      slideXml,
      /<a:latin typeface="Jost\*"[^/]*\/><a:ea typeface="Synthetic CJK"[^/]*\/><a:cs typeface="Jost\*"\/>/u,
    );
  });

  it("emits differential <a:rPr> slots for bound placeholders and falls back to fontFamily when EA is unconfigured", () => {
    const latinFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);
    const notoEa = font(
      "Noto Sans CJK SC",
      [{ weight: 400, data: notoScWoff2Bytes }],
      { substitute: true },
    );

    const layoutDef = defineSlideLayout({
      name: "DualPh",
      placeholders: {
        title: {
          type: "title",
          style: {
            left: 20,
            top: 20,
            width: 400,
            height: 60,
            fontFamily: "major",
            eastAsianFontFamily: "major",
          },
        },
      },
    });

    const doc = deck({
      width: 640,
      height: 360,
      theme: defineTheme({
        fonts: {
          major: latinFont,
          minor: latinFont,
          majorEastAsian: notoEa,
          minorEastAsian: notoEa,
        },
      }),
    }).append(
      slideMaster({ id: "m1" }).append(slideLayout(layoutDef)),
      slide({
        id: "ph-slide",
        layout: layoutDef,
        placeholders: {
          title: [
            paragraph([
              span("OnlyEA", { style: { eastAsianFontFamily: notoEa } }),
              span("OnlyLatin", { style: { fontFamily: latinFont } }),
            ]),
          ],
        },
      }),
    );

    const slideXml = strFromU8(
      unzipSync(emit(resolveText(doc)).pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      slideXml,
      /<a:rPr><a:ea typeface="Noto Sans CJK SC"[^/]*\/><\/a:rPr><a:t xml:space="preserve">OnlyEA<\/a:t>/u,
    );
    assert.match(
      slideXml,
      /<a:rPr><a:latin typeface="Jost\*"[^/]*\/><a:cs typeface="Jost\*"\/><\/a:rPr><a:t xml:space="preserve">OnlyLatin<\/a:t>/u,
    );

    // Legacy single-script fallback: when neither eastAsianFontFamily nor theme minorEastAsian is set,
    // CJK characters route to the author's fontFamily (e.g. reporting missing glyph on latin-only Jost).
    const legacyFailDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "legacy" }).append(
        text("中文", {
          id: "t1",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: "auto",
            fontFamily: latinFont,
          },
        }),
      ),
    );
    assert.throws(
      () => emit(resolveText(legacyFailDoc), { embedFonts: "full" }),
      (err: unknown) =>
        err instanceof BuildError &&
        err.diagnostics.some((d) => d.code === "font-missing-glyph"),
    );
  });

  it("roundtrips dual-script themes, fromFontCollection fonts, and differential slots across import and codegen", async () => {
    const cjkTtfBytes = createSyntheticCjkTtf(
      "Synthetic CJK",
      [0x4e2d, 0x6587],
    );
    const mixedTtcBytes = buildTestTtc([jostBookBytes, cjkTtfBytes]);
    const latinFont = fromFontCollection(mixedTtcBytes, "Jost*");
    const cjkFont = fromFontCollection(mixedTtcBytes, "Synthetic CJK");

    const dualTheme = defineTheme({
      name: "Dual Roundtrip",
      fonts: {
        major: latinFont,
        minor: latinFont,
        majorEastAsian: cjkFont,
        minorEastAsian: cjkFont,
      },
    });

    const doc = deck({ width: 640, height: 360, theme: dualTheme }).append(
      slide({
        id: "s1",
        footer: {
          text: "Footer 中文",
          style: { fontFamily: latinFont, eastAsianFontFamily: cjkFont },
        },
      }).append(
        text(
          [
            paragraph([
              span("OnlyEA 中文", { style: { eastAsianFontFamily: cjkFont } }),
              span("Both", {
                style: { fontFamily: latinFont, eastAsianFontFamily: cjkFont },
              }),
            ]),
          ],
          {
            id: "tx",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 300,
              height: 60,
              fontSize: 16,
            },
          },
        ),
        shape("Shape 中文", {
          id: "shp",
          style: {
            position: "absolute",
            left: 20,
            top: 90,
            width: 200,
            height: 50,
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
        }),
        table({
          id: "tbl",
          columns: [160],
          style: { position: "absolute", left: 20, top: 150, width: 160 },
        }).append(
          row().append(
            cell("Cell 中文", {
              style: { fontFamily: latinFont, eastAsianFontFamily: cjkFont },
            }),
          ),
        ),
        chart({
          id: "cht",
          title: {
            text: "Chart 中文",
            style: { fontFamily: latinFont, eastAsianFontFamily: cjkFont },
          },
          style: {
            position: "absolute",
            left: 240,
            top: 90,
            width: 360,
            height: 200,
            fontFamily: latinFont,
            eastAsianFontFamily: cjkFont,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "中"] },
              series: [{ name: "S1", values: [1, 2] }],
            },
          ],
        }),
      ),
    );

    const emitted = emit(resolveText(doc), { embedFonts: "full" });
    assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);

    const importer = createImporter(documentCapabilities);
    const imported = importer(emitted.pptx);
    const importedTheme = imported.document.props["theme"] as Theme;
    assert.equal(isFont(importedTheme.fonts.major), true);
    assert.equal(isFont(importedTheme.fonts.minor), true);
    assert.equal(isFont(importedTheme.fonts.majorEastAsian), true);
    assert.equal(isFont(importedTheme.fonts.minorEastAsian), true);
    assert.equal(
      (importedTheme.fonts.minorEastAsian as { family: string }).family,
      "Synthetic CJK",
    );

    const generate = createCodegen(documentCapabilities);
    for (const target of ["vanilla", "react"] as const) {
      const code = generate(imported.document, { target });
      const entrySource = code.files[code.entry] as string;
      assert.match(entrySource, /majorEastAsian/u);
      assert.match(entrySource, /minorEastAsian/u);
      assert.match(entrySource, /eastAsianFontFamily/u);
    }

    // Imported AST still builds and validates cleanly with preserved orthogonal slots.
    const rebuilt = emit(resolveText(imported.document), {
      embedFonts: "full",
    });
    assert.deepEqual((await validatePptx(rebuilt.pptx)).errors, []);
  });
});
