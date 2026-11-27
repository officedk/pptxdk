import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonContainer,
  CommonNode,
  build,
  codegen,
  defineSlideLayout,
  defineTheme,
  font,
  importPptx,
} from "@officedk/pptxdk";
import {
  Deck,
  P,
  Slide,
  SlideLayout,
  SlideMaster,
  Span,
  Text,
} from "@officedk/pptxdk/react";
import { unpackEotFontData } from "@pptxdk/import";
import { unzipSync } from "fflate";
import * as fontkit from "fontkit";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

function testFont(relPath: string): Uint8Array {
  return new Uint8Array(readFileSync(new URL(relPath, import.meta.url)));
}

function patchWeightAndName(
  source: Uint8Array,
  patch: {
    readonly weightClass: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly oblique?: boolean;
    readonly italicAngle?: number;
    readonly familyName?: string;
    readonly subfamily?: string;
  },
): Uint8Array {
  const view = new DataView(
    source.buffer,
    source.byteOffset,
    source.byteLength,
  );
  const sfntVersion = view.getUint32(0);
  const numTables = view.getUint16(4);
  const tables = new Map<string, Uint8Array>();

  for (let index = 0; index < numTables; index += 1) {
    const recOffset = 12 + index * 16;
    const tag = String.fromCharCode(
      ...source.subarray(recOffset, recOffset + 4),
    );
    const offset = view.getUint32(recOffset + 8);
    const length = view.getUint32(recOffset + 12);
    const tableBytes = source.slice(offset, offset + length);
    const tableView = new DataView(tableBytes.buffer);
    if (tag === "OS/2") {
      tableView.setUint16(4, patch.weightClass);
      let fsSelection = tableView.getUint16(62);
      if (patch.italic !== undefined) {
        fsSelection = patch.italic
          ? fsSelection | 0x0001
          : fsSelection & ~0x0001;
      }
      if (patch.bold !== undefined) {
        fsSelection = patch.bold ? fsSelection | 0x0020 : fsSelection & ~0x0020;
      }
      if (patch.oblique !== undefined) {
        fsSelection = patch.oblique
          ? fsSelection | 0x0200
          : fsSelection & ~0x0200;
      }
      tableView.setUint16(62, fsSelection);
    } else if (tag === "post" && patch.italicAngle !== undefined) {
      tableView.setInt32(4, Math.round(patch.italicAngle * 65536));
    }
    tables.set(tag, tableBytes);
  }

  if (patch.familyName !== undefined) {
    const sub = patch.subfamily ?? "Regular";
    const records = [
      { nameID: 1, value: patch.familyName },
      { nameID: 2, value: sub },
      { nameID: 4, value: `${patch.familyName} ${sub}` },
      { nameID: 5, value: "Version 1.0" },
    ].map((entry) => {
      const encoded = new Uint8Array(entry.value.length * 2);
      const ev = new DataView(encoded.buffer);
      for (let i = 0; i < entry.value.length; i += 1) {
        ev.setUint16(i * 2, entry.value.charCodeAt(i), false);
      }
      return { ...entry, bytes: encoded };
    });
    const headerSize = 6 + records.length * 12;
    const strBytes = records.reduce((sum, r) => sum + r.bytes.byteLength, 0);
    const nameTable = new Uint8Array(headerSize + strBytes);
    const nv = new DataView(nameTable.buffer);
    nv.setUint16(0, 0);
    nv.setUint16(2, records.length);
    nv.setUint16(4, headerSize);
    let strOff = 0;
    for (let i = 0; i < records.length; i += 1) {
      const rec = records[i]!;
      const off = 6 + i * 12;
      nv.setUint16(off, 3);
      nv.setUint16(off + 2, 1);
      nv.setUint16(off + 4, 0x0409);
      nv.setUint16(off + 6, rec.nameID);
      nv.setUint16(off + 8, rec.bytes.byteLength);
      nv.setUint16(off + 10, strOff);
      nameTable.set(rec.bytes, headerSize + strOff);
      strOff += rec.bytes.byteLength;
    }
    tables.set("name", nameTable);
  }

  const sorted = [...tables.entries()].toSorted(([a], [b]) =>
    a < b ? -1 : a > b ? 1 : 0,
  );
  const count = sorted.length;
  const headerSize = 12 + count * 16;
  let totalSize = headerSize;
  for (const [, data] of sorted) {
    totalSize += Math.ceil(data.byteLength / 4) * 4;
  }

  const checksum = (bytes: Uint8Array): number => {
    const padded = Math.ceil(bytes.byteLength / 4) * 4;
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let sum = 0;
    for (let o = 0; o < padded; o += 4) {
      if (o + 4 <= bytes.byteLength) {
        sum = (sum + dv.getUint32(o)) >>> 0;
      } else {
        let word = 0;
        for (let b = 0; b < 4; b += 1) {
          word = (word << 8) | (o + b < bytes.byteLength ? bytes[o + b]! : 0);
        }
        sum = (sum + (word >>> 0)) >>> 0;
      }
    }
    return sum >>> 0;
  };

  const out = new Uint8Array(totalSize);
  const ov = new DataView(out.buffer);
  ov.setUint32(0, sfntVersion >>> 0);
  ov.setUint16(4, count);
  const entrySelector = Math.floor(Math.log2(Math.max(1, count)));
  const searchRange = (1 << entrySelector) * 16;
  ov.setUint16(6, searchRange);
  ov.setUint16(8, entrySelector);
  ov.setUint16(10, count * 16 - searchRange);

  let dataOff = headerSize;
  let headOff: number | undefined;
  for (let i = 0; i < sorted.length; i += 1) {
    const [tag, raw] = sorted[i]!;
    const bytes =
      tag === "head" && raw.byteLength >= 12 ? new Uint8Array(raw) : raw;
    if (tag === "head" && bytes.byteLength >= 12) {
      new DataView(bytes.buffer).setUint32(8, 0);
      headOff = dataOff;
    }
    out.set(bytes, dataOff);
    const recOff = 12 + i * 16;
    for (let c = 0; c < 4; c += 1) out[recOff + c] = tag.charCodeAt(c);
    ov.setUint32(recOff + 4, checksum(bytes));
    ov.setUint32(recOff + 8, dataOff);
    ov.setUint32(recOff + 12, bytes.byteLength);
    dataOff += Math.ceil(bytes.byteLength / 4) * 4;
  }
  if (headOff !== undefined) {
    ov.setUint32(headOff + 8, (0xb1b0afba - checksum(out)) >>> 0);
  }
  return out;
}

function collectNodesByType(
  root: CommonNode,
  type: string,
  out: CommonNode[] = [],
): CommonNode[] {
  if (root.type === type) out.push(root);
  if (root instanceof CommonContainer) {
    for (const child of root.children) {
      if (child instanceof CommonNode) {
        collectNodesByType(child, type, out);
      }
    }
  }
  return out;
}

function readEmbeddedWeightsFromPptx(pptx: Uint8Array): readonly number[] {
  const entries = unzipSync(pptx);
  const weights: number[] = [];
  for (const [path, bytes] of Object.entries(entries)) {
    if (!path.startsWith("ppt/fonts/") || !path.endsWith(".fntdata")) continue;
    const unpacked = unpackEotFontData(bytes);
    assert.equal(unpacked.status, "ok");
    if (unpacked.status === "ok") {
      const opened = fontkit.create(Buffer.from(unpacked.ttfBytes));
      if ("OS/2" in opened && opened["OS/2"] !== undefined) {
        weights.push(opened["OS/2"].usWeightClass);
      }
    }
  }
  return weights.toSorted((a, b) => a - b);
}

const jostBook = testFont("./assets/fonts/Jost-400-Book.ttf");
const jostBold = testFont("./assets/fonts/Jost-700-Bold.ttf");
const jostItalic = testFont("./assets/fonts/Jost-400-Italic.ttf");
const jostBoldItalic = testFont("./assets/fonts/Jost-700-BoldItalic.ttf");
const darkerGrotesqueSemiBold = testFont(
  "./assets/fonts/DarkerGrotesque-SemiBold.ttf",
);

describe("End-to-end imported embedded font weight roundtrip", () => {
  it("preserves non-400/700 static weights across build -> importPptx -> codegen (vanilla & react) -> rebuild", async () => {
    const jost350 = patchWeightAndName(jostBook, {
      weightClass: 350,
      bold: false,
      familyName: "Jost*",
      subfamily: "Regular",
    });
    const jost800 = patchWeightAndName(jostBold, {
      weightClass: 800,
      bold: true,
      familyName: "Jost*",
      subfamily: "Bold",
    });
    const jostItalic350 = patchWeightAndName(jostItalic, {
      weightClass: 350,
      italic: false,
      oblique: true,
      italicAngle: -10,
      familyName: "Jost*",
      subfamily: "Oblique",
    });
    const jostBoldItalic800 = patchWeightAndName(jostBoldItalic, {
      weightClass: 800,
      bold: true,
      italic: true,
      familyName: "Jost*",
      subfamily: "Bold Italic",
    });
    const jostMedium500 = patchWeightAndName(jostBook, {
      weightClass: 500,
      bold: false,
      familyName: "Jost Medium",
      subfamily: "Regular",
    });

    const primaryFont = font("Jost*", [
      { data: jost350 },
      { data: jost800 },
      { data: jostItalic350 },
      { data: jostBoldItalic800 },
    ]);
    const mediumFont = font("Jost Medium", [{ data: jostMedium500 }]);
    const semiBold600Font = font("Darker Grotesque SemiBold", [
      { data: darkerGrotesqueSemiBold },
    ]);

    const brandTheme = defineTheme({
      name: "Non-RIBBI Weights",
      fonts: { major: primaryFont, minor: mediumFont },
    });

    const heroLayout = defineSlideLayout({
      id: "hero-weights",
      name: "Hero Weights",
      type: "obj",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 48,
            top: 32,
            width: 860,
            height: 60,
            fontFamily: "major",
            fontWeight: 350,
            fontSize: "28pt",
            color: "#0F172A",
          },
          listStyles: {
            1: { fontFamily: "major", fontWeight: 350, fontSize: "28pt" },
            2: { fontFamily: "major", fontWeight: 800, fontSize: "20pt" },
          },
        },
        body: {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 48,
            top: 112,
            width: 860,
            height: 220,
            fontFamily: "minor",
            fontWeight: 500,
            fontSize: "16pt",
            color: "#1E293B",
          },
        },
      },
    });

    for (const embedFonts of ["full", "subset"] as const) {
      const initial = await build(
        createElement(
          Deck,
          { width: 960, height: 540, theme: brandTheme },
          createElement(
            SlideMaster,
            {
              id: "master-weights",
              name: "Master Weights",
              theme: brandTheme,
              preserve: true,
              textStyles: {
                titleStyle: {
                  1: { fontFamily: "major", fontWeight: 350, fontSize: "30pt" },
                },
                bodyStyle: {
                  1: { fontFamily: "minor", fontWeight: 500, fontSize: "16pt" },
                },
              },
              placeholders: {
                title: {
                  type: "title",
                  style: {
                    position: "absolute",
                    left: 48,
                    top: 32,
                    width: 860,
                    height: 60,
                    fontFamily: "major",
                    fontWeight: 350,
                    fontSize: "28pt",
                  },
                },
              },
              footerPlaceholders: {
                footer: {
                  text: "Footer 600",
                  style: {
                    position: "absolute",
                    left: 48,
                    top: 492,
                    width: 400,
                    height: 24,
                    fontFamily: semiBold600Font,
                    fontWeight: 600,
                    fontSize: "11pt",
                    color: "#334155",
                  },
                },
                slideNumber: {
                  style: {
                    position: "absolute",
                    left: 800,
                    top: 492,
                    width: 112,
                    height: 24,
                    fontFamily: mediumFont,
                    fontWeight: 500,
                    fontSize: "11pt",
                    color: "#334155",
                  },
                },
              },
            },
            createElement(SlideLayout, { layout: heroLayout }),
          ),
          createElement(Slide, {
            id: "overview",
            layout: heroLayout,
            footer: { text: "Footer 600", style: { fontWeight: 600 } },
            slideNumber: { style: { fontWeight: 500 } },
            placeholders: {
              title: createElement(
                Text,
                { placeholder: "title" },
                createElement(
                  P,
                  null,
                  "Demi 350 ",
                  createElement(
                    Span,
                    { style: { fontWeight: 800 } },
                    "ExtraBold 800 ",
                  ),
                  createElement(
                    Span,
                    { style: { fontWeight: 800, fontStyle: "italic" } },
                    "BoldItal 800",
                  ),
                ),
              ),
              body: createElement(
                Text,
                { placeholder: "body" },
                createElement(
                  P,
                  null,
                  "Medium 500 ",
                  createElement(
                    Span,
                    {
                      style: {
                        fontFamily: semiBold600Font,
                        fontWeight: 600,
                      },
                    },
                    "DarkerSemiBold 600 ",
                  ),
                  createElement(
                    Span,
                    {
                      style: {
                        fontFamily: primaryFont,
                        fontWeight: 350,
                        fontStyle: "italic",
                      },
                    },
                    "Oblique 350",
                  ),
                ),
              ),
            },
          }),
        ),
        { embedFonts },
      );

      const copy = new Uint8Array(initial.pptx);
      const imported = importPptx(copy);
      copy.fill(0);
      assert.deepEqual(imported.report.diagnostics, []);

      // Verify exact FontFace weights and backfilled document styles.
      const importedTheme = imported.document.props["theme"] as {
        readonly fonts?: {
          readonly major?: {
            readonly faces?: readonly { readonly weight?: number }[];
          };
          readonly minor?: {
            readonly faces?: readonly { readonly weight?: number }[];
          };
        };
      };
      assert.deepEqual(
        importedTheme.fonts?.major?.faces?.map((f) => f.weight),
        [350, 800, 350, 800],
      );
      assert.deepEqual(
        importedTheme.fonts?.minor?.faces?.map((f) => f.weight),
        [500],
      );

      const importedSlide = collectNodesByType(imported.document, "slide")[0]!;
      assert.equal(
        (
          importedSlide.props["footer"] as {
            readonly style?: { readonly fontWeight?: unknown };
          }
        ).style?.fontWeight,
        600,
      );
      assert.equal(
        (
          importedSlide.props["slideNumber"] as {
            readonly style?: { readonly fontWeight?: unknown };
          }
        ).style?.fontWeight,
        500,
      );

      const vanillaCode = codegen(imported.document, { target: "vanilla" });
      const reactCode = codegen(imported.document, { target: "react" });

      if (embedFonts === "full") {
        await verifyGeneratedSourcesTypecheck([
          { label: "imported-weights-vanilla", result: vanillaCode },
          { label: "imported-weights-react", result: reactCode },
        ]);
      }

      for (const [target, generated] of [
        ["vanilla", vanillaCode],
        ["react", reactCode],
      ] as const) {
        const source = generated.files[generated.entry] as string;
        assert.match(source, /weight:\s*350\b/u);
        assert.match(source, /weight:\s*500\b/u);
        assert.match(source, /weight:\s*600\b/u);
        assert.match(source, /weight:\s*800\b/u);
        assert.match(source, /fontWeight:\s*350\b/u);
        assert.match(source, /fontWeight:\s*500\b/u);
        assert.match(source, /fontWeight:\s*600\b/u);
        assert.match(source, /fontWeight:\s*800\b/u);

        const rebuilt = await rebuildGeneratedSourceRelocated(generated, {
          verifyVanillaIsolation: target === "vanilla" && embedFonts === "full",
          buildOptions: { embedFonts: "full" },
        });
        assert.deepEqual(rebuilt.file.errors, []);
        assert.deepEqual(
          readEmbeddedWeightsFromPptx(rebuilt.built.pptx),
          [350, 350, 500, 600, 800, 800],
        );

        const reImported = importPptx(new Uint8Array(rebuilt.built.pptx));
        assert.deepEqual(reImported.report.diagnostics, []);
      }
    }
  }, 60_000);
});
