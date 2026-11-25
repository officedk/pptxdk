import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, shape, slide } from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { P, Span, Text } from "@pptxdk/capabilities/text-elements";
import { defineTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strToU8, unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

const jostRegular = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBold = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const jostFont = font("Jost*", [
  { weight: 400, data: jostRegular },
  { weight: 700, data: jostBold },
]);

function withSyntheticTable(source: Uint8Array, tag: string): Uint8Array {
  const buf = Buffer.from(source);
  const numTables = buf.readUInt16BE(4);
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
      tag: buf.subarray(off, off + 4).toString("ascii"),
      checkSum: buf.readUInt32BE(off + 4),
      offset: buf.readUInt32BE(off + 8) + 16,
      length: buf.readUInt32BE(off + 12),
    });
  }
  records.push({
    tag,
    checkSum: 0,
    offset: buf.length + 16,
    length: 16,
  });
  records.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const out = Buffer.alloc(buf.length + 32);
  buf.copy(out, 0, 0, 12);
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
  buf.copy(out, nextHeaderSize, headerSize);
  return new Uint8Array(out);
}

describe("Embedded font resource roundtrip & isolation", () => {
  it("imports full and subset embedded EOT fonts with detached bytes so paint() renders PNG after inputBytes.fill(0)", async () => {
    for (const embedFonts of ["full", "subset"] as const) {
      const doc = await expandDocument(
        createElement(
          Deck,
          {
            width: 640,
            height: 360,
            theme: defineTheme({ fonts: { minor: jostFont, major: jostFont } }),
          },
          createElement(
            Slide,
            { id: "fonts" },
            createElement(
              Text,
              {
                id: "headline",
                x: 32,
                y: 28,
                width: 360,
                height: 80,
                style: {
                  fontFamily: "minor",
                  fontSize: "24pt",
                  color: "#0F172A",
                },
              },
              createElement(
                P,
                null,
                "Theme ",
                createElement(
                  Span,
                  { style: { fontFamily: jostFont, fontWeight: "bold" } },
                  "Bold",
                ),
              ),
            ),
          ),
        ),
      );

      const initialEmit = emit(resolveDocument(doc), { embedFonts });
      const pptxCopy = new Uint8Array(initialEmit.pptx);
      const imported = importDocument(pptxCopy);
      // Zero caller buffer immediately to prove embedded font bytes are detached.
      pptxCopy.fill(0);

      assert.deepEqual(imported.report.diagnostics, []);

      const importedResolution = resolveDocument(imported.document);
      const painted = paint(importedResolution);
      assert.equal(painted.pages.length, 1);
      assert.deepEqual(painted.pages[0]?.diagnostics, []);
      assert.ok(painted.pages[0]?.png instanceof Uint8Array);
      const decoded = PNG.sync.read(Buffer.from(painted.pages[0].png));
      assert.equal(decoded.width, 640);
      assert.equal(decoded.height, 360);

      const reemitted = emit(importedResolution, { embedFonts });
      const validation = await validatePptx(reemitted.pptx);
      assert.deepEqual(validation.errors, []);
      assert.equal(validation.slides[0]?.objects[0]?.text?.text, "Theme Bold");
      assert.ok(
        Object.keys(unzipSync(reemitted.pptx)).some(
          (entry) =>
            entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
        ),
      );
    }
  });

  it("attaches imported embedded fonts to Shape.textStyle so paint() renders shape text after inputBytes.fill(0)", async () => {
    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "shape-font" }).append(
        shape({
          id: "badge",
          x: 40,
          y: 40,
          width: 240,
          height: 80,
          fill: "#2563EB",
          text: "Shape Embedded Font",
          textStyle: {
            fontFamily: jostFont,
            fontSize: 16,
            color: "#FFFFFF",
          },
        }),
      ),
    );
    const initialEmit = emit(resolveDocument(doc), { embedFonts: "full" });
    const copy = new Uint8Array(initialEmit.pptx);
    const imported = importDocument(copy);
    copy.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const painted = paint(resolveDocument(imported.document));
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);
  });

  it("isolates compressed, obfuscated, and corrupt .fntdata gracefully while preserving editable text", async () => {
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "t1",
              x: 20,
              y: 20,
              width: 240,
              height: 60,
              style: { fontFamily: jostFont, fontSize: "18pt" },
            },
            "Keep Editable",
          ),
        ),
      ),
    );
    const base = emit(resolveDocument(doc), { embedFonts: "full" });
    const baseEntries = unzipSync(base.pptx);
    const fontPart = Object.keys(baseEntries).find(
      (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
    )!;
    assert.ok(fontPart !== undefined);

    // 1. Compressed EOT (flags |= 0x00000004)
    const compressedBytes = new Uint8Array(baseEntries[fontPart]!);
    new DataView(compressedBytes.buffer).setUint32(12, 0x00000004, true);
    const compressedZip = zipSync({
      ...baseEntries,
      [fontPart]: compressedBytes,
    });
    const compressedImported = importDocument(compressedZip);
    assert.ok(
      compressedImported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" &&
          d.mechanismHint === "compressed-eot",
      ),
    );

    // 2. XOR-obfuscated EOT (flags |= 0x10000000)
    const obfuscatedBytes = new Uint8Array(baseEntries[fontPart]!);
    new DataView(obfuscatedBytes.buffer).setUint32(12, 0x10000000, true);
    const obfuscatedZip = zipSync({
      ...baseEntries,
      [fontPart]: obfuscatedBytes,
    });
    const obfuscatedImported = importDocument(obfuscatedZip);
    assert.ok(
      obfuscatedImported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" &&
          d.mechanismHint === "obfuscated-eot",
      ),
    );

    // 3. Corrupt font bytes
    const corruptZip = zipSync({
      ...baseEntries,
      [fontPart]: strToU8("not-a-font"),
    });
    const corruptImported = importDocument(corruptZip);
    assert.ok(
      corruptImported.report.diagnostics.some((d) => d.code === "corrupt-part"),
    );

    // Text remains editable and re-emittable in all cases.
    const rebuilt = emit(resolveDocument(corruptImported.document));
    const validation = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(validation.slides[0]?.objects[0]?.text?.text, "Keep Editable");
  });

  it("imports unembedded text cleanly and preserves native semantics even when paint() reports missing outline font", async () => {
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "t1",
              x: 20,
              y: 20,
              width: 240,
              height: 60,
              style: { fontFamily: "MissingPreviewFamily", fontSize: "20pt" },
            },
            "Still Native",
          ),
        ),
      ),
    );

    const initial = emit(resolveDocument(doc));
    const imported = importDocument(new Uint8Array(initial.pptx));
    assert.deepEqual(imported.report.diagnostics, []);

    const resolution = resolveDocument(imported.document);
    const painted = paint(resolution);
    assert.equal(painted.pages[0]?.png, undefined);
    assert.ok((painted.pages[0]?.diagnostics.length ?? 0) > 0);

    const reemitted = emit(resolution);
    const validation = await validatePptx(reemitted.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(
      validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontFamily,
      "MissingPreviewFamily",
    );
    assert.equal(
      validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      20,
    );

    // Color font (COLR table) imports cleanly and preserves native semantics while paint() rejects color font preview.
    const colrFont = font(
      "ColrJost",
      [{ weight: 400, data: withSyntheticTable(jostRegular, "COLR") }],
      { substitute: true },
    );
    const colrDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "colr" },
          createElement(
            Text,
            {
              id: "c1",
              x: 20,
              y: 20,
              width: 240,
              height: 60,
              style: { fontFamily: colrFont, fontSize: "16pt" },
            },
            "Color Font Text",
          ),
        ),
      ),
    );
    const colrInitial = emit(resolveDocument(colrDoc));
    const colrImported = importDocument(new Uint8Array(colrInitial.pptx));
    assert.deepEqual(colrImported.report.diagnostics, []);
    const colrRes = resolveDocument(colrImported.document);
    const colrPaint = paint(colrRes);
    assert.equal(colrPaint.pages[0]?.png, undefined);
    assert.ok((colrPaint.pages[0]?.diagnostics.length ?? 0) > 0);
    const colrValidation = await validatePptx(emit(colrRes).pptx);
    assert.deepEqual(colrValidation.errors, []);
    assert.equal(
      colrValidation.slides[0]?.objects[0]?.text?.text,
      "Color Font Text",
    );
  });
});
