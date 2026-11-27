import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { font } from "@pptxdk/capabilities/font";
import { P, Span, Text } from "@pptxdk/capabilities/text-elements";
import { BuildError, defineTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  createMinimalCffOtf,
  stripVariableFontTables,
} from "../../../test/sfnt-test-helpers.ts";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

const jostBook = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBold = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const jostBoldItalic = new Uint8Array(
  readFileSync(
    new URL(
      "../../../test/assets/fonts/Jost-700-BoldItalic.ttf",
      import.meta.url,
    ),
  ),
);
const rawNotoArabic = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/NotoNaskhArabic.ttf", import.meta.url),
  ),
);
const rawNotoHebrew = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/NotoSansHebrew.ttf", import.meta.url),
  ),
);

function patchSfntOs2AndPost(
  source: Uint8Array,
  patch: {
    readonly weightClass?: number;
    readonly italic?: boolean;
    readonly oblique?: boolean;
    readonly bold?: boolean;
    readonly italicAngle?: number;
  },
): Uint8Array {
  const copy = new Uint8Array(source);
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  const count = view.getUint16(4);
  for (let index = 0; index < count; index += 1) {
    const rec = 12 + index * 16;
    const tag = String.fromCharCode(...copy.subarray(rec, rec + 4));
    const offset = view.getUint32(rec + 8);
    if (tag === "OS/2") {
      if (patch.weightClass !== undefined) {
        view.setUint16(offset + 4, patch.weightClass);
      }
      let fsSelection = view.getUint16(offset + 62);
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
      view.setUint16(offset + 62, fsSelection);
    } else if (tag === "post" && patch.italicAngle !== undefined) {
      view.setInt32(offset + 4, Math.round(patch.italicAngle * 65536));
    }
  }
  return copy;
}

function readFirstEmbeddedEotItalicFlag(pptx: Uint8Array): number {
  const parts = unzipSync(pptx);
  const fontPart = Object.keys(parts).find(
    (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
  );
  assert.ok(fontPart !== undefined, "expected embedded .fntdata part");
  const eot = parts[fontPart]!;
  // EOT header offset 53 is the 1-byte Italic flag (4+4+4+4+10+1+4+2+2+16+4 = 55? Check:
  // ulEOTSize(4)+ulFontDataSize(4)+ulVersion(4)+ulFlags(4)+FontPANOSE(10)+Charset(1)=27, Italic at 27).
  return eot[27]!;
}

describe("Unified font module standards & validation", () => {
  it("supports 1..1000 integer fontWeight across resolve, full/subset embedFonts, paint, validatePptx, and import", async () => {
    const jost350 = patchSfntOs2AndPost(jostBook, { weightClass: 350 });
    const jost650 = patchSfntOs2AndPost(jostBold, {
      weightClass: 650,
      bold: true,
    });
    const customFont = font("Jost*", [
      { weight: 350, style: "normal", data: jost350 },
      { weight: 650, style: "normal", data: jost650 },
    ]);

    const doc = await expandDocument(
      createElement(
        Deck,
        {
          width: 640,
          height: 360,
          theme: defineTheme({
            fonts: { minor: customFont, major: customFont },
          }),
        },
        createElement(
          Slide,
          { id: "weights" },
          createElement(
            Text,
            {
              id: "non-hundreds",
              style: {
                position: "absolute",
                left: 32,
                top: 32,
                width: 400,
                height: "auto",
                fontFamily: customFont,
                fontSize: 20,
              },
            },
            createElement(
              P,
              null,
              createElement(Span, { style: { fontWeight: 350 } }, "Light 350 "),
              createElement(
                Span,
                { style: { fontWeight: 650 } },
                "SemiBold 650",
              ),
            ),
          ),
        ),
      ),
    );

    const resolution = resolveDocument(doc);
    assert.deepEqual(
      resolution.report.fonts.map((usage) => ({
        weight: usage.weight,
        hasResource: usage.hasResource,
        embeddable: usage.embeddable,
      })),
      [
        { weight: 350, hasResource: true, embeddable: true },
        { weight: 650, hasResource: true, embeddable: true },
      ],
    );

    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);

    for (const embedFonts of ["full", "subset"] as const) {
      const emitted = emit(resolution, { embedFonts });
      const validation = await validatePptx(emitted.pptx);
      assert.deepEqual(validation.errors, []);
      const runs =
        validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs ?? [];
      assert.equal(runs[0]?.bold, false);
      assert.equal(runs[1]?.bold, true);

      const imported = importDocument(emitted.pptx);
      assert.deepEqual(imported.report.diagnostics, []);
      const importedTheme = imported.document.props["theme"] as
        | {
            readonly fonts?: {
              readonly minor?: {
                readonly faces?: readonly { readonly weight?: number }[];
              };
            };
          }
        | undefined;
      const importedWeights = importedTheme?.fonts?.minor?.faces?.map(
        (face) => face.weight,
      );
      assert.deepEqual(importedWeights, [350, 650]);
    }
  });

  it("infers omitted FontFace weight and style from static font metadata, including oblique and italicAngle", async () => {
    const jost350 = patchSfntOs2AndPost(jostBook, { weightClass: 350 });
    const cffOblique = createMinimalCffOtf({
      familyName: "Oblique CFF",
      weightClass: 450,
      italic: false,
      oblique: true,
      bold: false,
      italicAngle: 0,
    });
    const cffAngleOnly = createMinimalCffOtf({
      familyName: "Angle CFF",
      weightClass: 550,
      italic: false,
      oblique: false,
      bold: false,
      italicAngle: -11,
    });

    const inferred350 = font("Jost*", [{ data: jost350 }]);
    assert.equal(inferred350.faces?.[0]?.weight, 350);
    assert.equal(inferred350.faces?.[0]?.style, "normal");

    const inferredBoldItalic = font("Jost*", [{ data: jostBoldItalic }]);
    assert.equal(inferredBoldItalic.faces?.[0]?.weight, 700);
    assert.equal(inferredBoldItalic.faces?.[0]?.style, "italic");

    const inferredOblique = font("Oblique CFF", [{ data: cffOblique }]);
    assert.equal(inferredOblique.faces?.[0]?.weight, 450);
    assert.equal(inferredOblique.faces?.[0]?.style, "italic");

    const inferredAngle = font("Angle CFF", [{ data: cffAngleOnly }]);
    assert.equal(inferredAngle.faces?.[0]?.weight, 550);
    assert.equal(inferredAngle.faces?.[0]?.style, "italic");

    // Also verifies literal Font objects bypassing font() resolve cleanly with inferred weight.
    const literalFont = { family: "Jost*", faces: [{ data: jost350 }] };
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "inferred" },
          createElement(
            Text,
            {
              id: "auto-weight",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 320,
                height: "auto",
                fontFamily: literalFont,
                fontWeight: 350,
                fontSize: 18,
              },
            },
            "Inferred 350",
          ),
        ),
      ),
    );
    const resolution = resolveDocument(expanded);
    assert.deepEqual(paint(resolution).pages[0]?.errors, []);
  });

  it("rejects explicit weight or style mismatches on fixed-size text boxes with BuildError naming declared and actual metadata", async () => {
    const jost350 = patchSfntOs2AndPost(jostBook, { weightClass: 350 });
    const cffOblique = createMinimalCffOtf({
      familyName: "Oblique CFF",
      weightClass: 450,
      oblique: true,
    });

    // font() must stay path-safe and not throw eagerly on mismatched metadata.
    const wrongWeightFont = font("Jost*", [{ weight: 400, data: jost350 }]);
    const wrongStyleFont = font("Oblique CFF", [
      { weight: 450, style: "normal", data: cffOblique },
    ]);

    const weightMismatchDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "fixed-wrong-weight",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 240,
                height: 40,
                fontFamily: wrongWeightFont,
                fontWeight: 400,
              },
            },
            "Wrong weight",
          ),
        ),
      ),
    );

    assert.throws(
      () => resolveDocument(weightMismatchDoc),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-unavailable");
        if (diag?.code === "font-unavailable") {
          assert.equal(diag.reason, "style-mismatch");
        }
        assert.equal(diag?.path, "deck/slide[s1]/text[fixed-wrong-weight]");
        assert.match(diag?.message ?? "", /400/u);
        assert.match(diag?.message ?? "", /350/u);
        return true;
      },
    );

    const styleMismatchDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "fixed-wrong-style",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 240,
                height: 40,
                fontFamily: wrongStyleFont,
                fontWeight: 450,
                fontStyle: "normal",
              },
            },
            "Wrong style",
          ),
        ),
      ),
    );

    assert.throws(
      () => resolveDocument(styleMismatchDoc),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-unavailable");
        if (diag?.code === "font-unavailable") {
          assert.equal(diag.reason, "style-mismatch");
        }
        assert.equal(diag?.path, "deck/slide[s1]/text[fixed-wrong-style]");
        assert.match(diag?.message ?? "", /normal/u);
        assert.match(diag?.message ?? "", /italic/u);
        return true;
      },
    );
  });

  it("rejects variable fonts with fvar in both resolve (unembedded, fixed and content-sized) and emit", async () => {
    // font() must remain non-throwing for variable fonts so BuildError carries node path.
    const varArabic = font("Noto Naskh Arabic", [
      { weight: 400, data: rawNotoArabic },
    ]);
    const varHebrewOmittedWeight = font("Noto Sans Hebrew Thin", [
      { data: rawNotoHebrew },
    ]);

    for (const [label, fontValue, textContent, fontWeight, height] of [
      ["arabic-fixed", varArabic, "مرحبا", 400, 40],
      ["arabic-auto", varArabic, "مرحبا", 400, "auto"],
      ["hebrew-omitted-weight", varHebrewOmittedWeight, "שלום", 100, 40],
    ] as const) {
      const doc = await expandDocument(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: "var-slide" },
            createElement(
              Text,
              {
                id: label,
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 240,
                  height,
                  fontFamily: fontValue,
                  fontWeight,
                },
              },
              textContent,
            ),
          ),
        ),
      );

      assert.throws(
        () => resolveDocument(doc),
        (error: unknown) => {
          assert.ok(error instanceof BuildError, label);
          const diag = error.diagnostics[0];
          assert.equal(diag?.code, "font-unavailable", label);
          if (diag?.code === "font-unavailable") {
            assert.equal(diag.reason, "unsupported-format", label);
          }
          assert.equal(diag?.path, `deck/slide[var-slide]/text[${label}]`);
          assert.match(diag?.message ?? "", /fvar/u, label);
          assert.match(diag?.message ?? "", /static/iu, label);
          return true;
        },
      );
    }

    // Confirm stripping fvar/variation tables makes the default static instance pass cleanly.
    const staticHebrew = font("Noto Sans Hebrew Thin", [
      { data: stripVariableFontTables(rawNotoHebrew) },
    ]);
    assert.equal(staticHebrew.faces?.[0]?.weight, 100);
    const staticDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "ok" },
          createElement(
            Text,
            {
              id: "hebrew-ok",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 240,
                height: "auto",
                fontFamily: staticHebrew,
                fontWeight: 100,
              },
            },
            "שלום",
          ),
        ),
      ),
    );
    const res = resolveDocument(staticDoc);
    assert.deepEqual(paint(res).pages[0]?.errors, []);
  });

  it("shapes and paints static monochrome CFF (.otf) fonts via HarfBuzz while rejecting embedFonts with unsupported-font-format", async () => {
    const otfBytes = createMinimalCffOtf({
      familyName: "Aether CFF",
      weightClass: 450,
    });
    const otfFont = font("Aether CFF", [{ data: otfBytes }]);
    assert.equal(otfFont.faces?.[0]?.weight, 450);
    assert.equal(otfFont.faces?.[0]?.style, "normal");

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "otf-slide" },
          createElement(
            Text,
            {
              id: "otf-text",
              style: {
                position: "absolute",
                left: 24,
                top: 24,
                width: 300,
                height: "auto",
                fontFamily: otfFont,
                fontWeight: 450,
                fontSize: 24,
                color: "#0F172A",
              },
            },
            "Hello CFF 123",
          ),
        ),
      ),
    );

    const resolution = resolveDocument(doc);
    const metrics = resolution.slide("otf-slide").get("otf-text").textMetrics;
    assert.equal(metrics?.status, "measured");
    if (metrics?.status === "measured") {
      assert.equal(metrics.fits, true);
      assert.ok(metrics.content.width > 0);
      assert.ok(metrics.content.height > 0);
    }

    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);
    const decoded = PNG.sync.read(Buffer.from(painted.pages[0].png));
    assert.equal(decoded.width, 480);
    assert.equal(decoded.height, 270);

    // Unembedded emit succeeds cleanly.
    const unembedded = emit(resolution);
    assert.deepEqual((await validatePptx(unembedded.pptx)).errors, []);

    // Both full and subset embedding reject .otf with unsupported-font-format.
    for (const embedFonts of ["full", "subset"] as const) {
      assert.throws(
        () => emit(resolution, { embedFonts }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          const diag = error.diagnostics[0];
          assert.equal(diag?.code, "font-unavailable");
          if (diag?.code === "font-unavailable") {
            assert.equal(diag.reason, "unsupported-format");
          }
          assert.match(diag?.message ?? "", /TrueType|CFF|\.otf/iu);
          return true;
        },
      );
    }
  });

  it("treats oblique TTF consistently with italic during font embedding and sets the EOT Italic flag", async () => {
    const obliqueTtf = patchSfntOs2AndPost(jostBook, {
      italic: false,
      oblique: true,
      italicAngle: -12,
    });
    const obliqueFont = font("Jost*", [{ data: obliqueTtf }]);
    assert.equal(obliqueFont.faces?.[0]?.weight, 400);
    assert.equal(obliqueFont.faces?.[0]?.style, "italic");

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "oblique" },
          createElement(
            Text,
            {
              id: "oblique-text",
              style: {
                position: "absolute",
                left: 24,
                top: 24,
                width: 300,
                height: "auto",
                fontFamily: obliqueFont,
                fontStyle: "italic",
                fontSize: 20,
              },
            },
            "Oblique Embedded",
          ),
        ),
      ),
    );

    const resolution = resolveDocument(doc);
    for (const embedFonts of ["full", "subset"] as const) {
      const emitted = emit(resolution, { embedFonts });
      const parts = unzipSync(emitted.pptx);
      const presentation = strFromU8(parts["ppt/presentation.xml"]!);
      assert.match(presentation, /<p:italic r:id="rId\d+"\/>/u);
      assert.equal(readFirstEmbeddedEotItalicFlag(emitted.pptx), 1);
      assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    }
  });
});
