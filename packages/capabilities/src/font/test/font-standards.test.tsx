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
  patchSfntOs2AndPost,
  replaceSfntNameTable,
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
const notoSansScWoff2 = new Uint8Array(
  readFileSync(
    new URL(
      "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
      import.meta.url,
    ),
  ),
);

function readFirstEmbeddedEotItalicFlag(pptx: Uint8Array): number {
  const parts = unzipSync(pptx);
  const fontPart = Object.keys(parts).find(
    (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
  );
  assert.ok(fontPart !== undefined, "expected embedded .fntdata part");
  const eot = parts[fontPart]!;
  // EOT header offset 27 holds the 1-byte Italic flag
  // (ulEOTSize(4) + ulFontDataSize(4) + ulVersion(4) + ulFlags(4) + FontPANOSE(10) + Charset(1)).
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

  it("matches non-English Unicode Name ID 1 and Name ID 16 records across Platform 0 and Platform 3 without substitute: true", async () => {
    // Platform 3 Encoding 1 Language 0x0804 (zh-CN) only — no 0x0409 English record.
    const zhOnlyTtf = replaceSfntNameTable(jostBook, [
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 1,
        value: "Aether ZH Legacy",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 2,
        value: "Regular",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 4,
        value: "Aether ZH Legacy Regular",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 5,
        value: "Version 1.000",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 16,
        value: "Aether ZH Typographic",
      },
    ]);

    // Platform 0 Encoding 3 Language 0 (Name ID 1) + Platform 3 Encoding 10 Language 0x040C (Name ID 16), no 0x0409.
    const multiUnicodeTtf = replaceSfntNameTable(jostBook, [
      {
        platformID: 0,
        encodingID: 3,
        languageID: 0,
        nameID: 1,
        value: "Aether Platform0 Legacy",
      },
      {
        platformID: 3,
        encodingID: 10,
        languageID: 0x040c,
        nameID: 16,
        value: "Aether French Preferred",
      },
    ]);

    // Bilingual font where 0x0409 English differs from 0x0804/0x040C records.
    const bilingualTtf = replaceSfntNameTable(jostBook, [
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0409,
        nameID: 1,
        value: "Aether English Legacy",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0804,
        nameID: 1,
        value: "思源测试黑体",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0409,
        nameID: 16,
        value: "Aether English Typo",
      },
      {
        platformID: 3,
        encodingID: 10,
        languageID: 0x040c,
        nameID: 16,
        value: "Aether Français Typo",
      },
    ]);

    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "multilang" },
          createElement(
            Text,
            {
              id: "zh-name1",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 220,
                height: "auto",
                fontFamily: font("Aether ZH Legacy", [{ data: zhOnlyTtf }]),
                fontSize: 16,
              },
            },
            "ZH Legacy",
          ),
          createElement(
            Text,
            {
              id: "zh-name16",
              style: {
                position: "absolute",
                left: 16,
                top: 56,
                width: 220,
                height: "auto",
                fontFamily: font("  aether   zh  typographic ", [
                  { data: zhOnlyTtf },
                ]),
                fontSize: 16,
              },
            },
            "ZH Preferred",
          ),
          createElement(
            Text,
            {
              id: "p0-name1",
              style: {
                position: "absolute",
                left: 16,
                top: 96,
                width: 220,
                height: "auto",
                fontFamily: font("Aether Platform0 Legacy", [
                  { data: multiUnicodeTtf },
                ]),
                fontSize: 16,
              },
            },
            "Platform 0",
          ),
          createElement(
            Text,
            {
              id: "fr-name16",
              style: {
                position: "absolute",
                left: 16,
                top: 136,
                width: 220,
                height: "auto",
                fontFamily: font("aether french preferred", [
                  { data: multiUnicodeTtf },
                ]),
                fontSize: 16,
              },
            },
            "French Preferred",
          ),
          createElement(
            Text,
            {
              id: "bilingual-zh-name1",
              style: {
                position: "absolute",
                left: 260,
                top: 16,
                width: 200,
                height: "auto",
                fontFamily: font("思源测试黑体", [{ data: bilingualTtf }]),
                fontSize: 16,
              },
            },
            "Localized Name 1",
          ),
          createElement(
            Text,
            {
              id: "bilingual-fr-name16",
              style: {
                position: "absolute",
                left: 260,
                top: 56,
                width: 200,
                height: "auto",
                fontFamily: font("aether français typo", [
                  { data: bilingualTtf },
                ]),
                fontSize: 16,
              },
            },
            "Localized Name 16",
          ),
          createElement(
            Text,
            {
              id: "woff2-zh",
              style: {
                position: "absolute",
                left: 16,
                top: 176,
                width: 220,
                height: "auto",
                fontFamily: font("Noto Sans CJK SC", [
                  { data: notoSansScWoff2 },
                ]),
                fontSize: 16,
              },
            },
            "中文",
          ),
        ),
      ),
    );

    const resolution = resolveDocument(doc);
    assert.deepEqual(paint(resolution).pages[0]?.errors, []);
    const unembedded = emit(resolution);
    assert.deepEqual((await validatePptx(unembedded.pptx)).errors, []);

    // A font whose non-English Unicode Name ID 1 matches declared family also embeds cleanly.
    const embedDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "zh-embed" },
          createElement(
            Text,
            {
              id: "zh-name1-embed",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 220,
                height: "auto",
                fontFamily: font("Aether ZH Legacy", [{ data: zhOnlyTtf }]),
                fontSize: 16,
              },
            },
            "ZH Legacy",
          ),
        ),
      ),
    );
    const embedRes = resolveDocument(embedDoc);
    for (const embedFonts of ["full", "subset"] as const) {
      const emitted = emit(embedRes, { embedFonts });
      assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    }
  });

  it("allows Name ID 16 match without embedFonts, rejects embedFonts with font-family-mismatch naming native Name ID 1, and never mutates font bytes", async () => {
    const mediumNonRibbi = replaceSfntNameTable(
      patchSfntOs2AndPost(jostBook, { weightClass: 500, bold: false }),
      [
        {
          platformID: 3,
          encodingID: 1,
          languageID: 0x0409,
          nameID: 1,
          value: "Aether Sans Medium",
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
          value: "Aether Sans Medium Regular",
        },
        {
          platformID: 3,
          encodingID: 1,
          languageID: 0x0409,
          nameID: 5,
          value: "Version 1.000",
        },
        {
          platformID: 3,
          encodingID: 1,
          languageID: 0x0409,
          nameID: 16,
          value: "Aether Sans",
        },
      ],
    );

    const typographicDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "non-ribbi" },
          createElement(
            Text,
            {
              id: "medium-typo",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: font("Aether Sans", [{ data: mediumNonRibbi }]),
                fontWeight: 500,
                fontSize: 18,
              },
            },
            "Medium via Name ID 16",
          ),
        ),
      ),
    );

    const resolution = resolveDocument(typographicDoc);
    assert.equal(
      resolution.report.fonts.find((entry) => entry.family === "Aether Sans")
        ?.embeddable,
      false,
    );
    assert.deepEqual(paint(resolution).pages[0]?.errors, []);
    const unembedded = emit(resolution);
    assert.deepEqual((await validatePptx(unembedded.pptx)).errors, []);

    for (const embedFonts of ["full", "subset"] as const) {
      assert.throws(
        () => emit(resolution, { embedFonts }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          const diag = error.diagnostics[0];
          assert.equal(diag?.code, "font-unavailable");
          if (diag?.code === "font-unavailable") {
            assert.equal(diag.reason, "family-mismatch");
          }
          const msg = diag?.message ?? "";
          assert.match(msg, /Name ID 16/u);
          assert.match(msg, /Name ID 1/u);
          assert.match(msg, /Aether Sans Medium/u);
          assert.match(msg, /font\("Aether Sans Medium"/u);
          return true;
        },
      );
    }

    // Registering with native Name ID 1 embeds cleanly and preserves exact source bytes in full mode.
    const nativeDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "non-ribbi-native" },
          createElement(
            Text,
            {
              id: "medium-native",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: font("Aether Sans Medium", [
                  { data: mediumNonRibbi },
                ]),
                fontWeight: 500,
                fontSize: 18,
              },
            },
            "Medium via Name ID 1",
          ),
        ),
      ),
    );
    const nativeRes = resolveDocument(nativeDoc);
    assert.equal(
      nativeRes.report.fonts.find(
        (entry) => entry.family === "Aether Sans Medium",
      )?.embeddable,
      true,
    );
    const fullEmitted = emit(nativeRes, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(fullEmitted.pptx)).errors, []);
    const parts = unzipSync(fullEmitted.pptx);
    const fontPartPath = Object.keys(parts).find(
      (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
    );
    assert.ok(fontPartPath !== undefined);
    const eot = parts[fontPartPath]!;
    const eotView = new DataView(eot.buffer, eot.byteOffset, eot.byteLength);
    const fontDataSize = eotView.getUint32(4, true);
    assert.deepEqual(
      eot.subarray(eot.byteLength - fontDataSize),
      mediumNonRibbi,
    );
  });

  it("classifies oblique-only and italicAngle-only TTFs into italic and boldItalic slots across resolve and full/subset embedFonts", async () => {
    for (const [variant, patch] of [
      ["oblique-only", { italic: false, oblique: true, italicAngle: 0 }],
      ["angle-only", { italic: false, oblique: false, italicAngle: -12 }],
    ] as const) {
      const italicFace = patchSfntOs2AndPost(jostBook, patch);
      const boldItalicFace = patchSfntOs2AndPost(jostBold, {
        ...patch,
        bold: true,
      });
      const testFont = font("Jost*", [
        { data: italicFace },
        { data: boldItalicFace },
      ]);
      assert.equal(testFont.faces?.[0]?.style, "italic", variant);
      assert.equal(testFont.faces?.[1]?.style, "italic", variant);

      const doc = await expandDocument(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: variant },
            createElement(
              Text,
              {
                id: "t-italic",
                style: {
                  position: "absolute",
                  left: 16,
                  top: 16,
                  width: 240,
                  height: "auto",
                  fontFamily: testFont,
                  fontWeight: 400,
                  fontStyle: "italic",
                  fontSize: 18,
                },
              },
              "Italic slot",
            ),
            createElement(
              Text,
              {
                id: "t-bold-italic",
                style: {
                  position: "absolute",
                  left: 16,
                  top: 56,
                  width: 240,
                  height: "auto",
                  fontFamily: testFont,
                  fontWeight: 700,
                  fontStyle: "italic",
                  fontSize: 18,
                },
              },
              "BoldItalic slot",
            ),
          ),
        ),
      );

      const resolution = resolveDocument(doc);
      assert.deepEqual(paint(resolution).pages[0]?.errors, [], variant);

      for (const embedFonts of ["full", "subset"] as const) {
        const emitted = emit(resolution, { embedFonts });
        const parts = unzipSync(emitted.pptx);
        const presentation = strFromU8(parts["ppt/presentation.xml"]!);
        assert.match(presentation, /<p:italic r:id="rId\d+"\/>/u, variant);
        assert.match(presentation, /<p:boldItalic r:id="rId\d+"\/>/u, variant);
        const fontParts = Object.keys(parts).filter(
          (entry) =>
            entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
        );
        assert.equal(fontParts.length, 2, variant);
        for (const fontPart of fontParts) {
          assert.equal(parts[fontPart]![27], 1, `${variant} ${fontPart}`);
        }
        assert.deepEqual(
          (await validatePptx(emitted.pptx)).errors,
          [],
          variant,
        );
      }
    }
  });

  it("reports native Name ID 1 in four-slot conflicts and distinguishes matching vs mismatched substitute: true diagnostics", async () => {
    // Face with Platform 3 0x0409 Name ID 1 = "Jost Medium" and Platform 0 Name ID 1 = "Jost*"
    // so it matches "Jost*" on Name ID 1 check and exposes its primary Name ID 1 ("Jost Medium") in 4-slot conflict.
    const jost500Conflict = replaceSfntNameTable(
      patchSfntOs2AndPost(jostBook, { weightClass: 500, bold: false }),
      [
        {
          platformID: 0,
          encodingID: 3,
          languageID: 0,
          nameID: 1,
          value: "Jost*",
        },
        {
          platformID: 3,
          encodingID: 1,
          languageID: 0x0409,
          nameID: 1,
          value: "Jost Medium",
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
          value: "Jost Medium Regular",
        },
        {
          platformID: 3,
          encodingID: 1,
          languageID: 0x0409,
          nameID: 5,
          value: "Version 1.000",
        },
      ],
    );

    const conflictFont = font("Jost*", [
      { weight: 400, data: jostBook },
      { weight: 500, data: jost500Conflict },
    ]);

    const sameTextConflictDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "slot-conflict" },
          createElement(
            Text,
            {
              id: "two-weights",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: conflictFont,
                fontSize: 18,
              },
            },
            createElement(
              P,
              null,
              createElement(Span, { style: { fontWeight: 400 } }, "Book "),
              createElement(Span, { style: { fontWeight: 500 } }, "Medium"),
            ),
          ),
        ),
      ),
    );

    const crossSlideConflictDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "book",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: conflictFont,
                fontWeight: 400,
                fontSize: 18,
              },
            },
            "Book",
          ),
        ),
        createElement(
          Slide,
          { id: "s2" },
          createElement(
            Text,
            {
              id: "medium",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: conflictFont,
                fontWeight: 500,
                fontSize: 18,
              },
            },
            "Medium",
          ),
        ),
      ),
    );

    for (const conflictDoc of [sameTextConflictDoc, crossSlideConflictDoc]) {
      const conflictRes = resolveDocument(conflictDoc);
      for (const embedFonts of ["full", "subset"] as const) {
        assert.throws(
          () => emit(conflictRes, { embedFonts }),
          (error: unknown) => {
            assert.ok(error instanceof BuildError);
            const diag = error.diagnostics[0];
            assert.equal(diag?.code, "font-unavailable");
            if (diag?.code === "font-unavailable") {
              assert.equal(diag.reason, "not-embeddable");
            }
            const msg = diag?.message ?? "";
            assert.match(msg, /Jost Medium/u);
            assert.match(msg, /regular/u);
            return true;
          },
        );
      }
    }

    // Redundant substitute: true when native Name ID 1 already matches declared family.
    const redundantSubDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "redundant-sub" },
          createElement(
            Text,
            {
              id: "sub-match",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: font("Jost*", [{ data: jostBook }], {
                  substitute: true,
                }),
                fontSize: 18,
              },
            },
            "Redundant substitute",
          ),
        ),
      ),
    );
    const redundantRes = resolveDocument(redundantSubDoc);
    for (const embedFonts of ["full", "subset"] as const) {
      assert.throws(
        () => emit(redundantRes, { embedFonts }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          const diag = error.diagnostics[0];
          assert.equal(diag?.code, "font-unavailable");
          if (diag?.code === "font-unavailable") {
            assert.equal(diag.reason, "not-embeddable");
          }
          const msg = diag?.message ?? "";
          assert.match(msg, /Name ID 1/u);
          assert.match(msg, /Jost\*/u);
          assert.match(msg, /already matches/iu);
          assert.match(msg, /remove.*substitute/iu);
          return true;
        },
      );
    }

    // Mismatched substitute: true includes actual native Name ID 1 explaining why it cannot embed under another name.
    const mismatchedSubDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "mismatch-sub" },
          createElement(
            Text,
            {
              id: "sub-mismatch",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 240,
                height: "auto",
                fontFamily: font("Display Alias", [{ data: jostBook }], {
                  substitute: true,
                }),
                fontSize: 18,
              },
            },
            "Mismatched substitute",
          ),
        ),
      ),
    );
    const mismatchedRes = resolveDocument(mismatchedSubDoc);
    for (const embedFonts of ["full", "subset"] as const) {
      assert.throws(
        () => emit(mismatchedRes, { embedFonts }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          const diag = error.diagnostics[0];
          assert.equal(diag?.code, "font-unavailable");
          if (diag?.code === "font-unavailable") {
            assert.equal(diag.reason, "not-embeddable");
          }
          const msg = diag?.message ?? "";
          assert.match(msg, /Name ID 1/u);
          assert.match(msg, /Jost\*/u);
          assert.match(msg, /Display Alias/u);
          return true;
        },
      );
    }
  });
});
