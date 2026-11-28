import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  font,
  fromFontCollection,
  slide,
  text,
} from "@pptxdk/capabilities/document";
import { fromFontCollection as fromFontCollectionDirect } from "@pptxdk/capabilities/font";
import { slideBackgroundCapability } from "@pptxdk/capabilities/slide-background";
import { fromFontCollection as fromFontCollectionText } from "@pptxdk/capabilities/text";
import { textCapability } from "@pptxdk/capabilities/text";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import {
  CapabilityError,
  defineTheme,
  hasFontFaceBytes,
  isSealedFont,
  peekFontFaceResource,
} from "@pptxdk/core/capability";
import { emit } from "@pptxdk/emit";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

import {
  buildTestTtc,
  createMinimalCffOtf,
  replaceSfntNameTable,
} from "../../../test/sfnt-test-helpers.ts";

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
const jostItalic = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Italic.ttf", import.meta.url),
  ),
);
const notoArabic = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/NotoNaskhArabic.ttf", import.meta.url),
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

const resolveText = createResolver([slideBackgroundCapability, textCapability]);

describe("Font OS/2 panose and pitchFamily metadata emission", () => {
  it("emits panose and pitchFamily on <a:latin>, <a:ea>, and <p:font> for static Font values including substitute: true", () => {
    const cffBytes = createMinimalCffOtf({ familyName: "Synthetic Sans" });
    const substituteCff = font(
      "Alias Sans",
      [{ weight: 400, style: "normal", data: cffBytes }],
      { substitute: true },
    );
    const woff2Substitute = font(
      "Noto Sans CJK SC",
      [{ weight: 400, style: "normal", data: notoSansScWoff2 }],
      { substitute: true },
    );
    const embeddableTtf = font("Jost*", [
      { weight: 400, style: "normal", data: jostBook },
    ]);

    const doc = deck({
      width: 640,
      height: 360,
      theme: defineTheme({
        name: "Metadata Theme",
        fonts: {
          major: substituteCff,
          minor: woff2Substitute,
        },
      }),
    }).append(
      slide({ id: "s1" }).append(
        text("Substitute CFF", {
          id: "t1",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 240,
            height: 40,
            fontFamily: substituteCff,
          },
        }),
        text("Embeddable Jost", {
          id: "t2",
          style: {
            position: "absolute",
            left: 20,
            top: 80,
            width: 240,
            height: 40,
            fontFamily: embeddableTtf,
          },
        }),
      ),
    );

    const emitted = emit(resolveText(doc));
    const parts = unzipSync(emitted.pptx);
    const themeXml = strFromU8(parts["ppt/theme/theme1.xml"]!);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);

    const embedDoc = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1" }).append(
        text("Embeddable Jost", {
          id: "t2",
          style: {
            position: "absolute",
            left: 20,
            top: 80,
            width: 240,
            height: 40,
            fontFamily: embeddableTtf,
          },
        }),
      ),
    );
    const embeddedPptx = emit(resolveText(embedDoc), { embedFonts: "subset" });
    const presentationXml = strFromU8(
      unzipSync(embeddedPptx.pptx)["ppt/presentation.xml"]!,
    );

    assert.match(
      themeXml,
      /<a:majorFont><a:latin typeface="Alias Sans" panose="020B0000000000000000" pitchFamily="34"\/><a:ea typeface="Alias Sans" panose="020B0000000000000000" pitchFamily="34"\/><a:cs typeface="Alias Sans"\/><\/a:majorFont>/u,
    );
    assert.match(
      themeXml,
      /<a:minorFont><a:latin typeface="Noto Sans CJK SC" panose="020B0500000000000000" pitchFamily="34"\/><a:ea typeface="Noto Sans CJK SC" panose="020B0500000000000000" pitchFamily="34"\/><a:cs typeface="Noto Sans CJK SC"\/><\/a:minorFont>/u,
    );
    assert.match(
      slideXml,
      /<a:latin typeface="Alias Sans" panose="020B0000000000000000" pitchFamily="34"\/><a:ea typeface="Alias Sans" panose="020B0000000000000000" pitchFamily="34"\/><a:cs typeface="Alias Sans"\/>/u,
    );
    assert.match(
      slideXml,
      /<a:latin typeface="Jost\*" panose="00000000000000000000" pitchFamily="2"\/><a:ea typeface="Jost\*" panose="00000000000000000000" pitchFamily="2"\/><a:cs typeface="Jost\*"\/>/u,
    );
    assert.match(
      presentationXml,
      /<p:font typeface="Jost\*" panose="00000000000000000000" pitchFamily="2" charset="0"\/>/u,
    );
  });
});

describe("fromFontCollection (.ttc / .otc) unpacking", () => {
  it("unpacks matching static faces from single-family and multi-family TTC/OTC containers into sealed Font values", () => {
    assert.equal(fromFontCollection, fromFontCollectionDirect);
    assert.equal(fromFontCollection, fromFontCollectionText);

    // Multi-family TTC: Jost* (Regular 400 + Bold 700, plus Name ID 16 Jost on Italic) + Other Family (CFF)
    const jostItalicNamed = replaceSfntNameTable(jostItalic, [
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0409,
        nameID: 1,
        value: "Jost",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0409,
        nameID: 2,
        value: "Italic",
      },
      {
        platformID: 3,
        encodingID: 1,
        languageID: 0x0409,
        nameID: 4,
        value: "Jost Italic",
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
        nameID: 16,
        value: "Jost*",
      },
    ]);
    const otherCff = createMinimalCffOtf({
      familyName: "Other CFF Family",
      weightClass: 400,
    });
    const multiTtc = buildTestTtc(
      [jostBook, jostBold, jostItalicNamed, otherCff],
      2,
    );

    const extractedJost = fromFontCollection(multiTtc, "Jost*");
    assert.equal(isSealedFont(extractedJost), true);
    assert.equal(extractedJost.family, "Jost*");
    assert.equal(extractedJost.faces?.length, 3);
    assert.deepEqual(
      extractedJost.faces?.map((f) => ({ weight: f.weight, style: f.style })),
      [
        { weight: 400, style: "normal" },
        { weight: 700, style: "normal" },
        { weight: 400, style: "italic" },
      ],
    );

    // Defensive getter isolation + zero-copy private bytes
    const firstFace = extractedJost.faces![0]!;
    assert.equal(hasFontFaceBytes(firstFace), true);
    const peeked = peekFontFaceResource(firstFace);
    assert.ok(peeked !== undefined);
    const copy1 = firstFace.data!;
    const copy2 = firstFace.data!;
    assert.notEqual(copy1, copy2);
    copy1[0] = 0xff;
    assert.notEqual(peeked.bytes[0], 0xff);

    // Extracted standalone SFNTs work seamlessly with resolve + subset emit
    const jostSubsettable = fromFontCollection(
      buildTestTtc([jostBook, jostBold]),
      "Jost*",
    );
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1" }).append(
        text("Subset TTC", {
          id: "t1",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 300,
            height: 40,
            fontFamily: jostSubsettable,
          },
        }),
      ),
    );
    const emitted = emit(resolveText(doc), { embedFonts: "subset" });
    assert.ok(
      Object.keys(unzipSync(emitted.pptx)).some((k) => k.endsWith(".fntdata")),
    );

    // Single-family collection with substitute: true accepts substitute family alias
    const singleFamilyTtc = buildTestTtc([jostBook, jostBold]);
    const aliased = fromFontCollection(singleFamilyTtc, "Alias Family", {
      substitute: true,
    });
    assert.equal(aliased.family, "Alias Family");
    assert.equal(aliased.substitute, true);
    assert.equal(aliased.faces?.length, 2);

    // OTC container of CFF faces
    const otc = buildTestTtc([
      createMinimalCffOtf({ familyName: "Collection Sans", weightClass: 400 }),
      createMinimalCffOtf({
        familyName: "Collection Sans",
        weightClass: 700,
        bold: true,
      }),
    ]);
    const extractedOtc = fromFontCollection(otc, "Collection Sans");
    assert.equal(extractedOtc.faces?.length, 2);
  });

  it("rejects corrupt containers, missing families, multi-family substitute ambiguity, fvar collections, and raw ttcf passed to font()", () => {
    assert.throws(
      () => fromFontCollection(jostBook, "Jost*"),
      (err: unknown) =>
        err instanceof CapabilityError &&
        err.code === "unsupported-font-format",
    );
    assert.throws(
      () =>
        fromFontCollection(
          new Uint8Array([0x74, 0x74, 0x63, 0x66, 0, 1]),
          "Jost*",
        ),
      (err: unknown) =>
        err instanceof CapabilityError &&
        err.code === "unsupported-font-format",
    );

    const singleTtc = buildTestTtc([jostBook]);
    assert.throws(
      () => fromFontCollection(singleTtc, "Nonexistent Family"),
      (err: unknown) =>
        err instanceof CapabilityError && err.code === "font-family-mismatch",
    );

    const multiTtc = buildTestTtc([
      jostBook,
      createMinimalCffOtf({ familyName: "Other Sans" }),
    ]);
    assert.throws(
      () => fromFontCollection(multiTtc, "Unknown Alias", { substitute: true }),
      (err: unknown) =>
        err instanceof CapabilityError && err.code === "font-family-mismatch",
    );

    // NotoNaskhArabic.ttf has an fvar table
    const fvarTtc = buildTestTtc([notoArabic]);
    assert.throws(
      () => fromFontCollection(fvarTtc, "Noto Naskh Arabic"),
      (err: unknown) =>
        err instanceof CapabilityError &&
        err.code === "unsupported-font-format",
    );

    // Passing raw ttcf directly to font() still rejects in resolve()
    const rawTtcFont = font("Jost*", [{ weight: 400, data: singleTtc }]);
    const badDoc = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1" }).append(
        text("Should fail", {
          id: "t1",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 200,
            height: "auto",
            fontFamily: rawTtcFont,
          },
        }),
      ),
    );
    assert.throws(
      () => resolveText(badDoc),
      (err: unknown) =>
        err instanceof BuildError &&
        err.diagnostics.some((d) => d.code === "font-unavailable"),
    );
  });
});
