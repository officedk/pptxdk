import assert from "node:assert/strict";

import {
  A_NS,
  OPC_REL_NS,
  P_NS,
  R_NS,
  REL_FONT,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_MASTER,
  REL_THEME,
  openPptxPackage,
  readEmbeddedFontList,
  readPresentationThemeFontScheme,
  unpackEotFontData,
  type ImportDiagnostic,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

function buildSyntheticTtf(): Uint8Array {
  // Minimal SFNT header with TrueType signature 0x00010000 and 1 dummy table record.
  const bytes = new Uint8Array(64);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000, false);
  view.setUint16(4, 1, false);
  return bytes;
}

function wrapInEot(
  payload: Uint8Array,
  options: {
    readonly version?: number;
    readonly flags?: number;
    readonly weight?: number;
    readonly italic?: number;
    readonly magic?: number;
  } = {},
): Uint8Array {
  const headerSize = 114;
  const totalSize = headerSize + payload.byteLength;
  const out = new Uint8Array(totalSize);
  const view = new DataView(out.buffer);
  view.setUint32(0, totalSize, true);
  view.setUint32(4, payload.byteLength, true);
  view.setUint32(8, options.version ?? 0x00020002, true);
  view.setUint32(12, options.flags ?? 0, true);
  out[27] = options.italic ?? 0;
  view.setUint32(28, options.weight ?? 400, true);
  view.setUint16(34, options.magic ?? 0x504c, true);
  view.setUint32(headerSize - 18, 0x50475342, true);
  out.set(payload, headerSize);
  return out;
}

function buildFontPackage(options: {
  readonly embeddedFontLstXml: string;
  readonly presRelsExtraXml?: string;
  readonly fontParts?: Readonly<Record<string, Uint8Array>>;
  readonly themeFonts?: { readonly major: string; readonly minor: string };
}): Uint8Array {
  const entries: Record<string, Uint8Array> = {
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/>
</Relationships>`,
    ),
    "ppt/presentation.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" embedTrueTypeFonts="1" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  ${options.embeddedFontLstXml}
</p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
  ${options.presRelsExtraXml ?? ""}
</Relationships>`,
    ),
    "ppt/slideMasters/slideMaster1.xml": strToU8(
      `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"/>`,
    ),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_THEME}" Target="../theme/theme1.xml"/>
</Relationships>`,
    ),
    "ppt/theme/theme1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A_NS}" name="Custom">
  <a:themeElements>
    <a:clrScheme name="Custom"/>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="${options.themeFonts?.major ?? "Aptos Display"}"/></a:majorFont>
      <a:minorFont><a:latin typeface="${options.themeFonts?.minor ?? "Aptos"}"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`,
    ),
    "ppt/slides/slide1.xml": strToU8(
      `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree/></p:cSld></p:sld>`,
    ),
  };
  for (const [part, bytes] of Object.entries(options.fontParts ?? {})) {
    entries[part] = bytes;
  }
  return zipSync(entries);
}

describe("@pptxdk/import embedded fonts and EOT unpacking", () => {
  it("unpacks uncompressed EOT payloads into detached TrueType bytes and flags compressed/obfuscated EOT cleanly", () => {
    const ttf = buildSyntheticTtf();
    const eot = wrapInEot(ttf, { weight: 700, italic: 1 });
    const eotCopy = new Uint8Array(eot);
    const ok = unpackEotFontData(eotCopy);
    eotCopy.fill(0);

    assert.equal(ok.status, "ok");
    if (ok.status === "ok") {
      assert.deepEqual(ok.ttfBytes, ttf);
      assert.equal(ok.eotWeight, 700);
      assert.equal(ok.eotItalic, true);
    }

    const compressed = unpackEotFontData(wrapInEot(ttf, { flags: 0x00000004 }));
    assert.equal(compressed.status, "unsupported");
    if (compressed.status === "unsupported") {
      assert.equal(compressed.reason, "compressed-eot");
    }

    const obfuscated = unpackEotFontData(wrapInEot(ttf, { flags: 0x10000000 }));
    assert.equal(obfuscated.status, "unsupported");
    if (obfuscated.status === "unsupported") {
      assert.equal(obfuscated.reason, "obfuscated-eot");
    }

    const badMagic = unpackEotFontData(wrapInEot(ttf, { magic: 0x1234 }));
    assert.equal(badMagic.status, "corrupt");
  });

  it("reads <p:embeddedFontLst>, theme fontScheme, and reports missing/broken font relationships without aborting", () => {
    const ttf = buildSyntheticTtf();
    const fntBytes = wrapInEot(ttf);
    const zipBytes = buildFontPackage({
      embeddedFontLstXml: `
        <p:embeddedFontLst>
          <p:embeddedFont>
            <p:font typeface="Jost*" pitchFamily="34" charset="0"/>
            <p:regular r:id="rId10"/>
            <p:bold r:id="rId11"/>
            <p:italic r:id="rId999"/>
          </p:embeddedFont>
        </p:embeddedFontLst>`,
      presRelsExtraXml: `
        <Relationship Id="rId10" Type="${REL_FONT}" Target="fonts/font1.fntdata"/>
        <Relationship Id="rId11" Type="${REL_FONT}" Target="fonts/missing.fntdata"/>`,
      fontParts: { "ppt/fonts/font1.fntdata": fntBytes },
      themeFonts: { major: "Jost*", minor: "Jost*" },
    });

    const pkg = openPptxPackage(zipBytes);
    zipBytes.fill(0);

    const diagnostics: ImportDiagnostic[] = [];
    const records = readEmbeddedFontList(
      pkg.presentationRoot,
      pkg.presentationRels,
      pkg,
      diagnostics,
    );
    assert.equal(records.length, 1);
    assert.equal(records[0]?.typeface, "Jost*");
    assert.equal(records[0]?.slots.length, 1);
    assert.equal(records[0]?.slots[0]?.slot, "regular");
    assert.deepEqual(records[0]?.slots[0]?.bytes, fntBytes);

    assert.ok(
      diagnostics.some(
        (d) => d.code === "missing-target-part" && d.relationshipId === "rId11",
      ),
    );
    assert.ok(
      diagnostics.some(
        (d) =>
          d.code === "broken-relationship" && d.relationshipId === "rId999",
      ),
    );

    const scheme = readPresentationThemeFontScheme(pkg);
    assert.equal(scheme?.themeName, "Custom");
    assert.equal(scheme?.majorTypeface, "Jost*");
    assert.equal(scheme?.minorTypeface, "Jost*");
  });
});
