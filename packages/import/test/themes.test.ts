import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonNode,
  DeckContainer,
  defaultColorMap,
  isTheme,
  type Theme,
  type ThemeOverride,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import {
  A_NS,
  OPC_REL_NS,
  P_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  createImporter,
  type ImportDiagnostic,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

import { testMastersCap } from "./test-masters-cap.ts";

describe("Native themes, color maps, and theme overrides import", () => {
  it("imports themes.pptx with 7 distinct Theme instances, sysClr values, inverted clrMap, and overrideClrMapping", () => {
    const bytes = new Uint8Array(
      readFileSync(new URL("./fixtures/themes.pptx", import.meta.url)),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.ok(result.document instanceof DeckContainer);
    assert.equal(result.report.slideCount, 10);
    assert.equal(result.document.masters.length, 7);

    const masters = result.document.masters;
    const themes = masters.map((m) => m.theme as Theme);
    for (const theme of themes) {
      assert.equal(isTheme(theme), true);
    }
    // Equal-value Office palettes (master1, master2, master4, master6) are distinct OPC parts -> distinct Theme instances.
    assert.notEqual(themes[0], themes[1]);
    assert.notEqual(themes[1], themes[3]);
    assert.notEqual(themes[3], themes[5]);
    assert.equal(themes[0]!.colors.dk1, "#000000");
    assert.equal(themes[0]!.colors.lt1, "#FFFFFF");
    assert.equal(themes[1]!.colors.dk1, "#000000");
    assert.equal(themes[2]!.name, "Austin");
    assert.equal(themes[2]!.colors.accent1, "#94C600");
    assert.equal(themes[4]!.name, "Trek");
    assert.equal(themes[4]!.colors.accent1, "#F0A22E");
    assert.equal(themes[6]!.name, "Tradeshow");
    assert.equal(themes[6]!.colors.dk1, "#3F3F3F");

    // Format scheme parsed including bgFillStyles (used by p:bgRef).
    assert.ok((themes[0]!.formatScheme?.bgFillStyles?.length ?? 0) >= 3);
    assert.ok((themes[6]!.formatScheme?.bgFillStyles?.length ?? 0) >= 3);

    // Mutating master 2's theme only updates master 2.
    themes[1]!.set({ colors: { accent1: "#FF0000" } });
    assert.equal(themes[1]!.colors.accent1, "#FF0000");
    assert.equal(themes[0]!.colors.accent1, "#4F81BD");
    assert.equal(themes[3]!.colors.accent1, "#4F81BD");

    // master7 has inverted bg1="dk1", tx1="lt1", bg2="dk2", tx2="lt2".
    assert.deepEqual(
      { ...(masters[6]!.props["colorMap"] as Record<string, string>) },
      {
        ...defaultColorMap,
        bg1: "dk1",
        tx1: "lt1",
        bg2: "dk2",
        tx2: "lt2",
      },
    );

    // slideLayout6 (master1 index 5) and slideLayout47 (master5 index 2) carry explicit overrideClrMapping.
    assert.deepEqual(
      {
        ...(masters[0]!.layouts[5]?.props["colorMapOverride"] as Record<
          string,
          string
        >),
      },
      {
        ...defaultColorMap,
      },
    );
    assert.deepEqual(
      {
        ...(masters[4]!.layouts[2]?.props["colorMapOverride"] as Record<
          string,
          string
        >),
      },
      {
        ...defaultColorMap,
        bg1: "dk1",
        tx1: "lt1",
        bg2: "dk2",
        tx2: "lt2",
      },
    );

    // Zero unconsumed XML diagnostics inside theme, clrMap, or clrMapOvr elements;
    // unsupported 3D effect elements inside theme1..7 produce localized unsupported-mechanism diagnostics.
    const themeOrColorMapUnconsumed = result.report.diagnostics.filter(
      (d) =>
        d.code === "unconsumed-xml" &&
        (d.location.part.startsWith("ppt/theme/") ||
          d.location.xpath?.includes("clrMap") === true),
    );
    assert.deepEqual(themeOrColorMapUnconsumed, []);
    const themeUnsupported = result.report.diagnostics.filter(
      (d) =>
        d.code === "unsupported-mechanism" &&
        d.location.part.startsWith("ppt/theme/"),
    );
    assert.ok(themeUnsupported.length > 0);
    assert.ok(
      themeUnsupported.every(
        (d) => d.disposition === "omitted" && d.category === "unsupported",
      ),
    );
  });

  it("imports placeholder-layout-color.pptx with non-default master clrMap and Simple Light theme", () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/placeholder-layout-color.pptx", import.meta.url),
      ),
    );
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.ok(result.document instanceof DeckContainer);
    assert.equal(result.document.masters.length, 1);
    const master = result.document.masters[0]!;
    const theme = master.theme as Theme;
    assert.equal(isTheme(theme), true);
    assert.equal(theme.name, "Modern Geometric Pitch Deck by Slidesgo");
    assert.equal(theme.colors.dk1, "#2D2842");
    assert.equal(theme.colors.lt1, "#FFFFFF");
    assert.deepEqual(
      { ...(master.props["colorMap"] as Record<string, string>) },
      {
        ...defaultColorMap,
        bg2: "dk2",
        tx2: "lt2",
      },
    );
  });

  it("deduplicates shared theme OPC parts to the exact same Theme instance and reads themeOverride parts", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/><p:sldMasterId id="2147483649" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId3"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster2.xml"/><Relationship Id="rId3" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(
        `<a:theme xmlns:a="${A_NS}" name="Shared Theme"><a:themeElements><a:clrScheme name="Shared"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="654321"/></a:accent2><a:accent3><a:srgbClr val="112233"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4><a:accent5><a:srgbClr val="778899"/></a:accent5><a:accent6><a:srgbClr val="AABBCC"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Shared"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface="Georgia"/><a:cs typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`,
      ),
      "ppt/theme/themeOverride1.xml": strToU8(
        `<a:themeOverride xmlns:a="${A_NS}"><a:clrScheme name="Ovr"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="059669"/></a:accent1><a:accent2><a:srgbClr val="654321"/></a:accent2><a:accent3><a:srgbClr val="112233"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4><a:accent5><a:srgbClr val="778899"/></a:accent5><a:accent6><a:srgbClr val="AABBCC"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeOverride>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster2.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M2"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483651" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout2.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L2"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster2.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
    assert.ok(result.document instanceof DeckContainer);
    const [m1, m2] = result.document.masters;
    assert.ok(m1?.theme);
    assert.ok(m2?.theme);
    assert.equal(m1.theme, m2.theme);
    (m1.theme as Theme).set({ colors: { accent1: "#AA1122" } });
    assert.equal((m2.theme as Theme).colors.accent1, "#AA1122");

    const slide1 = result.document.children.find(
      (node): node is CommonNode =>
        node instanceof CommonNode && node.kind === "slide",
    );
    assert.ok(slide1);
    assert.deepEqual(
      { ...(slide1.props["colorMapOverride"] as Record<string, string>) },
      {
        ...defaultColorMap,
        bg1: "dk1",
        tx1: "lt1",
      },
    );
    assert.equal(
      (slide1.props["themeOverride"] as { colors?: Record<string, string> })
        ?.colors?.["accent1"],
      "#059669",
    );
  });

  it("omits reserved font placeholders in theme1.xml and themeOverride1.xml with structured diagnostics so resolve succeeds", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(
        `<a:theme xmlns:a="${A_NS}" name="Placeholder Font Theme"><a:themeElements><a:clrScheme name="Base"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="654321"/></a:accent2><a:accent3><a:srgbClr val="112233"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4><a:accent5><a:srgbClr val="778899"/></a:accent5><a:accent6><a:srgbClr val="AABBCC"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="BadFonts"><a:majorFont><a:latin typeface="+mj-lt"/></a:majorFont><a:minorFont><a:latin typeface="minor"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
      ),
      "ppt/theme/themeOverride1.xml": strToU8(
        `<a:themeOverride xmlns:a="${A_NS}"><a:fontScheme name="BadOvrFonts"><a:majorFont><a:latin typeface="major"/></a:majorFont><a:minorFont><a:latin typeface="+mn-lt"/></a:minorFont></a:fontScheme></a:themeOverride>`,
      ),
      "ppt/theme/themeOverride2.xml": strToU8(
        `<a:themeOverride xmlns:a="${A_NS}"><a:clrScheme name="Ovr"><a:accent1><a:srgbClr val="059669"/></a:accent1></a:clrScheme><a:fontScheme name="MixedOvrFonts"><a:majorFont><a:latin typeface="+mj-ea"/></a:majorFont><a:minorFont><a:latin typeface="Georgia"/></a:minorFont></a:fontScheme></a:themeOverride>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride2.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.ok(result.document instanceof DeckContainer);
    const theme = result.document.masters[0]?.theme as Theme;
    assert.equal(theme.fonts.major, "Aptos Display");
    assert.equal(theme.fonts.minor, "Aptos");

    const layout1 = result.document.masters[0]?.layouts[0];
    assert.equal(layout1?.props["themeOverride"], undefined);

    const slide1 = result.document.children.find(
      (node): node is CommonNode =>
        node instanceof CommonNode && node.kind === "slide",
    );
    const slideOverride = slide1?.props["themeOverride"] as
      | ThemeOverride
      | undefined;
    assert.equal(slideOverride?.colors?.accent1, "#059669");
    assert.equal(slideOverride?.fonts?.major, undefined);
    assert.equal(slideOverride?.fonts?.minor, "Georgia");

    const fontDiagnostics = result.report.diagnostics.filter(
      (
        d,
      ): d is Extract<
        ImportDiagnostic,
        { readonly code: "invalid-xml-value" }
      > =>
        d.code === "invalid-xml-value" &&
        d.attributeOrElement === "a:latin/@typeface",
    );
    assert.equal(fontDiagnostics.length, 5);
    assert.ok(
      fontDiagnostics.every(
        (d) => d.category === "corrupt-input" && d.disposition === "omitted",
      ),
    );
    assert.deepEqual(
      fontDiagnostics.map((d) => [d.location.part, d.rawValue]),
      [
        ["ppt/theme/theme1.xml", "+mj-lt"],
        ["ppt/theme/theme1.xml", "minor"],
        ["ppt/theme/themeOverride1.xml", "major"],
        ["ppt/theme/themeOverride1.xml", "+mn-lt"],
        ["ppt/theme/themeOverride2.xml", "+mj-ea"],
      ],
    );

    const resolveImported = createResolver([testMastersCap]);
    const resolved = resolveImported(result.document);
    assert.equal(resolved.slides.length, 1);
  });

  it("imports independent majorEastAsian/minorEastAsian while omitting redundant EA slots when ea matches latin", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(
        `<a:theme xmlns:a="${A_NS}" name="DualTheme"><a:themeElements><a:clrScheme name="Base"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="654321"/></a:accent2><a:accent3><a:srgbClr val="112233"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4><a:accent5><a:srgbClr val="778899"/></a:accent5><a:accent6><a:srgbClr val="AABBCC"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Dual"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="PingFang SC" panose="020B0600000000000000" pitchFamily="34"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Noto Sans CJK SC"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
      ),
      "ppt/theme/themeOverride1.xml": strToU8(
        `<a:themeOverride xmlns:a="${A_NS}"><a:fontScheme name="DualOvr"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Inter"/><a:ea typeface="Noto Serif CJK SC"/></a:minorFont></a:fontScheme></a:themeOverride>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" Target="../theme/themeOverride1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.ok(result.document instanceof DeckContainer);
    const theme = result.document.masters[0]?.theme as Theme;
    assert.equal(theme.fonts.major, "Aptos Display");
    assert.equal(theme.fonts.minor, "Aptos");
    assert.equal(theme.fonts.majorEastAsian, "PingFang SC");
    assert.equal(theme.fonts.minorEastAsian, "Noto Sans CJK SC");

    const slide1 = result.document.children.find(
      (node): node is CommonNode =>
        node instanceof CommonNode && node.kind === "slide",
    );
    const slideOverride = slide1?.props["themeOverride"] as ThemeOverride;
    assert.equal(slideOverride.fonts?.major, "Georgia");
    assert.equal(slideOverride.fonts?.majorEastAsian, undefined);
    assert.equal(slideOverride.fonts?.minor, "Inter");
    assert.equal(slideOverride.fonts?.minorEastAsian, "Noto Serif CJK SC");
    assert.equal(
      result.report.diagnostics.filter((d) => d.code === "unconsumed-xml")
        .length,
      0,
    );
  });
});
