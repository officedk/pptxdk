import assert from "node:assert/strict";

import {
  CommonNode,
  DeckContainer,
  SlideMasterContainer,
  type ResolvedSlideLayout,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { createSlideNode } from "@pptxdk/core/document";
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
  XmlElementCursor,
  createImporter,
  parseXmlPart,
  readLineBorder,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

import { testMastersCap } from "./test-masters-cap.ts";

const minimalThemeXml = `<a:theme xmlns:a="${A_NS}" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="O"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

describe("SlideMaster and SlideLayout import", () => {
  it("imports a 0-slide multi-master deck with preserve distinction and idle layouts", () => {
    const bytes = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/><p:sldMasterId id="2147483649" r:id="rId2"/></p:sldMasterIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster2.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(minimalThemeXml),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster2.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" preserve="1"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483651" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout2.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster2.xml"/></Relationships>`,
      ),
    });
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 0);
    assert.ok(result.document instanceof DeckContainer);
    assert.equal(result.document.masters.length, 2);

    const [master1, master2] = result.document.masters;
    assert.ok(master1 instanceof SlideMasterContainer);
    assert.ok(master2 instanceof SlideMasterContainer);
    assert.equal(master1.id, "master-1");
    assert.equal(master1.props["name"], undefined);
    assert.equal(master1.props["preserve"], undefined);
    assert.equal(master1.layouts.length, 1);
    assert.equal(master1.layouts[0]?.id, "Blank");
    assert.equal(master1.layouts[0]?.props["name"], "Blank");
    assert.equal(master1.layouts[0]?.props["type"], "blank");
    assert.equal(master1.layouts[0]?.props["preserve"], true);

    assert.equal(master2.id, "master-2");
    assert.equal(master2.props["name"], undefined);
    assert.equal(master2.props["preserve"], true);
    assert.equal(master2.layouts.length, 1);
    assert.equal(master2.layouts[0]?.id, "Blank");
    assert.equal(master2.layouts[0]?.props["name"], "Blank");
    assert.equal(master2.layouts[0]?.props["type"], "blank");
    assert.equal(master2.layouts[0]?.props["preserve"], true);

    // Resolve the 0-slide deck as-is.
    const resolve = createResolver([testMastersCap]);
    const zeroSlideResolution = resolve(result.document);
    assert.equal(zeroSlideResolution.report.slides.length, 0);
    assert.equal(zeroSlideResolution.report.masters?.length, 2);
    assert.equal(zeroSlideResolution.report.masters?.[0]?.preserve, undefined);
    assert.equal(zeroSlideResolution.report.masters?.[1]?.preserve, true);

    // Append a new slide referencing master 2's previously idle layout.
    const idleLayout = master2.layouts[0];
    result.document.append(
      createSlideNode({ id: "cover", layout: idleLayout }),
    );
    const withSlideResolution = resolve(result.document);
    assert.equal(withSlideResolution.report.slides.length, 1);
    assert.equal(withSlideResolution.report.slides[0]?.id, "cover");
    assert.equal(withSlideResolution.report.slides[0]?.layout, "Blank");
  });

  it("imports used and idle vertical layouts and disambiguates Comparison placeholder keys", () => {
    const bytes = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(minimalThemeXml),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/><p:sldLayoutId id="2147483651" r:id="rId2"/><p:sldLayoutId id="2147483652" r:id="rId3"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId3" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout3.xml"/><Relationship Id="rId4" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="twoTxTwoObj" preserve="1"><p:cSld name="Comparison"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Left Text"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Left Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="2"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="5" name="Right Text"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="3"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="6" name="Right Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="4"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="7" name="Date"/><p:cNvSpPr/><p:nvPr><p:ph type="dt" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="8" name="Footer"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="9" name="Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout2.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="vertTx" preserve="1"><p:cSld name="Title and Vertical Text"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout3.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="vertTitleAndTx" preserve="1"><p:cSld name="Vertical Title and Text"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout3.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 1);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    assert.equal(master.props["preserve"], undefined);
    assert.equal(master.layouts.length, 3);
    assert.deepEqual(
      master.layouts.map((layout) => layout.props["name"]),
      ["Comparison", "Title and Vertical Text", "Vertical Title and Text"],
    );
    for (const layout of master.layouts) {
      assert.equal(layout.props["preserve"], true);
    }

    const queriedIdle = result.document.query
      .get(master.id!)
      .get("Title and Vertical Text").node;
    assert.equal(queriedIdle, master.layouts[1]);

    const slides = result.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.kind === "slide",
    );
    assert.equal(slides[0]?.props["layout"], master.layouts[0]);

    const resolve = createResolver([testMastersCap]);
    const resolved = resolve(result.document);
    const resolvedLayouts: readonly ResolvedSlideLayout[] =
      resolved.report.masters?.[0]?.layouts ?? [];
    assert.equal(resolvedLayouts.length, 3);
    assert.equal(resolvedLayouts[1]?.type, "vertTx");
    assert.equal(resolvedLayouts[2]?.type, "vertTitleAndTx");

    assert.ok(master.props["textStyles"] !== undefined);
    assert.ok(master.props["placeholders"] !== undefined);
    const comparisonPlaceholders = master.layouts[0]?.props["placeholders"] as
      | Readonly<
          Record<string, { readonly type: string; readonly idx?: number }>
        >
      | undefined;
    assert.ok(comparisonPlaceholders !== undefined);
    assert.deepEqual(Object.keys(comparisonPlaceholders), [
      "title",
      "body-1",
      "content-2",
      "body-3",
      "content-4",
      "dt",
      "ftr",
      "sldNum",
    ]);
    assert.equal(comparisonPlaceholders["body-1"]?.type, "body");
    assert.equal(comparisonPlaceholders["body-1"]?.idx, 1);
    assert.equal(comparisonPlaceholders["content-2"]?.type, "obj");
    assert.equal(comparisonPlaceholders["content-2"]?.idx, 2);
    assert.equal(comparisonPlaceholders["body-3"]?.type, "body");
    assert.equal(comparisonPlaceholders["body-3"]?.idx, 3);
    assert.equal(comparisonPlaceholders["content-4"]?.type, "obj");
    assert.equal(comparisonPlaceholders["content-4"]?.idx, 4);
  });

  it("imports layout listStyles, disambiguated duplicate title keys, hasCustomPrompt, and master headerFooter flags", () => {
    const bytes = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(minimalThemeXml),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:hf dt="0" ftr="0" hdr="0" sldNum="0"/><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="title"><p:cSld name="DualTitle"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="3500"><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Title 2"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="2" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="8000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle><a:p/></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });
    const importer = createImporter([testMastersCap]);
    const result = importer(bytes);
    bytes.fill(0);

    assert.equal(result.report.slideCount, 1);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    assert.deepEqual(
      { ...(master.props["headerFooter"] as Record<string, unknown>) },
      {
        dateTime: false,
        footer: false,
        header: false,
        slideNumber: false,
      },
    );

    const layout = master.layouts[0]!;
    const layoutPlaceholders = layout.props["placeholders"] as
      | Readonly<
          Record<
            string,
            {
              readonly type: string;
              readonly idx?: number;
              readonly hasCustomPrompt?: boolean;
              readonly listStyles?: Readonly<
                Record<
                  number,
                  {
                    readonly fontSize?: string;
                    readonly color?: string;
                    readonly fontFamily?: string;
                  }
                >
              >;
            }
          >
        >
      | undefined;
    assert.ok(layoutPlaceholders !== undefined);
    assert.deepEqual(Object.keys(layoutPlaceholders), ["title", "title-2"]);
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.fontSize,
      "35pt",
    );
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.color,
      "#00FFFF",
    );
    assert.equal(
      layoutPlaceholders["title"]?.listStyles?.[1]?.fontFamily,
      "major",
    );
    assert.equal(layoutPlaceholders["title-2"]?.idx, 2);
    assert.equal(layoutPlaceholders["title-2"]?.hasCustomPrompt, true);
    assert.equal(
      layoutPlaceholders["title-2"]?.listStyles?.[1]?.fontSize,
      "80pt",
    );
    assert.equal(
      layoutPlaceholders["title-2"]?.listStyles?.[1]?.color,
      "#FF0000",
    );
  });

  it("disambiguates duplicate display names deterministically within a master while preserving display names", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="Brand"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout2.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="cust" userDrawn="1"><p:cSld name="Shared Name"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout2.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Shared Name"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    assert.equal(master.id, "Brand");
    assert.equal(master.props["name"], "Brand");
    assert.equal(master.layouts.length, 2);

    const [first, second] = master.layouts;
    assert.equal(first?.id, "Shared Name");
    assert.equal(first?.props["name"], "Shared Name");
    assert.equal(first?.props["userDrawn"], true);
    assert.equal(second?.id, "Shared Name-2");
    assert.equal(second?.props["name"], "Shared Name");
    assert.equal(second?.props["preserve"], true);
  });

  it("imports placeholder geometry, flipH/flipV, background, border, and size onto unified style and size", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="Brand">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Master Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title" sz="half"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm flipH="1" flipV="1"><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="112233"/></a:solidFill>
          <a:ln w="19050"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:prstDash val="dash"/></a:ln>
        </p:spPr>
        <p:txBody><a:bodyPr wrap="none" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="title">
  <p:cSld name="Hero">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Layout Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="457200" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>
          <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.deepEqual(result.report.diagnostics, []);
    assert.ok(result.document instanceof DeckContainer);
    const master = result.document.masters[0]!;
    const masterPlaceholders = master.props["placeholders"] as Record<
      string,
      Record<string, unknown>
    >;
    const titlePh = masterPlaceholders["title"]!;
    assert.equal(titlePh["size"], "half");
    assert.equal(titlePh["sz"], undefined);
    assert.equal(titlePh["x"], undefined);
    assert.equal(titlePh["y"], undefined);
    assert.equal(titlePh["width"], undefined);
    assert.equal(titlePh["height"], undefined);
    assert.equal(titlePh["flipH"], undefined);
    assert.equal(titlePh["flipV"], undefined);
    assert.deepEqual(JSON.parse(JSON.stringify(titlePh["style"])), {
      position: "absolute",
      left: 96,
      top: 48,
      width: 192,
      height: 96,
      background: "#112233",
      border: { color: "#AABBCC", width: 2, style: "dashed" },
      flipH: true,
      flipV: true,
      textWrap: "nowrap",
      verticalAlign: "middle",
    });

    const layout = master.layouts[0]!;
    const layoutPlaceholders = layout.props["placeholders"] as Record<
      string,
      Record<string, unknown>
    >;
    const bodyPh = layoutPlaceholders["body"]!;
    assert.equal(bodyPh["size"], "quarter");
    assert.equal(bodyPh["sz"], undefined);
    assert.equal(bodyPh["x"], undefined);
    assert.equal(bodyPh["y"], undefined);
    assert.equal(bodyPh["width"], undefined);
    assert.equal(bodyPh["height"], undefined);
    assert.deepEqual(JSON.parse(JSON.stringify(bodyPh["style"])), {
      position: "absolute",
      left: 48,
      top: 96,
      width: 96,
      height: 48,
      borderRadius: 12,
      background: {
        type: "linear-gradient",
        angle: 180,
        stops: [
          { color: "#FF0000", position: 0 },
          { color: "#0000FF", position: 1 },
        ],
      },
    });
  });

  it("readLineBorder distinguishes <a:noFill/> from omitted line fill", () => {
    const noFillCursor = XmlElementCursor.createRoot(
      parseXmlPart("test.xml", `<a:ln xmlns:a="${A_NS}"><a:noFill/></a:ln>`),
      "test.xml",
    );
    assert.deepEqual(readLineBorder(noFillCursor, { part: "test.xml" }, []), {
      noFill: true,
    });

    const emptyLnCursor = XmlElementCursor.createRoot(
      parseXmlPart("test.xml", `<a:ln xmlns:a="${A_NS}" w="12700"/>`),
      "test.xml",
    );
    assert.deepEqual(
      readLineBorder(emptyLnCursor, { part: "test.xml" }, []),
      {},
    );
  });

  it("reads orthogonal eastAsianFontFamily and differential <a:ea>-only slots in master defRPr without polluting equal single-script styles", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Jost"/><a:ea typeface="Noto Sans CJK SC"/></a:defRPr></a:lvl1pPr><a:lvl2pPr><a:defRPr><a:ea typeface="Only EA Face"/></a:defRPr></a:lvl2pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([testMastersCap]);
    const result = importer(pptx);
    assert.ok(result.document instanceof DeckContainer);
    const textStyles = result.document.masters[0]?.props["textStyles"] as {
      titleStyle?: Record<
        number,
        { fontFamily?: string; eastAsianFontFamily?: string }
      >;
      bodyStyle?: Record<
        number,
        { fontFamily?: string; eastAsianFontFamily?: string }
      >;
    };
    assert.equal(textStyles.titleStyle?.[1]?.fontFamily, "major");
    assert.equal(textStyles.titleStyle?.[1]?.eastAsianFontFamily, undefined);
    assert.equal(textStyles.bodyStyle?.[1]?.fontFamily, "Jost");
    assert.equal(
      textStyles.bodyStyle?.[1]?.eastAsianFontFamily,
      "Noto Sans CJK SC",
    );
    assert.equal(textStyles.bodyStyle?.[2]?.fontFamily, undefined);
    assert.equal(
      textStyles.bodyStyle?.[2]?.eastAsianFontFamily,
      "Only EA Face",
    );
  });
});
