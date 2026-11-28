import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { font } from "@pptxdk/capabilities/font";
import { createCodegen } from "@pptxdk/codegen";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode, DeckContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import {
  A_NS,
  C_NS,
  DIAGRAM_NS,
  OPC_REL_NS,
  P_NS,
  REL_CHART,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const importAll = createImporter(documentCapabilities);
const resolveAll = createResolver(documentCapabilities);
const codegenAll = createCodegen(documentCapabilities);

function baseDeckFiles(args: {
  readonly slideXml: string;
  readonly slideRelsXml?: string;
  readonly extraFiles?: Readonly<Record<string, string | Uint8Array>>;
}): Uint8Array {
  return zipSync({
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "ppt/presentation.xml": strToU8(
      `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
    ),
    "ppt/theme/theme1.xml": strToU8(
      `<a:theme xmlns:a="${A_NS}" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Jost*"/></a:majorFont><a:minorFont><a:latin typeface="Jost*"/></a:minorFont></a:fontScheme><a:fmtScheme name="O"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
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
    "ppt/slides/slide1.xml": strToU8(args.slideXml),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(
      args.slideRelsXml ??
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
    ),
    ...Object.fromEntries(
      Object.entries(args.extraFiles ?? {}).map(([k, v]) => [
        k,
        typeof v === "string" ? strToU8(v) : v,
      ]),
    ),
  });
}

function assertCodegenAndEmitPass(
  document: ReturnType<typeof importAll>["document"],
  options?: { readonly attachJostFont?: boolean },
): void {
  const vanilla = codegenAll(document, { target: "vanilla" });
  const react = codegenAll(document, { target: "react" });
  assert.ok(Object.keys(vanilla.files).length > 0);
  assert.ok(Object.keys(react.files).length > 0);
  if (options?.attachJostFont && document instanceof DeckContainer) {
    for (const master of document.masters) {
      const theme = master.theme as { set?(patch: unknown): void } | undefined;
      theme?.set?.({
        fonts: {
          major: jostFont,
          minor: jostFont,
          majorEastAsian: jostFont,
          minorEastAsian: jostFont,
        },
      });
    }
  }
  const resolved = resolveAll(document);
  const emitted = emit(resolved);
  assert.ok(emitted.pptx.byteLength > 0);
}

function chartFrameXml(id: number, name: string, rId: string): string {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="457200"/><a:ext cx="4572000" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${rId}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

describe("importPptx postcondition invariants", () => {
  it("trims out-of-bounds chart dPt/dLbl indices, synthesizes missing categories, strips empty/axis-broken groups, and omits axisless broken charts", () => {
    const chart1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>S1</c:v></c:tx>
          <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="445566"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="9"/><c:spPr><a:solidFill><a:srgbClr val="778899"/></a:solidFill></c:spPr></c:dPt>
          <c:dLbls><c:dLbl><c:idx val="5"/><c:showVal val="1"/></c:dLbl></c:dLbls>
          <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:areaChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:v>BrokenArea</c:v></c:tx>
          <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="100"/><c:axId val="101"/>
      </c:areaChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
      <c:valAx><c:axId val="101"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="r"/><c:crossAx val="100"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

    const chart2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>OnlyBroken</c:v></c:tx>
          <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="50"/><c:axId val="51"/>
      </c:barChart>
      <c:dateAx><c:axId val="50"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="51"/></c:dateAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

    const chart3Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:v>MissingCat</c:v></c:tx>
          <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

    const pptx = baseDeckFiles({
      slideXml: `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${chartFrameXml(2, "Chart1", "rId2")}${chartFrameXml(3, "Chart2", "rId3")}${chartFrameXml(4, "Chart3", "rId4")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      slideRelsXml: `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_CHART}" Target="../charts/chart1.xml"/><Relationship Id="rId3" Type="${REL_CHART}" Target="../charts/chart2.xml"/><Relationship Id="rId4" Type="${REL_CHART}" Target="../charts/chart3.xml"/></Relationships>`,
      extraFiles: {
        "ppt/charts/chart1.xml": chart1Xml,
        "ppt/charts/chart2.xml": chart2Xml,
        "ppt/charts/chart3.xml": chart3Xml,
      },
    });

    const imported = importAll(pptx);
    const slideNode = imported.document.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    const charts = slideNode.children.filter(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "chart",
    );
    // Chart1 keeps only barChart (lineChart empty, areaChart missing catAx); Chart2 omitted entirely; Chart3 synthesizes ["1","2"].
    assert.equal(charts.length, 2);
    const c1Groups = charts[0]!.props["groups"] as readonly {
      readonly series: readonly { readonly points?: Record<number, unknown> }[];
    }[];
    assert.equal(c1Groups.length, 1);
    assert.deepEqual(Object.keys(c1Groups[0]!.series[0]!.points ?? {}), ["1"]);
    const c1Axes = charts[0]!.props["axes"] as Record<string, unknown>;
    assert.equal(Object.keys(c1Axes).length, 2);

    const c3Groups = charts[1]!.props["groups"] as readonly {
      readonly data: { readonly categories: readonly string[] };
      readonly series: readonly {
        readonly values: readonly (number | null)[];
      }[];
    }[];
    assert.deepEqual(c3Groups[0]!.data.categories, ["1", "2"]);
    assert.deepEqual(c3Groups[0]!.series[0]!.values, [10, 20]);

    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "invalid-xml-value" &&
          d.location.part === "ppt/charts/chart1.xml",
      ),
    );
    assert.ok(
      imported.report.diagnostics.some(
        (d) => d.location.part === "ppt/charts/chart2.xml",
      ),
    );

    assertCodegenAndEmitPass(imported.document);
  });

  it("sanitizes illegal fillColor/breathe targets, unsupported repeats, out-of-bounds pRg, and dangling master/layout timing spids", () => {
    // Master has one placeholder (spid=2) and one extra shape (spid=3), plus timing referencing valid spid=2 and dangling spid=99.
    const masterXml = `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="M1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TitlePh"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Master Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="300"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="8" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="9" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="99"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="10" dur="300"/><p:tgtEl><p:spTgt spid="99"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sldMaster>`;

    const slideXml = `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="10" name="NoFillShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="GradShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="2500000" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="445566"/></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="12" name="SolidBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1800000"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="13" name="SingleParaText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="2500000" y="1800000"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
        <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:t>Only one line</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="9" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="6" dur="400" fill="hold"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="111111"/></p:from><p:to><a:srgbClr val="FF0000"/></p:to></p:animClr>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="7" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond evt="end" delay="50"><p:tn val="5"/></p:cond></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="8" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="12"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="9" dur="300"/><p:tgtEl><p:spTgt spid="12"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="10" presetID="9" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="11" dur="400" fill="hold"/><p:tgtEl><p:spTgt spid="11"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="112233"/></p:from><p:to><a:srgbClr val="FFFFFF"/></p:to></p:animClr>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="12" presetID="9" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="13" dur="400" fill="hold"/><p:tgtEl><p:spTgt spid="13"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="000000"/></p:from><p:to><a:srgbClr val="FF0000"/></p:to></p:animClr>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="14" presetID="10" presetClass="entr" presetSubtype="0" repeatCount="2000" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="15" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="10"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="16" dur="300"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="17" presetID="8" presetClass="emph" presetSubtype="0" repeatCount="indefinite" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animRot by="21600000"><p:cBhvr><p:cTn id="18" dur="500"/><p:tgtEl><p:spTgt spid="12"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="19" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="20" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="13"><p:txEl><p:pRg st="5" end="5"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="21" dur="300"/><p:tgtEl><p:spTgt spid="13"><p:txEl><p:pRg st="5" end="5"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sld>`;

    const pptx = baseDeckFiles({
      slideXml,
      extraFiles: {
        "ppt/slideMasters/slideMaster1.xml": masterXml,
      },
    });
    const imported = importAll(pptx);

    // Master timing should drop spid=99 while keeping spid=2 ("title")
    const masterNode = (
      imported.document as { readonly masters?: readonly CommonNode[] }
    ).masters?.[0];
    assert.ok(
      masterNode !== undefined,
      `Expected masterNode; children=${JSON.stringify(imported.document.children.map((c) => (c instanceof CommonNode ? c.kind : c.type)))} diags=${JSON.stringify(imported.report.diagnostics)}`,
    );
    const ownerMotion = masterNode.props["motion"] as
      | {
          readonly steps: readonly {
            readonly effects: readonly { readonly target: unknown }[];
          }[];
        }
      | undefined;
    assert.ok(ownerMotion !== undefined);
    assert.equal(ownerMotion.steps.length, 1);
    assert.equal(ownerMotion.steps[0]!.effects.length, 1);
    assert.deepEqual(
      {
        ...(ownerMotion.steps[0]!.effects[0]!.target as Record<
          string,
          unknown
        >),
      },
      { placeholder: "title" },
    );

    // Slide objects: NoFillShape keeps fade (repeat stripped), SolidBox keeps spin (forever stripped, dependent fade stripped), Text drops out-of-range pRg=5
    const slideNode = imported.document.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    const findNamedNode = (
      root: CommonNode,
      targetName: string,
    ): CommonNode | undefined => {
      for (const child of root.children) {
        if (!(child instanceof CommonNode)) continue;
        if (child.props["name"] === targetName) return child;
        const nested = findNamedNode(child, targetName);
        if (nested !== undefined) return nested;
      }
      return undefined;
    };
    const shape10 = findNamedNode(slideNode, "NoFillShape");
    assert.ok(shape10 !== undefined);
    assert.equal(shape10.props["emphasis"], undefined);
    assert.equal(
      (shape10.props["enter"] as { readonly preset?: string } | undefined)
        ?.preset,
      "fade",
    );

    const shape11 = findNamedNode(slideNode, "GradShape");
    assert.equal(shape11?.props["emphasis"], undefined);

    const shape12 = findNamedNode(slideNode, "SolidBox");
    assert.equal(shape12?.props["enter"], undefined);
    assert.equal(
      (shape12?.props["emphasis"] as { readonly preset?: string } | undefined)
        ?.preset,
      "spin",
    );

    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" &&
          d.location.part === "ppt/slideMasters/slideMaster1.xml",
      ),
    );
    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" &&
          d.cascaded !== undefined &&
          d.cascaded.length > 0,
      ),
    );

    assertCodegenAndEmitPass(imported.document, { attachJostFont: true });
  });

  it("materializes 5-layer inheritance and unbinds duplicate, cross-role, and orphan placeholders", () => {
    const png1x1Base64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const pngBytes = Uint8Array.from(Buffer.from(png1x1Base64, "base64"));

    const masterXml = `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="M1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="MasterTitle"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="2800" b="1"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:latin typeface="Jost*"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="MasterBody"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="4000000" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="135000"/></a:lnSpc><a:buChar char="•"/><a:defRPr sz="1600"><a:solidFill><a:srgbClr val="334155"/></a:solidFill><a:latin typeface="Jost*"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="115000"/></a:lnSpc><a:defRPr sz="3200"><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:defRPr></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="125000"/></a:lnSpc><a:defRPr sz="1800"><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:defRPr></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`;

    const layoutXml = `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="obj">
  <p:cSld name="L1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="LayoutTitle"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="914400"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="LayoutPicSlot"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" idx="10"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="4800000" y="1371600"/><a:ext cx="3600450" cy="2286000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Pic</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

    const slideXml = `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="PrimaryTitle"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="DuplicateTitle"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="4" name="PicInTitleRole"/><p:cNvPicPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr/>
      </p:pic>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="TextInPicSlot"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="10"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Caption in Pic</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="OrphanBody"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="99"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Orphan Bullet</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="OrphanUnpopulatedShape"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="obj" idx="88"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="3800000"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

    const pptx = baseDeckFiles({
      slideXml,
      slideRelsXml: `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
      extraFiles: {
        "ppt/slideMasters/slideMaster1.xml": masterXml,
        "ppt/slideLayouts/slideLayout1.xml": layoutXml,
        "ppt/media/image1.png": pngBytes,
      },
    });

    const imported = importAll(pptx);
    const slideNode = imported.document.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    const byName = (targetName: string): CommonNode =>
      slideNode.children.find(
        (c): c is CommonNode =>
          c instanceof CommonNode && c.props["name"] === targetName,
      )!;

    // 1. PrimaryTitle retains binding
    const primaryTitle = byName("PrimaryTitle");
    assert.equal(primaryTitle.props["placeholder"], "title");
    assert.ok(primaryTitle.props["nativePlaceholder"] !== undefined);

    // 2. DuplicateTitle unbinds and materializes 5-layer geometry, background, font, eastAsianFontFamily, and % lineHeight
    const dupTitle = byName("DuplicateTitle");
    assert.equal(dupTitle.kind, "text");
    assert.equal(dupTitle.props["placeholder"], undefined);
    assert.equal(dupTitle.props["nativePlaceholder"], undefined);
    const dupStyle = dupTitle.props["style"] as Record<string, unknown>;
    assert.equal(dupStyle["position"], "absolute");
    assert.equal(dupStyle["left"], 48);
    assert.equal(dupStyle["top"], 28.8);
    assert.equal(dupStyle["width"], 864);
    assert.equal(dupStyle["height"], 96);
    assert.equal(dupStyle["background"], "#F1F5F9");
    assert.equal(dupStyle["fontSize"], "28pt");
    assert.equal(dupStyle["eastAsianFontFamily"], "minor");
    assert.equal(dupStyle["lineHeight"], "115%");

    // 3. PicInTitleRole unbinds and inherits title box dimensions
    const picInTitle = byName("PicInTitleRole");
    assert.equal(picInTitle.kind, "image");
    assert.equal(picInTitle.props["placeholder"], undefined);
    assert.equal(picInTitle.props["nativePlaceholder"], undefined);
    const picStyle = picInTitle.props["style"] as Record<string, unknown>;
    assert.equal(picStyle["position"], "absolute");
    assert.equal(picStyle["width"], 864);
    assert.equal(picStyle["height"], 96);

    // 4. TextInPicSlot unbinds and inherits pic slot geometry
    const textInPic = byName("TextInPicSlot");
    assert.equal(textInPic.kind, "text");
    assert.equal(textInPic.props["placeholder"], undefined);
    assert.equal(textInPic.props["nativePlaceholder"], undefined);
    assert.equal(
      (textInPic.props["style"] as Record<string, unknown>)["width"],
      378,
    );

    // 5. OrphanBody unbinds, inherits Master body geometry + bullet + eastAsianFontFamily + % lineHeight
    const orphanBody = byName("OrphanBody");
    assert.equal(orphanBody.kind, "text");
    assert.equal(orphanBody.props["placeholder"], undefined);
    assert.equal(orphanBody.props["nativePlaceholder"], undefined);
    const orphanPara = orphanBody.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "paragraph",
    );
    assert.ok(orphanPara !== undefined);
    const orphanParaStyle = orphanPara.props["style"] as Record<
      string,
      unknown
    >;
    assert.equal(orphanParaStyle["eastAsianFontFamily"], "major");
    assert.equal(orphanParaStyle["lineHeight"], "135%");
    assert.deepEqual(
      { ...(orphanParaStyle["list"] as Record<string, unknown>) },
      { type: "bullet", character: "•" },
    );

    // 6. OrphanUnpopulatedShape converts into free Shape without placeholder props
    const orphanShape = byName("OrphanUnpopulatedShape");
    assert.equal(orphanShape.kind, "shape");
    assert.equal(orphanShape.props["placeholder"], undefined);
    assert.equal(orphanShape.props["nativePlaceholder"], undefined);

    const fallbacks = imported.report.diagnostics.filter(
      (d) =>
        d.code === "compatibility-fallback" &&
        d.loss.lost === "master-layout-inheritance",
    );
    assert.equal(fallbacks.length, 5);

    assertCodegenAndEmitPass(imported.document, { attachJostFont: true });
  });

  it("recovers SmartArt with pre-computed dsp:drawing into editable Group and omits SmartArt missing drawing", () => {
    const DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    const png1x1Base64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const pngBytes = Uint8Array.from(Buffer.from(png1x1Base64, "base64"));

    const slideXml = `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="4" name="SmartArtWithDrawing"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="5486400" cy="2743200"/></p:xfrm>
        <a:graphic><a:graphicData uri="${DIAGRAM_NS}"><dgm:relIds xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${R_NS}" r:dm="rId2"/></a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="5" name="SmartArtMissingDrawing"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="457200" y="3400000"/><a:ext cx="2743200" cy="914400"/></p:xfrm>
        <a:graphic><a:graphicData uri="${DIAGRAM_NS}"><dgm:relIds xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${R_NS}" r:dm="rId4"/></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

    const data1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="${DIAGRAM_NS}" xmlns:a="${A_NS}">
  <dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/>
  <dgm:extLst><a:ext uri="${DSP_NS}"><dsp:dataModelExt xmlns:dsp="${DSP_NS}" relId="rId3"/></a:ext></dgm:extLst>
</dgm:dataModel>`;

    const data2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="${DIAGRAM_NS}" xmlns:a="${A_NS}"><dgm:ptLst/><dgm:cxnLst/></dgm:dataModel>`;

    const drawing1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <dsp:spTree>
    <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
    <dsp:grpSpPr/>
    <dsp:sp modelId="{A1}">
      <dsp:nvSpPr><dsp:cNvPr id="0" name="StepBox"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></dsp:spPr>
      <dsp:style><a:lnRef idx="0"/><a:fillRef idx="1"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></dsp:style>
      <dsp:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Node A</a:t></a:r></a:p></dsp:txBody>
      <dsp:txXfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></dsp:txXfrm>
    </dsp:sp>
    <dsp:cxnSp modelId="{C1}">
      <dsp:nvCxnSpPr><dsp:cNvPr id="0" name="StepArrow"/><dsp:cNvCxnSpPr/></dsp:nvCxnSpPr>
      <dsp:spPr><a:xfrm><a:off x="1828800" y="457200"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill></a:ln></dsp:spPr>
    </dsp:cxnSp>
    <dsp:pic modelId="{P1}">
      <dsp:nvPicPr><dsp:cNvPr id="0" name="NodeIcon"/><dsp:cNvPicPr/></dsp:nvPicPr>
      <dsp:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></dsp:blipFill>
      <dsp:spPr><a:xfrm><a:off x="2743200" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr>
    </dsp:pic>
  </dsp:spTree>
</dsp:drawing>`;

    const pptx = baseDeckFiles({
      slideXml,
      slideRelsXml: `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_DIAGRAM_DATA}" Target="../diagrams/data1.xml"/><Relationship Id="rId3" Type="${REL_DIAGRAM_DRAWING}" Target="../diagrams/drawing1.xml"/><Relationship Id="rId4" Type="${REL_DIAGRAM_DATA}" Target="../diagrams/data2.xml"/></Relationships>`,
      extraFiles: {
        "ppt/diagrams/data1.xml": data1Xml,
        "ppt/diagrams/data2.xml": data2Xml,
        "ppt/diagrams/drawing1.xml": drawing1Xml,
        "ppt/diagrams/_rels/drawing1.xml.rels": `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
        "ppt/media/image1.png": pngBytes,
      },
    });

    const imported = importAll(pptx);
    const slideNode = imported.document.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    const groups = slideNode.children.filter(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "group",
    );
    assert.equal(groups.length, 1);
    const recoveredGroup = groups[0]!;
    assert.equal(recoveredGroup.props["name"], "SmartArtWithDrawing");
    const memberKinds = recoveredGroup.children
      .filter((c): c is CommonNode => c instanceof CommonNode)
      .map((c) => c.kind);
    assert.deepEqual(memberKinds, ["shape", "line", "image"]);

    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "compatibility-fallback" &&
          d.selectedFallback === "group" &&
          d.loss.lost === "native-mechanism" &&
          d.loss.retainedAs === "group",
      ),
    );
    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" &&
          d.mechanismHint === DIAGRAM_NS &&
          d.location.objectName === "SmartArtMissingDrawing",
      ),
    );

    assertCodegenAndEmitPass(imported.document, { attachJostFont: true });
  });

  it("handles flat-text paragraph animations, omits zero-width collapsed groups, and remaps layout connector ownerMotion spids", () => {
    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="FlatText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="2743200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Single flat paragraph</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="3" name="CollapsedGroup"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="457200" y="1000000"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
        <p:pic>
          <p:nvPicPr><p:cNvPr id="4" name="MissingImage"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
          <p:blipFill><a:blip r:embed="rIdMissing"/></p:blipFill>
          <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:pic>
        <p:cxnSp>
          <p:nvCxnSpPr><p:cNvPr id="5" name="VerticalStem"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
          <p:spPr><a:xfrm><a:off x="457200" y="0"/><a:ext cx="0" cy="914400"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:ln></p:spPr>
        </p:cxnSp>
      </p:grpSp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="300"/><p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sld>`;

    const layoutWithAnimatedConnector = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank">
  <p:cSld name="L1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="9" name="LayoutLine"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:ln></p:spPr>
      </p:cxnSp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="9"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="300"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sldLayout>`;

    const imported = importAll(
      baseDeckFiles({
        slideXml,
        extraFiles: {
          "ppt/slideLayouts/slideLayout1.xml": layoutWithAnimatedConnector,
        },
      }),
    );
    const deckDoc = imported.document as DeckContainer;
    const slideNode = deckDoc.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    assert.equal(
      slideNode.children.some(
        (c) => c instanceof CommonNode && c.kind === "group",
      ),
      false,
    );
    const layoutMotion = deckDoc.masters[0]?.layouts[0]?.props["motion"] as {
      readonly steps: readonly {
        readonly effects: readonly {
          readonly target: { readonly spid?: number };
        }[];
      }[];
    };
    assert.equal(layoutMotion.steps[0]?.effects[0]?.target.spid, 2);

    assertCodegenAndEmitPass(imported.document, { attachJostFont: true });
  });

  it("strips out-of-bounds scalar run/paragraph/master/motion XML values with invalid-xml-value diagnostics", async () => {
    const masterXml = `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="M1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="MasterTitle"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="MasterFooter"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="4500000"/><a:ext cx="3600000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="50" b="1"/><a:t>Footer</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/><a:defRPr sz="50" spc="500000"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`;

    const layoutXml = `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="obj">
  <p:cSld name="L1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="LayoutBody"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="6400000" cy="2286000"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPts val="200000"/></a:lnSpc><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/><a:defRPr sz="50" spc="-500000" b="1"/></a:lvl1pPr></a:lstStyle>
          <a:p><a:r><a:t>Placeholder</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

    const slideXml = `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="10" name="OutOfBoundsText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:lnSpc><a:spcPts val="200000"/></a:lnSpc><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/></a:pPr>
            <a:r><a:rPr sz="50" spc="500000" baseline="3000000000" b="1"/><a:t>BadRun1 </a:t></a:r>
            <a:r><a:rPr sz="500000" spc="-500000" i="1"/><a:t>BadRun2</a:t></a:r>
          </a:p>
          <a:p>
            <a:pPr><a:lnSpc><a:spcPct val="20000000"/></a:lnSpc><a:spcAft><a:spcPct val="500000"/></a:spcAft></a:pPr>
            <a:r><a:rPr sz="100000"/><a:t>OverflowPct</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="OutOfBoundsShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="2500000"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr lvl="9"><a:lnSpc><a:spcPts val="200000"/></a:lnSpc><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/></a:pPr>
            <a:r><a:rPr sz="50" spc="500000" baseline="3000000000" b="1"/><a:t>ShapeBad</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="12" name="AnimShapeBadDur"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="3600000" y="2500000"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="10B981"/></a:solidFill></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="13" name="AnimShapeBadRot"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="5800000" y="2500000"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F59E0B"/></a:solidFill></p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="14" name="AnimShapeNegDelay"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="5800000" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="8B5CF6"/></a:solidFill></p:spPr>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="12"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="5000000000"/><p:tgtEl><p:spTgt spid="12"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="8" presetID="8" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animRot by="3000000000"><p:cBhvr><p:cTn id="9" dur="500"/><p:tgtEl><p:spTgt spid="13"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="10" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond evt="end" delay="50"><p:tn val="8"/></p:cond></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="11" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="14"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="12" dur="300"/><p:tgtEl><p:spTgt spid="14"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="13" presetID="0" presetClass="emph" presetSubtype="0" autoRev="1" fill="hold"><p:stCondLst><p:cond delay="-500"/></p:stCondLst><p:childTnLst>
              <p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="14" dur="500"/><p:tgtEl><p:spTgt spid="14"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="2563EB"/></p:from><p:to><a:srgbClr val="FFFFFF"/></p:to></p:animClr>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sld>`;

    const pptx = baseDeckFiles({
      slideXml,
      extraFiles: {
        "ppt/slideMasters/slideMaster1.xml": masterXml,
        "ppt/slideLayouts/slideLayout1.xml": layoutXml,
      },
    });
    const imported = importAll(pptx);
    const deck = imported.document as DeckContainer;
    const slideNode = deck.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    );
    assert.ok(
      slideNode !== undefined,
      `Expected slideNode; children=${JSON.stringify(deck.children.map((c) => (c instanceof CommonNode ? c.kind : c.type)))} diags=${JSON.stringify(imported.report.diagnostics)}`,
    );
    const byName = (
      targetName: string,
      root: CommonNode = slideNode,
    ): CommonNode | undefined => {
      for (const child of root.children) {
        if (!(child instanceof CommonNode)) continue;
        if (child.props["name"] === targetName) return child;
        const nested = byName(targetName, child);
        if (nested !== undefined) return nested;
      }
      return undefined;
    };

    // Preserved valid siblings and bounded lists on Text & Shape
    const textNode = byName("OutOfBoundsText")!;
    const firstPara = textNode.children[0] as CommonNode;
    const firstParaStyle = firstPara.props["style"] as Record<string, unknown>;
    assert.equal(firstParaStyle["lineHeight"], undefined);
    assert.equal(firstParaStyle["spacingAfter"], undefined);
    assert.deepEqual(
      { ...(firstParaStyle["list"] as Record<string, unknown>) },
      { type: "number", style: "arabicPeriod" },
    );
    const firstSpan = firstPara.children[0] as CommonNode;
    assert.equal(
      (firstSpan.props["style"] as Record<string, unknown>)["fontWeight"],
      "bold",
    );
    assert.equal(
      (firstSpan.props["style"] as Record<string, unknown>)["fontSize"],
      undefined,
    );

    const shapeNode = byName("OutOfBoundsShape")!;
    const shapePara = shapeNode.children[0] as CommonNode;
    const shapeParaStyle = shapePara.props["style"] as Record<string, unknown>;
    assert.equal(shapeParaStyle["lineHeight"], undefined);
    assert.equal(shapeParaStyle["spacingAfter"], undefined);
    assert.deepEqual(
      { ...(shapeParaStyle["list"] as Record<string, unknown>) },
      { type: "number", style: "arabicPeriod" },
    );

    // Master & layout placeholders sanitized cleanly
    const layoutBody = (
      deck.masters[0]?.layouts[0]?.props["placeholders"] as
        | Record<string, Record<string, unknown>>
        | undefined
    )?.["body"];
    assert.equal(
      (
        layoutBody?.["listStyles"] as
          | Record<number, Record<string, unknown>>
          | undefined
      )?.[1]?.["fontWeight"],
      "bold",
    );
    assert.equal(
      (
        layoutBody?.["listStyles"] as
          | Record<number, Record<string, unknown>>
          | undefined
      )?.[1]?.["fontSize"],
      undefined,
    );

    // Invalid motion dur & spin by omitted cleanly with cascaded cleanup; negative delay recovered
    assert.equal(byName("AnimShapeBadDur")?.props["enter"], undefined);
    assert.equal(byName("AnimShapeBadRot")?.props["emphasis"], undefined);
    assert.equal(byName("AnimShapeNegDelay")?.props["enter"], undefined);
    assert.equal(
      (
        byName("AnimShapeNegDelay")?.props["emphasis"] as
          | { readonly delay?: number }
          | undefined
      )?.delay,
      0,
    );

    assert.ok(
      imported.report.diagnostics.filter((d) => d.code === "invalid-xml-value")
        .length >= 8,
    );

    assertCodegenAndEmitPass(deck, { attachJostFont: true });
  });
});
