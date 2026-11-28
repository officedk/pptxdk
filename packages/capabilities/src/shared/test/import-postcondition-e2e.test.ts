import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { font } from "@pptxdk/capabilities/font";
import { createCodegen } from "@pptxdk/codegen";
import { isTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode, type DeckContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import {
  A_NS,
  C_NS,
  DIAGRAM_NS,
  DSP_NS,
  OPC_REL_NS,
  P_NS,
  REL_CHART,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  REL_IMAGE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

import { verifyBatchGeneratedSourcesAndRelocatedRebuilds } from "../../../test/codegen-batch-acceptance.ts";

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

const png1x1Bytes = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

function buildCombinedPostconditionDeck(): Uint8Array {
  const themeXml = `<a:theme xmlns:a="${A_NS}" name="ShortFmtTheme">
  <a:themeElements>
    <a:clrScheme name="C"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>
    <a:fontScheme name="F"><a:majorFont><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:majorFont><a:minorFont><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Short"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;

  const masterXml = `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="M1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="MasterTitle"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="2400"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="MasterBody"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="4000000" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="130000"/></a:lnSpc><a:buChar char="•"/><a:defRPr sz="1600"><a:solidFill><a:srgbClr val="334155"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="115000"/></a:lnSpc></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle/>
  </p:txStyles>
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
              <p:set><p:cBhvr><p:cTn id="9" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="999"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="10" dur="300"/><p:tgtEl><p:spTgt spid="999"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
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
      <p:cxnSp>
        <p:nvCxnSpPr><p:cNvPr id="9" name="LayoutDivider"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1200000"/><a:ext cx="8229600" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln></p:spPr>
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

  const chartXml = `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/><c:tx><c:v>S1</c:v></c:tx>
          <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="9"/><c:spPr><a:solidFill><a:srgbClr val="EF4444"/></a:solidFill></c:spPr></c:dPt>
          <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:lineChart><c:grouping val="standard"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

  const slide1Xml = `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="s1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TitlePrimary"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Primary Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="TitleDuplicate"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Duplicate Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="OrphanBullet"/><p:cNvSpPr><p:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="99"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buChar char="ab"/></a:pPr><a:r><a:t>Bullet item</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:pic>
        <p:nvPicPr><p:cNvPr id="5" name="ImageInTitle"/><p:cNvPicPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvPicPr>
        <p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr/>
      </p:pic>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="6" name="SanitizedChart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="4800000" y="1371600"/><a:ext cx="3600450" cy="2286000"/></p:xfrm>
        <a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="rId2"/></a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="NoFillAnimated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="3900000"/><a:ext cx="1828800" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="8" name="SmartArtGroup"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="2600000" y="3800000"/><a:ext cx="3657600" cy="914400"/></p:xfrm>
        <a:graphic><a:graphicData uri="${DIAGRAM_NS}"><dgm:relIds xmlns:dgm="${DIAGRAM_NS}" xmlns:r="${R_NS}" r:dm="rId4"/></a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="9" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animClr clrSpc="rgb"><p:cBhvr><p:cTn id="6" dur="400" fill="hold"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl><p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr><p:from><a:srgbClr val="111111"/></p:from><p:to><a:srgbClr val="FF0000"/></p:to></p:animClr>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="7" presetID="10" presetClass="entr" presetSubtype="0" repeatCount="2000" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="8" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="7"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="9" dur="300"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sld>`;

  const data1Xml = `<dgm:dataModel xmlns:dgm="${DIAGRAM_NS}" xmlns:a="${A_NS}"><dgm:ptLst/><dgm:cxnLst/><dgm:extLst><a:ext uri="${DSP_NS}"><dsp:dataModelExt xmlns:dsp="${DSP_NS}" relId="rId5"/></a:ext></dgm:extLst></dgm:dataModel>`;

  const drawing1Xml = `<dsp:drawing xmlns:dsp="${DSP_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">
  <dsp:spTree>
    <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
    <dsp:grpSpPr/>
    <dsp:sp modelId="{N1}">
      <dsp:nvSpPr><dsp:cNvPr id="0" name="DiagramCard"/><dsp:cNvSpPr/></dsp:nvSpPr>
      <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></dsp:spPr>
      <dsp:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Box</a:t></a:r></a:p></dsp:txBody>
    </dsp:sp>
  </dsp:spTree>
</dsp:drawing>`;

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
    "ppt/theme/theme1.xml": strToU8(themeXml),
    "ppt/slideMasters/slideMaster1.xml": strToU8(masterXml),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
      `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
    ),
    "ppt/slideLayouts/slideLayout1.xml": strToU8(layoutXml),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
      `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    ),
    "ppt/slides/slide1.xml": strToU8(slide1Xml),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(
      `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_CHART}" Target="../charts/chart1.xml"/><Relationship Id="rId3" Type="${REL_IMAGE}" Target="../media/image1.png"/><Relationship Id="rId4" Type="${REL_DIAGRAM_DATA}" Target="../diagrams/data1.xml"/><Relationship Id="rId5" Type="${REL_DIAGRAM_DRAWING}" Target="../diagrams/drawing1.xml"/></Relationships>`,
    ),
    "ppt/charts/chart1.xml": strToU8(chartXml),
    "ppt/media/image1.png": png1x1Bytes,
    "ppt/diagrams/data1.xml": strToU8(data1Xml),
    "ppt/diagrams/drawing1.xml": strToU8(drawing1Xml),
  });
}

describe("importPptx postcondition L2 end-to-end invariant", () => {
  it("passes dual-target codegen (with strict tsc), resolve, emit, validatePptx, and paint across all sanitized invariants", async () => {
    const imported = importAll(buildCombinedPostconditionDeck());
    const deck = imported.document as DeckContainer;
    const slideNode = deck.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    )!;
    assert.ok(slideNode !== undefined);

    // Dual-target codegen + strict single-batch tsc --noEmit verification
    const vanilla = codegenAll(deck, { target: "vanilla" });
    const react = codegenAll(deck, { target: "react" });
    const batch = await verifyBatchGeneratedSourcesAndRelocatedRebuilds(
      [
        {
          scenarioId: "import-postcondition-e2e",
          originLabel: "roundtrip-imported",
          result: vanilla,
        },
        {
          scenarioId: "import-postcondition-e2e",
          originLabel: "roundtrip-imported",
          result: react,
        },
      ],
      [],
    );
    assert.equal(batch.typecheckedCaseCount, 2);

    // Attach required Jost* outline fonts for measurement + resolve -> emit -> validatePptx -> paint
    for (const master of deck.masters) {
      if (isTheme(master.theme)) {
        master.theme.set({
          fonts: {
            major: jostFont,
            minor: jostFont,
            majorEastAsian: jostFont,
            minorEastAsian: jostFont,
          },
        });
      }
    }
    const resolution = resolveAll(deck);
    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);

    const preview = paint(resolution);
    assert.equal(preview.pages.length, 1);
    assert.deepEqual(preview.pages[0]!.errors, []);
  });

  it("sanitizes out-of-bounds scalar attributes end-to-end across dual-target codegen, tsc, emit, and validatePptx", async () => {
    const masterXml = `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="M1">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="MasterTitle"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="8229600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="50"/><a:t>Title</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="50" spc="500000"><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>
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
        <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/><a:defRPr sz="50"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:r><a:t>Body</a:t></a:r></a:p></p:txBody>
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
        <p:nvSpPr><p:cNvPr id="10" name="ScalarText"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:lnSpc><a:spcPts val="200000"/></a:lnSpc><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/></a:pPr><a:r><a:rPr sz="50" spc="500000" baseline="3000000000"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>Cleaned text</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="ScalarShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="2500000"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:spcAft><a:spcPts val="200000"/></a:spcAft><a:buAutoNum type="arabicPeriod" startAt="40000"/></a:pPr><a:r><a:rPr sz="50" spc="-500000"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:rPr><a:t>Cleaned shape</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
  <p:timing>
    <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"><p:childTnLst>
      <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
        <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
            <p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="5000000000"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="11"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="5000000000"/><p:tgtEl><p:spTgt spid="11"/></p:tgtEl></p:cBhvr></p:animEffect>
            </p:childTnLst></p:cTn></p:par>
            <p:par><p:cTn id="8" presetID="8" presetClass="emph" presetSubtype="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:animRot by="3000000000"><p:cBhvr><p:cTn id="9" dur="500"/><p:tgtEl><p:spTgt spid="11"/></p:tgtEl><p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:seq>
    </p:childTnLst></p:cTn></p:par></p:tnLst>
  </p:timing>
</p:sld>`;

    const pptx = zipSync({
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
        `<a:theme xmlns:a="${A_NS}" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="10B981"/></a:accent2><a:accent3><a:srgbClr val="F59E0B"/></a:accent3><a:accent4><a:srgbClr val="EF4444"/></a:accent4><a:accent5><a:srgbClr val="8B5CF6"/></a:accent5><a:accent6><a:srgbClr val="06B6D4"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:majorFont><a:minorFont><a:latin typeface="Jost*"/><a:ea typeface="Jost*"/></a:minorFont></a:fontScheme><a:fmtScheme name="O"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(masterXml),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(layoutXml),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(slideXml),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const imported = importAll(pptx);
    assert.ok(
      imported.report.diagnostics.some((d) => d.code === "invalid-xml-value"),
    );
    const deck = imported.document as DeckContainer;
    const vanilla = codegenAll(deck, { target: "vanilla" });
    const react = codegenAll(deck, { target: "react" });
    const batch = await verifyBatchGeneratedSourcesAndRelocatedRebuilds(
      [
        {
          scenarioId: "import-scalar-postcondition-e2e",
          originLabel: "roundtrip-imported",
          result: vanilla,
        },
        {
          scenarioId: "import-scalar-postcondition-e2e",
          originLabel: "roundtrip-imported",
          result: react,
        },
      ],
      [],
    );
    assert.equal(batch.typecheckedCaseCount, 2);

    for (const master of deck.masters) {
      if (isTheme(master.theme)) {
        master.theme.set({
          fonts: {
            major: jostFont,
            minor: jostFont,
            majorEastAsian: jostFont,
            minorEastAsian: jostFont,
          },
        });
      }
    }
    const resolution = resolveAll(deck);
    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);

    const preview = paint(resolution);
    assert.equal(preview.pages.length, 1);
    assert.deepEqual(preview.pages[0]!.errors, []);
  });
});
