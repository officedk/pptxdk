import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { A_NS, createImporter, P_NS } from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

function buildPptxWithSlideBody(slideSpTreeInner: string): Uint8Array {
  const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${P_NS}">
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
  const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`;
  const masterXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${P_NS}">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;
  const masterRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
  const layoutXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${P_NS}" type="blank">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
</p:sldLayout>`;
  const layoutRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${A_NS}" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="EA580C"/></a:accent2>
      <a:accent3><a:srgbClr val="16A34A"/></a:accent3><a:accent4><a:srgbClr val="9333EA"/></a:accent4>
      <a:accent5><a:srgbClr val="0891B2"/></a:accent5><a:accent6><a:srgbClr val="DC2626"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>`;
  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${P_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${slideSpTreeInner}
    </p:spTree>
  </p:cSld>
</p:sld>`;
  const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
  return zipSync({
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "ppt/presentation.xml": strToU8(presentationXml),
    "ppt/_rels/presentation.xml.rels": strToU8(presentationRels),
    "ppt/slideMasters/slideMaster1.xml": strToU8(masterXml),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(masterRels),
    "ppt/slideLayouts/slideLayout1.xml": strToU8(layoutXml),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(layoutRels),
    "ppt/theme/theme1.xml": strToU8(themeXml),
    "ppt/slides/slide1.xml": strToU8(slideXml),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(slideRels),
  });
}

function collectNodes(root: CommonNode): CommonNode[] {
  const out: CommonNode[] = [];
  const visit = (node: CommonNode): void => {
    out.push(node);
    for (const child of node.children) {
      if (child instanceof CommonNode) visit(child);
    }
  };
  visit(root);
  return out;
}

describe("Line import and <p:sp> normalization", () => {
  it("normalizes empty no-fill line/connector <p:sp> (including 0-extent) into <Line> while preserving filled or text <p:sp> as <Shape>", () => {
    const pptx = buildPptxWithSlideBody(`
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Horiz Sp Line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="1905000" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="19050"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Inv Sp Line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="952500" y="1905000"/><a:ext cx="952500" cy="952500"/></a:xfrm>
          <a:prstGeom prst="lineInv"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></p:style>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Text Line Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="3048000" y="952500"/><a:ext cx="952500" cy="476250"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="DBEAFE"/></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Label</a:t></a:r></a:p></p:txBody>
      </p:sp>
    `);

    const imported = importDocument(pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const nodes = collectNodes(imported.document);
    const lines = nodes.filter((n) => n.type === "line");
    const shapes = nodes.filter((n) => n.type === "shape");
    assert.equal(lines.length, 2);
    assert.equal(shapes.length, 1);

    const horiz = lines.find((n) => n.props["name"] === "Horiz Sp Line");
    assert.ok(horiz !== undefined);
    assert.equal((horiz.props["style"] as { height: number }).height, 0);
    assert.deepEqual(
      { ...(horiz.props["stroke"] as object) },
      {
        color: "#1E293B",
        width: 2,
        endArrow: "triangle",
      },
    );

    const inv = lines.find((n) => n.props["name"] === "Inv Sp Line");
    assert.ok(inv !== undefined);
    assert.equal((inv.props["style"] as { flipV?: boolean }).flipV, true);
    assert.deepEqual(
      { ...(inv.props["stroke"] as object) },
      {
        color: "accent1",
        width: 1,
        style: "solid",
      },
    );

    const resolution = resolveDocument(imported.document);
    assert.equal(
      resolution.slides[0]?.objects.filter((o) => o.kind === "line").length,
      2,
    );
  });

  it("imports <p:cxnSp> with missing anchor targets using box-layout fallback diagnostics", () => {
    const pptx = buildPptxWithSlideBody(`
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Live Target"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="476250"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="DBEAFE"/></a:solidFill>
        </p:spPr>
      </p:sp>
      <p:cxnSp>
        <p:nvCxnSpPr>
          <p:cNvPr id="3" name="Broken Connector"/>
          <p:cNvCxnSpPr><a:stCxn id="2" idx="3"/><a:endCxn id="999" idx="1"/></p:cNvCxnSpPr>
          <p:nvPr/>
        </p:nvCxnSpPr>
        <p:spPr>
          <a:xfrm><a:off x="1905000" y="1190625"/><a:ext cx="952500" cy="0"/></a:xfrm>
          <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="DC2626"/></a:solidFill></a:ln>
        </p:spPr>
      </p:cxnSp>
    `);

    const imported = importDocument(pptx);
    const nodes = collectNodes(imported.document);
    const broken = nodes.find((n) => n.props["name"] === "Broken Connector");
    assert.ok(broken !== undefined);
    assert.equal(broken.type, "line");
    assert.equal(broken.props["from"], undefined);
    assert.equal(broken.props["to"], undefined);
    assert.ok(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "compatibility-fallback" &&
          d.preferredMechanism === "connector-anchor",
      ),
    );
    const resolution = resolveDocument(imported.document);
    assert.equal(
      resolution.slides[0]?.objects.filter((o) => o.kind === "line").length,
      1,
    );
  });

  it("imports <p:cxnSp> and normalized <p:sp> inside <p:grpSp> while keeping group union bounds intact", () => {
    const pptx = buildPptxWithSlideBody(`
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="2" name="Diag Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="952500" y="952500"/><a:ext cx="2857500" cy="952500"/>
            <a:chOff x="952500" y="952500"/><a:chExt cx="2857500" cy="952500"/>
          </a:xfrm>
        </p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Left Node"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill>
          </p:spPr>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="4" name="Right Node"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="2857500" y="952500"/><a:ext cx="952500" cy="952500"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill>
          </p:spPr>
        </p:sp>
        <p:cxnSp>
          <p:nvCxnSpPr>
            <p:cNvPr id="5" name="Group Link"/>
            <p:cNvCxnSpPr><a:stCxn id="3" idx="3"/><a:endCxn id="4" idx="1"/></p:cNvCxnSpPr>
            <p:nvPr/>
          </p:nvCxnSpPr>
          <p:spPr>
            <a:xfrm><a:off x="1905000" y="1428750"/><a:ext cx="952500" cy="0"/></a:xfrm>
            <a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
          </p:spPr>
        </p:cxnSp>
      </p:grpSp>
    `);

    const imported = importDocument(pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const nodes = collectNodes(imported.document);
    const grp = nodes.find((n) => n.type === "group");
    const grpConn = nodes.find((n) => n.props["name"] === "Group Link");
    assert.ok(grp !== undefined);
    assert.ok(grpConn !== undefined);
    assert.equal(grpConn.type, "line");
    assert.deepEqual(
      { ...(grpConn.props["from"] as object) },
      {
        id: "Left Node",
        port: "right",
      },
    );
    assert.deepEqual(
      { ...(grpConn.props["to"] as object) },
      {
        id: "Right Node",
        port: "left",
      },
    );

    const resolution = resolveDocument(imported.document);
    assert.equal(resolution.slides[0]?.objects[0]?.kind, "group");
    assert.equal(
      resolution.slides[0]?.objects[0]?.members?.some((m) => m.kind === "line"),
      true,
    );
  });
});
