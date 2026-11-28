import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  A_NS,
  ImportError,
  OPC_REL_NS,
  P_NS,
  R_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  createImporter,
  fatalImportIssueSchema,
  importDiagnosticSchema,
  importReportSchema,
} from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

function minimalTransitionalPptx(options?: {
  readonly slideContent?: string;
  readonly slideRelsXml?: string;
  readonly presRelsXml?: string;
  readonly presentationXml?: string;
  readonly rootRelsXml?: string;
  readonly omitRootRels?: boolean;
  readonly omitPresentation?: boolean;
  readonly omitSlide1?: boolean;
}): Uint8Array {
  const rootRels =
    options?.rootRelsXml ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/>
</Relationships>`;

  const presentation =
    options?.presentationXml ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;

  const presRels =
    options?.presRelsXml ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
</Relationships>`;

  const slide1 =
    options?.slideContent ??
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="intro">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

  const entries: Record<string, Uint8Array> = {};
  if (!options?.omitRootRels) {
    entries["_rels/.rels"] = strToU8(rootRels);
  }
  if (!options?.omitPresentation) {
    entries["ppt/presentation.xml"] = strToU8(presentation);
  }
  entries["ppt/_rels/presentation.xml.rels"] = strToU8(presRels);
  if (!options?.omitSlide1) {
    entries["ppt/slides/slide1.xml"] = strToU8(slide1);
  }
  if (options?.slideRelsXml !== undefined) {
    entries["ppt/slides/_rels/slide1.xml.rels"] = strToU8(options.slideRelsXml);
  }
  return zipSync(entries);
}

describe("@pptxdk/import package entry and diagnostics", () => {
  it("rejects the real Strict OOXML fixture with structured unsupported-format-baseline diagnostics", () => {
    const strictBytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/strict-anchorModified.pptx", import.meta.url),
      ),
    );
    const importer = createImporter([]);

    let caught: unknown;
    try {
      importer(strictBytes);
    } catch (error) {
      caught = error;
    }

    assert.ok(caught instanceof ImportError);
    const parsed = fatalImportIssueSchema.parse(caught.issue);
    assert.equal(parsed.code, "unsupported-format-baseline");
    assert.equal(parsed.category, "unsupported");
    assert.equal(parsed.disposition, "aborted");
    if (parsed.code === "unsupported-format-baseline") {
      assert.equal(parsed.baseline, "ISO/IEC 29500 Strict");
      assert.ok(
        parsed.detectedNamespace.startsWith("http://purl.oclc.org/ooxml/"),
      );
      assert.equal(parsed.location.part, "_rels/.rels");
      assert.ok(parsed.location.xpath !== undefined);
    }
  });

  it("also detects Strict OOXML when presentation.xml uses Strict namespace despite Transitional rels", () => {
    const importer = createImporter([]);
    const hybridStrict = minimalTransitionalPptx({
      presentationXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://purl.oclc.org/ooxml/presentationml/main">
  <p:sldIdLst/>
</p:presentation>`,
    });

    assert.throws(
      () => importer(hybridStrict),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "unsupported-format-baseline" &&
        error.issue.category === "unsupported" &&
        error.issue.location.part === "ppt/presentation.xml" &&
        error.issue.detectedNamespace ===
          "http://purl.oclc.org/ooxml/presentationml/main",
    );
  });

  it("distinguishes corrupt archive, missing entry, malformed XML, broken entry rels, and invalid presentation root", () => {
    const importer = createImporter([]);

    assert.throws(
      () => importer(new Uint8Array([1, 2, 3, 4])),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "corrupt-archive" &&
        error.issue.category === "corrupt-input",
    );

    assert.throws(
      () => importer(minimalTransitionalPptx({ omitRootRels: true })),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "missing-entry" &&
        error.issue.location.part === "_rels/.rels",
    );

    assert.throws(
      () =>
        importer(
          minimalTransitionalPptx({
            rootRelsXml: "<Relationships><broken",
          }),
        ),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "malformed-xml" &&
        error.issue.location.part === "_rels/.rels",
    );

    assert.throws(
      () => importer(minimalTransitionalPptx({ omitPresentation: true })),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "broken-entry-relationship" &&
        error.issue.location.part === "_rels/.rels" &&
        error.issue.target === "ppt/presentation.xml",
    );

    assert.throws(
      () =>
        importer(
          minimalTransitionalPptx({
            presentationXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
          }),
        ),
      (error: unknown) =>
        error instanceof ImportError &&
        error.issue.code === "invalid-presentation-root" &&
        error.issue.location.part === "ppt/presentation.xml",
    );
  });

  it("imports a Transitional deck with unique slide IDs, isolates unsupported items, tracks unconsumed XML locations, and detaches from input bytes", () => {
    const importer = createImporter([]);
    const input = minimalTransitionalPptx({
      slideContent: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" customAttr="unexpected">
  <p:cSld name="">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:pic><p:nvPicPr><p:cNvPr id="2" name="pic"/></p:nvPicPr></p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
    });

    const result = importer(input);
    input.fill(0);

    const parsedReport = importReportSchema.parse(result.report);
    assert.equal(parsedReport.evidence, "import");
    assert.equal(
      parsedReport.sourceFormat,
      "ECMA-376 5th edition Transitional",
    );
    assert.equal(parsedReport.slideCount, 1);
    assert.equal(result.document.kind, "deck");
    assert.equal(result.document.props["width"], 960);
    assert.equal(result.document.props["height"], 540);
    assert.equal(result.document.children.length, 1);
    const slideNode = result.document.children[0];
    assert.ok(slideNode !== undefined && "id" in slideNode);
    assert.equal(slideNode.id, "slide-1");

    const parsedDiags = parsedReport.diagnostics.map((d) =>
      importDiagnosticSchema.parse(d),
    );
    const unsupported = parsedDiags.find(
      (d) => d.code === "unsupported-mechanism",
    );
    assert.ok(unsupported !== undefined);
    assert.equal(unsupported.localName, "pic");
    assert.equal(unsupported.location.part, "ppt/slides/slide1.xml");
    assert.equal(unsupported.location.slideIndex, 0);

    const unconsumed = parsedDiags.find((d) => d.code === "unconsumed-xml");
    assert.ok(unconsumed !== undefined);
    assert.equal(unconsumed.kind, "attribute");
    assert.equal(unconsumed.localName, "customAttr");
    assert.equal(unconsumed.location.part, "ppt/slides/slide1.xml");
    assert.equal(unconsumed.location.xpath, "/p:sld/@customAttr");
  });

  it("does not run master or section import when masters and section capabilities are omitted", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:extLst>
    <p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">
      <p14:sectionLst><p14:section name="Opening" id="{11111111-1111-1111-1111-111111111111}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst>
    </p:ext>
  </p:extLst>
</p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
</Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([]);
    const result = importer(pptx);
    assert.equal(result.document.kind, "deck");
    assert.equal(
      result.document.children.some(
        (child) =>
          typeof child === "object" &&
          child !== null &&
          "kind" in child &&
          child.kind === "slide-master",
      ),
      false,
    );
    assert.equal(
      result.document.children.some(
        (child) =>
          typeof child === "object" &&
          child !== null &&
          "kind" in child &&
          child.kind === "section",
      ),
      false,
    );
    assert.equal(result.document.children.length, 1);
    const onlySlide = result.document.children[0];
    assert.ok(onlySlide !== undefined && "id" in onlySlide);
    assert.equal(onlySlide.id, "s1");
  });
});
