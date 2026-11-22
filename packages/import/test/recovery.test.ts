import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import type { ObjectCapabilityModule } from "@pptxdk/core/capability";
import { CommonNode } from "@pptxdk/core/document";
import {
  A_NS,
  DIAGRAM_NS,
  MC_NS,
  OPC_REL_NS,
  P_NS,
  R_NS,
  REL_DIAGRAM_DATA,
  REL_HYPERLINK,
  REL_IMAGE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  createImporter,
  importDiagnosticSchema,
  openPptxPackage,
  readAlternateContent,
  buildRelationshipIndex,
  resolveExternalRelationshipTarget,
  resolveInternalRelationshipTarget,
} from "@pptxdk/import";
import { defineCanvasMechanism } from "@pptxdk/import";
import { strToU8, zipSync } from "fflate";
import { describe, it } from "vitest";

function buildTestPptx(
  entries: Record<string, string | Uint8Array>,
): Uint8Array {
  const normalized: Record<string, Uint8Array> = {};
  for (const [key, value] of Object.entries(entries)) {
    normalized[key] = typeof value === "string" ? strToU8(value) : value;
  }
  return zipSync(normalized);
}

function baseDeckEntries(
  slide1Xml: string,
  slide1RelsXml?: string,
): Record<string, string | Uint8Array> {
  return {
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/>
</Relationships>`,
    "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
</Relationships>`,
    "ppt/slides/slide1.xml": slide1Xml,
    ...(slide1RelsXml === undefined
      ? {}
      : { "ppt/slides/_rels/slide1.xml.rels": slide1RelsXml }),
  };
}

describe("@pptxdk/import recovery, resources, and mc:AlternateContent", () => {
  it("omits the bare SmartArt diagram in SmartArt.pptx with structured unsupported-mechanism and cascaded diagram dependencies while preserving the slide", () => {
    const smartArtBytes = new Uint8Array(
      readFileSync(new URL("./fixtures/SmartArt.pptx", import.meta.url)),
    );
    const importer = createImporter([]);
    const result = importer(smartArtBytes);
    smartArtBytes.fill(0);

    assert.equal(result.report.slideCount, 1);
    const slides = result.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide",
    );
    assert.equal(slides.length, 1);

    const parsedDiags = result.report.diagnostics.map((d) =>
      importDiagnosticSchema.parse(d),
    );
    const unsupported = parsedDiags.find(
      (d): d is Extract<typeof d, { readonly code: "unsupported-mechanism" }> =>
        d.code === "unsupported-mechanism" && d.localName === "graphicFrame",
    );
    assert.ok(unsupported !== undefined, "Expected graphicFrame diagnostic");
    assert.equal(unsupported.category, "unsupported");
    assert.equal(unsupported.disposition, "omitted");
    assert.equal(unsupported.location.part, "ppt/slides/slide1.xml");
    assert.equal(unsupported.location.slideIndex, 0);
    assert.equal(unsupported.location.objectId, 4);
    assert.equal(unsupported.location.objectName, "Diagram 3");
    assert.equal(unsupported.mechanismHint, DIAGRAM_NS);
    assert.ok(
      unsupported.cascaded !== undefined && unsupported.cascaded.length >= 4,
      "Expected cascaded diagram part omissions",
    );
    assert.ok(
      unsupported.cascaded.some(
        (c) => c.location.part === "ppt/diagrams/data1.xml",
      ),
    );
  });

  it("selects mc:Fallback when mc:Choice Requires an unsupported namespace and records compatibility-fallback", () => {
    const fallbackBoxMechanism = defineCanvasMechanism({
      key: "test-fallback-sp",
      role: "canvas",
      emit: () => ({ xml: "" }),
      paint: () => ({ svg: "" }),
      import: {
        target: "shape-tree",
        localName: "sp",
        read(ctx) {
          const nvSpPr = ctx.cursor.child(P_NS, "nvSpPr");
          const cNvPr = nvSpPr?.child(P_NS, "cNvPr");
          const idAttr = cNvPr?.attr("id");
          const nameAttr = cNvPr?.attr("name");
          nvSpPr?.ignoreEmptyChild(P_NS, "cNvSpPr");
          nvSpPr?.ignoreEmptyChild(P_NS, "nvPr");
          ctx.cursor.consumeChildrenWithoutVisiting();
          return {
            node: new CommonNode("shape", "shape", {
              ...(nameAttr === undefined ? {} : { name: nameAttr }),
              x: 10,
              y: 20,
              width: 100,
              height: 50,
            }),
            diagnostics: idAttr === undefined ? [] : [],
          };
        },
      },
    });

    const cap: ObjectCapabilityModule = {
      declaration: "shape",
      mechanisms: [fallbackBoxMechanism],
      bind: () => {
        throw new Error("not used in import test");
      },
    };

    const pptx = buildTestPptx(
      baseDeckEntries(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" xmlns:mc="${MC_NS}" xmlns:dgm="${DIAGRAM_NS}" mc:Ignorable="dgm">
  <p:cSld name="fallback-slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <mc:AlternateContent>
        <mc:Choice Requires="dgm">
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="4" name="SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          </p:graphicFrame>
        </mc:Choice>
        <mc:Fallback>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="5" name="FallbackShape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
          </p:sp>
        </mc:Fallback>
      </mc:AlternateContent>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`),
    );

    const importer = createImporter([cap]);
    const result = importer(pptx);
    const slideNode = result.document.children[0]!;
    assert.ok(slideNode instanceof CommonNode);
    assert.equal(slideNode.children.length, 1);
    const child = slideNode.children[0]!;
    assert.ok(child instanceof CommonNode);
    assert.equal(child.props["name"], "FallbackShape");

    const fallbackDiag = result.report.diagnostics.find(
      (d) => d.code === "compatibility-fallback",
    );
    assert.ok(fallbackDiag !== undefined);
    assert.equal(fallbackDiag.category, "unsupported");
    assert.equal(fallbackDiag.disposition, "fallback");
    assert.equal(fallbackDiag.location.slideIndex, 0);
    assert.equal(fallbackDiag.loss.lost, "native-mechanism");
    assert.equal(fallbackDiag.loss.retainedAs, "shape");
    // Rejected mc:Choice subtree must not emit noisy unconsumed-xml diagnostics.
    assert.equal(
      result.report.diagnostics.some((d) => d.code === "unconsumed-xml"),
      false,
    );
    void readAlternateContent;
  });

  it("resolves internal and external relationships cleanly and reports missing or broken targets", () => {
    const pptx = buildTestPptx(
      baseDeckEntries(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:cSld name="rels">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>`,
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_IMAGE}" Target="../media/missing.png"/>
  <Relationship Id="rId2" Type="${REL_IMAGE}" Target="https://example.com/photo.png" TargetMode="External"/>
  <Relationship Id="rId3" Type="${REL_HYPERLINK}" Target="https://example.com/docs" TargetMode="External"/>
  <Relationship Id="rId4" Type="${REL_DIAGRAM_DATA}" Target="../diagrams/data1.xml"/>
</Relationships>`,
      ),
    );

    const pkg = openPptxPackage(pptx);
    const { byId } = buildRelationshipIndex(pkg, "ppt/slides/slide1.xml");
    const loc = { part: "ppt/slides/slide1.xml", slideIndex: 0 };

    const missing = resolveInternalRelationshipTarget(
      pkg,
      byId,
      "ppt/slides/slide1.xml",
      "rId1",
      loc,
      REL_IMAGE,
    );
    assert.equal(missing.ok, false);
    if (!missing.ok) {
      assert.equal(missing.diagnostic.code, "missing-target-part");
      assert.equal(missing.diagnostic.targetPart, "ppt/media/missing.png");
    }

    const broken = resolveInternalRelationshipTarget(
      pkg,
      byId,
      "ppt/slides/slide1.xml",
      "rId99",
      loc,
      REL_IMAGE,
    );
    assert.equal(broken.ok, false);
    if (!broken.ok) {
      assert.equal(broken.diagnostic.code, "broken-relationship");
    }

    const extImg = resolveExternalRelationshipTarget(
      byId,
      "rId2",
      loc,
      REL_IMAGE,
    );
    assert.equal(extImg.ok, true);
    if (extImg.ok) {
      assert.equal(extImg.target, "https://example.com/photo.png");
    }

    const extLink = resolveExternalRelationshipTarget(
      byId,
      "rId3",
      loc,
      REL_HYPERLINK,
    );
    assert.equal(extLink.ok, true);
    if (extLink.ok) {
      assert.equal(extLink.target, "https://example.com/docs");
    }
  });
});
