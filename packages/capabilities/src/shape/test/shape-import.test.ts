import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, shape, slide } from "@pptxdk/capabilities/document";
import { CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { customBezierShape } from "./mechanisms.ts";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);
const emuPerPx = 9525;

function replaceSlide1Xml(pptx: Uint8Array, spTreeBodyXml: string): Uint8Array {
  const parts = unzipSync(pptx);
  const slide1 = strFromU8(parts["ppt/slides/slide1.xml"]!);
  const startMarker = "<p:spTree>";
  const endMarker = "</p:spTree>";
  const start = slide1.indexOf(startMarker);
  const end = slide1.indexOf(endMarker);
  assert.ok(start >= 0 && end > start, "Expected <p:spTree> in slide1.xml");
  const header = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
  const updatedSlide1 = `${slide1.slice(0, start + startMarker.length)}${header}${spTreeBodyXml}${slide1.slice(end)}`;
  parts["ppt/slides/slide1.xml"] = strToU8(updatedSlide1);
  return zipSync(parts);
}

function findNodesByType(
  root: CommonNode,
  type: string,
): readonly CommonNode[] {
  const result: CommonNode[] = [];
  const visit = (node: CommonNode): void => {
    if (node.type === type) result.push(node);
    for (const child of node.children) {
      if (child instanceof CommonNode) visit(child);
    }
  };
  visit(root);
  return result;
}

describe("Shape OOXML import", () => {
  it("recovers editable geometry, supported visuals, and text while reporting localized diagnostics for unsupported visual attributes", async () => {
    const base = deck({ width: 640, height: 360 }).append(
      slide({ id: "partial-recovery" }).append(
        shape({
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 20,
            height: 20,
            background: "#000000",
          },
        }),
      ),
    );
    const basePptx = emit(resolveDocument(base)).pptx;

    const mixedXml = `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="42" name="Partial Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${32 * emuPerPx}" y="${40 * emuPerPx}"/><a:ext cx="${180 * emuPerPx}" cy="${80 * emuPerPx}"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
          <a:pattFill prst="pct20"/>
          <a:ln w="${2 * emuPerPx}"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill><a:prstDash val="lgDash"/></a:ln>
          <a:effectLst>
            <a:outerShdw blurRad="${3 * emuPerPx}" dist="${4 * emuPerPx}" dir="5400000"><a:srgbClr val="000000"/></a:outerShdw>
            <a:innerShdw blurRad="${2 * emuPerPx}" dist="${2 * emuPerPx}" dir="0"/>
          </a:effectLst>
          <a:sp3d extrusionH="76200"/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr"/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr sz="1600"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Preserved Text</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="43" name="Arc With Guide"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${240 * emuPerPx}" y="${40 * emuPerPx}"/><a:ext cx="${100 * emuPerPx}" cy="${80 * emuPerPx}"/></a:xfrm>
          <a:custGeom>
            <a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="100" b="80"/>
            <a:pathLst>
              <a:path w="100" h="80">
                <a:moveTo><a:pt x="10" y="40"/></a:moveTo>
                <a:arcTo wR="30" hR="20" stAng="10800000" swAng="5400000"/>
                <a:lnTo><a:pt x="adj1" y="40"/></a:lnTo>
              </a:path>
            </a:pathLst>
          </a:custGeom>
          <a:blipFill><a:blip r:embed="rId999"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
        </p:spPr>
      </p:sp>`;

    const externalPptx = replaceSlide1Xml(basePptx, mixedXml);
    const imported = importDocument(externalPptx);

    const diagnostics = imported.report.diagnostics;
    const pattDiag = diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        "localName" in d &&
        d.localName === "pattFill",
    );
    assert.ok(pattDiag, "Expected localized diagnostic for <a:pattFill>");
    assert.equal(pattDiag.disposition, "omitted");
    assert.equal(pattDiag.location.part, "ppt/slides/slide1.xml");
    assert.equal(pattDiag.location.slideIndex, 0);
    assert.equal(pattDiag.location.objectId, 42);
    assert.equal(pattDiag.location.objectName, "Partial Card");
    assert.match(pattDiag.location.xpath ?? "", /a:pattFill/u);

    const sp3dDiag = diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        "localName" in d &&
        d.localName === "sp3d",
    );
    assert.ok(sp3dDiag, "Expected localized diagnostic for <a:sp3d>");
    assert.equal(sp3dDiag.location.objectId, 42);

    const innerShdwDiag = diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        "localName" in d &&
        d.localName === "innerShdw",
    );
    assert.ok(innerShdwDiag, "Expected localized diagnostic for <a:innerShdw>");
    assert.equal(innerShdwDiag.location.objectId, 42);

    const guideDiag = diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        "mechanismHint" in d &&
        d.mechanismHint === "custom-geometry-guide-reference",
    );
    assert.ok(guideDiag, "Expected localized diagnostic for custGeom formula");
    assert.equal(guideDiag.location.objectId, 43);
    assert.equal(guideDiag.location.objectName, "Arc With Guide");

    const blipDiag = diagnostics.find(
      (d) =>
        d.code === "broken-relationship" &&
        "relationshipId" in d &&
        d.relationshipId === "rId999",
    );
    assert.ok(blipDiag, "Expected broken-relationship diagnostic for rId999");
    assert.equal(blipDiag.location.objectId, 43);

    // Partial Card must survive with its editable geometry, border, outer shadow, and text intact.
    const resolution = resolveDocument(imported.document);
    const reemitted = emit(resolution);
    const validation = await validatePptx(reemitted.pptx);
    assert.deepEqual(validation.errors, []);

    const recoveredCard = validation.slides[0]!.objects.find(
      (o) => o.name === "Partial Card",
    );
    assert.ok(recoveredCard);
    assert.equal(recoveredCard.geometry, "roundRect");
    assert.equal(recoveredCard.text?.text, "Preserved Text");

    const recoveredXml = strFromU8(
      unzipSync(reemitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(recoveredXml, /<a:prstDash val="dash"\/>/u);
    assert.match(recoveredXml, /<a:outerShdw/u);
  });

  it("recovers custom bezier path control points cleanly from pure PPTX bytes", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "custom-path-recovery" }).append(
        shape({
          id: "custom-radial",
          name: "Custom Radial",
          geometry: customBezierShape,
          style: {
            position: "absolute",
            left: 24,
            top: 140,
            width: 140,
            height: 100,
            flipV: true,
            background: "#D97706",
          },
        }),
      ),
    );
    const emitted = emit(resolveDocument(doc)).pptx;
    const buffer = new Uint8Array(emitted);
    const imported = importDocument(buffer);
    buffer.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);
    const importedShape = findNodesByType(imported.document, "shape").find(
      (node) => node.props["name"] === "Custom Radial",
    );
    assert.ok(importedShape !== undefined);
    assert.deepEqual(
      JSON.parse(JSON.stringify(importedShape.props["geometry"])),
      {
        type: "custom",
        width: 100,
        height: 100,
        paths: [{ commands: customBezierShape.paths[0]!.commands }],
      },
    );
  });

  it("omits border arrows on external open-path shapes and records a shape-line-arrow diagnostic", async () => {
    const base = deck({ width: 640, height: 360 }).append(
      slide({ id: "open-path-arrow" }),
    );
    const basePptx = emit(resolveDocument(base)).pptx;
    const spXml = `
      <p:sp>
        <p:nvSpPr><p:cNvPr id="77" name="Open Arc With Arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${40 * emuPerPx}" y="${40 * emuPerPx}"/><a:ext cx="${120 * emuPerPx}" cy="${60 * emuPerPx}"/></a:xfrm>
          <a:prstGeom prst="arc"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="${2 * emuPerPx}">
            <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
            <a:tailEnd type="triangle"/>
          </a:ln>
        </p:spPr>
      </p:sp>`;
    const imported = importDocument(replaceSlide1Xml(basePptx, spXml));
    const arrowDiag = imported.report.diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        "mechanismHint" in d &&
        d.mechanismHint === "shape-line-arrow",
    );
    assert.ok(arrowDiag, "Expected shape-line-arrow diagnostic");
    assert.equal(arrowDiag.disposition, "omitted");
    assert.equal(arrowDiag.location.objectId, 77);
    const importedShape = findNodesByType(imported.document, "shape").find(
      (node) => node.props["name"] === "Open Arc With Arrow",
    );
    assert.ok(importedShape !== undefined);
    const style = importedShape.props["style"] as {
      border?: Record<string, unknown>;
    };
    assert.equal(style.border?.["startArrow"], undefined);
    assert.equal(style.border?.["endArrow"], undefined);
  });
});
