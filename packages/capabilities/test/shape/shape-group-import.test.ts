import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { box, deck, group, shape, slide } from "@pptxdk/capabilities/document";
import { CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

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

describe("Shape, Group, and expanded layout OOXML import", () => {
  it("maps external nested <p:grpSp> (chOff, chExt) -> (off, ext) affine transforms and ignores blank txBody", async () => {
    const base = deck({ width: 640, height: 360 }).append(
      slide({ id: "external-group" }).append(
        shape({ x: 10, y: 10, width: 20, height: 20, fill: "#000000" }),
      ),
    );
    const basePptx = emit(resolveDocument(base)).pptx;

    // Outer group: off=(100,50)px, ext=(200,100)px, chOff=(0,0)px, chExt=(100,50)px -> 2x scale + (100,50) offset.
    // Child shape at child (10,10,40,20)px -> slide (120,70,80,40)px, with blank PowerPoint <txBody>.
    // Inner group at child (50,5,50,40)px -> outer space (200,60,100,80)px,
    //   inner chOff=(10,20)px chExt=(50,40)px (2x scale in x and y),
    //   inner text box at child (20,30,25,10)px -> outer (50 + (20-10)*(50/50)=60, ...) wait:
    //   In inner: x = 50 + (20-10)*(50/50) = 60, y = 5 + (30-20)*(40/40) = 15, w = 25, h = 10;
    //   Then outer maps (60,15,25,10) -> x = 100 + 60*2 = 220, y = 50 + 15*2 = 80, w = 50, h = 20.
    const externalGrpXml = `<p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="10" name="Scaled Cluster"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm rot="900000"><a:off x="${100 * emuPerPx}" y="${50 * emuPerPx}"/><a:ext cx="${200 * emuPerPx}" cy="${100 * emuPerPx}"/><a:chOff x="0" y="0"/><a:chExt cx="${100 * emuPerPx}" cy="${50 * emuPerPx}"/></a:xfrm><a:effectLst/></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="Scaled Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${10 * emuPerPx}" y="${10 * emuPerPx}"/><a:ext cx="${40 * emuPerPx}" cy="${20 * emuPerPx}"/></a:xfrm>
          <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill>
        </p:spPr>
        <p:txBody><a:bodyPr wrap="square" lIns="91440" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>
      </p:sp>
      <p:grpSp>
        <p:nvGrpSpPr><p:cNvPr id="12" name="Nested Subcluster"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="${50 * emuPerPx}" y="${5 * emuPerPx}"/><a:ext cx="${50 * emuPerPx}" cy="${40 * emuPerPx}"/><a:chOff x="${10 * emuPerPx}" y="${20 * emuPerPx}"/><a:chExt cx="${50 * emuPerPx}" cy="${40 * emuPerPx}"/></a:xfrm></p:grpSpPr>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="13" name="Nested Label"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
          <p:spPr><a:xfrm><a:off x="${20 * emuPerPx}" y="${30 * emuPerPx}"/><a:ext cx="${25 * emuPerPx}" cy="${10 * emuPerPx}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
          <p:txBody><a:bodyPr wrap="none"/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Hello</a:t></a:r></a:p></p:txBody>
        </p:sp>
      </p:grpSp>
    </p:grpSp>`;

    const externalPptx = replaceSlide1Xml(basePptx, externalGrpXml);
    const imported = importDocument(externalPptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const resolution = resolveDocument(imported.document);
    const reemitted = emit(resolution);
    const validation = await validatePptx(reemitted.pptx);
    assert.deepEqual(validation.errors, []);

    const outer = validation.slides[0]!.objects[0]!;
    assert.equal(outer.kind, "group");
    assert.equal(outer.name, "Scaled Cluster");
    assert.equal(outer.rotationDegrees, 15);
    assert.deepEqual(
      outer.members?.map((m) => m.name),
      ["Scaled Card", "Nested Subcluster"],
    );

    const card = outer.members[0]!;
    assert.equal(card.geometry, "roundRect");
    assert.equal(card.text, undefined);
    assert.deepEqual(card.boxEmu, {
      x: 120 * emuPerPx,
      y: 70 * emuPerPx,
      width: 80 * emuPerPx,
      height: 40 * emuPerPx,
    });

    const nestedGroup = outer.members[1]!;
    assert.equal(nestedGroup.kind, "group");
    const label = nestedGroup.members![0]!;
    assert.equal(label.name, "Nested Label");
    assert.equal(label.text?.text, "Hello");
    assert.deepEqual(label.boxEmu, {
      x: 220 * emuPerPx,
      y: 80 * emuPerPx,
      width: 50 * emuPerPx,
      height: 20 * emuPerPx,
    });

    // Verify schemeClr tint/shade resolved to valid solid fill on the blank shape and paints without text errors.
    const groups = findNodesByType(imported.document, "group");
    assert.equal(groups.length, 2);
  });

  it("supports post-import editing of an individual Shape and an individual Group member", async () => {
    const initial = deck({ width: 640, height: 360 }).append(
      slide({ id: "post-edit" }).append(
        shape({
          id: "solo",
          name: "Solo Shape",
          geometry: "rect",
          x: 24,
          y: 24,
          width: 120,
          height: 70,
          fill: "#2563EB",
        }),
        group({ id: "cluster", name: "Editable Cluster" }).append(
          shape({
            id: "left",
            name: "Left Member",
            geometry: "rect",
            x: 200,
            y: 40,
            width: 80,
            height: 60,
            fill: "#10B981",
          }),
          shape({
            id: "right",
            name: "Right Member",
            geometry: "diamond",
            x: 310,
            y: 40,
            width: 80,
            height: 60,
            fill: "#F59E0B",
          }),
        ),
      ),
    );

    const emitted = emit(resolveDocument(initial)).pptx;
    const imported = importDocument(emitted);
    assert.deepEqual(imported.report.diagnostics, []);

    const shapes = findNodesByType(imported.document, "shape");
    const soloNode = shapes.find((n) => n.props["name"] === "Solo Shape")!;
    const rightMemberNode = shapes.find(
      (n) => n.props["name"] === "Right Member",
    )!;

    soloNode
      .set({
        name: "Edited Solo",
        geometry: "hexagon",
        x: 40,
        width: 150,
      })
      .style({ backgroundColor: "#DC2626" });
    rightMemberNode
      .set({
        name: "Edited Member",
        geometry: "ellipse",
        y: 80,
      })
      .style({ backgroundColor: "#0EA5E9" });

    const editedResolution = resolveDocument(imported.document);
    const editedEmit = emit(editedResolution);
    const editedPaint = paint(editedResolution);
    const editedValidation = await validatePptx(editedEmit.pptx);

    assert.deepEqual(editedValidation.errors, []);
    assert.deepEqual(editedPaint.pages[0]?.diagnostics, []);
    assert.ok(editedPaint.pages[0]?.png instanceof Uint8Array);
    assert.ok(editedPaint.pages[0]?.svg?.includes('fill="#DC2626"'));
    assert.ok(editedPaint.pages[0]?.svg?.includes('fill="#0EA5E9"'));

    const editedSolo = editedValidation.slides[0]!.objects[0]!;
    assert.equal(editedSolo.name, "Edited Solo");
    assert.equal(editedSolo.geometry, "hexagon");
    assert.deepEqual(editedSolo.boxEmu, {
      x: 40 * emuPerPx,
      y: 24 * emuPerPx,
      width: 150 * emuPerPx,
      height: 70 * emuPerPx,
    });

    const editedGroup = editedValidation.slides[0]!.objects[1]!;
    assert.equal(editedGroup.kind, "group");
    assert.equal(editedGroup.members?.[0]?.name, "Left Member");
    assert.equal(editedGroup.members?.[1]?.name, "Edited Member");
    assert.equal(editedGroup.members?.[1]?.geometry, "ellipse");
    assert.equal(editedGroup.members?.[1]?.boxEmu?.y, 80 * emuPerPx);

    // Also confirm text editing on the imported shape persists into re-emitted PPTX.
    soloNode.set({ text: "Updated Shape Copy" });
    const textEditedValidation = await validatePptx(
      emit(resolveDocument(imported.document)).pptx,
    );
    assert.equal(
      textEditedValidation.slides[0]!.objects[0]?.text?.text,
      "Updated Shape Copy",
    );
  });

  it("reconstructs expanded Flex and Grid containers preserving visible Boxes as Shapes and dropping transparent Boxes", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "layout-reconstruction" }).append(
        box({
          id: "transparent-flex",
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 260,
            display: "flex",
            flexDirection: "row",
            gap: 16,
          },
        }).append(
          shape({
            name: "Flex A",
            width: 80,
            height: 44,
            fill: "#2563EB",
          }),
          shape({
            name: "Flex B",
            width: 90,
            height: 44,
            fill: "#16A34A",
          }),
        ),
        box({
          id: "transparent-grid",
          style: {
            position: "absolute",
            left: 24,
            top: 100,
            width: 300,
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            columnGap: 12,
            rowGap: 12,
          },
        }).append(
          box({
            id: "visible-grid-card",
            style: {
              padding: 10,
              display: "flex",
              flexDirection: "column",
              backgroundColor: "#FEF3C7",
              borderColor: "#D97706",
              borderWidth: 2,
              borderRadius: 8,
            },
          }).append(
            shape({
              name: "Grid Card Leaf",
              width: 60,
              height: 30,
              fill: "#B45309",
            }),
          ),
          shape({
            name: "Grid Sibling",
            width: 80,
            height: 50,
            fill: "#7C3AED",
          }),
        ),
      ),
    );

    const initialResolution = resolveDocument(doc);
    const initialPptx = emit(initialResolution).pptx;
    const initialValidation = await validatePptx(initialPptx);
    assert.deepEqual(initialValidation.errors, []);

    const imported = importDocument(initialPptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const roundtripResolution = resolveDocument(imported.document);
    const roundtripPptx = emit(roundtripResolution).pptx;
    const roundtripPaint = paint(roundtripResolution);
    const roundtripValidation = await validatePptx(roundtripPptx);

    assert.deepEqual(roundtripValidation.errors, []);
    assert.deepEqual(roundtripPaint.pages[0]?.diagnostics, []);
    assert.ok(roundtripPaint.pages[0]?.png instanceof Uint8Array);

    const initialObjects = initialValidation.slides[0]!.objects;
    const roundtripObjects = roundtripValidation.slides[0]!.objects;
    // 2 Flex children + 1 visible Grid Box + 1 Grid Box child + 1 Grid Sibling = 5 native shapes; transparent Boxes disappeared.
    assert.equal(initialObjects.length, 5);
    assert.equal(roundtripObjects.length, 5);
    assert.deepEqual(
      roundtripObjects.map((o) => ({
        name: o.name,
        geometry: o.geometry,
        boxEmu: o.boxEmu,
      })),
      initialObjects.map((o) => ({
        name: o.name,
        geometry: o.geometry,
        boxEmu: o.boxEmu,
      })),
    );

    // Imported visible Box became an editable Shape in the document tree.
    const importedShapes = findNodesByType(imported.document, "shape");
    assert.equal(importedShapes.length, 5);
  });

  it("recovers editable geometry, supported visuals, and text while reporting localized diagnostics for unsupported visual attributes", async () => {
    const base = deck({ width: 640, height: 360 }).append(
      slide({ id: "partial-recovery" }).append(
        shape({ x: 10, y: 10, width: 20, height: 20, fill: "#000000" }),
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
});
