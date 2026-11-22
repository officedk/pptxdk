import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, image, slide, text } from "@pptxdk/capabilities/document";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import {
  A_NS,
  DIAGRAM_NS,
  MC_NS,
  OPC_REL_NS,
  P_NS,
  REL_HYPERLINK,
  REL_IMAGE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { describe, it } from "vitest";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

function createSamplePngBytes(): Uint8Array {
  const png = new PNG({ width: 24, height: 12 });
  for (let y = 0; y < 12; y++) {
    for (let x = 0; x < 24; x++) {
      png.data.set(
        x < 12 ? [230, 40, 40, 255] : [30, 130, 230, 255],
        (y * 24 + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(png));
}

const samplePng = createSamplePngBytes();

describe("Ticket 04: Image roundtrip, external links, and partial loss recovery", () => {
  it("detaches embedded image bytes on import and keeps image, crop, and links editable after zeroing inputBytes", async () => {
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "page-1" }).append(
        image({
          id: "hero",
          name: "HeroPhoto",
          src: samplePng,
          x: 30,
          y: 40,
          width: 160,
          height: 80,
          crop: { left: 0.25 },
          link: { id: "page-2" },
          style: { opacity: 0.8 },
        }),
      ),
      slide({ id: "page-2" }).append(
        image({
          id: "logo",
          name: "Logo",
          src: samplePng,
          x: 40,
          y: 40,
          width: 80,
          height: 40,
          link: { url: "https://example.com/about" },
        }),
      ),
    );

    const emitted = emit(resolveDocument(doc));
    const inputCopy = new Uint8Array(emitted.pptx);
    const imported = importDocument(inputCopy);
    // Prove zero retention of caller-owned ZIP bytes.
    inputCopy.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);
    const slide1 = imported.document.children[0] as CommonNode;
    const importedHero = slide1.children[0] as CommonNode;
    assert.equal(importedHero.kind, "image");
    assert.deepEqual(
      { ...(importedHero.props["link"] as Record<string, unknown>) },
      { id: "page-2" },
    );
    assert.deepEqual(
      { ...(importedHero.props["crop"] as Record<string, unknown>) },
      { left: 0.25 },
    );

    // Edit crop & position on the imported image node, then re-emit and paint offline.
    importedHero.set({ x: 80, crop: { right: 0.5 } });
    const reResolved = resolveDocument(imported.document);
    const reEmitted = emit(reResolved);
    const rePainted = paint(reResolved);
    const validation = await validatePptx(reEmitted.pptx);

    assert.deepEqual(validation.errors, []);
    assert.deepEqual(rePainted.pages[0]?.diagnostics, []);
    assert.ok(rePainted.pages[0]?.png instanceof Uint8Array);
    const reHero = validation.slides[0]?.objects[0];
    assert.equal(reHero?.boxEmu.x, 80 * 9525);
    assert.deepEqual(reHero?.crop, { right: 50000 });
    assert.deepEqual(reHero?.link, { slideId: "page-2" });
  });

  it("preserves external linked images (r:link) and external web links with zero network access during import and re-emit", async () => {
    const originalFetch = globalThis.fetch;
    let networkCalls = 0;
    globalThis.fetch = () => {
      networkCalls += 1;
      throw new Error("Network access is forbidden during offline import");
    };

    try {
      const doc = deck({ width: 640, height: 360 }).append(
        slide({ id: "ext-slide" }).append(
          image({
            id: "ext-img",
            name: "ExternalBanner",
            src: { externalUrl: "https://cdn.example.com/banner.png" },
            x: 40,
            y: 50,
            width: 240,
            height: 120,
            link: { url: "https://example.com/landing" },
          }),
          image({
            id: "local-img",
            name: "LocalSibling",
            src: samplePng,
            x: 320,
            y: 50,
            width: 120,
            height: 60,
          }),
        ),
      );

      const initialEmit = emit(resolveDocument(doc));
      const initialValidation = await validatePptx(initialEmit.pptx);
      assert.deepEqual(initialValidation.errors, []);
      assert.equal(
        initialValidation.slides[0]?.objects[0]?.image?.externalTarget,
        "https://cdn.example.com/banner.png",
      );

      const imported = importDocument(new Uint8Array(initialEmit.pptx));
      assert.equal(networkCalls, 0);
      assert.deepEqual(imported.report.diagnostics, []);

      const slideNode = imported.document.children[0] as CommonNode;
      const importedExt = slideNode.children[0] as CommonNode;
      assert.deepEqual(
        { ...(importedExt.props["src"] as Record<string, unknown>) },
        { externalUrl: "https://cdn.example.com/banner.png" },
      );
      assert.deepEqual(
        { ...(importedExt.props["link"] as Record<string, unknown>) },
        { url: "https://example.com/landing" },
      );

      const reResolved = resolveDocument(imported.document);
      const reEmitted = emit(reResolved);
      const rePainted = paint(reResolved);
      const reValidation = await validatePptx(reEmitted.pptx);

      assert.equal(networkCalls, 0);
      assert.deepEqual(reValidation.errors, []);
      assert.equal(
        reValidation.slides[0]?.objects[0]?.image?.externalTarget,
        "https://cdn.example.com/banner.png",
      );
      assert.deepEqual(reValidation.slides[0]?.objects[0]?.link, {
        url: "https://example.com/landing",
      });

      // Offline paint emits a warning for the external image placeholder while still rendering sibling content to PNG.
      assert.ok(
        rePainted.pages[0]?.diagnostics.some((d) => d.severity === "warning"),
      );
      assert.ok(rePainted.pages[0]?.png instanceof Uint8Array);

      const zip = unzipSync(reEmitted.pptx);
      const slide1Xml = strFromU8(zip["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:blip r:link="rId/u);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("isolates missing media parts, corrupt image bytes, and broken .rels references while preserving adjacent editable content", async () => {
    const validSibling = text("Healthy Sibling", {
      id: "healthy-text",
      name: "HealthySibling",
      x: 40,
      y: 220,
      width: 240,
      height: 40,
      style: { fontSize: "16pt", color: "#0F172A" },
    });
    const base = emit(
      resolveDocument(
        deck({ width: 640, height: 360 }).append(
          slide({ id: "recovery-slide" }).append(
            image({
              id: "missing-media",
              name: "MissingPic",
              src: samplePng,
              x: 40,
              y: 30,
              width: 100,
              height: 50,
              link: { url: "https://example.com/broken-pic-link" },
            }),
            validSibling,
          ),
        ),
      ),
    );

    const zipEntries = unzipSync(base.pptx);
    // 1. Delete the media part while keeping slide1.xml and slide1.xml.rels referencing it.
    const mediaKey = Object.keys(zipEntries).find((k) =>
      k.startsWith("ppt/media/"),
    )!;
    delete zipEntries[mediaKey];
    const missingMediaResult = importDocument(zipSync(zipEntries));
    const missingDiag = missingMediaResult.report.diagnostics.find(
      (d) => d.code === "missing-target-part",
    );
    assert.ok(missingDiag !== undefined);
    assert.equal(missingDiag.category, "corrupt-input");
    assert.equal(missingDiag.disposition, "omitted");
    assert.equal(missingDiag.location.part, "ppt/slides/slide1.xml");
    assert.equal(missingDiag.location.slideIndex, 0);
    assert.equal(missingDiag.location.objectName, "MissingPic");
    assert.equal(missingDiag.targetPart, mediaKey);
    assert.ok(
      missingDiag.cascaded?.some((c) => c.role === "link"),
      "Expected cascaded link omission on omitted picture",
    );

    // Adjacent valid text object survives and remains editable & re-emittable!
    const recoveredSlide = missingMediaResult.document
      .children[0] as CommonNode;
    assert.equal(recoveredSlide.children.length, 1);
    const siblingNode = recoveredSlide.children[0] as CommonNode;
    assert.equal(siblingNode.kind, "text");
    assert.equal(siblingNode.props["name"], "HealthySibling");
    siblingNode.set({ x: 60 });
    const reEmitted = emit(resolveDocument(missingMediaResult.document));
    const reValidated = await validatePptx(reEmitted.pptx);
    assert.deepEqual(reValidated.errors, []);

    // 2. Replace media bytes with corrupt non-image bytes -> corrupt-part diagnostic + sibling preserved.
    zipEntries[mediaKey] = new Uint8Array([0x00, 0x11, 0x22, 0x33]);
    const corruptBytesResult = importDocument(zipSync(zipEntries));
    const corruptDiag = corruptBytesResult.report.diagnostics.find(
      (d) => d.code === "corrupt-part",
    );
    assert.ok(corruptDiag !== undefined);
    assert.equal(corruptDiag.category, "corrupt-input");
    assert.equal(corruptDiag.disposition, "omitted");
    assert.equal(corruptDiag.location.objectName, "MissingPic");
    assert.ok(
      corruptDiag.cascaded?.some(
        (c) => c.role === "media" && c.location.part === mediaKey,
      ),
    );
    assert.equal(
      (corruptBytesResult.document.children[0] as CommonNode).children.length,
      1,
    );

    // 3. Corrupt a link relationship on a healthy image -> broken-relationship omits ONLY the link and keeps the image!
    const healthyBase = unzipSync(
      emit(
        resolveDocument(
          deck({ width: 640, height: 360 }).append(
            slide({ id: "broken-link-slide" }).append(
              image({
                id: "healthy-img",
                name: "HealthyImage",
                src: samplePng,
                x: 40,
                y: 40,
                width: 120,
                height: 60,
                link: { url: "https://example.com/ok" },
              }),
            ),
          ),
        ),
      ).pptx,
    );
    const relsXml = strFromU8(
      healthyBase["ppt/slides/_rels/slide1.xml.rels"]!,
    ).replace(/Target="https:\/\/example\.com\/ok"/u, 'Target=""');
    healthyBase["ppt/slides/_rels/slide1.xml.rels"] = strToU8(relsXml);
    const brokenRelResult = importDocument(zipSync(healthyBase));
    const brokenRelDiag = brokenRelResult.report.diagnostics.find(
      (d) => d.code === "broken-relationship",
    );
    assert.ok(brokenRelDiag !== undefined);
    assert.equal(brokenRelDiag.disposition, "omitted");
    const keptImage = (brokenRelResult.document.children[0] as CommonNode)
      .children[0] as CommonNode;
    assert.equal(keptImage.kind, "image");
    assert.equal(keptImage.props["link"], undefined);
  });

  it("imports positioned mc:Fallback <p:pic> from mc:AlternateContent with compatibility-fallback loss diagnostics", async () => {
    const pptx = zipSync({
      "_rels/.rels":
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/>
</Relationships>`),
      "ppt/presentation.xml":
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="6096000" cy="3429000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`),
      "ppt/_rels/presentation.xml.rels":
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/>
</Relationships>`),
      "ppt/slides/_rels/slide1.xml.rels":
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${OPC_REL_NS}">
  <Relationship Id="rId1" Type="${REL_IMAGE}" Target="../media/fallback.png"/>
  <Relationship Id="rId2" Type="${REL_HYPERLINK}" Target="https://example.com/fallback" TargetMode="External"/>
</Relationships>`),
      "ppt/media/fallback.png": samplePng,
      "ppt/slides/slide1.xml":
        strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" xmlns:mc="${MC_NS}" xmlns:dgm="${DIAGRAM_NS}" mc:Ignorable="dgm">
  <p:cSld name="mc-slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <mc:AlternateContent>
        <mc:Choice Requires="dgm">
          <p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="2" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="952500"/></p:xfrm>
            <a:graphic><a:graphicData uri="${DIAGRAM_NS}"/></a:graphic>
          </p:graphicFrame>
        </mc:Choice>
        <mc:Fallback>
          <p:pic>
            <p:nvPicPr>
              <p:cNvPr id="3" name="DiagramFallbackPic"><a:hlinkClick r:id="rId2"/></p:cNvPr>
              <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
              <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="rId1"/>
              <a:stretch><a:fillRect/></a:stretch>
            </p:blipFill>
            <p:spPr>
              <a:xfrm><a:off x="476250" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
          </p:pic>
        </mc:Fallback>
      </mc:AlternateContent>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`),
    });

    const imported = importDocument(pptx);
    const slideNode = imported.document.children[0] as CommonNode;
    assert.equal(slideNode.children.length, 1);
    const fallbackPic = slideNode.children[0] as CommonNode;
    assert.equal(fallbackPic.kind, "image");
    assert.equal(fallbackPic.props["name"], "DiagramFallbackPic");
    assert.equal(fallbackPic.props["x"], 50);
    assert.equal(fallbackPic.props["y"], 50);
    assert.equal(fallbackPic.props["width"], 200);
    assert.equal(fallbackPic.props["height"], 100);
    assert.deepEqual(
      { ...(fallbackPic.props["link"] as Record<string, unknown>) },
      { url: "https://example.com/fallback" },
    );

    const fallbackDiag = imported.report.diagnostics.find(
      (d) => d.code === "compatibility-fallback",
    );
    assert.ok(fallbackDiag !== undefined);
    assert.equal(fallbackDiag.category, "unsupported");
    assert.equal(fallbackDiag.disposition, "fallback");
    assert.equal(fallbackDiag.selectedFallback, "image");
    assert.deepEqual(fallbackDiag.loss, {
      lost: "native-mechanism",
      retainedAs: "image",
    });

    // Re-resolve, re-emit, paint, and validate the recovered fallback image!
    const resolved = resolveDocument(imported.document);
    const reEmitted = emit(resolved);
    const painted = paint(resolved);
    const validation = await validatePptx(reEmitted.pptx);
    assert.deepEqual(validation.errors, []);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);
  });
});
