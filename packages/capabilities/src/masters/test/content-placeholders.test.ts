import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { font } from "@pptxdk/capabilities/document";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode, SlideContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

const expectedFixtureSha256: Readonly<Record<string, string>> = {
  "ph-populated-placeholders.pptx":
    "b08599276b4d708757951f2d45ac74cb4e81e94463d575e4d857690d0d231a60",
  "ph-unpopulated-placeholders.pptx":
    "4a73c52d57e944abfc3367d7521f0ae8740030309af5644f6702ded0508cca53",
};

function readFixture(name: string): Uint8Array {
  const bytes = new Uint8Array(
    readFileSync(new URL(`../../../test/fixtures/${name}`, import.meta.url)),
  );
  const expectedSha = expectedFixtureSha256[name];
  expect(expectedSha).toBeDefined();
  expect(createHash("sha256").update(bytes).digest("hex")).toBe(expectedSha);
  return bytes;
}

describe("Native content placeholders fixtures", () => {
  it("imports, emits, validates, and paints ph-populated-placeholders.pptx cleanly", async () => {
    const bytes = readFixture("ph-populated-placeholders.pptx");
    const imported = importDocument(bytes);

    // Every diagnostic on the raw Office fixture is an unsupported non-core mechanism (theme 3D, slide 3 tableStyleId, slide 8 SmartArt dgm).
    expect(
      imported.report.diagnostics.every((d) => d.category === "unsupported"),
    ).toBe(true);
    expect(
      imported.report.diagnostics.some(
        (d) =>
          d.code === "unsupported-mechanism" && d.location.slideIndex === 7,
      ),
    ).toBe(true);

    const resolution = resolveDocument(imported.document);
    const { pptx } = emit(resolution);
    const validation = await validatePptx(pptx);
    expect(validation.errors).toEqual([]);
    expect(validation.slides).toHaveLength(9);

    // Slide 1: pic placeholder with omitted local transform (<p:spPr/>) and inherited layout boxEmu
    const s1Pic = validation.slides[0]?.objects[0];
    expect(s1Pic?.kind).toBe("image");
    expect(s1Pic?.placeholder).toEqual({ type: "pic", idx: 10 });
    expect(s1Pic?.hasLocalTransform).toBe(false);
    expect(s1Pic?.crop).toEqual({ top: 23715, bottom: 23715 });
    expect(s1Pic?.boxEmu.width).toBeGreaterThan(0);
    expect(s1Pic?.boxEmu.height).toBeGreaterThan(0);

    // Slide 2: clipArt placeholder with omitted local transform (<p:spPr/>)
    const s2Clip = validation.slides[1]?.objects[0];
    expect(s2Clip?.kind).toBe("image");
    expect(s2Clip?.placeholder).toEqual({ type: "clipArt", idx: 10 });
    expect(s2Clip?.hasLocalTransform).toBe(false);
    expect(s2Clip?.crop).toEqual({ left: 23333, right: 23333 });

    // Slide 3: tbl placeholder
    const s3Tbl = validation.slides[2]?.objects[0];
    expect(s3Tbl?.kind).toBe("table");
    expect(s3Tbl?.placeholder).toEqual({ type: "tbl", idx: 10 });
    expect(s3Tbl?.hasLocalTransform).toBe(true);

    // Slide 4: chart placeholder
    const s4Chart = validation.slides[3]?.objects[0];
    expect(s4Chart?.kind).toBe("chart");
    expect(s4Chart?.placeholder).toEqual({ type: "chart", idx: 10 });
    expect(s4Chart?.hasLocalTransform).toBe(true);

    // Slide 5-7: text placeholders (title, untyped obj, body)
    expect(validation.slides[4]?.objects[0]?.placeholder?.type).toBe("title");
    expect(validation.slides[5]?.objects[0]?.placeholder).toEqual({
      type: "obj",
      idx: 10,
    });
    expect(validation.slides[6]?.objects[0]?.placeholder?.type).toBe("body");

    // Slide 9: unpopulated media placeholder (<p:sp> without <p:txBody>)
    const s9Media = validation.slides[8]?.objects[0];
    expect(s9Media?.kind).toBe("shape");
    expect(s9Media?.placeholder).toEqual({ type: "media", idx: 10 });
    expect(s9Media?.hasLocalTransform).toBe(false);
    expect(s9Media?.text).toBeUndefined();

    const entries = unzipSync(pptx);
    const s1Xml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    expect(s1Xml).toMatch(/<p:ph type="pic" sz="quarter" idx="10"\/>/u);
    expect(s1Xml).toMatch(/<p:spPr\/>/u);

    const s9Xml = strFromU8(entries["ppt/slides/slide9.xml"]!);
    expect(s9Xml).toMatch(/<p:ph type="media" sz="quarter" idx="10"\/>/u);
    expect(s9Xml).toMatch(/<p:spPr\/>/u);
    expect(s9Xml).not.toMatch(/<p:txBody/u);

    // Paint non-text placeholder slides (pic, clipArt, tbl, empty dgm slot, unpopulated media) cleanly with zero diagnostics.
    const preview = paint(resolution, { pages: [1, 2, 3, 8, 9] });
    expect(preview.pages).toHaveLength(5);
    for (const page of preview.pages) {
      expect(page.diagnostics).toEqual([]);
      expect(page.png).toBeInstanceOf(Uint8Array);
    }

    // Re-importing the emitted PPTX preserves all 9 slides and placeholder bindings without errors.
    const reimported = importDocument(pptx);
    expect(reimported.report.diagnostics).toEqual([]);

    // Supplying an outline font to Slide 4 (chart) and Slides 5-7 (text) paints cleanly with zero diagnostics.
    const reimportedSlides = reimported.document.children.filter(
      (child): child is SlideContainer =>
        child instanceof SlideContainer && child.kind === "slide",
    );
    for (const slideIndex of [3, 4, 5, 6]) {
      const node = reimportedSlides[slideIndex]?.children[0];
      expect(node).toBeInstanceOf(CommonNode);
      if (node instanceof CommonNode) {
        node.style({ fontFamily: jostFont });
      }
    }
    const fontBackedPreview = paint(resolveDocument(reimported.document), {
      pages: [4, 5, 6, 7],
    });
    expect(fontBackedPreview.pages).toHaveLength(4);
    for (const page of fontBackedPreview.pages) {
      expect(page.diagnostics).toEqual([]);
      expect(page.png).toBeInstanceOf(Uint8Array);
    }
  });

  it("roundtrips ph-unpopulated-placeholders.pptx preserving empty content placeholders and painting blank previews", async () => {
    const bytes = readFixture("ph-unpopulated-placeholders.pptx");
    const imported = importDocument(bytes);
    expect(
      imported.report.diagnostics.every((d) => d.category === "unsupported"),
    ).toBe(true);
    expect(
      imported.report.diagnostics.some(
        (d) =>
          d.location.part.startsWith("ppt/slides/") &&
          d.location.xpath?.includes("/p:spTree/"),
      ),
    ).toBe(false);

    const resolution = resolveDocument(imported.document);
    const { pptx } = emit(resolution);
    const validation = await validatePptx(pptx);
    expect(validation.errors).toEqual([]);
    expect(validation.slides).toHaveLength(9);

    const entries = unzipSync(pptx);
    const expectedTypes = [
      "title",
      "obj",
      "body",
      "chart",
      "tbl",
      "dgm",
      "media",
      "clipArt",
      "pic",
    ] as const;

    for (const [index, expectedType] of expectedTypes.entries()) {
      const obj = validation.slides[index]?.objects[0];
      expect(obj?.kind).toBe("shape");
      expect(obj?.placeholder?.type).toBe(expectedType);
      expect(obj?.hasLocalTransform).toBe(false);

      const slideXml = strFromU8(entries[`ppt/slides/slide${index + 1}.xml`]!);
      expect(slideXml).toMatch(/<p:spPr\/>/u);
      if (index >= 3) {
        expect(slideXml).not.toMatch(/<p:txBody/u);
      }
    }

    const preview = paint(resolution, { pages: [4, 5, 6, 7, 8, 9] });
    expect(preview.pages).toHaveLength(6);
    for (const page of preview.pages) {
      expect(page.diagnostics).toEqual([]);
      expect(page.png).toBeInstanceOf(Uint8Array);
      // Static preview omits interactive empty placeholder prompts; only slide background rect is drawn.
      expect(page.svg).not.toMatch(/<path\b/u);
    }

    const reimported = importDocument(pptx);
    expect(reimported.report.diagnostics).toEqual([]);
  });

  it("reports unconsumed-xml when an unpopulated placeholder spPr carries unsupported fills or effects", () => {
    const bytes = readFixture("ph-unpopulated-placeholders.pptx");
    const entries = unzipSync(bytes);
    const slide4Xml = strFromU8(entries["ppt/slides/slide4.xml"]!);
    entries["ppt/slides/slide4.xml"] = strToU8(
      slide4Xml.replace(
        "<p:spPr/>",
        '<p:spPr><a:solidFill><a:srgbClr val="DC2626"/></a:solidFill><a:effectLst><a:outerShdw blurRad="50800"/></a:effectLst></p:spPr>',
      ),
    );
    const mutated = zipSync(entries);
    const imported = importDocument(mutated);
    const slide4Unconsumed = imported.report.diagnostics.filter(
      (
        d,
      ): d is Extract<
        (typeof imported.report.diagnostics)[number],
        { readonly code: "unconsumed-xml" }
      > =>
        d.code === "unconsumed-xml" &&
        d.location.part === "ppt/slides/slide4.xml",
    );
    expect(slide4Unconsumed.some((d) => d.localName === "solidFill")).toBe(
      true,
    );
    expect(slide4Unconsumed.some((d) => d.localName === "effectLst")).toBe(
      true,
    );
  });
});
