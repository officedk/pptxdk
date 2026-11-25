import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  defineSlideLayout,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import {
  createPlaceholderNode,
  type contentPlaceholderMechanism,
} from "@pptxdk/capabilities/placeholder";
import { CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../../../test/mechanism-verification.ts";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

const unpopulatedLayout = defineSlideLayout({
  id: "unpopulated-obj-layout",
  name: "Unpopulated Object Layout",
  type: "obj",
  preserve: true,
  placeholders: {
    content: {
      type: "obj",
      idx: 1,
      x: 48,
      y: 96,
      width: 544,
      height: 220,
    },
  },
});

/**
 * Dual-author scenario exercising unpopulated content placeholders (`placeholder` capability)
 * with both inherited layout geometry (`<p:spPr/>`) and explicit local transform overrides.
 */
export const contentPlaceholderRoundtripScenario: CapabilityScenario = {
  id: "mechanism-content-placeholder-roundtrip",
  capability: "placeholder",
  description:
    "Dual-author unpopulated content placeholder scenario verifying omitted <p:txBody>, inherited <p:spPr/> vs local transform, and static preview blanking.",
  visualCheckpoints: [
    "Pages 1 and 2 render clean blank white canvases with zero diagnostics (interactive empty placeholder prompts are omitted in static preview).",
  ],
  nativeCheckpoints: [
    "Slide 1 emits <p:spPr/> without <p:txBody>; Slide 2 retains local transform without <p:txBody>; importedDocument keeps placeholder nodes.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "ph-master", name: "Placeholder Master", preserve: true },
        createElement(SlideLayout, { layout: unpopulatedLayout }),
      ),
      createElement(
        Slide,
        { id: "unpopulated-inherited", layout: unpopulatedLayout },
        createElement("placeholder" as never, {
          id: "empty-obj",
          placeholder: "content",
        }),
      ),
      createElement(
        Slide,
        { id: "unpopulated-local-box", layout: unpopulatedLayout },
        createElement("placeholder" as never, {
          id: "local-obj",
          placeholder: "content",
          x: 96,
          y: 120,
          width: 320,
          height: 140,
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "ph-master",
        name: "Placeholder Master",
        preserve: true,
      }).append(slideLayout(unpopulatedLayout)),
      slide({ id: "unpopulated-inherited", layout: unpopulatedLayout }).append(
        createPlaceholderNode({ id: "empty-obj", placeholder: "content" }),
      ),
      slide({ id: "unpopulated-local-box", layout: unpopulatedLayout }).append(
        createPlaceholderNode({
          id: "local-obj",
          placeholder: "content",
          x: 96,
          y: 120,
          width: 320,
          height: 140,
        }),
      ),
    );
  },
  assert({ validation, preview }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 2);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.equal(page.svg?.includes("<path "), false);
    }
  },
};

export const placeholderMechanismVerifications = {
  "content-placeholder": {
    scenario: contentPlaceholderRoundtripScenario,
    async assertNative({ stage, pptx, validation, importedDocument }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 2);

      const s1Ph = validation.slides[0]?.objects[0];
      assert.equal(s1Ph?.kind, "shape");
      assert.deepEqual(s1Ph?.placeholder, { type: "obj", idx: 1 });
      assert.equal(s1Ph?.hasLocalTransform, false);
      assert.equal(s1Ph?.text, undefined);

      const s2Ph = validation.slides[1]?.objects[0];
      assert.equal(s2Ph?.kind, "shape");
      assert.deepEqual(s2Ph?.placeholder, { type: "obj", idx: 1 });
      assert.equal(s2Ph?.hasLocalTransform, true);
      assert.deepEqual(s2Ph?.boxEmu, {
        x: 96 * 9525,
        y: 120 * 9525,
        width: 320 * 9525,
        height: 140 * 9525,
      });
      assert.equal(s2Ph?.text, undefined);

      const scenarioEntries = unzipSync(pptx);
      const s1Xml = strFromU8(scenarioEntries["ppt/slides/slide1.xml"]!);
      assert.match(s1Xml, /<p:spPr\/>/u);
      assert.doesNotMatch(s1Xml, /<p:txBody/u);
      const s2Xml = strFromU8(scenarioEntries["ppt/slides/slide2.xml"]!);
      assert.doesNotMatch(s2Xml, /<p:txBody/u);

      if (stage === "roundtrip" && importedDocument !== undefined) {
        const slides = importedDocument.children.filter(
          (child): child is CommonNode =>
            child instanceof CommonNode && child.kind === "slide",
        );
        const importedPh = slides[0]?.children[0];
        assert.ok(importedPh instanceof CommonNode);
        assert.equal(importedPh.kind, "placeholder");
        assert.equal(importedPh.props["placeholder"], "content");
      }

      if (stage !== "initial") return;

      // Also verify pure-file roundtrip on the shared python-pptx unpopulated placeholders reference fixture.
      const fixtureBytes = new Uint8Array(
        readFileSync(
          new URL(
            "../../../test/fixtures/ph-unpopulated-placeholders.pptx",
            import.meta.url,
          ),
        ),
      );
      const imported = importDocument(fixtureBytes);
      fixtureBytes.fill(0);
      assert.ok(
        imported.report.diagnostics.every((d) => d.category === "unsupported"),
      );
      assert.equal(
        imported.report.diagnostics.some(
          (d) =>
            d.location.part.startsWith("ppt/slides/") &&
            d.location.xpath?.includes("/p:spTree/"),
        ),
        false,
      );
      const resolution = resolveDocument(imported.document);
      const { pptx: fixturePptx } = emit(resolution);
      const fixtureValidation = await validatePptx(fixturePptx);
      assert.deepEqual(fixtureValidation.errors, []);
      assert.equal(fixtureValidation.slides.length, 9);
      const reimported = importDocument(fixturePptx);
      assert.deepEqual(reimported.report.diagnostics, []);

      const entries = unzipSync(fixturePptx);
      for (let slideNum = 4; slideNum <= 9; slideNum += 1) {
        const slideXml = strFromU8(entries[`ppt/slides/slide${slideNum}.xml`]!);
        assert.match(slideXml, /<p:spPr\/>/u);
        assert.doesNotMatch(slideXml, /<p:txBody/u);
        const obj = fixtureValidation.slides[slideNum - 1]?.objects[0];
        assert.equal(obj?.kind, "shape");
        assert.equal(obj?.hasLocalTransform, false);
        assert.equal(obj?.text, undefined);
      }

      const preview = paint(resolution, { pages: [4, 5, 6, 7, 8, 9] });
      assert.equal(preview.pages.length, 6);
      for (const page of preview.pages) {
        assert.deepEqual(page.diagnostics, []);
        assert.ok(page.png instanceof Uint8Array);
        assert.equal(page.svg?.includes("<path "), false);
      }
    },
    assertPaint({ initial, roundtrip }) {
      for (const preview of [initial, roundtrip]) {
        assert.equal(preview.pages.length, 2);
        for (const page of preview.pages) {
          assert.deepEqual(page.diagnostics, []);
          assert.ok(page.png instanceof Uint8Array);
          assert.equal(page.svg?.includes("<path "), false);
        }
      }
    },
  } satisfies MechanismVerificationEntry<typeof contentPlaceholderMechanism>,
} satisfies ExhaustiveMechanismVerificationMap<
  typeof contentPlaceholderMechanism
>;
