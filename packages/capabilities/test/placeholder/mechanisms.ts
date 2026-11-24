import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";

import type { contentPlaceholderMechanism } from "../../src/placeholder/placeholder.ts";
import { masterSharedCanvasScenario } from "../masters/shared-canvas-mechanisms.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../mechanism-verification.ts";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

export const placeholderMechanismVerifications = {
  "content-placeholder": {
    // Share the master-shared-canvas scenario for cached build execution, and directly verify
    // unpopulated content placeholder roundtrip from the python-pptx reference fixture.
    scenario: masterSharedCanvasScenario,
    async assertNative({ stage }) {
      if (stage !== "initial") return;
      const fixtureBytes = new Uint8Array(
        readFileSync(
          new URL(
            "../fixtures/ph-unpopulated-placeholders.pptx",
            import.meta.url,
          ),
        ),
      );
      const imported = importDocument(fixtureBytes);
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
      const { pptx } = emit(resolution);
      const validation = await validatePptx(pptx);
      assert.deepEqual(validation.errors, []);
      assert.equal(validation.slides.length, 9);
      const reimported = importDocument(pptx);
      assert.deepEqual(reimported.report.diagnostics, []);

      const entries = unzipSync(pptx);
      for (let slideNum = 4; slideNum <= 9; slideNum += 1) {
        const slideXml = strFromU8(entries[`ppt/slides/slide${slideNum}.xml`]!);
        assert.match(slideXml, /<p:spPr\/>/u);
        assert.doesNotMatch(slideXml, /<p:txBody/u);
        const obj = validation.slides[slideNum - 1]?.objects[0];
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
    assertPaint() {},
  } satisfies MechanismVerificationEntry<typeof contentPlaceholderMechanism>,
} satisfies ExhaustiveMechanismVerificationMap<
  typeof contentPlaceholderMechanism
>;
