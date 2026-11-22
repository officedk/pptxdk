import assert from "node:assert/strict";

import { deck, image, slide } from "@pptxdk/capabilities/document";
import { Image } from "@pptxdk/capabilities/image-elements";
import { path, preset } from "@pptxdk/capabilities/shape";
import { Deck, Slide } from "@pptxdk/react";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { MechanismVerificationEntry } from "../mechanism-verification.ts";

function createTestPng(): Uint8Array {
  const png = new PNG({ width: 40, height: 20 });
  for (let y = 0; y < 20; y++) {
    for (let x = 0; x < 40; x++) {
      png.data.set(
        x < 20 ? [220, 40, 40, 255] : [30, 110, 220, 255],
        (y * 40 + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(png));
}

const samplePng = createTestPng();
const diamondCustomMask = path("M5 0 L10 5 L5 10 L0 5 Z", {
  viewBox: [10, 10],
});

export const imageRoundtripScenario: CapabilityScenario = {
  id: "mechanism-image-roundtrip",
  capability: "image",
  description:
    "Dual-author image mechanism scenario verifying embedded PNG, crop, opacity, preset & custom clipping geometries, and external/slide/home hyperlinks through full import roundtrip.",
  visualCheckpoints: [
    "Page 1 renders cropped/masked/semi-transparent/custom-clipped images with zero diagnostics.",
    "Page 2 renders target image with home action.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain <p:pic>, <a:srcRect>, <a:alphaModFix>, <a:prstGeom>/<a:custGeom>, and <a:hlinkClick> relationships.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "images-main" },
        createElement(Image, {
          id: "cropped-ellipse",
          name: "CroppedEllipse",
          alt: "Red and blue split sample",
          src: samplePng,
          x: 24,
          y: 24,
          width: 160,
          height: 120,
          geometry: "ellipse",
          crop: { left: 0.25, top: 0.1 },
          link: { url: "https://example.com/docs" },
          style: {
            opacity: 0.75,
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
          },
        }),
        createElement(Image, {
          id: "custom-mask-jump",
          name: "CustomMaskJump",
          src: samplePng,
          x: 220,
          y: 24,
          width: 140,
          height: 120,
          geometry: diamondCustomMask,
          link: { id: "images-target" },
        }),
        createElement(Image, {
          id: "roundrect-home",
          name: "RoundRectHome",
          src: samplePng,
          x: 400,
          y: 24,
          width: 140,
          height: 120,
          geometry: preset("roundRect", { adj: 20000 }),
          link: { action: "home" },
        }),
      ),
      createElement(
        Slide,
        { id: "images-target" },
        createElement(Image, {
          id: "target-img",
          name: "TargetImg",
          src: samplePng,
          x: 40,
          y: 40,
          width: 200,
          height: 100,
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "images-main" }).append(
        image({
          id: "cropped-ellipse",
          name: "CroppedEllipse",
          alt: "Red and blue split sample",
          src: samplePng,
          x: 24,
          y: 24,
          width: 160,
          height: 120,
          geometry: "ellipse",
          crop: { left: 0.25, top: 0.1 },
          link: { url: "https://example.com/docs" },
          style: {
            opacity: 0.75,
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
          },
        }),
        image({
          id: "custom-mask-jump",
          name: "CustomMaskJump",
          src: samplePng,
          x: 220,
          y: 24,
          width: 140,
          height: 120,
          geometry: diamondCustomMask,
          link: { id: "images-target" },
        }),
        image({
          id: "roundrect-home",
          name: "RoundRectHome",
          src: samplePng,
          x: 400,
          y: 24,
          width: 140,
          height: 120,
          geometry: preset("roundRect", { adj: 20000 }),
          link: { action: "home" },
        }),
      ),
      slide({ id: "images-target" }).append(
        image({
          id: "target-img",
          name: "TargetImg",
          src: samplePng,
          x: 40,
          y: 40,
          width: 200,
          height: 100,
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 2);
    assert.equal(validation.slides[0]?.objects.length, 3);
  },
};

export const imageMechanismVerifications = {
  image: {
    scenario: imageRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 2);
      const s1 = validation.slides[0]!;
      assert.equal(s1.objects.length, 3);

      const ellipse = s1.objects[0]!;
      assert.equal(ellipse.kind, "image");
      assert.equal(ellipse.name, "CroppedEllipse");
      assert.equal(ellipse.geometry, "ellipse");
      assert.deepEqual(ellipse.crop, { left: 25000, top: 10000 });
      assert.deepEqual(ellipse.link, { url: "https://example.com/docs" });
      assert.equal(ellipse.image?.contentType, "image/png");

      const customMask = s1.objects[1]!;
      assert.equal(customMask.kind, "image");
      assert.equal(customMask.geometry, "custom");
      assert.deepEqual(customMask.link, { slideId: "images-target" });

      const homeImg = s1.objects[2]!;
      assert.equal(homeImg.kind, "image");
      assert.equal(homeImg.geometry, "roundRect");
      assert.deepEqual(homeImg.link, { action: "home" });
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 2);
      assert.equal(roundtrip.pages.length, 2);
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.ok(roundtrip.pages[0]?.png instanceof Uint8Array);
      assert.ok(roundtrip.pages[1]?.png instanceof Uint8Array);
    },
  } satisfies MechanismVerificationEntry<
    Extract<
      import("../mechanism-verification.ts").ProductionMechanism,
      { readonly key: "image" }
    >
  >,
} as const;
