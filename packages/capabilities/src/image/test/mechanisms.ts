import assert from "node:assert/strict";

import {
  deck,
  defineSlideLayout,
  image,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { Image } from "@pptxdk/capabilities/image-elements";
import { path, preset } from "@pptxdk/capabilities/shape";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type { MechanismVerificationEntry } from "../../../test/mechanism-verification.ts";

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
          geometry: "ellipse",
          crop: { left: 0.25, top: 0.1 },
          link: { url: "https://example.com/docs" },
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 160,
            height: 120,
            opacity: 0.75,
            border: {
              color: "#1E293B",
              width: 2,
              style: "dashed",
            },
          },
        }),
        createElement(Image, {
          id: "custom-mask-jump",
          name: "CustomMaskJump",
          src: samplePng,
          style: {
            position: "absolute",
            left: 220,
            top: 24,
            width: 140,
            height: 120,
          },
          geometry: diamondCustomMask,
          link: { id: "images-target" },
        }),
        createElement(Image, {
          id: "roundrect-home",
          name: "RoundRectHome",
          src: samplePng,
          style: {
            position: "absolute",
            left: 400,
            top: 24,
            width: 140,
            height: 120,
          },
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
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 200,
            height: 100,
          },
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
          geometry: "ellipse",
          crop: { left: 0.25, top: 0.1 },
          link: { url: "https://example.com/docs" },
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 160,
            height: 120,
            opacity: 0.75,
            border: {
              color: "#1E293B",
              width: 2,
              style: "dashed",
            },
          },
        }),
        image({
          id: "custom-mask-jump",
          name: "CustomMaskJump",
          src: samplePng,
          style: {
            position: "absolute",
            left: 220,
            top: 24,
            width: 140,
            height: 120,
          },
          geometry: diamondCustomMask,
          link: { id: "images-target" },
        }),
        image({
          id: "roundrect-home",
          name: "RoundRectHome",
          src: samplePng,
          style: {
            position: "absolute",
            left: 400,
            top: 24,
            width: 140,
            height: 120,
          },
          geometry: preset("roundRect", { adj: 20000 }),
          link: { action: "home" },
        }),
      ),
      slide({ id: "images-target" }).append(
        image({
          id: "target-img",
          name: "TargetImg",
          src: samplePng,
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 200,
            height: 100,
          },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 2);
    assert.equal(validation.slides[0]?.objects.length, 3);
  },
};

const pictureHolderLayout = defineSlideLayout({
  id: "picture-holder-layout",
  name: "Picture Holder",
  placeholders: {
    photo: {
      type: "pic",
      explicitType: true,
      idx: 10,
      size: "quarter",
      style: {
        position: "absolute",
        left: 48,
        top: 48,
        width: 240,
        height: 120,
      },
      name: "Photo Slot",
    },
  },
});

export const picturePlaceholderRoundtripScenario: CapabilityScenario = {
  id: "mechanism-picture-placeholder-roundtrip",
  capability: "image",
  description:
    "Dual-author picture-placeholder mechanism scenario verifying inherited bounds without local <a:xfrm>, crop preservation, and roundtrip.",
  visualCheckpoints: [
    "Page 1 paints image inside inherited layout bounds (48, 48, 240x120) with zero diagnostics.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain <p:nvPr><p:ph type='pic' sz='quarter' idx='10'/></p:nvPr>, empty <p:spPr/> without <a:xfrm>, and inherited boxEmu.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "pic-master", name: "Pic Master" },
        createElement(SlideLayout, { layout: pictureHolderLayout }),
      ),
      createElement(Slide, {
        id: "pic-ph-slide",
        layout: pictureHolderLayout,
        placeholders: {
          photo: createElement(Image, {
            id: "hero-photo",
            src: samplePng,
            crop: { left: 0.1, right: 0.1 },
          }),
        },
      }),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "pic-master", name: "Pic Master" }).append(
        slideLayout(pictureHolderLayout),
      ),
      slide({
        id: "pic-ph-slide",
        layout: pictureHolderLayout,
        placeholders: {
          photo: image({
            id: "hero-photo",
            src: samplePng,
            crop: { left: 0.1, right: 0.1 },
          }),
        },
      }),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 1);
    assert.equal(validation.slides[0]?.objects.length, 1);
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
      assert.deepEqual(initial.pages[0]?.errors, []);
      assert.deepEqual(roundtrip.pages[0]?.errors, []);
      assert.ok(roundtrip.pages[0]?.png instanceof Uint8Array);
      assert.ok(roundtrip.pages[1]?.png instanceof Uint8Array);
    },
  } satisfies MechanismVerificationEntry<
    Extract<
      import("../../../test/mechanism-verification.ts").ProductionMechanism,
      { readonly key: "image" }
    >
  >,
  "picture-placeholder": {
    scenario: picturePlaceholderRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 1);
      const pic = validation.slides[0]?.objects[0];
      assert.equal(pic?.kind, "image", `${stage} pic kind`);
      assert.deepEqual(pic?.placeholder, { type: "pic", idx: 10 });
      assert.equal(pic?.hasLocalTransform, false);
      assert.deepEqual(pic?.boxEmu, {
        x: 48 * 9525,
        y: 48 * 9525,
        width: 240 * 9525,
        height: 120 * 9525,
      });
      assert.deepEqual(pic?.crop, { left: 10000, right: 10000 });
      const slideXml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slideXml, /<p:ph type="pic" sz="quarter" idx="10"\/>/u);
      assert.match(slideXml, /<p:spPr\/>/u);
    },
    assertPaint({ initial, roundtrip }) {
      for (const preview of [initial, roundtrip]) {
        assert.equal(preview.pages.length, 1);
        assert.deepEqual(preview.pages[0]?.errors, []);
        assert.ok(preview.pages[0]?.png instanceof Uint8Array);
      }
    },
  } satisfies MechanismVerificationEntry<
    Extract<
      import("../../../test/mechanism-verification.ts").ProductionMechanism,
      { readonly key: "picture-placeholder" }
    >
  >,
} as const;
