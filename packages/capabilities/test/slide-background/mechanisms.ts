import assert from "node:assert/strict";

import {
  deck,
  defineSlideLayout,
  linearGradient,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { defineTheme } from "@pptxdk/core";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../mechanism-verification.ts";

function solidPng(r: number, g: number, b: number): Uint8Array {
  const image = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < image.data.length; i += 4) {
    image.data[i] = r;
    image.data[i + 1] = g;
    image.data[i + 2] = b;
    image.data[i + 3] = 255;
  }
  return new Uint8Array(PNG.sync.write(image));
}

function readPngPixelHex(pngBytes: Uint8Array, x: number, y: number): string {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  const r = decoded.data[offset] ?? 0;
  const g = decoded.data[offset + 1] ?? 0;
  const b = decoded.data[offset + 2] ?? 0;
  return `#${[r, g, b]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

const masterImageBytes = solidPng(0x1e, 0x40, 0xaf);

const refTheme = defineTheme({
  name: "Background Ref Theme",
  colors: {
    dk1: "#111827",
    lt1: "#F8FAFC",
    dk2: "#1E293B",
    lt2: "#E2E8F0",
    accent1: "#2563EB",
    accent2: "#059669",
  },
  formatScheme: {
    name: "Custom",
    bgFillStyles: [
      {
        kind: "solid",
        color: { color: "phClr" },
      },
      {
        kind: "gradient",
        linear: { angle: 0, scaled: true },
        stops: [
          { position: 0, color: { color: "phClr", tint: 0.5 } },
          { position: 1, color: { color: "phClr" } },
        ],
      },
      {
        kind: "solid",
        color: { color: "phClr", shade: 0.5 },
      },
    ],
  },
});

const inheritMasterBgLayout = defineSlideLayout({
  id: "inherit-master-bg",
  name: "Inherit Master Background",
  type: "blank",
});

const solidLayoutBg = defineSlideLayout({
  id: "layout-solid-bg",
  name: "Layout Solid Background",
  type: "blank",
  background: "#0F766E",
});

export const slideBackgroundMechanismsScenario: CapabilityScenario = {
  id: "mechanism-slide-background-variants",
  capability: "slide-background",
  description:
    "Master/layout/slide backgrounds covering solid fill, linear gradient, master-owned image resource registration, and p:bgRef format-scheme evaluation.",
  visualCheckpoints: [
    "Page 1 inherits master image background (#1E40AF) without duplicating p:bg on slide1.",
    "Page 2 inherits layout solid background (#0F766E) overriding master image.",
    "Page 3 renders slide-level linear gradient overriding layout solid background.",
    "Page 4 resolves slide-level p:bgRef idx=1002 gradient via theme.formatScheme.bgFillStyles and colorMap.",
    "Page 5 resolves p:bgRef idx=1003 shaded accent2 (#059669 shaded 0.5 -> #034B35).",
  ],
  nativeCheckpoints: [
    "slideMaster1 owns its background image relationship in slideMaster1.xml.rels; slideLayout2 emits p:bg solidFill; slide3..5 emit gradFill/bgRef cleanly.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        SlideMaster,
        {
          id: "bg-master",
          name: "BG Master",
          theme: refTheme,
          background: { image: masterImageBytes },
        },
        createElement(SlideLayout, { layout: inheritMasterBgLayout }),
        createElement(SlideLayout, { layout: solidLayoutBg }),
      ),
      createElement(Slide, {
        id: "s1-inherit-image",
        layout: inheritMasterBgLayout,
      }),
      createElement(Slide, {
        id: "s2-inherit-layout-solid",
        layout: solidLayoutBg,
      }),
      createElement(Slide, {
        id: "s3-slide-gradient",
        layout: solidLayoutBg,
        background: linearGradient(90, "#0F172A", "#38BDF8"),
      }),
      createElement(Slide, {
        id: "s4-bg-ref-grad",
        layout: inheritMasterBgLayout,
        background: { ref: 1002, color: "accent1" },
      }),
      createElement(Slide, {
        id: "s5-bg-ref-shade",
        layout: inheritMasterBgLayout,
        background: { ref: 1003, color: "accent2" },
      }),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slideMaster({
        id: "bg-master",
        name: "BG Master",
        theme: refTheme,
        background: { image: masterImageBytes },
      }).append(slideLayout(inheritMasterBgLayout), slideLayout(solidLayoutBg)),
      slide({ id: "s1-inherit-image", layout: inheritMasterBgLayout }),
      slide({ id: "s2-inherit-layout-solid", layout: solidLayoutBg }),
      slide({
        id: "s3-slide-gradient",
        layout: solidLayoutBg,
        background: linearGradient(90, "#0F172A", "#38BDF8"),
      }),
      slide({
        id: "s4-bg-ref-grad",
        layout: inheritMasterBgLayout,
        background: { ref: 1002, color: "accent1" },
      }),
      slide({
        id: "s5-bg-ref-shade",
        layout: inheritMasterBgLayout,
        background: { ref: 1003, color: "accent2" },
      }),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.masters?.[0]?.background?.type, "image");
    assert.equal(
      validation.masters?.[0]?.layouts[1]?.background?.type,
      "solid",
    );
    assert.equal(validation.slides[0]?.background, undefined);
    assert.equal(validation.slides[1]?.background, undefined);
    assert.equal(validation.slides[2]?.background?.type, "gradient");
    assert.equal(validation.slides[3]?.background?.type, "ref");
    assert.equal(validation.slides[4]?.background?.type, "ref");
  },
};

export const slideBackgroundMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  | "slide-background-solid"
  | "slide-background-gradient"
  | "slide-background-image"
  | "slide-background-ref"
> = {
  "slide-background-solid": {
    scenario: slideBackgroundMechanismsScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const layout2Bg = validation.masters?.[0]?.layouts[1]?.background;
      assert.equal(layout2Bg?.type, "solid");
      assert.equal(layout2Bg?.color, "#0F766E");
      assert.equal(validation.slides[1]?.background, undefined);
    },
    assertPaint({ initial, roundtrip }) {
      for (const result of [initial, roundtrip]) {
        const page = result.pages[1]!;
        assert.deepEqual(page.diagnostics, []);
        assert.ok(page.png instanceof Uint8Array);
        assert.equal(readPngPixelHex(page.png, 100, 100), "#0F766E");
      }
    },
  },
  "slide-background-gradient": {
    scenario: slideBackgroundMechanismsScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const bg = validation.slides[2]?.background;
      assert.equal(bg?.type, "gradient");
      assert.equal(bg?.kind, "linear");
    },
    assertPaint({ initial, roundtrip }) {
      for (const result of [initial, roundtrip]) {
        const page = result.pages[2]!;
        assert.deepEqual(page.diagnostics, []);
        assert.ok(page.png instanceof Uint8Array);
        assert.ok(page.svg?.includes("<linearGradient"));
        const left = PNG.sync.read(Buffer.from(page.png)).data[
          (135 * 480 + 2) * 4 + 2
        ]!;
        const right = PNG.sync.read(Buffer.from(page.png)).data[
          (135 * 480 + 477) * 4 + 2
        ]!;
        assert.ok(left < 100 && right > 200);
      }
    },
  },
  "slide-background-image": {
    scenario: slideBackgroundMechanismsScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const masterBg = validation.masters?.[0]?.background;
      assert.equal(masterBg?.type, "image");
      const files = unzipSync(pptx);
      const masterRels = strFromU8(
        files["ppt/slideMasters/_rels/slideMaster1.xml.rels"]!,
      );
      assert.match(masterRels, /relationships\/image/u);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);
      assert.equal(slide1Xml.includes("<p:bg"), false);
    },
    assertPaint({ initial, roundtrip }) {
      for (const result of [initial, roundtrip]) {
        const page = result.pages[0]!;
        assert.deepEqual(page.diagnostics, []);
        assert.ok(page.png instanceof Uint8Array);
        assert.equal(readPngPixelHex(page.png, 100, 100), "#1E40AF");
      }
    },
  },
  "slide-background-ref": {
    scenario: slideBackgroundMechanismsScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const bg4 = validation.slides[3]?.background;
      assert.equal(bg4?.type, "ref");
      assert.equal(bg4?.index, 1002);
      const bg5 = validation.slides[4]?.background;
      assert.equal(bg5?.type, "ref");
      assert.equal(bg5?.index, 1003);
    },
    assertPaint({ initial, roundtrip }) {
      for (const result of [initial, roundtrip]) {
        assert.deepEqual(result.pages[3]?.diagnostics, []);
        assert.ok(result.pages[3]?.svg?.includes("<linearGradient"));
        const page5 = result.pages[4]!;
        assert.deepEqual(page5.diagnostics, []);
        assert.ok(page5.png instanceof Uint8Array);
        assert.equal(readPngPixelHex(page5.png, 100, 100), "#034B35");
      }
    },
  },
};
