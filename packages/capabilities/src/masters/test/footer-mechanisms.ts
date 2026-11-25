import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  defineSlideLayout,
  font,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { Text } from "@pptxdk/capabilities/text-elements";
import { defineTheme } from "@pptxdk/core";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { PNG } from "pngjs";
import { createElement } from "react";

import type {
  CapabilityScenario,
  NativeRoundtripAssertContext,
} from "../../../test/acceptance-runner.ts";
import type { MechanismVerificationEntry } from "../../../test/mechanism-verification.ts";
import type { footerPlaceholderMechanism } from "../footer-placeholder.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const footerTheme = defineTheme({
  name: "Footer Brand Theme",
  fonts: { major: jostFont, minor: jostFont },
});

function countNonWhitePixelsInBand(
  pngBytes: Uint8Array,
  xMin: number,
  xMax: number,
  yMin: number,
  yMax: number,
): number {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  let count = 0;
  for (let y = yMin; y < Math.min(yMax, decoded.height); y += 1) {
    for (let x = xMin; x < Math.min(xMax, decoded.width); x += 1) {
      const idx = (y * decoded.width + x) * 4;
      const r = decoded.data[idx] ?? 255;
      const g = decoded.data[idx + 1] ?? 255;
      const b = decoded.data[idx + 2] ?? 255;
      if (r < 245 || g < 245 || b < 245) count += 1;
    }
  }
  return count;
}

const titleLayout = defineSlideLayout({
  id: "title-layout",
  name: "Title Slide",
  type: "title",
  preserve: true,
  headerFooter: { footer: false, slideNumber: false, dateTime: false },
});

const contentLayout = defineSlideLayout({
  id: "content-layout",
  name: "Title and Content",
  type: "obj",
  preserve: true,
});

const sharedFooterPlaceholders = {
  dateTime: {
    x: 36,
    y: 316,
    width: 150,
    height: 24,
    text: "2026-11-22",
    style: {
      fontFamily: jostFont,
      fontSize: 12,
      color: "#475569",
      align: "left" as const,
    },
  },
  footer: {
    x: 200,
    y: 316,
    width: 240,
    height: 24,
    text: "Acme Engineering",
    style: {
      fontFamily: jostFont,
      fontSize: 12,
      color: "#475569",
      align: "center" as const,
    },
  },
  slideNumber: {
    x: 490,
    y: 316,
    width: 110,
    height: 24,
    style: {
      fontFamily: jostFont,
      fontSize: 12,
      color: "#0F172A",
      align: "right" as const,
    },
  },
};

export const footerPlaceholderScenario: CapabilityScenario = {
  id: "mechanism-footer-placeholder",
  capability: "footer-placeholder",
  description:
    "SlideMaster footer/date/slideNumber slots with p:hf suppression on title layout, dynamic per-page numbers, and slide override.",
  visualCheckpoints: [
    "Page 1 hides date, footer, and slide number via Title layout headerFooter flags.",
    "Page 2 displays shared date, 'Acme Engineering' footer, and dynamic slide number '2'.",
    "Page 3 overrides footer to 'Quarterly Review' and displays dynamic slide number '3'.",
  ],
  nativeCheckpoints: [
    "Master holds dt/ftr/sldNum slots with xfrm; title layout carries p:hf; active slides materialize dt/ftr/sldNum placeholders cleanly.",
  ],
  emitOptions: { embedFonts: "subset" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: footerTheme },
      createElement(
        SlideMaster,
        {
          id: "brand",
          name: "Brand Master",
          footerPlaceholders: sharedFooterPlaceholders,
        },
        createElement(SlideLayout, { layout: titleLayout }),
        createElement(SlideLayout, { layout: contentLayout }),
      ),
      createElement(
        Slide,
        { id: "cover", layout: titleLayout },
        createElement(
          Text,
          {
            id: "cover-title",
            x: 48,
            y: 120,
            width: 540,
            height: 60,
            style: { fontFamily: jostFont, fontSize: 28, color: "#0F172A" },
          },
          "Cover Slide",
        ),
      ),
      createElement(
        Slide,
        { id: "page-two", layout: contentLayout },
        createElement(
          Text,
          {
            id: "p2-title",
            x: 48,
            y: 48,
            width: 540,
            height: 48,
            style: { fontFamily: jostFont, fontSize: 24, color: "#0F172A" },
          },
          "Second Slide",
        ),
      ),
      createElement(
        Slide,
        {
          id: "page-three",
          layout: contentLayout,
          footer: "Quarterly Review",
        },
        createElement(
          Text,
          {
            id: "p3-title",
            x: 48,
            y: 48,
            width: 540,
            height: 48,
            style: { fontFamily: jostFont, fontSize: 24, color: "#0F172A" },
          },
          "Third Slide",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: footerTheme }).append(
      slideMaster({
        id: "brand",
        name: "Brand Master",
        footerPlaceholders: sharedFooterPlaceholders,
      }).append(slideLayout(titleLayout), slideLayout(contentLayout)),
      slide({ id: "cover", layout: titleLayout }).append(
        text("Cover Slide", {
          id: "cover-title",
          x: 48,
          y: 120,
          width: 540,
          height: 60,
          style: { fontFamily: jostFont, fontSize: 28, color: "#0F172A" },
        }),
      ),
      slide({ id: "page-two", layout: contentLayout }).append(
        text("Second Slide", {
          id: "p2-title",
          x: 48,
          y: 48,
          width: 540,
          height: 48,
          style: { fontFamily: jostFont, fontSize: 24, color: "#0F172A" },
        }),
      ),
      slide({
        id: "page-three",
        layout: contentLayout,
        footer: "Quarterly Review",
      }).append(
        text("Third Slide", {
          id: "p3-title",
          x: 48,
          y: 48,
          width: 540,
          height: 48,
          style: { fontFamily: jostFont, fontSize: 24, color: "#0F172A" },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 3);
  },
};

export const footerPlaceholderMechanismVerifications: {
  readonly "footer-placeholder": MechanismVerificationEntry<
    typeof footerPlaceholderMechanism
  >;
} = {
  "footer-placeholder": {
    scenario: footerPlaceholderScenario,
    assertNative({ stage, validation }: NativeRoundtripAssertContext) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const master = validation.masters?.[0];
      assert.ok(master, `${stage} master missing`);
      assert.deepEqual(
        master.footerPlaceholders?.map((p) => p.kind),
        ["dt", "ftr", "sldNum"],
      );
      assert.deepEqual(master.layouts[0]?.headerFooter, {
        dateTime: false,
        footer: false,
        slideNumber: false,
      });

      assert.equal(validation.slides.length, 3);
      assert.equal(validation.slides[0]?.footerPlaceholders, undefined);

      const s2Ph = validation.slides[1]?.footerPlaceholders ?? [];
      assert.deepEqual(
        s2Ph.map((p) => p.kind),
        ["dt", "ftr", "sldNum"],
      );
      assert.equal(s2Ph.find((p) => p.kind === "dt")?.text, "2026-11-22");
      assert.equal(
        s2Ph.find((p) => p.kind === "ftr")?.text,
        "Acme Engineering",
      );
      assert.equal(
        s2Ph.find((p) => p.kind === "sldNum")?.hasSlideNumberField,
        true,
      );
      assert.equal(s2Ph.find((p) => p.kind === "sldNum")?.text, "2");

      const s3Ph = validation.slides[2]?.footerPlaceholders ?? [];
      assert.equal(
        s3Ph.find((p) => p.kind === "ftr")?.text,
        "Quarterly Review",
      );
      assert.equal(s3Ph.find((p) => p.kind === "sldNum")?.text, "3");
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 3);
      assert.equal(roundtrip.pages.length, 3);
      for (const page of [...initial.pages, ...roundtrip.pages]) {
        assert.deepEqual(page.diagnostics, []);
        assert.ok(page.png && page.png.byteLength > 0);
      }
      // Title layout suppresses dt/ftr/sldNum; footer band (y >= 310) stays blank white.
      assert.equal(
        countNonWhitePixelsInBand(initial.pages[0]!.png!, 20, 620, 312, 344),
        0,
      );
      assert.equal(
        countNonWhitePixelsInBand(roundtrip.pages[0]!.png!, 20, 620, 312, 344),
        0,
      );
      // Pages 2 and 3 draw date/footer/slideNumber glyphs in footer band, and differ visually.
      assert.ok(
        countNonWhitePixelsInBand(initial.pages[1]!.png!, 20, 620, 312, 344) >
          20,
      );
      assert.ok(
        countNonWhitePixelsInBand(roundtrip.pages[2]!.png!, 20, 620, 312, 344) >
          20,
      );
      assert.notDeepEqual(initial.pages[1]?.png, initial.pages[2]?.png);
      assert.notDeepEqual(roundtrip.pages[1]?.png, roundtrip.pages[2]?.png);
    },
  },
};
