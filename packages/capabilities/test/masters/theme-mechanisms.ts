import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  defineSlideLayout,
  font,
  shape,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import { defaultColorMap, defineTheme } from "@pptxdk/core";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../mechanism-verification.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const invertedColorMap = {
  ...defaultColorMap,
  bg1: "dk1",
  tx1: "lt1",
  bg2: "dk2",
  tx2: "lt2",
} as const;

const alphaTheme = defineTheme({
  name: "Brand Alpha",
  colors: {
    dk1: "#111827",
    lt1: "#F8FAFC",
    dk2: "#1E293B",
    lt2: "#E2E8F0",
    accent1: "#2563EB",
    accent2: "#059669",
  },
  fonts: { major: jostFont, minor: jostFont },
});

const betaEqualTheme = defineTheme({
  name: "Brand Alpha",
  colors: {
    dk1: "#111827",
    lt1: "#F8FAFC",
    dk2: "#1E293B",
    lt2: "#E2E8F0",
    accent1: "#2563EB",
    accent2: "#059669",
  },
  fonts: { major: jostFont, minor: jostFont },
});

const alphaDefaultLayout = defineSlideLayout({
  id: "alpha-default",
  name: "Alpha Default",
  type: "blank",
});

const alphaOverrideLayout = defineSlideLayout({
  id: "alpha-ovr",
  name: "Alpha Overridden",
  type: "blank",
  colorMapOverride: { ...defaultColorMap, tx1: "accent2" },
  themeOverride: { colors: { accent2: "#DC2626" } },
});

const betaLayout = defineSlideLayout({
  id: "beta-blank",
  name: "Beta Blank",
  type: "blank",
});

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

export const nativeThemesRoundtripScenario: CapabilityScenario = {
  id: "mechanism-native-theme-clrMap-override",
  capability: "masters",
  description:
    "Multi-master themes, equal-value independent theme parts, master clrMap inversion, layout clrMapOvr, and layout/slide themeOverride roundtrip.",
  visualCheckpoints: [
    "Page 1 shape resolves tx1 through inverted master clrMap (tx1 -> lt1 = #F8FAFC) while text uses accent1 (#2563EB).",
    "Page 2 shape resolves tx1 through layout clrMapOvr + layout themeOverride (tx1 -> accent2 = #DC2626) while text uses slide themeOverride accent1 (#9333EA).",
    "Page 3 uses independent theme2 accent1 (#2563EB) and default clrMap tx1 (#111827).",
  ],
  nativeCheckpoints: [
    "Two equal-value Theme instances emit distinct theme1.xml and theme2.xml, preserve clrMap, clrMapOvr, and themeOverride parts across roundtrip.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        {
          id: "master-alpha",
          name: "Master Alpha",
          theme: alphaTheme,
          colorMap: invertedColorMap,
        },
        createElement(SlideLayout, { layout: alphaDefaultLayout }),
        createElement(SlideLayout, { layout: alphaOverrideLayout }),
      ),
      createElement(
        SlideMaster,
        {
          id: "master-beta",
          name: "Master Beta",
          theme: betaEqualTheme,
        },
        createElement(SlideLayout, { layout: betaLayout }),
      ),
      createElement(
        Slide,
        { id: "s1-clr-map", layout: alphaDefaultLayout },
        createElement(Shape, {
          id: "box-tx1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "tx1",
        }),
        createElement(
          Text,
          {
            id: "text-acc1",
            x: 40,
            y: 140,
            width: 240,
            height: 40,
            style: { color: "accent1", fontFamily: "major", fontSize: 18 },
          },
          "Master clrMap Slide",
        ),
      ),
      createElement(
        Slide,
        {
          id: "s2-theme-ovr",
          layout: alphaOverrideLayout,
          themeOverride: { colors: { accent1: "#9333EA" } },
        },
        createElement(Shape, {
          id: "box-ovr-tx1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "tx1",
        }),
        createElement(
          Text,
          {
            id: "text-ovr-acc1",
            x: 40,
            y: 140,
            width: 240,
            height: 40,
            style: { color: "accent1", fontFamily: "minor", fontSize: 18 },
          },
          "ThemeOverride Slide",
        ),
      ),
      createElement(
        Slide,
        { id: "s3-beta", layout: betaLayout },
        createElement(Shape, {
          id: "box-beta-acc1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "accent1",
        }),
        createElement(
          Text,
          {
            id: "text-beta-tx1",
            x: 40,
            y: 140,
            width: 240,
            height: 40,
            style: { color: "tx1", fontFamily: "minor", fontSize: 18 },
          },
          "Independent Theme Slide",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "master-alpha",
        name: "Master Alpha",
        theme: alphaTheme,
        colorMap: invertedColorMap,
      }).append(
        slideLayout(alphaDefaultLayout),
        slideLayout(alphaOverrideLayout),
      ),
      slideMaster({
        id: "master-beta",
        name: "Master Beta",
        theme: betaEqualTheme,
      }).append(slideLayout(betaLayout)),
      slide({ id: "s1-clr-map", layout: alphaDefaultLayout }).append(
        shape({
          id: "box-tx1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "tx1",
        }),
        text("Master clrMap Slide", {
          id: "text-acc1",
          x: 40,
          y: 140,
          width: 240,
          height: 40,
          style: { color: "accent1", fontFamily: "major", fontSize: 18 },
        }),
      ),
      slide({
        id: "s2-theme-ovr",
        layout: alphaOverrideLayout,
        themeOverride: { colors: { accent1: "#9333EA" } },
      }).append(
        shape({
          id: "box-ovr-tx1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "tx1",
        }),
        text("ThemeOverride Slide", {
          id: "text-ovr-acc1",
          x: 40,
          y: 140,
          width: 240,
          height: 40,
          style: { color: "accent1", fontFamily: "minor", fontSize: 18 },
        }),
      ),
      slide({ id: "s3-beta", layout: betaLayout }).append(
        shape({
          id: "box-beta-acc1",
          x: 40,
          y: 40,
          width: 120,
          height: 80,
          fill: "accent1",
        }),
        text("Independent Theme Slide", {
          id: "text-beta-tx1",
          x: 40,
          y: 140,
          width: 240,
          height: 40,
          style: { color: "tx1", fontFamily: "minor", fontSize: 18 },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.masters?.length, 2);
    assert.equal(validation.masters?.[0]?.themePart, "ppt/theme/theme1.xml");
    assert.equal(validation.masters?.[1]?.themePart, "ppt/theme/theme2.xml");
    assert.equal(validation.slides.length, 3);
  },
};

export const themeMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  "native-theme" | "color-map" | "theme-override"
> = {
  "native-theme": {
    scenario: nativeThemesRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.masters?.length, 2, `${stage} master count`);
      assert.equal(validation.masters?.[0]?.themePart, "ppt/theme/theme1.xml");
      assert.equal(validation.masters?.[1]?.themePart, "ppt/theme/theme2.xml");
      assert.equal(
        validation.masters?.[0]?.theme?.colors.accent1,
        "#2563EB",
        `${stage} theme1 accent1`,
      );
      assert.equal(
        validation.masters?.[1]?.theme?.colors.accent1,
        "#2563EB",
        `${stage} theme2 accent1`,
      );
      const parts = unzipSync(pptx);
      assert.ok(parts["ppt/theme/theme1.xml"], `${stage} theme1.xml exists`);
      assert.ok(parts["ppt/theme/theme2.xml"], `${stage} theme2.xml exists`);
      const s3 = validation.slides[2];
      assert.equal(
        s3?.objects[1]?.text?.paragraphs[0]?.runs[0]?.color,
        "#111827",
      );
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 3);
      assert.deepEqual(initial.pages[2]?.diagnostics, []);
      assert.ok(initial.pages[2]?.svg?.includes('fill="#2563EB"'));
      const png = initial.pages[2]?.png;
      assert.ok(png instanceof Uint8Array);
      assert.equal(readPngPixelHex(png, 80, 80), "#2563EB");
      // Until Ticket 02 imports embedded fonts, unembedded imported text reports missing outline font on roundtrip paint.
      assert.equal(roundtrip.pages.length, 3);
    },
  },
  "color-map": {
    scenario: nativeThemesRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.masters?.[0]?.colorMap?.tx1, "lt1");
      assert.equal(validation.masters?.[0]?.colorMap?.bg1, "dk1");
      assert.equal(validation.masters?.[1]?.colorMap?.tx1, "dk1");
      const layoutOvr = validation.masters?.[0]?.layouts[1]?.colorMapOverride;
      assert.ok(typeof layoutOvr === "object" && layoutOvr !== null);
      assert.equal(layoutOvr.tx1, "accent2");
      const parts = unzipSync(pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:schemeClr val="tx1"\/>/u);
      assert.equal(
        validation.slides[0]?.objects[1]?.text?.paragraphs[0]?.runs[0]?.color,
        "#2563EB",
      );
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 3);
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.svg?.includes('fill="#F8FAFC"'));
      const png = initial.pages[0]?.png;
      assert.ok(png instanceof Uint8Array);
      assert.equal(readPngPixelHex(png, 80, 80), "#F8FAFC");
      assert.equal(roundtrip.pages.length, 3);
    },
  },
  "theme-override": {
    scenario: nativeThemesRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const layout2 = validation.masters?.[0]?.layouts[1];
      assert.ok(layout2?.themeOverridePart);
      assert.equal(layout2?.themeOverride?.colors?.accent2, "#DC2626");
      const slide2 = validation.slides[1];
      assert.ok(slide2?.themeOverridePart);
      assert.equal(slide2?.themeOverride?.colors?.accent1, "#9333EA");
      assert.equal(
        slide2?.objects[1]?.text?.paragraphs[0]?.runs[0]?.color,
        "#9333EA",
      );
      const parts = unzipSync(pptx);
      assert.ok(parts[layout2.themeOverridePart]);
      assert.ok(parts[slide2.themeOverridePart]);
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 3);
      assert.deepEqual(initial.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[1]?.svg?.includes('fill="#DC2626"'));
      const png = initial.pages[1]?.png;
      assert.ok(png instanceof Uint8Array);
      assert.equal(readPngPixelHex(png, 80, 80), "#DC2626");
      assert.equal(roundtrip.pages.length, 3);
    },
  },
};
