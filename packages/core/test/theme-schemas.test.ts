import assert from "node:assert/strict";

import { CapabilityError, deckPropsSchema } from "@pptxdk/core";
import { createDeckNode } from "@pptxdk/core/document";
import {
  backgroundValueSchema,
  borderObjectSchema,
  borderValueSchema,
  boxShadowValueSchema,
  colorMapOverrideSchema,
  colorMapSchema,
  colorValueSchema,
  createThemeContext,
  defaultColorMap,
  defaultTheme,
  defineTheme,
  fontFaceSchema,
  fontFaceStyleSchema,
  fontOptionsSchema,
  fontReferenceSchema,
  fontSchema,
  fontSizeSchema,
  formatSchemeSchema,
  gradientStopSchema,
  isTheme,
  isThemeFontSlot,
  linearGradientSchema,
  nativeColorSlotSchema,
  radialGradientSchema,
  schemeColorNameSchema,
  schemeColorSlotSchema,
  shadowSchema,
  textAlignmentSchema,
  textStyleSchema,
  themeColorSchema,
  themeFontSlotSchema,
  themeInputSchema,
  themeOverrideSchema,
  themeSchema,
  themeShadeSchema,
  validateTextStyle,
  visualStyleSchema,
  type BackgroundValue,
  type BorderObject,
  type BorderValue,
  type BoxShadowValue,
  type ColorMap,
  type ColorMapOverride,
  type ColorValue,
  type Font,
  type FontFace,
  type FontOptions,
  type FontReference,
  type FontSize,
  type FormatScheme,
  type HexColor,
  type LinearGradient,
  type NativeColorSlot,
  type RadialGradient,
  type Shadow,
  type TextAlignment,
  type TextStyle,
  type Theme,
  type ThemeColor,
  type ThemeFontSlot,
  type ThemeInput,
  type ThemeOverride,
  type ThemeShade,
  type VisualStyle,
} from "@pptxdk/core/theme";
import { describe, it } from "vitest";

describe("Slice 2: theme, color, font, and textStyle Zod schemas", () => {
  it("exports Zod schemas for color, font, visual, and text style primitives and preserves units", () => {
    const faceStyle = fontFaceStyleSchema.parse("italic");
    assert.equal(faceStyle, "italic");

    const face: FontFace = fontFaceSchema.parse({
      weight: 700,
      style: "normal",
      data: new Uint8Array([1, 2, 3]),
    });
    assert.equal(face.weight, 700);
    assert(face.data instanceof Uint8Array);

    const fontOpts: FontOptions = fontOptionsSchema.parse({
      substitute: true,
      embeddable: false,
    });
    assert.equal(fontOpts.substitute, true);

    const fontValue: Font = fontSchema.parse({
      family: "Jost",
      faces: [face],
      substitute: true,
    });
    const fontRef: FontReference = fontReferenceSchema.parse(fontValue);
    const fontSlot: ThemeFontSlot = themeFontSlotSchema.parse("minor");
    assert.deepEqual(fontRef, fontValue);
    assert.equal(fontRef.faces?.[0]?.data, face.data);
    assert.equal(fontSlot, "minor");
    assert.equal(isThemeFontSlot("major"), true);
    assert.equal(isThemeFontSlot("body"), false);
    const hexColor: HexColor = "#112233";
    assert.equal(themeColorSchema.parse(hexColor), "#112233");

    const sizes: readonly FontSize[] = ["32pt", 32, "32px"];
    for (const size of sizes) {
      assert.equal(fontSizeSchema.parse(size), size);
    }

    const align: TextAlignment = textAlignmentSchema.parse("justify");
    assert.equal(align, "justify");

    const slot: NativeColorSlot = nativeColorSlotSchema.parse("accent1");
    const shade: ThemeShade = themeShadeSchema.parse({
      slot,
      lumMod: 0.75,
      lumOff: 0.25,
    });
    const themeColor: ThemeColor = themeColorSchema.parse(shade);
    const colorValue: ColorValue = colorValueSchema.parse({
      color: themeColor,
      alpha: 0.8,
    });
    assert.equal((colorValue as { readonly alpha?: number }).alpha, 0.8);

    const stop = gradientStopSchema.parse({ color: "#112233", position: 0.5 });
    const linear: LinearGradient = linearGradientSchema.parse({
      type: "linear-gradient",
      angle: 90,
      stops: [stop],
    });
    const radial: RadialGradient = radialGradientSchema.parse({
      type: "radial-gradient",
      stops: [stop],
    });
    const shadow: Shadow = shadowSchema.parse({
      color: "#000000",
      blur: 4,
      distance: 2,
      angle: 45,
    });
    const bgSolid: BackgroundValue = backgroundValueSchema.parse("#FFFFFF");
    const bgNone: BackgroundValue = backgroundValueSchema.parse("none");
    const bgGrad: BackgroundValue = backgroundValueSchema.parse(linear);
    const borderObj: BorderObject = borderObjectSchema.parse({
      color: "#000000",
      width: 1,
      style: "dashed",
    });
    const borderVal: BorderValue = borderValueSchema.parse(borderObj);
    const borderNone: BorderValue = borderValueSchema.parse("none");
    const shadowNone: BoxShadowValue = boxShadowValueSchema.parse("none");
    const visual: VisualStyle = visualStyleSchema.parse({
      background: bgGrad,
      border: borderVal,
      borderRadius: 6,
      opacity: 0.9,
      boxShadow: shadow,
      softEdge: 2,
      rotate: 15,
      flipH: true,
      flipV: false,
    });
    const visualClear: VisualStyle = visualStyleSchema.parse({
      background: bgNone,
      border: borderNone,
      boxShadow: shadowNone,
      flipH: false,
      flipV: true,
    });
    assert.equal(bgSolid, "#FFFFFF");
    assert.equal(radial.type, "radial-gradient");
    assert.equal(visual.flipH, true);
    assert.equal(visual.flipV, false);
    assert.deepEqual(visual.border, borderObj);
    assert.equal(visualClear.background, "none");
    assert.equal(visualClear.border, "none");
    assert.equal(visualClear.boxShadow, "none");

    for (const fontWeight of [
      1,
      150,
      250,
      350,
      450,
      650,
      1000,
      "normal",
      "bold",
    ] as const) {
      const validated = validateTextStyle({
        fontFamily: {
          family: "Custom",
          faces: [
            { weight: typeof fontWeight === "number" ? fontWeight : 400 },
          ],
        },
        fontWeight,
      });
      assert.equal(validated.fontWeight, fontWeight);
    }

    // Preserve native units ("32pt" vs 32 vs "32px") through schema & validateTextStyle
    for (const fontSize of ["32pt", 32, "32px"] as const) {
      const style: TextStyle = textStyleSchema.parse({
        fontFamily: fontValue,
        fontSize,
        fontWeight: 600,
        fontStyle: "italic",
        color: colorValue,
        lineHeight: "36pt",
        verticalAlign: "middle",
        align: "center",
        spacingAfter: 8,
        letterSpacing: 1.5,
        textTransform: "uppercase",
        textDecoration: "underline",
        textShadow: shadow,
        textWrap: "nowrap",
        lang: "en-US",
      });
      assert.equal(style.fontSize, fontSize);
      assert.equal(validateTextStyle(style).fontSize, fontSize);
    }
  });

  it("preserves exact CapabilityError codes in validateTextStyle", () => {
    const cases: readonly [Record<string, unknown>, string][] = [
      [{ fontFamily: "" }, "invalid-font-family"],
      [{ fontFamily: { family: "  " } }, "invalid-font-family"],
      [
        { fontFamily: { family: "Bad", faces: [{ weight: 0 }] } },
        "invalid-font-face",
      ],
      [
        { fontFamily: { family: "Bad", faces: [{ weight: 1001 }] } },
        "invalid-font-face",
      ],
      [
        { fontFamily: { family: "Bad", faces: [{ weight: 400.5 }] } },
        "invalid-font-face",
      ],
      [{ fontSize: "20em" }, "invalid-text-length"],
      [{ fontSize: 0 }, "invalid-font-size"],
      [{ fontSize: -4 }, "invalid-font-size"],
      [{ fontSize: "6000pt" }, "invalid-font-size"],
      [{ fontWeight: 0 }, "invalid-font-weight"],
      [{ fontWeight: 1001 }, "invalid-font-weight"],
      [{ fontWeight: -100 }, "invalid-font-weight"],
      [{ fontWeight: 400.5 }, "invalid-font-weight"],
      [{ fontStyle: "oblique" }, "invalid-font-style"],
      [{ color: 123 }, "invalid-color"],
      [{ color: "ink" }, "invalid-color"],
      [{ color: "red" }, "invalid-color"],
      [{ textShadow: { color: "tileAccent" } }, "invalid-color"],
      [{ lineHeight: 0 }, "invalid-text-length"],
      [{ verticalAlign: "baseline" }, "invalid-vertical-align"],
      [{ align: "start" }, "invalid-text-alignment"],
      [{ spacingAfter: -1 }, "invalid-paragraph-spacing"],
      [{ letterSpacing: 10_000 }, "invalid-letter-spacing"],
      [{ textTransform: "lowercase" }, "invalid-text-transform"],
      [{ textDecoration: "overline" }, "invalid-text-decoration"],
      [{ textWrap: "balance" }, "invalid-text-wrap"],
      [{ textShadow: { unknown: 1 } }, "invalid-shadow"],
      [{ lang: "not valid!" }, "invalid-text-lang"],
    ];

    for (const [style, expectedCode] of cases) {
      assert.throws(
        () => validateTextStyle(style),
        (error: unknown) => {
          assert(error instanceof CapabilityError);
          assert.equal(
            error.code,
            expectedCode,
            `Expected ${expectedCode} for ${JSON.stringify(style)}, got ${error.code}`,
          );
          return true;
        },
      );
    }
  });

  it("validates closed ThemeInput, Theme, ThemeOverride, and ThemeContext color/font narrowing", () => {
    const input: ThemeInput = themeInputSchema.parse({
      name: "Brand",
      colors: { accent1: " #abc ", accent2: "#1a2b3c" },
      fonts: { major: "Aptos Display", minor: "Aptos" },
    });
    assert.equal(input.colors?.accent1, "#AABBCC");
    assert.equal(input.colors?.accent2, "#1A2B3C");
    const theme: Theme = themeSchema.parse(defineTheme(input));
    assert.equal(theme.name, "Brand");
    assert.equal(theme.colors.dk1, "#000000");
    assert.equal(theme.colors.accent1, "#AABBCC");
    assert.equal(theme.colors.accent2, "#1A2B3C");

    const ctx = createThemeContext(input, {
      fontSize: "14pt",
      color: "accent1",
    });
    const derived = ctx.with({ fontWeight: "bold" });
    assert.equal(derived.resolveColor("accent1").hex, "AABBCC");
    assert.equal(derived.resolveColor("#1234").hex, "112233");
    assert.equal(derived.resolveColor({ color: "bg1", alpha: 0.5 }).alpha, 0.5);
    assert.equal(derived.resolveFont("major").family, "Aptos Display");
    assert.equal(derived.resolveFont("major").themeFont, "major");
    assert.equal(derived.resolveFont("Georgia").family, "Georgia");
    assert.equal(derived.resolveFont("Georgia").themeFont, undefined);

    for (const invalidColor of [
      "ink",
      "red",
      "tileAccent",
      "#12345",
    ] as const) {
      assert.equal(colorValueSchema.safeParse(invalidColor).success, false);
      assert.throws(
        () => derived.resolveColor(invalidColor as never),
        (error: unknown) => {
          assert(error instanceof CapabilityError);
          assert.equal(error.code, "invalid-color");
          return true;
        },
      );
    }

    for (const badThemeInput of [
      { colors: { followHlink: "#112233" } },
      { colors: { tileAccent: "#112233" } },
      { colors: { accent1: "#11223380" } },
      { colors: { accent1: "#1234" } },
      { colors: { dk1: "accent1" } },
      { colors: { accent1: { slot: "accent1", lumMod: 0.5 } } },
      { fonts: { minro: "Aptos" } },
      { fonts: { mono: "Consolas" } },
      { fonts: { major: "major" } },
      { fonts: { minor: "MINOR" } },
      { fonts: { major: "+mj-lt" } },
      { fonts: { minor: "+mn-ea" } },
      { fonts: { major: { family: "major" } } },
      { fonts: { minor: { family: "+mn-lt" } } },
    ]) {
      assert.equal(themeInputSchema.safeParse(badThemeInput).success, false);
      assert.throws(
        () => defineTheme(badThemeInput as never),
        (error: unknown) => {
          assert(error instanceof CapabilityError);
          assert.equal(error.code, "invalid-theme");
          return true;
        },
      );
    }

    for (const badOverride of [
      { colors: { followHlink: "#112233" } },
      { colors: { accent1: "#11223380" } },
      { colors: { dk1: "accent2" } },
      { fonts: { mono: "Consolas" } },
      { fonts: { major: "major" } },
      { fonts: { minor: "+mn-lt" } },
    ]) {
      assert.equal(themeOverrideSchema.safeParse(badOverride).success, false);
    }
  });

  it("preserves Theme instance identity, supports .set() normalization, and guards defaultTheme", () => {
    const brand = defineTheme({
      name: "Brand A",
      colors: { accent1: "#112233", accent2: "#abc" },
      fonts: { major: "Aptos Display" },
    });
    assert.equal(isTheme(brand), true);
    assert.equal(isTheme(defaultTheme), true);
    assert.equal(isTheme({ name: "Plain", colors: {}, fonts: {} }), false);
    assert.equal(brand.colors.accent2, "#AABBCC");

    // CommonNode preserves Theme reference identity when snapshotting props.
    assert.equal(createDeckNode({ theme: brand }).props["theme"], brand);

    const same = brand.set({
      name: "Brand A v2",
      colors: { accent1: " #aa0000 " as HexColor },
      fonts: { minor: "Jost*" },
    });
    assert.equal(same, brand);
    assert.equal(brand.name, "Brand A v2");
    assert.equal(brand.colors.accent1, "#AA0000");
    assert.equal(brand.colors.accent2, "#AABBCC");
    assert.equal(brand.fonts.major, "Aptos Display");
    assert.equal(brand.fonts.minor, "Jost*");

    assert.throws(
      () => brand.set({ colors: { custom: "#112233" } } as never),
      (error: unknown) => {
        assert(error instanceof CapabilityError);
        assert.equal(error.code, "invalid-theme");
        return true;
      },
    );

    assert.throws(() => defaultTheme.set({ name: "Mutated" }), TypeError);

    // deckPropsSchema accepts both identity-bearing Theme instances and inline ThemeInput.
    const parsedBrandDeck = deckPropsSchema.parse({ theme: brand });
    assert.equal(parsedBrandDeck.theme, brand);
    assert.equal(isTheme(parsedBrandDeck.theme), true);
    const parsedInlineDeck = deckPropsSchema.parse({
      theme: { name: "Inline", colors: { accent1: "#abc" } },
    });
    assert.equal(isTheme(parsedInlineDeck.theme), false);
    assert.equal(parsedInlineDeck.theme?.colors?.accent1, "#AABBCC");
    assert.equal(
      deckPropsSchema.safeParse({
        theme: { colors: { followHlink: "#112233" } },
      }).success,
      false,
    );
  });

  it("validates color maps, format schemes, and theme overrides and resolves scheme colors through active colorMap", () => {
    assert.equal(schemeColorSlotSchema.parse("tx1"), "tx1");
    assert.equal(schemeColorNameSchema.parse("lt2"), "lt2");

    const invertedMap: ColorMap = colorMapSchema.parse({
      ...defaultColorMap,
      bg1: "dk1",
      tx1: "lt1",
      bg2: "dk2",
      tx2: "lt2",
    });
    const overrideMap: ColorMapOverride =
      colorMapOverrideSchema.parse(invertedMap);
    assert.equal(colorMapOverrideSchema.parse("master"), "master");

    const fmt: FormatScheme = formatSchemeSchema.parse({
      name: "CustomFmt",
      bgFillStyles: [
        { kind: "solid", color: { color: "phClr" } },
        {
          kind: "gradient",
          stops: [
            { position: 0, color: { color: "phClr", tint: 0.4 } },
            { position: 1, color: { color: "phClr", shade: 0.2 } },
          ],
          path: {
            shape: "circle",
            fillToRect: { l: 0.5, t: 0.5, r: 0.5, b: 0.5 },
          },
        },
      ],
    });
    const override: ThemeOverride = themeOverrideSchema.parse({
      colors: { accent1: "#f60" },
      fonts: { major: "Georgia" },
      formatScheme: fmt,
    });
    assert.equal(override.colors?.accent1, "#FF6600");

    const theme = defineTheme({
      colors: {
        dk1: "#111111",
        lt1: "#EEEEEE",
        dk2: "#222222",
        lt2: "#DDDDDD",
        accent1: "#0055AA",
      },
      fonts: { major: "Aptos Display", minor: "Aptos" },
    });
    const baseCtx = createThemeContext(theme, undefined, {
      colorMap: invertedMap,
    });
    // tx1 resolves via invertedMap (tx1 -> lt1 -> #EEEEEE) while preserving scheme: "tx1"
    const tx1Resolved = baseCtx.resolveColor("tx1");
    assert.equal(tx1Resolved.hex, "EEEEEE");
    assert.equal(tx1Resolved.slot, "lt1");
    assert.equal(tx1Resolved.scheme, "tx1");

    // Direct native slot lt2 resolves directly to lt2 (#DDDDDD) with scheme: "lt2"
    const lt2Resolved = baseCtx.resolveColor("lt2");
    assert.equal(lt2Resolved.hex, "DDDDDD");
    assert.equal(lt2Resolved.slot, "lt2");
    assert.equal(lt2Resolved.scheme, "lt2");

    // ThemeOverride overrides accent1 & major font without mutating base Theme
    const overriddenCtx = baseCtx.withNativeOverrides(override, overrideMap);
    assert.equal(overriddenCtx.resolveColor("accent1").hex, "FF6600");
    assert.equal(overriddenCtx.resolveFont("major").family, "Georgia");
    assert.equal(theme.colors.accent1, "#0055AA");
    assert.equal(overriddenCtx.theme, theme);
  });
});
