import assert from "node:assert/strict";

import { CapabilityError } from "@pptxdk/core";
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
  fontTokenSchema,
  formatSchemeSchema,
  gradientStopSchema,
  isTheme,
  linearGradientSchema,
  nativeColorSlotSchema,
  radialGradientSchema,
  schemeColorNameSchema,
  schemeColorSlotSchema,
  shadowSchema,
  textAlignmentSchema,
  textStyleSchema,
  themeColorSchema,
  themeInputSchema,
  themeOverrideSchema,
  themeSchema,
  themeScopeSchema,
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
  type LinearGradient,
  type NativeColorSlot,
  type RadialGradient,
  type Shadow,
  type TextAlignment,
  type TextStyle,
  type Theme,
  type ThemeColor,
  type ThemeInput,
  type ThemeOverride,
  type ThemeScope,
  type ThemeShade,
  type VisualStyle,
} from "@pptxdk/core/theme";
import { describe, it } from "vitest";

import { snapshotValue } from "../src/input-values.ts";

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
    const fontToken: FontReference = fontTokenSchema.parse("minor");
    assert.deepEqual(fontRef, fontValue);
    assert.equal(fontRef.faces?.[0]?.data, face.data);
    assert.equal(fontToken, "minor");

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
        { fontFamily: { family: "Bad", faces: [{ weight: 150 }] } },
        "invalid-font-face",
      ],
      [{ fontSize: "20em" }, "invalid-text-length"],
      [{ fontSize: 0 }, "invalid-font-size"],
      [{ fontSize: -4 }, "invalid-font-size"],
      [{ fontSize: "6000pt" }, "invalid-font-size"],
      [{ fontWeight: 250 }, "invalid-font-weight"],
      [{ fontStyle: "oblique" }, "invalid-font-style"],
      [{ color: 123 }, "invalid-color"],
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

  it("validates ThemeInput, Theme, and ThemeScope and preserves invalid-theme-scope errors", () => {
    const input: ThemeInput = themeInputSchema.parse({
      name: "Brand",
      colors: { accent1: "#112233", ink: "accent1" },
      fonts: { major: "Aptos Display", body: "Aptos" },
    });
    const theme: Theme = themeSchema.parse(defineTheme(input));
    assert.equal(theme.name, "Brand");
    assert.equal(theme.colors.dk1, "#000000");

    const scope: ThemeScope = themeScopeSchema.parse({
      colors: { ink: "#445566" },
      fonts: { body: "Jost" },
    });
    const ctx = createThemeContext(input, { fontSize: "14pt", color: "ink" });
    const scoped = ctx.with(scope, { fontWeight: "bold" });
    assert.equal(scoped.resolveColor("ink").hex, "445566");
    assert.equal(scoped.resolveFont("body").family, "Jost");

    for (const badScope of [
      "not-an-object",
      { extra: {} },
      { colors: { dk1: "#111111" } },
      { colors: { missing: "#111111" } },
      { fonts: { missing: "Arial" } },
      { fonts: { body: "" } },
    ]) {
      assert.throws(
        () => ctx.with(badScope as ThemeScope),
        (error: unknown) => {
          assert(error instanceof CapabilityError);
          assert.equal(error.code, "invalid-theme-scope");
          return true;
        },
      );
    }
  });

  it("preserves Theme instance identity, supports .set() mutation, and guards defaultTheme", () => {
    const brand = defineTheme({
      name: "Brand A",
      colors: { accent1: "#112233", ink: "accent1" },
      fonts: { major: "Aptos Display" },
    });
    assert.equal(isTheme(brand), true);
    assert.equal(isTheme(defaultTheme), true);
    assert.equal(isTheme({ name: "Plain", colors: {}, fonts: {} }), false);

    // snapshotValue preserves Theme reference identity (used by CommonNode props).
    assert.equal(snapshotValue(brand), brand);

    const same = brand.set({
      name: "Brand A v2",
      colors: { accent1: "#AA0000" },
      fonts: { minor: "Jost*" },
    });
    assert.equal(same, brand);
    assert.equal(brand.name, "Brand A v2");
    assert.equal(brand.colors.accent1, "#AA0000");
    assert.equal(brand.colors["ink"], "accent1");
    assert.equal(brand.fonts.major, "Aptos Display");
    assert.equal(brand.fonts.minor, "Jost*");

    assert.throws(() => defaultTheme.set({ name: "Mutated" }), TypeError);
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
      colors: { accent1: "#FF6600" },
      fonts: { major: "Georgia" },
      formatScheme: fmt,
    });

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
