import assert from "node:assert/strict";

import { BuildError, defineTheme, isTheme } from "@pptxdk/core";
import { createResolver, targets } from "@pptxdk/core/build";
import type { SlideCapabilityModule } from "@pptxdk/core/capability";
import {
  CommonNode,
  createDeckNode,
  createSectionNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { createThemeContext } from "@pptxdk/core/theme";
import {
  assertXml10Characters,
  emit,
  emittedBuildQuery,
  layoutXml,
  masterXml,
  pxToEmu,
  resolvedColorXml,
  themeOverrideXml,
  themeXml,
  validateOptions,
  xmlDocument,
  xmlEscape,
} from "@pptxdk/emit";
import { describe, it } from "vitest";

describe("@pptxdk/emit pure unit contract", () => {
  it("converts pixels to EMUs, escapes XML 1.0 strings, and formats color and document XML", () => {
    assert.equal(pxToEmu(1), 9525);
    assert.equal(pxToEmu(96), 914400);

    assert.equal(
      xmlEscape(`A & B <tag> "q" 's'\t\r\n`),
      "A &amp; B &lt;tag&gt; &quot;q&quot; &apos;s&apos;&#x9;&#xD;&#xA;",
    );
    assertXml10Characters("Valid text \t\n\r 中文");
    assert.throws(
      () => assertXml10Characters("bad\u0001char"),
      /not allowed in XML 1\.0/u,
    );
    assert.throws(() => xmlEscape("bad\u000bchar"), /not allowed in XML 1\.0/u);

    assert.equal(
      resolvedColorXml({ hex: "112233" }),
      '<a:srgbClr val="112233"/>',
    );
    assert.equal(
      resolvedColorXml({
        hex: "112233",
        slot: "accent1",
        lumMod: 0.6,
        lumOff: 0.4,
        alpha: 0.5,
      }),
      '<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/><a:alpha val="50000"/></a:schemeClr>',
    );
    assert.equal(
      xmlDocument("<p:sld/>"),
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld/>',
    );
  });

  it("validates emit options and rejects invalid target or option shapes with BuildError", () => {
    assert.deepEqual(validateOptions({}), targets.transitional);
    assert.deepEqual(
      validateOptions({ embedFonts: "subset" }),
      targets.transitional,
    );

    for (const badOptions of [
      { embedFonts: "none" },
      { target: { format: "strict" } },
      { target: { format: "transitional", extensions: ["unknown-ext"] } },
      { unknownField: true },
    ]) {
      assert.throws(
        () => validateOptions(badOptions as never),
        (error: unknown) =>
          error instanceof BuildError && error.diagnostics.length > 0,
      );
    }
  });

  it("emits a valid PPTX ZIP archive for a resolved deck and rejects aborted or forbidden-extension builds", () => {
    const resolve = createResolver([]);
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "intro" }),
    );
    const resolution = resolve(doc);
    const emitted = emit(resolution);

    assert.ok(emitted.pptx instanceof Uint8Array);
    assert.ok(emitted.pptx.byteLength > 0);
    assert.equal(
      Buffer.from(emitted.pptx).readUInt32LE(0),
      0x04034b50,
      "Expected PKZIP signature",
    );
    assert.equal(emitted.report.slides.length, 1);
    assert.equal(emitted.report.slides[0]?.id, "intro");

    const query = emittedBuildQuery(emitted.report, resolution);
    assert.equal(query.slides.length, 1);
    assert.equal(query.slide("intro").id, "intro");

    const controller = new AbortController();
    const reason = new Error("aborted emit");
    controller.abort(reason);
    assert.throws(
      () => emit(resolution, { signal: controller.signal }),
      (error: unknown) => error === reason,
    );

    const sectionCapability: SlideCapabilityModule = {
      declaration: "test-sections",
      slideProps: [],
      slot: "timing",
      bindSections: () => ({
        extensions: ["p14"],
        emit: () => ({ xml: "" }),
      }),
      bindSlide: () => ({
        resolveSlide: () => ({
          extensions: [],
          emit: () => ({ xml: "" }),
        }),
      }),
    };
    const resolveWithSections = createResolver([sectionCapability]);
    const sectioned = createDeckNode({ width: 960, height: 540 }).append(
      createSectionNode({ name: "Sec" }).append(createSlideNode({ id: "s1" })),
    );
    const sectionedResolution = resolveWithSections(sectioned);
    assert.throws(
      () => emit(sectionedResolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "target-violation",
    );

    assert.ok(CommonNode);
  });

  it("emits OS/2 panose and pitchFamily attributes on theme font slots when Font faces carry SFNT bytes", () => {
    // Build a minimal valid SFNT containing OS/2 (panose=02 0B 06 03 02 02 02 02 02 04 -> Swiss/Sans proportional = 0x22 = 34) and post.
    const os2 = new Uint8Array(78);
    const panose = [2, 11, 6, 3, 2, 2, 2, 2, 2, 4];
    os2.set(panose, 32);
    const post = new Uint8Array(32);

    const sfnt = new Uint8Array(12 + 2 * 16 + 80 + 32);
    const view = new DataView(sfnt.buffer);
    view.setUint32(0, 0x00010000);
    view.setUint16(4, 2);
    sfnt.set([0x4f, 0x53, 0x2f, 0x32], 12); // "OS/2"
    view.setUint32(20, 44);
    view.setUint32(24, 78);
    sfnt.set([0x70, 0x6f, 0x73, 0x74], 28); // "post"
    view.setUint32(36, 44 + 80);
    view.setUint32(40, 32);
    sfnt.set(os2, 44);
    sfnt.set(post, 44 + 80);

    const theme = defineTheme({
      name: "Classified Theme",
      fonts: {
        major: {
          family: "Custom Sans",
          faces: [{ weight: 400, style: "normal", data: sfnt }],
          substitute: true,
        },
        minor: "Plain System Font",
      },
    });
    assert.equal(isTheme(theme), true);
    const xml = themeXml(theme);
    assert.match(
      xml,
      /<a:majorFont><a:latin typeface="Custom Sans" panose="020B0603020202020204" pitchFamily="34"\/><a:ea typeface="Custom Sans" panose="020B0603020202020204" pitchFamily="34"\/><a:cs typeface="Custom Sans"\/><\/a:majorFont>/u,
    );
    assert.match(
      xml,
      /<a:minorFont><a:latin typeface="Plain System Font"\/><a:ea typeface="Plain System Font"\/><a:cs typeface="Plain System Font"\/><\/a:minorFont>/u,
    );
  });

  it("emits orthogonal majorEastAsian and minorEastAsian slots on themeXml and themeOverrideXml", () => {
    const base = defineTheme({
      name: "Bilingual",
      fonts: {
        major: "Inter",
        minor: "Aptos",
        majorEastAsian: "PingFang SC",
        minorEastAsian: "Noto Sans SC",
      },
    });
    const xml = themeXml(base);
    assert.match(
      xml,
      /<a:majorFont><a:latin typeface="Inter"\/><a:ea typeface="PingFang SC"\/><a:cs typeface="Inter"\/><\/a:majorFont>/u,
    );
    assert.match(
      xml,
      /<a:minorFont><a:latin typeface="Aptos"\/><a:ea typeface="Noto Sans SC"\/><a:cs typeface="Aptos"\/><\/a:minorFont>/u,
    );

    const ovrXml = themeOverrideXml(
      { fonts: { minorEastAsian: "Microsoft YaHei" } },
      base,
    );
    assert.match(
      ovrXml,
      /<a:minorFont><a:latin typeface="Aptos"\/><a:ea typeface="Microsoft YaHei"\/><a:cs typeface="Aptos"\/><\/a:minorFont>/u,
    );
  });

  it("pads theme formatScheme style lists to at least 3 items in themeXml", () => {
    const theme = defineTheme({
      name: "PartialFmt",
      formatScheme: {
        fillStyles: [{ kind: "solid", color: { color: "phClr" } }],
        lineStyles: [{ widthEmu: 9525 }],
        effectStyles: [{}],
        bgFillStyles: [{ kind: "solid", color: { color: "phClr" } }],
      },
    });
    const xml = themeXml(theme);
    const fillMatch = /<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/u.exec(xml);
    const lnMatch = /<a:lnStyleLst>([\s\S]*?)<\/a:lnStyleLst>/u.exec(xml);
    const effMatch = /<a:effectStyleLst>([\s\S]*?)<\/a:effectStyleLst>/u.exec(
      xml,
    );
    const bgMatch = /<a:bgFillStyleLst>([\s\S]*?)<\/a:bgFillStyleLst>/u.exec(
      xml,
    );
    assert.ok(fillMatch && lnMatch && effMatch && bgMatch);
    assert.equal(fillMatch[1]!.match(/<a:solidFill>/gu)?.length, 3);
    assert.equal(lnMatch[1]!.match(/<a:ln\b/gu)?.length, 3);
    assert.equal(effMatch[1]!.match(/<a:effectStyle>/gu)?.length, 3);
    assert.equal(bgMatch[1]!.match(/<a:solidFill>/gu)?.length, 3);
  });

  it("keeps <a:latin> and <a:ea> orthogonal in master/layout <a:defRPr> without leaking fontFamily into <a:ea>", () => {
    const theme = defineTheme({
      name: "DualScript",
      fonts: {
        major: "Inter",
        minor: "Aptos",
        majorEastAsian: "PingFang SC",
        minorEastAsian: "Noto Sans SC",
      },
    });
    const themeContext = createThemeContext(theme);
    const mXml = masterXml(
      {
        key: "master-1",
        theme,
        layouts: [],
        textStyles: {
          titleStyle: {
            1: { fontFamily: "Arial" },
          },
          bodyStyle: {
            1: { eastAsianFontFamily: "Microsoft YaHei" },
          },
        },
      },
      [{ id: 2147483649, rId: "rId1" }],
      themeContext,
    );
    assert.match(
      mXml,
      /<p:titleStyle><a:lvl1pPr><a:defRPr><a:latin typeface="Arial"\/><a:cs typeface="Arial"\/><\/a:defRPr><\/a:lvl1pPr><\/p:titleStyle>/u,
    );
    assert.doesNotMatch(mXml, /<a:ea typeface="Arial"\/>/u);
    assert.match(
      mXml,
      /<p:bodyStyle><a:lvl1pPr><a:defRPr><a:ea typeface="Microsoft YaHei"\/><\/a:defRPr><\/a:lvl1pPr><\/p:bodyStyle>/u,
    );

    const lXml = layoutXml(
      {
        key: "layout-1",
        name: "DualScriptLayout",
        type: "obj",
        placeholders: {
          title: {
            type: "title",
            style: { fontFamily: "Georgia" },
          },
          body: {
            type: "body",
            idx: 1,
            listStyles: {
              1: { fontFamily: "Helvetica" },
              2: { eastAsianFontFamily: "SimSun" },
            },
          },
        },
      },
      themeContext,
    );
    assert.match(
      lXml,
      /<a:latin typeface="Georgia"\/><a:cs typeface="Georgia"\/>/u,
    );
    assert.doesNotMatch(lXml, /<a:ea typeface="Georgia"\/>/u);
    assert.match(
      lXml,
      /<a:latin typeface="Helvetica"\/><a:cs typeface="Helvetica"\/>/u,
    );
    assert.doesNotMatch(lXml, /<a:ea typeface="Helvetica"\/>/u);
    assert.match(lXml, /<a:ea typeface="SimSun"\/>/u);
  });
});
