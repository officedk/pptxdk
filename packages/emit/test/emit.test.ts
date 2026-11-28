import assert from "node:assert/strict";

import { BuildError } from "@pptxdk/core";
import { createResolver, targets } from "@pptxdk/core/build";
import type { SlideCapabilityModule } from "@pptxdk/core/capability";
import {
  CommonNode,
  createDeckNode,
  createSectionNode,
  createSlideNode,
} from "@pptxdk/core/document";
import {
  assertXml10Characters,
  emit,
  emittedBuildQuery,
  pxToEmu,
  resolvedColorXml,
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
});
