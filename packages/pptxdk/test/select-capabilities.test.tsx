import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonContainer,
  CommonNode,
  build as buildDefault,
  defaultColorMap,
  defineSlideLayout,
  defineTheme,
  font,
  transitions,
} from "@officedk/pptxdk";
import {
  BuildError,
  build as buildTrimmed,
  createImporter,
  emit,
  officialCapabilityNames,
  resolve as resolveTrimmed,
  selectCapabilities,
} from "@officedk/pptxdk/build";
import { paint } from "@officedk/pptxdk/preview";
import {
  Deck,
  P,
  Section,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Span,
  Text,
} from "@officedk/pptxdk/react";
import {
  deck,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@officedk/pptxdk/vanilla";
import { createPlaceholderNode } from "@pptxdk/capabilities/placeholder";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { describe, it } from "vitest";

const jostBook = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

function testTheme() {
  const family = font("Jost", [{ weight: 400, data: jostBook }], {
    substitute: true,
  });
  return defineTheme({
    name: "Test",
    fonts: { major: family, minor: family },
  });
}

describe("@officedk/pptxdk/build capability selection", () => {
  it("diagnoses missing independent Text declared inside composed components and vanilla documents, reuses one Resolution for emit and paint, and keeps default build intact", async () => {
    function NestedText() {
      return (
        <Text
          id="missing-text"
          style={{
            position: "absolute",
            left: 0,
            top: 0,
            width: 100,
            height: 32,
          }}
        >
          not available
        </Text>
      );
    }

    let reactError: unknown;
    try {
      await buildTrimmed(
        <Deck width={240} height={120}>
          <Slide id="nested">
            <NestedText />
          </Slide>
        </Deck>,
        { capabilities: ["shape"] },
      );
    } catch (error) {
      reactError = error;
    }
    assert.ok(reactError instanceof BuildError);
    assert.equal(reactError.diagnostics[0]?.code, "missing-capability");
    assert.equal(
      reactError.diagnostics[0]?.path,
      "deck/slide[nested]/text[missing-text]",
    );
    assert.equal(
      reactError.diagnostics[0]?.message,
      "No capability module registered for text",
    );

    let vanillaError: unknown;
    try {
      await buildTrimmed(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "nested" }).append(
            text("not available", {
              id: "missing-text",
              style: {
                position: "absolute",
                left: 0,
                top: 0,
                width: 100,
                height: 32,
              },
            }),
          ),
        ),
        { capabilities: ["shape"] },
      );
    } catch (error) {
      vanillaError = error;
    }
    assert.ok(vanillaError instanceof BuildError);
    assert.equal(vanillaError.diagnostics[0]?.code, "missing-capability");
    assert.equal(
      vanillaError.diagnostics[0]?.path,
      "deck/slide[nested]/text[missing-text]",
    );

    const sharedDoc = (
      <Deck width={320} height={180} theme={testTheme()}>
        <Slide id="shared" transition={transitions.fade()}>
          <Shape
            id="card"
            style={{
              position: "absolute",
              left: 24,
              top: 24,
              width: 200,
              height: 60,
              background: "#E2E8F0",
            }}
          >
            <P>
              <Span>Shared snapshot</Span>
            </P>
          </Shape>
        </Slide>
      </Deck>
    );

    const resolution = await resolveTrimmed(sharedDoc, {
      capabilities: ["shape", "transition"],
    });
    const emitted = emit(resolution);
    const painted = paint(resolution);

    const validated = await validatePptx(emitted.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(
      validated.slides[0]?.objects[0]?.text?.text,
      "Shared snapshot",
    );
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.deepEqual(painted.pages[0]?.placeholders, []);
    const png = PNG.sync.read(Buffer.from(painted.pages[0].png!));
    assert.ok(png.width > 0 && png.height > 0);

    const fullResult = await buildDefault(
      <Deck width={320} height={180}>
        <Slide id="full">
          <Text
            id="heading"
            style={{
              position: "absolute",
              left: 24,
              top: 24,
              width: 200,
              height: 40,
            }}
          >
            Full default build
          </Text>
        </Slide>
      </Deck>,
    );
    const fullValidated = await validatePptx(fullResult.pptx);
    assert.deepEqual(fullValidated.errors, []);
    assert.equal(
      fullValidated.slides[0]?.objects[0]?.text?.text,
      "Full default build",
    );
  });

  it("validates capability names, deduplicates repeated selections, and respects AbortSignal", async () => {
    assert.equal(officialCapabilityNames.length, 15);
    assert.ok(officialCapabilityNames.includes("line"));
    assert.ok(officialCapabilityNames.includes("footer-placeholder"));
    assert.ok(officialCapabilityNames.includes("masters"));
    assert.ok(officialCapabilityNames.includes("placeholder"));

    const modules = await selectCapabilities(["shape", "shape", "transition"]);
    assert.equal(modules.length, 2);

    await assert.rejects(
      () =>
        selectCapabilities([
          "unknown-capability" as (typeof officialCapabilityNames)[number],
        ]),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "missing-capability" &&
        error.diagnostics[0]?.path === "build.capabilities[0]",
    );

    await assert.rejects(
      () =>
        buildTrimmed(deck({ width: 100, height: 100 }), {
          capabilities: "shape" as unknown as readonly [],
        }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-options",
    );

    const controller = new AbortController();
    const reason = new Error("trimmed build aborted");
    controller.abort(reason);
    await assert.rejects(
      () =>
        buildTrimmed(deck({ width: 100, height: 100 }), {
          capabilities: ["shape"],
          signal: controller.signal,
        }),
      (error: unknown) => error === reason,
    );
  });

  it("supports masters, footer-placeholder, and placeholder in trimmed build and createImporter(selectCapabilities(officialCapabilityNames))", async () => {
    const contentLayout = defineSlideLayout({
      id: "content-layout",
      name: "Content Layout",
      type: "obj",
      preserve: true,
      footerPlaceholders: {
        footer: {
          style: {
            position: "absolute",
            left: 40,
            top: 480,
            width: 320,
            height: 28,
          },
          text: "Footer",
        },
      },
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 40,
            top: 32,
            width: 560,
            height: 48,
          },
        },
        body: {
          type: "obj",
          idx: 1,
          style: {
            position: "absolute",
            left: 40,
            top: 96,
            width: 560,
            height: 320,
          },
        },
      },
    });

    const unpopulatedDeck = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "brand", name: "Brand" }).append(
        slideLayout(contentLayout),
      ),
      slide({
        id: "s1",
        layout: contentLayout,
        footer: "Quarterly",
        placeholders: { title: "Title" },
      }).append(
        createPlaceholderNode({
          placeholder: "body",
          nativePlaceholder: { type: "obj", idx: 1 },
        }),
      ),
    );
    const trimmed = await buildTrimmed(unpopulatedDeck, {
      capabilities: [
        "masters",
        "footer-placeholder",
        "placeholder",
        "text",
      ] as unknown as readonly (typeof officialCapabilityNames)[number][],
    });
    const validated = await validatePptx(trimmed.pptx);
    assert.deepEqual(validated.errors, []);

    const fullModules = await selectCapabilities(officialCapabilityNames);
    const importer = createImporter(fullModules);
    const imported = importer(trimmed.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    assert.equal(
      imported.document.children.some(
        (c) => c instanceof CommonNode && c.kind === "slide-master",
      ),
      true,
    );
    const importedSlide = imported.document.children.find(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide",
    );
    const footerProp = importedSlide?.props["footer"] as
      | string
      | { readonly text?: string }
      | undefined;
    assert.equal(
      typeof footerProp === "string" ? footerProp : footerProp?.text,
      "Quarterly",
    );
    assert.equal(
      importedSlide?.children.some(
        (c) => c instanceof CommonNode && c.kind === "placeholder",
      ),
      true,
    );
  });

  it("rejects SlideMaster when masters capability is trimmed and Section when notes capability is trimmed", async () => {
    const blankLayout = defineSlideLayout({
      id: "blank",
      name: "Blank",
      type: "blank",
    });

    await assert.rejects(
      () =>
        buildTrimmed(
          <Deck width={960} height={540}>
            <SlideMaster id="m1">
              <SlideLayout layout={blankLayout} />
            </SlideMaster>
            <Slide id="s1" layout={blankLayout}>
              <Text
                style={{
                  position: "absolute",
                  left: 40,
                  top: 40,
                  width: 200,
                  height: 40,
                }}
              >
                Hello
              </Text>
            </Slide>
          </Deck>,
          { capabilities: ["text"] },
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "missing-capability",
    );

    await assert.rejects(
      () =>
        buildTrimmed(
          <Deck width={960} height={540}>
            <Section name="Intro">
              <Slide id="s1">
                <Shape
                  style={{
                    position: "absolute",
                    left: 40,
                    top: 40,
                    width: 100,
                    height: 60,
                    background: "#2563EB",
                  }}
                />
              </Slide>
            </Section>
          </Deck>,
          { capabilities: ["shape"] },
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "missing-capability",
    );
  });

  it("paints resolved dk1 canvas backdrop via mastersCanvas when colorMap inverts bg1 to dk1 without explicit background", async () => {
    const invertedTheme = defineTheme({
      name: "Inverted",
      colors: { dk1: "#1E293B", lt1: "#F8FAFC" },
    });
    const blankLayout = defineSlideLayout({
      id: "blank",
      name: "Blank",
      type: "blank",
    });

    const resolution = await resolveTrimmed(
      <Deck width={320} height={180}>
        <SlideMaster
          id="m-inv"
          theme={invertedTheme}
          colorMap={{
            ...defaultColorMap,
            bg1: "dk1",
            tx1: "lt1",
            bg2: "dk2",
            tx2: "lt2",
          }}
        >
          <SlideLayout layout={blankLayout} />
        </SlideMaster>
        <Slide id="s-inv" layout={blankLayout} />
      </Deck>,
      { capabilities: ["masters"] },
    );

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.deepEqual(painted.pages[0]?.placeholders, []);
    assert.match(
      painted.pages[0]?.svg ?? "",
      /<rect width="100%" height="100%" fill="#1E293B"\/>/u,
    );
    const png = PNG.sync.read(Buffer.from(painted.pages[0].png!));
    const idx = (10 * png.width + 10) * 4;
    assert.deepEqual(
      [png.data[idx], png.data[idx + 1], png.data[idx + 2]],
      [0x1e, 0x29, 0x3b],
    );
  });

  it("paints backdrop SVG and PNG across nativeTheme, custom colorMap, and themeOverride paths", async () => {
    const plainTheme = defineTheme({
      name: "PlainBackdrop",
      colors: { lt1: "#EFF6FF", dk1: "#0F172A" },
    });
    const invertedTheme = defineTheme({
      name: "InvertedBackdrop",
      colors: { dk1: "#1E293B", lt1: "#F8FAFC" },
    });
    const plainLayout = defineSlideLayout({
      id: "plain-blank",
      name: "Plain Blank",
      type: "blank",
    });
    const invertedLayout = defineSlideLayout({
      id: "inv-blank",
      name: "Inverted Blank",
      type: "blank",
    });

    const resolution = await resolveTrimmed(
      <Deck width={320} height={180}>
        <SlideMaster id="m-plain" theme={plainTheme}>
          <SlideLayout layout={plainLayout} />
        </SlideMaster>
        <SlideMaster
          id="m-inv"
          theme={invertedTheme}
          colorMap={{
            ...defaultColorMap,
            bg1: "dk1",
            tx1: "lt1",
            bg2: "dk2",
            tx2: "lt2",
          }}
        >
          <SlideLayout layout={invertedLayout} />
        </SlideMaster>
        <Slide id="s-plain" layout={plainLayout} />
        <Slide id="s-custom-cmap" layout={invertedLayout} />
        <Slide
          id="s-theme-ovr"
          layout={plainLayout}
          themeOverride={{ colors: { lt1: "#FEF2F2" } }}
        />
      </Deck>,
      { capabilities: ["masters"] },
    );

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 3);
    const expected = [
      { hex: "#EFF6FF", rgb: [0xef, 0xf6, 0xff] },
      { hex: "#1E293B", rgb: [0x1e, 0x29, 0x3b] },
      { hex: "#FEF2F2", rgb: [0xfe, 0xf2, 0xf2] },
    ] as const;
    for (const [pageIndex, spec] of expected.entries()) {
      const page = painted.pages[pageIndex];
      assert.ok(page);
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.placeholders, []);
      assert.match(
        page.svg ?? "",
        new RegExp(
          `<rect width="100%" height="100%" fill="${spec.hex}"/>`,
          "u",
        ),
      );
      const png = PNG.sync.read(Buffer.from(page.png!));
      const idx = (10 * png.width + 10) * 4;
      assert.deepEqual(
        [png.data[idx], png.data[idx + 1], png.data[idx + 2]],
        spec.rgb,
      );
    }
  });
});
