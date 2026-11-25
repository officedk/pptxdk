import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

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

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
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

function runSubprocess(conditions: readonly string[], script: string): string {
  const child = spawnSync(
    process.execPath,
    [...conditions, "--input-type=module", "--eval", script],
    { cwd: packageRoot, encoding: "utf8" },
  );
  assert.equal(child.status, 0, child.stderr || child.stdout);
  return child.stdout;
}

describe("@officedk/pptxdk/build capability selection", () => {
  it("builds shapes with text, theme, geometry, gradient, morph, and motion via @officedk/pptxdk/vanilla with capabilities: ['shape', 'motion', 'transition'] across source and dist without loading unselected implementations, React, or preview", () => {
    const vanillaScript = `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js") ||
      resolved.url.endsWith("/capabilities/src/chart/chart.ts") ||
      resolved.url.endsWith("/capabilities/dist/chart/chart.js") ||
      resolved.url.endsWith("/capabilities/src/table/table.ts") ||
      resolved.url.endsWith("/capabilities/dist/table/table.js")
    ) {
      throw new Error("unselected capability implementation was loaded: " + resolved.url);
    }
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("vanilla trimmed build loaded forbidden dependency: " + specifier);
    }
    return resolved;
  },
});

const {
  createMorphScope,
  createMotionScopeHandle,
  deck,
  defineTheme,
  fade,
  font,
  linearGradient,
  motionScope,
  paragraph,
  preset,
  shape,
  slide,
  span,
  step,
  transitions,
} = await import("@officedk/pptxdk/vanilla");
const { build } = await import("@officedk/pptxdk/build");
const { validatePptx } = await import("@pptxdk/validation");

const jostBytes = new Uint8Array(readFileSync("./test/assets/fonts/Jost-400-Book.ttf"));
const family = font("Jost", [{ weight: 400, data: jostBytes }], { substitute: true });
const theme = defineTheme({ name: "TrimmedVanilla", fonts: { major: family, minor: family } });
const morph = createMorphScope();
const hero = morph.handle("hero");
const scope = createMotionScopeHandle();

const doc = deck({ width: 240, height: 120, theme }).append(
  slide({ id: "trimmed-vanilla", transition: transitions.fade() }).append(
    step({ n: 1 }).append(
      motionScope(scope).append(
        shape({
          id: "plain",
          morphKey: hero,
          geometry: preset("roundRect", { adj: 20000 }),
          x: 0,
          y: 0,
          width: 20,
          height: 20,
          style: { backgroundImage: linearGradient(90, "#0F766E", "#14B8A6") },
          enter: fade({ duration: 150 }),
        }),
      ),
    ),
    shape({
      id: "labeled",
      name: "labeled",
      x: 24,
      y: 0,
      width: 100,
      height: 32,
      text: "Shape text",
    }),
    shape({
      id: "rich",
      name: "rich",
      x: 24,
      y: 40,
      width: 120,
      height: 32,
    }).content(paragraph(span("Rich shape"))),
  ),
  slide({ id: "trimmed-vanilla-next", transition: transitions.morph({ fallback: transitions.fade() }) }).append(
    shape({
      id: "plain-next",
      morphKey: hero,
      geometry: preset("roundRect", { adj: 20000 }),
      x: 20,
      y: 10,
      width: 40,
      height: 40,
      fill: "#0F766E",
    }),
  ),
);

const result = await build(doc, { capabilities: ["shape", "motion", "transition"] });
const file = await validatePptx(result.pptx);
if (file.errors.length !== 0) {
  throw new Error(JSON.stringify(file.errors));
}
if (file.slides[0]?.transition?.kind !== "fade" || file.slides[1]?.transition?.kind !== "morph") {
  throw new Error("Expected fade and morph transitions on trimmed vanilla slides");
}
if (file.slides[1]?.objects[0]?.name !== "!!morph-1") {
  throw new Error("Expected matching morph pair name on slide 2: " + file.slides[1]?.objects[0]?.name);
}
const texts = file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]);
process.stdout.write(JSON.stringify(texts));
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(JSON.parse(runSubprocess(conditions, vanillaScript)), [
        ["!!morph-1", ""],
        ["labeled", "Shape text"],
        ["rich", "Rich shape"],
      ]);
    }
  }, 45_000);

  it("builds shapes with text, theme, geometry, gradient, morph, and motion via @officedk/pptxdk/react with capabilities: ['shape', 'motion', 'transition'] across source and dist without loading unselected implementations or preview", () => {
    const reactScript = `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js") ||
      resolved.url.endsWith("/capabilities/src/chart/chart.ts") ||
      resolved.url.endsWith("/capabilities/dist/chart/chart.js") ||
      resolved.url.endsWith("/capabilities/src/table/table.ts") ||
      resolved.url.endsWith("/capabilities/dist/table/table.js")
    ) {
      throw new Error("unselected capability implementation was loaded: " + resolved.url);
    }
    if (
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("react trimmed build loaded preview dependency: " + specifier);
    }
    return resolved;
  },
});

const { createElement } = await import("react");
const {
  createMorphScope,
  createMotionScopeHandle,
  Deck,
  defineTheme,
  fade,
  font,
  linearGradient,
  MotionScope,
  P,
  preset,
  Shape,
  Slide,
  Span,
  Step,
  transitions,
} = await import("@officedk/pptxdk/react");
const { build } = await import("@officedk/pptxdk/build");
const { validatePptx } = await import("@pptxdk/validation");

const jostBytes = new Uint8Array(readFileSync("./test/assets/fonts/Jost-400-Book.ttf"));
const family = font("Jost", [{ weight: 400, data: jostBytes }], { substitute: true });
const theme = defineTheme({ name: "TrimmedReact", fonts: { major: family, minor: family } });
const morph = createMorphScope();
const hero = morph.handle("hero");
const scope = createMotionScopeHandle();

const tree = createElement(
  Deck,
  { width: 240, height: 120, theme },
  createElement(
    Slide,
    { id: "trimmed-react", transition: transitions.fade() },
    createElement(
      Step,
      { n: 1 },
      createElement(
        MotionScope,
        { scope },
        createElement(Shape, {
          id: "plain",
          morphKey: hero,
          geometry: preset("roundRect", { adj: 20000 }),
          x: 0,
          y: 0,
          width: 20,
          height: 20,
          style: { backgroundImage: linearGradient(90, "#0F766E", "#14B8A6") },
          enter: fade({ duration: 150 }),
        }),
      ),
    ),
    createElement(Shape, {
      id: "labeled",
      name: "labeled",
      x: 24,
      y: 0,
      width: 100,
      height: 32,
      text: "React shape text",
    }),
    createElement(
      Shape,
      { id: "rich", name: "rich", x: 24, y: 40, width: 120, height: 32 },
      createElement(P, null, createElement(Span, null, "Rich react shape")),
    ),
  ),
  createElement(
    Slide,
    { id: "trimmed-react-next", transition: transitions.morph({ fallback: transitions.fade() }) },
    createElement(Shape, {
      id: "plain-next",
      morphKey: hero,
      geometry: preset("roundRect", { adj: 20000 }),
      x: 20,
      y: 10,
      width: 40,
      height: 40,
      fill: "#0F766E",
    }),
  ),
);

const result = await build(tree, { capabilities: ["shape", "motion", "transition"] });
const file = await validatePptx(result.pptx);
if (file.errors.length !== 0) {
  throw new Error(JSON.stringify(file.errors));
}
if (file.slides[0]?.transition?.kind !== "fade" || file.slides[1]?.transition?.kind !== "morph") {
  throw new Error("Expected fade and morph transitions on trimmed react slides");
}
if (file.slides[1]?.objects[0]?.name !== "!!morph-1") {
  throw new Error("Expected matching morph pair name on slide 2: " + file.slides[1]?.objects[0]?.name);
}
const texts = file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]);
process.stdout.write(JSON.stringify(texts));
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(JSON.parse(runSubprocess(conditions, reactScript)), [
        ["!!morph-1", ""],
        ["labeled", "React shape text"],
        ["rich", "Rich react shape"],
      ]);
    }
  }, 45_000);

  it("diagnoses missing independent Text declared inside composed components and vanilla documents, reuses one Resolution for emit and paint, and keeps default build intact", async () => {
    function NestedText() {
      return (
        <Text id="missing-text" x={0} y={0} width={100} height={32}>
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
              x: 0,
              y: 0,
              width: 100,
              height: 32,
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
          <Shape id="card" x={24} y={24} width={200} height={60} fill="#E2E8F0">
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
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    const png = PNG.sync.read(Buffer.from(painted.pages[0].png!));
    assert.ok(png.width > 0 && png.height > 0);

    const fullResult = await buildDefault(
      <Deck width={320} height={180}>
        <Slide id="full">
          <Text id="heading" x={24} y={24} width={200} height={40}>
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
    assert.equal(officialCapabilityNames.length, 14);
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
        footer: { x: 40, y: 480, width: 320, height: 28, text: "Footer" },
      },
      placeholders: {
        title: { type: "title", x: 40, y: 32, width: 560, height: 48 },
        body: { type: "obj", idx: 1, x: 40, y: 96, width: 560, height: 320 },
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
              <Text x={40} y={40} width={200} height={40}>
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
                <Shape x={40} y={40} width={100} height={60} fill="#2563EB" />
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
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
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
      assert.deepEqual(page.diagnostics, []);
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
