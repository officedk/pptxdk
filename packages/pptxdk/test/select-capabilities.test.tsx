import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { createPlaceholderNode } from "@pptxdk/capabilities/placeholder";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import {
  CommonContainer,
  CommonNode,
  build as buildDefault,
  defineSlideLayout,
  defineTheme,
  font,
  transitions,
} from "pptxdk";
import {
  BuildError,
  build as buildTrimmed,
  createImporter,
  emit,
  officialCapabilityNames,
  resolve as resolveTrimmed,
  selectCapabilities,
} from "pptxdk/build";
import { paint } from "pptxdk/preview";
import { Deck, P, Shape, Slide, Span, Text } from "pptxdk/react";
import { deck, slide, slideLayout, slideMaster, text } from "pptxdk/vanilla";
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

describe("pptxdk/build capability selection", () => {
  it("builds shapes with text, theme, geometry, gradient, morph, and motion via pptxdk/vanilla with capabilities: ['shape', 'motion', 'transition'] across source and dist without loading unselected implementations, React, or preview", () => {
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
} = await import("pptxdk/vanilla");
const { build } = await import("pptxdk/build");
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

  it("builds shapes with text, theme, geometry, gradient, morph, and motion via pptxdk/react with capabilities: ['shape', 'motion', 'transition'] across source and dist without loading unselected implementations or preview", () => {
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
} = await import("pptxdk/react");
const { build } = await import("pptxdk/build");
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
    assert.equal(reactError.diagnostics[0]?.severity, "error");
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
        error.diagnostics[0]?.code === "unknown-capability" &&
        error.diagnostics[0]?.path === "build.capabilities[0]",
    );

    await assert.rejects(
      () =>
        buildTrimmed(deck({ width: 100, height: 100 }), {
          capabilities: "shape" as unknown as readonly [],
        }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-capability-selection",
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
});
