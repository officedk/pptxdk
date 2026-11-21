import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { build as buildDefault, defineTheme, font, transitions } from "pptxdk";
import {
  BuildError,
  build as buildTrimmed,
  emit,
  officialCapabilityNames,
  resolve as resolveTrimmed,
  selectCapabilities,
} from "pptxdk/build";
import { paint } from "pptxdk/preview";
import { Deck, P, Shape, Slide, Span, Text } from "pptxdk/react";
import { deck, slide, text } from "pptxdk/vanilla";
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
  it("builds shapes with text via pptxdk/vanilla with capabilities: ['shape'] across source and dist without loading text box implementation, React, or preview", () => {
    const vanillaScript = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js")
    ) {
      throw new Error("text-box implementation was loaded: " + resolved.url);
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

const { deck, paragraph, shape, slide, span } = await import("pptxdk/vanilla");
const { build } = await import("pptxdk/build");
const { validatePptx } = await import("@pptxdk/validation");

const doc = deck({ width: 240, height: 120 }).append(
  slide({ id: "trimmed-vanilla" }).append(
    shape({ id: "plain", name: "plain", x: 0, y: 0, width: 20, height: 20 }),
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
);

const result = await build(doc, { capabilities: ["shape"] });
const file = await validatePptx(result.pptx);
if (file.errors.length !== 0) {
  throw new Error(JSON.stringify(file.errors));
}
const texts = file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]);
process.stdout.write(JSON.stringify(texts));
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(JSON.parse(runSubprocess(conditions, vanillaScript)), [
        ["plain", ""],
        ["labeled", "Shape text"],
        ["rich", "Rich shape"],
      ]);
    }
  }, 45_000);

  it("builds shapes with text via pptxdk/react with capabilities: ['shape', 'transition'] across source and dist without loading text box implementation or preview", () => {
    const reactScript = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js")
    ) {
      throw new Error("text-box implementation was loaded: " + resolved.url);
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
const { Deck, P, Shape, Slide, Span, transitions } = await import("pptxdk/react");
const { build } = await import("pptxdk/build");
const { validatePptx } = await import("@pptxdk/validation");

const tree = createElement(
  Deck,
  { width: 240, height: 120 },
  createElement(
    Slide,
    { id: "trimmed-react", transition: transitions.fade() },
    createElement(Shape, { id: "plain", name: "plain", x: 0, y: 0, width: 20, height: 20 }),
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
);

const result = await build(tree, { capabilities: ["shape", "transition"] });
const file = await validatePptx(result.pptx);
if (file.errors.length !== 0) {
  throw new Error(JSON.stringify(file.errors));
}
if (file.slides[0]?.transition?.kind !== "fade") {
  throw new Error("Expected fade transition on trimmed slide");
}
const texts = file.slides[0]?.objects.map(({ name, text }) => [name, text?.text ?? ""]);
process.stdout.write(JSON.stringify(texts));
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(JSON.parse(runSubprocess(conditions, reactScript)), [
        ["plain", ""],
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
    assert.equal(officialCapabilityNames.length, 11);

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

  it("retires legacy fixed-capability subpaths pptxdk/shape, pptxdk/text, and pptxdk/chart", () => {
    const script = `
const codes = [];
for (const legacySubpath of ["pptxdk/shape", "pptxdk/text", "pptxdk/chart"]) {
  try {
    await import(legacySubpath);
    throw new Error("Expected rejection for " + legacySubpath);
  } catch (error) {
    codes.push(error && typeof error === "object" && "code" in error ? error.code : String(error));
  }
}
process.stdout.write(JSON.stringify(codes));
`;
    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(JSON.parse(runSubprocess(conditions, script)), [
        "ERR_PACKAGE_PATH_NOT_EXPORTED",
        "ERR_PACKAGE_PATH_NOT_EXPORTED",
        "ERR_PACKAGE_PATH_NOT_EXPORTED",
      ]);
    }
  });
});
