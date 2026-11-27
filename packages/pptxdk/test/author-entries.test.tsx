import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import {
  BuildError,
  build,
  CommonContainer,
  defineTheme,
  font,
  RichTextContainer,
  transitions,
  type Transition,
} from "@officedk/pptxdk";
import {
  Deck,
  expandDocument,
  Shape,
  Slide,
  Text,
  type SlideProps,
} from "@officedk/pptxdk/react";
import {
  deck,
  shape,
  slide,
  text,
  type SlideNode,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
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

describe("public author entrypoints", () => {
  it("shares editable documents between @officedk/pptxdk/react and @officedk/pptxdk/vanilla and builds via @officedk/pptxdk", async () => {
    const fadeTransition: Transition = transitions.fade();
    const slideProps: SlideProps = {
      id: "intro",
      transition: fadeTransition,
    };

    const document = await expandDocument(
      <Deck width={960} height={540} theme={testTheme()}>
        <Slide {...slideProps}>
          <Text
            id="heading"
            style={{
              position: "absolute",
              left: 48,
              top: 48,
              width: 400,
              height: 48,
            }}
          >
            Initial title
          </Text>
          <Shape
            id="badge"
            style={{
              position: "absolute",
              left: 48,
              top: 120,
              width: 160,
              height: 48,
            }}
          >
            JSX badge
          </Shape>
        </Slide>
      </Deck>,
    );

    const introSlide = document.query.getAs(
      "intro",
      (node): node is SlideNode =>
        node instanceof CommonContainer && node.kind === "slide",
    );
    document.query
      .get("intro")
      .getAs(
        "badge",
        (node): node is RichTextContainer => node instanceof RichTextContainer,
      )
      .content("Edited badge")
      .style({ background: "#2563EB" });

    introSlide.append(
      shape("Added by vanilla", {
        id: "vanilla-box",
        style: {
          position: "absolute",
          left: 240,
          top: 120,
          width: 220,
          height: 48,
        },
      }),
    );

    document.append(
      slide({ id: "follow-up" }).append(
        text("Second slide from vanilla", {
          id: "follow-up-title",
          style: {
            position: "absolute",
            left: 48,
            top: 48,
            width: 480,
            height: 48,
          },
        }),
      ),
    );

    const result = await build(document);
    assert.equal(result.slides.length, 2);

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides.length, 2);
    assert.equal(file.slides[0]?.id, "intro");
    assert.equal(file.slides[0]?.transition?.kind, "fade");
    assert.deepEqual(
      file.slides[0]?.objects.map((object) => object.text?.text),
      ["Initial title", "Edited badge", "Added by vanilla"],
    );
    assert.equal(file.slides[1]?.id, "follow-up");
    assert.equal(
      file.slides[1]?.objects[0]?.text?.text,
      "Second slide from vanilla",
    );

    const directVanilla = await build(
      deck({ width: 960, height: 540, theme: testTheme() }).append(
        slide({ id: "only-vanilla" }).append(
          shape("Pure vanilla", {
            id: "v-shape",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 200,
              height: 60,
            },
          }),
        ),
      ),
    );
    const vanillaFile = await validatePptx(directVanilla.pptx);
    assert.deepEqual(vanillaFile.errors, []);
    assert.equal(vanillaFile.slides[0]?.objects[0]?.text?.text, "Pure vanilla");

    const directJsx = await build(
      <Deck width={960} height={540} theme={testTheme()}>
        <Slide id="only-jsx">
          <Shape
            id="j-shape"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 200,
              height: 60,
            }}
          >
            Pure JSX
          </Shape>
        </Slide>
      </Deck>,
    );
    const jsxFile = await validatePptx(directJsx.pptx);
    assert.deepEqual(jsxFile.errors, []);
    assert.equal(jsxFile.slides[0]?.objects[0]?.text?.text, "Pure JSX");
  });

  it("validates Slide id and numeric width at authoring time via tsc, and locates invalid slide nesting in resolve", async () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const tscBin = join(packageRoot, "node_modules/.bin/tsc");
    const runConsumerTsc = async (source: string) => {
      const dir = await mkdtemp(join(tmpdir(), "pptxdk-consumer-tsc-"));
      try {
        const nodeModulesDir = join(dir, "node_modules");
        await mkdir(join(nodeModulesDir, "@officedk"), { recursive: true });
        await mkdir(join(nodeModulesDir, "@types"), { recursive: true });
        await symlink(
          packageRoot,
          join(nodeModulesDir, "@officedk/pptxdk"),
          "dir",
        );
        await symlink(
          join(packageRoot, "node_modules/react"),
          join(nodeModulesDir, "react"),
          "dir",
        );
        await symlink(
          join(packageRoot, "node_modules/@types/react"),
          join(nodeModulesDir, "@types/react"),
          "dir",
        );
        await writeFile(
          join(dir, "tsconfig.json"),
          JSON.stringify({
            compilerOptions: {
              target: "es2024",
              module: "nodenext",
              moduleResolution: "nodenext",
              jsx: "react-jsx",
              jsxImportSource: "react",
              strict: true,
              noEmit: true,
              skipLibCheck: false,
            },
            include: ["deck.tsx"],
          }),
        );
        await writeFile(join(dir, "deck.tsx"), source);
        return spawnSync(tscBin, ["-p", "tsconfig.json", "--pretty", "false"], {
          cwd: dir,
          encoding: "utf8",
        });
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    };

    const validResult = await runConsumerTsc(`
import { transitions } from "@officedk/pptxdk";
import { Deck, Shape, Slide, type ShapeProps, type SlideProps } from "@officedk/pptxdk/react";
import { deck, shape, slide, type SlideNode } from "@officedk/pptxdk/vanilla";

const slideProps: SlideProps = { id: "cover", transition: transitions.fade() };
const shapeProps: ShapeProps = {
  id: "box",
  style: { position: "absolute", left: 10, top: 20, width: 200, height: 100 },
};

export const validDeck = (
  <Deck width={960} height={540}>
    <Slide {...slideProps}>
      <Shape {...shapeProps} />
    </Slide>
  </Deck>
);

const vanillaSlide: SlideNode = slide({ id: "v-cover", transition: transitions.fade() }).append(
  shape({ id: "v-box", style: { position: "absolute", left: 10, top: 20, width: 200, height: 100 } }),
);
export const validVanilla = deck({ width: 960, height: 540 }).append(vanillaSlide);
`);
    assert.equal(
      validResult.status,
      0,
      `Expected valid TSX to pass tsc:\n${validResult.stdout}\n${validResult.stderr}`,
    );

    const missingIdResult = await runConsumerTsc(`
import { Deck, Shape, Slide } from "@officedk/pptxdk/react";

export const missingIdDeck = (
  <Deck width={960} height={540}>
    <Slide>
      <Shape style={{ position: "absolute", left: 10, top: 20, width: 200, height: 100 }} />
    </Slide>
  </Deck>
);
`);
    assert.notEqual(missingIdResult.status, 0);
    assert.match(
      `${missingIdResult.stdout}\n${missingIdResult.stderr}`,
      /\bid\b/u,
    );

    const stringWidthResult = await runConsumerTsc(`
import { Deck, Shape, Slide } from "@officedk/pptxdk/react";

export const stringWidthDeck = (
  <Deck width={960} height={540}>
    <Slide id="cover">
      <Shape style={{ position: "absolute", left: 10, top: 20, width: "200", height: 100 }} />
    </Slide>
  </Deck>
);
`);
    assert.notEqual(stringWidthResult.status, 0);
    assert.match(
      `${stringWidthResult.stdout}\n${stringWidthResult.stderr}`,
      /width|string|number|LayoutLength/u,
    );

    let structuralError: unknown;
    try {
      await build(
        <Deck width={960} height={540} theme={testTheme()}>
          <Slide id="outer">
            <Slide id="nested">
              <Shape
                id="inner"
                style={{
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 40,
                }}
              />
            </Slide>
          </Slide>
        </Deck>,
      );
    } catch (error) {
      structuralError = error;
    }
    assert.ok(structuralError instanceof BuildError);
    assert.ok(
      structuralError.diagnostics.some(
        (diagnostic) =>
          diagnostic.path.includes("slide[outer]") &&
          diagnostic.path.includes("slide[nested]"),
      ),
      `Unexpected diagnostics: ${JSON.stringify(structuralError.diagnostics)}`,
    );
  }, 45_000);

  it("loads vanilla + default build without React or preview and expands JSX on demand across source and dist conditions", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const runSubprocess = (conditions: readonly string[], script: string) => {
      const child = spawnSync(
        process.execPath,
        [...conditions, "--input-type=module", "--eval", script],
        { cwd: packageRoot, encoding: "utf8" },
      );
      assert.equal(child.status, 0, child.stderr || child.stdout);
      return child.stdout;
    };

    const vanillaScript = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier === "@pptxdk/capabilities" ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("vanilla + build loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build, transitions } = await import("@officedk/pptxdk");
const { deck, shape, slide, text } = await import("@officedk/pptxdk/vanilla");
const doc = deck({ width: 960, height: 540 }).append(
  slide({ id: "s1", transition: transitions.fade() }).append(
    shape("Shape", {
      id: "sh1",
      style: { position: "absolute", left: 20, top: 20, width: 120, height: 60 },
    }),
    text("Hello", {
      id: "t1",
      style: { position: "absolute", left: 20, top: 100, width: 120, height: 40 },
    }),
  ),
);
const result = await build(doc);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX bytes");
}
process.stdout.write("ok");
`;

    const jsxOnDemandScript = `
import { registerHooks } from "node:module";

const loaded = new Set();
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("JSX build loaded preview dependency: " + specifier);
    }
    const resolved = nextResolve(specifier, context);
    if (specifier === "react" || specifier === "react-reconciler") {
      loaded.add(specifier);
    }
    if (specifier === "@pptxdk/react" || resolved.url.endsWith("/react/dist/index.js")) {
      loaded.add("@pptxdk/react");
    }
    return resolved;
  },
});

const { build } = await import("@officedk/pptxdk");
if (loaded.size !== 0) {
  throw new Error("Importing @officedk/pptxdk root eagerly loaded: " + [...loaded].join(", "));
}
const { createElement } = await import("react");
const { Deck, Shape, Slide } = await import("@officedk/pptxdk/react");
const result = await build(
  createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(
      Slide,
      { id: "jsx-s1" },
      createElement(
        Shape,
        {
          id: "jsx-sh1",
          style: { position: "absolute", left: 20, top: 20, width: 120, height: 60 },
        },
        "JSX",
      ),
    ),
  ),
);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX bytes");
}
if (!loaded.has("@pptxdk/react") || !loaded.has("react") || !loaded.has("react-reconciler")) {
  throw new Error("Expected JSX build to dynamically load React adapter: " + [...loaded].join(", "));
}
process.stdout.write("ok");
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.equal(runSubprocess(conditions, vanillaScript), "ok");
      assert.equal(runSubprocess(conditions, jsxOnDemandScript), "ok");
    }
  }, 45_000);
});
