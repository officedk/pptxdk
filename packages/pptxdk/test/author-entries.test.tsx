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
  fromFontCollection,
  RichTextContainer,
  transitions,
  type Transition,
} from "@officedk/pptxdk";
import {
  Deck,
  expandDocument,
  fromFontCollection as reactFromFontCollection,
  Shape,
  Slide,
  Text,
  type SlideProps,
} from "@officedk/pptxdk/react";
import {
  deck,
  fromFontCollection as vanillaFromFontCollection,
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

    assert.equal(fromFontCollection, reactFromFontCollection);
    assert.equal(fromFontCollection, vanillaFromFontCollection);
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
});
