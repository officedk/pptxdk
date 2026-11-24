import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { validatePptx, type ValidationResult } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  BuildError,
  CodegenError,
  build,
  codegen,
  importPptx,
  type CodegenResult,
} from "pptxdk";
import { buildFromSource } from "pptxdk/node";
import {
  Deck,
  Section,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  createMorphScope,
  defineSlideLayout,
  expandDocument,
  transitions,
} from "pptxdk/react";
import {
  deck,
  defineSlideLayout as vanillaDefineSlideLayout,
  group,
  image,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
} from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
} from "./helpers/codegen-acceptance.ts";

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

function createDirectVanillaSectionsAndTransitionsDeck() {
  const scopeA = createMorphScope();
  const scopeB = createMorphScope();
  const cardA = scopeA.handle("card");
  const cardB = scopeB.handle("card");
  const entering = scopeA.handle("entering-card");

  const inheritedLayout = vanillaDefineSlideLayout({
    id: "inherited-layout",
    name: "Inherited Layout",
    type: "blank",
  });
  const morphLayout = vanillaDefineSlideLayout({
    id: "morph-layout",
    name: "Morph Layout",
    type: "cust",
    transition: transitions.morph({
      duration: 1100,
      by: "word",
      fallback: transitions.push({ from: "left" }),
    }),
    advanceAfter: 3500,
  });

  return deck({ width: 960, height: 540 }).append(
    slideMaster({
      id: "brand-master",
      name: "Brand Master",
      transition: transitions.fade({ duration: 500 }),
      advanceAfter: 3000,
    }).append(slideLayout(inheritedLayout), slideLayout(morphLayout)),
    section({ name: "Empty" }),
    section({ name: "Morph Flow" }).append(
      slide({
        id: "s1-origin",
        layout: inheritedLayout,
        background: "#F8FAFC",
      }).append(
        shape({
          id: "hero-a",
          morphKey: cardA,
          x: 60,
          y: 80,
          width: 140,
          height: 100,
          fill: "#2563EB",
        }),
        shape({
          id: "isolated-a",
          morphKey: cardB,
          x: 60,
          y: 230,
          width: 120,
          height: 80,
          fill: "#7C3AED",
        }),
      ),
      slide({
        id: "s2-morph-inherited",
        layout: morphLayout,
        background: "#F1F5F9",
      }).append(
        shape({
          id: "hero-b",
          morphKey: cardA,
          x: 420,
          y: 120,
          width: 200,
          height: 120,
          fill: "#059669",
        }),
        shape({
          id: "isolated-b",
          morphKey: cardB,
          x: 420,
          y: 290,
          width: 150,
          height: 90,
          fill: "#DC2626",
        }),
        shape({
          id: "entering-only",
          morphKey: entering,
          x: 100,
          y: 320,
          width: 160,
          height: 70,
          fill: "#D97706",
        }),
      ),
    ),
    section({ name: "Overrides" }).append(
      slide({
        id: "s3-local-override",
        layout: morphLayout,
        background: "#0F172A",
        transition: transitions.flip({
          direction: "left",
          duration: 800,
          fallback: transitions.strips({ direction: "rightDown" }),
        }),
        advanceAfter: 1800,
      }).append(
        shape({
          id: "outro-flip",
          x: 200,
          y: 160,
          width: 220,
          height: 100,
          fill: "#F59E0B",
        }),
      ),
      slide({
        id: "s4-plain-fade",
        layout: inheritedLayout,
        background: "#FFFFFF",
        transition: transitions.fade(),
      }).append(
        shape({
          id: "outro-fade",
          x: 200,
          y: 160,
          width: 220,
          height: 100,
          fill: "#0EA5E9",
        }),
      ),
    ),
  );
}

function createDirectReactSectionsAndTransitionsElement() {
  const scopeA = createMorphScope();
  const scopeB = createMorphScope();
  const cardA = scopeA.handle("card");
  const cardB = scopeB.handle("card");
  const entering = scopeA.handle("entering-card");

  const inheritedLayout = defineSlideLayout({
    id: "inherited-layout",
    name: "Inherited Layout",
    type: "blank",
  });
  const morphLayout = defineSlideLayout({
    id: "morph-layout",
    name: "Morph Layout",
    type: "cust",
    transition: transitions.morph({
      duration: 1100,
      by: "word",
      fallback: transitions.push({ from: "left" }),
    }),
    advanceAfter: 3500,
  });

  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(
      SlideMaster,
      {
        id: "brand-master",
        name: "Brand Master",
        transition: transitions.fade({ duration: 500 }),
        advanceAfter: 3000,
      },
      createElement(SlideLayout, { layout: inheritedLayout }),
      createElement(SlideLayout, { layout: morphLayout }),
    ),
    createElement(Section, { name: "Empty" }),
    createElement(
      Section,
      { name: "Morph Flow" },
      createElement(
        Slide,
        {
          id: "s1-origin",
          layout: inheritedLayout,
          background: "#F8FAFC",
        },
        createElement(Shape, {
          id: "hero-a",
          morphKey: cardA,
          x: 60,
          y: 80,
          width: 140,
          height: 100,
          fill: "#2563EB",
        }),
        createElement(Shape, {
          id: "isolated-a",
          morphKey: cardB,
          x: 60,
          y: 230,
          width: 120,
          height: 80,
          fill: "#7C3AED",
        }),
      ),
      createElement(
        Slide,
        {
          id: "s2-morph-inherited",
          layout: morphLayout,
          background: "#F1F5F9",
        },
        createElement(Shape, {
          id: "hero-b",
          morphKey: cardA,
          x: 420,
          y: 120,
          width: 200,
          height: 120,
          fill: "#059669",
        }),
        createElement(Shape, {
          id: "isolated-b",
          morphKey: cardB,
          x: 420,
          y: 290,
          width: 150,
          height: 90,
          fill: "#DC2626",
        }),
        createElement(Shape, {
          id: "entering-only",
          morphKey: entering,
          x: 100,
          y: 320,
          width: 160,
          height: 70,
          fill: "#D97706",
        }),
      ),
    ),
    createElement(
      Section,
      { name: "Overrides" },
      createElement(
        Slide,
        {
          id: "s3-local-override",
          layout: morphLayout,
          background: "#0F172A",
          transition: transitions.flip({
            direction: "left",
            duration: 800,
            fallback: transitions.strips({ direction: "rightDown" }),
          }),
          advanceAfter: 1800,
        },
        createElement(Shape, {
          id: "outro-flip",
          x: 200,
          y: 160,
          width: 220,
          height: 100,
          fill: "#F59E0B",
        }),
      ),
      createElement(
        Slide,
        {
          id: "s4-plain-fade",
          layout: inheritedLayout,
          background: "#FFFFFF",
          transition: transitions.fade(),
        },
        createElement(Shape, {
          id: "outro-fade",
          x: 200,
          y: 160,
          width: 220,
          height: 100,
          fill: "#0EA5E9",
        }),
      ),
    ),
  );
}

function assertSectionsTransitionsAndMorphFile(
  pptx: Uint8Array,
  file: ValidationResult,
  label: string,
): void {
  assert.deepEqual(file.errors, [], `${label} validation errors`);

  // Sections order, empty section, and intra-section slide assignment
  assert.equal(file.presentation.sections?.length, 3, `${label} section count`);
  assert.equal(file.presentation.sections?.[0]?.name, "Empty", label);
  assert.deepEqual(
    file.presentation.sections?.[0]?.slideIds,
    [],
    `${label} empty section`,
  );
  assert.equal(file.presentation.sections?.[1]?.name, "Morph Flow", label);
  assert.equal(file.presentation.sections?.[1]?.slideIds.length, 2, label);
  assert.equal(file.presentation.sections?.[2]?.name, "Overrides", label);
  assert.equal(file.presentation.sections?.[2]?.slideIds.length, 2, label);

  assert.equal(file.slides.length, 4, `${label} slide count`);
  assert.equal(file.slides[0]?.section, "Morph Flow", label);
  assert.equal(file.slides[1]?.section, "Morph Flow", label);
  assert.equal(file.slides[2]?.section, "Overrides", label);
  assert.equal(file.slides[3]?.section, "Overrides", label);
  assert.deepEqual(
    file.presentation.sections?.[1]?.slideIds,
    [file.slides[0]?.id, file.slides[1]?.id],
    `${label} Morph Flow slide ordering`,
  );
  assert.deepEqual(
    file.presentation.sections?.[2]?.slideIds,
    [file.slides[2]?.id, file.slides[3]?.id],
    `${label} Overrides slide ordering`,
  );

  // Three-layer transition ownership (master -> layout -> slide)
  assert.equal(file.masters?.[0]?.advanceAfterMs, 3000, label);
  assert.deepEqual(
    file.masters?.[0]?.transition,
    {
      kind: "fade",
      speed: "fast",
      durationMs: 500,
      fallback: { kind: "fade", speed: "fast" },
    },
    `${label} master transition`,
  );
  assert.equal(file.masters?.[0]?.layouts[0]?.transition, undefined, label);
  assert.equal(file.masters?.[0]?.layouts[0]?.advanceAfterMs, undefined, label);
  assert.equal(file.masters?.[0]?.layouts[1]?.advanceAfterMs, 3500, label);
  assert.deepEqual(
    file.masters?.[0]?.layouts[1]?.transition,
    {
      kind: "morph",
      speed: "slow",
      durationMs: 1100,
      by: "word",
      fallback: { kind: "push", speed: "slow", direction: "l" },
    },
    `${label} morph layout transition`,
  );

  // Inherited slides omit local transition declarations
  assert.equal(file.slides[0]?.transition, undefined, `${label} s1 inherited`);
  assert.equal(file.slides[0]?.advanceAfterMs, undefined, label);
  assert.equal(file.slides[1]?.transition, undefined, `${label} s2 inherited`);
  assert.equal(file.slides[1]?.advanceAfterMs, undefined, label);

  // Local override slide retains explicit transition & advanceAfter
  assert.equal(file.slides[2]?.advanceAfterMs, 1800, label);
  assert.deepEqual(
    file.slides[2]?.transition,
    {
      kind: "flip",
      speed: "med",
      durationMs: 800,
      direction: "l",
      fallback: { kind: "strips", speed: "med", direction: "rd" },
    },
    `${label} s3 flip override`,
  );
  assert.deepEqual(
    file.slides[3]?.transition,
    { kind: "fade", speed: "slow" },
    `${label} s4 plain fade`,
  );

  // Verify mc:AlternateContent Choice/Fallback XML branches and p14:sectionLst
  const archive = unzipSync(pptx);
  const presXml = strFromU8(archive["ppt/presentation.xml"]!);
  assert.match(presXml, /<p14:sectionLst\b/u, `${label} sectionLst XML`);

  const masterXml = strFromU8(archive["ppt/slideMasters/slideMaster1.xml"]!);
  assert.match(
    masterXml,
    /<mc:Choice Requires="p14">.*p14:dur="500".*<mc:Fallback><p:transition spd="fast" advTm="3000"><p:fade\/><\/p:transition><\/mc:Fallback>/u,
    `${label} master mc:AlternateContent`,
  );

  const layout2Xml = strFromU8(archive["ppt/slideLayouts/slideLayout2.xml"]!);
  assert.match(
    layout2Xml,
    /<mc:Choice Requires="p14 p159">.*p14:dur="1100".*<p159:morph option="byWord"\/>.*<mc:Fallback><p:transition spd="slow" advTm="3500"><p:push dir="l"\/><\/p:transition><\/mc:Fallback>/u,
    `${label} layout2 mc:AlternateContent`,
  );

  const slide1Xml = strFromU8(archive["ppt/slides/slide1.xml"]!);
  assert.doesNotMatch(slide1Xml, /<p:transition/u, `${label} s1 no transition`);

  const slide3Xml = strFromU8(archive["ppt/slides/slide3.xml"]!);
  assert.match(
    slide3Xml,
    /<mc:Choice Requires="p14">.*p14:dur="800".*<p14:flip dir="l"\/>.*<mc:Fallback><p:transition spd="med" advTm="1800"><p:strips dir="rd"\/><\/p:transition><\/mc:Fallback>/u,
    `${label} s3 mc:AlternateContent`,
  );

  // Cross-slide Morph pairing & MorphScope isolation + entering-only object
  const heroAName = file.slides[0]?.objects[0]?.name;
  const isolatedAName = file.slides[0]?.objects[1]?.name;
  const heroBName = file.slides[1]?.objects[0]?.name;
  const isolatedBName = file.slides[1]?.objects[1]?.name;
  const enteringName = file.slides[1]?.objects[2]?.name;

  assert.ok(heroAName?.startsWith("!!morph-"), `${label} heroA morph prefix`);
  assert.equal(heroBName, heroAName, `${label} hero pair matches`);
  assert.ok(
    isolatedAName?.startsWith("!!morph-"),
    `${label} isolatedA morph prefix`,
  );
  assert.equal(isolatedBName, isolatedAName, `${label} isolated pair matches`);
  assert.notEqual(
    heroAName,
    isolatedAName,
    `${label} distinct MorphScopes with identical label remain isolated`,
  );
  assert.ok(
    enteringName?.startsWith("!!morph-"),
    `${label} entering-only morph prefix`,
  );
  assert.notEqual(enteringName, heroAName, label);
  assert.notEqual(enteringName, isolatedAName, label);

  const heroABox = file.slides[0]?.objects[0]?.boxEmu;
  const heroBBox = file.slides[1]?.objects[0]?.boxEmu;
  assert.deepEqual(
    heroABox,
    { x: 60 * 9525, y: 80 * 9525, width: 140 * 9525, height: 100 * 9525 },
    `${label} heroA boxEmu`,
  );
  assert.deepEqual(
    heroBBox,
    { x: 420 * 9525, y: 120 * 9525, width: 200 * 9525, height: 120 * 9525 },
    `${label} heroB moved/resized boxEmu`,
  );
  assert.match(slide1Xml, /<a:srgbClr val="2563EB"\/>/u, label);
  const slide2Xml = strFromU8(archive["ppt/slides/slide2.xml"]!);
  assert.match(slide2Xml, /<a:srgbClr val="059669"\/>/u, label);
}

describe("codegen sections, cross-slide morph, and multi-layer transitions acceptance", () => {
  it("rebuilds sections, empty sections, MorphScope isolation, and three-layer transitions for direct and imported sources across vanilla and react", async () => {
    const vanillaDoc = createDirectVanillaSectionsAndTransitionsDeck();
    const reactDoc = await expandDocument(
      createDirectReactSectionsAndTransitionsElement(),
    );

    // Single formal build + validation reused for the baseline direct scenario and import input.
    const vanillaBuilt = await build(vanillaDoc);
    const reactBuilt = await build(reactDoc);
    const baselineValidation = await validatePptx(vanillaBuilt.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assertSectionsTransitionsAndMorphFile(
      vanillaBuilt.pptx,
      baselineValidation,
      "direct-vanilla-baseline",
    );
    const reactBaselineValidation = await validatePptx(reactBuilt.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assertSectionsTransitionsAndMorphFile(
      reactBuilt.pptx,
      reactBaselineValidation,
      "direct-react-baseline",
    );

    // Import PPTX and zero out caller bytes to guarantee zero retention of original archive bytes.
    const importBytes = new Uint8Array(vanillaBuilt.pptx);
    const imported = importPptx(importBytes);
    importBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const directVanillaResult = codegen(vanillaDoc, { target: "vanilla" });
    const directVanillaRepeat = codegen(vanillaDoc, { target: "vanilla" });
    assert.deepEqual(directVanillaRepeat, directVanillaResult);

    const directReactResult = codegen(reactDoc, { target: "react" });
    const importedVanillaResult = codegen(imported.document, {
      target: "vanilla",
    });
    const importedReactResult = codegen(imported.document, { target: "react" });
    const importedReactRepeat = codegen(imported.document, { target: "react" });
    assert.deepEqual(importedReactRepeat, importedReactResult);

    // Direct source preserves distinct MorphScope buckets (morphScope & morphScope2) for same-label handles.
    for (const directRes of [directVanillaResult, directReactResult]) {
      const source = directRes.files[directRes.entry] as string;
      assert.match(source, /createMorphScope\(\)/u);
      assert.match(source, /morphScope\.handle\("card"\)/u);
      assert.match(source, /morphScope2\.handle\("card"\)/u);
      assert.match(source, /morphScope\.handle\("entering-card"\)/u);
      assert.match(source, /transitions\.fade\(\)/u);
    }

    // Batch strict tsc --noEmit check against public consumer SDK types
    await verifyGeneratedSourcesTypecheck([
      { label: "direct-vanilla", result: directVanillaResult },
      { label: "direct-react", result: directReactResult },
      { label: "imported-vanilla", result: importedVanillaResult },
      { label: "imported-react", result: importedReactResult },
    ]);

    // Relocated buildFromSource rebuilds and independent validatePptx checks
    const rebuiltDirectVanilla =
      await rebuildGeneratedSourceRelocated(directVanillaResult);
    assertSectionsTransitionsAndMorphFile(
      rebuiltDirectVanilla.built.pptx,
      rebuiltDirectVanilla.file,
      "rebuilt-direct-vanilla",
    );

    const rebuiltDirectReact =
      await rebuildGeneratedSourceRelocated(directReactResult);
    assertSectionsTransitionsAndMorphFile(
      rebuiltDirectReact.built.pptx,
      rebuiltDirectReact.file,
      "rebuilt-direct-react",
    );

    const rebuiltImportedVanilla = await rebuildGeneratedSourceRelocated(
      importedVanillaResult,
    );
    assertSectionsTransitionsAndMorphFile(
      rebuiltImportedVanilla.built.pptx,
      rebuiltImportedVanilla.file,
      "rebuilt-imported-vanilla",
    );

    const rebuiltImportedReact =
      await rebuildGeneratedSourceRelocated(importedReactResult);
    assertSectionsTransitionsAndMorphFile(
      rebuiltImportedReact.built.pptx,
      rebuiltImportedReact.file,
      "rebuilt-imported-react",
    );
  }, 60_000);

  it("propagates edited layout transitions to inherited slides while preserving local slide transition overrides", async () => {
    const sharedLayout = vanillaDefineSlideLayout({
      id: "shared-layout",
      name: "Shared Layout",
      type: "blank",
      transition: transitions.fade({ duration: 500 }),
      advanceAfter: 2000,
    });
    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1", name: "Master" }).append(
        slideLayout(sharedLayout),
      ),
      slide({ id: "s-inherited", layout: sharedLayout }).append(
        shape({ id: "a", x: 40, y: 40, width: 120, height: 60 }),
      ),
      slide({
        id: "s-local",
        layout: sharedLayout,
        transition: transitions.push({ from: "right", duration: 400 }),
        advanceAfter: 1200,
      }).append(shape({ id: "b", x: 40, y: 40, width: 120, height: 60 })),
    );

    for (const target of ["vanilla", "react"] as const) {
      const generated: CodegenResult = codegen(doc, { target });
      const workDir = await mkdtemp(join(tmpdir(), "pptxdk-edit-layout-tr-"));
      try {
        await writeCodegenFiles(workDir, generated.files);
        const entryPath = join(workDir, generated.entry);
        const originalSource = await readFile(entryPath, "utf8");

        const editedSource = originalSource
          .replace(
            /transition:\s*transitions\.fade\(\{\s*duration:\s*500\s*\}\)/u,
            'transition: transitions.flip({ direction: "left", duration: 900, fallback: transitions.strips({ direction: "leftUp" }) })',
          )
          .replace(/advanceAfter:\s*2000/u, "advanceAfter: 4500");
        assert.notEqual(
          editedSource,
          originalSource,
          `${target} layout transition edit must match generated source`,
        );

        await writeFile(entryPath, editedSource, "utf8");
        const rebuilt = await buildFromSource(entryPath);
        const validated = await validatePptx(rebuilt.pptx, {
          allowedExtensions: ["p14", "p159"],
        });
        assert.deepEqual(validated.errors, []);

        // Layout transition updated and inherited slide still omits local override
        assert.equal(validated.masters?.[0]?.layouts[0]?.advanceAfterMs, 4500);
        assert.deepEqual(validated.masters?.[0]?.layouts[0]?.transition, {
          kind: "flip",
          speed: "slow",
          durationMs: 900,
          direction: "l",
          fallback: { kind: "strips", speed: "slow", direction: "lu" },
        });
        assert.equal(validated.slides[0]?.transition, undefined);
        assert.equal(validated.slides[0]?.advanceAfterMs, undefined);

        // Local override slide kept its original push transition & advanceAfter
        assert.equal(validated.slides[1]?.advanceAfterMs, 1200);
        assert.deepEqual(validated.slides[1]?.transition, {
          kind: "push",
          speed: "fast",
          durationMs: 400,
          direction: "r",
          fallback: { kind: "push", speed: "fast", direction: "r" },
        });
      } finally {
        await rm(workDir, { recursive: true, force: true });
      }
    }
  }, 60_000);

  it("rejects invalid Morph pairings with slide/object diagnostics while keeping healthy Morph pairings intact", async () => {
    const scope = createMorphScope();
    const healthyKey = scope.handle("healthy");
    const mismatchKey = scope.handle("mismatch");
    const groupedKey = scope.handle("grouped");

    // 1. Cross-kind Morph pairing mismatch + healthy pair in the same deck
    const mismatchDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1-src" }).append(
        shape({
          id: "healthy-a",
          morphKey: healthyKey,
          x: 40,
          y: 40,
          width: 100,
          height: 60,
          fill: "#2563EB",
        }),
        shape({
          id: "mismatch-a",
          morphKey: mismatchKey,
          x: 200,
          y: 40,
          width: 100,
          height: 60,
          fill: "#DC2626",
        }),
      ),
      slide({
        id: "s2-dst",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(
        shape({
          id: "healthy-b",
          morphKey: healthyKey,
          x: 320,
          y: 80,
          width: 140,
          height: 80,
          fill: "#059669",
        }),
        image({
          id: "mismatch-b",
          morphKey: mismatchKey,
          src: png1x1,
          x: 500,
          y: 80,
          width: 100,
          height: 60,
        }),
      ),
    );

    // 2. Morph inside Group + healthy pair in the same deck
    const groupedDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1-grp" }).append(
        shape({
          id: "healthy-a",
          morphKey: healthyKey,
          x: 40,
          y: 40,
          width: 100,
          height: 60,
          fill: "#2563EB",
        }),
        group({ id: "bad-group" }).append(
          shape({
            id: "bad-grouped-shape",
            morphKey: groupedKey,
            x: 200,
            y: 40,
            width: 80,
            height: 60,
          }),
        ),
      ),
      slide({
        id: "s2-grp",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(
        shape({
          id: "healthy-b",
          morphKey: healthyKey,
          x: 320,
          y: 80,
          width: 140,
          height: 80,
          fill: "#059669",
        }),
      ),
    );

    for (const target of ["vanilla", "react"] as const) {
      const mismatchRes = codegen(mismatchDoc, { target });
      const groupedRes = codegen(groupedDoc, { target });

      const workDir = await mkdtemp(join(tmpdir(), "pptxdk-morph-reject-"));
      try {
        const mismatchDir = join(workDir, "mismatch");
        await writeCodegenFiles(mismatchDir, mismatchRes.files);
        await assert.rejects(
          buildFromSource(join(mismatchDir, mismatchRes.entry)),
          (error: unknown) => {
            assert.ok(error instanceof BuildError);
            const diag = error.diagnostics.find(
              (item) => item.code === "morph-object-kind-mismatch",
            );
            assert.ok(diag !== undefined, `${target} mismatch diagnostic`);
            assert.equal(diag.path, "deck/slide[s2-dst]/image[mismatch-b]");
            return true;
          },
        );

        // Removing only the mismatched image leaves the healthy Morph pair building cleanly
        const mismatchSource = await readFile(
          join(mismatchDir, mismatchRes.entry),
          "utf8",
        );
        const fixedMismatchSource =
          target === "vanilla"
            ? mismatchSource.replace(
                /,?\s*image\(\s*\{[\s\S]*?id:\s*"mismatch-b"[\s\S]*?\},?\s*\)/u,
                "",
              )
            : mismatchSource.replace(
                /\s*<Image\b[\s\S]*?id="mismatch-b"[\s\S]*?\/>/u,
                "",
              );
        assert.notEqual(fixedMismatchSource, mismatchSource);
        const healedDir = join(workDir, "healed");
        await writeCodegenFiles(healedDir, mismatchRes.files);
        await writeFile(
          join(healedDir, mismatchRes.entry),
          fixedMismatchSource,
          "utf8",
        );
        const healed = await buildFromSource(
          join(healedDir, mismatchRes.entry),
        );
        const healedValidation = await validatePptx(healed.pptx, {
          allowedExtensions: ["p14", "p159"],
        });
        assert.deepEqual(healedValidation.errors, []);
        assert.ok(
          healedValidation.slides[0]?.objects[0]?.name?.startsWith("!!morph-"),
        );
        assert.equal(
          healedValidation.slides[0]?.objects[0]?.name,
          healedValidation.slides[1]?.objects[0]?.name,
        );

        const groupedDir = join(workDir, "grouped");
        await writeCodegenFiles(groupedDir, groupedRes.files);
        await assert.rejects(
          buildFromSource(join(groupedDir, groupedRes.entry)),
          (error: unknown) => {
            assert.ok(error instanceof BuildError);
            const diag = error.diagnostics.find(
              (item) => item.code === "morph-inside-group",
            );
            assert.ok(diag !== undefined, `${target} morph-inside-group`);
            assert.equal(
              diag.path,
              "deck/slide[s1-grp]/group[bad-group]/shape[bad-grouped-shape]",
            );
            return true;
          },
        );
      } finally {
        await rm(workDir, { recursive: true, force: true });
      }
    }

    // Invalid transition / advanceAfter fails codegen with exact path while healthy sibling succeeds
    const invalidTransitionDoc = deck().append(
      slide({
        id: "bad-tr",
        transition: { type: "transition", kind: "unknown" } as never,
      }),
    );
    assert.throws(
      () => codegen(invalidTransitionDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[bad-tr]@transition",
    );
  }, 60_000);
});
