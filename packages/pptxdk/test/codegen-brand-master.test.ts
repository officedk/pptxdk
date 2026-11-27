import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import {
  DeckContainer,
  build,
  codegen,
  defineTheme,
  importPptx,
  lighter,
  type CodegenResult,
} from "@officedk/pptxdk";
import {
  Deck,
  Image,
  Slide,
  SlideLayout,
  SlideMaster,
  Text,
  expandDocument,
  font as reactFont,
} from "@officedk/pptxdk/react";
import {
  deck,
  defineSlideLayout,
  font as vanillaFont,
  image,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@officedk/pptxdk/vanilla";
import { unzipSync, strFromU8 } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

const BRAND_FIXTURE_SHA256 =
  "de3376969353a2322ccce48182393f63eea9a44f5a231565985bcb5d3d336536";

const jostRegularBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

function createBrandBadgePng(): Uint8Array {
  const png = new PNG({ width: 32, height: 16 });
  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      png.data.set(
        x < 16 ? [29, 78, 216, 255] : [15, 23, 42, 255],
        (y * png.width + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(png));
}

function buildDirectBrandMasterFixtures() {
  const logoPng = createBrandBadgePng();
  const vFont = vanillaFont("Jost*", [{ weight: 400, data: jostRegularBytes }]);
  const rFont = reactFont("Jost*", [{ weight: 400, data: jostRegularBytes }]);

  const makeTheme = (fontValue: typeof vFont) =>
    defineTheme({
      name: "Apex Brand",
      colors: {
        accent1: "#1D4ED8",
        brandTint: lighter("accent1", 0.35),
      },
      fonts: { major: fontValue, minor: fontValue },
    });

  const vTheme = makeTheme(vFont);
  const rTheme = makeTheme(rFont);

  const heroLayout = defineSlideLayout({
    id: "brand-hero",
    name: "Brand Hero",
    type: "title",
    preserve: true,
    headerFooter: { dateTime: false, footer: true, slideNumber: true },
    footerPlaceholders: {
      footer: { text: "Apex Confidential", style: { fontSize: "10pt" } },
    },
    placeholders: {
      title: {
        type: "ctrTitle",
        style: {
          position: "absolute",
          left: 64,
          top: 88,
          width: 832,
          height: 80,
          fontSize: "32pt",
          color: "#0F172A",
        },
      },
      subtitle: {
        type: "subTitle",
        idx: 1,
        style: {
          position: "absolute",
          left: 64,
          top: 192,
          width: 832,
          height: 120,
        },
      },
    },
  });

  const idleContentLayout = defineSlideLayout({
    id: "brand-idle-content",
    name: "Brand Idle Content",
    type: "obj",
    preserve: true,
    headerFooter: { dateTime: false, footer: true, slideNumber: true },
    placeholders: {
      title: {
        type: "title",
        style: {
          position: "absolute",
          left: 64,
          top: 56,
          width: 832,
          height: 64,
          fontSize: "26pt",
          color: "#0F172A",
        },
      },
      content: {
        type: "obj",
        explicitType: false,
        idx: 1,
        style: {
          position: "absolute",
          left: 64,
          top: 144,
          width: 832,
          height: 296,
        },
      },
    },
  });

  const vanillaDoc = deck({ width: 960, height: 540, theme: vTheme }).append(
    slideMaster({
      id: "brand-master",
      name: "Brand Master",
      preserve: true,
      theme: vTheme,
      background: "#F8FAFC",
      headerFooter: { dateTime: false, footer: true, slideNumber: true },
      footerPlaceholders: {
        footer: { text: "Apex Confidential" },
        slideNumber: { idx: 4 },
      },
    }).append(
      image({
        src: logoPng,
        id: "brand-logo",
        name: "Brand Logo",
        style: {
          position: "absolute",
          left: 24,
          top: 16,
          width: 64,
          height: 32,
        },
      }),
      text("APEX BRAND", {
        id: "brand-wordmark",
        name: "Brand Wordmark",
        style: {
          position: "absolute",
          left: 96,
          top: 16,
          width: 160,
          height: 32,
          fontSize: "12pt",
          color: "#1D4ED8",
        },
      }),
      slideLayout(heroLayout),
      slideLayout(idleContentLayout),
    ),
    slide({
      id: "inherited-slide",
      layout: heroLayout,
      footer: "Apex Confidential",
      slideNumber: true,
    }).append(
      text("Inherited Brand Title", { placeholder: "title" }),
      text("Layout-driven size", { placeholder: "subtitle" }),
    ),
    slide({
      id: "explicit-equal-slide",
      layout: heroLayout,
      footer: "Apex Confidential",
      slideNumber: true,
    }).append(
      text("Locked Brand Title", {
        placeholder: "title",
        style: { fontSize: "32pt" },
      }),
      text("Equal explicit size", { placeholder: "subtitle" }),
    ),
  );

  const reactElement = createElement(
    Deck,
    { width: 960, height: 540, theme: rTheme },
    createElement(
      SlideMaster,
      {
        id: "brand-master",
        name: "Brand Master",
        preserve: true,
        theme: rTheme,
        background: "#F8FAFC",
        headerFooter: { dateTime: false, footer: true, slideNumber: true },
        footerPlaceholders: {
          footer: { text: "Apex Confidential" },
          slideNumber: { idx: 4 },
        },
      },
      createElement(Image, {
        id: "brand-logo",
        name: "Brand Logo",
        src: logoPng,
        style: {
          position: "absolute",
          left: 24,
          top: 16,
          width: 64,
          height: 32,
        },
      }),
      createElement(
        Text,
        {
          id: "brand-wordmark",
          name: "Brand Wordmark",
          style: {
            position: "absolute",
            left: 96,
            top: 16,
            width: 160,
            height: 32,
            fontSize: "12pt",
            color: "#1D4ED8",
          },
        },
        "APEX BRAND",
      ),
      createElement(SlideLayout, { layout: heroLayout }),
      createElement(SlideLayout, { layout: idleContentLayout }),
    ),
    createElement(
      Slide,
      {
        id: "inherited-slide",
        layout: heroLayout,
        footer: "Apex Confidential",
        slideNumber: true,
      },
      createElement(Text, { placeholder: "title" }, "Inherited Brand Title"),
      createElement(Text, { placeholder: "subtitle" }, "Layout-driven size"),
    ),
    createElement(
      Slide,
      {
        id: "explicit-equal-slide",
        layout: heroLayout,
        footer: "Apex Confidential",
        slideNumber: true,
      },
      createElement(
        Text,
        { placeholder: "title", style: { fontSize: "32pt" } },
        "Locked Brand Title",
      ),
      createElement(Text, { placeholder: "subtitle" }, "Equal explicit size"),
    ),
  );

  return { vanillaDoc, reactElement };
}

function readEntrySource(result: CodegenResult): string {
  const source = result.files[result.entry];
  if (typeof source !== "string") {
    throw new Error(`Expected string entry file ${result.entry}`);
  }
  return source;
}

/**
 * Extracts the top two sections (imports, asset declarations, fonts, themes, layouts, and
 * brand master) before slide declarations from a generated main file and appends a new slide
 * referencing the previously idle layout and its named placeholders.
 */
function buildCopiedBrandMasterDeck(
  generated: CodegenResult,
  options: {
    readonly idleLayoutVar: string;
    readonly masterRef: string;
    readonly themeVar: string;
    readonly titlePlaceholder: string;
    readonly bodyPlaceholder: string;
  },
): CodegenResult {
  const textSrc = readEntrySource(generated);

  const slideBoundary =
    generated.target === "vanilla"
      ? textSrc.search(/\nconst \w+Slide = slide\(/u)
      : textSrc.search(/\nfunction \w+Slide\(/u);
  assert.ok(
    slideBoundary > 0,
    "Expected slide declarations after master section",
  );

  const prefix = textSrc.slice(0, slideBoundary);
  assert.ok(
    prefix.includes(options.idleLayoutVar),
    `Extracted top prefix missing idle layout ${options.idleLayoutVar}`,
  );

  const appendedSource =
    generated.target === "vanilla"
      ? `${prefix}

const continuedSlide = slide({
  id: "continued-from-idle",
  layout: ${options.idleLayoutVar},
  footer: "Continued Brand Deck",
  slideNumber: true,
}).append(
  text("Continued From Copied Master", { placeholder: ${JSON.stringify(options.titlePlaceholder)} }),
  text("Created using previously idle layout", { placeholder: ${JSON.stringify(options.bodyPlaceholder)} }),
);

export default deck({ width: 960, height: 540, theme: ${options.themeVar} }).append(
  ${options.masterRef},
  continuedSlide,
);
`
      : `${prefix}

function ContinuedSlide() {
  return (
    <Slide
      id="continued-from-idle"
      layout={${options.idleLayoutVar}}
      footer="Continued Brand Deck"
      slideNumber={true}
    >
      <Text placeholder=${JSON.stringify(options.titlePlaceholder)}>
        Continued From Copied Master
      </Text>
      <Text placeholder=${JSON.stringify(options.bodyPlaceholder)}>
        Created using previously idle layout
      </Text>
    </Slide>
  );
}

export default function GeneratedDeck() {
  return (
    <Deck width={960} height={540} theme={${options.themeVar}}>
      ${options.masterRef}
      <ContinuedSlide />
    </Deck>
  );
}
`;

  return {
    target: generated.target,
    entry: generated.entry,
    files: {
      ...generated.files,
      [generated.entry]: appendedSource,
    },
  };
}

function readBrandFixtureBytes(): Uint8Array {
  const raw = new Uint8Array(
    readFileSync(
      new URL(
        "./fixtures/brand-master-inheritance-override.pptx",
        import.meta.url,
      ),
    ),
  );
  assert.equal(
    createHash("sha256").update(raw).digest("hex"),
    BRAND_FIXTURE_SHA256,
  );
  return raw;
}

describe("pptxdk editable brand master codegen acceptance", () => {
  it("rebuilds vanilla and react sources for direct (with embedded image/font assets) and imported PowerPoint brand masters", async () => {
    const { vanillaDoc, reactElement } = buildDirectBrandMasterFixtures();
    const reactDoc = await expandDocument(reactElement);

    const builtDirect = await build(vanillaDoc, { embedFonts: "subset" });
    const builtDirectBytes = new Uint8Array(builtDirect.pptx);
    const importedDirect = importPptx(builtDirectBytes);
    builtDirectBytes.fill(0);
    assert.deepEqual(importedDirect.report.diagnostics, []);

    const fixtureRaw = readBrandFixtureBytes();
    const importedFixture = importPptx(fixtureRaw);
    fixtureRaw.fill(0);
    assert.equal(importedFixture.report.slideCount, 2);
    assert.ok(importedFixture.document instanceof DeckContainer);
    assert.equal(importedFixture.document.masters.length, 1);

    const cases = [
      { label: "direct-vanilla", doc: vanillaDoc, kind: "direct" as const },
      { label: "direct-react", doc: reactDoc, kind: "direct" as const },
      {
        label: "imported-direct-assets",
        doc: importedDirect.document,
        kind: "direct" as const,
      },
      {
        label: "imported-brand-fixture",
        doc: importedFixture.document,
        kind: "imported" as const,
      },
    ].flatMap(({ label, doc, kind }) => {
      const vanilla = codegen(doc, { target: "vanilla" });
      const react = codegen(doc, { target: "react" });
      return [
        { label: `${label}:vanilla`, kind, result: vanilla },
        { label: `${label}:react`, kind, result: react },
      ];
    });

    await verifyGeneratedSourcesTypecheck(cases);

    for (const item of cases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
        item.kind === "direct"
          ? { buildOptions: { embedFonts: "subset" } }
          : {},
      );
      assert.deepEqual(file.errors, [], `${item.label} validation errors`);
      assert.equal(file.masters?.length, 1, `${item.label} master count`);
      assert.equal(
        file.masters?.[0]?.layouts.length,
        2,
        `${item.label} layout count`,
      );
      assert.equal(file.slides.length, 2, `${item.label} slide count`);

      // Native layout/master bindings on slides
      assert.equal(
        file.slides[0]?.layoutPart,
        file.masters?.[0]?.layouts[0]?.part,
      );
      assert.equal(file.slides[0]?.masterPart, file.masters?.[0]?.part);
      assert.equal(
        file.slides[1]?.layoutPart,
        file.masters?.[0]?.layouts[0]?.part,
      );

      // Master background & shared objects stay on master, never flattened onto slides
      const zip = unzipSync(built.pptx);
      const masterXml = strFromU8(zip["ppt/slideMasters/slideMaster1.xml"]!);
      assert.match(masterXml, /F8FAFC/u, `${item.label} master background`);

      const masterObjectNames = (file.masters?.[0]?.objects ?? []).map(
        (obj) => obj.name,
      );
      if (item.kind === "direct") {
        assert.equal(file.masters?.[0]?.theme?.name, "Apex Brand");
        assert.deepEqual(masterObjectNames, ["Brand Logo", "Brand Wordmark"]);
        assert.ok(
          Object.keys(zip).some((name) => name.startsWith("ppt/media/")),
          `${item.label} expected embedded media image`,
        );
        assert.ok(
          Object.keys(zip).some((name) => name.startsWith("ppt/fonts/")),
          `${item.label} expected embedded font`,
        );
      } else {
        assert.equal(file.masters?.[0]?.theme?.fonts.major, "Arial");
        assert.deepEqual(masterObjectNames, ["BrandBadge"]);
      }

      for (const [slideIndex, s] of file.slides.entries()) {
        assert.equal(
          s.objects.some(
            (obj) =>
              obj.name === "Brand Logo" ||
              obj.name === "Brand Wordmark" ||
              obj.name === "BrandBadge",
          ),
          false,
          `${item.label} shared master object must not flatten onto slide ${slideIndex + 1}`,
        );
        assert.equal(
          s.footerPlaceholders?.some((ph) => ph.kind === "ftr"),
          true,
          `${item.label} slide ${slideIndex + 1} footer preserved`,
        );
        // Named placeholders retain <p:ph> bindings
        assert.ok(s.objects[0]?.placeholder !== undefined);
        assert.ok(s.objects[1]?.placeholder !== undefined);
      }
    }
  }, 90_000);

  it("copies brand master definitions and assets from the top of generated TS/TSX files to author a new slide using the idle layout", async () => {
    const { vanillaDoc } = buildDirectBrandMasterFixtures();
    const fixtureRaw = readBrandFixtureBytes();
    const importedFixture = importPptx(fixtureRaw);
    fixtureRaw.fill(0);

    const directVanillaCopied = buildCopiedBrandMasterDeck(
      codegen(vanillaDoc, { target: "vanilla" }),
      {
        idleLayoutVar: "brandMasterBrandIdleContentLayout",
        masterRef: "brandMaster",
        themeVar: "apexBrandTheme",
        titlePlaceholder: "title",
        bodyPlaceholder: "content",
      },
    );
    const directReactCopied = buildCopiedBrandMasterDeck(
      codegen(vanillaDoc, { target: "react" }),
      {
        idleLayoutVar: "brandMasterBrandIdleContentLayout",
        masterRef: "<BrandMaster />",
        themeVar: "apexBrandTheme",
        titlePlaceholder: "title",
        bodyPlaceholder: "content",
      },
    );

    const importedVanillaCopied = buildCopiedBrandMasterDeck(
      codegen(importedFixture.document, { target: "vanilla" }),
      {
        idleLayoutVar: "brandMasterBrandIdleContentLayout",
        masterRef: "brandMaster",
        themeVar: "officeTheme",
        titlePlaceholder: "title",
        bodyPlaceholder: "content",
      },
    );
    const importedReactCopied = buildCopiedBrandMasterDeck(
      codegen(importedFixture.document, { target: "react" }),
      {
        idleLayoutVar: "brandMasterBrandIdleContentLayout",
        masterRef: "<BrandMaster />",
        themeVar: "officeTheme",
        titlePlaceholder: "title",
        bodyPlaceholder: "content",
      },
    );

    const copiedCases = [
      {
        label: "copied-direct:vanilla",
        kind: "direct" as const,
        result: directVanillaCopied,
      },
      {
        label: "copied-direct:react",
        kind: "direct" as const,
        result: directReactCopied,
      },
      {
        label: "copied-imported:vanilla",
        kind: "imported" as const,
        result: importedVanillaCopied,
      },
      {
        label: "copied-imported:react",
        kind: "imported" as const,
        result: importedReactCopied,
      },
    ];

    await verifyGeneratedSourcesTypecheck(copiedCases);

    for (const item of copiedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(file.errors, [], `${item.label} validation errors`);
      assert.equal(file.masters?.length, 1);
      assert.equal(file.slides.length, 1);

      const newSlide = file.slides[0]!;
      assert.equal(
        newSlide.layoutPart,
        file.masters?.[0]?.layouts[1]?.part,
        `${item.label} new slide must bind to the previously idle second layout`,
      );
      assert.equal(newSlide.masterPart, file.masters?.[0]?.part);
      assert.equal(
        newSlide.objects[0]?.text?.text,
        "Continued From Copied Master",
      );
      assert.equal(newSlide.objects[0]?.placeholder?.type, "title");
      assert.equal(
        newSlide.objects[1]?.text?.text,
        "Created using previously idle layout",
      );
      assert.equal(newSlide.objects[1]?.placeholder?.idx, 1);
      assert.equal(newSlide.objects.length, 2);

      const zip = unzipSync(built.pptx);
      assert.match(
        strFromU8(zip["ppt/slideMasters/slideMaster1.xml"]!),
        /F8FAFC/u,
      );
      if (item.kind === "direct") {
        assert.deepEqual(
          (file.masters?.[0]?.objects ?? []).map((o) => o.name),
          ["Brand Logo", "Brand Wordmark"],
        );
      } else {
        assert.deepEqual(
          (file.masters?.[0]?.objects ?? []).map((o) => o.name),
          ["BrandBadge"],
        );
      }
    }
  }, 90_000);

  it("propagates layout title fontSize edit (32pt -> 22pt) only to the inheriting slide while preserving equal explicit 32pt override in both TS and TSX", async () => {
    const { vanillaDoc } = buildDirectBrandMasterFixtures();
    const fixtureRaw = readBrandFixtureBytes();
    const importedFixture = importPptx(fixtureRaw);
    fixtureRaw.fill(0);

    const editCases: { label: string; result: CodegenResult }[] = [];
    for (const [sourceLabel, doc] of [
      ["direct", vanillaDoc],
      ["imported", importedFixture.document],
    ] as const) {
      for (const target of ["vanilla", "react"] as const) {
        const generated = codegen(doc, { target });
        const originalSrc = readEntrySource(generated);
        // Edit only the first occurrence of fontSize: "32pt" (inside Brand Hero layout definition).
        const layoutTitleIndex = originalSrc.indexOf('fontSize: "32pt"');
        assert.ok(
          layoutTitleIndex >= 0,
          `${sourceLabel}:${target} missing layout 32pt`,
        );
        const editedSrc = `${originalSrc.slice(0, layoutTitleIndex)}fontSize: "22pt"${originalSrc.slice(layoutTitleIndex + 'fontSize: "32pt"'.length)}`;
        assert.notEqual(editedSrc, originalSrc);

        editCases.push({
          label: `${sourceLabel}:${target}`,
          result: {
            target,
            entry: generated.entry,
            files: {
              ...generated.files,
              [generated.entry]: editedSrc,
            },
          },
        });
      }
    }

    await verifyGeneratedSourcesTypecheck(editCases);

    for (const item of editCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(file.errors, [], `${item.label} validation errors`);
      assert.equal(file.slides.length, 2);

      const zip = unzipSync(built.pptx);
      const layout1Xml = strFromU8(zip["ppt/slideLayouts/slideLayout1.xml"]!);
      assert.match(
        layout1Xml,
        /sz="2200"/u,
        `${item.label} layout updated to 22pt`,
      );

      // Slide 1 omits local fontSize and inherits 22pt from slideLayout1.xml.
      assert.equal(
        file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        undefined,
        `${item.label} slide 1 must keep inheriting from layout`,
      );
      // Slide 2 retains explicit equal-value override at 32pt.
      assert.equal(
        file.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        32,
        `${item.label} slide 2 must remain locked at 32pt`,
      );
      assert.match(
        strFromU8(zip["ppt/slides/slide2.xml"]!),
        /sz="3200"/u,
        `${item.label} slide 2 XML retains sz="3200"`,
      );
    }
  }, 90_000);
});
