import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import {
  build,
  DeckContainer,
  defaultColorMap,
  defineSlideLayout,
  defineTheme,
  font,
  importPptx,
  isTheme,
  type Theme,
} from "pptxdk";
import { preview } from "pptxdk/preview";
import {
  Deck,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Text,
} from "pptxdk/react";
import { deck, slide, slideLayout, slideMaster, text } from "pptxdk/vanilla";
import { describe, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
]);

describe("Slice 2: multi-master theme identity, clrMap, clrMapOvr, and themeOverride emission + validation", () => {
  it("shares one OPC theme part when two SlideMasters reference the same defineTheme() instance", async () => {
    const sharedTheme = defineTheme({
      name: "Shared Brand",
      colors: { accent1: "#123456", dk1: "#111827" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const layoutA = defineSlideLayout({
      id: "la",
      name: "Layout A",
      type: "blank",
    });
    const layoutB = defineSlideLayout({
      id: "lb",
      name: "Layout B",
      type: "blank",
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1", name: "Master 1", theme: sharedTheme }).append(
        slideLayout(layoutA),
      ),
      slideMaster({ id: "m2", name: "Master 2", theme: sharedTheme }).append(
        slideLayout(layoutB),
      ),
      slide({ id: "s1", layout: layoutA }).append(
        text("Master 1 text", {
          id: "t1",
          x: 40,
          y: 40,
          width: 300,
          height: 40,
          style: { color: "accent1", fontFamily: "minor" },
        }),
      ),
      slide({ id: "s2", layout: layoutB }).append(
        text("Master 2 text", {
          id: "t2",
          x: 40,
          y: 40,
          width: 300,
          height: 40,
          style: { color: "accent1", fontFamily: "minor" },
        }),
      ),
    );

    const built = await build(doc);
    const files = unzipSync(built.pptx);
    assert.ok(files["ppt/theme/theme1.xml"]);
    assert.equal(files["ppt/theme/theme2.xml"], undefined);

    const validated = await validatePptx(built.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.masters?.length, 2);
    assert.equal(validated.masters?.[0]?.themePart, "ppt/theme/theme1.xml");
    assert.equal(validated.masters?.[1]?.themePart, "ppt/theme/theme1.xml");
    assert.equal(
      validated.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#123456",
    );
    assert.equal(
      validated.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#123456",
    );
  });

  it("emits separate theme parts for two equal-value independent defineTheme() instances", async () => {
    const themeA = defineTheme({
      name: "Equal Office",
      colors: { accent1: "#1D4ED8" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const themeB = defineTheme({
      name: "Equal Office",
      colors: { accent1: "#1D4ED8" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const layoutA = defineSlideLayout({
      id: "la",
      name: "Layout A",
      type: "blank",
    });
    const layoutB = defineSlideLayout({
      id: "lb",
      name: "Layout B",
      type: "blank",
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1", name: "Master A", theme: themeA }).append(
        slideLayout(layoutA),
      ),
      slideMaster({ id: "m2", name: "Master B", theme: themeB }).append(
        slideLayout(layoutB),
      ),
      slide({ id: "s1", layout: layoutA }).append(
        text("A", {
          id: "t1",
          x: 40,
          y: 40,
          width: 200,
          height: 40,
          style: { color: "accent1" },
        }),
      ),
      slide({ id: "s2", layout: layoutB }).append(
        text("B", {
          id: "t2",
          x: 40,
          y: 40,
          width: 200,
          height: 40,
          style: { color: "accent1" },
        }),
      ),
    );

    const firstBuilt = await build(doc);
    const firstFiles = unzipSync(firstBuilt.pptx);
    assert.ok(firstFiles["ppt/theme/theme1.xml"]);
    assert.ok(firstFiles["ppt/theme/theme2.xml"]);

    // Mutating themeA affects ONLY m1 / s1, leaving m2 / s2 untouched!
    themeA.set({ colors: { accent1: "#DC2626" } });
    const secondBuilt = await build(doc);
    const validated = await validatePptx(secondBuilt.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.masters?.[0]?.themePart, "ppt/theme/theme1.xml");
    assert.equal(validated.masters?.[1]?.themePart, "ppt/theme/theme2.xml");
    assert.equal(
      validated.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#DC2626",
    );
    assert.equal(
      validated.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#1D4ED8",
    );
  });

  it("emits and validates master clrMap, layout/slide clrMapOvr, themeOverride, and explicit schemeClr slots per slide", async () => {
    const invertedMap = {
      ...defaultColorMap,
      bg1: "dk1",
      tx1: "lt1",
      bg2: "dk2",
      tx2: "lt2",
    } as const;
    const layoutOverridden = defineSlideLayout({
      id: "layout-ovr",
      name: "Overridden Layout",
      type: "blank",
      colorMapOverride: { ...defaultColorMap, tx1: "accent2" },
      themeOverride: {
        colors: { accent2: "#059669" },
      },
    });
    const layoutDefault = defineSlideLayout({
      id: "layout-default",
      name: "Default Layout",
      type: "blank",
    });

    const masterTheme = defineTheme({
      name: "Multi Palette",
      colors: {
        dk1: "#111111",
        lt1: "#FAFAFA",
        dk2: "#1E293B",
        lt2: "#E2E8F0",
        accent1: "#2563EB",
        accent2: "#EA580C",
      },
      fonts: { major: jostFont, minor: jostFont },
    });

    const jsxBuilt = await build(
      <Deck width={960} height={540}>
        <SlideMaster
          id="m1"
          name="Inverted Master"
          theme={masterTheme}
          colorMap={invertedMap}
        >
          <SlideLayout layout={layoutDefault} />
          <SlideLayout layout={layoutOverridden} />
        </SlideMaster>
        <Slide id="s-master-map" layout={layoutDefault}>
          <Text
            id="t-tx1"
            x={40}
            y={40}
            width={300}
            height={40}
            style={{ color: "tx1" }}
          >
            Uses master tx1 (= lt1)
          </Text>
          <Text
            id="t-lt2"
            x={40}
            y={90}
            width={300}
            height={40}
            style={{ color: "lt2" }}
          >
            Direct lt2
          </Text>
        </Slide>
        <Slide
          id="s-layout-ovr"
          layout={layoutOverridden}
          themeOverride={{ colors: { accent1: "#9333EA" } }}
        >
          <Text
            id="t-ovr-tx1"
            x={40}
            y={40}
            width={300}
            height={40}
            style={{ color: "tx1" }}
          >
            Layout overridden tx1 (= accent2 from layout themeOverride)
          </Text>
          <Text
            id="t-ovr-acc1"
            x={40}
            y={90}
            width={300}
            height={40}
            style={{ color: "accent1" }}
          >
            Slide themeOverride accent1
          </Text>
        </Slide>
      </Deck>,
    );

    const validated = await validatePptx(jsxBuilt.pptx);
    assert.deepEqual(validated.errors, []);

    const files = unzipSync(jsxBuilt.pptx);
    const masterXml = strFromU8(files["ppt/slideMasters/slideMaster1.xml"]!);
    assert.match(masterXml, /<p:clrMap\b[^>]*bg1="dk1"[^>]*tx1="lt1"/u);

    const layout2Xml = strFromU8(files["ppt/slideLayouts/slideLayout2.xml"]!);
    assert.match(layout2Xml, /<a:overrideClrMapping\b[^>]*tx1="accent2"/u);

    const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);
    assert.match(slide1Xml, /<a:schemeClr val="tx1"\/>/u);
    assert.match(slide1Xml, /<a:schemeClr val="lt2"\/>/u);

    // Independent validator evaluates each slide's active theme + themeOverride + clrMap/clrMapOvr chain!
    assert.equal(
      validated.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#FAFAFA",
    );
    assert.equal(
      validated.slides[0]?.objects[1]?.text?.paragraphs[0]?.runs[0]?.color,
      "#E2E8F0",
    );
    assert.equal(
      validated.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.color,
      "#059669",
    );
    assert.equal(
      validated.slides[1]?.objects[1]?.text?.paragraphs[0]?.runs[0]?.color,
      "#9333EA",
    );
  });

  it("roundtrips shared Theme identity vs equal-value independent Themes through importPptx -> .set() -> preview -> build -> validatePptx", async () => {
    const sharedTheme = defineTheme({
      name: "Shared Identity",
      colors: { accent1: "#2563EB" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const layoutA = defineSlideLayout({ id: "la", name: "LA", type: "blank" });
    const layoutB = defineSlideLayout({ id: "lb", name: "LB", type: "blank" });

    const sharedBuilt = await build(
      <Deck width={640} height={360}>
        <SlideMaster id="m1" name="M1" theme={sharedTheme}>
          <SlideLayout layout={layoutA} />
        </SlideMaster>
        <SlideMaster id="m2" name="M2" theme={sharedTheme}>
          <SlideLayout layout={layoutB} />
        </SlideMaster>
        <Slide id="s1" layout={layoutA}>
          <Shape
            id="bg-s1"
            x={0}
            y={0}
            width={640}
            height={360}
            fill="accent1"
          />
          <Shape id="fg-s1" x={40} y={40} width={120} height={40} fill="tx1" />
        </Slide>
        <Slide id="s2" layout={layoutB}>
          <Shape
            id="bg-s2"
            x={0}
            y={0}
            width={640}
            height={360}
            fill="accent1"
          />
          <Shape id="fg-s2" x={40} y={40} width={120} height={40} fill="tx1" />
        </Slide>
      </Deck>,
    );
    const sharedBytes = new Uint8Array(sharedBuilt.pptx);
    const importedShared = importPptx(sharedBytes);
    sharedBytes.fill(0);
    assert.ok(importedShared.document instanceof DeckContainer);
    const [sm1, sm2] = importedShared.document.masters;
    assert.ok(sm1?.theme && isTheme(sm1.theme));
    assert.equal(sm1.theme, sm2?.theme);

    sm1.theme.set({ colors: { accent1: "#BE123C" } });
    const sharedPreview = await preview(importedShared.document);
    assert.equal(sharedPreview.pages.length, 2);
    for (const page of sharedPreview.pages) {
      assert.ok(page.png instanceof Uint8Array);
      const png = PNG.sync.read(Buffer.from(page.png));
      const idx = (10 * png.width + 10) * 4;
      assert.deepEqual(
        [png.data[idx], png.data[idx + 1], png.data[idx + 2]],
        [0xbe, 0x12, 0x3c],
      );
    }
    const sharedRebuilt = await build(importedShared.document);
    const sharedParts = unzipSync(sharedRebuilt.pptx);
    assert.ok(sharedParts["ppt/theme/theme1.xml"]);
    assert.equal(sharedParts["ppt/theme/theme2.xml"], undefined);
    const sharedValidated = await validatePptx(sharedRebuilt.pptx);
    assert.deepEqual(sharedValidated.errors, []);

    // Equal-value independent Themes remain distinct across importPptx and edit independently.
    const indA = defineTheme({
      name: "Equal",
      colors: { accent1: "#1D4ED8" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const indB = defineTheme({
      name: "Equal",
      colors: { accent1: "#1D4ED8" },
      fonts: { major: jostFont, minor: jostFont },
    });
    const indBuilt = await build(
      <Deck width={640} height={360}>
        <SlideMaster id="m1" name="M1" theme={indA}>
          <SlideLayout layout={layoutA} />
        </SlideMaster>
        <SlideMaster id="m2" name="M2" theme={indB}>
          <SlideLayout layout={layoutB} />
        </SlideMaster>
        <Slide id="s1" layout={layoutA}>
          <Shape
            id="bg-i1"
            x={0}
            y={0}
            width={640}
            height={360}
            fill="accent1"
          />
        </Slide>
        <Slide id="s2" layout={layoutB}>
          <Shape
            id="bg-i2"
            x={0}
            y={0}
            width={640}
            height={360}
            fill="accent1"
          />
        </Slide>
      </Deck>,
    );
    const indBytes = new Uint8Array(indBuilt.pptx);
    const importedInd = importPptx(indBytes);
    indBytes.fill(0);
    assert.ok(importedInd.document instanceof DeckContainer);
    const [im1, im2] = importedInd.document.masters;
    assert.ok(im1?.theme && im2?.theme);
    assert.notEqual(im1.theme, im2.theme);

    (im2.theme as Theme).set({ colors: { accent1: "#0D9488" } });
    const indPreview = await preview(importedInd.document);
    assert.ok(indPreview.pages[0]?.png instanceof Uint8Array);
    assert.ok(indPreview.pages[1]?.png instanceof Uint8Array);
    const p1 = PNG.sync.read(Buffer.from(indPreview.pages[0].png));
    const p2 = PNG.sync.read(Buffer.from(indPreview.pages[1].png));
    const o1 = (10 * p1.width + 10) * 4;
    const o2 = (10 * p2.width + 10) * 4;
    assert.deepEqual(
      [p1.data[o1], p1.data[o1 + 1], p1.data[o1 + 2]],
      [0x1d, 0x4e, 0xd8],
    );
    assert.deepEqual(
      [p2.data[o2], p2.data[o2 + 1], p2.data[o2 + 2]],
      [0x0d, 0x94, 0x88],
    );

    const indRebuilt = await build(importedInd.document);
    const indValidated = await validatePptx(indRebuilt.pptx);
    assert.deepEqual(indValidated.errors, []);
    assert.equal(indValidated.masters?.[0]?.theme?.colors.accent1, "#1D4ED8");
    assert.equal(indValidated.masters?.[1]?.theme?.colors.accent1, "#0D9488");
  });

  it("imports external themes.pptx and placeholder-layout-color.pptx via SDK importPptx", () => {
    const themesBytes = new Uint8Array(
      readFileSync(new URL("./fixtures/themes.pptx", import.meta.url)),
    );
    const themesResult = importPptx(themesBytes);
    themesBytes.fill(0);
    assert.ok(themesResult.document instanceof DeckContainer);
    assert.equal(themesResult.document.masters.length, 7);
    assert.equal(
      (themesResult.document.masters[6]!.theme as Theme).name,
      "Tradeshow",
    );
    assert.deepEqual(
      {
        ...(themesResult.document.masters[6]!.props["colorMap"] as Record<
          string,
          string
        >),
      },
      {
        ...defaultColorMap,
        bg1: "dk1",
        tx1: "lt1",
        bg2: "dk2",
        tx2: "lt2",
      },
    );

    const plcBytes = new Uint8Array(
      readFileSync(
        new URL("./fixtures/placeholder-layout-color.pptx", import.meta.url),
      ),
    );
    const plcResult = importPptx(plcBytes);
    plcBytes.fill(0);
    assert.ok(plcResult.document instanceof DeckContainer);
    assert.equal(plcResult.document.masters.length, 1);
    assert.equal(
      (plcResult.document.masters[0]!.theme as Theme).colors.dk1,
      "#2D2842",
    );
    assert.deepEqual(
      {
        ...(plcResult.document.masters[0]!.props["colorMap"] as Record<
          string,
          string
        >),
      },
      {
        ...defaultColorMap,
        bg2: "dk2",
        tx2: "lt2",
      },
    );
  });

  it("preserves schemeClr tx1/bg1/tx2/bg2 in master txStyles across emit -> import -> resolve -> paint so dark layout colorMapOverride resolves inherited title color to lt1", async () => {
    const darkLayout = defineSlideLayout({
      id: "dark-title",
      name: "Dark Title",
      type: "title",
      colorMapOverride: {
        ...defaultColorMap,
        bg1: "dk1",
        tx1: "lt1",
        bg2: "dk2",
        tx2: "lt2",
      },
      background: "bg1",
      placeholders: {
        title: {
          type: "title",
          x: 40,
          y: 40,
          width: 480,
          height: 80,
        },
      },
    });
    const lightTheme = defineTheme({
      name: "LightBase",
      colors: {
        dk1: "#111827",
        lt1: "#F8FAFC",
      },
      fonts: { major: jostFont, minor: jostFont },
    });

    const built = await build(
      <Deck width={640} height={360}>
        <SlideMaster
          id="m-scheme"
          name="Scheme Master"
          theme={lightTheme}
          colorMap={defaultColorMap}
          textStyles={{
            titleStyle: {
              1: { fontFamily: jostFont, fontSize: 28, color: "tx1" },
            },
          }}
        >
          <SlideLayout layout={darkLayout} />
        </SlideMaster>
        <Slide id="dark-slide" layout={darkLayout}>
          <Text id="s-title" placeholder="title">
            Dark Layout Title
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );

    const initialParts = unzipSync(built.pptx);
    const initialMasterXml = strFromU8(
      initialParts["ppt/slideMasters/slideMaster1.xml"]!,
    );
    assert.match(initialMasterXml, /<a:schemeClr val="tx1"\/>/u);

    const initialValidated = await validatePptx(built.pptx);
    assert.deepEqual(initialValidated.errors, []);
    assert.equal(
      initialValidated.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.color,
      undefined,
    );
    const initialPainted = await preview(
      <Deck width={640} height={360}>
        <SlideMaster
          id="m-scheme"
          name="Scheme Master"
          theme={lightTheme}
          colorMap={defaultColorMap}
          textStyles={{
            titleStyle: {
              1: { fontFamily: jostFont, fontSize: 28, color: "tx1" },
            },
          }}
        >
          <SlideLayout layout={darkLayout} />
        </SlideMaster>
        <Slide id="dark-slide" layout={darkLayout}>
          <Text id="s-title" placeholder="title">
            Dark Layout Title
          </Text>
        </Slide>
      </Deck>,
    );
    assert.deepEqual(initialPainted.pages[0]?.diagnostics, []);
    assert.ok(
      initialPainted.pages[0]?.svg?.includes("#F8FAFC"),
      "Expected initial dark layout inherited title color to resolve through colorMapOverride to lt1 (#F8FAFC)",
    );

    const imported = importPptx(new Uint8Array(built.pptx));
    assert.deepEqual(imported.report.diagnostics, []);
    assert.ok(imported.document instanceof DeckContainer);
    const importedTextStyles = imported.document.masters[0]?.props[
      "textStyles"
    ] as
      | {
          readonly titleStyle?: {
            readonly 1?: { readonly color?: unknown };
          };
        }
      | undefined;
    assert.equal(importedTextStyles?.titleStyle?.[1]?.color, "tx1");

    const reemitted = await build(imported.document);
    const reemittedParts = unzipSync(reemitted.pptx);
    const reemittedMasterXml = strFromU8(
      reemittedParts["ppt/slideMasters/slideMaster1.xml"]!,
    );
    assert.match(reemittedMasterXml, /<a:schemeClr val="tx1"\/>/u);

    const reemittedValidated = await validatePptx(reemitted.pptx);
    assert.deepEqual(reemittedValidated.errors, []);

    const painted = await preview(imported.document);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(
      painted.pages[0]?.svg?.includes("#F8FAFC"),
      "Expected dark layout inherited title color to resolve through colorMapOverride to lt1 (#F8FAFC)",
    );
  });
});
