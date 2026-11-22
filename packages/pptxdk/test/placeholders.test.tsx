import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  BuildError,
  build,
  defineSlideLayout,
  emit,
  font,
  importPptx,
  resolve,
} from "pptxdk";
import { preview } from "pptxdk/preview";
import {
  Box,
  Deck,
  Image,
  P,
  Slide,
  SlideLayout,
  SlideMaster,
  Span,
  Text,
} from "pptxdk/react";
import {
  DeckContainer,
  SlideContainer,
  box,
  deck,
  image,
  paragraph,
  slide,
  slideLayout,
  slideMaster,
  span,
  text,
  type PlaceholderMap,
  type TextNode,
} from "pptxdk/vanilla";
import { type ReactNode } from "react";
import { describe, it } from "vitest";

const jostBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [
  { weight: 400, data: jostBytes },
  { weight: 700, data: jostBoldBytes },
]);
const tinyPng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

const sectionLayout = defineSlideLayout({
  id: "section",
  name: "Section",
  type: "title",
  placeholders: {
    title: {
      type: "title",
      x: 64,
      y: 48,
      width: 832,
      height: 72,
      style: { fontFamily: jostFont, fontSize: "32pt", color: "#0F172A" },
    },
    body: {
      type: "body",
      idx: 1,
      x: 64,
      y: 148,
      width: 832,
      height: 300,
      style: { fontFamily: jostFont, fontSize: "18pt", color: "#334155" },
    },
  },
});

const heroPictureLayout = defineSlideLayout({
  id: "hero-pic",
  name: "Hero Picture",
  type: "picTx",
  placeholders: {
    title: {
      type: "title",
      x: 40,
      y: 32,
      width: 400,
      height: 64,
      style: { fontFamily: jostFont, fontSize: "24pt" },
    },
    hero: {
      type: "pic",
      idx: 1,
      x: 480,
      y: 40,
      width: 400,
      height: 360,
    },
  },
});

function SectionSlide(props: {
  readonly id: string;
  readonly heading: string;
  readonly summary: string;
  readonly children?: ReactNode;
}) {
  return (
    <Slide
      id={props.id}
      layout={sectionLayout}
      placeholders={{
        title: props.heading,
        body: props.summary,
      }}
    >
      {props.children}
    </Slide>
  );
}

describe("typed text placeholders and style inheritance", () => {
  it("infers placeholder keys and content types without explicit generics in TSX and vanilla, supports free coexistence and wrapper components", async () => {
    const jsxTree = (
      <Deck width={960} height={540}>
        <SlideMaster id="m1" name="Brand">
          <SlideLayout layout={sectionLayout} />
        </SlideMaster>
        <SectionSlide id="s1" heading="Q3 Overview" summary="Revenue grew 22%">
          <Text
            id="free-note"
            x={64}
            y={470}
            width={400}
            height={32}
            style={{ fontFamily: jostFont, fontSize: "12pt", color: "#64748B" }}
          >
            Free canvas note
          </Text>
        </SectionSlide>
        <Slide
          id="s2"
          layout={sectionLayout}
          placeholders={{
            title: <Text style={{ color: "#1D4ED8" }}>Custom JSX Title</Text>,
            body: [
              <P key="p1">
                First <Span style={{ fontWeight: "bold" }}>bullet</Span>
              </P>,
            ],
          }}
        />
      </Deck>
    );

    const s1Vanilla = slide({
      id: "s1",
      layout: sectionLayout,
      placeholders: {
        title: "Q3 Overview",
        body: "Revenue grew 22%",
      },
    }).append(
      text("Free canvas note", {
        id: "free-note",
        x: 64,
        y: 470,
        width: 400,
        height: 32,
        style: { fontFamily: jostFont, fontSize: "12pt", color: "#64748B" },
      }),
    );

    const s2Vanilla = slide({
      id: "s2",
      layout: sectionLayout,
      placeholders: {
        title: text("Custom JSX Title", { style: { color: "#1D4ED8" } }),
        body: [
          paragraph([
            "First ",
            span("bullet", { style: { fontWeight: "bold" } }),
          ]),
        ],
      },
    });

    const vanillaDeck = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1", name: "Brand" }).append(
        slideLayout(sectionLayout),
      ),
      s1Vanilla,
      s2Vanilla,
    );

    const titleNode = s1Vanilla.placeholder("title") as TextNode | undefined;
    assert.ok(titleNode !== undefined);
    assert.equal(titleNode.props["placeholder"], "title");

    const jsxBuilt = await build(jsxTree);
    const vanillaBuilt = await build(vanillaDeck);
    assert.deepEqual(jsxBuilt.report, vanillaBuilt.report);

    const jsxPreview = await preview(jsxTree);
    assert.deepEqual(jsxPreview.pages[0]?.diagnostics, []);
    assert.deepEqual(jsxPreview.pages[1]?.diagnostics, []);

    const validated = await validatePptx(vanillaBuilt.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.slides[0]?.objects.length, 3);
    assert.equal(validated.slides[0]?.objects[0]?.text?.text, "Q3 Overview");
    assert.equal(
      validated.slides[0]?.objects[1]?.text?.text,
      "Revenue grew 22%",
    );
    assert.equal(
      validated.slides[0]?.objects[2]?.text?.text,
      "Free canvas note",
    );
  });

  it("enforces compile-time checks for misspelled keys, wrong content types, and missing layout", () => {
    const checkTypeErrors = () => {
      void (
        <Slide
          id="bad-key-jsx"
          layout={sectionLayout}
          // @ts-expect-error -- misspelled placeholder key must fail tsc
          placeholders={{ titel: "Typo" }}
        />
      );
      void (
        <Slide
          id="bad-type-jsx"
          layout={sectionLayout}
          // @ts-expect-error -- non-text ImageNode is invalid in title placeholder
          placeholders={{ title: image({ src: tinyPng }) }}
        />
      );
      void (
        <Slide
          id="bad-pic-jsx"
          layout={heroPictureLayout}
          // @ts-expect-error -- TextNode or string is invalid in pic placeholder
          placeholders={{ hero: "Not an image" }}
        />
      );
      void (
        <Slide
          id="bad-pic-textnode-jsx"
          layout={heroPictureLayout}
          // @ts-expect-error -- TextNode is invalid in pic placeholder
          placeholders={{ hero: text("Not an image") }}
        />
      );
      void (
        <Slide
          id="missing-layout-jsx"
          // @ts-expect-error -- placeholders cannot be supplied without layout
          placeholders={{ title: "No layout" }}
        />
      );

      slide({
        id: "bad-key-vanilla",
        layout: sectionLayout,
        // @ts-expect-error -- misspelled placeholder key must fail tsc
        placeholders: { titel: "Typo" },
      });
      slide({
        id: "bad-type-vanilla",
        layout: sectionLayout,
        // @ts-expect-error -- ImageNode is invalid in title placeholder
        placeholders: { title: image({ src: tinyPng }) },
      });
      slide({
        id: "bad-pic-vanilla",
        layout: heroPictureLayout,
        // @ts-expect-error -- string is invalid in pic placeholder
        placeholders: { hero: "Not an image" },
      });
      slide({
        id: "missing-layout-vanilla",
        // @ts-expect-error -- placeholders cannot be supplied without layout
        placeholders: { title: "No layout" },
      });
    };
    void checkTypeErrors;
    assert.ok(true);
  });

  it("enforces 3-way 32pt differential OOXML emission, layout 32pt->36pt cascade, fontSize undefined clearing, and single-slide box widening", async () => {
    const headingLayout = defineSlideLayout({
      id: "heading-32",
      name: "Heading 32pt",
      type: "title",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          x: 64,
          y: 48,
          width: 480,
          height: 64,
          style: { fontFamily: jostFont, fontSize: "32pt", color: "#1E293B" },
        },
      },
    });

    const jsxTree = (
      <Deck width={960} height={540}>
        <SlideMaster id="m-32" name="Brand 32" preserve>
          <SlideLayout layout={headingLayout} />
        </SlideMaster>
        <Slide
          id="s-inherited"
          layout={headingLayout}
          placeholders={{ title: "Inherited 32pt" }}
        >
          <Text
            id="free-32"
            x={64}
            y={140}
            width={400}
            height={56}
            style={{ fontFamily: jostFont, fontSize: "32pt", color: "#475569" }}
          >
            Free 32pt
          </Text>
        </Slide>
        <Slide
          id="s-override"
          layout={headingLayout}
          placeholders={{
            title: (
              <Text style={{ fontSize: "32pt", color: "#CC0000" }}>
                Explicit 32pt
              </Text>
            ),
          }}
        />
      </Deck>
    );

    const vanillaTree = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m-32", name: "Brand 32", preserve: true }).append(
        slideLayout(headingLayout),
      ),
      slide({
        id: "s-inherited",
        layout: headingLayout,
        placeholders: { title: "Inherited 32pt" },
      }).append(
        text("Free 32pt", {
          id: "free-32",
          x: 64,
          y: 140,
          width: 400,
          height: 56,
          style: { fontFamily: jostFont, fontSize: "32pt", color: "#475569" },
        }),
      ),
      slide({
        id: "s-override",
        layout: headingLayout,
        placeholders: {
          title: text("Explicit 32pt", {
            style: { fontSize: "32pt", color: "#CC0000" },
          }),
        },
      }),
    );

    const built = await build(jsxTree, { embedFonts: "subset" });
    const builtVanilla = await build(vanillaTree, { embedFonts: "subset" });
    assert.deepEqual(built.report, builtVanilla.report);

    const zipEntries = unzipSync(built.pptx);
    const slide1Xml = strFromU8(zipEntries["ppt/slides/slide1.xml"]!);
    const slide2Xml = strFromU8(zipEntries["ppt/slides/slide2.xml"]!);
    const slide1TitleSp = slide1Xml.split("<p:sp>")[1] ?? "";
    const slide1FreeSp = slide1Xml.split("<p:sp>")[2] ?? "";
    assert.match(slide1TitleSp, /<p:spPr\/>/u);
    assert.doesNotMatch(slide1TitleSp, /sz="3200"/u);
    assert.doesNotMatch(slide1TitleSp, /val="1E293B"/u);
    assert.match(slide1FreeSp, /sz="3200"/u);

    // Slide 2 declares explicit 32pt + crimson #CC0000, so BOTH must be emitted even though 32pt equals layout 32pt.
    assert.match(slide2Xml, /<p:spPr\/>/u);
    assert.match(slide2Xml, /sz="3200"/u);
    assert.match(slide2Xml, /val="CC0000"/u);

    // Import and mutate layout placeholder 32pt -> 36pt without touching slides.
    const imported = importPptx(built.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    assert.ok(imported.document instanceof DeckContainer);
    const importedMaster = imported.document.masters[0]!;
    const importedLayout = importedMaster.layouts[0]!;
    const oldPhs = importedLayout.props["placeholders"] as PlaceholderMap;
    const oldTitle = oldPhs["title"]!;
    const oldListStyles = oldTitle.listStyles ?? {};
    importedLayout.set({
      placeholders: {
        ...oldPhs,
        title: {
          ...oldTitle,
          listStyles: {
            ...oldListStyles,
            1: {
              ...oldListStyles[1],
              fontSize: "36pt",
            },
          },
        },
      },
    });

    const mutatedEmit = emit(resolve(imported.document));
    const mutatedZip = unzipSync(mutatedEmit.pptx);
    const mutatedSlide1 = strFromU8(mutatedZip["ppt/slides/slide1.xml"]!);
    const mutatedSlide2 = strFromU8(mutatedZip["ppt/slides/slide2.xml"]!);
    assert.doesNotMatch(mutatedSlide1.split("<p:sp>")[1] ?? "", /sz="3600"/u);
    assert.match(mutatedSlide1.split("<p:sp>")[2] ?? "", /sz="3200"/u);
    assert.match(mutatedSlide2, /sz="3200"/u);
    assert.match(mutatedSlide2, /val="CC0000"/u);

    // Verify Slide 1 title renders with 36pt in layout XML, Slide 2 retains 32pt, Slide 1 free stays 32pt.
    const mutatedLayoutXml = strFromU8(
      mutatedZip["ppt/slideLayouts/slideLayout1.xml"]!,
    );
    assert.match(mutatedLayoutXml, /sz="3600"/u);
    const mutatedValidated = await validatePptx(mutatedEmit.pptx);
    assert.deepEqual(mutatedValidated.errors, []);
    assert.equal(
      mutatedValidated.slides[0]?.objects[1]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      32,
    );
    assert.equal(
      mutatedValidated.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      32,
    );

    // Clear ONLY fontSize on Slide 2 title while retaining #CC0000, and widen Slide 1 title only.
    const importedSlides = imported.document.children.filter(
      (child): child is SlideContainer =>
        child instanceof SlideContainer && child.kind === "slide",
    );
    const s1Title = importedSlides[0]!.placeholder("title") as TextNode;
    const s2Title = importedSlides[1]!.placeholder("title") as TextNode;
    s2Title.style({ fontSize: undefined });
    s1Title.set({ width: 560 });

    const clearedEmit = emit(resolve(imported.document));
    const clearedZip = unzipSync(clearedEmit.pptx);
    const clearedSlide1 = strFromU8(clearedZip["ppt/slides/slide1.xml"]!);
    const clearedSlide2 = strFromU8(clearedZip["ppt/slides/slide2.xml"]!);
    assert.match(clearedSlide1.split("<p:sp>")[1] ?? "", /<a:xfrm>/u);
    assert.match(
      clearedSlide1.split("<p:sp>")[1] ?? "",
      new RegExp(`cx="${560 * 9525}"`, "u"),
    );
    assert.match(clearedSlide2, /<p:spPr\/>/u);
    assert.doesNotMatch(clearedSlide2, /sz="/u);
    assert.match(clearedSlide2, /val="CC0000"/u);
  });

  it("isolates Slide.style and Deck.style from bound placeholders, cascades master txStyles, and emits differential Span overrides", async () => {
    const cascadeLayout = defineSlideLayout({
      id: "cascade",
      name: "Cascade",
      type: "tx",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          style: { color: "#0F172A" },
        },
        body: {
          type: "body",
          idx: 1,
          x: 64,
          y: 140,
          width: 832,
          height: 300,
        },
      },
    });

    const built = await build(
      <Deck
        width={960}
        height={540}
        style={{ fontFamily: jostFont, fontSize: "14pt", color: "#22C55E" }}
      >
        <SlideMaster
          id="m-cascade"
          name="Cascade Master"
          placeholders={{
            title: {
              type: "title",
              x: 64,
              y: 40,
              width: 832,
              height: 64,
            },
          }}
          textStyles={{
            titleStyle: {
              1: { fontFamily: jostFont, fontSize: "30pt", color: "#1D4ED8" },
            },
            bodyStyle: {
              1: { fontFamily: jostFont, fontSize: "20pt", color: "#334155" },
              2: {
                fontFamily: jostFont,
                fontSize: "16pt",
                color: "#64748B",
                list: { type: "bullet", character: "–" },
              },
            },
          }}
        >
          <SlideLayout layout={cascadeLayout} />
        </SlideMaster>
        <Slide
          id="s-cascade"
          layout={cascadeLayout}
          style={{ fontSize: "15pt", color: "#EF4444" }}
          placeholders={{
            title: "Master Geometry & TitleStyle",
            body: [
              <P key="l1">
                Level 1{" "}
                <Span style={{ fontWeight: "bold", color: "#0284C7" }}>
                  accent
                </Span>
              </P>,
              <P key="l2" style={{ list: { type: "bullet", level: 1 } }}>
                Level 2 inherited bullet
              </P>,
            ],
          }}
        >
          <Text id="free-child" x={64} y={460} width={400} height={36}>
            Free inherits Slide.style
          </Text>
        </Slide>
      </Deck>,
    );

    const validated = await validatePptx(built.pptx);
    assert.deepEqual(validated.errors, []);
    const s1 = validated.slides[0]!;
    assert.equal(s1.objects.length, 3);
    // Title inherited geometry from masterPlaceholder because layout omitted x/y/width/height.
    assert.deepEqual(s1.objects[0]?.boxEmu, {
      x: 64 * 9525,
      y: 40 * 9525,
      width: 832 * 9525,
      height: 64 * 9525,
    });
    // Free text absorbed Slide.style (15pt, #EF4444).
    assert.equal(s1.objects[2]?.text?.paragraphs[0]?.runs[0]?.fontSizePt, 15);
    assert.equal(s1.objects[2]?.text?.paragraphs[0]?.runs[0]?.color, "#EF4444");

    const zipEntries = unzipSync(built.pptx);
    const slideXml = strFromU8(zipEntries["ppt/slides/slide1.xml"]!);
    const bodyXml = slideXml.split("<p:sp>")[2] ?? "";
    // Body placeholder does NOT absorb Deck.style (14pt #22C55E) or Slide.style (15pt #EF4444).
    assert.doesNotMatch(bodyXml, /EF4444|22C55E/u);
    assert.doesNotMatch(bodyXml, /sz="1400"|sz="1500"/u);
    // Inline Span only emits its local bold + color overrides.
    assert.match(bodyXml, /b="1"/u);
    assert.match(bodyXml, /val="0284C7"/u);
    assert.doesNotMatch(bodyXml, /sz="2000"/u);
  });

  it("imports, resolves, paints, and re-emits layouts.pptx (Comparison + idx=4294967295 fallback) and placeholder-layout-color.pptx", async () => {
    const layoutsRaw = new Uint8Array(
      readFileSync(new URL("./fixtures/layouts.pptx", import.meta.url)),
    );
    const layoutsImported = importPptx(layoutsRaw);
    layoutsRaw.fill(0);
    const slideSpDiagnostics = layoutsImported.report.diagnostics.filter(
      (d) =>
        d.location.part.startsWith("ppt/slides/") &&
        d.location.xpath?.includes("/p:sp["),
    );
    assert.deepEqual(slideSpDiagnostics, []);

    assert.ok(layoutsImported.document instanceof DeckContainer);
    const slides = layoutsImported.document.children.filter(
      (child): child is SlideContainer =>
        child instanceof SlideContainer && child.kind === "slide",
    );
    // Slide 5 (Comparison layout) preserves deterministic body-1/body-3 keys and <p:spPr/> (no local box).
    const slide5 = slides[4]!;
    assert.ok(slide5.placeholder("body-1") !== undefined);
    assert.ok(slide5.placeholder("content-2") !== undefined);
    assert.ok(slide5.placeholder("body-3") !== undefined);
    assert.ok(slide5.placeholder("content-4") !== undefined);
    // Slide 7 (Blank with <p:ph type="title" idx="4294967295"/>) falls back to master title placeholder.
    const layoutsResolved = resolve(layoutsImported.document);
    assert.equal(
      layoutsResolved.report.slides[6]?.objects[0]?.text,
      "Blank with Default Title",
    );
    const layoutsReemitted = emit(layoutsResolved);
    const layoutsValidated = await validatePptx(layoutsReemitted.pptx);
    assert.deepEqual(layoutsValidated.errors, []);
    assert.equal(
      layoutsValidated.slides[4]?.objects[1]?.hasLocalTransform,
      false,
    );
    assert.equal(
      layoutsValidated.slides[6]?.objects[0]?.hasLocalTransform,
      false,
    );
    assert.equal(
      layoutsValidated.slides[9]?.objects[4]?.text?.paragraphs[0]?.runs[0]
        ?.field,
      "slidenum",
    );
    const layoutsReimported = importPptx(layoutsReemitted.pptx);
    assert.deepEqual(layoutsReimported.report.diagnostics, []);

    const plcRaw = new Uint8Array(
      readFileSync(
        new URL("./fixtures/placeholder-layout-color.pptx", import.meta.url),
      ),
    );
    const plcImported = importPptx(plcRaw);
    plcRaw.fill(0);
    const plcSlideSpDiagnostics = plcImported.report.diagnostics.filter(
      (d) =>
        d.location.part.startsWith("ppt/slides/") &&
        d.location.xpath?.includes("/p:sp["),
    );
    assert.deepEqual(plcSlideSpDiagnostics, []);
    const plcResolved = resolve(plcImported.document);
    const plcReemitted = emit(plcResolved);
    const plcValidated = await validatePptx(plcReemitted.pptx);
    assert.deepEqual(plcValidated.errors, []);
    const plcTitle = plcValidated.slides[0]?.objects[0];
    assert.equal(plcTitle?.hasLocalTransform, true);
    assert.equal(plcTitle?.text?.text, "CYAN TEXT");
    const plcZip = unzipSync(plcReemitted.pptx);
    assert.match(
      strFromU8(plcZip["ppt/slideLayouts/slideLayout1.xml"]!),
      /sz="3500".*00FFFF/u,
    );
    assert.doesNotMatch(
      strFromU8(plcZip["ppt/slides/slide1.xml"]!),
      /00FFFF|sz="3500"/u,
    );
    const plcReimported = importPptx(plcReemitted.pptx);
    assert.deepEqual(plcReimported.report.diagnostics, []);

    const plcPreview = await preview(plcImported.document);
    assert.equal(plcPreview.pages.length, 1);
  });

  it("reports runtime diagnostics for nested placeholders, duplicates, unknown keys, missing layout, and incompatible role", async () => {
    const hasDiagnosticCode = (error: unknown, code: string): boolean =>
      error instanceof BuildError &&
      error.diagnostics.some((entry) => entry.code === code);

    await assert.rejects(
      () =>
        build(
          <Deck width={960} height={540}>
            <SlideMaster id="m-err">
              <SlideLayout layout={sectionLayout} />
            </SlideMaster>
            <Slide id="s-nested" layout={sectionLayout}>
              <Box style={{ width: 300, height: 100 }}>
                <Text placeholder="title">Nested</Text>
              </Box>
            </Slide>
          </Deck>,
        ),
      (error: unknown) =>
        hasDiagnosticCode(error, "placeholder-inside-container"),
    );

    await assert.rejects(
      () =>
        build(
          deck({ width: 960, height: 540 }).append(
            slideMaster({ id: "m-dup" }).append(slideLayout(sectionLayout)),
            slide({ id: "s-dup", layout: sectionLayout }).append(
              text("First", { placeholder: "title" }),
              text("Second", { placeholder: "title" }),
            ),
          ),
        ),
      (error: unknown) =>
        hasDiagnosticCode(error, "duplicate-slide-placeholder"),
    );

    await assert.rejects(
      () =>
        build(
          deck({ width: 960, height: 540 }).append(
            slideMaster({ id: "m-unk" }).append(slideLayout(sectionLayout)),
            slide({ id: "s-unk", layout: sectionLayout }).append(
              text("Unknown", { placeholder: "missing" }),
            ),
          ),
        ),
      (error: unknown) => hasDiagnosticCode(error, "unknown-slide-placeholder"),
    );

    await assert.rejects(
      () =>
        build(
          deck({ width: 960, height: 540 }).append(
            slide({ id: "s-nolayout" }).append(
              text("Orphan", { placeholder: "title" }),
            ),
          ),
        ),
      (error: unknown) =>
        hasDiagnosticCode(error, "placeholder-requires-layout"),
    );

    await assert.rejects(
      () =>
        build(
          deck({ width: 960, height: 540 }).append(
            slideMaster({ id: "m-incompat" }).append(
              slideLayout(heroPictureLayout),
            ),
            slide({ id: "s-incompat", layout: heroPictureLayout }).append(
              text("Wrong role", { placeholder: "hero" }),
            ),
          ),
        ),
      (error: unknown) =>
        hasDiagnosticCode(error, "incompatible-placeholder-content"),
    );

    await assert.rejects(
      () =>
        build(
          <Deck width={960} height={540}>
            <SlideMaster id="m-incompat-img">
              <SlideLayout layout={sectionLayout} />
            </SlideMaster>
            <Slide id="s-incompat-img" layout={sectionLayout}>
              <Image src={tinyPng} placeholder="title" />
            </Slide>
          </Deck>,
        ),
      (error: unknown) =>
        hasDiagnosticCode(error, "incompatible-placeholder-content"),
    );

    // Unused helper check in case box isn't referenced directly.
    void box;
  });
});
