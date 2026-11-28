import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  defineSlideLayout,
  font,
  shape,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { Text } from "@pptxdk/capabilities/text-elements";
import {
  CommonNode,
  DeckContainer,
  SlideContainer,
  SlideLayoutContainer,
  SlideMasterContainer,
  defaultColorMap,
  defineTheme,
  type LayoutPlaceholderSpec,
  type Theme,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import {
  runCapabilityScenario,
  type CapabilityScenario,
} from "../../../test/acceptance-runner.ts";
import { footerPlaceholderScenario } from "./footer-mechanisms.ts";
import { masterLayoutOwnershipScenario } from "./master-layout-mechanisms.ts";
import { masterSharedCanvasScenario } from "./shared-canvas-mechanisms.ts";
import { nativeThemesRoundtripScenario } from "./theme-mechanisms.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const notoScFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: readFileSync(
      new URL(
        "../../../test/assets/fonts/NotoSansSC-Regular.woff2",
        import.meta.url,
      ),
    ),
  },
]);

const jostFullFamilyFont = font(
  "Jost*",
  [
    {
      weight: 400,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-400-Book.ttf",
          import.meta.url,
        ),
      ),
    },
    {
      weight: 400,
      style: "italic",
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-400-Italic.ttf",
          import.meta.url,
        ),
      ),
    },
    {
      weight: 700,
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-700-Bold.ttf",
          import.meta.url,
        ),
      ),
    },
    {
      weight: 700,
      style: "italic",
      data: readFileSync(
        new URL(
          "../../../test/assets/fonts/Jost-700-BoldItalic.ttf",
          import.meta.url,
        ),
      ),
    },
  ],
  { substitute: true },
);

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

const expectedMasterFixtureSha256: Readonly<Record<string, string>> = {
  "prs-slide-masters.pptx":
    "6f2a9083194f537179b582ba962c9fc324bba76fd18f50e8446e7760ffb5d7eb",
  "layouts.pptx":
    "9c3d53afa3115de2ba72ec77637a616aeed505a6ff3b30a46db871c4edd32be5",
  "WithMaster.pptx":
    "54ae7677fd2813a815160cc72bd8960a4e482690dfc369b392868521c9b61549",
  "placeholder-layout-color.pptx":
    "b683af99cf4e71db112fe87db15840846c53dbc7a6839cbf84804980d21af9ce",
};

function readMasterFixture(name: string): Uint8Array {
  const bytes = new Uint8Array(
    readFileSync(new URL(`./fixtures/${name}`, import.meta.url)),
  );
  const expected = expectedMasterFixtureSha256[name];
  assert.ok(expected !== undefined, `Missing expected SHA-256 for ${name}`);
  assert.equal(
    createHash("sha256").update(bytes).digest("hex"),
    expected,
    `SHA-256 mismatch for ${name}`,
  );
  return bytes;
}

function readPngPixelHex(
  pngBytes: Uint8Array | undefined,
  x: number,
  y: number,
): string {
  assert.ok(pngBytes instanceof Uint8Array);
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  const r = decoded.data[offset] ?? 0;
  const g = decoded.data[offset + 1] ?? 0;
  const b = decoded.data[offset + 2] ?? 0;
  return `#${[r, g, b]
    .map((c) => c.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

function collectSlides(root: CommonNode): readonly SlideContainer[] {
  return root.children.filter(
    (child): child is SlideContainer =>
      child instanceof SlideContainer && child.kind === "slide",
  );
}

const inheritanceTitleLayout = defineSlideLayout({
  id: "inheritance-title-layout",
  name: "Inheritance Title Layout",
  type: "title",
  preserve: true,
  placeholders: {
    title: {
      type: "title",
      style: {
        position: "absolute",
        left: 48,
        top: 36,
        width: 544,
        height: 64,
        fontFamily: jostFont,
        fontSize: "32pt",
        color: "#0F172A",
      },
    },
  },
});

const layoutInheritanceOverrideScenario: CapabilityScenario = {
  id: "masters-layout-inheritance-vs-equal-override",
  capability: "masters",
  description:
    "Dual-author scenario contrasting a slide title inheriting layout 32pt fontSize with a sibling slide explicitly declaring an equal 32pt override.",
  visualCheckpoints: [
    "Page 1 renders layout-inherited 32pt title; Page 2 renders locally overridden equal-value 32pt crimson title.",
  ],
  nativeCheckpoints: [
    "slide1 omits local sz while slide2 emits sz='3200'; independent layout edit (32pt -> 20pt) updates slide1 while preserving slide2 at 32pt.",
  ],
  emitOptions: { embedFonts: "subset" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "m-inherit", name: "Inheritance Master", preserve: true },
        createElement(SlideLayout, { layout: inheritanceTitleLayout }),
      ),
      createElement(Slide, {
        id: "ph-inherited",
        layout: inheritanceTitleLayout,
        placeholders: { title: "Inherited 32pt" },
      }),
      createElement(Slide, {
        id: "ph-explicit-equal",
        layout: inheritanceTitleLayout,
        placeholders: {
          title: createElement(
            Text,
            { style: { fontSize: "32pt", color: "#CC0000" } },
            "Explicit Equal 32pt",
          ),
        },
      }),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m-inherit",
        name: "Inheritance Master",
        preserve: true,
      }).append(slideLayout(inheritanceTitleLayout)),
      slide({
        id: "ph-inherited",
        layout: inheritanceTitleLayout,
        placeholders: { title: "Inherited 32pt" },
      }),
      slide({
        id: "ph-explicit-equal",
        layout: inheritanceTitleLayout,
        placeholders: {
          title: text("Explicit Equal 32pt", {
            style: { fontSize: "32pt", color: "#CC0000" },
          }),
        },
      }),
    );
  },
  assert({ validation, preview }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 2);
    // Slide 1 omits local sz and inherits 32pt from layout; Slide 2 explicitly declares equal-value 32pt.
    assert.equal(
      validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      undefined,
    );
    assert.equal(
      validation.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      32,
    );
    assert.deepEqual(preview.pages[0]?.errors, []);
    assert.deepEqual(preview.pages[1]?.errors, []);
  },
};

describe("Masters capability acceptance, external fixtures, inheritance & isolation", () => {
  it("runs masters dual-author mechanism scenarios through shared runner with single pure-file roundtrip", async () => {
    for (const scenario of [
      masterLayoutOwnershipScenario,
      masterSharedCanvasScenario,
      footerPlaceholderScenario,
      nativeThemesRoundtripScenario,
    ]) {
      const execution = await runCapabilityScenario(scenario);
      assert.equal(execution.runs.length, 2);
      assert.deepEqual(
        execution.runs.map((run) => run.author),
        ["jsx", "document"],
      );
      assert.deepEqual(execution.roundtrip.importReport.diagnostics, []);
    }
  });

  it("isolates pure-file recovery from extra-font-environment observation and verifies embedded-font recovery", async () => {
    // 1. Unembedded theme fonts: pure-file roundtrip preserves native theme/clrMap/themeOverride without font reattachment,
    //    while static paint requires an explicit extra font environment supplied on an independent editable instance.
    const themeExec = await runCapabilityScenario(
      nativeThemesRoundtripScenario,
    );
    for (const page of themeExec.roundtrip.preview.pages) {
      assert.equal(
        page.png,
        undefined,
        `Pure-file roundtrip page ${page.page} without embedded font must omit png`,
      );
      assert.ok(
        page.errors.some((err) => err.code === "preview-missing-font"),
        `Pure-file roundtrip page ${page.page} must report preview-missing-font`,
      );
    }

    // Extra font environment observation on an independent editable instance (never mutating themeExec.roundtrip).
    const editableTheme = themeExec.openEditableRoundtripDocument();
    assert.ok(editableTheme.document instanceof DeckContainer);
    for (const master of editableTheme.document.masters) {
      const existingTheme = master.theme as Theme | undefined;
      if (existingTheme !== undefined) {
        master.set({
          theme: {
            ...existingTheme,
            fonts: { major: jostFont, minor: jostFont },
          },
        });
      }
    }
    const extraFontResolution = resolveDocument(editableTheme.document);
    const extraFontPaint = paint(extraFontResolution);
    assert.equal(
      extraFontPaint.pages.length,
      3,
      "Extra font environment observation (explicit Jost* TTF attached on editable instance): expected 3 pages",
    );
    for (const page of extraFontPaint.pages) {
      assert.deepEqual(
        page.errors,
        [],
        `Extra font environment observation page ${page.page} expected zero errors`,
      );
      assert.ok(page.png instanceof Uint8Array);
    }
    assert.equal(
      readPngPixelHex(extraFontPaint.pages[0]!.png, 80, 80),
      "#F8FAFC",
      "Extra font environment observation page 1 clrMap tx1 -> lt1",
    );
    assert.equal(
      readPngPixelHex(extraFontPaint.pages[1]!.png, 80, 80),
      "#DC2626",
      "Extra font environment observation page 2 clrMapOvr + themeOverride",
    );
    assert.equal(
      readPngPixelHex(extraFontPaint.pages[2]!.png, 80, 80),
      "#2563EB",
      "Extra font environment observation page 3 independent theme2 accent1",
    );

    // 2. Actual font-carrying sample (embedFonts: "subset"): pure-file roundtrip recovers embedded TTF from ppt/fonts/*.fntdata.
    const embeddedExec = await runCapabilityScenario(footerPlaceholderScenario);
    const embeddedZip = unzipSync(embeddedExec.primary.pptx);
    assert.ok(
      Object.keys(embeddedZip).some((name) => name.startsWith("ppt/fonts/")),
      "Embedded-font sample must carry ppt/fonts/*.fntdata",
    );
    for (const page of embeddedExec.roundtrip.preview.pages) {
      assert.deepEqual(
        page.errors,
        [],
        `Embedded-font pure-file roundtrip page ${page.page} must paint without extra font environment`,
      );
      assert.ok(page.png instanceof Uint8Array);
    }
  });

  it("distinguishes layout fontSize inheritance from equal-value explicit override across shared read-only inspection and independent editable mutation", async () => {
    const execution = await runCapabilityScenario(
      layoutInheritanceOverrideScenario,
    );

    // 1. Inspect shared read-only roundtrip artifact: slide1 omits sz; slide2 preserves explicit sz="3200".
    const rtZip = unzipSync(execution.roundtrip.pptx);
    const rtLayoutXml = strFromU8(rtZip["ppt/slideLayouts/slideLayout1.xml"]!);
    const rtSlide1Xml = strFromU8(rtZip["ppt/slides/slide1.xml"]!);
    const rtSlide2Xml = strFromU8(rtZip["ppt/slides/slide2.xml"]!);
    assert.match(rtLayoutXml, /sz="3200"/u);
    assert.doesNotMatch(rtSlide1Xml, /sz="3200"/u);
    assert.match(rtSlide2Xml, /sz="3200"/u);

    const rtSlides = collectSlides(execution.roundtrip.importedDocument);
    const rtSlide1Title = rtSlides[0]?.placeholder("title");
    const rtSlide2Title = rtSlides[1]?.placeholder("title");
    assert.ok(rtSlide1Title !== undefined && rtSlide2Title !== undefined);
    assert.equal(
      (rtSlide1Title.props["style"] as Record<string, unknown> | undefined)?.[
        "fontSize"
      ],
      undefined,
      "Read-only imported Slide 1 title must omit local fontSize declaration",
    );
    assert.equal(
      (rtSlide2Title.props["style"] as Record<string, unknown> | undefined)?.[
        "fontSize"
      ],
      "32pt",
      "Read-only imported Slide 2 title must retain explicit equal-value fontSize ('32pt')",
    );

    // 2. Open an independent editable instance, mutate layout title fontSize 32pt -> 20pt, and re-emit.
    const editable = execution.openEditableRoundtripDocument();
    assert.ok(editable.document instanceof DeckContainer);
    const layoutNode = editable.document.masters[0]?.layouts[0];
    assert.ok(layoutNode instanceof SlideLayoutContainer);
    const currentPlaceholders = layoutNode.props["placeholders"] as Readonly<
      Record<string, LayoutPlaceholderSpec>
    >;
    const currentTitle = currentPlaceholders["title"]!;
    layoutNode.set({
      placeholders: {
        ...currentPlaceholders,
        title: {
          ...currentTitle,
          style: {
            ...currentTitle.style,
            fontSize: "20pt",
          },
          ...(currentTitle.listStyles === undefined
            ? {}
            : {
                listStyles: {
                  ...currentTitle.listStyles,
                  1: {
                    ...currentTitle.listStyles[1],
                    fontSize: "20pt",
                  },
                },
              }),
        },
      },
    });

    const editedResolution = resolveDocument(editable.document);
    const editedEmit = emit(editedResolution, { embedFonts: "subset" });
    const editedPaint = paint(editedResolution);
    const editedValidation = await validatePptx(editedEmit.pptx);
    assert.deepEqual(editedValidation.errors, []);
    assert.equal(
      editedValidation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      undefined,
      "Inherited Slide 1 title continues omitting local fontSize while inheriting 20pt from layout",
    );
    assert.equal(
      editedValidation.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]
        ?.fontSizePt,
      32,
      "Explicit equal-value override Slide 2 title must stay at 32pt after layout edit",
    );
    const editedZip = unzipSync(editedEmit.pptx);
    assert.match(
      strFromU8(editedZip["ppt/slideLayouts/slideLayout1.xml"]!),
      /sz="2000"/u,
    );
    assert.doesNotMatch(
      strFromU8(editedZip["ppt/slides/slide1.xml"]!),
      /sz="/u,
    );
    assert.match(strFromU8(editedZip["ppt/slides/slide2.xml"]!), /sz="3200"/u);

    // Visual confirmation: Slide 1 rendered glyphs change when layout changes 32pt -> 20pt,
    // whereas Slide 2 with equal-value explicit 32pt override renders identically.
    assert.deepEqual(editedPaint.pages[0]?.errors, []);
    assert.deepEqual(editedPaint.pages[1]?.errors, []);
    assert.notEqual(
      editedPaint.pages[0]?.svg,
      execution.roundtrip.preview.pages[0]?.svg,
      "Inherited Slide 1 SVG must update after layout fontSize edit 32pt -> 20pt",
    );
    assert.equal(
      editedPaint.pages[1]?.svg,
      execution.roundtrip.preview.pages[1]?.svg,
      "Explicit equal-value override Slide 2 SVG must stay identical at 32pt",
    );
    assert.equal(
      Buffer.compare(
        editedPaint.pages[1].png!,
        execution.roundtrip.preview.pages[1]!.png!,
      ),
      0,
      "Explicit equal-value override Slide 2 PNG must stay identical at 32pt",
    );

    // 3. Shared read-only execution.roundtrip must remain untouched at 32pt.
    assert.match(
      strFromU8(
        unzipSync(execution.roundtrip.pptx)[
          "ppt/slideLayouts/slideLayout1.xml"
        ]!,
      ),
      /sz="3200"/u,
    );
  });

  it("imports external prs-slide-masters.pptx (0 slides, 2 masters, idle layouts) and binds a new slide to Master 2 idle layout", async () => {
    const raw = readMasterFixture("prs-slide-masters.pptx");
    const imported = importDocument(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 0);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 2);

    // Pure-file 0-slide roundtrip retains both masters and all idle layouts.
    const zeroEmitted = emit(resolveDocument(imported.document));
    const zeroValidated = await validatePptx(zeroEmitted.pptx);
    assert.deepEqual(zeroValidated.errors, []);
    assert.equal(zeroValidated.presentation.slideCount, 0);
    assert.equal(zeroValidated.masters?.length, 2);
    assert.equal(zeroValidated.masters?.[0]?.preserve, undefined);
    assert.equal(zeroValidated.masters?.[1]?.preserve, true);
    assert.equal(zeroValidated.masters?.[0]?.layouts[0]?.preserve, true);
    assert.equal(zeroValidated.masters?.[1]?.layouts[0]?.preserve, true);

    // Separate independent editable instance: select Master 2's idle Blank layout and add a slide.
    const editableRaw = readMasterFixture("prs-slide-masters.pptx");
    const editable = importDocument(editableRaw);
    editableRaw.fill(0);
    assert.ok(editable.document instanceof DeckContainer);
    const master2 = editable.document.masters[1];
    assert.ok(master2 instanceof SlideMasterContainer);
    const idleLayout2 = master2.layouts[0];
    assert.ok(idleLayout2 instanceof SlideLayoutContainer);

    editable.document.append(
      slide({ id: "added-from-idle", layout: idleLayout2 }).append(
        text("Created from Master 2 idle layout", {
          id: "idle-heading",
          style: {
            position: "absolute",
            left: 48,
            top: 48,
            width: 480,
            height: 56,
            fontFamily: jostFont,
            fontSize: "24pt",
            color: "#0F172A",
          },
        }),
      ),
    );

    const rebuiltResolution = resolveDocument(editable.document);
    const rebuiltEmit = emit(rebuiltResolution);
    const rebuiltValidation = await validatePptx(rebuiltEmit.pptx);
    assert.deepEqual(rebuiltValidation.errors, []);
    assert.equal(rebuiltValidation.presentation.slideCount, 1);
    assert.equal(rebuiltValidation.slides[0]?.id, "added-from-idle");
    assert.equal(
      rebuiltValidation.slides[0]?.layoutPart,
      rebuiltValidation.masters?.[1]?.layouts[0]?.part,
    );
    assert.equal(
      rebuiltValidation.slides[0]?.masterPart,
      rebuiltValidation.masters?.[1]?.part,
    );

    // Extra font environment observation (explicit Jost* TTF on the newly added Text node).
    const painted = paint(rebuiltResolution);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(
      painted.pages[0]?.errors,
      [],
      "Extra font environment observation (explicit Jost* TTF on added slide)",
    );
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);
  });

  it("imports external WithMaster.pptx preserving non-flattened shared objects, bgRef, footer, and independent edit propagation", async () => {
    const raw = readMasterFixture("WithMaster.pptx");
    const imported = importDocument(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 2);
    assert.ok(imported.document instanceof DeckContainer);
    const master = imported.document.masters[0]!;
    assert.deepEqual(
      { ...(master.props["background"] as Record<string, unknown>) },
      { ref: 1001, color: "bg1" },
    );
    assert.equal(master.layouts[0]?.children.length, 1);
    const sharedText = master.layouts[0]?.children[0] as CommonNode;
    assert.equal(sharedText.props["name"], "TextBox 6");

    // Pure-file roundtrip: shared object stays in slideLayout1.xml, never flattened into slides.
    const pureResolution = resolveDocument(imported.document);
    const pureEmit = emit(pureResolution);
    const pureValidation = await validatePptx(pureEmit.pptx);
    assert.deepEqual(pureValidation.errors, []);
    assert.equal(pureValidation.masters?.[0]?.layouts[0]?.objects?.length, 1);
    assert.equal(
      pureValidation.masters?.[0]?.layouts[0]?.objects?.[0]?.name,
      "TextBox 6",
    );
    for (const s of pureValidation.slides) {
      assert.equal(
        s.objects.some((obj) => obj.name === "TextBox 6"),
        false,
        "Shared layout object TextBox 6 must not be flattened into ordinary slides",
      );
      assert.equal(
        s.objects.every((obj) => obj.placeholder !== undefined),
        true,
      );
    }
    const pureZip = unzipSync(pureEmit.pptx);
    assert.equal(
      strFromU8(pureZip["ppt/slides/slide1.xml"]!).includes("TextBox 6"),
      false,
    );
    assert.equal(
      strFromU8(pureZip["ppt/slides/slide2.xml"]!).includes("TextBox 6"),
      false,
    );
    assert.deepEqual(
      pureValidation.slides[1]?.footerPlaceholders?.map((p) => p.kind),
      ["ftr"],
    );
    assert.equal(
      pureValidation.slides[1]?.footerPlaceholders?.[0]?.text,
      "Footer from the master slide",
    );

    // Pure-file paint without embedded fonts reports explicit monochrome outline font file.
    const purePaint = paint(pureResolution);
    assert.equal(purePaint.pages.length, 2);
    for (const page of purePaint.pages) {
      assert.equal(page.png, undefined);
      assert.ok(
        page.errors.some((err) => err.code === "preview-missing-font"),
        "Pure-file WithMaster.pptx paint without embedded font must report preview-missing-font",
      );
    }

    // Independent editable instance: move shared TextBox 6 and supply explicit Jost* TTF for extra font environment observation.
    const editRaw = readMasterFixture("WithMaster.pptx");
    const editable = importDocument(editRaw);
    editRaw.fill(0);
    assert.ok(editable.document instanceof DeckContainer);
    const editMaster = editable.document.masters[0]!;
    const editLayout = editMaster.layouts[0]!;
    const editSharedText = editLayout.children[0] as CommonNode;
    const nextX = 80;
    editSharedText.style({ left: nextX });

    // Attach explicit test font family (400/700 normal & italic) to master theme & nodes for extra font environment observation.
    const editTheme = editMaster.theme as Theme | undefined;
    if (editTheme !== undefined) {
      editMaster.set({
        theme: {
          ...editTheme,
          fonts: { major: jostFullFamilyFont, minor: jostFullFamilyFont },
        },
      });
    }
    editSharedText.style({ fontFamily: jostFullFamilyFont });
    for (const s of collectSlides(editable.document)) {
      for (const child of s.children) {
        if (child instanceof CommonNode) {
          child.style({ fontFamily: jostFullFamilyFont });
        }
      }
    }

    const editedResolution = resolveDocument(editable.document);
    const editedEmit = emit(editedResolution);
    const editedValidation = await validatePptx(editedEmit.pptx);
    assert.deepEqual(editedValidation.errors, []);
    assert.equal(
      editedValidation.masters?.[0]?.layouts[0]?.objects?.[0]?.boxEmu.x,
      nextX * 9525,
    );
    const editedPaint = paint(editedResolution);
    assert.equal(editedPaint.pages.length, 2);
    for (const page of editedPaint.pages) {
      assert.deepEqual(
        page.errors,
        [],
        `Extra font environment observation (WithMaster.pptx + Jost* TTF) page ${page.page}`,
      );
      assert.ok(page.png instanceof Uint8Array);
    }
  });

  it("imports external layouts.pptx and placeholder-layout-color.pptx preserving idle layouts and layout color/size inheritance", async () => {
    const layoutsRaw = readMasterFixture("layouts.pptx");
    const layoutsImported = importDocument(layoutsRaw);
    layoutsRaw.fill(0);
    assert.equal(layoutsImported.report.slideCount, 10);
    assert.ok(layoutsImported.document instanceof DeckContainer);
    assert.equal(layoutsImported.document.masters[0]?.layouts.length, 11);

    const layoutSlides = collectSlides(layoutsImported.document);
    // Slide 5 (Comparison layout) preserves deterministic body-1..content-4 keys and <p:spPr/> (no local box).
    const slide5 = layoutSlides[4]!;
    assert.ok(slide5.placeholder("body-1") !== undefined);
    assert.ok(slide5.placeholder("content-2") !== undefined);
    assert.ok(slide5.placeholder("body-3") !== undefined);
    assert.ok(slide5.placeholder("content-4") !== undefined);
    assert.equal(layoutSlides[0]?.props["footer"], undefined);
    assert.equal(layoutSlides[0]?.props["slideNumber"], undefined);
    assert.equal(
      (
        layoutSlides[9]?.props["dateTime"] as
          | { readonly text?: string }
          | undefined
      )?.text,
      "Friday, October 21, 2011",
    );
    assert.equal(
      (
        layoutSlides[9]?.props["footer"] as
          | { readonly text?: string }
          | undefined
      )?.text,
      "Apache Software Foundation",
    );
    assert.ok(layoutSlides[9]?.props["slideNumber"] !== undefined);

    // Slide 7 (Blank with <p:ph type="title" idx="4294967295"/>) falls back to master title placeholder.
    const layoutsResolved = resolveDocument(layoutsImported.document);
    assert.equal(
      layoutsResolved.report.slides[6]?.objects[0]?.text,
      "Blank with Default Title",
    );
    const layoutsEmit = emit(layoutsResolved);
    const layoutsValidated = await validatePptx(layoutsEmit.pptx);
    assert.deepEqual(layoutsValidated.errors, []);
    assert.equal(layoutsValidated.masters?.[0]?.layouts.length, 11);
    for (const layout of layoutsValidated.masters?.[0]?.layouts ?? []) {
      assert.equal(layout.preserve, true);
    }
    assert.equal(layoutsValidated.slides.length, 10);
    assert.equal(
      layoutsValidated.slides[4]?.objects[1]?.hasLocalTransform,
      false,
    );
    assert.equal(
      layoutsValidated.slides[6]?.objects[0]?.hasLocalTransform,
      false,
    );
    assert.deepEqual(
      layoutsValidated.slides[9]?.footerPlaceholders?.map((p) => p.kind),
      ["dt", "ftr", "sldNum"],
    );
    assert.equal(
      layoutsValidated.slides[9]?.footerPlaceholders?.[2]?.hasSlideNumberField,
      true,
    );
    const layoutsReimported = importDocument(layoutsEmit.pptx);
    assert.deepEqual(layoutsReimported.report.diagnostics, []);

    const plcRaw = readMasterFixture("placeholder-layout-color.pptx");
    const plcImported = importDocument(plcRaw);
    plcRaw.fill(0);
    assert.ok(plcImported.document instanceof DeckContainer);
    const plcImportedMaster = plcImported.document.masters[0]!;
    assert.equal((plcImportedMaster.theme as Theme).colors.dk1, "#2D2842");
    assert.deepEqual(
      { ...(plcImportedMaster.props["colorMap"] as Record<string, string>) },
      { ...defaultColorMap, bg2: "dk2", tx2: "lt2" },
    );
    const plcResolution = resolveDocument(plcImported.document);
    const plcEmit = emit(plcResolution);
    const plcValidated = await validatePptx(plcEmit.pptx);
    assert.deepEqual(plcValidated.errors, []);
    assert.equal(plcValidated.slides[0]?.objects[0]?.hasLocalTransform, true);
    assert.equal(plcValidated.slides[0]?.objects[0]?.text?.text, "CYAN TEXT");
    const plcZip = unzipSync(plcEmit.pptx);
    assert.match(
      strFromU8(plcZip["ppt/slideLayouts/slideLayout1.xml"]!),
      /sz="3500".*00FFFF/u,
    );
    assert.doesNotMatch(
      strFromU8(plcZip["ppt/slides/slide1.xml"]!),
      /00FFFF|sz="3500"/u,
    );
    const plcReimported = importDocument(plcEmit.pptx);
    assert.deepEqual(plcReimported.report.diagnostics, []);

    // Pure-file paint reports unembedded font error; separate editable instance with Jost* TTF paints cleanly.
    const plcPurePaint = paint(plcResolution);
    assert.equal(plcPurePaint.pages[0]?.png, undefined);
    assert.ok(
      plcPurePaint.pages[0]?.errors.some(
        (err) => err.code === "preview-missing-font",
      ),
    );

    const plcEditRaw = readMasterFixture("placeholder-layout-color.pptx");
    const plcEditable = importDocument(plcEditRaw);
    plcEditRaw.fill(0);
    assert.ok(plcEditable.document instanceof DeckContainer);
    const plcMaster = plcEditable.document.masters[0]!;
    const plcTheme = plcMaster.theme as Theme | undefined;
    if (plcTheme !== undefined) {
      plcMaster.set({
        theme: { ...plcTheme, fonts: { major: jostFont, minor: jostFont } },
      });
    }
    const plcExtraPaint = paint(resolveDocument(plcEditable.document));
    assert.deepEqual(
      plcExtraPaint.pages[0]?.errors,
      [],
      "Extra font environment observation (placeholder-layout-color.pptx + Jost* TTF)",
    );
    assert.ok(plcExtraPaint.pages[0]?.png instanceof Uint8Array);
  });

  it("isolates a corrupt slide shape with exact fault type, location, and disposition while preserving all healthy slides, objects, and masters", async () => {
    const cleanLayout = defineSlideLayout({
      id: "clean-layout",
      name: "Clean Layout",
      type: "blank",
      preserve: true,
    });
    const sourceDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m-healthy",
        name: "Healthy Master",
        preserve: true,
      }).append(slideLayout(cleanLayout)),
      slide({ id: "s1", layout: cleanLayout }).append(
        shape({
          id: "s1-corrupt-target",
          name: "Corrupt Target",
          style: {
            position: "absolute",
            left: 32,
            top: 32,
            width: 120,
            height: 60,
            background: "#EF4444",
          },
        }),
        shape({
          id: "s1-healthy-sibling",
          name: "Healthy Sibling",
          style: {
            position: "absolute",
            left: 200,
            top: 32,
            width: 120,
            height: 60,
            background: "#10B981",
          },
        }),
      ),
      slide({ id: "s2", layout: cleanLayout }).append(
        shape({
          id: "s2-healthy",
          name: "Healthy Slide 2 Shape",
          style: {
            position: "absolute",
            left: 64,
            top: 64,
            width: 160,
            height: 80,
            background: "#2563EB",
          },
        }),
      ),
    );

    const basePptx = emit(resolveDocument(sourceDoc)).pptx;
    const entries = unzipSync(basePptx);
    const slide1Xml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    // Corrupt only the first shape on slide 1 by stripping its required <p:spPr>...</p:spPr>.
    const corruptedSlide1Xml = slide1Xml.replace(/<p:spPr>.*?<\/p:spPr>/u, "");
    assert.notEqual(corruptedSlide1Xml, slide1Xml);
    entries["ppt/slides/slide1.xml"] = strToU8(corruptedSlide1Xml);
    const corruptedBytes = new Uint8Array(zipSync(entries));

    const imported = importDocument(corruptedBytes);
    corruptedBytes.fill(0);

    // Exact fault type, location, and disposition check.
    expect(imported.report.diagnostics).toHaveLength(1);
    const diag = imported.report.diagnostics[0]!;
    expect(diag.code).toBe("corrupt-part");
    expect(diag.category).toBe("corrupt-input");
    expect(diag.disposition).toBe("omitted");
    expect(diag.location.part).toBe("ppt/slides/slide1.xml");
    expect(diag.location.slideIndex).toBe(0);
    expect(diag.location.xpath).toMatch(/\/p:sp\[1\]$/u);

    // Exact healthy masters, slides, and objects check (any unexpected extra loss fails immediately).
    const recoveredResolution = resolveDocument(imported.document);
    const recoveredEmit = emit(recoveredResolution);
    const recoveredValidation = await validatePptx(recoveredEmit.pptx);
    assert.deepEqual(recoveredValidation.errors, []);
    expect(recoveredValidation.masters).toHaveLength(1);
    expect(recoveredValidation.masters?.[0]?.layouts).toHaveLength(1);
    expect(recoveredValidation.slides).toHaveLength(2);
    expect(recoveredValidation.slides[0]?.objects).toHaveLength(1);
    expect(recoveredValidation.slides[0]?.objects[0]?.name).toBe(
      "Healthy Sibling",
    );
    expect(recoveredValidation.slides[1]?.objects).toHaveLength(1);
    expect(recoveredValidation.slides[1]?.objects[0]?.name).toBe(
      "Healthy Slide 2 Shape",
    );

    const recoveredPaint = paint(recoveredResolution);
    expect(recoveredPaint.pages).toHaveLength(2);
    expect(recoveredPaint.pages[0]?.errors).toEqual([]);
    expect(recoveredPaint.pages[1]?.errors).toEqual([]);
    expect(readPngPixelHex(recoveredPaint.pages[0]?.png, 260, 62)).toBe(
      "#10B981",
    );
    expect(readPngPixelHex(recoveredPaint.pages[1]?.png, 144, 104)).toBe(
      "#2563EB",
    );
  });

  it("roundtrips master, layout, and slide footer placeholder letterSpacing (including empty-text endParaRPr slots) with zero unconsumed-xml diagnostics", async () => {
    const layoutWithFooterLetterSpacing = defineSlideLayout({
      id: "footer-ls-layout",
      name: "Footer LetterSpacing Layout",
      type: "obj",
      preserve: true,
      footerPlaceholders: {
        footer: {
          text: "Layout Footer",
          style: {
            position: "absolute",
            left: 200,
            top: 316,
            width: 240,
            height: 24,
            fontFamily: jostFont,
            fontSize: 12,
            color: "#334155",
            letterSpacing: 2.4,
          },
        },
      },
    });

    const sourceDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m-footer-ls",
        name: "Footer LetterSpacing Master",
        preserve: true,
        footerPlaceholders: {
          dateTime: {
            text: "2026-11-27",
            style: {
              position: "absolute",
              left: 36,
              top: 316,
              width: 150,
              height: 24,
              fontFamily: jostFont,
              fontSize: 12,
              color: "#475569",
              letterSpacing: 2,
            },
          },
          footer: {
            style: {
              position: "absolute",
              left: 200,
              top: 316,
              width: 240,
              height: 24,
              fontFamily: jostFont,
              fontSize: 12,
              color: "#475569",
              letterSpacing: -1.2,
            },
          },
          slideNumber: {
            style: {
              position: "absolute",
              left: 490,
              top: 316,
              width: 110,
              height: 24,
              fontFamily: jostFont,
              fontSize: 12,
              color: "#0F172A",
              letterSpacing: 0.8,
            },
          },
        },
      }).append(slideLayout(layoutWithFooterLetterSpacing)),
      slide({
        id: "s-footer-inherited",
        layout: layoutWithFooterLetterSpacing,
      }),
      slide({
        id: "s-footer-override",
        layout: layoutWithFooterLetterSpacing,
        dateTime: {
          text: "Q4 2026",
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#0F172A",
            letterSpacing: 1.6,
          },
        },
        footer: {
          // Empty text emits <a:pPr><a:defRPr> and <a:endParaRPr> without <a:r>.
          text: "",
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#0F172A",
            letterSpacing: -1.2,
          },
        },
        slideNumber: {
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#0F172A",
            letterSpacing: 3.2,
          },
        },
      }),
    );

    const emitted = emit(resolveDocument(sourceDoc), { embedFonts: "subset" });
    const imported = importDocument(emitted.pptx);

    assert.deepEqual(imported.report.diagnostics, []);
    assert.ok(imported.document instanceof DeckContainer);

    const masterPlaceholders = imported.document.masters[0]?.props[
      "placeholders"
    ] as
      | Record<string, { listStyles?: Record<number, Record<string, unknown>> }>
      | undefined;
    assert.equal(
      masterPlaceholders?.["dt"]?.listStyles?.[1]?.["letterSpacing"],
      2,
    );

    const layoutPlaceholders = imported.document.masters[0]?.layouts[0]?.props[
      "placeholders"
    ] as
      | Record<string, { listStyles?: Record<number, Record<string, unknown>> }>
      | undefined;
    assert.equal(
      layoutPlaceholders?.["ftr"]?.listStyles?.[1]?.["letterSpacing"],
      2.4,
    );

    const importedSlides = collectSlides(imported.document);
    const inheritedSlide = importedSlides[0]!;
    assert.equal(
      (inheritedSlide.props["dateTime"] as { style?: Record<string, unknown> })
        ?.style?.["letterSpacing"],
      2,
    );
    assert.equal(
      (inheritedSlide.props["footer"] as { style?: Record<string, unknown> })
        ?.style?.["letterSpacing"],
      2.4,
    );
    assert.equal(
      (
        inheritedSlide.props["slideNumber"] as {
          style?: Record<string, unknown>;
        }
      )?.style?.["letterSpacing"],
      0.8,
    );

    const overrideSlide = importedSlides[1]!;
    assert.equal(
      (overrideSlide.props["dateTime"] as { style?: Record<string, unknown> })
        ?.style?.["letterSpacing"],
      1.6,
    );
    assert.equal(
      (overrideSlide.props["footer"] as { style?: Record<string, unknown> })
        ?.style?.["letterSpacing"],
      -1.2,
    );
    assert.equal(
      (
        overrideSlide.props["slideNumber"] as {
          style?: Record<string, unknown>;
        }
      )?.style?.["letterSpacing"],
      3.2,
    );

    const rebuilt = emit(resolveDocument(imported.document), {
      embedFonts: "subset",
    });
    const validated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validated.errors, []);

    const reimported = importDocument(rebuilt.pptx);
    assert.deepEqual(reimported.report.diagnostics, []);
  });

  it("isolates invalid and out-of-range @spc on footer placeholders with invalid-xml-value while preserving sibling attributes and slide content", async () => {
    const cleanLayout = defineSlideLayout({
      id: "footer-invalid-spc-layout",
      name: "Footer Invalid Spc Layout",
      type: "blank",
      preserve: true,
    });
    const sourceDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m-invalid-spc",
        name: "Invalid Spc Master",
        preserve: true,
      }).append(slideLayout(cleanLayout)),
      slide({
        id: "s-invalid-spc-1",
        layout: cleanLayout,
        dateTime: {
          text: "2026-11-27",
          style: {
            fontFamily: jostFullFamilyFont,
            fontSize: 14,
            fontWeight: "bold",
            fontStyle: "italic",
            color: "#1D4ED8",
            letterSpacing: 2,
          },
        },
        footer: {
          text: "Healthy Footer",
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#0F172A",
            letterSpacing: 1.6,
          },
        },
        slideNumber: {
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#0F172A",
            letterSpacing: 0.8,
          },
        },
      }).append(
        shape({
          id: "healthy-shape",
          name: "Healthy Sibling Shape",
          style: {
            position: "absolute",
            left: 48,
            top: 48,
            width: 160,
            height: 60,
            background: "#10B981",
          },
        }),
      ),
      slide({
        id: "s-invalid-spc-2",
        layout: cleanLayout,
        footer: {
          text: "",
          style: {
            fontFamily: jostFont,
            fontSize: 12,
            color: "#334155",
            letterSpacing: -1.2,
          },
        },
      }),
    );

    const basePptx = emit(resolveDocument(sourceDoc)).pptx;
    const entries = unzipSync(basePptx);
    const slide1Xml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    const mutatedSlide1Xml = slide1Xml
      .replace('spc="150"', 'spc="12.5"')
      .replace('spc="60"', 'spc="400001"');
    assert.notEqual(mutatedSlide1Xml, slide1Xml);
    entries["ppt/slides/slide1.xml"] = strToU8(mutatedSlide1Xml);

    const slide2Xml = strFromU8(entries["ppt/slides/slide2.xml"]!);
    // Strip pPr/defRPr so empty-paragraph endParaRPr is the sole run style source, and set out-of-range spc.
    const mutatedSlide2Xml = slide2Xml
      .replace(/<a:pPr>.*?<\/a:pPr>/u, "")
      .replace('spc="-90"', 'spc="-400001"');
    assert.notEqual(mutatedSlide2Xml, slide2Xml);
    entries["ppt/slides/slide2.xml"] = strToU8(mutatedSlide2Xml);

    const mutatedBytes = new Uint8Array(zipSync(entries));
    const imported = importDocument(mutatedBytes);
    mutatedBytes.fill(0);

    assert.equal(imported.report.diagnostics.length, 3);
    const [diagDt, diagSldNum, diagS2Ftr] = imported.report.diagnostics;
    assert.deepEqual(
      {
        code: diagDt?.code,
        category: diagDt?.category,
        disposition: diagDt?.disposition,
        attributeOrElement:
          diagDt?.code === "invalid-xml-value"
            ? diagDt.attributeOrElement
            : undefined,
        rawValue:
          diagDt?.code === "invalid-xml-value" ? diagDt.rawValue : undefined,
        part: diagDt?.location.part,
        slideIndex: diagDt?.location.slideIndex,
      },
      {
        code: "invalid-xml-value",
        category: "corrupt-input",
        disposition: "omitted",
        attributeOrElement: "@spc",
        rawValue: "12.5",
        part: "ppt/slides/slide1.xml",
        slideIndex: 0,
      },
    );
    assert.ok(diagDt?.location.xpath?.endsWith("/a:rPr/@spc"));

    assert.equal(diagSldNum?.code, "invalid-xml-value");
    assert.equal(diagSldNum?.location.slideIndex, 0);
    if (diagSldNum?.code === "invalid-xml-value") {
      assert.equal(diagSldNum.attributeOrElement, "@spc");
      assert.equal(diagSldNum.rawValue, "400001");
    }
    assert.ok(diagSldNum?.location.xpath?.endsWith("/a:fld/a:rPr/@spc"));

    assert.equal(diagS2Ftr?.code, "invalid-xml-value");
    assert.equal(diagS2Ftr?.location.slideIndex, 1);
    if (diagS2Ftr?.code === "invalid-xml-value") {
      assert.equal(diagS2Ftr.attributeOrElement, "@spc");
      assert.equal(diagS2Ftr.rawValue, "-400001");
    }
    assert.ok(diagS2Ftr?.location.xpath?.endsWith("/a:endParaRPr/@spc"));

    const importedSlides = collectSlides(imported.document);
    const dtStyle = (
      importedSlides[0]?.props["dateTime"] as {
        style?: Record<string, unknown>;
      }
    )?.style;
    assert.equal(dtStyle?.["letterSpacing"], undefined);
    assert.equal(dtStyle?.["fontSize"], "10.5pt");
    assert.equal(dtStyle?.["fontWeight"], "bold");
    assert.equal(dtStyle?.["fontStyle"], "italic");
    assert.equal(dtStyle?.["color"], "#1D4ED8");
    assert.equal(dtStyle?.["fontFamily"], "Jost*");

    const ftrStyle = (
      importedSlides[0]?.props["footer"] as { style?: Record<string, unknown> }
    )?.style;
    assert.equal(ftrStyle?.["letterSpacing"], 1.6);

    const s2FtrStyle = (
      importedSlides[1]?.props["footer"] as { style?: Record<string, unknown> }
    )?.style;
    assert.equal(s2FtrStyle?.["letterSpacing"], undefined);
    assert.equal(s2FtrStyle?.["fontSize"], "9pt");
    assert.equal(s2FtrStyle?.["color"], "#334155");

    const rebuilt = emit(resolveDocument(imported.document));
    const validated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(
      validated.slides[0]?.objects[0]?.name,
      "Healthy Sibling Shape",
    );
  });

  it("preserves inherited eastAsianFontFamily across layout placeholder defRPr emit, paint, and import roundtrip when only fontFamily is declared", async () => {
    const dualScriptTheme = defineTheme({
      name: "DualScript Master Theme",
      fonts: {
        major: jostFont,
        minor: jostFont,
        majorEastAsian: notoScFont,
        minorEastAsian: notoScFont,
      },
    });

    const dualLayout = defineSlideLayout({
      id: "dual-script-layout",
      name: "Dual Script Layout",
      type: "obj",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 48,
            top: 32,
            width: 544,
            height: 64,
            fontFamily: jostFont,
            fontSize: "24pt",
          },
        },
        body: {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 48,
            top: 120,
            width: 544,
            height: 180,
          },
          listStyles: {
            1: {
              fontFamily: "major",
              fontSize: "16pt",
            },
          },
        },
      },
    });

    const doc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "dual-master",
        name: "Dual Master",
        theme: dualScriptTheme,
        preserve: true,
        textStyles: {
          titleStyle: {
            1: { eastAsianFontFamily: "major" },
          },
          bodyStyle: {
            1: { eastAsianFontFamily: "minor" },
          },
        },
      }).append(slideLayout(dualLayout)),
      slide({
        id: "dual-slide",
        layout: dualLayout,
        placeholders: {
          title: text("Latin Title 中文标题"),
          body: text("Body ABC 中文正文"),
        },
      }),
    );

    const resolved = resolveDocument(doc);
    const emitted = emit(resolved);
    const painted = paint(resolved);
    const validated = await validatePptx(emitted.pptx);
    assert.deepEqual(validated.errors, []);
    assert.deepEqual(painted.pages[0]?.errors, []);

    const zipEntries = unzipSync(emitted.pptx);
    const layoutXmlStr = strFromU8(
      zipEntries["ppt/slideLayouts/slideLayout1.xml"]!,
    );
    assert.match(layoutXmlStr, /<a:latin typeface="Jost\*"/u);
    assert.doesNotMatch(layoutXmlStr, /<a:ea typeface="Jost\*"/u);
    assert.match(layoutXmlStr, /<a:latin typeface="\+mj-lt"\/>/u);
    assert.doesNotMatch(layoutXmlStr, /<a:ea typeface="\+mj-ea"\/>/u);

    const imported = importDocument(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const importedDeck = imported.document as DeckContainer;
    const importedMaster = importedDeck.masters[0]!;
    assert.equal(importedMaster.theme?.fonts?.major, "Jost*");
    assert.equal(
      importedMaster.theme?.fonts?.majorEastAsian,
      "Noto Sans CJK SC",
    );
    assert.equal(
      importedMaster.theme?.fonts?.minorEastAsian,
      "Noto Sans CJK SC",
    );
    const importedLayout = importedMaster.layouts[0]!;
    (importedMaster.theme as { set?(patch: unknown): void } | undefined)?.set?.(
      {
        fonts: {
          major: jostFont,
          minor: jostFont,
          majorEastAsian: notoScFont,
          minorEastAsian: notoScFont,
        },
      },
    );
    const importedPlaceholders = importedLayout.props[
      "placeholders"
    ] as Readonly<Record<string, LayoutPlaceholderSpec>>;
    const importedTitleSpec = importedPlaceholders["title"];
    const importedBodySpec = importedPlaceholders["body"];
    assert.ok(importedTitleSpec !== undefined);
    assert.ok(importedBodySpec !== undefined);
    assert.equal(importedTitleSpec.style?.eastAsianFontFamily, undefined);
    assert.equal(
      importedBodySpec.listStyles?.[1]?.eastAsianFontFamily,
      undefined,
    );
    (importedLayout as unknown as { set(patch: unknown): void }).set({
      placeholders: {
        ...importedPlaceholders,
        title: {
          ...importedTitleSpec,
          style: {
            ...importedTitleSpec.style,
            fontFamily: jostFont,
          },
          ...(importedTitleSpec.listStyles === undefined
            ? {}
            : {
                listStyles: {
                  ...importedTitleSpec.listStyles,
                  1: {
                    ...importedTitleSpec.listStyles[1],
                    fontFamily: jostFont,
                  },
                },
              }),
        },
      },
    });
    const importedResolved = resolveDocument(importedDeck);
    const importedPainted = paint(importedResolved);
    assert.deepEqual(importedPainted.pages[0]?.errors, []);
    assert.ok(
      importedPainted.pages[0]?.png !== undefined &&
        importedPainted.pages[0].png.byteLength > 0,
    );
  });
});
