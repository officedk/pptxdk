import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  BuildError,
  CommonNode,
  DeckContainer,
  SlideMasterContainer,
  build,
  defineSlideLayout,
  emit,
  font,
  importPptx,
  resolve,
} from "pptxdk";
import { preview } from "pptxdk/preview";
import { Deck, Slide, SlideLayout, SlideMaster, Text } from "pptxdk/react";
import { deck, slide, slideLayout, slideMaster, text } from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBytes }]);

describe("explicit SlideMaster and SlideLayout ownership", () => {
  it("supports zero-slide multi-master decks through both JSX and vanilla, omitting sldIdLst", async () => {
    const coverDef = defineSlideLayout({
      id: "cover",
      name: "Cover",
      type: "title",
      preserve: true,
    });
    const blankDef = defineSlideLayout({
      id: "blank",
      name: "Blank",
      type: "blank",
    });

    const jsxResult = await build(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          SlideMaster,
          { id: "m1", name: "Primary", preserve: true },
          createElement(SlideLayout, { layout: coverDef }),
        ),
        createElement(
          SlideMaster,
          { id: "m2", name: "Secondary" },
          createElement(SlideLayout, { layout: blankDef }),
        ),
      ),
    );

    const vanillaDeck = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1", name: "Primary", preserve: true }).append(
        slideLayout(coverDef),
      ),
      slideMaster({ id: "m2", name: "Secondary" }).append(
        slideLayout(blankDef),
      ),
    );
    const vanillaResult = await build(vanillaDeck);

    assert.deepEqual(jsxResult.report, vanillaResult.report);
    assert.equal(jsxResult.report.slides.length, 0);
    assert.equal(jsxResult.report.masters?.length, 2);

    const painted = await preview(vanillaDeck);
    assert.equal(painted.pages.length, 0);

    assert.equal(vanillaDeck.masters.length, 2);
    assert.equal(
      vanillaDeck.query.get("m1").get("cover").node,
      vanillaDeck.masters[0]?.layouts[0],
    );
    assert.equal(
      vanillaDeck.query.get("m2").get("blank").node,
      vanillaDeck.masters[1]?.layouts[0],
    );

    for (const bytes of [jsxResult.pptx, vanillaResult.pptx]) {
      const files = unzipSync(bytes);
      const presentationXml = strFromU8(files["ppt/presentation.xml"]!);
      assert.equal(presentationXml.includes("<p:sldIdLst"), false);

      const validated = await validatePptx(bytes);
      assert.deepEqual(validated.errors, []);
      assert.equal(validated.presentation.slideCount, 0);
      assert.equal(validated.masters?.length, 2);
      assert.equal(validated.masters?.[0]?.themePart, "ppt/theme/theme1.xml");
      assert.equal(validated.masters?.[1]?.themePart, "ppt/theme/theme2.xml");
      assert.equal(validated.masters?.[0]?.preserve, true);
      assert.equal(validated.masters?.[1]?.preserve, undefined);
      assert.equal(validated.masters?.[0]?.layouts[0]?.preserve, true);
      assert.equal(validated.masters?.[1]?.layouts[0]?.preserve, undefined);
    }
  });

  it("reports empty-slide-master, foreign-slide-layout, and ambiguous-slide-layout with clear diagnostics", async () => {
    const emptyMasterDeck = deck().append(slideMaster({ id: "empty-m" }));
    assert.throws(
      () => resolve(emptyMasterDeck),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (d) =>
            d.code === "empty-slide-master" &&
            d.path === "deck/slide-master[empty-m]",
        ),
    );

    const foreignLayout = slideLayout({ id: "other-layout", name: "Foreign" });
    const otherDeck = deck().append(
      slideMaster({ id: "other-master" }).append(foreignLayout),
    );
    void otherDeck;

    const hostDeck = deck().append(
      slideMaster({ id: "host-master" }).append(
        slideLayout({ id: "host-blank", type: "blank" }),
      ),
      slide({ id: "bad-slide", layout: foreignLayout }),
    );
    assert.throws(
      () => resolve(hostDeck),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (d) =>
            d.code === "foreign-slide-layout" &&
            d.path === "deck/slide[bad-slide]" &&
            d.message.includes("other-layout"),
        ),
    );

    const unmountedDef = defineSlideLayout({
      id: "unmounted-def",
      name: "Unmounted",
    });
    const plainDeckWithForeignDef = deck().append(
      slide({ id: "s1", layout: unmountedDef }),
    );
    assert.throws(
      () => resolve(plainDeckWithForeignDef),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (d) =>
            d.code === "foreign-slide-layout" &&
            d.path === "deck/slide[s1]" &&
            d.message.includes("unmounted-def"),
        ),
    );

    const sharedDef = defineSlideLayout({ id: "shared", name: "Shared" });
    const ambiguousDeck = deck().append(
      slideMaster({ id: "m1" }).append(slideLayout(sharedDef)),
      slideMaster({ id: "m2" }).append(slideLayout(sharedDef)),
      slide({ id: "s1", layout: sharedDef }),
    );
    assert.throws(
      () => resolve(ambiguousDeck),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (d) =>
            d.code === "ambiguous-slide-layout" && d.path === "deck/slide[s1]",
        ),
    );
  });

  it("imports prs-slide-masters.pptx, emits valid 0-slide PPTX, and supports adding a slide from an idle layout", async () => {
    const raw = new Uint8Array(
      readFileSync(
        new URL("./fixtures/prs-slide-masters.pptx", import.meta.url),
      ),
    );
    const imported = importPptx(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 0);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 2);

    const zeroBuilt = await build(imported.document);
    const zeroValidated = await validatePptx(zeroBuilt.pptx);
    assert.deepEqual(zeroValidated.errors, []);
    assert.equal(zeroValidated.presentation.slideCount, 0);
    assert.equal(zeroValidated.masters?.length, 2);
    assert.equal(zeroValidated.masters?.[0]?.preserve, undefined);
    assert.equal(zeroValidated.masters?.[1]?.preserve, true);
    assert.equal(zeroValidated.masters?.[0]?.layouts[0]?.preserve, true);
    assert.equal(zeroValidated.masters?.[1]?.layouts[0]?.preserve, true);

    // Add a new slide targeting Master 2's previously idle layout and build.
    const master2 = imported.document.masters[1]!;
    assert.ok(master2 instanceof SlideMasterContainer);
    const idleLayout = master2.layouts[0]!;
    imported.document.append(
      slide({ id: "added-cover", layout: idleLayout }).append(
        text("Created from idle layout", {
          id: "headline",
          x: 48,
          y: 48,
          width: 480,
          height: 60,
          style: { fontFamily: jostFont, fontSize: "24pt", color: "#1E293B" },
        }),
      ),
    );

    const rebuilt = await build(imported.document);
    const painted = await preview(imported.document);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);

    const rebuiltValidated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(rebuiltValidated.errors, []);
    assert.equal(rebuiltValidated.presentation.slideCount, 1);
    assert.equal(rebuiltValidated.slides[0]?.id, "added-cover");
    assert.equal(
      rebuiltValidated.slides[0]?.layoutPart,
      rebuiltValidated.masters?.[1]?.layouts[0]?.part,
    );
    assert.equal(
      rebuiltValidated.slides[0]?.masterPart,
      rebuiltValidated.masters?.[1]?.part,
    );
  });

  it("imports layouts.pptx, retains all 11 layouts (used and idle) and preserve flags across re-emit", async () => {
    const raw = new Uint8Array(
      readFileSync(new URL("./fixtures/layouts.pptx", import.meta.url)),
    );
    const imported = importPptx(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 10);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 1);
    const master = imported.document.masters[0]!;
    assert.equal(master.layouts.length, 11);

    const reemitted = emit(resolve(imported.document));
    const validated = await validatePptx(reemitted.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.masters?.length, 1);
    assert.equal(validated.masters?.[0]?.preserve, undefined);
    assert.equal(validated.masters?.[0]?.layouts.length, 11);
    for (const layout of validated.masters?.[0]?.layouts ?? []) {
      assert.equal(layout.preserve, true);
    }
    assert.equal(validated.slides.length, 10);
    assert.equal(
      validated.slides[0]?.layoutPart,
      validated.masters?.[0]?.layouts[0]?.part,
    );
    assert.equal(
      validated.slides[9]?.layoutPart,
      validated.masters?.[0]?.layouts[1]?.part,
    );

    const reimported = importPptx(reemitted.pptx);
    assert.deepEqual(reimported.report.diagnostics, []);
    assert.ok(reimported.document instanceof DeckContainer);
    assert.equal(reimported.document.masters[0]?.layouts.length, 11);

    const slides = imported.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.kind === "slide",
    );
    // Only slide 10 declares dt/ftr/sldNum in layouts.pptx.
    assert.equal(slides[0]?.props["footer"], undefined);
    assert.equal(slides[0]?.props["slideNumber"], undefined);
    assert.equal(
      (slides[9]?.props["dateTime"] as { readonly text?: string } | undefined)
        ?.text,
      "Friday, October 21, 2011",
    );
    assert.equal(
      (slides[9]?.props["footer"] as { readonly text?: string } | undefined)
        ?.text,
      "Apache Software Foundation",
    );
    assert.ok(slides[9]?.props["slideNumber"] !== undefined);

    assert.equal(validated.slides[0]?.footerPlaceholders, undefined);
    assert.deepEqual(
      validated.slides[9]?.footerPlaceholders?.map((p) => p.kind),
      ["dt", "ftr", "sldNum"],
    );
    assert.equal(
      validated.slides[9]?.footerPlaceholders?.[2]?.hasSlideNumberField,
      true,
    );
  });

  it("imports WithMaster.pptx preserving bgRef, layout shared TextBox, and slide2 footer across edit & re-emit", async () => {
    const raw = new Uint8Array(
      readFileSync(new URL("./fixtures/WithMaster.pptx", import.meta.url)),
    );
    const imported = importPptx(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 2);
    assert.ok(imported.document instanceof DeckContainer);
    const master = imported.document.masters[0]!;
    assert.deepEqual(
      { ...(master.props["background"] as Record<string, unknown>) },
      { ref: 1001, color: "bg1" },
    );

    const titleLayout = master.layouts[0]!;
    assert.equal(titleLayout.children.length, 1);
    const sharedText = titleLayout.children[0] as CommonNode;
    assert.equal(sharedText.props["name"], "TextBox 6");

    const slides = imported.document.children.filter(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.kind === "slide",
    );
    assert.equal(slides[0]?.props["footer"], undefined);
    assert.equal(
      (slides[1]?.props["footer"] as { readonly text?: string } | undefined)
        ?.text,
      "Footer from the master slide",
    );

    // Edit the shared TextBox position once on SlideLayout and re-emit.
    const nextX = 80;
    sharedText.set({ x: nextX });

    const reemitted = emit(resolve(imported.document));
    const validated = await validatePptx(reemitted.pptx);
    assert.deepEqual(validated.errors, []);

    const validatedLayout1 = validated.masters?.[0]?.layouts[0];
    assert.equal(validatedLayout1?.objects?.length, 1);
    assert.equal(validatedLayout1?.objects?.[0]?.name, "TextBox 6");
    assert.equal(validatedLayout1?.objects?.[0]?.boxEmu.x, nextX * 9525);
    // Shared layout object stays in slideLayout1.xml, not flattened into slides (slides only hold their own placeholders).
    assert.equal(
      validated.slides[0]?.objects.some((o) => o.name === "TextBox 6"),
      false,
    );
    assert.equal(
      validated.slides[1]?.objects.some((o) => o.name === "TextBox 6"),
      false,
    );
    assert.equal(
      validated.slides[0]?.objects.every((o) => o.placeholder !== undefined),
      true,
    );
    assert.equal(
      validated.slides[1]?.objects.every((o) => o.placeholder !== undefined),
      true,
    );
    assert.equal(validated.slides[0]?.footerPlaceholders, undefined);
    assert.deepEqual(
      validated.slides[1]?.footerPlaceholders?.map((p) => p.kind),
      ["ftr"],
    );
    assert.equal(
      validated.slides[1]?.footerPlaceholders?.[0]?.text,
      "Footer from the master slide",
    );
  });

  it("matches JSX and vanilla output when a Slide binds to an explicit SlideLayout alongside Text content", async () => {
    const titleLayout = defineSlideLayout({
      id: "title-slide",
      name: "Title Slide",
      type: "title",
      preserve: true,
    });
    const jsxOut = await build(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          SlideMaster,
          { id: "m1", name: "Brand", preserve: true },
          createElement(SlideLayout, { layout: titleLayout }),
        ),
        createElement(
          Slide,
          { id: "s1", layout: titleLayout },
          createElement(
            Text,
            {
              id: "t1",
              x: 40,
              y: 40,
              width: 300,
              height: 48,
              style: { fontFamily: jostFont, fontSize: "20pt" },
            },
            "Hello Master",
          ),
        ),
      ),
    );
    const docOut = await build(
      deck({ width: 640, height: 360 }).append(
        slideMaster({ id: "m1", name: "Brand", preserve: true }).append(
          slideLayout(titleLayout),
        ),
        slide({ id: "s1", layout: titleLayout }).append(
          text("Hello Master", {
            id: "t1",
            x: 40,
            y: 40,
            width: 300,
            height: 48,
            style: { fontFamily: jostFont, fontSize: "20pt" },
          }),
        ),
      ),
    );
    assert.deepEqual(jsxOut.report, docOut.report);
    const validated = await validatePptx(jsxOut.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(
      validated.slides[0]?.layoutPart,
      validated.masters?.[0]?.layouts[0]?.part,
    );
  });
});
