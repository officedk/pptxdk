import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  build,
  defineSlideLayout,
  font,
  resolve,
} from "@officedk/pptxdk";
import { preview } from "@officedk/pptxdk/preview";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  Text,
} from "@officedk/pptxdk/react";
import {
  deck,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
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
            d.code === "invalid-structure" &&
            d.path === "deck/slide-master[empty-m]",
        ),
    );

    const foreignLayout = slideLayout({ id: "other-layout", name: "Foreign" });
    deck().append(slideMaster({ id: "other-master" }).append(foreignLayout));

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
            d.code === "broken-reference" &&
            d.targetKind === "layout" &&
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
            d.code === "broken-reference" &&
            d.targetKind === "layout" &&
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
            d.code === "broken-reference" &&
            d.targetKind === "layout" &&
            d.path === "deck/slide[s1]",
        ),
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
              style: {
                position: "absolute",
                left: 40,
                top: 40,
                width: 300,
                height: 48,
                fontFamily: jostFont,
                fontSize: "20pt",
              },
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
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 300,
              height: 48,
              fontFamily: jostFont,
              fontSize: "20pt",
            },
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
