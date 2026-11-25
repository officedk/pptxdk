import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  deck,
  defineSlideLayout,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
} from "../../document.ts";
import { shapeCapability } from "../../shape/shape.ts";
import { transitions } from "../transition-api.ts";

describe("transition capability codegen", () => {
  it("emits transitions.<kind>(...) and advanceAfter across SlideMaster, SlideLayout, and Slide for vanilla and react while keeping inherited slides clean", () => {
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
        fallback: transitions.push({ from: "left", duration: 400 }),
      }),
      advanceAfter: 3500,
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "brand-master",
        name: "Brand Master",
        transition: transitions.fade(),
        advanceAfter: 2500,
      }).append(slideLayout(inheritedLayout), slideLayout(morphLayout)),
      section({ name: "Empty" }),
      section({ name: "Showcase" }).append(
        slide({ id: "inherited-slide", layout: inheritedLayout }).append(
          shape({ id: "s1-box", x: 40, y: 40, width: 100, height: 60 }),
        ),
        slide({
          id: "fade-timed",
          layout: inheritedLayout,
          transition: transitions.fade({ duration: 500 }),
          advanceAfter: 1800,
        }),
        slide({
          id: "strips-slide",
          layout: inheritedLayout,
          transition: transitions.strips({
            direction: "leftUp",
            duration: 650,
          }),
        }),
        slide({
          id: "flip-slide",
          layout: morphLayout,
          transition: transitions.flip({
            direction: "left",
            duration: 800,
            fallback: transitions.strips({ direction: "rightDown" }),
          }),
        }),
      ),
    );

    const generate = createCodegen(documentCapabilities);
    for (const target of ["vanilla", "react"] as const) {
      const result = generate(doc, { target });
      const source = result.files[result.entry] as string;

      assert.match(source, /\btransitions\b/u, `${target} imports transitions`);
      assert.match(
        source,
        /transitions\.fade\(\)/u,
        `${target} concise fade()`,
      );
      assert.match(
        source,
        /transitions\.fade\(\{\s*duration:\s*500\s*\}\)/u,
        `${target} timed fade`,
      );
      assert.match(
        source,
        /transitions\.strips\(\{\s*duration:\s*650,\s*direction:\s*"leftUp"\s*\}\)/u,
        `${target} strips`,
      );
      assert.match(
        source,
        /transitions\.push\(\{\s*duration:\s*400,\s*from:\s*"left"\s*\}\)/u,
        `${target} push fallback`,
      );
      assert.match(source, /transitions\.flip\(/u, `${target} flip`);
      assert.match(
        source,
        /transitions\.morph\(/u,
        `${target} morph on layout`,
      );
      assert.match(
        source,
        /advanceAfter(?::\s*|=\{)2500/u,
        `${target} master advanceAfter`,
      );
      assert.match(
        source,
        /advanceAfter:\s*3500/u,
        `${target} layout advanceAfter`,
      );
      assert.match(
        source,
        /advanceAfter(?::\s*|=\{)1800/u,
        `${target} slide advanceAfter`,
      );

      if (target === "vanilla") {
        assert.match(
          source,
          /const inheritedSlide = slide\(\{\s*id:\s*"inherited-slide",\s*layout:\s*brandMasterInheritedLayout\s*\}\)/u,
        );
        assert.match(source, /section\(\{\s*name:\s*"Empty"\s*\}\)/u);
      } else {
        assert.match(
          source,
          /<Slide id="inherited-slide" layout=\{brandMasterInheritedLayout\}>/u,
        );
        assert.match(source, /<Section name="Empty" \/>/u);
      }
    }
  });

  it("reports CodegenError with exact owner path for invalid transition or advanceAfter on Slide, SlideLayout, and SlideMaster", () => {
    const generate = createCodegen(documentCapabilities);

    const invalidCases = [
      {
        label: "slide invalid transition object",
        doc: deck().append(
          slide({ id: "s1", transition: { kind: "fade" } as never }),
        ),
        expectedPath: "deck/slide[s1]@transition",
      },
      {
        label: "slide unsupported transition kind",
        doc: deck().append(
          slide({
            id: "s-kind",
            transition: { type: "transition", kind: "blinds" } as never,
          }),
        ),
        expectedPath: "deck/slide[s-kind]@transition",
      },
      {
        label: "slide morph missing fallback",
        doc: deck().append(
          slide({
            id: "s-morph",
            transition: {
              type: "transition",
              kind: "morph",
              by: "object",
            } as never,
          }),
        ),
        expectedPath: "deck/slide[s-morph]@transition",
      },
      {
        label: "slide flip with non-standard fallback",
        doc: deck().append(
          slide({
            id: "s-flip",
            transition: {
              type: "transition",
              kind: "flip",
              direction: "left",
              fallback: transitions.morph({ fallback: transitions.fade() }),
            } as never,
          }),
        ),
        expectedPath: "deck/slide[s-flip]@transition",
      },
      {
        label: "layout invalid transition duration",
        doc: deck().append(
          slideMaster({ id: "m1" }).append(
            slideLayout({
              id: "l1",
              transition: {
                type: "transition",
                kind: "fade",
                duration: Number.NaN,
              } as never,
            }),
          ),
        ),
        expectedPath: "deck/slide-master[m1]/slide-layout[l1]@transition",
      },
      {
        label: "master negative advanceAfter",
        doc: deck().append(
          slideMaster({ id: "m-adv", advanceAfter: -50 }).append(
            slideLayout({ id: "l1" }),
          ),
        ),
        expectedPath: "deck/slide-master[m-adv]@advanceAfter",
      },
      {
        label: "slide out-of-range advanceAfter",
        doc: deck().append(slide({ id: "s-adv", advanceAfter: 4_294_967_296 })),
        expectedPath: "deck/slide[s-adv]@advanceAfter",
      },
    ];

    for (const item of invalidCases) {
      for (const target of ["vanilla", "react"] as const) {
        assert.throws(
          () => generate(item.doc, { target }),
          (error: unknown) =>
            error instanceof CodegenError &&
            error.diagnostics[0]?.code === "unexpressible-value" &&
            error.diagnostics[0]?.path === item.expectedPath,
          `${item.label} (${target})`,
        );
      }
    }
  });

  it("reports unsupported-slide-property when transitionCapability is omitted", () => {
    const generateTrimmed = createCodegen([shapeCapability]);

    assert.throws(
      () =>
        generateTrimmed(
          deck().append(slide({ id: "s-tr", transition: transitions.fade() })),
          { target: "vanilla" },
        ),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-slide-property" &&
        error.diagnostics[0]?.path === "deck/slide[s-tr]@transition",
    );

    assert.throws(
      () =>
        generateTrimmed(
          deck().append(slide({ id: "s-adv", advanceAfter: 1000 })),
          { target: "react" },
        ),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-slide-property" &&
        error.diagnostics[0]?.path === "deck/slide[s-adv]@advanceAfter",
    );
  });
});
