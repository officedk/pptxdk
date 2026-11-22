import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  defineSlideLayout,
  image,
  shape,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { Image } from "@pptxdk/capabilities/image-elements";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import {
  AdvanceAfterSchema,
  FlipTransitionSchema,
  MorphTransitionSchema,
  StandardTransitionSchema,
  StripsDirectionSchema,
  TransitionDirectionSchema,
  TransitionDurationSchema,
  TransitionSchema,
  TransitionSpeedSchema,
  transitionCapability,
  transitions,
} from "@pptxdk/capabilities/transition";
import {
  BuildError,
  CommonNode,
  DeckContainer,
  createMorphScope,
  targets,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import {
  A_NS,
  MC_NS,
  OPC_REL_NS,
  P_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strToU8, zipSync } from "fflate";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("Transition Zod schemas", () => {
  it("validates baseline transitions, flip, morph, directions, speeds, durations, and advanceAfter", () => {
    assert.equal(TransitionSpeedSchema.parse("med"), "med");
    assert.equal(TransitionDirectionSchema.parse("left"), "left");
    assert.equal(StripsDirectionSchema.parse("leftUp"), "leftUp");
    assert.equal(TransitionDurationSchema.parse(750), 750);
    assert.equal(AdvanceAfterSchema.parse(2000), 2000);

    const fade = transitions.fade({ duration: 500 });
    assert.deepEqual(StandardTransitionSchema.parse(fade), fade);

    const strips = transitions.strips({ direction: "rightUp" });
    assert.deepEqual(StandardTransitionSchema.parse(strips), strips);

    const push = transitions.push({ from: "left", duration: 700 });
    assert.deepEqual(StandardTransitionSchema.parse(push), push);

    const flip = transitions.flip({
      direction: "left",
      fallback: transitions.push({ from: "right" }),
    });
    assert.deepEqual(FlipTransitionSchema.parse(flip), flip);

    const morph = transitions.morph({
      duration: 1200,
      by: "word",
      fallback: transitions.fade(),
    });
    assert.deepEqual(MorphTransitionSchema.parse(morph), morph);
    assert.deepEqual(TransitionSchema.parse(morph), morph);

    assert.equal(TransitionDurationSchema.safeParse(-1).success, false);
    assert.equal(
      TransitionDurationSchema.safeParse(Number.POSITIVE_INFINITY).success,
      false,
    );
    assert.equal(AdvanceAfterSchema.safeParse(4_294_967_296).success, false);
    assert.equal(
      TransitionSchema.safeParse({
        type: "transition",
        kind: "flip",
        direction: "left",
        fallback: morph,
      }).success,
      false,
    );
  });

  it("preserves three-layer transition and advanceAfter ownership across SlideMaster, SlideLayout, and Slide, triggers Morph continuity checks, and leaves static paint unchanged", async () => {
    const resolve = createResolver(documentCapabilities);
    const scope = createMorphScope();
    const hero = scope.handle("hero");

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
        duration: 1000,
        by: "word",
        fallback: transitions.push({ from: "left" }),
      }),
      advanceAfter: 4000,
    });

    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
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
        createElement(
          Slide,
          { id: "s1", layout: inheritedLayout, background: "#FFFFFF" },
          createElement(Shape, {
            id: "hero-1",
            morphKey: hero,
            x: 40,
            y: 40,
            width: 100,
            height: 60,
            fill: "#2563EB",
          }),
        ),
        createElement(
          Slide,
          { id: "s2", layout: morphLayout, background: "#FFFFFF" },
          createElement(Shape, {
            id: "hero-2",
            morphKey: hero,
            x: 240,
            y: 80,
            width: 140,
            height: 80,
            fill: "#059669",
          }),
        ),
        createElement(
          Slide,
          {
            id: "s3",
            layout: morphLayout,
            background: "#FFFFFF",
            transition: transitions.strips({ direction: "leftUp" }),
            advanceAfter: 1500,
          },
          createElement(Shape, {
            id: "outro",
            x: 100,
            y: 100,
            width: 120,
            height: 60,
            fill: "#D97706",
          }),
        ),
      ),
    );

    const imperativeDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "brand-master",
        name: "Brand Master",
        transition: transitions.fade({ duration: 500 }),
        advanceAfter: 3000,
      }).append(slideLayout(inheritedLayout), slideLayout(morphLayout)),
      slide({
        id: "s1",
        layout: inheritedLayout,
        background: "#FFFFFF",
      }).append(
        shape({
          id: "hero-1",
          morphKey: hero,
          x: 40,
          y: 40,
          width: 100,
          height: 60,
          fill: "#2563EB",
        }),
      ),
      slide({ id: "s2", layout: morphLayout, background: "#FFFFFF" }).append(
        shape({
          id: "hero-2",
          morphKey: hero,
          x: 240,
          y: 80,
          width: 140,
          height: 80,
          fill: "#059669",
        }),
      ),
      slide({
        id: "s3",
        layout: morphLayout,
        background: "#FFFFFF",
        transition: transitions.strips({ direction: "leftUp" }),
        advanceAfter: 1500,
      }).append(
        shape({
          id: "outro",
          x: 100,
          y: 100,
          width: 120,
          height: 60,
          fill: "#D97706",
        }),
      ),
    );

    const jsxResolution = resolve(jsxDoc);
    const impResolution = resolve(imperativeDoc);
    assert.deepEqual(jsxResolution.report, impResolution.report);

    const jsxEmitted = emit(jsxResolution);
    const impEmitted = emit(impResolution);
    const jsxPainted = paint(jsxResolution);
    for (const page of jsxPainted.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png instanceof Uint8Array);
    }

    // Static paint matches an identical deck without any transitions.
    const staticDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "brand-master", name: "Brand Master" }).append(
        slideLayout({
          id: "inherited-layout",
          name: "Inherited Layout",
          type: "blank",
        }),
        slideLayout({ id: "morph-layout", name: "Morph Layout", type: "cust" }),
      ),
      slide({ id: "s1", background: "#FFFFFF" }).append(
        shape({
          id: "hero-1",
          morphKey: hero,
          x: 40,
          y: 40,
          width: 100,
          height: 60,
          fill: "#2563EB",
        }),
      ),
      slide({ id: "s2", background: "#FFFFFF" }).append(
        shape({
          id: "hero-2",
          morphKey: hero,
          x: 240,
          y: 80,
          width: 140,
          height: 80,
          fill: "#059669",
        }),
      ),
      slide({ id: "s3", background: "#FFFFFF" }).append(
        shape({
          id: "outro",
          x: 100,
          y: 100,
          width: 120,
          height: 60,
          fill: "#D97706",
        }),
      ),
    );
    const staticPainted = paint(resolve(staticDoc));
    for (let i = 0; i < 3; i += 1) {
      assert.equal(
        Buffer.compare(jsxPainted.pages[i]!.png!, staticPainted.pages[i]!.png!),
        0,
      );
    }

    const validated = await validatePptx(impEmitted.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.masters?.[0]?.advanceAfterMs, 3000);
    assert.deepEqual(validated.masters?.[0]?.transition, {
      kind: "fade",
      speed: "fast",
      durationMs: 500,
      fallback: { kind: "fade", speed: "fast" },
    });
    assert.equal(validated.masters?.[0]?.layouts[0]?.transition, undefined);
    assert.equal(validated.masters?.[0]?.layouts[0]?.advanceAfterMs, undefined);
    assert.equal(validated.masters?.[0]?.layouts[1]?.advanceAfterMs, 4000);
    assert.deepEqual(validated.masters?.[0]?.layouts[1]?.transition, {
      kind: "morph",
      speed: "slow",
      durationMs: 1000,
      by: "word",
      fallback: { kind: "push", speed: "slow", direction: "l" },
    });
    assert.equal(validated.slides[0]?.transition, undefined);
    assert.equal(validated.slides[0]?.advanceAfterMs, undefined);
    assert.equal(validated.slides[1]?.transition, undefined);
    assert.equal(validated.slides[1]?.advanceAfterMs, undefined);
    assert.equal(validated.slides[2]?.advanceAfterMs, 1500);
    assert.deepEqual(validated.slides[2]?.transition, {
      kind: "strips",
      speed: "slow",
      direction: "lu",
    });

    // baselineOnly target rejects master/layout p14/p159 transitions.
    assert.throws(
      () => emit(jsxResolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "forbidden-mechanism"),
    );

    // Effective Morph from SlideLayout triggers Morph continuity kind-matching on the slide.
    const layoutMorphMismatch = deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "m1" }).append(slideLayout(morphLayout)),
      slide({ id: "src", layout: morphLayout }).append(
        shape({ id: "a", morphKey: hero, x: 20, y: 20, width: 60, height: 40 }),
      ),
      slide({ id: "dst", layout: morphLayout }).append(
        image({
          id: "b",
          src: png1x1,
          morphKey: hero,
          x: 20,
          y: 20,
          width: 60,
          height: 40,
        }),
      ),
    );
    assert.throws(
      () => resolve(layoutMorphMismatch),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "morph-object-kind-mismatch"),
    );

    // Effective Morph from SlideMaster triggers Morph continuity kind-matching unless overridden on Slide.
    const masterMorphLayout = defineSlideLayout({ id: "plain", type: "blank" });
    const masterMorphMismatch = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m1",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(slideLayout(masterMorphLayout)),
      slide({ id: "src", layout: masterMorphLayout }).append(
        shape({ id: "a", morphKey: hero, x: 20, y: 20, width: 60, height: 40 }),
      ),
      slide({ id: "dst", layout: masterMorphLayout }).append(
        image({
          id: "b",
          src: png1x1,
          morphKey: hero,
          x: 20,
          y: 20,
          width: 60,
          height: 40,
        }),
      ),
    );
    assert.throws(
      () => resolve(masterMorphMismatch),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "morph-object-kind-mismatch"),
    );

    // Overriding master Morph with fade on the Slide suppresses Morph continuity kind-matching.
    const slideOverrideOk = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m1",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(slideLayout(masterMorphLayout)),
      slide({
        id: "src",
        layout: masterMorphLayout,
        transition: transitions.fade(),
      }).append(
        shape({ id: "a", morphKey: hero, x: 20, y: 20, width: 60, height: 40 }),
      ),
      slide({
        id: "dst",
        layout: masterMorphLayout,
        transition: transitions.fade(),
      }).append(
        image({
          id: "b",
          src: png1x1,
          morphKey: hero,
          x: 20,
          y: 20,
          width: 60,
          height: 40,
        }),
      ),
    );
    assert.equal(resolve(slideOverrideOk).report.slides.length, 2);

    // Trimmed resolver without transitionCapability rejects transition on SlideMaster & SlideLayout.
    const resolveNoTransition = createResolver([shapeCapability]);
    assert.throws(
      () =>
        resolveNoTransition(
          deck().append(
            slideMaster({ id: "m1", transition: transitions.fade() }).append(
              slideLayout({ id: "l1" }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "missing-capability"),
    );

    void jsxEmitted;
    void Image;
    void SlideMaster;
    void SlideLayout;
    void Slide;
    void Deck;
  });

  it("roundtrips three-layer transitions across emit -> import -> edit -> re-emit preserving exact ownership", async () => {
    const resolve = createResolver(documentCapabilities);
    const importDocument = createImporter(documentCapabilities);

    const inheritedLayout = defineSlideLayout({
      id: "inherited",
      name: "Inherited",
      type: "blank",
    });
    const morphLayout = defineSlideLayout({
      id: "morph-layout",
      name: "Morph Layout",
      type: "cust",
      transition: transitions.morph({
        duration: 1000,
        by: "word",
        fallback: transitions.push({ from: "left" }),
      }),
      advanceAfter: 4000,
    });

    const initialDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "brand-master",
        name: "Brand Master",
        transition: transitions.fade({ duration: 500 }),
        advanceAfter: 3000,
      }).append(slideLayout(inheritedLayout), slideLayout(morphLayout)),
      slide({ id: "s1", layout: inheritedLayout }).append(
        shape({ id: "box-1", x: 40, y: 40, width: 100, height: 60 }),
      ),
      slide({ id: "s2", layout: morphLayout }).append(
        shape({ id: "box-2", x: 140, y: 80, width: 100, height: 60 }),
      ),
      slide({
        id: "s3",
        layout: morphLayout,
        transition: transitions.flip({
          direction: "left",
          duration: 750,
          fallback: transitions.strips({ direction: "rightUp" }),
        }),
        advanceAfter: 1500,
      }).append(shape({ id: "box-3", x: 240, y: 120, width: 100, height: 60 })),
    );

    const initialBytes = new Uint8Array(emit(resolve(initialDoc)).pptx);
    const imported = importDocument(initialBytes);
    initialBytes.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);
    assert.ok(imported.document instanceof DeckContainer);
    const importedMaster = imported.document.masters[0]!;
    expect(importedMaster.props["transition"]).toEqual(
      transitions.fade({ duration: 500 }),
    );
    assert.equal(importedMaster.props["advanceAfter"], 3000);

    const [importedInherited, importedMorphLayout] = importedMaster.layouts;
    assert.equal(importedInherited?.props["transition"], undefined);
    assert.equal(importedInherited?.props["advanceAfter"], undefined);
    expect(importedMorphLayout?.props["transition"]).toEqual(
      transitions.morph({
        duration: 1000,
        by: "word",
        fallback: transitions.push({ from: "left" }),
      }),
    );
    assert.equal(importedMorphLayout?.props["advanceAfter"], 4000);

    const importedSlides = imported.document.children.filter(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    );
    assert.equal(importedSlides[0]?.props["transition"], undefined);
    assert.equal(importedSlides[0]?.props["advanceAfter"], undefined);
    assert.equal(importedSlides[1]?.props["transition"], undefined);
    assert.equal(importedSlides[1]?.props["advanceAfter"], undefined);
    expect(importedSlides[2]?.props["transition"]).toEqual(
      transitions.flip({
        direction: "left",
        duration: 750,
        fallback: transitions.strips({ direction: "rightUp" }),
      }),
    );
    assert.equal(importedSlides[2]?.props["advanceAfter"], 1500);

    // Edit each layer independently and verify ownership & updates survive re-emit.
    importedMaster.set({
      transition: transitions.push({ from: "top" }),
      advanceAfter: 3500,
    });
    importedMorphLayout?.set({
      advanceAfter: 5000,
    });
    importedSlides[2]?.set({
      transition: undefined,
      advanceAfter: undefined,
    });

    const reemitted = emit(resolve(imported.document));
    const revalidated = await validatePptx(reemitted.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(revalidated.errors, []);
    assert.deepEqual(revalidated.masters?.[0]?.transition, {
      kind: "push",
      speed: "slow",
      direction: "u",
    });
    assert.equal(revalidated.masters?.[0]?.advanceAfterMs, 3500);
    assert.equal(revalidated.masters?.[0]?.layouts[0]?.transition, undefined);
    assert.equal(revalidated.masters?.[0]?.layouts[1]?.advanceAfterMs, 5000);
    assert.deepEqual(revalidated.masters?.[0]?.layouts[1]?.transition, {
      kind: "morph",
      speed: "slow",
      durationMs: 1000,
      by: "word",
      fallback: { kind: "push", speed: "slow", direction: "l" },
    });
    assert.equal(revalidated.slides[2]?.transition, undefined);
    assert.equal(revalidated.slides[2]?.advanceAfterMs, undefined);
  });

  it("handles mc:AlternateContent transition fallbacks, preferred+fallback preservation, lexical units, and unsupported omissions", async () => {
    const p14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";
    const p159Ns =
      "http://schemas.microsoft.com/office/powerpoint/2015/09/main";

    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="${REL_SLIDE}" Target="slides/slide2.xml"/><Relationship Id="rId4" Type="${REL_SLIDE}" Target="slides/slide3.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      // Slide 1 (Case A): Supported Morph Choice (default option + 1.5s unit dur) + explicit Push Fallback -> reconstruct BOTH with zero loss diagnostics.
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="case-a"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${p14Ns}" xmlns:p159="${p159Ns}"><mc:Choice Requires="p14 p159"><p:transition spd="slow" p14:dur="1.5s" advTm="2200"><p159:morph/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med" advTm="2200"><p:push dir="l"/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      // Slide 2 (Case B): Unsupported p14:vortex Choice + supported p:push Fallback -> adopt fallback + emit compatibility-fallback.
      "ppt/slides/slide2.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="case-b"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${p14Ns}"><mc:Choice Requires="p14"><p:transition spd="med"><p14:vortex dir="l"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med" advTm="1800"><p:push dir="r"/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>`,
      ),
      "ppt/slides/_rels/slide2.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      // Slide 3 (Case C): Unsupported transition p:blinds with valid advTm -> omit transition, emit unsupported-mechanism, retain advanceAfter.
      "ppt/slides/slide3.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="case-c"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition spd="fast" advTm="2500"><p:blinds dir="horz"/></p:transition></p:sld>`,
      ),
      "ppt/slides/_rels/slide3.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([transitionCapability]);
    const result = importer(pptx);
    const slides = result.document.children.filter(
      (c): c is CommonNode => c instanceof CommonNode && c.kind === "slide",
    );

    // Slide 1 (Case A)
    expect(slides[0]?.props["transition"]).toEqual(
      transitions.morph({
        duration: 1500,
        by: "object",
        fallback: transitions.push({ from: "left" }),
      }),
    );
    assert.equal(slides[0]?.props["advanceAfter"], 2200);
    assert.equal(
      result.report.diagnostics.some((d) => d.location.slideIndex === 0),
      false,
    );

    // Slide 2 (Case B)
    expect(slides[1]?.props["transition"]).toEqual(
      transitions.push({ from: "right" }),
    );
    assert.equal(slides[1]?.props["advanceAfter"], 1800);
    const fallbackDiags = result.report.diagnostics.filter(
      (d) => d.location.slideIndex === 1,
    );
    assert.equal(fallbackDiags.length, 1);
    assert.deepEqual(fallbackDiags[0], {
      code: "compatibility-fallback",
      category: "unsupported",
      disposition: "fallback",
      location: {
        part: "ppt/slides/slide2.xml",
        xpath: "/p:sld/mc:AlternateContent",
        slideIndex: 1,
      },
      preferredMechanism: "p14:vortex",
      selectedFallback: "p:push",
      loss: {
        lost: "preferred-compatibility-branch",
        retainedAs: "transition:push",
      },
    });

    // Slide 3 (Case C)
    assert.equal(slides[2]?.props["transition"], undefined);
    assert.equal(slides[2]?.props["advanceAfter"], 2500);
    const omittedDiags = result.report.diagnostics.filter(
      (d) => d.location.slideIndex === 2,
    );
    assert.equal(omittedDiags.length, 1);
    assert.equal(omittedDiags[0]?.code, "unsupported-mechanism");
    assert.equal(omittedDiags[0]?.disposition, "omitted");
    if (omittedDiags[0]?.code === "unsupported-mechanism") {
      assert.equal(omittedDiags[0].localName, "blinds");
      assert.equal(omittedDiags[0].namespaceUri, P_NS);
    }

    // Re-emitting imported document produces valid PPTX preserving Case A's push fallback, Case B's adopted push, and Case C's advanceAfter.
    const resolve = createResolver(documentCapabilities);
    const rebuilt = emit(resolve(result.document));
    const validated = await validatePptx(rebuilt.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(validated.errors, []);
    assert.deepEqual(validated.slides[0]?.transition, {
      kind: "morph",
      speed: "slow",
      durationMs: 1500,
      by: "object",
      fallback: { kind: "push", speed: "slow", direction: "l" },
    });
    assert.deepEqual(validated.slides[1]?.transition, {
      kind: "push",
      speed: "slow",
      direction: "r",
    });
    assert.equal(validated.slides[2]?.transition, undefined);
    assert.equal(validated.slides[2]?.advanceAfterMs, 2500);
  });

  it("handles layout fallback, unsupported Choice+Fallback pairs, and leaves non-transition mc:AlternateContent unconsumed", () => {
    const p14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
      // Layout with unsupported p14:glitter Choice + supported p:fade Fallback
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${p14Ns}"><mc:Choice Requires="p14"><p:transition spd="med"><p14:glitter dir="l"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      // Slide with both unsupported AlternateContent (honeycomb + checker) AND an unrelated non-transition AlternateContent
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${p14Ns}"><mc:Choice Requires="p14"><p:transition spd="med"><p14:honeycomb/></p:transition></mc:Choice><mc:Fallback><p:transition spd="med"><p:checker dir="horz"/></p:transition></mc:Fallback></mc:AlternateContent><mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${p14Ns}"><mc:Choice Requires="p14"><p:extLst/></mc:Choice></mc:AlternateContent></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter([transitionCapability]);
    const result = importer(pptx);
    assert.ok(result.document instanceof DeckContainer);
    const layout = result.document.masters[0]!.layouts[0]!;
    expect(layout.props["transition"]).toEqual(transitions.fade());

    const layoutFallback = result.report.diagnostics.find(
      (d) =>
        d.code === "compatibility-fallback" &&
        d.location.part === "ppt/slideLayouts/slideLayout1.xml",
    );
    assert.ok(layoutFallback !== undefined);
    if (layoutFallback.code === "compatibility-fallback") {
      assert.equal(layoutFallback.preferredMechanism, "p14:glitter");
      assert.equal(layoutFallback.selectedFallback, "p:fade");
      assert.equal(layoutFallback.loss.retainedAs, "transition:fade");
    }

    const slideUnsupported = result.report.diagnostics.find(
      (d) =>
        d.code === "unsupported-mechanism" &&
        d.location.part === "ppt/slides/slide1.xml",
    );
    assert.ok(slideUnsupported !== undefined);

    const unrelatedUnconsumed = result.report.diagnostics.find(
      (d) =>
        d.code === "unconsumed-xml" &&
        d.location.part === "ppt/slides/slide1.xml" &&
        d.localName === "AlternateContent",
    );
    assert.ok(unrelatedUnconsumed !== undefined);
  });
});
