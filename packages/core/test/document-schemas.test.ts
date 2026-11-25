import assert from "node:assert/strict";

import {
  BuildError,
  createMorphScope,
  deckPropsSchema,
  sectionPropsSchema,
  slideBackgroundSchema,
  slidePropsSchema,
  type DeckProps,
  type SectionProps,
  type SlideBackground,
  type SlideProps,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import {
  CommonNode,
  createDeckNode,
  createSectionNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { describe, it } from "vitest";

describe("Slice 3: Deck/Slide/Section schemas, resolve validation, and declaration preservation", () => {
  it("exports DeckProps, SlideProps, SectionProps, and SlideBackground schemas", () => {
    const bg: SlideBackground = slideBackgroundSchema.parse({
      image: new URL("https://example.test/bg.png"),
    });
    assert(typeof bg === "object" && "image" in bg && bg.image instanceof URL);

    const deckProps: DeckProps = deckPropsSchema.parse({
      width: 960,
      height: 540,
      title: "Quarterly Review",
      style: { fontSize: "32pt", color: "#112233" },
    });
    assert.equal(deckProps.style?.fontSize, "32pt");

    const slideProps: SlideProps = slidePropsSchema.parse({
      id: "cover",
      hidden: false,
      background: bg,
      style: { fontSize: 32, display: "flex", flexDirection: "column" },
    });
    assert.equal(slideProps.id, "cover");

    const sectionProps: SectionProps = sectionPropsSchema.parse({
      name: "Introduction",
    });
    assert.equal(sectionProps.name, "Introduction");

    assert.equal(deckPropsSchema.safeParse({ extra: true }).success, false);
    assert.equal(deckPropsSchema.safeParse({ width: 10 }).success, false);
    assert.equal(sectionPropsSchema.safeParse({ name: "  " }).success, false);
    assert.equal(slidePropsSchema.safeParse({ id: "" }).success, false);
    assert.equal(
      slidePropsSchema.safeParse({ id: "s1", style: { zIndex: 1 } }).success,
      false,
    );
  });

  it("preserves resolve BuildError codes and paths for Deck, Section, Slide, and LinkTarget", () => {
    const resolve = createResolver([]);

    const expectBuildError = (
      fn: () => unknown,
      expectedCode: string,
      expectedPath: string,
    ) => {
      assert.throws(fn, (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(
          error.diagnostics[0]?.code,
          expectedCode,
          `Expected ${expectedCode}, got ${error.diagnostics[0]?.code}: ${error.diagnostics[0]?.message}`,
        );
        assert.equal(error.diagnostics[0]?.path, expectedPath);
        return true;
      });
    };

    const badDeckUnknown = createDeckNode({ unexpected: true } as never).append(
      createSlideNode({ id: "s1" }),
    );
    expectBuildError(() => resolve(badDeckUnknown), "unknown-prop", "deck");

    const badDeckDim = createDeckNode({ width: 10 }).append(
      createSlideNode({ id: "s1" }),
    );
    expectBuildError(() => resolve(badDeckDim), "invalid-prop", "deck");

    const badDeckTitle = createDeckNode({ title: "bad\u0000title" }).append(
      createSlideNode({ id: "s1" }),
    );
    expectBuildError(() => resolve(badDeckTitle), "invalid-prop", "deck");

    const badSectionProp = createDeckNode({}).append(
      createSectionNode({ name: "S1", extra: 1 } as never).append(
        createSlideNode({ id: "s1" }),
      ),
    );
    expectBuildError(
      () => resolve(badSectionProp),
      "unknown-prop",
      "deck/section[1]",
    );

    const badSectionName = createDeckNode({}).append(
      createSectionNode({ name: "   " }).append(createSlideNode({ id: "s1" })),
    );
    expectBuildError(
      () => resolve(badSectionName),
      "invalid-prop",
      "deck/section[1]",
    );

    const badSlideHidden = createDeckNode({}).append(
      createSlideNode({ id: "s1", hidden: "yes" as never }),
    );
    expectBuildError(
      () => resolve(badSlideHidden),
      "invalid-prop",
      "deck/slide[s1]",
    );

    const badSlideStyle = createDeckNode({}).append(
      createSlideNode({ id: "s1", style: { zIndex: 2 } as never }),
    );
    expectBuildError(
      () => resolve(badSlideStyle),
      "invalid-style",
      "deck/slide[s1]",
    );
  });

  it("preserves unit declarations, omitted vs explicit overrides, patch undefined removal, snapshotting, and identity", () => {
    // 1. "32pt" vs 32 vs "32px" preserved verbatim in CommonNode props
    const sPt = createSlideNode({ id: "s-pt", style: { fontSize: "32pt" } });
    const sNum = createSlideNode({ id: "s-num", style: { fontSize: 32 } });
    const sPx = createSlideNode({ id: "s-px", style: { fontSize: "32px" } });
    assert.equal(
      (sPt.props["style"] as { fontSize?: unknown }).fontSize,
      "32pt",
    );
    assert.equal((sNum.props["style"] as { fontSize?: unknown }).fontSize, 32);
    assert.equal(
      (sPx.props["style"] as { fontSize?: unknown }).fontSize,
      "32px",
    );

    // 2. Omitted default vs explicit equal-value override
    const omittedWidthDeck = createDeckNode({});
    const explicitWidthDeck = createDeckNode({ width: 1280 });
    assert.equal(Object.hasOwn(omittedWidthDeck.props, "width"), false);
    assert.equal(Object.hasOwn(explicitWidthDeck.props, "width"), true);
    assert.equal(explicitWidthDeck.props["width"], 1280);

    // 3. .style(patch) omission vs explicit undefined
    sPt.style({ color: "#AA0000" });
    assert.deepEqual(
      { ...(sPt.props["style"] as object) },
      {
        fontSize: "32pt",
        color: "#AA0000",
      },
    );
    sPt.style({ fontSize: undefined });
    assert.deepEqual(
      { ...(sPt.props["style"] as object) },
      {
        color: "#AA0000",
      },
    );
    assert.equal(
      Object.hasOwn(sPt.props["style"] as Record<string, unknown>, "fontSize"),
      false,
    );

    // 4. Call-time value snapshotting & instance identity preservation
    const morphScope = createMorphScope();
    const morphKey = morphScope.handle("hero");
    const url = new URL("https://example.test/asset.png");
    const handle = Object.assign(() => {}, { type: "motion-scope-handle" });
    const child = new CommonNode("box", "box", { id: "child" });
    const mutableStyle = { color: "#111111" };
    const node = new CommonNode("text", "text", {
      morphKey,
      background: { image: url },
      motionHandle: handle,
      targetNode: child,
      style: mutableStyle,
    });
    mutableStyle.color = "#999999";

    assert.equal((node.props["style"] as { color?: string }).color, "#111111");
    assert.equal(node.props["morphKey"], morphKey);
    assert.equal(node.props["motionHandle"], handle);
    assert.equal(node.props["targetNode"], child);
    assert(
      (node.props["background"] as { image?: unknown }).image instanceof URL,
    );
    assert.equal(
      (node.props["background"] as { image: URL }).image.href,
      "https://example.test/asset.png",
    );
  });
});
