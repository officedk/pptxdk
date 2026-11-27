import assert from "node:assert/strict";

import {
  BuildError,
  CapabilityError,
  authorIdSchema,
  layoutLengthSchema,
  layoutSpacingSchema,
  layoutStyleSchema,
  linkTargetSchema,
  nonNegativeLayoutSpacingSchema,
  paddingStyleSchema,
  placeholderSizeSchema,
  placeholderStyleSchema,
  type AuthorId,
  type LayoutLength,
  type LayoutSpacing,
  type LayoutStyle,
  type LinkTarget,
  type NonNegativeLayoutSpacing,
  type PaddingStyle,
  type PlaceholderSize,
} from "@pptxdk/core";
import {
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  parseLinkTarget,
  readAuthorId,
} from "@pptxdk/core/capability";
import { describe, it } from "vitest";

describe("Slice 1: core author primitives & layout schemas", () => {
  it("validates AuthorId via authorIdSchema and readAuthorId with localized BuildError", () => {
    const validId: AuthorId = authorIdSchema.parse("hero-title");
    assert.equal(validId, "hero-title");
    assert.equal(readAuthorId(undefined, "deck/slide[1]"), undefined);
    assert.equal(readAuthorId("cover", "deck/slide[1]"), "cover");

    for (const invalid of ["", "   ", 123, null, {}]) {
      assert.equal(authorIdSchema.safeParse(invalid).success, false);
      assert.throws(
        () => readAuthorId(invalid, "deck/slide[1]/text[1]"),
        (error: unknown) => {
          assert(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-prop");
          assert.equal(error.diagnostics[0]?.path, "deck/slide[1]/text[1]");
          return true;
        },
      );
    }
  });

  it("validates LayoutLength and LayoutSpacing including percentages, tuples, and auto", () => {
    const validLengths: readonly LayoutLength[] = [
      0,
      120,
      -20,
      "50%",
      ".5%",
      "-12.5%",
      "auto",
    ];
    for (const value of validLengths) {
      assert.equal(layoutLengthSchema.parse(value), value);
    }

    for (const invalid of [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      "20px",
      "20pt",
      "",
      "%",
      null,
    ]) {
      assert.equal(layoutLengthSchema.safeParse(invalid).success, false);
    }

    const validSpacings: readonly LayoutSpacing[] = [
      12,
      [10, 20],
      [1, 2, 3, 4],
      [-4, 8],
    ];
    for (const value of validSpacings) {
      assert.deepEqual(layoutSpacingSchema.parse(value), value);
    }

    for (const invalid of [
      [10],
      [1, 2, 3],
      [1, 2, 3, 4, 5],
      ["10", 20],
      Number.NaN,
    ]) {
      assert.equal(layoutSpacingSchema.safeParse(invalid).success, false);
    }
  });

  it("validates PaddingStyle and Block/Flex/Grid LayoutStyle via Zod schemas", () => {
    const pad: PaddingStyle = paddingStyleSchema.parse({
      padding: [8, 16],
      paddingTop: 12,
    });
    assert.deepEqual(pad, { padding: [8, 16], paddingTop: 12 });
    assert.equal(paddingStyleSchema.safeParse({ padding: -1 }).success, false);
    assert.equal(
      paddingStyleSchema.safeParse({ paddingLeft: -2 }).success,
      false,
    );
    assert.equal(paddingStyleSchema.safeParse({ margin: 10 }).success, false);

    const flexLayout: LayoutStyle = layoutStyleSchema.parse({
      display: "flex",
      flexDirection: "column",
      justifyContent: "space-between",
      alignItems: "center",
      gap: 12,
      width: "100%",
      height: "auto",
      padding: [8, 12, 16, 20],
      zIndex: 2,
    });
    assert.equal(flexLayout.display, "flex");

    const gridLayout: LayoutStyle = layoutStyleSchema.parse({
      display: "grid",
      gridTemplateColumns: "120px 1fr auto",
      gridTemplateRows: "auto",
      gridAutoRows: "40px",
      gridColumn: "1 / span 2",
      gridRow: 2,
      justifyItems: "stretch",
      justifySelf: "center",
    });
    assert.equal(gridLayout.gridColumn, "1 / span 2");

    assert.equal(
      layoutStyleSchema.safeParse({ display: "inline" }).success,
      false,
    );
    assert.equal(
      layoutStyleSchema.safeParse({ aspectRatio: 0 }).success,
      false,
    );
    assert.equal(
      layoutStyleSchema.safeParse({ flexShrink: -1 }).success,
      false,
    );
    assert.equal(
      layoutStyleSchema.safeParse({ unknownLayoutProp: 1 }).success,
      false,
    );

    for (const size of ["full", "half", "quarter"] as const) {
      const parsed: PlaceholderSize = placeholderSizeSchema.parse(size);
      assert.equal(parsed, size);
    }
    assert.equal(placeholderSizeSchema.safeParse("third").success, false);

    const phStyle = placeholderStyleSchema.parse({
      position: "absolute",
      left: 60,
      top: 100,
      width: 480,
      height: 320,
      background: "#FFFFFF",
      border: { color: "#112233", width: 1, style: "solid" },
      boxShadow: "none",
      flipH: true,
      flipV: false,
      fontSize: 18,
    });
    assert.equal(phStyle.left, 60);
    assert.equal(phStyle.flipH, true);
    assert.equal(phStyle.boxShadow, "none");
  });

  it("validates LinkTarget via linkTargetSchema and parseLinkTarget", () => {
    const idTarget: LinkTarget = linkTargetSchema.parse({ id: "details" });
    const urlTarget: LinkTarget = linkTargetSchema.parse({
      url: "https://example.test",
    });
    const homeTarget: LinkTarget = linkTargetSchema.parse({ action: "home" });
    assert.deepEqual(idTarget, { id: "details" });
    assert.deepEqual(urlTarget, { url: "https://example.test" });
    assert.deepEqual(homeTarget, { action: "home" });

    assert.deepEqual(parseLinkTarget({ id: "details", url: undefined }), {
      id: "details",
    });
    assert.equal(parseLinkTarget(undefined), undefined);

    for (const invalid of [
      "not-an-object",
      null,
      {},
      { id: undefined },
      { id: "a", url: "https://x.test" },
      { action: "next" },
      { unknown: "value" },
    ]) {
      assert.throws(
        () => parseLinkTarget(invalid, "box[cta]"),
        (error: unknown) => {
          assert(error instanceof CapabilityError);
          assert.equal(error.code, "invalid-link");
          assert.equal(error.path, "box[cta]");
          return true;
        },
      );
    }
  });

  it("exports readonly nonNegativeLayoutSpacingSchema and derives explicitPixelSize, fixedAxes, positioned, and auto preservesIntrinsicSize in parseCapabilityLayoutContract", () => {
    const pad2 = [9, 12] as const;
    const pad4 = [9, 12, 9, 12] as const;
    const parsed2: NonNegativeLayoutSpacing =
      nonNegativeLayoutSpacingSchema.parse(pad2);
    const parsed4: NonNegativeLayoutSpacing =
      nonNegativeLayoutSpacingSchema.parse(pad4);
    assert.deepEqual(parsed2, [9, 12]);
    assert.deepEqual(parsed4, [9, 12, 9, 12]);
    assert.equal(nonNegativeLayoutSpacingSchema.safeParse(-1).success, false);
    assert.equal(
      nonNegativeLayoutSpacingSchema.safeParse([4, -1]).success,
      false,
    );

    const allowed = new Set<string>(layoutStyleKeys);

    const percentContract = parseCapabilityLayoutContract({
      nativeKind: "graphicFrame",
      props: { style: { width: "100%", height: "auto" } },
      allowedStyleKeys: allowed,
    });
    assert.deepEqual(percentContract.explicitPixelSize, {});
    assert.deepEqual(percentContract.fixedAxes, { width: true, height: false });
    assert.equal(percentContract.positioned, false);

    const pixelContract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: { style: { position: "relative", width: 240, height: 120 } },
      allowedStyleKeys: allowed,
    });
    assert.deepEqual(pixelContract.explicitPixelSize, {
      width: 240,
      height: 120,
    });
    assert.deepEqual(pixelContract.fixedAxes, { width: true, height: true });
    assert.equal(pixelContract.positioned, true);

    const insetContract = parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: { style: { position: "absolute", inset: 24 } },
      allowedStyleKeys: allowed,
    });
    assert.deepEqual(insetContract.explicitPixelSize, {});
    assert.deepEqual(insetContract.fixedAxes, { width: true, height: true });
    assert.equal(insetContract.positioned, true);

    const autoWidthGroup = parseCapabilityLayoutContract({
      nativeKind: "grpSp",
      props: { style: { width: "auto" } },
      allowedStyleKeys: allowed,
    });
    assert.equal(autoWidthGroup.preservesIntrinsicSize, true);
    assert.deepEqual(autoWidthGroup.fixedAxes, { width: false, height: false });
  });
});
