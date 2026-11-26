import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  defineSlideLayout,
  font,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@pptxdk/capabilities/document";
import { Text } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { createElement } from "react";

import type {
  CapabilityScenario,
  NativeRoundtripAssertContext,
} from "../../../test/acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../../../test/mechanism-verification.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const primaryCoverLayout = defineSlideLayout({
  id: "cover",
  name: "Cover Layout",
  type: "title",
  preserve: true,
});

const primaryIdleLayout = defineSlideLayout({
  id: "custom-idle",
  name: "Custom Idle",
  matchingName: "Custom Idle",
  userDrawn: true,
  preserve: true,
  showMasterShapes: false,
  showMasterPlaceholderAnimations: false,
});

const secondaryIdleLayout = defineSlideLayout({
  id: "secondary-blank",
  name: "Secondary Blank",
  type: "blank",
  preserve: true,
});

/**
 * Dual-author scenario verifying multiple SlideMasters, used and idle SlideLayouts,
 * preserve flags, and slide-to-layout/master relationship binding across roundtrip.
 */
export const masterLayoutOwnershipScenario: CapabilityScenario = {
  id: "mechanism-master-layout-ownership",
  capability: "masters",
  description:
    "Multi-master deck with explicit SlideMasters, used and idle SlideLayouts, preserve semantics, and slide-to-layout relationship binding.",
  visualCheckpoints: [
    "Slide 1 renders cleanly against explicit multi-master layout ownership.",
  ],
  nativeCheckpoints: [
    "Both masters, all 3 layouts (including 2 idle), preserve settings, and slide-to-layout relationship survive initial emit and import roundtrip.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "brand-a", name: "Brand A Master", preserve: true },
        createElement(SlideLayout, { layout: primaryCoverLayout }),
        createElement(SlideLayout, { layout: primaryIdleLayout }),
      ),
      createElement(
        SlideMaster,
        { id: "brand-b", name: "Brand B Master" },
        createElement(SlideLayout, { layout: secondaryIdleLayout }),
      ),
      createElement(
        Slide,
        { id: "cover-slide", layout: primaryCoverLayout },
        createElement(
          Text,
          {
            id: "heading",
            name: "Heading",
            style: {
              position: "absolute",
              left: 32,
              top: 28,
              width: 320,
              height: 48,
              fontFamily: jostFont,
              fontSize: "24pt",
              color: "#0F172A",
            },
          },
          "Master Owned Slide",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "brand-a",
        name: "Brand A Master",
        preserve: true,
      }).append(
        slideLayout(primaryCoverLayout),
        slideLayout(primaryIdleLayout),
      ),
      slideMaster({ id: "brand-b", name: "Brand B Master" }).append(
        slideLayout(secondaryIdleLayout),
      ),
      slide({ id: "cover-slide", layout: primaryCoverLayout }).append(
        text("Master Owned Slide", {
          id: "heading",
          name: "Heading",
          style: {
            position: "absolute",
            left: 32,
            top: 28,
            width: 320,
            height: 48,
            fontFamily: jostFont,
            fontSize: "24pt",
            color: "#0F172A",
          },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.masters?.length, 2);
    assert.equal(validation.masters?.[0]?.layouts.length, 2);
    assert.equal(validation.masters?.[1]?.layouts.length, 1);
    assert.equal(validation.slides.length, 1);
    assert.equal(
      validation.slides[0]?.layoutPart,
      validation.masters?.[0]?.layouts[0]?.part,
    );
  },
};

const assertMasterLayoutNative = ({
  stage,
  validation,
}: NativeRoundtripAssertContext): void => {
  assert.equal(validation.errors.length, 0, `${stage} validation errors`);
  assert.equal(validation.masters?.length, 2, `${stage} master count`);
  const [m1, m2] = validation.masters;
  assert.equal(m1?.name, "Brand A Master");
  assert.equal(m1?.preserve, true);
  assert.equal(m1?.layouts.length, 2);
  assert.equal(m1?.layouts[0]?.name, "Cover Layout");
  assert.equal(m1?.layouts[0]?.type, "title");
  assert.equal(m1?.layouts[0]?.preserve, true);
  assert.equal(m1?.layouts[1]?.name, "Custom Idle");
  assert.equal(m1?.layouts[1]?.matchingName, "Custom Idle");
  assert.equal(m1?.layouts[1]?.userDrawn, true);
  assert.equal(m1?.layouts[1]?.preserve, true);
  assert.equal(m1?.layouts[1]?.showMasterShapes, false);
  assert.equal(m1?.layouts[1]?.showMasterPlaceholderAnimations, false);

  assert.equal(m2?.name, "Brand B Master");
  assert.equal(m2?.preserve, undefined);
  assert.equal(m2?.layouts.length, 1);
  assert.equal(m2?.layouts[0]?.name, "Secondary Blank");
  assert.equal(m2?.layouts[0]?.type, "blank");
  assert.equal(m2?.layouts[0]?.preserve, true);

  const allIds = [
    m1.id,
    m2.id,
    ...m1.layouts.map((l) => l.id),
    ...m2.layouts.map((l) => l.id),
  ];
  assert.equal(new Set(allIds).size, 5);
  for (const id of allIds) {
    assert.ok(id >= 2_147_483_648);
  }
  assert.equal(validation.slides.length, 1);
  assert.equal(validation.slides[0]?.layoutPart, m1?.layouts[0]?.part);
  assert.equal(validation.slides[0]?.masterPart, m1?.part);
};

export const masterOwnershipVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  "slide-master" | "slide-layout"
> = {
  "slide-master": {
    scenario: masterLayoutOwnershipScenario,
    assertNative: assertMasterLayoutNative,
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.errors, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
      // Pure-file roundtrip uses unembedded Jost*; static preview reports preview-missing-font.
      assert.ok(
        roundtrip.pages[0]?.errors.some(
          (err) => err.code === "preview-missing-font",
        ),
        "Pure-file roundtrip without embedded font must report preview-missing-font",
      );
    },
  },
  "slide-layout": {
    scenario: masterLayoutOwnershipScenario,
    assertNative: assertMasterLayoutNative,
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.errors, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
      assert.ok(
        roundtrip.pages[0]?.errors.some(
          (err) => err.code === "preview-missing-font",
        ),
        "Pure-file roundtrip without embedded font must report preview-missing-font",
      );
    },
  },
};
