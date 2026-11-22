import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
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
import type { NativeMechanismDefinition } from "@pptxdk/core/capability";
import type { PreviewResult } from "@pptxdk/preview";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { createElement } from "react";

import {
  runCapabilityScenario,
  runScenarioRoundtrip,
  type CapabilityScenario,
  type NativeRoundtripAssertContext,
} from "./acceptance-runner.ts";
import { chartMechanismVerifications } from "./chart/mechanisms.ts";
import { imageMechanismVerifications } from "./image/mechanisms.ts";
import { notesMechanismVerifications } from "./notes/mechanisms.ts";
import { shapeAndGroupMechanismVerifications } from "./shape/mechanisms.ts";
import {
  textBoxRoundtripScenario,
  textMechanismVerifications,
} from "./text/mechanisms.ts";

export { textBoxRoundtripScenario };

export type ExtractMechanisms<Modules extends readonly unknown[]> =
  Modules[number] extends infer M
    ? M extends { readonly mechanisms: readonly (infer D)[] }
      ? Extract<D, NativeMechanismDefinition>
      : never
    : never;

export type ProductionMechanism = ExtractMechanisms<
  typeof documentCapabilities
>;
export type ProductionMechanismKey = ProductionMechanism["key"];

export type MechanismVerificationEntry<M extends NativeMechanismDefinition> =
  M extends { readonly role: "canvas" }
    ? {
        readonly scenario: CapabilityScenario;
        assertNative(ctx: NativeRoundtripAssertContext): void | Promise<void>;
        assertPaint(ctx: {
          readonly initial: PreviewResult;
          readonly roundtrip: PreviewResult;
        }): void | Promise<void>;
      }
    : M extends { readonly role: "show" }
      ? {
          readonly scenario: CapabilityScenario;
          assertNative(ctx: NativeRoundtripAssertContext): void | Promise<void>;
          readonly assertPaint?: never;
        }
      : never;

export type ExhaustiveMechanismVerificationMap<
  M extends NativeMechanismDefinition = ProductionMechanism,
> = {
  readonly [K in M["key"]]: MechanismVerificationEntry<
    Extract<M, { readonly key: K }>
  >;
};

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
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
            x: 32,
            y: 28,
            width: 320,
            height: 48,
            style: {
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
          x: 32,
          y: 28,
          width: 320,
          height: 48,
          style: {
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

export const productionMechanismVerifications: ExhaustiveMechanismVerificationMap =
  {
    "slide-master": {
      scenario: masterLayoutOwnershipScenario,
      assertNative: assertMasterLayoutNative,
      assertPaint({ initial, roundtrip }) {
        assert.deepEqual(initial.pages[0]?.diagnostics, []);
        assert.ok(initial.pages[0]?.png instanceof Uint8Array);
        assert.equal(roundtrip.pages.length, 1);
      },
    },
    "slide-layout": {
      scenario: masterLayoutOwnershipScenario,
      assertNative: assertMasterLayoutNative,
      assertPaint({ initial, roundtrip }) {
        assert.deepEqual(initial.pages[0]?.diagnostics, []);
        assert.ok(initial.pages[0]?.png instanceof Uint8Array);
        assert.equal(roundtrip.pages.length, 1);
      },
    },
    ...imageMechanismVerifications,
    ...notesMechanismVerifications,
    ...shapeAndGroupMechanismVerifications,
    ...textMechanismVerifications,
    ...chartMechanismVerifications,
  };

/** Runs every mechanism verification entry, deduplicating shared scenarios across initial and roundtrip artifacts. */
export async function runProductionMechanismVerifications(): Promise<void> {
  const entries = Object.entries(productionMechanismVerifications) as [
    ProductionMechanismKey,
    MechanismVerificationEntry<ProductionMechanism>,
  ][];
  const cachedRuns = new Map<
    CapabilityScenario,
    Promise<{
      readonly initialPptx: Uint8Array;
      readonly initialValidation: Awaited<
        ReturnType<typeof runCapabilityScenario>
      >[number]["validation"];
      readonly initialReport: Awaited<
        ReturnType<typeof runCapabilityScenario>
      >[number]["report"];
      readonly initialPreview: PreviewResult;
      readonly roundtrip: Awaited<ReturnType<typeof runScenarioRoundtrip>>;
    }>
  >();

  for (const [key, entry] of entries) {
    let runPromise = cachedRuns.get(entry.scenario);
    if (runPromise === undefined) {
      runPromise = (async () => {
        const runs = await runCapabilityScenario(entry.scenario);
        const primary = runs[0]!;
        const roundtrip = await runScenarioRoundtrip(primary.pptx, {
          ...(entry.scenario.pages === undefined
            ? {}
            : { pages: entry.scenario.pages }),
          ...(entry.scenario.emitOptions === undefined
            ? {}
            : { emitOptions: entry.scenario.emitOptions }),
        });
        assert.deepEqual(
          roundtrip.importReport.diagnostics,
          [],
          `${entry.scenario.id} roundtrip import must have zero unexpected diagnostics`,
        );
        return {
          initialPptx: primary.pptx,
          initialValidation: primary.validation,
          initialReport: primary.report,
          initialPreview: primary.preview,
          roundtrip,
        };
      })();
      cachedRuns.set(entry.scenario, runPromise);
    }

    const data = await runPromise;
    await entry.assertNative({
      stage: "initial",
      pptx: data.initialPptx,
      validation: data.initialValidation,
      report: data.initialReport,
    });
    await entry.assertNative({
      stage: "roundtrip",
      pptx: data.roundtrip.pptx,
      validation: data.roundtrip.validation,
      report: data.roundtrip.report,
    });
    if (entry.assertPaint !== undefined) {
      await entry.assertPaint({
        initial: data.initialPreview,
        roundtrip: data.roundtrip.preview,
      });
    }
    void key;
  }
}
