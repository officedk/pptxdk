import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, font, slide, text } from "@pptxdk/capabilities/document";
import { Text } from "@pptxdk/capabilities/text-elements";
import type { NativeMechanismDefinition } from "@pptxdk/core/capability";
import type { PreviewResult } from "@pptxdk/preview";
import { Deck, Slide } from "@pptxdk/react";
import { createElement } from "react";

import {
  runCapabilityScenario,
  runScenarioRoundtrip,
  type CapabilityScenario,
  type NativeRoundtripAssertContext,
} from "./acceptance-runner.ts";
import { imageMechanismVerifications } from "./image/mechanisms.ts";

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

export const textBoxRoundtripScenario: CapabilityScenario = {
  id: "mechanism-text-box-roundtrip",
  capability: "text",
  description:
    "Minimal single-slide Transitional text-box scenario proving dual-author equivalence and import/edit/re-emit roundtrip.",
  visualCheckpoints: [
    "Single-slide white canvas with left-aligned headline text box and centered body text box.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain editable <p:sp txBox='1'> objects with matching geometry and text.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "text-roundtrip" },
        createElement(
          Text,
          {
            id: "headline",
            name: "Headline",
            x: 32,
            y: 28,
            width: 320,
            height: 48,
            style: {
              fontFamily: jostFont,
              fontSize: "24pt",
              color: "#111827",
              align: "left",
              verticalAlign: "top",
            },
          },
          "Imported Headline",
        ),
        createElement(
          Text,
          {
            id: "caption",
            name: "Caption",
            x: 32,
            y: 96,
            width: 320,
            height: 64,
            style: {
              fontFamily: jostFont,
              fontSize: "14pt",
              color: "#334155",
              align: "center",
              verticalAlign: "middle",
              padding: [6, 12, 6, 12],
            },
          },
          "Roundtrip Body Text",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "text-roundtrip" }).append(
        text("Imported Headline", {
          id: "headline",
          name: "Headline",
          x: 32,
          y: 28,
          width: 320,
          height: 48,
          style: {
            fontFamily: jostFont,
            fontSize: "24pt",
            color: "#111827",
            align: "left",
            verticalAlign: "top",
          },
        }),
        text("Roundtrip Body Text", {
          id: "caption",
          name: "Caption",
          x: 32,
          y: 96,
          width: 320,
          height: 64,
          style: {
            fontFamily: jostFont,
            fontSize: "14pt",
            color: "#334155",
            align: "center",
            verticalAlign: "middle",
            padding: [6, 12, 6, 12],
          },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 1);
    assert.equal(validation.slides[0]?.objects.length, 2);
  },
};

export const productionMechanismVerifications: ExhaustiveMechanismVerificationMap =
  {
    ...imageMechanismVerifications,
    "text-box": {
      scenario: textBoxRoundtripScenario,
      assertNative({ stage, validation }) {
        assert.equal(validation.errors.length, 0, `${stage} validation errors`);
        assert.equal(validation.presentation.widthEmu, 640 * 9525);
        assert.equal(validation.presentation.heightEmu, 360 * 9525);
        assert.equal(validation.slides.length, 1);
        const s = validation.slides[0]!;
        assert.equal(s.id, "text-roundtrip");
        assert.equal(s.objects.length, 2);

        const headline = s.objects[0]!;
        assert.equal(headline.name, "Headline");
        assert.deepEqual(headline.boxEmu, {
          x: 32 * 9525,
          y: 28 * 9525,
          width: 320 * 9525,
          height: 48 * 9525,
        });
        assert.equal(headline.text?.text, "Imported Headline");
        assert.equal(headline.text?.paragraphs[0]?.runs[0]?.fontSizePt, 24);
        assert.equal(headline.text?.paragraphs[0]?.runs[0]?.color, "#111827");

        const caption = s.objects[1]!;
        assert.equal(caption.name, "Caption");
        assert.deepEqual(caption.boxEmu, {
          x: 32 * 9525,
          y: 96 * 9525,
          width: 320 * 9525,
          height: 64 * 9525,
        });
        assert.equal(caption.text?.text, "Roundtrip Body Text");
        assert.equal(caption.text?.paragraphs[0]?.alignment, "ctr");
        assert.equal(caption.text?.paragraphs[0]?.runs[0]?.fontSizePt, 14);
      },
      assertPaint({ initial, roundtrip }) {
        assert.deepEqual(initial.pages[0]?.diagnostics, []);
        assert.ok(initial.pages[0]?.png instanceof Uint8Array);
        // Until Ticket 02 imports embedded fonts, unembedded imported text reports missing outline font on paint.
        assert.equal(roundtrip.pages.length, 1);
      },
    },
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
        const roundtrip = await runScenarioRoundtrip(
          primary.pptx,
          entry.scenario.pages === undefined
            ? {}
            : { pages: entry.scenario.pages },
        );
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
