import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import type { NativeMechanismDefinition } from "@pptxdk/core/capability";
import type { PreviewResult } from "@pptxdk/preview";

import {
  runCapabilityScenario,
  runScenarioRoundtrip,
  type CapabilityScenario,
  type NativeRoundtripAssertContext,
} from "./acceptance-runner.ts";
import { imageMechanismVerifications } from "./image/mechanisms.ts";
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

export const productionMechanismVerifications: ExhaustiveMechanismVerificationMap =
  {
    ...imageMechanismVerifications,
    ...shapeAndGroupMechanismVerifications,
    ...textMechanismVerifications,
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
