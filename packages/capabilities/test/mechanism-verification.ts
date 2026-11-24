import { documentCapabilities } from "@pptxdk/capabilities/all";
import type { NativeMechanismDefinition } from "@pptxdk/core/capability";
import type { PreviewResult } from "@pptxdk/preview";

import { shapeMechanismVerifications } from "../src/shape/test/mechanisms.ts";
import {
  runCapabilityScenario,
  type CapabilityScenario,
  type NativeRoundtripAssertContext,
  type ScenarioExecutionResult,
  type StageCounters,
} from "./acceptance-runner.ts";
import { contentMechanismVerifications } from "./mechanism-registry-content.ts";
import { presentationMechanismVerifications } from "./mechanism-registry-presentation.ts";

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
    ...shapeMechanismVerifications,
    ...contentMechanismVerifications,
    ...presentationMechanismVerifications,
  };

/**
 * Runs every mechanism verification entry, executing each unique scenario once
 * and sharing its read-only initial and pure-file roundtrip artifacts across assertions.
 */
export async function runProductionMechanismVerifications(
  options: { readonly counters?: StageCounters } = {},
): Promise<void> {
  const entries = Object.entries(productionMechanismVerifications) as [
    ProductionMechanismKey,
    MechanismVerificationEntry<ProductionMechanism>,
  ][];
  const cachedExecutions = new Map<
    CapabilityScenario,
    Promise<ScenarioExecutionResult>
  >();

  for (const [, entry] of entries) {
    let executionPromise = cachedExecutions.get(entry.scenario);
    if (executionPromise === undefined) {
      executionPromise = runCapabilityScenario(entry.scenario, options);
      cachedExecutions.set(entry.scenario, executionPromise);
    }

    const execution = await executionPromise;
    await entry.assertNative({
      stage: "initial",
      pptx: execution.primary.pptx,
      validation: execution.primary.validation,
      report: execution.primary.report,
      openEditableDocument: () => execution.openEditableRoundtripDocument(),
    });
    await entry.assertNative({
      stage: "roundtrip",
      pptx: execution.roundtrip.pptx,
      validation: execution.roundtrip.validation,
      report: execution.roundtrip.report,
      importedDocument: execution.roundtrip.importedDocument,
      openEditableDocument: () => execution.openEditableRoundtripDocument(),
    });
    if (entry.assertPaint !== undefined) {
      await entry.assertPaint({
        initial: execution.primary.preview,
        roundtrip: execution.roundtrip.preview,
      });
    }
  }
}
