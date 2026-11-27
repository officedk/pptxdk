import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { createCodegen, type CodegenResult } from "@pptxdk/codegen";
import type { NativeMechanismDefinition } from "@pptxdk/core/capability";
import {
  CommonNode,
  SlideLayoutContainer,
  SlideMasterContainer,
} from "@pptxdk/core/document";
import type { PreviewResult } from "@pptxdk/preview";

import { chartMechanismVerifications } from "../src/chart/test/mechanisms.ts";
import { groupMechanismVerifications } from "../src/group/test/mechanisms.ts";
import { imageMechanismVerifications } from "../src/image/test/mechanisms.ts";
import { footerPlaceholderMechanismVerifications } from "../src/masters/test/footer-mechanisms.ts";
import { masterOwnershipVerifications } from "../src/masters/test/master-layout-mechanisms.ts";
import { masterSharedCanvasMechanismVerifications } from "../src/masters/test/shared-canvas-mechanisms.ts";
import { themeMechanismVerifications } from "../src/masters/test/theme-mechanisms.ts";
import { motionMechanismVerifications } from "../src/motion/test/mechanisms.ts";
import { notesMechanismVerifications } from "../src/notes/test/mechanisms.ts";
import { placeholderMechanismVerifications } from "../src/placeholder/test/mechanisms.ts";
import { shapeMechanismVerifications } from "../src/shape/test/mechanisms.ts";
import { slideBackgroundMechanismVerifications } from "../src/slide-background/test/mechanisms.ts";
import { tableMechanismVerifications } from "../src/table/test/mechanisms.ts";
import { textMechanismVerifications } from "../src/text/test/mechanisms.ts";
import { transitionMechanismVerifications } from "../src/transition/test/mechanisms.ts";
import {
  runCapabilityScenario,
  type CapabilityScenario,
  type NativeRoundtripAssertContext,
  type ScenarioExecutionResult,
  type StageCounters,
} from "./acceptance-runner.ts";

export type ExtractMechanisms<Modules extends readonly unknown[]> =
  Modules[number] extends infer M
    ? M extends { readonly mechanisms?: readonly (infer D)[] | undefined }
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
    ...groupMechanismVerifications,
    ...textMechanismVerifications,
    ...imageMechanismVerifications,
    ...tableMechanismVerifications,
    ...chartMechanismVerifications,
    ...masterOwnershipVerifications,
    ...themeMechanismVerifications,
    ...masterSharedCanvasMechanismVerifications,
    ...footerPlaceholderMechanismVerifications,
    ...slideBackgroundMechanismVerifications,
    ...placeholderMechanismVerifications,
    ...notesMechanismVerifications,
    ...transitionMechanismVerifications,
    ...motionMechanismVerifications,
  };

const generateProductionSource = createCodegen(documentCapabilities);

function collectDocumentExpectations(doc: CommonNode): {
  readonly masterCount: number;
  readonly layoutCount: number;
  readonly slideCount: number;
  readonly sectionNames: readonly string[];
  readonly nodeIds: readonly string[];
  readonly declarationCounts: ReadonlyMap<string, number>;
} {
  let masterCount = 0;
  let layoutCount = 0;
  let slideCount = 0;
  const sectionNames: string[] = [];
  const nodeIds: string[] = [];
  const declarationCounts = new Map<string, number>();

  const visitTree = (node: CommonNode): void => {
    if (node.id !== undefined) {
      nodeIds.push(node.id);
    }
    if (
      node.type !== "deck" &&
      node.type !== "slide-master" &&
      node.type !== "slide-layout" &&
      node.type !== "section" &&
      node.type !== "slide"
    ) {
      declarationCounts.set(
        node.type,
        (declarationCounts.get(node.type) ?? 0) + 1,
      );
    }
    for (const child of node.children) {
      if (child instanceof CommonNode) {
        visitTree(child);
      }
    }
  };

  for (const child of doc.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child instanceof SlideMasterContainer) {
      masterCount += 1;
      layoutCount += child.layouts.length;
      if (child.id !== undefined) nodeIds.push(child.id);
      for (const mChild of child.children) {
        if (mChild instanceof SlideLayoutContainer) {
          if (mChild.id !== undefined) nodeIds.push(mChild.id);
          for (const lChild of mChild.children) {
            if (lChild instanceof CommonNode) visitTree(lChild);
          }
        } else if (mChild instanceof CommonNode) {
          visitTree(mChild);
        }
      }
      continue;
    }
    if (child.type === "section") {
      sectionNames.push(
        typeof child.props["name"] === "string" ? child.props["name"] : "",
      );
      for (const secChild of child.children) {
        if (secChild instanceof CommonNode && secChild.type === "slide") {
          slideCount += 1;
          visitTree(secChild);
        }
      }
      continue;
    }
    if (child.type === "slide") {
      slideCount += 1;
      visitTree(child);
    }
  }

  return {
    masterCount,
    layoutCount,
    slideCount,
    sectionNames,
    nodeIds,
    declarationCounts,
  };
}

const vanillaCalleeByKind: Readonly<Record<string, RegExp>> = {
  box: /\bbox\(/gu,
  cell: /\bcell\(/gu,
  chart: /\bchart\(/gu,
  group: /\bgroup\(/gu,
  image: /\bimage\(/gu,
  "motion-scope": /\bmotionScope\(/gu,
  "motion-step": /\bstep\(/gu,
  notes: /\bnotes\(/gu,
  paragraph: /\bparagraph\(/gu,
  placeholder: /\bplaceholder\(/gu,
  row: /\brow\(/gu,
  shape: /\bshape\(/gu,
  "slide-number": /\bslideNumber\(/gu,
  span: /\bspan\(/gu,
  table: /\btable\(/gu,
  text: /\btext\(/gu,
};

const reactTagByKind: Readonly<Record<string, RegExp>> = {
  box: /<Box[\s>]/gu,
  cell: /<Cell[\s>]/gu,
  chart: /<Chart[\s>]/gu,
  group: /<Group[\s>]/gu,
  image: /<Image[\s>]/gu,
  "motion-scope": /<MotionScope[\s>]/gu,
  "motion-step": /<Step[\s>]/gu,
  notes: /<Notes[\s>]/gu,
  paragraph: /<P[\s>]/gu,
  placeholder: /<Placeholder[\s>]/gu,
  row: /<Row[\s>]/gu,
  shape: /<Shape[\s/>]/gu,
  "slide-number": /<SlideNumber[\s/>]/gu,
  span: /<Span[\s>]/gu,
  table: /<Table[\s>]/gu,
  text: /<Text[\s>]/gu,
};

function assertGeneratedSourceNonEmptyAndComplete(
  scenarioId: string,
  originLabel: string,
  doc: CommonNode,
  result: CodegenResult,
): void {
  const source = result.files[result.entry];
  assert.equal(
    typeof source,
    "string",
    `${scenarioId} (${originLabel}:${result.target}) missing string entry ${result.entry}`,
  );
  const sourceText = source as string;
  assert.ok(
    sourceText.trim().length > 0,
    `${scenarioId} (${originLabel}:${result.target}) generated empty source`,
  );

  const summary = collectDocumentExpectations(doc);
  const countMatches = (pattern: RegExp): number =>
    sourceText.match(pattern)?.length ?? 0;

  if (result.target === "vanilla") {
    assert.equal(
      countMatches(/\bslideMaster\(/gu),
      summary.masterCount,
      `${scenarioId} (${originLabel}:vanilla) master count mismatch`,
    );
    assert.equal(
      countMatches(/\bdefineSlideLayout\(/gu),
      summary.layoutCount,
      `${scenarioId} (${originLabel}:vanilla) layout count mismatch`,
    );
    assert.equal(
      countMatches(/\bslide\(/gu),
      summary.slideCount,
      `${scenarioId} (${originLabel}:vanilla) slide count mismatch`,
    );
    assert.equal(
      countMatches(/\bsection\(/gu),
      summary.sectionNames.length,
      `${scenarioId} (${originLabel}:vanilla) section count mismatch`,
    );
  } else {
    assert.equal(
      countMatches(/<SlideMaster[\s>]/gu),
      summary.masterCount,
      `${scenarioId} (${originLabel}:react) SlideMaster count mismatch`,
    );
    assert.equal(
      countMatches(/\bdefineSlideLayout\(/gu),
      summary.layoutCount,
      `${scenarioId} (${originLabel}:react) layout count mismatch`,
    );
    assert.equal(
      countMatches(/<Slide[\s>]/gu),
      summary.slideCount,
      `${scenarioId} (${originLabel}:react) Slide count mismatch`,
    );
    assert.equal(
      countMatches(/<Section[\s>]/gu),
      summary.sectionNames.length,
      `${scenarioId} (${originLabel}:react) Section count mismatch`,
    );
  }

  const kindPatterns =
    result.target === "vanilla" ? vanillaCalleeByKind : reactTagByKind;
  for (const [kind, count] of summary.declarationCounts) {
    const pattern = kindPatterns[kind];
    assert.ok(
      pattern !== undefined,
      `${scenarioId} (${originLabel}:${result.target}) missing matcher for declaration ${kind}`,
    );
    assert.ok(
      countMatches(pattern) >= count,
      `${scenarioId} (${originLabel}:${result.target}) omitted ${kind} declarations: expected >= ${count}, got ${countMatches(pattern)}`,
    );
  }

  for (const id of summary.nodeIds) {
    assert.ok(
      sourceText.includes(JSON.stringify(id)),
      `${scenarioId} (${originLabel}:${result.target}) omitted id ${id}`,
    );
  }

  for (const name of summary.sectionNames) {
    assert.ok(
      sourceText.includes(JSON.stringify(name)),
      `${scenarioId} (${originLabel}:${result.target}) omitted section ${name}`,
    );
  }

  for (const [relPath, content] of Object.entries(result.files)) {
    if (relPath === result.entry) continue;
    assert.ok(
      content instanceof Uint8Array && content.byteLength > 0,
      `${scenarioId} (${originLabel}:${result.target}) asset ${relPath} must be non-empty Uint8Array`,
    );
    assert.ok(
      sourceText.includes(`./${relPath}`),
      `${scenarioId} (${originLabel}:${result.target}) asset ${relPath} is unreferenced in source`,
    );
  }
}

/**
 * Runs every mechanism verification entry, executing each unique scenario once
 * and sharing its read-only initial and pure-file roundtrip artifacts across assertions.
 */
export async function runProductionMechanismVerifications(
  options: { readonly counters?: StageCounters } = {},
): Promise<{ readonly codegenScenarioCount: number }> {
  const entries = Object.entries(productionMechanismVerifications) as [
    ProductionMechanismKey,
    MechanismVerificationEntry<ProductionMechanism>,
  ][];
  const cachedExecutions = new Map<
    CapabilityScenario,
    Promise<ScenarioExecutionResult>
  >();
  const checkedCodegenScenarios = new Set<CapabilityScenario>();

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

    if (
      entry.scenario.kind !== "rejection" &&
      !checkedCodegenScenarios.has(entry.scenario)
    ) {
      checkedCodegenScenarios.add(entry.scenario);
      for (const [originLabel, doc] of [
        ["direct", execution.document.document],
        ["roundtrip-imported", execution.roundtrip.importedDocument],
      ] as const) {
        for (const target of ["vanilla", "react"] as const) {
          const generated = generateProductionSource(doc, { target });
          assertGeneratedSourceNonEmptyAndComplete(
            entry.scenario.id,
            originLabel,
            doc,
            generated,
          );
        }
      }
    }
  }

  return { codegenScenarioCount: checkedCodegenScenarios.size };
}
