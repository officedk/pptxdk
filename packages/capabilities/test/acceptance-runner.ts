import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import type { BuildReport, CommonNode } from "@pptxdk/core";
import { createResolver, type Resolution } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import {
  paint,
  type PreviewPageResult,
  type PreviewResult,
} from "@pptxdk/preview";
import { expandDocument } from "@pptxdk/react";
import { validatePptx, type ValidationResult } from "@pptxdk/validation";
import { PNG } from "pngjs";
import type { ReactNode } from "react";

const emuPerPx = 9525;

export type AuthorKind = "jsx" | "document";

export type ScenarioAssertContext = {
  readonly author: AuthorKind;
  readonly resolution: Resolution;
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
};

export type CapabilityScenario = {
  readonly id: string;
  readonly capability: string;
  readonly description: string;
  readonly kind?: "positive" | "rejection";
  readonly visualCheckpoints: readonly string[];
  readonly nativeCheckpoints: readonly string[];
  readonly pages?: readonly number[];
  jsx(): ReactNode;
  document(): CommonNode;
  assert(context: ScenarioAssertContext): void | Promise<void>;
};

export type CollectedScenarioArtifact = {
  readonly scenarioId: string;
  readonly author: AuthorKind;
  readonly page: number;
  readonly slideId: string;
  readonly pngPath?: string;
  readonly pptxPath: string;
};

export type ScenarioRunOutput = {
  readonly scenarioId: string;
  readonly author: AuthorKind;
  readonly resolution: Resolution;
  readonly report: BuildReport;
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
  readonly artifacts: readonly CollectedScenarioArtifact[];
};

const resolveDocument = createResolver(documentCapabilities);

function assertPageOutput(
  page: PreviewPageResult,
  scenario: CapabilityScenario,
  expectedSize: { readonly width?: number; readonly height?: number },
  label: string,
): void {
  const isRejection = scenario.kind === "rejection";
  const hasError = page.diagnostics.some((item) => item.severity === "error");

  if (isRejection && hasError) {
    assert.equal(page.png, undefined, `${label} failed page must omit png`);
    assert.equal(page.svg, undefined, `${label} failed page must omit svg`);
    for (const diagnostic of page.diagnostics) {
      assert.notEqual(diagnostic.path.trim(), "", `${label} diagnostic path`);
      assert.notEqual(
        diagnostic.message.trim(),
        "",
        `${label} diagnostic message`,
      );
    }
    return;
  }

  if (!isRejection) {
    assert.deepEqual(
      page.diagnostics,
      [],
      `${label} positive scenario must have zero preview diagnostics`,
    );
  }

  assert.ok(page.png instanceof Uint8Array, `${label} expected png`);
  assert.ok(
    typeof page.svg === "string" && page.svg.trim().length > 0,
    `${label} expected non-empty svg`,
  );
  const decoded = PNG.sync.read(Buffer.from(page.png));
  assert.ok(decoded.width > 0 && decoded.height > 0, `${label} PNG dimensions`);
  if (expectedSize.width !== undefined && expectedSize.height !== undefined) {
    assert.equal(
      decoded.width,
      expectedSize.width,
      `${label} PNG width must match slide width`,
    );
    assert.equal(
      decoded.height,
      expectedSize.height,
      `${label} PNG height must match slide height`,
    );
  }
}

function assertDualAuthorEquivalence(
  scenario: CapabilityScenario,
  jsxRun: ScenarioRunOutput,
  docRun: ScenarioRunOutput,
): void {
  assert.deepEqual(
    jsxRun.report,
    docRun.report,
    `${scenario.id} build report must match between jsx and document`,
  );
  assert.deepEqual(
    jsxRun.validation.presentation,
    docRun.validation.presentation,
    `${scenario.id} presentation file evidence must match between jsx and document`,
  );
  assert.deepEqual(
    jsxRun.validation.slides,
    docRun.validation.slides,
    `${scenario.id} slide file evidence must match between jsx and document`,
  );
  assert.equal(
    jsxRun.preview.pages.length,
    docRun.preview.pages.length,
    `${scenario.id} preview page count must match between jsx and document`,
  );
  for (let i = 0; i < jsxRun.preview.pages.length; i += 1) {
    const jsxPage = jsxRun.preview.pages[i]!;
    const docPage = docRun.preview.pages[i]!;
    const pageLabel = `${scenario.id} page ${jsxPage.page}`;
    assert.equal(jsxPage.page, docPage.page, `${pageLabel} page number`);
    assert.equal(jsxPage.id, docPage.id, `${pageLabel} slide id`);
    assert.deepEqual(
      jsxPage.diagnostics,
      docPage.diagnostics,
      `${pageLabel} diagnostics`,
    );
    assert.deepEqual(jsxPage.nodes, docPage.nodes, `${pageLabel} nodes`);
    assert.deepEqual(jsxPage.layouts, docPage.layouts, `${pageLabel} layouts`);
    assert.equal(jsxPage.svg, docPage.svg, `${pageLabel} svg`);
    if (jsxPage.png !== undefined) {
      assert.ok(docPage.png !== undefined, `${pageLabel} doc png`);
      assert.equal(
        Buffer.compare(jsxPage.png, docPage.png),
        0,
        `${pageLabel} png bytes must match between jsx and document`,
      );
    } else {
      assert.equal(docPage.png, undefined, `${pageLabel} doc png omitted`);
    }
  }
}

/**
 * Shared runner executing one capability-owned scenario across both real JSX and imperative
 * document authoring entries by resolving once, then feeding emit and paint independently.
 */
export async function runCapabilityScenario(
  scenario: CapabilityScenario,
  options: { readonly outputDir?: string } = {},
): Promise<readonly ScenarioRunOutput[]> {
  const results: ScenarioRunOutput[] = [];
  const previewOptions =
    scenario.pages === undefined ? {} : { pages: scenario.pages };

  for (const author of ["jsx", "document"] as const) {
    const document =
      author === "jsx"
        ? await expandDocument(scenario.jsx())
        : scenario.document();
    const resolution = resolveDocument(document);
    const { pptx, report } = emit(resolution);
    const { pages } = paint(resolution, previewOptions);
    const preview: PreviewResult = Object.freeze({
      report: resolution.report,
      query: resolution.query,
      pages,
    });

    const validation = await validatePptx(pptx);
    assert.deepEqual(
      validation.errors,
      [],
      `${scenario.id} (${author}) PPTX validation failed`,
    );

    const expectedSize = {
      ...(validation.presentation.widthEmu === undefined
        ? {}
        : {
            width: Math.round(validation.presentation.widthEmu / emuPerPx),
          }),
      ...(validation.presentation.heightEmu === undefined
        ? {}
        : {
            height: Math.round(validation.presentation.heightEmu / emuPerPx),
          }),
    };
    for (const page of preview.pages) {
      assertPageOutput(
        page,
        scenario,
        expectedSize,
        `${scenario.id}:${author}:p${page.page}`,
      );
    }

    await scenario.assert({ author, resolution, pptx, preview, validation });

    const artifacts: CollectedScenarioArtifact[] = [];
    if (options.outputDir !== undefined) {
      const baseDir = join(options.outputDir, scenario.id, author);
      await mkdir(baseDir, { recursive: true });
      const pptxPath = join(baseDir, "deck.pptx");
      await writeFile(pptxPath, pptx);
      for (const page of preview.pages) {
        let pngPath: string | undefined;
        if (page.png !== undefined) {
          pngPath = join(baseDir, `page-${page.page}-${page.id}.png`);
          await writeFile(pngPath, page.png);
        }
        artifacts.push({
          scenarioId: scenario.id,
          author,
          page: page.page,
          slideId: page.id,
          ...(pngPath === undefined ? {} : { pngPath }),
          pptxPath,
        });
      }
      await writeFile(
        join(baseDir, "summary.json"),
        JSON.stringify(
          {
            scenarioId: scenario.id,
            capability: scenario.capability,
            author,
            description: scenario.description,
            visualCheckpoints: scenario.visualCheckpoints,
            nativeCheckpoints: scenario.nativeCheckpoints,
            pages: preview.pages.map((page) => ({
              page: page.page,
              id: page.id,
              hasPng: page.png !== undefined,
              diagnostics: page.diagnostics,
              nodes: page.nodes.map((node) => ({
                id: node.id,
                authorId: node.authorId,
                kind: node.kind,
                box: node.box,
                screenBox: node.screenBox,
                path: node.path,
              })),
            })),
          },
          null,
          2,
        ),
      );
    }

    results.push({
      scenarioId: scenario.id,
      author,
      resolution,
      report,
      pptx,
      preview,
      validation,
      artifacts,
    });
  }

  assertDualAuthorEquivalence(scenario, results[0]!, results[1]!);

  return results;
}
