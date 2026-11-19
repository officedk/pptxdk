import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";

import { documentCapabilities } from "@react-pptx/capabilities/document";
import type { CommonNode } from "@react-pptx/core";
import { createResolver, type Resolution } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import {
  paint,
  type PreviewPageResult,
  type PreviewResult,
} from "@react-pptx/preview";
import { expandDocument } from "@react-pptx/react";
import { validatePptx, type ValidationResult } from "@react-pptx/validation";
import { PNG } from "pngjs";
import type { ReactNode } from "react";

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
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
  readonly artifacts: readonly CollectedScenarioArtifact[];
};

const resolveDocument = createResolver(documentCapabilities);

function assertPngValid(page: PreviewPageResult, label: string): void {
  const hasError = page.diagnostics.some((item) => item.severity === "error");
  if (hasError) {
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
  assert.ok(page.png instanceof Uint8Array, `${label} expected png`);
  assert.ok(typeof page.svg === "string", `${label} expected svg`);
  const decoded = PNG.sync.read(Buffer.from(page.png));
  assert.ok(decoded.width > 0 && decoded.height > 0, `${label} PNG dimensions`);
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
    const { pptx } = emit(resolution);
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

    for (const page of preview.pages) {
      assertPngValid(page, `${scenario.id}:${author}:p${page.page}`);
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
      pptx,
      preview,
      validation,
      artifacts,
    });
  }

  return results;
}
