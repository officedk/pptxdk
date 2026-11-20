import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { documentCapabilities } from "@react-pptx/capabilities/document";
import {
  BuildError,
  type CapabilityModule,
  type CommonNode,
} from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { expandDocument } from "@react-pptx/react";
import type { ReactNode } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  layoutBlockFlexPositiveScenario,
  layoutBlockFlexRejectionCases,
} from "./scenarios-layout-block-flex.ts";
import {
  layoutGridPositiveScenario,
  layoutGridRejectionCases,
} from "./scenarios-layout-grid.ts";
import { boxPositiveScenario, boxRejectionScenario } from "./scenarios.ts";

function previewDocWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      resolution,
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

function previewJsxWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (element: ReactNode) => {
    const resolution = resolveWith(await expandDocument(element));
    return {
      resolution,
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Box and layout capability acceptance", () => {
  it("runs Box and Block/Flex/Grid dual-author scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-box-accept-"));
    try {
      for (const scenario of [
        boxPositiveScenario,
        boxRejectionScenario,
        layoutBlockFlexPositiveScenario,
        layoutGridPositiveScenario,
      ]) {
        const runs = await runCapabilityScenario(scenario, {
          outputDir: outDir,
        });
        assert.equal(runs.length, 2);
        assert.deepEqual(
          runs.map((r) => r.author),
          ["jsx", "document"],
        );
      }
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 30_000);

  it("enforces unsupported Block/Flex and Grid layout rejections consistently across JSX and document entries", async () => {
    const jsxPreview = previewJsxWith(documentCapabilities);
    const docPreview = previewDocWith(documentCapabilities);
    for (const entry of [
      ...layoutBlockFlexRejectionCases,
      ...layoutGridRejectionCases,
    ]) {
      for (const author of ["jsx", "document"] as const) {
        await assert.rejects(
          () =>
            author === "jsx"
              ? jsxPreview(entry.jsx())
              : docPreview(entry.document()),
          (error: unknown) => {
            assert(
              error instanceof BuildError,
              `${entry.id} (${author}) should throw BuildError`,
            );
            assert.ok(
              error.diagnostics.some(
                (d) =>
                  d.code === entry.expectedCode &&
                  d.path.includes(entry.expectedPathSubstring),
              ),
              `${entry.id} (${author}) expected ${entry.expectedCode} at ${entry.expectedPathSubstring}, got ${JSON.stringify(error.diagnostics)}`,
            );
            return true;
          },
        );
      }
    }
  });
});
