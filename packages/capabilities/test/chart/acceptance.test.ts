import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@react-pptx/capabilities";
import { chartCapability } from "@react-pptx/capabilities/chart";
import { Chart } from "@react-pptx/capabilities/chart-elements";
import {
  chart,
  deck,
  font,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import {
  BuildError,
  type CapabilityModule,
  type CommonNode,
} from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { createElement, type ReactNode } from "react";
import { describe, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  chartAxesAndCompactScenario,
  chartBarColumnVariantsScenario,
  chartComboDualAxesScenario,
  chartLineAreaParityScenario,
  chartPositiveScenario,
  chartRejectionScenario,
  chartScatterParityScenario,
} from "./scenarios.ts";

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

describe("Chart capability acceptance", () => {
  it("runs Chart dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-chart-accept-"));
    try {
      for (const scenario of [
        chartPositiveScenario,
        chartAxesAndCompactScenario,
        chartBarColumnVariantsScenario,
        chartLineAreaParityScenario,
        chartScatterParityScenario,
        chartComboDualAxesScenario,
        chartRejectionScenario,
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
  }, 15_000);

  it("enforces [slideBackgroundCapability, chartCapability] trimming across JSX and document entries", async () => {
    const chartFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const chartOnlyModules = [
      slideBackgroundCapability,
      chartCapability,
    ] as const;
    const chartTrimmedDoc = previewDocWith(chartOnlyModules);
    const chartTrimmedJsx = previewJsxWith(chartOnlyModules);
    const chartGroup = {
      kind: "column" as const,
      data: { kind: "shared-categories" as const, categories: ["Q1", "Q2"] },
      series: [{ name: "Sales", values: [12, 24] }],
    };
    const chartDocOut = await chartTrimmedDoc(
      deck({ width: 320, height: 200 }).append(
        slide({ id: "chart-trim", background: "#FFFFFF" }).append(
          chart({
            id: "c1",
            x: 20,
            y: 20,
            width: 280,
            height: 160,
            style: { fontFamily: chartFont },
            groups: [chartGroup],
          }),
        ),
      ),
    );
    const chartJsxOut = await chartTrimmedJsx(
      createElement(
        Deck,
        { width: 320, height: 200 },
        createElement(
          Slide,
          { id: "chart-trim", background: "#FFFFFF" },
          createElement(Chart, {
            id: "c1",
            x: 20,
            y: 20,
            width: 280,
            height: 160,
            style: { fontFamily: chartFont },
            groups: [chartGroup],
          }),
        ),
      ),
    );
    assert.ok(chartDocOut.pages[0]?.png !== undefined);
    assert.ok(chartJsxOut.pages[0]?.png !== undefined);
    assert.deepEqual((await validatePptx(chartDocOut.pptx)).errors, []);
    assert.deepEqual((await validatePptx(chartJsxOut.pptx)).errors, []);
    await assert.rejects(
      () =>
        chartTrimmedDoc(
          deck({ width: 320, height: 200 }).append(
            slide({ id: "missing-shape" }).append(
              shape({ id: "s1", x: 10, y: 10, width: 50, height: 50 }),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        return true;
      },
    );
  });
});
