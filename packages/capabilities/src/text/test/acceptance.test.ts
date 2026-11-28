import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { font } from "@pptxdk/capabilities/document";
import { Text } from "@pptxdk/capabilities/text-elements";
import type { CapabilityModule, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { RichTextContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import {
  textDualScriptScenario,
  textPositiveScenario,
  textRejectionScenario,
  textRtlListsAndFillsScenario,
} from "./scenarios.ts";

function previewDocWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      resolution,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Text capability acceptance", () => {
  it("runs Text dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-text-accept-"));
    try {
      for (const scenario of [
        textPositiveScenario,
        textRtlListsAndFillsScenario,
        textDualScriptScenario,
        textRejectionScenario,
      ]) {
        const { runs } = await runCapabilityScenario(scenario, {
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

  it("supports JSX expansion followed by imperative Text edit, rebuild, and preview without mutating prior outputs", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL(
            "../../../test/assets/fonts/Jost-400-Book.ttf",
            import.meta.url,
          ),
        ),
      },
    ]);
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "handoff" },
          createElement(
            Text,
            {
              id: "headline",
              style: {
                position: "absolute",
                left: 24,
                top: 24,
                width: 200,
                fontFamily: jostFont,
                fontSize: 20,
                color: "#1E293B",
                padding: 8,
                background: "#F1F5F9",
              },
            },
            "Initial title",
          ),
        ),
      ),
    );

    const previewDocument = previewDocWith(documentCapabilities);
    const first = await previewDocument(expanded);
    assert.deepEqual(first.pages[0]?.errors, []);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const firstPng = first.pages[0].png!.slice();
    const firstSvg = first.pages[0].svg!;
    const firstPptx = first.pptx.slice();
    const firstReportSnapshot = JSON.stringify(first.report);
    const firstHeight = first.resolution.slide("handoff").get("headline")
      .box!.height;

    const headline = expanded.query
      .get("handoff")
      .getAs(
        "headline",
        (node): node is RichTextContainer => node instanceof RichTextContainer,
      );
    headline.content("Updated longer headline wrapping across multiple lines");
    headline.style({ color: "#059669" });

    const second = await previewDocument(expanded);
    assert.deepEqual(second.pages[0]?.errors, []);
    assert.deepEqual((await validatePptx(second.pptx)).errors, []);
    assert.equal(
      second.resolution.slide("handoff").get("headline").text,
      "Updated longer headline wrapping across multiple lines",
    );
    const secondHeight = second.resolution.slide("handoff").get("headline")
      .box!.height;
    assert.ok(secondHeight > firstHeight);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // Prior BuildResult and PreviewResult remain unchanged.
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.pptx, firstPptx);
    assert.equal(JSON.stringify(first.report), firstReportSnapshot);
    assert.equal(
      first.resolution.slide("handoff").get("headline").text,
      "Initial title",
    );
  });
});
