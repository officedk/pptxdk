import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { deck, shape, slide } from "@pptxdk/capabilities/document";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import {
  CapabilityError,
  type CapabilityModule,
  type CommonNode,
  type SlideCapabilityModule,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement, type ReactNode } from "react";
import { describe, expect, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  slideBackgroundPositiveScenario,
  slideBackgroundRejectionScenario,
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

describe("SlideBackground capability acceptance", () => {
  it("runs slide-background dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-bg-accept-"));
    try {
      for (const scenario of [
        slideBackgroundPositiveScenario,
        slideBackgroundRejectionScenario,
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

  it("supports custom bound SlideCapabilityModule with paint and rejects missing background paint", async () => {
    const customSlideCalls: string[] = [];
    const customBackdropModule: SlideCapabilityModule = {
      declaration: "custom-backdrop",
      slideProp: "background",
      slot: "background",
      bindSlide(input, parseContext) {
        const raw = input.props["background"];
        if (typeof raw !== "string" || !/^#[0-9A-Fa-f]{6}$/u.test(raw)) {
          throw new CapabilityError(
            "invalid-background",
            "Custom backdrop expects a 6-digit hex color",
            parseContext.path,
          );
        }
        const hex = raw.slice(1).toUpperCase();
        customSlideCalls.push(`bindSlide:${parseContext.path}`);
        return {
          resolveSlide(resolveContext) {
            customSlideCalls.push(`resolveSlide:${resolveContext.path}`);
            return {
              extensions: [],
              observation: { backdropColor: `#${hex}` },
              emit() {
                return {
                  xml: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
                };
              },
              paint(paintContext) {
                customSlideCalls.push(`paint:${paintContext.path}`);
                return {
                  svg: `<rect width="${paintContext.dimensions.width}" height="${paintContext.dimensions.height}" fill="#${hex}"/>`,
                };
              },
            };
          },
        };
      },
    };

    const backdropDocPreview = previewDocWith([
      customBackdropModule,
      shapeCapability,
    ]);
    const backdropJsxPreview = previewJsxWith([
      customBackdropModule,
      shapeCapability,
    ]);

    const docResult = await backdropDocPreview(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "bg-doc", background: "#0F766E" }).append(
          shape({
            id: "card",
            x: 40,
            y: 20,
            width: 120,
            height: 80,
            fill: "#F59E0B",
          }),
        ),
      ),
    );
    const jsxResult = await backdropJsxPreview(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "bg-jsx", background: "#0F766E" },
          createElement(Shape, {
            id: "card",
            x: 40,
            y: 20,
            width: 120,
            height: 80,
            fill: "#F59E0B",
          }),
        ),
      ),
    );

    for (const result of [docResult, jsxResult]) {
      assert.deepEqual(result.pages[0]?.diagnostics, []);
      assert.deepEqual((await validatePptx(result.pptx)).errors, []);
      assert.deepEqual(result.report.slides[0]?.capabilities, {
        "custom-backdrop": { backdropColor: "#0F766E" },
      });
      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      expect(slideXml).toContain('<a:srgbClr val="0F766E"/>');
      const png = PNG.sync.read(Buffer.from(result.pages[0].png!));
      // Top-left corner (5, 5) shows custom backdrop #0F766E; center (100, 60) shows shape #F59E0B.
      const cornerOffset = (5 * png.width + 5) * 4;
      assert.deepEqual(
        [...png.data.subarray(cornerOffset, cornerOffset + 4)],
        [0x0f, 0x76, 0x6e, 255],
      );
      const centerOffset = (60 * png.width + 100) * 4;
      assert.deepEqual(
        [...png.data.subarray(centerOffset, centerOffset + 4)],
        [0xf5, 0x9e, 0x0b, 255],
      );
    }
    assert.deepEqual(customSlideCalls, [
      "bindSlide:deck/slide[bg-doc]",
      "resolveSlide:deck/slide[bg-doc]",
      "paint:deck/slide[bg-doc]",
      "bindSlide:deck/slide[bg-jsx]",
      "resolveSlide:deck/slide[bg-jsx]",
      "paint:deck/slide[bg-jsx]",
    ]);

    // A background-slot SlideCapabilityModule without paint produces png: undefined + located error.
    const backdropWithoutPreview: SlideCapabilityModule = {
      declaration: "custom-backdrop-no-preview",
      slideProp: "background",
      slot: "background",
      bindSlide() {
        return {
          resolveSlide() {
            return {
              extensions: [],
              emit() {
                return {
                  xml: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
                };
              },
            };
          },
        };
      },
    };
    const noPreviewResult = await previewDocWith([backdropWithoutPreview])(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "no-bg-prev", background: "#1E293B" }),
      ),
    );
    assert.equal(noPreviewResult.pages[0]?.png, undefined);
    assert.deepEqual(noPreviewResult.pages[0]?.diagnostics, [
      {
        severity: "error",
        page: 1,
        path: "deck/slide[no-bg-prev]",
        message: "Preview does not support custom-backdrop-no-preview",
      },
    ]);
  });
});
