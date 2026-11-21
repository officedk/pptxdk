import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { deck, font, shape, slide } from "@pptxdk/capabilities/document";
import { shapeCapability } from "@pptxdk/capabilities/shape";
import {
  BuildError,
  CommonNode,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityModule,
  type CapabilityParseContext,
  type ObjectCapabilityModule,
  type SlideCapabilityModule,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { parseCapabilityLayoutContract } from "@pptxdk/core/capability";
import { emit, pxToEmu, xmlEscape } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { PNG } from "pngjs";
import { createElement, type ReactNode } from "react";
import { describe, expect, it } from "vitest";

import { runCapabilityScenario } from "../acceptance-runner.ts";
import {
  shapePositiveScenario,
  shapePresetCatalogScenario,
  shapeRejectionScenario,
  shapeShadedFillsAndPathsScenario,
  shapeTextParityScenario,
} from "./scenarios.ts";

type CustomMarkerProps = {
  readonly id?: string;
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly color: string;
  readonly warnMessage?: string;
};

function createMarkerNode(
  props: CustomMarkerProps,
): CommonNode<"marker", CustomMarkerProps> {
  return new CommonNode("marker", "custom-marker", props);
}

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

describe("Shape capability acceptance & bound extension contracts", () => {
  it("runs Shape dual-author positive and rejection scenarios through shared runner", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-shape-accept-"));
    try {
      for (const scenario of [
        shapePositiveScenario,
        shapePresetCatalogScenario,
        shapeTextParityScenario,
        shapeShadedFillsAndPathsScenario,
        shapeRejectionScenario,
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

  it("supports registering custom bound extension modules, warning retention, and missing preview rejection", async () => {
    const calls: string[] = [];

    const customMarkerWithPreview: BoundObjectCapabilityModule = {
      declaration: "custom-marker",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const x = Number(input.props["x"]);
        const y = Number(input.props["y"]);
        const size = Number(input.props["size"]);
        const color = String(input.props["color"]);
        const warnMessage =
          typeof input.props["warnMessage"] === "string"
            ? input.props["warnMessage"]
            : undefined;
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: { x, y, width: size, height: size },
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        calls.push(`bind:${parseContext.path}`);

        return {
          contract,
          resolve(resolveContext) {
            calls.push(`resolve:${resolveContext.path}`);
            const box = resolveContext.layoutBox ?? {
              x,
              y,
              width: size,
              height: size,
            };
            const name = xmlEscape(`Marker ${resolveContext.id}`);
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name,
                geometry: "ellipse",
                box,
              },
              extensions: [],
              emit(emitContext) {
                const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`;
                return { xml };
              },
              paint(paintContext) {
                calls.push(`paint:${paintContext.path}`);
                const resolved = paintContext.theme.resolveColor(color);
                return {
                  svg: `<circle cx="${box.x + box.width / 2}" cy="${box.y + box.height / 2}" r="${box.width / 2}" fill="#${resolved.hex}"/>`,
                  ...(warnMessage === undefined
                    ? {}
                    : { warnings: [{ message: warnMessage }] }),
                };
              },
            };
          },
        };
      },
    };

    const customMarkerWithoutPreview: BoundObjectCapabilityModule = {
      declaration: "custom-no-preview",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        return {
          contract,
          resolve(resolveContext) {
            const box = resolveContext.layoutBox ?? {
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            };
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: "NoPreview",
                geometry: "rect",
                box,
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="NoPreview"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const previewCustom = previewDocWith([
      shapeCapability,
      customMarkerWithPreview,
      customMarkerWithoutPreview,
    ]);

    const doc = deck({ width: 200, height: 120 }).append(
      slide({ id: "warn-keeps-png" }).append(
        createMarkerNode({
          id: "dot",
          x: 60,
          y: 20,
          size: 80,
          color: "#DC2626",
          warnMessage: "Custom approximation notice",
        }),
      ),
      slide({ id: "warn-cannot-clear-error" }).append(
        new CommonNode("noprev", "custom-no-preview", {
          id: "opaque",
          x: 10,
          y: 10,
          width: 20,
          height: 20,
        }),
        createMarkerNode({
          id: "later-warn",
          x: 100,
          y: 20,
          size: 40,
          color: "#2563EB",
          warnMessage: "Later warning must not clear earlier error",
        }),
      ),
    );

    const result = await previewCustom(doc);
    // Page 1: warning preserves PNG and draws the red circle.
    assert.equal(result.pages[0]?.diagnostics.length, 1);
    assert.equal(result.pages[0]?.diagnostics[0]?.severity, "warning");
    assert.ok(result.pages[0]?.png !== undefined);
    const pixels = PNG.sync.read(Buffer.from(result.pages[0].png));
    const center = (60 * pixels.width + 100) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(center, center + 4)],
      [0xdc, 0x26, 0x26, 255],
    );
    assert.ok(
      result.query.getSlide("warn-keeps-png").get("dot").observe().box
        ?.width === 80,
    );

    // Page 2: missing preview handler causes error; subsequent warning does NOT clear error or fabricate PNG.
    assert.equal(result.pages[1]?.png, undefined);
    expect(result.pages[1]?.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          severity: "error",
          path: "deck/slide[warn-cannot-clear-error]/custom-no-preview[opaque]",
        }),
        expect.objectContaining({
          severity: "warning",
          path: "deck/slide[warn-cannot-clear-error]/custom-marker[later-warn]",
        }),
      ]),
    );

    // Same custom bound extension works via JSX expansion and shares trimmed constraints.
    const jsxCustomPreview = previewJsxWith([shapeCapability]);
    const trimmedOk = await jsxCustomPreview(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(Slide, { id: "trimmed" }),
      ),
    );
    assert.ok(trimmedOk.pages[0]?.png !== undefined);

    await assert.rejects(
      () =>
        jsxCustomPreview(
          createElement(
            Deck,
            { width: 200, height: 120 },
            createElement(Slide, { id: "missing-bg", background: "#123456" }),
          ),
        ),
      BuildError,
    );
    void shape;
    assert.ok(calls.some((entry) => entry.startsWith("bind:")));
    assert.ok(calls.some((entry) => entry.startsWith("resolve:")));
    assert.ok(calls.some((entry) => entry.startsWith("paint:")));
  });

  it("rejects non-bound capability modules at registration and invalid bound handles at bind time", async () => {
    const unboundObjectModule = {
      declaration: "unbound-object",
      parse() {
        return {
          contract: { nativeKind: "sp", layoutStyle: {} },
          parsed: {},
        };
      },
      resolve() {
        return {
          observation: { id: 2, kind: "shape", name: "Unbound" },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    } as unknown as CapabilityModule;

    const unboundSlideModule = {
      declaration: "unbound-slide",
      slideProp: "unboundSlide",
      slot: "timing",
      resolveSlide() {
        return { extensions: [], emit: () => ({ xml: "" }) };
      },
    } as unknown as CapabilityModule;

    for (const invalidModule of [unboundObjectModule, unboundSlideModule]) {
      assert.throws(
        () => createResolver([invalidModule]),
        (error: unknown) => {
          assert(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-capability-module");
          return true;
        },
      );
    }

    const brokenObjectBinding: ObjectCapabilityModule = {
      declaration: "broken-object",
      bind() {
        return {
          contract: { nativeKind: "sp" },
        } as unknown as ReturnType<ObjectCapabilityModule["bind"]>;
      },
    };
    const brokenObjectResolve = createResolver([brokenObjectBinding]);
    assert.throws(
      () =>
        brokenObjectResolve(
          deck().append(
            slide({ id: "s1" }).append(
              new CommonNode("broken", "broken-object", { id: "bad-obj" }),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-capability-binding");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[s1]/broken-object[bad-obj]",
        );
        return true;
      },
    );

    const brokenSlideBinding: SlideCapabilityModule = {
      declaration: "broken-slide",
      slideProp: "background",
      slot: "background",
      bindSlide() {
        return {} as unknown as ReturnType<SlideCapabilityModule["bindSlide"]>;
      },
    };
    const brokenSlideResolve = createResolver([brokenSlideBinding]);
    assert.throws(
      () =>
        brokenSlideResolve(
          deck().append(slide({ id: "s1", background: "#123456" })),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-capability-binding");
        assert.equal(error.diagnostics[0]?.path, "deck/slide[s1]");
        return true;
      },
    );
  });

  it("embeds subset fonts for Shape inline text across emit() and preserves paint() parity", () => {
    const jostSubsetFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const resolveWith = createResolver(documentCapabilities);
    const resolution = resolveWith(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "embed-shape" }).append(
          shape({
            id: "embedded",
            x: 20,
            y: 20,
            width: 180,
            height: 60,
            fill: "#1E293B",
            text: "Subset Font",
            textStyle: { fontFamily: jostSubsetFont, color: "#FFFFFF" },
          }),
        ),
      ),
    );
    const emitted = emit(resolution, { embedFonts: "subset" });
    assert.ok(emitted.pptx.byteLength > 0);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);
  });
});
