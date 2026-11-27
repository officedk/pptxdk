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
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement, type ReactNode } from "react";
import { describe, expect, it } from "vitest";

import {
  runCapabilityScenario,
  type StageCounters,
} from "../../../test/acceptance-runner.ts";
import {
  customBezierShape,
  shapeAndGroupRoundtripScenario,
} from "./mechanisms.ts";
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

  it("supports registering custom bound extension modules and missing preview rejection", async () => {
    const calls: string[] = [];

    const customMarkerWithPreview: BoundObjectCapabilityModule = {
      declaration: "custom-marker",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const x = Number(input.props["x"]);
        const y = Number(input.props["y"]);
        const size = Number(input.props["size"]);
        const color = String(input.props["color"]);
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: {
            style: {
              position: "absolute",
              left: x,
              top: y,
              width: size,
              height: size,
            },
          },
          allowedStyleKeys: new Set<string>([
            "position",
            "left",
            "top",
            "width",
            "height",
          ]),
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
          allowedStyleKeys: new Set<string>([
            "position",
            "left",
            "top",
            "width",
            "height",
          ]),
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
      slide({ id: "valid-custom" }).append(
        createMarkerNode({
          id: "dot",
          x: 60,
          y: 20,
          size: 80,
          color: "#DC2626",
        }),
      ),
      slide({ id: "missing-preview-slide" }).append(
        new CommonNode("noprev", "custom-no-preview", {
          id: "opaque",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 20,
            height: 20,
          },
        }),
        createMarkerNode({
          id: "later-marker",
          x: 100,
          y: 20,
          size: 40,
          color: "#2563EB",
        }),
      ),
    );

    const result = await previewCustom(doc);
    // Page 1: renders cleanly with zero errors and draws the red circle.
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    const pixels = PNG.sync.read(Buffer.from(result.pages[0].png));
    const center = (60 * pixels.width + 100) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(center, center + 4)],
      [0xdc, 0x26, 0x26, 255],
    );
    assert.equal(
      result.resolution.slide("valid-custom").get("dot").box?.width,
      80,
    );

    // Page 2: missing preview handler records PreviewError and omits PNG.
    assert.equal(result.pages[1]?.png, undefined);
    expect(result.pages[1]?.errors).toEqual([
      {
        code: "preview-unsupported-capability",
        path: "deck/slide[missing-preview-slide]/custom-no-preview[opaque]",
        message: "Preview does not support custom-no-preview",
        capability: "custom-no-preview",
      },
    ]);

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
          assert.equal(error.diagnostics[0]?.code, "missing-capability");
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
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
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
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
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
          new URL(
            "../../../test/assets/fonts/Jost-400-Book.ttf",
            import.meta.url,
          ),
        ),
      },
    ]);
    const resolveWith = createResolver(documentCapabilities);
    const resolution = resolveWith(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "embed-shape" }).append(
          shape("Subset Font", {
            id: "embedded",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 180,
              height: 60,
              background: "#1E293B",
              fontFamily: jostSubsetFont,
              color: "#FFFFFF",
            },
          }),
        ),
      ),
    );
    const emitted = emit(resolution, { embedFonts: "subset" });
    assert.ok(emitted.pptx.byteLength > 0);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png !== undefined);
  });

  it("executes shapeAndGroupRoundtripScenario with single pure-file roundtrip counters and isolated editable instances", async () => {
    const counters: StageCounters = {
      expandJsx: 0,
      buildDocument: 0,
      resolve: 0,
      emit: 0,
      paint: 0,
      import: 0,
      validate: 0,
    };
    const execution = await runCapabilityScenario(
      shapeAndGroupRoundtripScenario,
      { counters },
    );

    assert.deepEqual(counters, {
      expandJsx: 1,
      buildDocument: 1,
      resolve: 3,
      emit: 3,
      paint: 3,
      import: 1,
      validate: 3,
    });

    const findShape = (
      root: CommonNode,
      name: string,
    ): CommonNode | undefined => {
      if (root.type === "shape" && root.props["name"] === name) return root;
      for (const child of root.children) {
        if (child instanceof CommonNode) {
          const found = findShape(child, name);
          if (found !== undefined) return found;
        }
      }
      return undefined;
    };

    const expectedImportedGeometry = {
      type: "custom",
      width: 100,
      height: 100,
      paths: [{ commands: customBezierShape.paths[0]!.commands }],
    };

    const sharedCustom = findShape(
      execution.roundtrip.importedDocument,
      "Custom Radial",
    );
    assert.ok(sharedCustom !== undefined);
    assert.deepEqual(
      JSON.parse(JSON.stringify(sharedCustom.props["geometry"])),
      expectedImportedGeometry,
    );

    const editable = execution.openEditableRoundtripDocument();
    const editableCustom = findShape(editable.document, "Custom Radial");
    assert.ok(editableCustom !== undefined);
    editableCustom.set({ name: "Edited Radial", geometry: "diamond" });

    const editedValidation = await validatePptx(
      emit(createResolver(documentCapabilities)(editable.document)).pptx,
    );
    assert.deepEqual(editedValidation.errors, []);
    const editedObj = editedValidation.slides[0]?.objects.find(
      (o) => o.name === "Edited Radial",
    );
    assert.equal(editedObj?.geometry, "diamond");

    // Mutating the independent editable instance must not pollute the shared read-only roundtrip artifact.
    assert.equal(sharedCustom.props["name"], "Custom Radial");
    assert.deepEqual(
      JSON.parse(JSON.stringify(sharedCustom.props["geometry"])),
      expectedImportedGeometry,
    );
    assert.ok(
      execution.roundtrip.validation.slides[0]?.objects.some(
        (o) => o.name === "Custom Radial" && o.geometry === "custom",
      ),
    );
  });
});
