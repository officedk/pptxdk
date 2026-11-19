import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { shapeCapability } from "@react-pptx/capabilities/shape";
import {
  BuildError,
  CommonNode,
  Deck,
  Shape,
  Slide,
  Text,
  emit as emitFromMain,
  expandDocument,
  font,
  resolve as resolveFromMain,
} from "@react-pptx/react-pptx";
import {
  createResolver,
  emit as emitFromBuild,
  type CapabilityModule,
} from "@react-pptx/react-pptx/build";
import {
  deck,
  emit as emitFromDocument,
  resolve as resolveFromDocument,
  shape,
  slide,
  text,
} from "@react-pptx/react-pptx/document";
import { paint, preview } from "@react-pptx/react-pptx/preview";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

describe("Seam 1 & Seam 3: public resolve / emit / paint & createResolver", () => {
  it("isolates Resolution from subsequent editable document edits across emit and paint", async () => {
    const label = text("Before edit", {
      id: "headline",
      x: 20,
      y: 20,
      width: 220,
      height: 40,
      style: { fontFamily: jostFont, fontSize: 16 },
    });
    const card = shape({
      id: "card",
      x: 20,
      y: 80,
      width: 120,
      height: 50,
      fill: "#1D4ED8",
    });
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "s1", background: "#F8FAFC" }).append(label, card),
    );

    const firstResolution = resolveFromMain(doc);
    assert.equal(
      firstResolution.query.getSlide("s1").get("headline").observe().text,
      "Before edit",
    );

    // Mutate the editable document after resolve().
    label.content("After edit");
    card.set({ fill: "#DC2626" });
    doc.append(slide({ id: "s2" }));

    const firstEmit = emitFromMain(firstResolution);
    const firstPaint = paint(firstResolution);
    const firstFile = await validatePptx(firstEmit.pptx);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(firstFile.slides.length, 1);
    assert.equal(firstFile.slides[0]?.objects[0]?.text?.text, "Before edit");
    assert.equal(firstPaint.pages.length, 1);
    expect(firstPaint.pages[0]?.svg).toContain('fill="#1D4ED8"');
    expect(firstPaint.pages[0]?.svg).not.toContain('fill="#DC2626"');

    // Re-resolving reflects the updated document.
    const secondResolution = resolveFromDocument(doc);
    const secondEmit = emitFromDocument(secondResolution);
    const secondPaint = paint(secondResolution);
    const secondFile = await validatePptx(secondEmit.pptx);
    assert.deepEqual(secondFile.errors, []);
    assert.equal(secondFile.slides.length, 2);
    assert.equal(secondFile.slides[0]?.objects[0]?.text?.text, "After edit");
    assert.equal(secondPaint.pages.length, 2);
    expect(secondPaint.pages[0]?.svg).toContain('fill="#DC2626"');
  });

  it("scopes pages selection to paint and preview without omitting build slides or returning preview.build", async () => {
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "p1", background: "#FFFFFF" }).append(
        shape({
          id: "s1",
          x: 10,
          y: 10,
          width: 50,
          height: 40,
          fill: "#111111",
        }),
      ),
      slide({ id: "p2", background: "#FFFFFF" }).append(
        shape({
          id: "s2",
          x: 10,
          y: 10,
          width: 50,
          height: 40,
          fill: "#222222",
        }),
      ),
    );

    const resolution = resolveFromDocument(doc);
    assert.equal(paint(resolution).pages.length, 2);
    assert.equal(paint(resolution, { pages: [] }).pages.length, 0);
    assert.deepEqual(
      paint(resolution, { pages: [2] }).pages.map((page) => page.id),
      ["p2"],
    );

    for (const invalidPage of [0, -1, 1.5, 3]) {
      assert.throws(
        () => paint(resolution, { pages: [invalidPage] }),
        RangeError,
      );
      await assert.rejects(
        () => preview(doc, { pages: [invalidPage] }),
        RangeError,
      );
    }

    const previewResult = await preview(doc, { pages: [2] });
    assert.equal("build" in previewResult, false);
    assert.equal(previewResult.report.slides.length, 2);
    assert.equal(previewResult.query.slides.length, 2);
    assert.deepEqual(
      previewResult.pages.map((page) => page.id),
      ["p2"],
    );

    // Emit still writes the full deck regardless of paint page selection.
    const emitted = emitFromDocument(resolution);
    const file = await validatePptx(emitted.pptx);
    assert.equal(file.slides.length, 2);
  });

  it("converges JSX expansion and imperative document authoring into matching resolution reports and painted canvases", async () => {
    const jsxTree = createElement(
      Deck,
      { width: 320, height: 180 },
      createElement(
        Slide,
        { id: "shared", background: "#F1F5F9" },
        createElement(Shape, {
          id: "card",
          x: 24,
          y: 24,
          width: 120,
          height: 72,
          fill: "#0F766E",
        }),
        createElement(
          Text,
          {
            id: "label",
            x: 24,
            y: 112,
            width: 160,
            height: 32,
            style: { fontFamily: jostFont, fontSize: 16 },
          },
          "Shared stage",
        ),
      ),
    );
    const imperativeDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "shared", background: "#F1F5F9" }).append(
        shape({
          id: "card",
          x: 24,
          y: 24,
          width: 120,
          height: 72,
          fill: "#0F766E",
        }),
        text("Shared stage", {
          id: "label",
          x: 24,
          y: 112,
          width: 160,
          height: 32,
          style: { fontFamily: jostFont, fontSize: 16 },
        }),
      ),
    );

    const expanded = await expandDocument(jsxTree);
    const jsxResolution = resolveFromMain(expanded);
    const docResolution = resolveFromDocument(imperativeDoc);

    assert.deepEqual(jsxResolution.report, docResolution.report);
    const jsxPaint = paint(jsxResolution);
    const docPaint = paint(docResolution);
    assert.equal(jsxPaint.pages[0]?.svg, docPaint.pages[0]?.svg);
    assert.deepEqual(jsxPaint.pages[0]?.png, docPaint.pages[0]?.png);
  });

  it("never loads ./preview.ts or @resvg/resvg-js when importing or running resolve/emit/build on the main entry", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@resvg/resvg-js") ||
      specifier.endsWith("/preview.ts") ||
      specifier.endsWith("/preview.js") ||
      specifier === "./preview.ts"
    ) {
      throw new Error("main entry loaded preview dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build, emit, resolve } = await import("@react-pptx/react-pptx");
const { deck, slide, shape } = await import("@react-pptx/react-pptx/document");
const doc = deck({ width: 320, height: 180 }).append(
  slide({ id: "s1" }).append(
    shape({ id: "box", x: 10, y: 10, width: 40, height: 40 }),
  ),
);
const resolution = resolve(doc);
const emitted = emit(resolution);
const built = await build(doc);
if (emitted.pptx.byteLength === 0 || built.pptx.byteLength === 0) {
  throw new Error("expected PPTX bytes");
}
process.stdout.write("ok");
`;
    const child = spawnSync(
      process.execPath,
      ["--conditions=source", "--input-type=module", "--eval", script],
      { cwd: packageRoot, encoding: "utf8" },
    );
    assert.equal(child.status, 0, child.stderr);
    assert.equal(child.stdout, "ok");
  });

  it("assembles trimmed modules with createResolver, validates roots, and reports unsupported paint without blocking emit", async () => {
    const customNoPaintObject: CapabilityModule = {
      declaration: "rp:custom-opaque",
      bind() {
        return {
          contract: {
            nativeKind: "sp",
            layoutStyle: {},
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            isContainer: false,
            hasExplicitSize: true,
          },
          resolve(context) {
            return {
              observation: {
                id: context.id,
                kind: "shape",
                name: "Opaque",
                box: { x: 10, y: 10, width: 60, height: 40 },
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="Opaque"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="571500" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const customNoPaintBackground: CapabilityModule = {
      declaration: "custom-opaque-bg",
      slideProp: "background",
      slot: "background",
      bindSlide() {
        return {
          resolveSlide() {
            return {
              extensions: [],
              emit() {
                return {
                  xml: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
                };
              },
            };
          },
        };
      },
    };

    const resolveCustom = createResolver([
      shapeCapability,
      customNoPaintObject,
      customNoPaintBackground,
    ]);

    // Invalid root CommonNode throws BuildError ("document-root").
    assert.throws(
      () => resolveCustom(slide({ id: "not-a-deck" })),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "document-root");
        return true;
      },
    );

    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "s1", background: "#112233" }).append(
        new CommonNode("opaque", "rp:custom-opaque", { id: "op1" }),
      ),
      slide({ id: "s2" }).append(
        shape({
          id: "ok",
          x: 10,
          y: 10,
          width: 60,
          height: 40,
          fill: "#2563EB",
        }),
      ),
    );

    const resolution = resolveCustom(doc);
    const emitted = emitFromBuild(resolution);
    const file = await validatePptx(emitted.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides.length, 2);

    const painted = paint(resolution);
    assert.equal(painted.pages[0]?.svg, undefined);
    assert.equal(painted.pages[0]?.png, undefined);
    assert.deepEqual(
      painted.pages[0]?.diagnostics.map((diag) => diag.message),
      [
        "Preview does not support rp:custom-opaque",
        "Preview does not support custom-opaque-bg",
      ],
    );
    assert.ok(painted.pages[1]?.png !== undefined);
    assert.deepEqual(painted.pages[1]?.diagnostics, []);
  });
});
