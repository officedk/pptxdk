import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { shapeCapability } from "@pptxdk/capabilities/shape";
import { validatePptx } from "@pptxdk/validation";
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
  formatBaseline,
  observeMotion,
  resolve as resolveFromMain,
  targets,
} from "@react-pptx/react-pptx";
import {
  createResolver,
  emit as emitFromBuild,
  type CapabilityModule,
} from "@react-pptx/react-pptx/build";
import {
  box,
  chart,
  createMorphScope,
  deck,
  emit as emitFromDocument,
  fade,
  resolve as resolveFromDocument,
  section,
  shape,
  slide,
  text,
  transitions,
} from "@react-pptx/react-pptx/document";
import { paint, preview } from "@react-pptx/react-pptx/preview";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: jostBookBytes,
  },
]);

function withFontFsType(bytes: Uint8Array, fsType: number): Uint8Array {
  const copy = new Uint8Array(bytes);
  const view = new DataView(copy.buffer, copy.byteOffset, copy.byteLength);
  const tableCount = view.getUint16(4);
  for (let index = 0; index < tableCount; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...copy.subarray(record, record + 4));
    if (tag === "OS/2") {
      view.setUint16(view.getUint32(record + 8) + 8, fsType);
      return copy;
    }
  }
  throw new Error("Jost fixture must contain an OS/2 table");
}

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
    assert.deepEqual(Object.keys(firstEmit).toSorted(), ["pptx", "report"]);
    assert.equal("query" in firstEmit, false);
    const firstPaint = paint(firstResolution);
    const firstFile = await validatePptx(firstEmit.pptx);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(firstFile.slides.length, 1);
    assert.equal(firstFile.slides[0]?.objects[0]?.text?.text, "Before edit");
    assert.equal(firstPaint.pages.length, 1);
    assert.ok(firstPaint.pages[0]?.svg?.includes('fill="#1D4ED8"'));
    assert.equal(firstPaint.pages[0]?.svg?.includes('fill="#DC2626"'), false);

    // Re-resolving reflects the updated document.
    const secondResolution = resolveFromDocument(doc);
    const secondEmit = emitFromDocument(secondResolution);
    const secondPaint = paint(secondResolution);
    const secondFile = await validatePptx(secondEmit.pptx);
    assert.deepEqual(secondFile.errors, []);
    assert.equal(secondFile.slides.length, 2);
    assert.equal(secondFile.slides[0]?.objects[0]?.text?.text, "After edit");
    assert.equal(secondPaint.pages.length, 2);
    assert.ok(secondPaint.pages[0]?.svg?.includes('fill="#DC2626"'));
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
    assert.deepEqual(Object.keys(previewResult).toSorted(), [
      "pages",
      "query",
      "report",
    ]);
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

  it("never loads ./preview.ts, @pptxdk/preview, or @resvg/resvg-js when importing or running build subpaths", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("@pptxdk/preview") ||
      specifier.endsWith("/packages/preview/src/index.ts") ||
      specifier.endsWith("/preview.ts") ||
      specifier.endsWith("/preview.js") ||
      specifier === "./preview.ts"
    ) {
      throw new Error("build path loaded preview dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build, emit, resolve } = await import("@react-pptx/react-pptx");
const docEntry = await import("@react-pptx/react-pptx/document");
const buildEntry = await import("@react-pptx/react-pptx/build");
await import("@react-pptx/react-pptx/node");
await import("@react-pptx/react-pptx/shape");
await import("@react-pptx/react-pptx/chart");
await import("@react-pptx/react-pptx/text");
void buildEntry;

const doc = docEntry.deck({ width: 320, height: 180 }).append(
  docEntry.slide({ id: "s1" }).append(
    docEntry.shape({ id: "box", x: 10, y: 10, width: 40, height: 40 }),
  ),
);
const resolution = resolve(doc);
const emitted = emit(resolution);
const built = await build(doc);
const docBuilt = await docEntry.build(doc);
if (
  emitted.pptx.byteLength === 0 ||
  built.pptx.byteLength === 0 ||
  docBuilt.pptx.byteLength === 0
) {
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

  it("never loads react, react-reconciler, or @pptxdk/react from any @pptxdk/core subpath and keeps host operations off public core entries", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/emit") ||
      specifier.includes("@pptxdk/preview")
    ) {
      throw new Error("core subpath loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const coreMain = await import("@pptxdk/core");
const coreBuild = await import("@pptxdk/core/build");
const coreCap = await import("@pptxdk/core/capability");
const coreDoc = await import("@pptxdk/core/document");
const coreHost = await import("@pptxdk/core/host");
await import("@pptxdk/core/query");
await import("@pptxdk/core/resolution");
await import("@pptxdk/core/theme");
void coreMain;
void coreBuild;
void coreCap;
void coreDoc;

for (const key of [
  "hostNodeCreator",
  "insertDeclarationChild",
  "removeDeclarationChild",
  "replaceDeclarationProps",
  "setDeclarationText",
  "validateAuthorIdScopes",
]) {
  if (!(key in coreHost)) {
    throw new Error("@pptxdk/core/host must export " + key);
  }
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

  it("never loads react or react-reconciler when running imperative paint via @pptxdk/preview", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/")
    ) {
      throw new Error("preview package loaded React dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { paint } = await import("@pptxdk/preview");
const { createResolver } = await import("@pptxdk/core/build");
const { createDeckNode, createSlideNode } = await import("@pptxdk/core/document");
const { shapeCapability, createShapeNode } = await import("@pptxdk/capabilities/shape");

const d = createDeckNode({ width: 320, height: 180 });
const s = createSlideNode({ id: "s1" });
s.append(createShapeNode({ id: "box", x: 10, y: 10, width: 40, height: 40, fill: "#2563EB" }));
d.append(s);

const resolution = createResolver([shapeCapability])({ children: [d] });
const painted = paint(resolution);
if (!painted.pages[0]?.png || painted.pages[0].png.byteLength === 0) {
  throw new Error("expected PNG output");
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
      declaration: "custom-opaque",
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
        new CommonNode("opaque", "custom-opaque", { id: "op1" }),
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
        "Preview does not support custom-opaque",
        "Preview does not support custom-opaque-bg",
      ],
    );
    assert.ok(painted.pages[1]?.png !== undefined);
    assert.deepEqual(painted.pages[1]?.diagnostics, []);
  });

  it("splits ResolutionReport from BuildReport, rejects target on resolve, validates target on emit, and emits one Resolution to two targets", async () => {
    const doc = deck({
      title: "Multi-target Deck",
      width: 320,
      height: 180,
    }).append(
      slide({
        id: "s1",
        background: "#F8FAFC",
        transition: transitions.fade({ duration: 500 }),
      }).append(
        shape({
          id: "hero",
          x: 20,
          y: 20,
          width: 100,
          height: 50,
          fill: "#2563EB",
          enter: fade(),
        }),
      ),
    );

    // Runtime callers passing target to resolve() must fail immediately with unknown-parameter.
    assert.throws(
      () => resolveFromMain(doc, { target: targets.baselineOnly } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "unknown-parameter",
    );

    const resolution = resolveFromMain(doc);
    assert.equal(resolution.report.evidence, "generation");
    assert.equal(resolution.report.title, "Multi-target Deck");
    assert.deepEqual(resolution.report.extensions, ["p14"]);
    assert.equal("target" in resolution.report, false);

    // Invalid target shapes are rejected by emit().
    for (const invalidTarget of [
      null,
      { name: "", format: formatBaseline, extensions: [] },
      { name: "bad-ext", format: formatBaseline, extensions: ["p99"] },
      {
        name: "extra-key",
        format: formatBaseline,
        extensions: [],
        unexpected: true,
      },
    ]) {
      assert.throws(
        () => emitFromMain(resolution, { target: invalidTarget as never }),
        BuildError,
      );
    }

    const p14OnlyTarget = {
      name: "p14-only",
      format: formatBaseline,
      extensions: ["p14"] as const,
    };
    const firstEmit = emitFromMain(resolution, { target: p14OnlyTarget });
    const secondEmit = emitFromMain(resolution, {
      target: targets.transitional,
    });

    assert.deepEqual(firstEmit.report.target, {
      name: "p14-only",
      format: formatBaseline,
      extensions: ["p14"],
    });
    assert.deepEqual(firstEmit.report.extensions, ["p14"]);
    assert.equal(firstEmit.report.title, "Multi-target Deck");
    assert.equal(firstEmit.report.slides, resolution.report.slides);
    assert.ok(observeMotion(firstEmit.report.slides[0]!) !== undefined);

    assert.deepEqual(secondEmit.report.target, {
      name: targets.transitional.name,
      format: targets.transitional.format,
      extensions: ["p14", "p159", "c15"],
    });
    assert.deepEqual(secondEmit.report.extensions, ["p14"]);
    assert.equal(secondEmit.report.slides, resolution.report.slides);
    assert.ok(observeMotion(secondEmit.report.slides[0]!) !== undefined);

    const firstFile = await validatePptx(firstEmit.pptx, {
      allowedExtensions: ["p14"],
    });
    const secondFile = await validatePptx(secondEmit.pptx, {
      allowedExtensions: ["p14", "p159", "c15"],
    });
    assert.deepEqual(firstFile.errors, []);
    assert.deepEqual(secondFile.errors, []);
  });

  it("allows resolve, paint, and preview to succeed when target forbids required mechanisms while emit throws forbidden-mechanism", async () => {
    const doc = deck({ width: 320, height: 180 }).append(
      section({ name: "Part 1" }).append(
        slide({ id: "intro", background: "#FFFFFF" }).append(
          shape({
            id: "box-1",
            x: 20,
            y: 20,
            width: 80,
            height: 50,
            fill: "#2563EB",
          }),
        ),
        slide({
          id: "morph-slide",
          background: "#FFFFFF",
          transition: transitions.morph({
            duration: 800,
            fallback: transitions.fade(),
          }),
        }).append(
          text("Morph target", {
            id: "label-1",
            x: 20,
            y: 20,
            width: 160,
            height: 40,
            style: { fontFamily: jostFont, fontSize: 16 },
          }),
        ),
      ),
    );

    const resolution = resolveFromMain(doc);
    assert.deepEqual(
      new Set(resolution.report.extensions),
      new Set(["p14", "p159"]),
    );
    assert.equal("target" in resolution.report, false);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 2);
    for (const page of painted.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(typeof page.svg === "string" && page.svg.length > 0);
      assert.ok(page.png instanceof Uint8Array && page.png.byteLength > 0);
    }

    const previewResult = await preview(doc);
    assert.deepEqual(
      new Set(previewResult.report.extensions),
      new Set(["p14", "p159"]),
    );
    assert.equal("target" in previewResult.report, false);
    assert.equal(previewResult.pages.length, 2);
    for (const page of previewResult.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png instanceof Uint8Array && page.png.byteLength > 0);
    }

    assert.throws(
      () => emitFromMain(resolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "forbidden-mechanism"),
    );

    const emitted = emitFromMain(resolution, { target: targets.transitional });
    const file = await validatePptx(emitted.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, []);

    // c15 chart resolves cleanly and rejects targets omitting c15 during emit.
    const c15Doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "c15-slide", background: "#FFFFFF" }).append(
        chart({
          id: "rings",
          x: 20,
          y: 20,
          width: 320,
          height: 200,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [
                { name: "R1", categories: ["A", "B"], values: [10, 20] },
                { name: "R2", categories: ["B", "A"], values: [15, 25] },
              ],
              dataLabels: { value: true, category: true },
            },
          ],
        }),
      ),
    );
    const c15Resolution = resolveFromMain(c15Doc);
    assert.deepEqual(c15Resolution.report.extensions, ["c15"]);
    assert.throws(
      () => emitFromMain(c15Resolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "forbidden-mechanism"),
    );
    assert.throws(
      () =>
        emitFromMain(c15Resolution, {
          target: {
            name: "no-c15",
            format: formatBaseline,
            extensions: ["p14", "p159"],
          },
        }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "forbidden-mechanism"),
    );
    const c15Emitted = emitFromMain(c15Resolution, {
      target: targets.transitional,
    });
    const c15File = await validatePptx(c15Emitted.pptx, {
      allowedExtensions: ["c15"],
    });
    assert.deepEqual(c15File.errors, []);
  });

  it("rejects invalid documents directly during resolve without waiting for emit", () => {
    // 1. Declaration parse error
    const invalidParseDoc = deck({ width: 320, height: 180 }).append(
      slide({
        id: "bad-transition",
        transition: transitions.fade({ duration: -1 }),
      }),
    );
    assert.throws(
      () => resolveFromMain(invalidParseDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-transition-duration",
    );

    // 2. Author ID conflict
    const duplicateIdDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "dup-slide" }).append(
        shape({ id: "dup", x: 10, y: 10, width: 40, height: 40 }),
        shape({ id: "dup", x: 60, y: 10, width: 40, height: 40 }),
      ),
    );
    assert.throws(
      () => resolveFromMain(duplicateIdDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "duplicate-author-id",
    );

    // 3. Morph continuity object kind mismatch (sp vs graphicFrame)
    const morph = createMorphScope();
    const hero = morph.handle("hero");
    const morphMismatchDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "m1" }).append(
        shape({
          id: "hero-shape",
          morphKey: hero,
          x: 20,
          y: 20,
          width: 80,
          height: 50,
        }),
      ),
      slide({
        id: "m2",
        transition: transitions.morph({ fallback: transitions.fade() }),
      }).append(
        chart({
          id: "hero-chart",
          morphKey: hero,
          x: 20,
          y: 20,
          width: 200,
          height: 150,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S1", values: [1, 2] }],
            },
          ],
        }),
      ),
    );
    assert.throws(
      () => resolveFromMain(morphMismatchDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "morph-object-kind-mismatch",
    );

    // 4. Layout error
    const invalidLayoutDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "bad-layout" }).append(
        box({
          id: "negative-padding",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: 80,
            display: "flex",
            padding: -10,
          },
        }),
      ),
    );
    assert.throws(
      () => resolveFromMain(invalidLayoutDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-layout-style",
    );
  });

  it("rejects embedFonts in resolve, allows resolve/paint/preview/unembedded emit with restricted font license, and fails only when emit embeds", async () => {
    const restrictedFont = font("Jost*", [
      { weight: 400, data: withFontFsType(jostBookBytes, 0x0002) },
    ]);
    const restrictedDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "restricted-slide", background: "#FFFFFF" }).append(
        text("Licensed text", {
          id: "licensed-text",
          x: 20,
          y: 20,
          width: 240,
          height: 48,
          style: { fontFamily: restrictedFont, fontSize: 18, color: "#0F172A" },
        }),
      ),
    );

    // Passing embedFonts to resolve() must fail immediately with unknown-parameter.
    for (const mode of ["full", "subset"] as const) {
      assert.throws(
        () => resolveFromMain(restrictedDoc, { embedFonts: mode } as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "unknown-parameter",
      );
    }

    // Resolve, paint, preview, and unembedded emit all succeed.
    const resolution = resolveFromMain(restrictedDoc);
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.svg?.includes("<path"));
    assert.ok(
      painted.pages[0]?.png instanceof Uint8Array &&
        painted.pages[0].png.byteLength > 0,
    );

    const previewed = await preview(restrictedDoc);
    assert.equal(previewed.pages.length, 1);
    assert.deepEqual(previewed.pages[0]?.diagnostics, []);
    assert.ok(
      previewed.pages[0]?.png instanceof Uint8Array &&
        previewed.pages[0].png.byteLength > 0,
    );

    const unembedded = emitFromMain(resolution);
    const unembeddedValidation = await validatePptx(unembedded.pptx);
    assert.deepEqual(unembeddedValidation.errors, []);

    // Requesting full or subset font embedding at emit() throws BuildError.
    for (const mode of ["full", "subset"] as const) {
      assert.throws(
        () => emitFromMain(resolution, { embedFonts: mode }),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "font-embedding-restricted",
      );
    }

    // Subsetting-forbidden fsType (0x0100) still allows full embedding at emit, failing only on subset.
    const noSubsetFont = font("Jost*", [
      { weight: 400, data: withFontFsType(jostBookBytes, 0x0100) },
    ]);
    const noSubsetDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "no-subset-slide", background: "#FFFFFF" }).append(
        text("Full only", {
          id: "full-only-text",
          x: 20,
          y: 20,
          width: 240,
          height: 48,
          style: { fontFamily: noSubsetFont, fontSize: 18 },
        }),
      ),
    );
    const noSubsetResolution = resolveFromMain(noSubsetDoc);
    const fullEmitted = emitFromMain(noSubsetResolution, {
      embedFonts: "full",
    });
    assert.deepEqual((await validatePptx(fullEmitted.pptx)).errors, []);
    assert.throws(
      () => emitFromMain(noSubsetResolution, { embedFonts: "subset" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "font-subsetting-forbidden",
    );
  });

  it("resolves, emits, and paints tableCapability and chartCapability text outlines without registering or loading textCapability", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "@pptxdk/capabilities/text" ||
      specifier === "./text.ts" ||
      specifier === "./text.js" ||
      specifier.endsWith("/capabilities/src/text/text.ts") ||
      specifier.endsWith("/capabilities/dist/text/text.js")
    ) {
      throw new Error("Loaded forbidden text capability module: " + specifier);
    }
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js")
    ) {
      throw new Error("Loaded forbidden text capability module: " + resolved.url);
    }
    return resolved;
  },
});

const { createResolver } = await import("@pptxdk/core/build");
const { createDeckNode, createSlideNode } = await import("@pptxdk/core/document");
const {
  tableCapability,
  createTableNode,
  createRowNode,
  createCellNode,
} = await import("@pptxdk/capabilities/table");
const { chartCapability, createChartNode } = await import("@pptxdk/capabilities/chart");
const { font } = await import("@pptxdk/capabilities/font");
const { emit } = await import("@pptxdk/emit");
const { paint } = await import("@pptxdk/preview");

const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(readFileSync("./test/assets/fonts/Jost-400-Book.ttf")),
  },
]);

const doc = createDeckNode({ width: 640, height: 360 });
const s1 = createSlideNode({ id: "table-and-chart" });
const tbl = createTableNode({
  id: "tbl",
  x: 20,
  y: 20,
  width: 240,
  columns: [120, 120],
  style: { fontFamily: jost, fontSize: 16, color: "#111827" },
});
const row1 = createRowNode({ style: { height: 40 } });
row1.append(createCellNode({}).content("Alpha"), createCellNode({}).content("Beta"));
tbl.append(row1);
s1.append(
  tbl,
  createChartNode({
    id: "cht",
    x: 280,
    y: 20,
    width: 320,
    height: 240,
    title: "Sales",
    legend: "bottom",
    style: { fontFamily: jost, fontSize: 14 },
    groups: [
      {
        kind: "column",
        data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
        series: [{ name: "Actual", values: [12, 24] }],
      },
    ],
  }),
);
doc.append(s1);

const resolution = createResolver([tableCapability, chartCapability])({ children: [doc] });
const emitted = emit(resolution);
if (emitted.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output");
}
const painted = paint(resolution);
const page = painted.pages[0];
if (!page || page.diagnostics.some((diag) => diag.severity === "error")) {
  throw new Error("Expected clean page paint without errors");
}
if (!page.png || page.png.byteLength === 0 || !page.svg?.includes("<path")) {
  throw new Error("Expected rendered glyph outline SVG and PNG");
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
});
