import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { build } from "@officedk/pptxdk";
import { Deck, Slide, Text } from "@officedk/pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const png1x1Base64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=";
const jpeg1x1Base64 =
  "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAFR/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIRAxEAPwD/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2Q==";

function runIsolated(conditions: readonly string[], script: string): string {
  const child = spawnSync(
    process.execPath,
    [...conditions, "--input-type=module", "--eval", script],
    { cwd: packageRoot, encoding: "utf8" },
  );
  assert.equal(child.status, 0, child.stderr || child.stdout);
  return child.stdout;
}

describe("@officedk/pptxdk centralized on-demand module isolation", () => {
  it("runs vanilla build, resolve, emit, importPptx, codegen, and @pptxdk/core without loading React, preview, resvg, image decoders, validation, or cli-contract across source and dist", () => {
    const script = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier === "@pptxdk/capabilities" ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@pptxdk/validation") ||
      specifier.includes("cli-contract") ||
      specifier.endsWith("/preview.ts") ||
      specifier.endsWith("/preview.js") ||
      specifier === "./preview.ts" ||
      specifier === "./preview.js"
    ) {
      throw new Error("Non-React/non-preview path loaded forbidden module: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const root = await import("@officedk/pptxdk");
const vanilla = await import("@officedk/pptxdk/vanilla");
const buildEntry = await import("@officedk/pptxdk/build");
await import("@officedk/pptxdk/node");
const coreHost = await import("@pptxdk/core/host");
await import("@pptxdk/core");
await import("@pptxdk/core/build");
await import("@pptxdk/core/capability");
await import("@pptxdk/core/document");
await import("@pptxdk/core/query");
await import("@pptxdk/core/resolution");
await import("@pptxdk/core/theme");

for (const key of [
  "hostNodeCreator",
  "insertDeclarationChild",
  "removeDeclarationChild",
  "replaceDeclarationProps",
  "setDeclarationText",
  "validateAuthorIdScopes",
]) {
  assert.ok(key in coreHost, "@pptxdk/core/host missing " + key);
}

assert.equal(typeof root.preview, "function");
assert.equal(typeof root.PreviewPageRangeError, "function");
assert.equal(typeof root.codegen, "function");
assert.equal(typeof root.CodegenError, "function");
assert.equal(root.codegenTargetSchema.parse("vanilla"), "vanilla");
assert.equal(typeof buildEntry.createResolver, "function");
assert.equal(typeof buildEntry.emit, "function");

const png = new Uint8Array(Buffer.from("${png1x1Base64}", "base64"));
const jpeg = new Uint8Array(Buffer.from("${jpeg1x1Base64}", "base64"));
const doc = vanilla.deck({ width: 960, height: 540 }).append(
  vanilla.slide({ id: "s1", background: "#112233", transition: root.transitions.fade() }).append(
    vanilla.box().style({ width: 320, height: 80 }).append(
      vanilla.text("No React", { id: "t1", style: { position: "absolute", left: 8, top: 8, width: 200, height: 32 } }),
    ),
    vanilla.shape("Shape", { id: "sh1", style: { position: "absolute", left: 8, top: 100, width: 80, height: 40, background: "#336699" } }),
    vanilla.image({ id: "img-png", src: png, fit: "cover", style: { position: "absolute", left: 100, top: 100, width: 40, height: 40 } }),
    vanilla.image({ id: "img-jpg", src: jpeg, style: { position: "absolute", left: 150, top: 100, width: 40, height: 40 } }),
    vanilla.group({ id: "g1" }).append(
      vanilla.shape({ id: "g-sh", style: { position: "absolute", left: 200, top: 100, width: 40, height: 40 } }),
    ),
    vanilla.table({ id: "tbl1", columns: [80, "1fr"], style: { position: "absolute", left: 8, top: 160, width: 200, fontSize: 12 } }).append(
      vanilla.row({ style: { height: 24 } }).append(vanilla.cell("A"), vanilla.cell("B")),
    ),
    vanilla.chart({
      id: "cht1",
      style: { position: "absolute", left: 240, top: 160, width: 200, height: 120 },
      groups: [{ kind: "column", data: { kind: "shared-categories", categories: ["Q1"] }, series: [{ name: "S", values: [1] }] }],
    }),
  ),
);

const resolution = root.resolve(doc);
const emitted = root.emit(resolution);
const built = await root.build(doc);
assert.ok(emitted.pptx.byteLength > 0 && built.pptx.byteLength > 0);
assert.equal(built.report.slides[0]?.objects.length, 7);

const imported = root.importPptx(new Uint8Array(built.pptx));
assert.equal(imported.report.slideCount, 1);
const reEmitted = root.emit(root.resolve(imported.document));
assert.ok(reEmitted.pptx.byteLength > 0);
const generated = root.codegen(imported.document, { target: "vanilla" });
assert.equal(typeof generated.files[generated.entry], "string");

process.stdout.write("ok");
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.equal(runIsolated(conditions, script), "ok");
    }
  }, 45_000);

  it("defers React loading until JSX build and never loads preview, rasterizer, decoders, or validation during JSX build across source and dist", () => {
    const script = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";

const loaded = new Set();
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("JSX build loaded forbidden dependency: " + specifier);
    }
    const resolved = nextResolve(specifier, context);
    if (specifier === "react" || specifier === "react-reconciler") {
      loaded.add(specifier);
    }
    if (specifier === "@pptxdk/react" || resolved.url.endsWith("/react/dist/index.js")) {
      loaded.add("@pptxdk/react");
    }
    return resolved;
  },
});

const { build } = await import("@officedk/pptxdk");
assert.equal(loaded.size, 0, "Importing @officedk/pptxdk eagerly loaded: " + [...loaded].join(", "));

const { createElement } = await import("react");
const { Deck, Image, Shape, Slide } = await import("@officedk/pptxdk/react");
const pngBytes = new Uint8Array(Buffer.from("${png1x1Base64}", "base64"));
const jpegBytes = new Uint8Array(Buffer.from("${jpeg1x1Base64}", "base64"));

const result = await build(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "jsx-s1" },
      createElement(Shape, { id: "jsx-sh1", style: { position: "absolute", left: 10, top: 10, width: 40, height: 40 } }, "JSX"),
      createElement(Image, { style: { position: "absolute", left: 60, top: 10, width: 40, height: 40 }, src: pngBytes, fit: "cover" }),
      createElement(Image, { style: { position: "absolute", left: 110, top: 10, width: 40, height: 40 }, src: jpegBytes }),
    ),
  ),
);
assert.ok(result.pptx instanceof Uint8Array && result.pptx.byteLength > 0);
assert.ok(loaded.has("@pptxdk/react") && loaded.has("react") && loaded.has("react-reconciler"));
process.stdout.write("ok");
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.equal(runIsolated(conditions, script), "ok");
    }
  }, 45_000);

  it("isolates preview entrypoints: plain document and @pptxdk/preview paint avoid React, preview schema/zero-page paint avoids resvg, and JSX preview excludes test/validation", () => {
    const plainPreviewScript = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("Plain preview loaded React or validation: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { deck, slide, shape } = await import("@officedk/pptxdk/vanilla");
const { preview } = await import("@officedk/pptxdk/preview");
const result = await preview(
  deck({ width: 320, height: 180 }).append(
    slide({ id: "s1", background: "#112233" }).append(
      shape({ style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#FF0000" } }),
    ),
  ),
);
assert.ok(result.pages[0]?.png && result.pages[0].png.byteLength > 0);

const { paint } = await import("@pptxdk/preview");
const { createResolver } = await import("@pptxdk/core/build");
const { createDeckNode, createSlideNode } = await import("@pptxdk/core/document");
const { shapeCapability, createShapeNode } = await import("@pptxdk/capabilities/shape");
const d = createDeckNode({ width: 320, height: 180 });
const s = createSlideNode({ id: "s1" });
s.append(createShapeNode({ id: "box", style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#2563EB" } }));
d.append(s);
const painted = paint(createResolver([shapeCapability])({ children: [d] }));
assert.ok(painted.pages[0]?.png && painted.pages[0].png.byteLength > 0);
process.stdout.write("ok");
`;

    const zeroPagePreviewScript = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier.includes("@resvg/resvg-js")) {
      throw new Error("forbidden @resvg/resvg-js load: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const previewEntry = await import("@officedk/pptxdk/preview");
const { resolve } = await import("@officedk/pptxdk");
const { deck, slide } = await import("@officedk/pptxdk/vanilla");
assert.equal(previewEntry.previewPageNumberSchema.safeParse(2).success, true);
const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));
const painted = previewEntry.paint(resolve(doc), { pages: [] });
assert.equal(painted.pages.length, 0);
process.stdout.write("ok");
`;

    const jsxPreviewClosureScript = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";
import { createElement } from "react";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@pptxdk/validation") ||
      specifier === "vitest" ||
      specifier.includes("/capabilities/test/")
    ) {
      throw new Error("Production preview loaded test/validation dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { Deck, Slide, Shape } = await import("@officedk/pptxdk/react");
const { preview } = await import("@officedk/pptxdk/preview");
const result = await preview(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "s1" },
      createElement(Shape, { style: { position: "absolute", left: 10, top: 10, width: 40, height: 40, background: "#2563EB" } }),
    ),
  ),
);
assert.ok(result.pages[0]?.png && result.pages[0].png.byteLength > 0);
process.stdout.write("ok");
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.equal(runIsolated(conditions, plainPreviewScript), "ok");
      assert.equal(runIsolated(conditions, zeroPagePreviewScript), "ok");
      assert.equal(runIsolated(conditions, jsxPreviewClosureScript), "ok");
    }
  }, 45_000);

  it("isolates trimmed capability selection (selectCapabilities) and standalone table+chart paint without loading unselected capabilities", async () => {
    const sample = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "trimmed" },
          createElement(
            Text,
            {
              id: "t1",
              name: "Label",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 200,
                height: 40,
              },
            },
            "Selective Import",
          ),
        ),
      ),
    );
    const sampleBase64 = Buffer.from(sample.pptx).toString("base64");
    const jostReadPath = "./test/assets/fonts/Jost-400-Book.ttf";
    assert.ok(
      readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url))
        .byteLength > 0,
    );

    const trimmedVanillaScript = `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js") ||
      resolved.url.endsWith("/capabilities/src/chart/chart.ts") ||
      resolved.url.endsWith("/capabilities/dist/chart/chart.js") ||
      resolved.url.endsWith("/capabilities/src/table/table.ts") ||
      resolved.url.endsWith("/capabilities/dist/table/table.js")
    ) {
      throw new Error("unselected capability implementation loaded: " + resolved.url);
    }
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/") ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("vanilla trimmed build loaded forbidden dependency: " + specifier);
    }
    return resolved;
  },
});

const { createMorphScope, createMotionScopeHandle, deck, defineTheme, fade, font, linearGradient, motionScope, paragraph, preset, shape, slide, span, step, transitions } = await import("@officedk/pptxdk/vanilla");
const { build } = await import("@officedk/pptxdk/build");

const jostBytes = new Uint8Array(readFileSync(${JSON.stringify(jostReadPath)}));
const family = font("Jost", [{ weight: 400, data: jostBytes }], { substitute: true });
const theme = defineTheme({ name: "TrimmedVanilla", fonts: { major: family, minor: family } });
const morph = createMorphScope();
const hero = morph.handle("hero");
const scope = createMotionScopeHandle();
const doc = deck({ width: 240, height: 120, theme }).append(
  slide({ id: "trimmed-vanilla", transition: transitions.fade() }).append(
    step({ n: 1 }).append(
      motionScope(scope).append(
        shape({
          id: "plain",
          morphKey: hero,
          geometry: preset("roundRect", { adj: 20000 }),
          style: { position: "absolute", left: 0, top: 0, width: 20, height: 20, background: linearGradient(90, "#0F766E", "#14B8A6") },
          enter: fade({ duration: 150 }),
        }),
      ),
    ),
    shape("Shape text", { id: "labeled", name: "labeled", style: { position: "absolute", left: 24, top: 0, width: 100, height: 32 } }),
    shape({ id: "rich", name: "rich", style: { position: "absolute", left: 24, top: 40, width: 120, height: 32 } }).content(paragraph(span("Rich shape"))),
  ),
  slide({ id: "trimmed-vanilla-next", transition: transitions.morph({ fallback: transitions.fade() }) }).append(
    shape({ id: "plain-next", morphKey: hero, geometry: preset("roundRect", { adj: 20000 }), style: { position: "absolute", left: 20, top: 10, width: 40, height: 40, background: "#0F766E" } }),
  ),
);
const result = await build(doc, { capabilities: ["shape", "motion", "transition"] });
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) throw new Error("Expected PPTX bytes");
const texts = result.report.slides[0]?.objects.map((o) => o.name);
process.stdout.write(JSON.stringify(texts));
`;

    const trimmedReactScript = `
import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      resolved.url.endsWith("/capabilities/src/text/text.ts") ||
      resolved.url.endsWith("/capabilities/dist/text/text.js") ||
      resolved.url.endsWith("/capabilities/src/chart/chart.ts") ||
      resolved.url.endsWith("/capabilities/dist/chart/chart.js") ||
      resolved.url.endsWith("/capabilities/src/table/table.ts") ||
      resolved.url.endsWith("/capabilities/dist/table/table.js")
    ) {
      throw new Error("unselected capability implementation loaded: " + resolved.url);
    }
    if (specifier.includes("@pptxdk/preview") || specifier.includes("@resvg/resvg-js")) {
      throw new Error("react trimmed build loaded preview dependency: " + specifier);
    }
    return resolved;
  },
});

const { createElement } = await import("react");
const { createMorphScope, createMotionScopeHandle, Deck, defineTheme, fade, font, linearGradient, MotionScope, P, preset, Shape, Slide, Span, Step, transitions } = await import("@officedk/pptxdk/react");
const { build } = await import("@officedk/pptxdk/build");

const jostBytes = new Uint8Array(readFileSync(${JSON.stringify(jostReadPath)}));
const family = font("Jost", [{ weight: 400, data: jostBytes }], { substitute: true });
const theme = defineTheme({ name: "TrimmedReact", fonts: { major: family, minor: family } });
const morph = createMorphScope();
const hero = morph.handle("hero");
const scope = createMotionScopeHandle();
const tree = createElement(
  Deck,
  { width: 240, height: 120, theme },
  createElement(
    Slide,
    { id: "trimmed-react", transition: transitions.fade() },
    createElement(
      Step,
      { n: 1 },
      createElement(
        MotionScope,
        { scope },
        createElement(Shape, { id: "plain", morphKey: hero, geometry: preset("roundRect", { adj: 20000 }), style: { position: "absolute", left: 0, top: 0, width: 20, height: 20, background: linearGradient(90, "#0F766E", "#14B8A6") }, enter: fade({ duration: 150 }) }),
      ),
    ),
    createElement(Shape, { id: "labeled", name: "labeled", style: { position: "absolute", left: 24, top: 0, width: 100, height: 32 } }, "React shape text"),
    createElement(Shape, { id: "rich", name: "rich", style: { position: "absolute", left: 24, top: 40, width: 120, height: 32 } }, createElement(P, null, createElement(Span, null, "Rich react shape"))),
  ),
  createElement(
    Slide,
    { id: "trimmed-react-next", transition: transitions.morph({ fallback: transitions.fade() }) },
    createElement(Shape, { id: "plain-next", morphKey: hero, geometry: preset("roundRect", { adj: 20000 }), style: { position: "absolute", left: 20, top: 10, width: 40, height: 40, background: "#0F766E" } }),
  ),
);
const result = await build(tree, { capabilities: ["shape", "motion", "transition"] });
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) throw new Error("Expected PPTX bytes");
const texts = result.report.slides[0]?.objects.map((o) => o.name);
process.stdout.write(JSON.stringify(texts));
`;

    const selectiveImportScript = `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    const resolved = nextResolve(specifier, context);
    if (
      specifier === "@pptxdk/capabilities/all" ||
      resolved.url.endsWith("/capabilities/src/all.ts") ||
      resolved.url.endsWith("/capabilities/dist/all.js") ||
      resolved.url.endsWith("/capabilities/src/chart/chart.ts") ||
      resolved.url.endsWith("/capabilities/dist/chart/chart.js") ||
      specifier === "react" ||
      specifier === "@pptxdk/react" ||
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("@pptxdk/validation")
    ) {
      throw new Error("selective createImporter loaded forbidden module: " + resolved.url);
    }
    return resolved;
  },
});

const { createImporter, selectCapabilities, createResolver, emit } = await import("@officedk/pptxdk/build");
const modules = await selectCapabilities(["text"]);
const importer = createImporter(modules);
const bytes = new Uint8Array(Buffer.from("${sampleBase64}", "base64"));
const imported = importer(bytes);
bytes.fill(0);
assert.equal(imported.report.slideCount, 1);
assert.deepEqual(imported.report.diagnostics, []);
const emitted = emit(createResolver(modules)(imported.document));
assert.ok(emitted.pptx instanceof Uint8Array && emitted.pptx.byteLength > 0);
process.stdout.write("ok");
`;

    const tableChartWithoutTextScript = `
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
const { tableCapability, createTableNode, createRowNode, createCellNode } = await import("@pptxdk/capabilities/table");
const { chartCapability, createChartNode } = await import("@pptxdk/capabilities/chart");
const { font } = await import("@pptxdk/capabilities/font");
const { emit } = await import("@pptxdk/emit");
const { paint } = await import("@pptxdk/preview");

const jost = font("Jost*", [{ weight: 400, data: new Uint8Array(readFileSync(${JSON.stringify(jostReadPath)})) }]);
const doc = createDeckNode({ width: 640, height: 360 });
const s1 = createSlideNode({ id: "table-and-chart" });
const tbl = createTableNode({ id: "tbl", columns: [120, 120], style: { position: "absolute", left: 20, top: 20, width: 240, fontFamily: jost, fontSize: 16, color: "#111827" } });
const row1 = createRowNode({ style: { height: 40 } });
row1.append(createCellNode({}).content("Alpha"), createCellNode({}).content("Beta"));
tbl.append(row1);
s1.append(
  tbl,
  createChartNode({
    id: "cht",
    title: "Sales",
    legend: "bottom",
    style: { position: "absolute", left: 280, top: 20, width: 320, height: 240, fontFamily: jost, fontSize: 14 },
    groups: [{ kind: "column", data: { kind: "shared-categories", categories: ["Q1", "Q2"] }, series: [{ name: "Actual", values: [12, 24] }] }],
  }),
);
doc.append(s1);
const resolution = createResolver([tableCapability, chartCapability])({ children: [doc] });
if (emit(resolution).pptx.byteLength === 0) throw new Error("Expected PPTX output");
const page = paint(resolution).pages[0];
if (!page || page.errors.length > 0 || !page.png || page.png.byteLength === 0 || !page.svg?.includes("<path")) {
  throw new Error("Expected clean page paint without textCapability");
}
process.stdout.write("ok");
`;

    for (const conditions of [["--conditions=source"], []] as const) {
      assert.deepEqual(
        JSON.parse(runIsolated(conditions, trimmedVanillaScript)),
        ["!!morph-1", "labeled", "rich"],
      );
      assert.deepEqual(
        JSON.parse(runIsolated(conditions, trimmedReactScript)),
        ["!!morph-1", "labeled", "rich"],
      );
      assert.equal(runIsolated(conditions, selectiveImportScript), "ok");
    }
    assert.equal(
      runIsolated(["--conditions=source"], tableChartWithoutTextScript),
      "ok",
    );
  }, 45_000);
});
