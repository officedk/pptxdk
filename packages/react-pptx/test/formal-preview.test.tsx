import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  Box,
  BuildError,
  Deck,
  Group,
  Shape,
  Slide,
  Text,
  font,
  preview as previewJsx,
} from "@react-pptx/react-pptx";
import {
  box,
  chart,
  deck,
  group,
  shape,
  slide,
  text,
} from "@react-pptx/react-pptx/document";
import { preview as previewDocument } from "@react-pptx/react-pptx/preview";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runIsolatedScript(script: string) {
  const packageRoot = fileURLToPath(new URL("../", import.meta.url));
  const child = spawnSync(
    process.execPath,
    ["--conditions=source", "--input-type=module", "--eval", script],
    { cwd: packageRoot, encoding: "utf8" },
  );
  assert.equal(child.status, 0, child.stderr);
  return child.stdout;
}

describe("Slice 3: formal static preview public entrypoints and runtime isolation", () => {
  it("supports JSX preview and plain document preview from the same formal pipeline", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const jsxResult = await previewJsx(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "cover", background: "#EEEEEE" },
          createElement(Shape, {
            id: "card",
            x: 20,
            y: 20,
            width: 120,
            height: 60,
            fill: "#336699",
          }),
          createElement(
            Text,
            {
              id: "caption",
              x: 20,
              y: 96,
              width: 160,
              style: { fontFamily: jostFont, fontSize: 18, color: "#111827" },
            },
            "Shared preview",
          ),
        ),
      ),
    );
    const docResult = await previewDocument(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "cover", background: "#EEEEEE" }).append(
          shape({
            id: "card",
            x: 20,
            y: 20,
            width: 120,
            height: 60,
            fill: "#336699",
          }),
          text("Shared preview", {
            id: "caption",
            x: 20,
            y: 96,
            width: 160,
            style: { fontFamily: jostFont, fontSize: 18, color: "#111827" },
          }),
        ),
      ),
    );

    assert.equal(jsxResult.pages.length, 1);
    assert.equal(docResult.pages.length, 1);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.deepEqual(docResult.pages[0]?.diagnostics, []);
    expect(jsxResult.pages[0]?.svg).toContain("<path");
    assert.deepEqual(jsxResult.pages[0]?.png, docResult.pages[0]?.png);
    assert.deepEqual((await validatePptx(jsxResult.build.pptx)).errors, []);
    assert.deepEqual((await validatePptx(docResult.build.pptx)).errors, []);
  });

  it("validates 1-based pages, skips drawing unselected pages, and still validates full-deck build", async () => {
    const validTwoPages = deck({ width: 320, height: 180 }).append(
      slide({ id: "unselected-unmigrated" }).append(
        group({ id: "g-unselected" }).append(
          shape({
            id: "s-unselected",
            x: 10,
            y: 10,
            width: 80,
            height: 40,
            fill: "#334155",
          }),
        ),
      ),
      slide({ id: "selected", background: "#123456" }).append(
        shape({
          id: "s",
          x: 10,
          y: 10,
          width: 80,
          height: 40,
          fill: "#FFFFFF",
        }),
      ),
    );

    const selectedOnly = await previewDocument(validTwoPages, { pages: [2] });
    assert.deepEqual(
      selectedOnly.pages.map((p) => p.id),
      ["selected"],
    );
    assert.deepEqual(selectedOnly.pages[0]?.diagnostics, []);
    assert.ok(selectedOnly.pages[0]?.png !== undefined);
    assert.equal(selectedOnly.build.report.slides.length, 2);

    const noneSelected = await previewDocument(validTwoPages, { pages: [] });
    assert.equal(noneSelected.pages.length, 0);
    assert.equal(noneSelected.build.report.slides.length, 2);

    for (const invalidPage of [0, -1, 1.5]) {
      await assert.rejects(
        () => previewDocument(validTwoPages, { pages: [invalidPage] }),
        RangeError,
      );
      await assert.rejects(
        () =>
          previewJsx(
            createElement(
              Deck,
              { width: 320, height: 180 },
              createElement(Slide, { id: "one" }),
            ),
            { pages: [invalidPage] },
          ),
        RangeError,
      );
    }
    await assert.rejects(
      () => previewDocument(validTwoPages, { pages: [3] }),
      RangeError,
    );

    // Full-deck formal build validation still catches build errors on unselected pages.
    const invalidUnselected = deck({ width: 320, height: 180 }).append(
      slide({ id: "broken-unselected" }).append(
        shape({ id: "bad-x", x: Number.NaN, y: 0, width: 50, height: 50 }),
      ),
      slide({ id: "valid-selected" }),
    );
    await assert.rejects(
      () => previewDocument(invalidUnselected, { pages: [2] }),
      BuildError,
    );
  });

  it("rejects unmigrated object capabilities on selected pages without silently omitting them", async () => {
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "with-group" }).append(
        group({ id: "g1" }).append(
          shape({
            id: "m1",
            x: 10,
            y: 10,
            width: 40,
            height: 40,
            fill: "#FF0000",
          }),
        ),
      ),
      slide({ id: "with-chart" }).append(
        chart({
          id: "c1",
          x: 10,
          y: 10,
          width: 200,
          height: 120,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", values: [1] }],
            },
          ],
        }),
      ),
      slide({ id: "ok-box" }).append(
        box({
          id: "b1",
          style: {
            width: 120,
            height: 60,
            padding: 8,
            backgroundColor: "#E2E8F0",
          },
        }).append(
          shape({
            id: "ok",
            width: 60,
            height: 40,
            fill: "#00AA00",
          }),
        ),
      ),
    );
    const result = await previewDocument(doc);
    expect(result.pages[0]?.png).toBeUndefined();
    expect(result.pages[0]?.diagnostics[0]?.path).toContain("group[g1]");
    expect(result.pages[1]?.png).toBeUndefined();
    expect(result.pages[1]?.diagnostics[0]?.path).toContain("chart[c1]");
    expect(result.pages[2]?.png).toBeDefined();
    expect(result.pages[2]?.diagnostics).toEqual([]);

    const jsxResult = await previewJsx(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "with-group" },
          createElement(
            Group,
            { id: "g1" },
            createElement(Shape, {
              id: "m1",
              x: 10,
              y: 10,
              width: 40,
              height: 40,
              fill: "#FF0000",
            }),
          ),
        ),
        createElement(
          Slide,
          { id: "with-jsx-box" },
          createElement(
            Box,
            { id: "jb1", style: { width: 80, height: 40, padding: 4 } },
            createElement(Shape, { width: 40, height: 20, fill: "#0000FF" }),
          ),
        ),
      ),
    );
    expect(jsxResult.pages[0]?.png).toBeUndefined();
    expect(jsxResult.pages[0]?.diagnostics[0]?.path).toContain("group[g1]");
    expect(jsxResult.pages[1]?.png).toBeDefined();
    expect(jsxResult.pages[1]?.diagnostics).toEqual([]);
  });

  it("never loads @resvg/resvg-js when calling JSX build or document build", () => {
    const jsxOut = runIsolatedScript(`
import { registerHooks } from "node:module";
import { createElement } from "react";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@react-pptx/validation")
    ) {
      throw new Error("JSX build loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build, Deck, Image, Slide, Shape } = await import("@react-pptx/react-pptx");
const pngBytes = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);
const result = await build(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "s1" },
      createElement(Shape, { x: 10, y: 10, width: 40, height: 40 }),
      createElement(Image, { x: 60, y: 10, width: 40, height: 40, src: pngBytes }),
    ),
  ),
);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output");
}
process.stdout.write("ok");
`);
    assert.equal(jsxOut, "ok");

    const docBuildOut = runIsolatedScript(`
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier.includes("@resvg/resvg-js") ||
      specifier.includes("pngjs") ||
      specifier.includes("jpeg-js") ||
      specifier.includes("@react-pptx/validation")
    ) {
      throw new Error("Document build loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build, deck, image, slide, shape } = await import("@react-pptx/react-pptx/document");
const { createFormalBuild } = await import("@react-pptx/react-pptx/build");
void createFormalBuild;
const pngBytes = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);
const result = await build(
  deck({ width: 320, height: 180 }).append(
    slide({ id: "s1" }).append(
      shape({ x: 10, y: 10, width: 40, height: 40 }),
      image({ x: 60, y: 10, width: 40, height: 40, src: pngBytes }),
    ),
  ),
);
if (!(result.pptx instanceof Uint8Array) || result.pptx.byteLength === 0) {
  throw new Error("Expected PPTX output");
}
process.stdout.write("ok");
`);
    assert.equal(docBuildOut, "ok");
  });

  it("never loads react or react-reconciler when calling plain document preview", () => {
    const out = runIsolatedScript(`
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier.includes("@react-pptx/validation")
    ) {
      throw new Error("Plain document preview loaded React or test dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { deck, slide, shape } = await import("@react-pptx/react-pptx/document");
const { preview } = await import("@react-pptx/react-pptx/preview");
const result = await preview(
  deck({ width: 320, height: 180 }).append(
    slide({ id: "s1", background: "#112233" }).append(
      shape({ x: 10, y: 10, width: 40, height: 40, fill: "#FF0000" }),
    ),
  ),
);
if (!result.pages[0]?.png || result.pages[0].png.byteLength === 0) {
  throw new Error("Expected preview PNG");
}
process.stdout.write("ok");
`);
    assert.equal(out, "ok");
  });

  it("never loads test scenarios, vitest, or validation from production preview entrypoints", () => {
    const out = runIsolatedScript(`
import { registerHooks } from "node:module";
import { createElement } from "react";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier.includes("@react-pptx/validation") ||
      specifier === "vitest" ||
      specifier.includes("/capabilities/test/")
    ) {
      throw new Error("Production entrypoint loaded test/validation dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { Deck, Slide, Shape, preview } = await import("@react-pptx/react-pptx");
const result = await preview(
  createElement(
    Deck,
    { width: 320, height: 180 },
    createElement(
      Slide,
      { id: "s1" },
      createElement(Shape, { x: 10, y: 10, width: 40, height: 40, fill: "#2563EB" }),
    ),
  ),
);
if (!result.pages[0]?.png || result.pages[0].png.byteLength === 0) {
  throw new Error("Expected preview PNG");
}
process.stdout.write("ok");
`);
    assert.equal(out, "ok");
  });
});
