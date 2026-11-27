import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { getEventListeners } from "node:events";
import { mkdir, mkdtemp, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { BuildError, type Resolution } from "@officedk/pptxdk";
import * as nodeEntry from "@officedk/pptxdk/node";
import { preview } from "@officedk/pptxdk/preview";
import { Deck, Shape, Slide } from "@officedk/pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const jostFontUrl = new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)
  .href;

async function createSourceWorkspace() {
  const root = await mkdtemp(join(tmpdir(), "pptxdk source "));
  const nodeModules = join(root, "node_modules");
  await mkdir(join(nodeModules, "@officedk"), { recursive: true });
  await symlink(packageRoot, join(nodeModules, "@officedk/pptxdk"), "dir");
  await symlink(
    join(packageRoot, "node_modules/react"),
    join(nodeModules, "react"),
    "dir",
  );
  return root;
}

function runInNode(conditions: readonly string[], script: string, cwd: string) {
  const child = spawnSync(
    process.execPath,
    [...conditions, "--input-type=module", "--eval", script],
    { cwd, encoding: "utf8" },
  );
  assert.equal(child.status, 0, child.stderr || child.stdout);
  return child.stdout;
}

describe("@officedk/pptxdk/node source SDK", () => {
  it("loads TS, TSX, MTS, and JSX entries with local deps, deferred imports, paths/URLs, and exportName parity", async () => {
    const root = await createSourceWorkspace();
    try {
      const spaceDir = join(root, "deck dir");
      await mkdir(spaceDir, { recursive: true });

      await writeFile(
        join(spaceDir, "theme-token.ts"),
        [
          'import { readFileSync } from "node:fs";',
          'import { defineTheme, font } from "@officedk/pptxdk";',
          'export const accentColor = "#2563EB";',
          `const jost = font("Jost", [{ weight: 400, data: readFileSync(new URL(${JSON.stringify(jostFontUrl)})) }], { substitute: true });`,
          'export const deckTheme = defineTheme({ name: "Test", fonts: { major: jost, minor: jost } });',
          "",
        ].join("\n"),
      );
      await writeFile(
        join(spaceDir, "lazy-label.mts"),
        'export const lazyTitle = "Deferred MTS";\n',
      );
      await writeFile(
        join(spaceDir, "shared-slide.jsx"),
        [
          'import { Shape, Slide } from "@officedk/pptxdk/react";',
          "export function SharedSlide({ id, label, color }) {",
          "  return (",
          "    <Slide id={id}>",
          "      <Shape id={`${id}-shape`} style={{ position: 'absolute', left: 20, top: 20, width: 160, height: 48, background: color }}>{label}</Shape>",
          "    </Slide>",
          "  );",
          "}",
          "",
        ].join("\n"),
      );

      const tsxEntry = join(spaceDir, "multi deck.tsx");
      await writeFile(
        tsxEntry,
        [
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          'import { accentColor, deckTheme } from "./theme-token.ts";',
          'import { SharedSlide } from "./shared-slide.jsx";',
          'const { lazyTitle } = await import("./lazy-label.mts");',
          "export function DefaultDeck() {",
          "  return (",
          "    <Deck width={320} height={180} theme={deckTheme}>",
          '      <SharedSlide id="first" label={lazyTitle} color={accentColor} />',
          '      <Slide id="second">',
          '        <Shape id="second-shape" style={{ position: "absolute", left: 20, top: 20, width: 160, height: 48 }}>Second Slide</Shape>',
          "      </Slide>",
          "    </Deck>",
          "  );",
          "}",
          "export const NamedElement = (",
          "  <Deck width={320} height={180} theme={deckTheme}>",
          '    <Slide id="named-slide">',
          '      <Shape id="named-shape" style={{ position: "absolute", left: 16, top: 16, width: 120, height: 40, background: "#10B981" }}>Named TSX</Shape>',
          "    </Slide>",
          "  </Deck>",
          ");",
          "export default DefaultDeck;",
          "",
        ].join("\n"),
      );

      const jsxEntry = join(spaceDir, "plain-jsx.jsx");
      await writeFile(
        jsxEntry,
        [
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          "export default (",
          "  <Deck width={320} height={180}>",
          '    <Slide id="jsx-only">',
          '      <Shape id="jsx-shape" style={{ position: "absolute", left: 16, top: 16, width: 120, height: 40 }}>JSX Entry</Shape>',
          "    </Slide>",
          "  </Deck>",
          ");",
          "",
        ].join("\n"),
      );

      const mtsEntry = join(spaceDir, "entry-mts.mts");
      await writeFile(
        mtsEntry,
        [
          'import { createElement } from "react";',
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          'const { lazyTitle } = await import("./lazy-label.mts");',
          "export default createElement(",
          "  Deck,",
          "  { width: 320, height: 180 },",
          "  createElement(",
          "    Slide,",
          '    { id: "mts-slide" },',
          '    createElement(Shape, { id: "mts-shape", style: { position: "absolute", left: 16, top: 16, width: 120, height: 40 } }, lazyTitle),',
          "  ),",
          ");",
          "",
        ].join("\n"),
      );

      const tsEntry = join(spaceDir, "entry-ts.ts");
      await writeFile(
        tsEntry,
        [
          'import { createElement } from "react";',
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          'import { accentColor } from "./theme-token.ts";',
          "export default createElement(",
          "  Deck,",
          "  { width: 320, height: 180 },",
          "  createElement(",
          "    Slide,",
          '    { id: "ts-slide" },',
          '    createElement(Shape, { id: "ts-shape", style: { position: "absolute", left: 16, top: 16, width: 120, height: 40, background: accentColor } }, "TS Entry"),',
          "  ),",
          ");",
          "",
        ].join("\n"),
      );

      const {
        buildFromSource,
        previewFromSource,
        resolveFromSource,
        resolveFromSourceOptionsSchema,
        SourceLoadError,
      } = nodeEntry as typeof nodeEntry & {
        previewFromSource: (
          entry: string | URL,
          options?: { exportName?: string; pages?: readonly number[] },
        ) => ReturnType<typeof preview>;
        resolveFromSource: (
          entry: string | URL,
          options?: { exportName?: string; signal?: AbortSignal },
        ) => Promise<Resolution>;
      };

      assert.equal(typeof previewFromSource, "function");
      assert.equal(typeof resolveFromSource, "function");
      assert.equal(resolveFromSourceOptionsSchema.safeParse({}).success, true);
      assert.equal(
        resolveFromSourceOptionsSchema.safeParse({ exportName: "" }).success,
        false,
      );

      await assert.rejects(
        buildFromSource(new URL("https://example.com/deck.tsx")),
        (error: unknown) => error instanceof SourceLoadError,
      );
      await assert.rejects(
        previewFromSource(new URL("https://example.com/deck.tsx")),
        (error: unknown) => error instanceof SourceLoadError,
      );
      await assert.rejects(
        resolveFromSource(new URL("https://example.com/deck.tsx")),
        (error: unknown) => error instanceof SourceLoadError,
      );

      const resolvedDefault = await resolveFromSource(tsxEntry);
      assert.equal(resolvedDefault.slides.length, 2);
      assert.equal(
        resolvedDefault.slide("first").get("first-shape").text,
        "Deferred MTS",
      );

      const resolvedFromUrl = await resolveFromSource(pathToFileURL(tsxEntry), {
        exportName: "NamedElement",
      });
      assert.equal(resolvedFromUrl.slides.length, 1);
      assert.equal(
        resolvedFromUrl.slide("named-slide").get("named-shape").text,
        "Named TSX",
      );

      const builtDefault = await buildFromSource(tsxEntry);
      const parsedDefault = await validatePptx(builtDefault.pptx);
      assert.deepEqual(parsedDefault.errors, []);
      assert.equal(parsedDefault.slides.length, 2);
      assert.equal(
        parsedDefault.slides[0]?.objects[0]?.text?.text,
        "Deferred MTS",
      );

      const builtFromUrl = await buildFromSource(pathToFileURL(tsxEntry), {
        exportName: "NamedElement",
      });
      const parsedNamed = await validatePptx(builtFromUrl.pptx);
      assert.deepEqual(parsedNamed.errors, []);
      assert.equal(parsedNamed.slides[0]?.objects[0]?.text?.text, "Named TSX");

      const { readFileSync } = await import("node:fs");
      const { defineTheme, font } = await import("@officedk/pptxdk");
      const jost = font(
        "Jost",
        [{ weight: 400, data: readFileSync(new URL(jostFontUrl)) }],
        { substitute: true },
      );
      const testTheme = defineTheme({
        name: "Test",
        fonts: { major: jost, minor: jost },
      });

      const directPreview = await preview(
        createElement(
          Deck,
          { width: 320, height: 180, theme: testTheme },
          createElement(
            Slide,
            { id: "first" },
            createElement(
              Shape,
              {
                id: "first-shape",
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 160,
                  height: 48,
                  background: "#2563EB",
                },
              },
              "Deferred MTS",
            ),
          ),
          createElement(
            Slide,
            { id: "second" },
            createElement(
              Shape,
              {
                id: "second-shape",
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 160,
                  height: 48,
                },
              },
              "Second Slide",
            ),
          ),
        ),
        { pages: [2] },
      );

      const previewFromPath = await previewFromSource(tsxEntry, {
        pages: [2],
      });
      assert.equal(previewFromPath.report.slides.length, 2);
      assert.equal(previewFromPath.pages.length, 1);
      assert.equal(previewFromPath.pages[0]?.page, 2);
      assert.equal(previewFromPath.pages[0]?.id, "second");
      assert.deepEqual(previewFromPath.pages[0]?.errors, []);
      assert.deepEqual(previewFromPath.pages[0]?.placeholders, []);
      assert.ok((previewFromPath.pages[0]?.png?.byteLength ?? 0) > 0);
      assert.deepEqual(
        previewFromPath.pages[0]?.png,
        directPreview.pages[0]?.png,
      );

      const previewFromUrl = await previewFromSource(pathToFileURL(tsxEntry), {
        exportName: "NamedElement",
      });
      assert.equal(previewFromUrl.report.slides.length, 1);
      assert.equal(previewFromUrl.pages[0]?.id, "named-slide");
      assert.ok((previewFromUrl.pages[0]?.png?.byteLength ?? 0) > 0);

      const jsxResult = await buildFromSource(jsxEntry);
      assert.equal(jsxResult.report.slides[0]?.id, "jsx-only");
      const mtsResult = await buildFromSource(mtsEntry);
      assert.equal(mtsResult.report.slides[0]?.id, "mts-slide");
      const tsResult = await buildFromSource(tsEntry);
      assert.equal(tsResult.report.slides[0]?.id, "ts-slide");
      assert.equal(
        (await resolveFromSource(jsxEntry)).report.slides[0]?.id,
        "jsx-only",
      );
      assert.equal(
        (await resolveFromSource(mtsEntry)).report.slides[0]?.id,
        "mts-slide",
      );
      assert.equal(
        (await resolveFromSource(tsEntry)).report.slides[0]?.id,
        "ts-slide",
      );

      const relativeOutput = runInNode(
        ["--conditions=source"],
        [
          'import assert from "node:assert/strict";',
          'import { buildFromSource, previewFromSource, resolveFromSource } from "@officedk/pptxdk/node";',
          'const b = await buildFromSource("./deck dir/multi deck.tsx");',
          'const p = await previewFromSource("./deck dir/multi deck.tsx", { pages: [1] });',
          'const r = await resolveFromSource("./deck dir/multi deck.tsx");',
          "assert.equal(b.report.slides.length, 2);",
          'assert.equal(p.pages[0]?.id, "first");',
          "assert.equal(r.slides.length, 2);",
          'process.stdout.write("ok");',
        ].join("\n"),
        root,
      );
      assert.equal(relativeOutput, "ok");
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  }, 45_000);

  it("supports all 4 valid export shapes (CommonNode, Promise<CommonNode>, JSX element, Hooks component) with direct SDK parity", async () => {
    const root = await createSourceWorkspace();
    try {
      const shapesEntry = join(root, "four-shapes.tsx");
      await writeFile(
        shapesEntry,
        [
          'import { readFileSync } from "node:fs";',
          'import { useEffect, useMemo, useState } from "react";',
          'import { defineTheme, font } from "@officedk/pptxdk";',
          'import { deck, shape, slide } from "@officedk/pptxdk/vanilla";',
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          "",
          `const jost = font("Jost", [{ weight: 400, data: readFileSync(new URL(${JSON.stringify(jostFontUrl)})) }], { substitute: true });`,
          'const theme = defineTheme({ name: "Test", fonts: { major: jost, minor: jost } });',
          "",
          "function createVanillaDoc(id: string, label: string) {",
          "  return deck({ width: 320, height: 180, theme }).append(",
          "    slide({ id }).append(",
          "      shape(label, { id: `${id}-box`, style: { position: 'absolute', left: 20, top: 20, width: 180, height: 48, background: '#1D4ED8' } }),",
          "    ),",
          "  );",
          "}",
          "",
          'export const syncDoc = createVanillaDoc("sync-doc", "Sync Document");',
          "",
          "export const asyncDoc = (async () => {",
          "  await Promise.resolve();",
          '  return createVanillaDoc("async-doc", "Async Document");',
          "})();",
          "",
          "export const jsxElement = (",
          "  <Deck width={320} height={180} theme={theme}>",
          '    <Slide id="jsx-el">',
          '      <Shape id="jsx-el-box" style={{ position: "absolute", left: 20, top: 20, width: 180, height: 48, background: "#059669" }}>JSX Element</Shape>',
          "    </Slide>",
          "  </Deck>",
          ");",
          "",
          "export function HookComponent() {",
          '  const [label, setLabel] = useState("initial");',
          '  const prefix = useMemo(() => "Hook:", []);',
          "  useEffect(() => {",
          '    setLabel("committed");',
          "  }, []);",
          "  return (",
          "    <Deck width={320} height={180} theme={theme}>",
          '      <Slide id="hook-comp">',
          '        <Shape id="hook-box" style={{ position: "absolute", left: 20, top: 20, width: 180, height: 48 }}>{`${prefix} ${label}`}</Shape>',
          "      </Slide>",
          "    </Deck>",
          "  );",
          "}",
          "",
        ].join("\n"),
      );

      const { buildFromSource, previewFromSource, resolveFromSource } =
        nodeEntry;
      const hookBuilt = await buildFromSource(shapesEntry, {
        exportName: "HookComponent",
      });
      const hookResolved = await resolveFromSource(shapesEntry, {
        exportName: "HookComponent",
      });
      assert.equal(hookResolved.slides.length, 1);
      assert.equal(
        hookResolved.slide("hook-comp").get("hook-box").text,
        "Hook: committed",
      );
      const hookValidated = await validatePptx(hookBuilt.pptx);
      assert.deepEqual(hookValidated.errors, []);
      assert.equal(
        hookValidated.slides[0]?.objects[0]?.text?.text,
        "Hook: committed",
      );

      const parityScript = [
        'import assert from "node:assert/strict";',
        'import { readFileSync } from "node:fs";',
        'import { createElement } from "react";',
        'import { build, defineTheme, font } from "@officedk/pptxdk";',
        'import { preview } from "@officedk/pptxdk/preview";',
        'import { buildFromSource, previewFromSource, resolveFromSource } from "@officedk/pptxdk/node";',
        'import { deck, shape, slide } from "@officedk/pptxdk/vanilla";',
        'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
        "",
        `const jost = font("Jost", [{ weight: 400, data: readFileSync(new URL(${JSON.stringify(jostFontUrl)})) }], { substitute: true });`,
        'const theme = defineTheme({ name: "Test", fonts: { major: jost, minor: jost } });',
        "const checks = [",
        "  {",
        '    exportName: "syncDoc",',
        '    direct: deck({ width: 320, height: 180, theme }).append(slide({ id: "sync-doc" }).append(shape("Sync Document", { id: "sync-doc-box", style: { position: "absolute", left: 20, top: 20, width: 180, height: 48, background: "#1D4ED8" } }))),',
        "  },",
        "  {",
        '    exportName: "asyncDoc",',
        '    direct: deck({ width: 320, height: 180, theme }).append(slide({ id: "async-doc" }).append(shape("Async Document", { id: "async-doc-box", style: { position: "absolute", left: 20, top: 20, width: 180, height: 48, background: "#1D4ED8" } }))),',
        "  },",
        "  {",
        '    exportName: "jsxElement",',
        '    direct: createElement(Deck, { width: 320, height: 180, theme }, createElement(Slide, { id: "jsx-el" }, createElement(Shape, { id: "jsx-el-box", style: { position: "absolute", left: 20, top: 20, width: 180, height: 48, background: "#059669" } }, "JSX Element"))),',
        "  },",
        "];",
        "for (const check of checks) {",
        '  const sourceBuilt = await buildFromSource("./four-shapes.tsx", { exportName: check.exportName });',
        "  const directBuilt = await build(check.direct);",
        "  assert.deepEqual(sourceBuilt.report, directBuilt.report);",
        "  assert.equal(sourceBuilt.pptx.byteLength > 0, true);",
        '  const sourcePreview = await previewFromSource("./four-shapes.tsx", { exportName: check.exportName });',
        "  const directPreview = await preview(check.direct);",
        "  assert.deepEqual(sourcePreview.report, directPreview.report);",
        "  assert.deepEqual(sourcePreview.pages[0]?.errors, []);",
        "  assert.deepEqual(sourcePreview.pages[0]?.placeholders, []);",
        "  assert.equal((sourcePreview.pages[0]?.png?.byteLength ?? 0) > 0, true);",
        "  assert.deepEqual(sourcePreview.pages[0]?.png, directPreview.pages[0]?.png);",
        '  const sourceResolved = await resolveFromSource("./four-shapes.tsx", { exportName: check.exportName });',
        "  assert.deepEqual(sourceResolved.report, directPreview.report);",
        "  assert.equal(sourceResolved.slides.length, 1);",
        "}",
        'process.stdout.write("ok");',
      ].join("\n");
      for (const conditions of [["--conditions=source"], []] as const) {
        assert.equal(runInNode(conditions, parityScript, root), "ok");
      }

      const hookPreview = await previewFromSource(shapesEntry, {
        exportName: "HookComponent",
      });
      assert.equal(hookPreview.pages[0]?.id, "hook-comp");
      assert.ok((hookPreview.pages[0]?.png?.byteLength ?? 0) > 0);
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  }, 45_000);

  it("reports source diagnostics, preserves BuildError/RangeError, and releases loader hooks on errors", async () => {
    const root = await createSourceWorkspace();
    try {
      const badExports = join(root, "bad-exports.tsx");
      await writeFile(
        badExports,
        [
          'import { Deck, Slide } from "@officedk/pptxdk/react";',
          "export const undefExport = undefined;",
          "export const nullExport = null;",
          "export const numExport = 42;",
          'export const strExport = "deck";',
          "export const objExport = { a: 1 };",
          "export const arrExport = [];",
          "export const promiseNum = Promise.resolve(123);",
          'export const promiseJsx = Promise.resolve(<Deck><Slide id="p" /></Deck>);',
          "",
        ].join("\n"),
      );

      const promiseRejected = join(root, "promise-rejected.ts");
      await writeFile(
        promiseRejected,
        'export default Promise.reject(new Error("async reject"));\n',
      );

      const syntaxErr = join(root, "syntax-error.ts");
      await writeFile(syntaxErr, "export const broken = (\n");

      const runtimeErr = join(root, "runtime-error.ts");
      await writeFile(runtimeErr, 'throw new Error("top-level boom");\n');

      const invalidDeck = join(root, "invalid-deck.tsx");
      await writeFile(
        invalidDeck,
        [
          'import { Deck, Slide } from "@officedk/pptxdk/react";',
          'export default <Deck width={320} height={180}><Slide id="outer"><Slide id="inner" /></Slide></Deck>;',
          "",
        ].join("\n"),
      );

      const probeUnloaded = join(root, "probe-unloaded.tsx");
      await writeFile(probeUnloaded, "export const x = <div />;\n");

      const {
        buildFromSource,
        previewFromSource,
        resolveFromSource,
        SourceLoadError,
      } = nodeEntry;

      for (const exportName of [
        "missingExport",
        "undefExport",
        "nullExport",
        "numExport",
        "strExport",
        "objExport",
        "arrExport",
        "promiseNum",
        "promiseJsx",
      ]) {
        await assert.rejects(
          buildFromSource(badExports, { exportName }),
          (error: unknown) => error instanceof SourceLoadError,
          `Expected SourceLoadError for export ${exportName}`,
        );
        await assert.rejects(
          previewFromSource(badExports, { exportName }),
          (error: unknown) => error instanceof SourceLoadError,
          `Expected SourceLoadError in preview for export ${exportName}`,
        );
        await assert.rejects(
          resolveFromSource(badExports, { exportName }),
          (error: unknown) => error instanceof SourceLoadError,
          `Expected SourceLoadError in resolve for export ${exportName}`,
        );
      }

      await assert.rejects(
        buildFromSource(promiseRejected),
        (error: unknown) =>
          error instanceof SourceLoadError && error.cause instanceof Error,
      );
      await assert.rejects(
        previewFromSource(promiseRejected),
        (error: unknown) =>
          error instanceof SourceLoadError && error.cause instanceof Error,
      );
      await assert.rejects(
        resolveFromSource(promiseRejected),
        (error: unknown) =>
          error instanceof SourceLoadError && error.cause instanceof Error,
      );

      const hookReleaseOutput = runInNode(
        ["--conditions=source"],
        [
          'import assert from "node:assert/strict";',
          'import { pathToFileURL } from "node:url";',
          'import { buildFromSource, previewFromSource, resolveFromSource, SourceLoadError } from "@officedk/pptxdk/node";',
          `const badFiles = ${JSON.stringify([join(root, "does-not-exist.ts"), syntaxErr, runtimeErr])};`,
          `const probeBase = ${JSON.stringify(probeUnloaded)};`,
          "for (let i = 0; i < badFiles.length; i += 1) {",
          "  await assert.rejects(",
          "    buildFromSource(badFiles[i]),",
          "    (error) => error instanceof SourceLoadError && error.cause !== undefined,",
          "  );",
          "  await assert.rejects(",
          "    previewFromSource(badFiles[i]),",
          "    (error) => error instanceof SourceLoadError && error.cause !== undefined,",
          "  );",
          "  await assert.rejects(",
          "    resolveFromSource(badFiles[i]),",
          "    (error) => error instanceof SourceLoadError && error.cause !== undefined,",
          "  );",
          "  await assert.rejects(",
          "    import(`${pathToFileURL(probeBase).href}?attempt=${i}`),",
          '    (error) => error?.code === "ERR_UNKNOWN_FILE_EXTENSION",',
          "  );",
          "}",
          'process.stdout.write("ok");',
        ].join("\n"),
        root,
      );
      assert.equal(hookReleaseOutput, "ok");

      await assert.rejects(
        buildFromSource(invalidDeck),
        (error: unknown) => error instanceof BuildError,
      );
      await assert.rejects(
        buildFromSource(invalidDeck, { target: "invalid" as never }),
        (error: unknown) => error instanceof BuildError,
      );
      await assert.rejects(
        resolveFromSource(invalidDeck),
        (error: unknown) => error instanceof BuildError,
      );
      await assert.rejects(
        resolveFromSource(invalidDeck, null as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-options" &&
          error.diagnostics[0]?.path === "resolve",
      );
      await assert.rejects(
        resolveFromSource(invalidDeck, { target: "invalid" } as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "unknown-prop" &&
          error.diagnostics[0]?.path === "resolve" &&
          error.diagnostics[0]?.prop === "target",
      );
      await assert.rejects(
        resolveFromSource(invalidDeck, { signal: "bad" } as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-options" &&
          error.diagnostics[0]?.path === "resolve.signal" &&
          error.diagnostics[0]?.field === "signal",
      );
      await assert.rejects(
        resolveFromSource(invalidDeck, { exportName: "" }),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.code === "invalid-options" &&
          error.diagnostics[0]?.path === "resolve.exportName" &&
          error.diagnostics[0]?.field === "exportName",
      );
      await assert.rejects(
        previewFromSource(invalidDeck, { pages: [0] }),
        (error: unknown) => error instanceof RangeError,
      );
      await assert.rejects(
        previewFromSource(invalidDeck, { pages: [1.5] }),
        (error: unknown) => error instanceof RangeError,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  }, 45_000);

  it("propagates cancellation reasons without unhandled rejections or leftover listeners across all stages", async () => {
    const root = await createSourceWorkspace();
    try {
      const { buildFromSource, previewFromSource, resolveFromSource } =
        nodeEntry;
      const preEntry = join(root, "pre-cancel.tsx");
      await writeFile(
        preEntry,
        [
          'import { Deck, Slide } from "@officedk/pptxdk/react";',
          "globalThis.__pptxdkPreCancelExecuted = true;",
          'export default <Deck width={320} height={180}><Slide id="pre" /></Deck>;',
          "",
        ].join("\n"),
      );

      for (const invoke of [
        buildFromSource,
        previewFromSource,
        resolveFromSource,
      ]) {
        const preController = new AbortController();
        const preReason = new Error("pre-cancelled");
        preController.abort(preReason);
        Reflect.deleteProperty(globalThis, "__pptxdkPreCancelExecuted");

        await assert.rejects(
          invoke(preEntry, { signal: preController.signal }),
          (error: unknown) => error === preReason,
        );
        assert.equal(
          Reflect.get(globalThis, "__pptxdkPreCancelExecuted"),
          undefined,
        );
        assert.equal(
          getEventListeners(preController.signal, "abort").length,
          0,
        );
      }

      const unhandled: unknown[] = [];
      const onUnhandled = (reason: unknown) => {
        unhandled.push(reason);
      };
      process.on("unhandledRejection", onUnhandled);
      try {
        for (const [index, invoke] of [
          buildFromSource,
          previewFromSource,
          resolveFromSource,
        ].entries()) {
          const importPending = join(root, `import-pending-${index}.tsx`);
          const importStarted = Promise.withResolvers<void>();
          const importReject = Promise.withResolvers<void>();
          Reflect.set(globalThis, "__pptxdkImportState", {
            start: () => importStarted.resolve(),
            wait: importReject.promise,
          });
          await writeFile(
            importPending,
            [
              'import { Deck, Slide } from "@officedk/pptxdk/react";',
              "const state = globalThis.__pptxdkImportState;",
              "state.start();",
              "await state.wait;",
              'export default <Deck width={320} height={180}><Slide id="late" /></Deck>;',
              "",
            ].join("\n"),
          );

          const controller = new AbortController();
          const reason = new Error(`aborted during import ${index}`);
          const pendingCall = invoke(importPending, {
            signal: controller.signal,
          });
          await importStarted.promise;
          controller.abort(reason);

          await assert.rejects(
            pendingCall,
            (error: unknown) => error === reason,
          );
          importReject.reject(new Error("late module error"));
          await new Promise((resolve) => setImmediate(resolve));
          assert.equal(getEventListeners(controller.signal, "abort").length, 0);
        }

        for (const [index, invoke] of [
          buildFromSource,
          previewFromSource,
          resolveFromSource,
        ].entries()) {
          const promisePending = join(root, `promise-pending-${index}.ts`);
          const exported = Promise.withResolvers<unknown>();
          const subscribed = Promise.withResolvers<void>();
          Reflect.set(globalThis, "__pptxdkPendingPromise", {
            // Expose subscription timing deterministically before aborting.
            // oxlint-disable-next-line unicorn/no-thenable
            then(
              onFulfilled?: ((value: unknown) => unknown) | null,
              onRejected?: ((reason: unknown) => unknown) | null,
            ) {
              subscribed.resolve();
              return exported.promise.then(onFulfilled, onRejected);
            },
          });
          await writeFile(
            promisePending,
            "export default globalThis.__pptxdkPendingPromise;\n",
          );

          const controller = new AbortController();
          const reason = new Error(`aborted waiting for deck promise ${index}`);
          const pendingCall = invoke(promisePending, {
            signal: controller.signal,
          });
          await subscribed.promise;
          controller.abort(reason);

          await assert.rejects(
            pendingCall,
            (error: unknown) => error === reason,
          );
          exported.reject(new Error("late deck promise rejection"));
          await new Promise((resolve) => setImmediate(resolve));
          assert.equal(getEventListeners(controller.signal, "abort").length, 0);
        }

        for (const [index, invoke] of [
          buildFromSource,
          previewFromSource,
          resolveFromSource,
        ].entries()) {
          const suspensePending = join(root, `suspense-pending-${index}.tsx`);
          const mounted = Promise.withResolvers<void>();
          let cleaned = 0;
          Reflect.set(globalThis, "__pptxdkSuspenseState", {
            onMount: () => mounted.resolve(),
            onCleanup: () => {
              cleaned += 1;
            },
          });
          await writeFile(
            suspensePending,
            [
              'import { Suspense, use, useEffect } from "react";',
              'import { Deck, Slide, Text } from "@officedk/pptxdk/react";',
              "const never = new Promise<string>(() => {});",
              "function PendingChild() {",
              "  useEffect(() => {",
              "    globalThis.__pptxdkSuspenseState.onMount();",
              "    return () => globalThis.__pptxdkSuspenseState.onCleanup();",
              "  }, []);",
              "  const value = use(never);",
              "  return <Text style={{ position: 'absolute', left: 0, top: 0, width: 100, height: 30 }}>{value}</Text>;",
              "}",
              "function MountedProbe() {",
              "  useEffect(() => {",
              "    globalThis.__pptxdkSuspenseState.onMount();",
              "    return () => globalThis.__pptxdkSuspenseState.onCleanup();",
              "  }, []);",
              "  return null;",
              "}",
              "export default function SuspenseDeck() {",
              "  return (",
              "    <Deck width={320} height={180}>",
              '      <Slide id="s1">',
              "        <MountedProbe />",
              "        <Suspense fallback={null}>",
              "          <PendingChild />",
              "        </Suspense>",
              "      </Slide>",
              "    </Deck>",
              "  );",
              "}",
              "",
            ].join("\n"),
          );

          const controller = new AbortController();
          const reason = new Error(`aborted during react expansion ${index}`);
          const pendingCall = invoke(suspensePending, {
            signal: controller.signal,
          });
          await mounted.promise;
          controller.abort(reason);

          await assert.rejects(
            pendingCall,
            (error: unknown) => error === reason,
          );
          assert.equal(cleaned, 1);
          assert.equal(getEventListeners(controller.signal, "abort").length, 0);
        }

        assert.deepEqual(unhandled, []);
      } finally {
        process.removeListener("unhandledRejection", onUnhandled);
        Reflect.deleteProperty(globalThis, "__pptxdkPreCancelExecuted");
        Reflect.deleteProperty(globalThis, "__pptxdkImportState");
        Reflect.deleteProperty(globalThis, "__pptxdkPendingPromise");
        Reflect.deleteProperty(globalThis, "__pptxdkSuspenseState");
      }
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  it("loads React and preview strictly on demand across source and dist conditions", async () => {
    const root = await createSourceWorkspace();
    try {
      const plainEntry = join(root, "plain-vanilla.ts");
      await writeFile(
        plainEntry,
        [
          'import { readFileSync } from "node:fs";',
          'import { defineTheme, font } from "@officedk/pptxdk";',
          'import { deck, shape, slide } from "@officedk/pptxdk/vanilla";',
          `const jost = font("Jost", [{ weight: 400, data: readFileSync(new URL(${JSON.stringify(jostFontUrl)})) }], { substitute: true });`,
          'const theme = defineTheme({ name: "Test", fonts: { major: jost, minor: jost } });',
          "function makeDeck(id: string, text: string) {",
          "  return deck({ width: 320, height: 180, theme }).append(",
          "    slide({ id }).append(",
          '      shape(text, { id: `${id}-box`, style: { position: "absolute", left: 20, top: 20, width: 160, height: 48, background: "#2563EB" } }),',
          "    ),",
          "  );",
          "}",
          'export default makeDeck("sync-plain", "Sync Plain");',
          "export const asyncPlain = (async () => {",
          "  await Promise.resolve();",
          '  return makeDeck("async-plain", "Async Plain");',
          "})();",
          "",
        ].join("\n"),
      );

      const jsxEntry = join(root, "jsx-source.tsx");
      await writeFile(
        jsxEntry,
        [
          'import { Deck, Shape, Slide } from "@officedk/pptxdk/react";',
          "export default function JsxSourceDeck() {",
          "  return (",
          "    <Deck width={320} height={180}>",
          '      <Slide id="jsx-src">',
          '        <Shape id="jsx-src-box" style={{ position: "absolute", left: 20, top: 20, width: 160, height: 48 }}>JSX Source</Shape>',
          "      </Slide>",
          "    </Deck>",
          "  );",
          "}",
          "",
        ].join("\n"),
      );

      const forbidReactAndPreview = [
        'specifier === "react"',
        'specifier === "react-reconciler"',
        'specifier.startsWith("react/")',
        'specifier === "@pptxdk/react"',
        'specifier.startsWith("@pptxdk/react/")',
        'specifier.includes("@resvg/resvg-js")',
        'specifier.includes("@pptxdk/preview")',
        'specifier.endsWith("/preview.ts")',
        'specifier.endsWith("/preview.js")',
        'specifier === "./preview.ts"',
        'specifier === "./preview.js"',
        'specifier.includes("@pptxdk/validation")',
      ].join(" || ");

      const forbidReactOnly = [
        'specifier === "react"',
        'specifier === "react-reconciler"',
        'specifier.startsWith("react/")',
        'specifier === "@pptxdk/react"',
        'specifier.startsWith("@pptxdk/react/")',
        'specifier.includes("@pptxdk/validation")',
      ].join(" || ");

      const forbidPreviewOnly = [
        'specifier.includes("@resvg/resvg-js")',
        'specifier.includes("@pptxdk/preview")',
        'specifier.endsWith("/preview.ts")',
        'specifier.endsWith("/preview.js")',
        'specifier === "./preview.ts"',
        'specifier === "./preview.js"',
        'specifier.includes("@pptxdk/validation")',
      ].join(" || ");

      const plainBuildScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidReactAndPreview}) {`,
        '      throw new Error("Plain build loaded forbidden dependency: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { buildFromSource } = await import("@officedk/pptxdk/node");',
        'const syncBuilt = await buildFromSource("./plain-vanilla.ts");',
        'const asyncBuilt = await buildFromSource("./plain-vanilla.ts", { exportName: "asyncPlain" });',
        "assert.ok(syncBuilt.pptx.byteLength > 0);",
        "assert.ok(asyncBuilt.pptx.byteLength > 0);",
        'process.stdout.write("ok");',
      ].join("\n");

      const plainPreviewScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidReactOnly}) {`,
        '      throw new Error("Plain preview loaded React or validation: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { previewFromSource } = await import("@officedk/pptxdk/node");',
        'const syncPreview = await previewFromSource("./plain-vanilla.ts");',
        'const asyncPreview = await previewFromSource("./plain-vanilla.ts", { exportName: "asyncPlain" });',
        "assert.deepEqual(syncPreview.pages[0]?.errors, []);",
        "assert.deepEqual(syncPreview.pages[0]?.placeholders, []);",
        "assert.deepEqual(asyncPreview.pages[0]?.errors, []);",
        "assert.deepEqual(asyncPreview.pages[0]?.placeholders, []);",
        "assert.ok((syncPreview.pages[0]?.png?.byteLength ?? 0) > 0);",
        "assert.ok((asyncPreview.pages[0]?.png?.byteLength ?? 0) > 0);",
        'process.stdout.write("ok");',
      ].join("\n");

      const jsxBuildScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidPreviewOnly}) {`,
        '      throw new Error("JSX source build loaded preview or validation: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { buildFromSource } = await import("@officedk/pptxdk/node");',
        'const built = await buildFromSource("./jsx-source.tsx");',
        "assert.ok(built.pptx.byteLength > 0);",
        'process.stdout.write("ok");',
      ].join("\n");

      const plainResolveScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidReactAndPreview}) {`,
        '      throw new Error("Plain resolve loaded forbidden dependency: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { resolveFromSource } = await import("@officedk/pptxdk/node");',
        'const syncResolved = await resolveFromSource("./plain-vanilla.ts");',
        'const asyncResolved = await resolveFromSource("./plain-vanilla.ts", { exportName: "asyncPlain" });',
        "assert.equal(syncResolved.slides.length, 1);",
        "assert.equal(asyncResolved.slides.length, 1);",
        'assert.equal(syncResolved.report.slides[0]?.id, "sync-plain");',
        'assert.equal(asyncResolved.report.slides[0]?.id, "async-plain");',
        'process.stdout.write("ok");',
      ].join("\n");

      const jsxResolveScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidPreviewOnly}) {`,
        '      throw new Error("JSX source resolve loaded preview or validation: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { resolveFromSource } = await import("@officedk/pptxdk/node");',
        'const resolved = await resolveFromSource("./jsx-source.tsx");',
        "assert.equal(resolved.slides.length, 1);",
        'assert.equal(resolved.report.slides[0]?.id, "jsx-src");',
        'process.stdout.write("ok");',
      ].join("\n");

      const preCancelIsolationScript = [
        'import assert from "node:assert/strict";',
        'import { registerHooks } from "node:module";',
        "registerHooks({",
        "  resolve(specifier, context, nextResolve) {",
        `    if (${forbidReactAndPreview}) {`,
        '      throw new Error("Pre-cancelled call loaded forbidden dependency: " + specifier);',
        "    }",
        "    return nextResolve(specifier, context);",
        "  },",
        "});",
        'const { buildFromSource, previewFromSource, resolveFromSource } = await import("@officedk/pptxdk/node");',
        "const controller = new AbortController();",
        'const reason = new Error("pre-cancelled isolation");',
        "controller.abort(reason);",
        'await assert.rejects(buildFromSource("./jsx-source.tsx", { signal: controller.signal }), (e) => e === reason);',
        'await assert.rejects(previewFromSource("./jsx-source.tsx", { signal: controller.signal }), (e) => e === reason);',
        'await assert.rejects(resolveFromSource("./jsx-source.tsx", { signal: controller.signal }), (e) => e === reason);',
        'process.stdout.write("ok");',
      ].join("\n");

      for (const conditions of [["--conditions=source"], []] as const) {
        assert.equal(runInNode(conditions, plainBuildScript, root), "ok");
        assert.equal(runInNode(conditions, plainPreviewScript, root), "ok");
        assert.equal(runInNode(conditions, plainResolveScript, root), "ok");
        assert.equal(runInNode(conditions, jsxBuildScript, root), "ok");
        assert.equal(runInNode(conditions, jsxResolveScript, root), "ok");
        assert.equal(
          runInNode(conditions, preCancelIsolationScript, root),
          "ok",
        );
      }
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  }, 45_000);
});
