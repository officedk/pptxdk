import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import {
  BuildError,
  buildOptionsSchema,
  diagnosticSchema,
  diagnosticSeveritySchema,
  emit,
  emitOptionsSchema,
  extensionSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  resolve,
  resolveOptionsSchema,
  targetNameSchema,
  targetSchema,
  targets,
} from "@officedk/pptxdk";
import {
  build as buildTrimmed,
  officialCapabilityNameSchema,
  officialCapabilityNames,
  resolve as resolveTrimmed,
  trimmedBuildOptionsSchema,
  trimmedResolveOptionsSchema,
} from "@officedk/pptxdk/build";
import {
  buildFromSource,
  buildFromSourceOptionsSchema,
  exportNameSchema,
  previewFromSource,
  previewFromSourceOptionsSchema,
} from "@officedk/pptxdk/node";
import {
  paint,
  paintOptionsSchema,
  preview,
  previewDiagnosticSchema,
  previewOptionsSchema,
  previewPageNumberSchema,
  previewPagesSchema,
  previewSeveritySchema,
  screenGeometrySchema,
  screenPointSchema,
} from "@officedk/pptxdk/preview";
import { deck, slide } from "@officedk/pptxdk/vanilla";
import {
  buildOptionsSchema as coreBuildOptionsSchema,
  diagnosticSchema as coreDiagnosticSchema,
  diagnosticSeveritySchema as coreDiagnosticSeveritySchema,
  emitOptionsSchema as coreEmitOptionsSchema,
  extensionSchema as coreExtensionSchema,
  fontEmbeddingModeSchema as coreFontEmbeddingModeSchema,
  formatBaselineSchema as coreFormatBaselineSchema,
  resolveOptionsSchema as coreResolveOptionsSchema,
  targetNameSchema as coreTargetNameSchema,
  targetSchema as coreTargetSchema,
} from "@pptxdk/core";
import {
  buildOptionsSchema as buildEntryBuildOptionsSchema,
  resolveOptionsSchema as buildEntryResolveOptionsSchema,
} from "@pptxdk/core/build";
import { validateOptions } from "@pptxdk/emit";
import {
  paintOptionsSchema as pkgPaintOptionsSchema,
  previewDiagnosticSchema as pkgPreviewDiagnosticSchema,
  previewOptionsSchema as pkgPreviewOptionsSchema,
  previewPageNumberSchema as pkgPreviewPageNumberSchema,
  previewPagesSchema as pkgPreviewPagesSchema,
  previewSeveritySchema as pkgPreviewSeveritySchema,
  screenGeometrySchema as pkgScreenGeometrySchema,
  screenPointSchema as pkgScreenPointSchema,
  validatePaintOptions,
} from "@pptxdk/preview";
import { describe, it } from "vitest";

describe("stage option schemas (core + emit)", () => {
  it("exposes shared stage option schemas across @pptxdk/core and @officedk/pptxdk", () => {
    assert.equal(formatBaselineSchema, coreFormatBaselineSchema);
    assert.equal(extensionSchema, coreExtensionSchema);
    assert.equal(fontEmbeddingModeSchema, coreFontEmbeddingModeSchema);
    assert.equal(targetSchema, coreTargetSchema);
    assert.equal(targetNameSchema, coreTargetNameSchema);
    assert.equal(resolveOptionsSchema, coreResolveOptionsSchema);
    assert.equal(resolveOptionsSchema, buildEntryResolveOptionsSchema);
    assert.equal(buildOptionsSchema, coreBuildOptionsSchema);
    assert.equal(buildOptionsSchema, buildEntryBuildOptionsSchema);
    assert.equal(emitOptionsSchema, coreEmitOptionsSchema);
    assert.equal(diagnosticSeveritySchema, coreDiagnosticSeveritySchema);
    assert.equal(diagnosticSchema, coreDiagnosticSchema);

    assert.deepEqual(fontEmbeddingModeSchema.options, ["full", "subset"]);
    assert.deepEqual(targetNameSchema.options, [
      "transitional",
      "baseline-only",
    ]);
    assert.deepEqual(extensionSchema.options, ["p14", "p159", "c15"]);
    assert.equal(formatBaselineSchema.safeParse(formatBaseline).success, true);
    assert.equal(formatBaselineSchema.safeParse("Strict").success, false);
    assert.equal(targetSchema.safeParse(targets.transitional).success, true);
    assert.equal(targetSchema.safeParse(targets.baselineOnly).success, true);
    assert.equal(
      diagnosticSchema.safeParse({
        severity: "warning",
        path: "deck/slide[1]",
        code: "overflow",
        message: "Object exceeds slide bounds",
      }).success,
      true,
    );
  });

  it("validates resolve options via schema with exact BuildError codes and paths", () => {
    const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));

    for (const invalidOptions of [null, [], "bad"]) {
      assert.throws(
        () => resolve(doc, invalidOptions as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.path === "resolve" &&
          error.diagnostics[0]?.code === "invalid-options",
      );
    }

    assert.throws(
      () => resolve(doc, { target: targets.baselineOnly } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "resolve" &&
        error.diagnostics[0]?.code === "unknown-prop",
    );

    assert.throws(
      () => resolve(doc, { signal: "not-a-signal" } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "resolve.signal" &&
        error.diagnostics[0]?.code === "invalid-options",
    );

    const controller = new AbortController();
    const resolution = resolve(doc, { signal: controller.signal });
    assert.equal(resolution.report.slides.length, 1);
  });

  it("validates emit/build options and targets via schema with exact BuildError codes and paths", () => {
    const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));
    const resolution = resolve(doc);

    const cases: readonly {
      readonly input: unknown;
      readonly path: string;
      readonly code: string;
    }[] = [
      { input: null, path: "build", code: "invalid-options" },
      { input: [], path: "build", code: "invalid-options" },
      { input: { extra: true }, path: "build", code: "unknown-parameter" },
      {
        input: { embedFonts: "partial" },
        path: "build.embedFonts",
        code: "unsupported-font-embedding",
      },
      {
        input: { signal: {} },
        path: "build.signal",
        code: "invalid-options",
      },
      {
        input: { target: null },
        path: "build.target",
        code: "invalid-target",
      },
      {
        input: {
          target: { name: "   ", format: formatBaseline, extensions: [] },
        },
        path: "build.target",
        code: "invalid-target",
      },
      {
        input: {
          target: {
            name: "extra",
            format: formatBaseline,
            extensions: [],
            unexpected: true,
          },
        },
        path: "build.target",
        code: "unknown-parameter",
      },
      {
        input: {
          target: { name: "strict", format: "Strict", extensions: [] },
        },
        path: "build.target",
        code: "unsupported-format",
      },
      {
        input: {
          target: {
            name: "bad-ext-type",
            format: formatBaseline,
            extensions: "p14",
          },
        },
        path: "build.target",
        code: "invalid-target",
      },
      {
        input: {
          target: {
            name: "bad-ext",
            format: formatBaseline,
            extensions: ["p99"],
          },
        },
        path: "build.target",
        code: "unsupported-extension",
      },
      {
        input: {
          target: {
            name: "dup-ext",
            format: formatBaseline,
            extensions: ["p14", "p14"],
          },
        },
        path: "build.target",
        code: "duplicate-extension",
      },
    ];

    for (const testCase of cases) {
      assert.throws(
        () => validateOptions(testCase.input as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.path === testCase.path &&
          error.diagnostics[0]?.code === testCase.code,
        `Expected ${testCase.path} [${testCase.code}]`,
      );
      assert.throws(
        () => emit(resolution, testCase.input as never),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics[0]?.path === testCase.path &&
          error.diagnostics[0]?.code === testCase.code,
      );
    }

    const normalizedDefault = validateOptions({});
    assert.equal(normalizedDefault, targets.transitional);

    const normalizedCustom = validateOptions({
      target: {
        name: "custom-p14",
        format: formatBaseline,
        extensions: ["p14"],
      },
      embedFonts: "subset",
    });
    assert.equal(Object.isFrozen(normalizedCustom), true);
    assert.equal(Object.isFrozen(normalizedCustom.extensions), true);
    assert.deepEqual(normalizedCustom, {
      name: "custom-p14",
      format: formatBaseline,
      extensions: ["p14"],
    });
  });
});

describe("preview schemas and lazy resvg loading", () => {
  it("exposes preview schemas across @pptxdk/preview and @officedk/pptxdk/preview", () => {
    assert.equal(previewPageNumberSchema, pkgPreviewPageNumberSchema);
    assert.equal(previewPagesSchema, pkgPreviewPagesSchema);
    assert.equal(paintOptionsSchema, pkgPaintOptionsSchema);
    assert.equal(previewOptionsSchema, pkgPreviewOptionsSchema);
    assert.equal(previewSeveritySchema, pkgPreviewSeveritySchema);
    assert.equal(previewDiagnosticSchema, pkgPreviewDiagnosticSchema);
    assert.equal(screenPointSchema, pkgScreenPointSchema);
    assert.equal(screenGeometrySchema, pkgScreenGeometrySchema);

    assert.equal(previewPageNumberSchema.safeParse(1).success, true);
    assert.equal(previewPageNumberSchema.safeParse(0).success, false);
    assert.equal(previewPageNumberSchema.safeParse(-1).success, false);
    assert.equal(previewPageNumberSchema.safeParse(1.5).success, false);
    assert.equal(
      previewPageNumberSchema.safeParse(Number.MAX_SAFE_INTEGER + 1).success,
      false,
    );
    assert.equal(previewPagesSchema.safeParse([]).success, true);
    assert.equal(previewPagesSchema.safeParse([1, 3]).success, true);
    assert.equal(
      previewDiagnosticSchema.safeParse({
        severity: "error",
        page: 1,
        path: "deck/slide[1]/shape[a]",
        message: "Unsupported",
      }).success,
      true,
    );
    assert.equal(
      screenGeometrySchema.safeParse({
        corners: [
          { x: 0, y: 0 },
          { x: 10, y: 0 },
          { x: 10, y: 10 },
          { x: 0, y: 10 },
        ],
        screenBox: { x: 0, y: 0, width: 10, height: 10 },
      }).success,
      true,
    );
  });

  it("validates paint and preview options including invalid pages, non-object options, and empty pages", async () => {
    const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));
    const resolution = resolve(doc);

    for (const invalidPages of [[0], [-1], [1.5], [Number.NaN], "1"]) {
      assert.throws(
        () => validatePaintOptions({ pages: invalidPages }),
        RangeError,
      );
      assert.throws(
        () => paint(resolution, { pages: invalidPages as never }),
        RangeError,
      );
      await assert.rejects(
        () => preview(doc, { pages: invalidPages as never }),
        RangeError,
      );
    }

    for (const badOptions of [null, [], { unknown: true }, { signal: "bad" }]) {
      assert.throws(() => validatePaintOptions(badOptions), TypeError);
      assert.throws(() => paint(resolution, badOptions as never), TypeError);
      await assert.rejects(() => preview(doc, badOptions as never), TypeError);
    }

    const emptyPaint = paint(resolution, { pages: [] });
    assert.equal(emptyPaint.pages.length, 0);

    const emptyPreview = await preview(doc, { pages: [] });
    assert.equal(emptyPreview.pages.length, 0);
  });

  it("never loads @resvg/resvg-js when importing schemas from @officedk/pptxdk/preview or painting zero pages", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
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

if (previewEntry.previewPageNumberSchema.safeParse(2).success !== true) {
  throw new Error("expected valid page number");
}
const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));
const res = resolve(doc);
const painted = previewEntry.paint(res, { pages: [] });
if (painted.pages.length !== 0) {
  throw new Error("expected 0 painted pages");
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

describe("@officedk/pptxdk/build and @officedk/pptxdk/node schemas", () => {
  it("exposes and validates trimmed build/resolve schemas in @officedk/pptxdk/build", async () => {
    assert.deepEqual(
      officialCapabilityNameSchema.options,
      officialCapabilityNames,
    );
    assert.equal(
      trimmedResolveOptionsSchema.safeParse({ capabilities: ["shape"] })
        .success,
      true,
    );
    assert.equal(
      trimmedResolveOptionsSchema.safeParse({
        capabilities: ["shape"],
        target: targets.baselineOnly,
      }).success,
      false,
    );
    assert.equal(
      trimmedBuildOptionsSchema.safeParse({
        capabilities: ["shape"],
        embedFonts: "subset",
        target: targets.baselineOnly,
      }).success,
      true,
    );

    const doc = deck({ width: 320, height: 180 }).append(slide({ id: "s1" }));

    await assert.rejects(
      () => resolveTrimmed(doc, null as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "resolve" &&
        error.diagnostics[0]?.code === "invalid-options",
    );
    await assert.rejects(
      () =>
        resolveTrimmed(doc, {
          capabilities: ["shape"],
          target: targets.baselineOnly,
        } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "resolve" &&
        error.diagnostics[0]?.code === "unknown-prop",
    );
    await assert.rejects(
      () => resolveTrimmed(doc, { capabilities: "shape" } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "build.capabilities" &&
        error.diagnostics[0]?.code === "invalid-options",
    );
    await assert.rejects(
      () => resolveTrimmed(doc, { capabilities: ["unknown"] } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "build.capabilities[0]" &&
        error.diagnostics[0]?.code === "missing-capability",
    );

    await assert.rejects(
      () => buildTrimmed(doc, null as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "build" &&
        error.diagnostics[0]?.code === "invalid-options",
    );
    await assert.rejects(
      () =>
        buildTrimmed(doc, {
          capabilities: ["shape"],
          extra: true,
        } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "build" &&
        error.diagnostics[0]?.code === "unknown-prop",
    );
    await assert.rejects(
      () =>
        buildTrimmed(doc, {
          capabilities: ["shape"],
          embedFonts: "bad",
        } as never),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "build.embedFonts" &&
        error.diagnostics[0]?.code === "target-violation",
    );
  });

  it("exposes and validates node entry option schemas without loading preview on import", async () => {
    assert.equal(exportNameSchema.safeParse("default").success, true);
    assert.equal(exportNameSchema.safeParse("").success, false);
    assert.equal(
      buildFromSourceOptionsSchema.safeParse({
        exportName: "DeckA",
        embedFonts: "full",
        target: targets.baselineOnly,
      }).success,
      true,
    );
    assert.equal(
      buildFromSourceOptionsSchema.safeParse({ exportName: "" }).success,
      false,
    );
    assert.equal(
      previewFromSourceOptionsSchema.safeParse({
        exportName: "DeckA",
        pages: [1, 2],
      }).success,
      true,
    );
    assert.equal(
      previewFromSourceOptionsSchema.safeParse({ pages: [0] }).success,
      false,
    );

    await assert.rejects(
      () =>
        buildFromSource("/non-existent/deck.ts", {
          exportName: "",
        }),
      (error: unknown) => error instanceof BuildError,
    );
    await assert.rejects(
      () =>
        previewFromSource("/non-existent/deck.ts", {
          exportName: "",
        }),
      TypeError,
    );
    await assert.rejects(
      () =>
        previewFromSource("/non-existent/deck.ts", {
          pages: [0],
        }),
      RangeError,
    );
  });
});
