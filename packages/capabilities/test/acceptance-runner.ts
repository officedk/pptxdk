import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  defaultTheme,
  isTheme,
  type BuildReport,
  type CommonNode,
  type EmitOptions,
  type Font,
  type Theme,
} from "@pptxdk/core";
import { createResolver, type Resolution } from "@pptxdk/core/build";
import {
  CommonNode as DocumentCommonNode,
  DeckContainer,
  SlideMasterContainer,
} from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { createImporter, type ImportReport } from "@pptxdk/import";
import {
  paint,
  type PreviewPageResult,
  type PreviewResult,
} from "@pptxdk/preview";
import { expandDocument } from "@pptxdk/react";
import {
  validatePptx,
  type FileObject,
  type ValidationResult,
} from "@pptxdk/validation";
import { PNG } from "pngjs";
import type { ReactNode } from "react";

const emuPerPx = 9525;

export type AuthorKind = "jsx" | "document";

export type ScenarioAssertContext = {
  readonly author: AuthorKind;
  readonly resolution: Resolution;
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
};

export type NativeRoundtripAssertContext = {
  readonly stage: "initial" | "roundtrip";
  readonly pptx: Uint8Array;
  readonly validation: ValidationResult;
  readonly report: BuildReport;
};

export type RoundtripRunOutput = {
  readonly importReport: ImportReport;
  readonly importedDocument: CommonNode;
  readonly resolution: Resolution;
  readonly report: BuildReport;
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
};

export type CapabilityScenario = {
  readonly id: string;
  readonly capability: string;
  readonly description: string;
  readonly kind?: "positive" | "rejection";
  readonly visualCheckpoints: readonly string[];
  readonly nativeCheckpoints: readonly string[];
  readonly pages?: readonly number[];
  readonly emitOptions?: EmitOptions;
  jsx(): ReactNode;
  document(): CommonNode;
  assert(context: ScenarioAssertContext): void | Promise<void>;
  assertNative?(context: NativeRoundtripAssertContext): void | Promise<void>;
};

export type CollectedScenarioArtifact = {
  readonly scenarioId: string;
  readonly author: AuthorKind;
  readonly page: number;
  readonly slideId: string;
  readonly pngPath?: string;
  readonly pptxPath: string;
};

export type ScenarioRunOutput = {
  readonly scenarioId: string;
  readonly author: AuthorKind;
  readonly resolution: Resolution;
  readonly report: BuildReport;
  readonly pptx: Uint8Array;
  readonly preview: PreviewResult;
  readonly validation: ValidationResult;
  readonly artifacts: readonly CollectedScenarioArtifact[];
};

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

function isFontRecord(value: unknown): value is Font {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return false;
  }
  const record = value as Readonly<Record<string, unknown>>;
  if (typeof record["family"] !== "string" || record["family"].trim() === "") {
    return false;
  }
  const faces = record["faces"];
  return faces === undefined || Array.isArray(faces);
}

function isPlainRecord(
  value: unknown,
): value is Readonly<Record<string, unknown>> {
  if (typeof value !== "object" || value === null) return false;
  const proto = Object.getPrototypeOf(value) as unknown;
  return proto === Object.prototype || proto === null;
}

function collectFontsFromDocument(root: CommonNode): ReadonlyMap<string, Font> {
  const fonts = new Map<string, Font>();
  const visited = new Set<unknown>();

  const recordFont = (candidate: Font): void => {
    const key = candidate.family.trim().toLowerCase();
    const existing = fonts.get(key);
    const candidateCount = candidate.faces?.length ?? 0;
    const existingCount = existing?.faces?.length ?? 0;
    if (existing === undefined || candidateCount > existingCount) {
      fonts.set(key, candidate);
    }
  };

  const visit = (value: unknown): void => {
    if (value === null || typeof value !== "object") return;
    if (visited.has(value)) return;
    visited.add(value);

    if (value instanceof Uint8Array || value instanceof URL) return;

    if (isFontRecord(value)) {
      recordFont(value);
      return;
    }

    if (isTheme(value)) {
      for (const fontToken of Object.values(value.fonts)) {
        visit(fontToken);
      }
      return;
    }

    if (value instanceof DocumentCommonNode) {
      visit(value.props);
      for (const child of value.children) {
        visit(child);
      }
      if (value instanceof DeckContainer) {
        for (const master of value.masters) {
          visit(master);
        }
      } else if (value instanceof SlideMasterContainer) {
        for (const layout of value.layouts) {
          visit(layout);
        }
      }
      return;
    }

    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }

    if (isPlainRecord(value)) {
      for (const entry of Object.values(value)) {
        visit(entry);
      }
    }
  };

  visit(root);
  return fonts;
}

function inferNonBoldFaceWeight(
  matched: Font,
  styleRecord: Readonly<Record<string, unknown>>,
): number | undefined {
  if (matched.substitute === true) return undefined;
  const currentWeight = styleRecord["fontWeight"];
  if (
    currentWeight !== undefined &&
    currentWeight !== "normal" &&
    currentWeight !== 400
  ) {
    return undefined;
  }
  const faces = matched.faces ?? [];
  if (
    faces.length === 0 ||
    faces.some(
      (face) =>
        (face.weight ?? 400) === 400 && (face.style ?? "normal") === "normal",
    )
  ) {
    return undefined;
  }
  const nonBoldFace = faces.find(
    (face) =>
      (face.weight ?? 400) < 600 && (face.style ?? "normal") === "normal",
  );
  return nonBoldFace?.weight;
}

function reattachUnembeddedFonts(
  root: CommonNode,
  fontsByFamily: ReadonlyMap<string, Font>,
  inferNonBoldCustomWeight = false,
): void {
  if (fontsByFamily.size === 0) return;
  const visited = new Set<unknown>();

  const lookupFont = (family: string): Font | undefined =>
    fontsByFamily.get(family.trim().toLowerCase());

  const rewriteTheme = (theme: Theme): void => {
    if (theme === defaultTheme) return;
    const patch: Record<string, Font> = {};
    for (const [token, val] of Object.entries(theme.fonts)) {
      if (typeof val === "string") {
        const matched = lookupFont(val);
        if (matched !== undefined) patch[token] = matched;
      }
    }
    if (Object.keys(patch).length > 0) {
      theme.set({ fonts: patch });
    }
  };

  const rewriteValue = (
    value: unknown,
  ): { readonly changed: boolean; readonly value: unknown } => {
    if (value === null || typeof value !== "object") {
      return { changed: false, value };
    }
    if (
      value instanceof Uint8Array ||
      value instanceof URL ||
      value instanceof DocumentCommonNode ||
      isFontRecord(value)
    ) {
      return { changed: false, value };
    }
    if (isTheme(value)) {
      rewriteTheme(value);
      return { changed: false, value };
    }
    if (Array.isArray(value)) {
      let changed = false;
      const next = value.map((item) => {
        const res = rewriteValue(item);
        if (res.changed) changed = true;
        return res.value;
      });
      return { changed, value: changed ? next : value };
    }
    if (isPlainRecord(value)) {
      let changed = false;
      const next: Record<string, unknown> = {};
      let inferredWeight: number | undefined;
      for (const [key, entry] of Object.entries(value)) {
        if (key === "fontFamily" && typeof entry === "string") {
          const matched = lookupFont(entry);
          if (matched !== undefined) {
            next[key] = matched;
            changed = true;
            if (inferNonBoldCustomWeight) {
              inferredWeight = inferNonBoldFaceWeight(matched, value);
            }
            continue;
          }
        }
        const res = rewriteValue(entry);
        if (res.changed) changed = true;
        next[key] = res.value;
      }
      if (inferredWeight !== undefined) {
        // OOXML <a:rPr b="0"> collapses all non-bold weights; restore the thin/light face weight matching the source font file's OS/2 usWeightClass.
        next["fontWeight"] = inferredWeight;
        changed = true;
      }
      return { changed, value: changed ? next : value };
    }
    return { changed: false, value };
  };

  const visitNode = (node: DocumentCommonNode): void => {
    if (visited.has(node)) return;
    visited.add(node);

    const props = node.props;
    if (isTheme(props["theme"])) {
      rewriteTheme(props["theme"]);
    }
    const patch: Record<string, unknown> = {};
    let hasPropPatch = false;
    for (const [key, val] of Object.entries(props)) {
      if (key === "style" || key === "theme" || key === "children") continue;
      const res = rewriteValue(val);
      if (res.changed) {
        patch[key] = res.value;
        hasPropPatch = true;
      }
    }
    if (hasPropPatch) {
      node.set(patch);
    }

    const styleVal = props["style"];
    if (isPlainRecord(styleVal)) {
      const styleRes = rewriteValue(styleVal);
      if (styleRes.changed) {
        node.style(styleRes.value as Readonly<Record<string, unknown>>);
      }
    }

    for (const child of node.children) {
      if (child instanceof DocumentCommonNode) {
        visitNode(child);
      }
    }
    if (node instanceof DeckContainer) {
      for (const master of node.masters) {
        visitNode(master);
      }
    } else if (node instanceof SlideMasterContainer) {
      for (const layout of node.layouts) {
        visitNode(layout);
      }
    }
  };

  if (root instanceof DocumentCommonNode) {
    visitNode(root);
  }
}

function summarizeFileObject(obj: FileObject): unknown {
  return {
    kind: obj.kind,
    name: obj.name,
    geometry: obj.geometry,
    boxEmu: obj.boxEmu,
    ...(obj.placeholder === undefined ? {} : { placeholder: obj.placeholder }),
    ...(obj.link === undefined ? {} : { link: obj.link }),
    ...(obj.crop === undefined ? {} : { crop: obj.crop }),
    ...(obj.text === undefined ? {} : { text: obj.text.text }),
    ...(obj.table === undefined
      ? {}
      : {
          table: {
            cells: obj.table.cells,
            merges: obj.table.merges,
            columnWidthsEmu: obj.table.columnWidthsEmu,
            rowHeightsEmu: obj.table.rowHeightsEmu,
          },
        }),
    ...(obj.chart === undefined
      ? {}
      : {
          chart: {
            types: obj.chart.types,
            series: obj.chart.series.map((s) => ({
              name: s.name,
              groupIndex: s.groupIndex,
              ...(s.categories === undefined
                ? {}
                : { categories: s.categories }),
              ...(s.values === undefined ? {} : { values: s.values }),
              ...(s.x === undefined ? {} : { x: s.x }),
              ...(s.y === undefined ? {} : { y: s.y }),
            })),
          },
        }),
    ...(obj.members === undefined
      ? {}
      : { members: obj.members.map(summarizeFileObject) }),
  };
}

function assertNativeRoundtripEquivalence(
  scenarioId: string,
  initial: ValidationResult,
  roundtrip: ValidationResult,
): void {
  const label = `${scenarioId} roundtrip`;
  assert.equal(
    roundtrip.presentation.widthEmu,
    initial.presentation.widthEmu,
    `${label} widthEmu`,
  );
  assert.equal(
    roundtrip.presentation.heightEmu,
    initial.presentation.heightEmu,
    `${label} heightEmu`,
  );
  assert.deepEqual(
    roundtrip.presentation.sections,
    initial.presentation.sections,
    `${label} sections`,
  );
  assert.equal(
    roundtrip.masters?.length,
    initial.masters?.length,
    `${label} master count`,
  );
  for (let m = 0; m < (initial.masters?.length ?? 0); m += 1) {
    const im = initial.masters![m]!;
    const rm = roundtrip.masters![m]!;
    assert.equal(rm.name, im.name, `${label} master[${m}] name`);
    assert.equal(rm.preserve, im.preserve, `${label} master[${m}] preserve`);
    assert.deepEqual(rm.colorMap, im.colorMap, `${label} master[${m}] clrMap`);
    assert.equal(
      rm.background?.type,
      im.background?.type,
      `${label} master[${m}] background`,
    );
    assert.equal(
      rm.objects?.length,
      im.objects?.length,
      `${label} master[${m}] object count`,
    );
    assert.equal(
      rm.layouts.length,
      im.layouts.length,
      `${label} master[${m}] layout count`,
    );
    for (let l = 0; l < im.layouts.length; l += 1) {
      const il = im.layouts[l]!;
      const rl = rm.layouts[l]!;
      assert.equal(rl.name, il.name, `${label} master[${m}] layout[${l}] name`);
      assert.equal(rl.type, il.type, `${label} master[${m}] layout[${l}] type`);
      assert.equal(
        rl.preserve,
        il.preserve,
        `${label} master[${m}] layout[${l}] preserve`,
      );
      assert.equal(
        rl.showMasterShapes,
        il.showMasterShapes,
        `${label} master[${m}] layout[${l}] showMasterShapes`,
      );
      assert.equal(
        rl.showMasterPlaceholderAnimations,
        il.showMasterPlaceholderAnimations,
        `${label} master[${m}] layout[${l}] showMasterPhAnim`,
      );
      assert.deepEqual(
        rl.transition,
        il.transition,
        `${label} master[${m}] layout[${l}] transition`,
      );
      assert.equal(
        rl.advanceAfterMs,
        il.advanceAfterMs,
        `${label} master[${m}] layout[${l}] advanceAfterMs`,
      );
      assert.equal(
        rl.objects?.length,
        il.objects?.length,
        `${label} master[${m}] layout[${l}] object count`,
      );
    }
  }

  assert.equal(
    roundtrip.slides.length,
    initial.slides.length,
    `${label} slide count`,
  );
  for (let s = 0; s < initial.slides.length; s += 1) {
    const is = initial.slides[s]!;
    const rs = roundtrip.slides[s]!;
    assert.equal(rs.id, is.id, `${label} slide[${s}] id`);
    assert.equal(rs.hidden, is.hidden, `${label} slide[${s}] hidden`);
    assert.equal(rs.section, is.section, `${label} slide[${s}] section`);
    assert.equal(
      rs.showMasterShapes,
      is.showMasterShapes,
      `${label} slide[${s}] showMasterShapes`,
    );
    assert.equal(
      rs.showMasterPlaceholderAnimations,
      is.showMasterPlaceholderAnimations,
      `${label} slide[${s}] showMasterPhAnim`,
    );
    assert.deepEqual(
      rs.transition,
      is.transition,
      `${label} slide[${s}] transition`,
    );
    assert.equal(
      rs.advanceAfterMs,
      is.advanceAfterMs,
      `${label} slide[${s}] advanceAfterMs`,
    );
    assert.equal(
      rs.notes?.text.text,
      is.notes?.text.text,
      `${label} slide[${s}] notes`,
    );
    assert.equal(
      rs.background?.type,
      is.background?.type,
      `${label} slide[${s}] background`,
    );
    assert.deepEqual(
      rs.objects.map(summarizeFileObject),
      is.objects.map(summarizeFileObject),
      `${label} slide[${s}] objects`,
    );
    assert.equal(
      rs.animations?.length,
      is.animations?.length,
      `${label} slide[${s}] animations count`,
    );
  }
}

/** Executes the formal `import -> resolve -> emit -> paint -> validatePptx` roundtrip over an emitted PPTX. */
export async function runScenarioRoundtrip(
  pptx: Uint8Array,
  options: {
    readonly pages?: readonly number[];
    readonly emitOptions?: EmitOptions;
    readonly fontSourceDocument?: CommonNode;
    readonly allowFontWeight400Fallback?: boolean;
  } = {},
): Promise<RoundtripRunOutput> {
  const detachedCopy = new Uint8Array(pptx);
  const imported = importDocument(detachedCopy);
  // Immediately zero caller bytes to prove the imported document retains no buffer references.
  detachedCopy.fill(0);

  if (options.fontSourceDocument !== undefined) {
    reattachUnembeddedFonts(
      imported.document,
      collectFontsFromDocument(options.fontSourceDocument),
      options.allowFontWeight400Fallback === true,
    );
  }

  const previewOptions =
    options.pages === undefined ? {} : { pages: options.pages };
  const resolution = resolveDocument(imported.document);
  const { pptx: reemittedPptx, report } = emit(resolution, options.emitOptions);
  const { pages } = paint(resolution, previewOptions);
  const preview: PreviewResult = Object.freeze({
    report: resolution.report,
    query: resolution.query,
    pages,
  });
  const validation = await validatePptx(reemittedPptx);
  assert.deepEqual(
    validation.errors,
    [],
    "Roundtrip re-emitted PPTX validation failed",
  );
  return {
    importReport: imported.report,
    importedDocument: imported.document,
    resolution,
    report,
    pptx: reemittedPptx,
    preview,
    validation,
  };
}

function assertPageOutput(
  page: PreviewPageResult,
  scenario: CapabilityScenario,
  expectedSize: { readonly width?: number; readonly height?: number },
  label: string,
): void {
  const isRejection = scenario.kind === "rejection";
  const hasError = page.diagnostics.some((item) => item.severity === "error");

  if (isRejection && hasError) {
    assert.equal(page.png, undefined, `${label} failed page must omit png`);
    assert.equal(page.svg, undefined, `${label} failed page must omit svg`);
    for (const diagnostic of page.diagnostics) {
      assert.notEqual(diagnostic.path.trim(), "", `${label} diagnostic path`);
      assert.notEqual(
        diagnostic.message.trim(),
        "",
        `${label} diagnostic message`,
      );
    }
    return;
  }

  if (!isRejection) {
    assert.deepEqual(
      page.diagnostics,
      [],
      `${label} positive scenario must have zero preview diagnostics`,
    );
  }

  assert.ok(page.png instanceof Uint8Array, `${label} expected png`);
  assert.ok(
    typeof page.svg === "string" && page.svg.trim().length > 0,
    `${label} expected non-empty svg`,
  );
  const decoded = PNG.sync.read(Buffer.from(page.png));
  assert.ok(decoded.width > 0 && decoded.height > 0, `${label} PNG dimensions`);
  if (expectedSize.width !== undefined && expectedSize.height !== undefined) {
    assert.equal(
      decoded.width,
      expectedSize.width,
      `${label} PNG width must match slide width`,
    );
    assert.equal(
      decoded.height,
      expectedSize.height,
      `${label} PNG height must match slide height`,
    );
  }
}

function assertDualAuthorEquivalence(
  scenario: CapabilityScenario,
  jsxRun: ScenarioRunOutput,
  docRun: ScenarioRunOutput,
): void {
  assert.deepEqual(
    jsxRun.report,
    docRun.report,
    `${scenario.id} build report must match between jsx and document`,
  );
  assert.deepEqual(
    jsxRun.validation.presentation,
    docRun.validation.presentation,
    `${scenario.id} presentation file evidence must match between jsx and document`,
  );
  assert.deepEqual(
    jsxRun.validation.slides,
    docRun.validation.slides,
    `${scenario.id} slide file evidence must match between jsx and document`,
  );
  assert.equal(
    jsxRun.preview.pages.length,
    docRun.preview.pages.length,
    `${scenario.id} preview page count must match between jsx and document`,
  );
  for (let i = 0; i < jsxRun.preview.pages.length; i += 1) {
    const jsxPage = jsxRun.preview.pages[i]!;
    const docPage = docRun.preview.pages[i]!;
    const pageLabel = `${scenario.id} page ${jsxPage.page}`;
    assert.equal(jsxPage.page, docPage.page, `${pageLabel} page number`);
    assert.equal(jsxPage.id, docPage.id, `${pageLabel} slide id`);
    assert.deepEqual(
      jsxPage.diagnostics,
      docPage.diagnostics,
      `${pageLabel} diagnostics`,
    );
    assert.deepEqual(jsxPage.nodes, docPage.nodes, `${pageLabel} nodes`);
    assert.deepEqual(jsxPage.layouts, docPage.layouts, `${pageLabel} layouts`);
    assert.equal(jsxPage.svg, docPage.svg, `${pageLabel} svg`);
    if (jsxPage.png !== undefined) {
      assert.ok(docPage.png !== undefined, `${pageLabel} doc png`);
      assert.equal(
        Buffer.compare(jsxPage.png, docPage.png),
        0,
        `${pageLabel} png bytes must match between jsx and document`,
      );
    } else {
      assert.equal(docPage.png, undefined, `${pageLabel} doc png omitted`);
    }
  }
}

/**
 * Shared runner executing one capability-owned scenario across both real JSX and imperative
 * document authoring entries by resolving once, then feeding emit and paint independently.
 */
export async function runCapabilityScenario(
  scenario: CapabilityScenario,
  options: { readonly outputDir?: string } = {},
): Promise<readonly ScenarioRunOutput[]> {
  const results: ScenarioRunOutput[] = [];
  let primaryDocument: CommonNode | undefined;
  const previewOptions =
    scenario.pages === undefined ? {} : { pages: scenario.pages };

  for (const author of ["jsx", "document"] as const) {
    const document =
      author === "jsx"
        ? await expandDocument(scenario.jsx())
        : scenario.document();
    if (primaryDocument === undefined) {
      primaryDocument = document;
    }
    const resolution = resolveDocument(document);
    const { pptx, report } = emit(resolution, scenario.emitOptions);
    const { pages } = paint(resolution, previewOptions);
    const preview: PreviewResult = Object.freeze({
      report: resolution.report,
      query: resolution.query,
      pages,
    });

    const validation = await validatePptx(pptx);
    assert.deepEqual(
      validation.errors,
      [],
      `${scenario.id} (${author}) PPTX validation failed`,
    );

    const expectedSize = {
      ...(validation.presentation.widthEmu === undefined
        ? {}
        : {
            width: Math.round(validation.presentation.widthEmu / emuPerPx),
          }),
      ...(validation.presentation.heightEmu === undefined
        ? {}
        : {
            height: Math.round(validation.presentation.heightEmu / emuPerPx),
          }),
    };
    for (const page of preview.pages) {
      assertPageOutput(
        page,
        scenario,
        expectedSize,
        `${scenario.id}:${author}:p${page.page}`,
      );
    }

    await scenario.assert({ author, resolution, pptx, preview, validation });

    const artifacts: CollectedScenarioArtifact[] = [];
    if (options.outputDir !== undefined) {
      const baseDir = join(options.outputDir, scenario.id, author);
      await mkdir(baseDir, { recursive: true });
      const pptxPath = join(baseDir, "deck.pptx");
      await writeFile(pptxPath, pptx);
      for (const page of preview.pages) {
        let pngPath: string | undefined;
        if (page.png !== undefined) {
          pngPath = join(baseDir, `page-${page.page}-${page.id}.png`);
          await writeFile(pngPath, page.png);
        }
        artifacts.push({
          scenarioId: scenario.id,
          author,
          page: page.page,
          slideId: page.id,
          ...(pngPath === undefined ? {} : { pngPath }),
          pptxPath,
        });
      }
      await writeFile(
        join(baseDir, "summary.json"),
        JSON.stringify(
          {
            scenarioId: scenario.id,
            capability: scenario.capability,
            author,
            description: scenario.description,
            visualCheckpoints: scenario.visualCheckpoints,
            nativeCheckpoints: scenario.nativeCheckpoints,
            pages: preview.pages.map((page) => ({
              page: page.page,
              id: page.id,
              hasPng: page.png !== undefined,
              diagnostics: page.diagnostics,
              nodes: page.nodes.map((node) => ({
                id: node.id,
                authorId: node.authorId,
                kind: node.kind,
                box: node.box,
                screenBox: node.screenBox,
                path: node.path,
              })),
            })),
          },
          null,
          2,
        ),
      );
    }

    results.push({
      scenarioId: scenario.id,
      author,
      resolution,
      report,
      pptx,
      preview,
      validation,
      artifacts,
    });
  }

  assertDualAuthorEquivalence(scenario, results[0]!, results[1]!);

  const primary = results[0]!;
  const roundtrip = await runScenarioRoundtrip(primary.pptx, {
    ...(scenario.pages === undefined ? {} : { pages: scenario.pages }),
    ...(scenario.emitOptions === undefined
      ? {}
      : { emitOptions: scenario.emitOptions }),
    ...(primaryDocument === undefined
      ? {}
      : { fontSourceDocument: primaryDocument }),
    ...(scenario.kind === "rejection"
      ? {}
      : { allowFontWeight400Fallback: true }),
  });

  if (scenario.kind !== "rejection") {
    assert.deepEqual(
      roundtrip.importReport.diagnostics,
      [],
      `${scenario.id} roundtrip import must have zero unexpected diagnostics`,
    );
    const roundtripExpectedSize = {
      ...(roundtrip.validation.presentation.widthEmu === undefined
        ? {}
        : {
            width: Math.round(
              roundtrip.validation.presentation.widthEmu / emuPerPx,
            ),
          }),
      ...(roundtrip.validation.presentation.heightEmu === undefined
        ? {}
        : {
            height: Math.round(
              roundtrip.validation.presentation.heightEmu / emuPerPx,
            ),
          }),
    };
    for (const page of roundtrip.preview.pages) {
      assertPageOutput(
        page,
        scenario,
        roundtripExpectedSize,
        `${scenario.id}:roundtrip:p${page.page}`,
      );
    }
    assertNativeRoundtripEquivalence(
      scenario.id,
      primary.validation,
      roundtrip.validation,
    );
  } else {
    assert.ok(
      roundtrip.importReport.diagnostics.every(
        (diag) =>
          diag.category === "corrupt-input" && diag.code === "corrupt-part",
      ),
      `${scenario.id} rejection roundtrip import reported unexpected diagnostics: ${JSON.stringify(roundtrip.importReport.diagnostics)}`,
    );
    const isolatedCorruptParts = roundtrip.importReport.diagnostics.length > 0;
    const roundtripExpectedSize = {
      ...(roundtrip.validation.presentation.widthEmu === undefined
        ? {}
        : {
            width: Math.round(
              roundtrip.validation.presentation.widthEmu / emuPerPx,
            ),
          }),
      ...(roundtrip.validation.presentation.heightEmu === undefined
        ? {}
        : {
            height: Math.round(
              roundtrip.validation.presentation.heightEmu / emuPerPx,
            ),
          }),
    };
    for (const page of roundtrip.preview.pages) {
      assertPageOutput(
        page,
        scenario,
        roundtripExpectedSize,
        `${scenario.id}:roundtrip:p${page.page}`,
      );
    }
    if (!isolatedCorruptParts) {
      assert.equal(
        roundtrip.preview.pages.length,
        primary.preview.pages.length,
        `${scenario.id} rejection roundtrip page count`,
      );
      for (let i = 0; i < primary.preview.pages.length; i += 1) {
        const initPage = primary.preview.pages[i]!;
        const rtPage = roundtrip.preview.pages[i]!;
        const initHasError = initPage.diagnostics.some(
          (diag) => diag.severity === "error",
        );
        const rtHasError = rtPage.diagnostics.some(
          (diag) => diag.severity === "error",
        );
        assert.equal(
          rtHasError,
          initHasError,
          `${scenario.id} rejection roundtrip page ${rtPage.page} error state mismatch`,
        );
      }
    }
  }

  if (scenario.assertNative !== undefined) {
    await scenario.assertNative({
      stage: "initial",
      pptx: primary.pptx,
      validation: primary.validation,
      report: primary.report,
    });
    await scenario.assertNative({
      stage: "roundtrip",
      pptx: roundtrip.pptx,
      validation: roundtrip.validation,
      report: roundtrip.report,
    });
  }

  return results;
}
