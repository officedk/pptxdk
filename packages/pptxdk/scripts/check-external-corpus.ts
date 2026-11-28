import { spawn, spawnSync } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rm,
  stat,
  symlink,
  writeFile,
} from "node:fs/promises";
import { availableParallelism, tmpdir } from "node:os";
import { dirname, extname, join, resolve as resolvePath } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import {
  BuildError,
  CodegenError,
  CommonNode,
  ImportError,
  codegen,
  defineTheme,
  emit,
  font,
  fromFontCollection,
  importPptx,
  isTheme,
  isThemeFontSlot,
  resolve,
  type CodegenResult,
  type Font,
  type FontFace,
  type FontFaceStyle,
  type Theme,
} from "@officedk/pptxdk";
import { paint, type PreviewError } from "@officedk/pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import * as fontkit from "fontkit";
import { z } from "zod";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const tscBin = join(packageRoot, "node_modules/.bin/tsc");
const defaultTimeoutMs = 30_000;
const defaultConcurrency = Math.min(8, Math.max(1, availableParallelism()));

export type ExternalCorpusOptions = {
  readonly input: string;
  readonly fontsDir?: string | undefined;
  readonly reportPath?: string | undefined;
  readonly timeoutMs?: number | undefined;
  readonly concurrency?: number | undefined;
};

export type SampleCheckResult = {
  readonly file: string;
  readonly status: "passed" | "rejected-typed" | "violated";
  readonly importErrorCode?: string | undefined;
  readonly slideCount?: number | undefined;
  readonly importDiagnosticsCount?: number | undefined;
  readonly codegenStatus?: "passed" | "failed" | "not-run" | undefined;
  readonly emitStatus?:
    | "passed"
    | "skipped-font-unavailable"
    | "failed"
    | "not-run"
    | undefined;
  readonly paintStatus?:
    | "passed"
    | "completed-with-page-errors"
    | "failed"
    | "not-run"
    | undefined;
  readonly previewErrors?: readonly PreviewError[] | undefined;
  readonly signal?: string | undefined;
  readonly exitCode?: number | null | undefined;
  readonly timedOut?: boolean | undefined;
  readonly error?: string | undefined;
};

export type ExternalCorpusReport = {
  readonly inputPath: string;
  readonly fontsDir?: string | undefined;
  readonly totalFiles: number;
  readonly passed: boolean;
  readonly violations: number;
  readonly summary: {
    readonly importedCount: number;
    readonly typedRejectedCount: number;
    readonly codegenPassedCount: number;
    readonly batchTscPassed: boolean;
    readonly emitPassedCount: number;
    readonly emitSkippedFontUnavailableCount: number;
    readonly paintCompletedCount: number;
    readonly paintPageErrorsCount: number;
    readonly crashedCount: number;
    readonly timedOutCount: number;
  };
  readonly batchTscDiagnostics?: string | undefined;
  readonly results: readonly SampleCheckResult[];
};

const workerStageSchema = z.object({
  file: z.string(),
  status: z.enum(["passed", "rejected-typed", "violated"]),
  importErrorCode: z.string().optional(),
  slideCount: z.number().int().nonnegative().optional(),
  importDiagnosticsCount: z.number().int().nonnegative().optional(),
  codegenStatus: z.enum(["passed", "failed", "not-run"]).optional(),
  emitStatus: z
    .enum(["passed", "skipped-font-unavailable", "failed", "not-run"])
    .optional(),
  paintStatus: z
    .enum(["passed", "completed-with-page-errors", "failed", "not-run"])
    .optional(),
  previewErrors: z.array(z.unknown()).optional(),
  error: z.string().optional(),
});

type LoadedFontPool = {
  readonly byFamily: ReadonlyMap<string, readonly FontFace[]>;
  readonly fallback: readonly FontFace[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

async function discoverTopLevelPptxFiles(
  inputPath: string,
): Promise<readonly string[]> {
  const absInput = resolvePath(inputPath);
  const info = await stat(absInput);
  if (info.isFile()) {
    if (!absInput.toLowerCase().endsWith(".pptx")) {
      throw new Error(`Input file must have a .pptx extension: ${absInput}`);
    }
    return [absInput];
  }
  if (!info.isDirectory()) {
    throw new Error(
      `Input path must be a .pptx file or directory: ${absInput}`,
    );
  }
  const entries = await readdir(absInput, { withFileTypes: true });
  // Only read direct first-layer .pptx files; never recurse or unpack archives.
  const files = entries
    .filter(
      (entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".pptx"),
    )
    .map((entry) => join(absInput, entry.name))
    .toSorted((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  if (files.length === 0) {
    throw new Error(`No top-level .pptx files found in directory: ${absInput}`);
  }
  return files;
}

function supportsCjkFallback(fontData: Uint8Array | undefined): boolean {
  if (fontData === undefined) return false;
  try {
    const opened = fontkit.create(Buffer.from(fontData));
    return (
      "hasGlyphForCodePoint" in opened &&
      typeof opened.hasGlyphForCodePoint === "function" &&
      opened.hasGlyphForCodePoint(0x4e00) &&
      opened.glyphForCodePoint(0x4e00).id !== 0
    );
  } catch {
    return false;
  }
}

async function loadTopLevelSubstituteFontPool(
  fontsDir: string,
): Promise<LoadedFontPool> {
  const absDir = resolvePath(fontsDir);
  const info = await stat(absDir);
  if (!info.isDirectory()) {
    throw new Error(`Fonts path must be a directory: ${absDir}`);
  }
  const entries = await readdir(absDir, { withFileTypes: true });
  const fontFiles = entries
    .filter((entry) => {
      if (!entry.isFile()) return false;
      const ext = extname(entry.name).toLowerCase();
      return (
        ext === ".ttf" || ext === ".otf" || ext === ".ttc" || ext === ".otc"
      );
    })
    .map((entry) => join(absDir, entry.name))
    .toSorted((a, b) => (a < b ? -1 : a > b ? 1 : 0));

  const byFamilyBuckets = new Map<string, Map<string, FontFace>>();
  const fallbackBySlot = new Map<string, FontFace>();
  const fallbackPreferredSlots = new Set<string>();

  const registerFamilyFace = (
    familyKey: string,
    face: FontFace,
    isPreferredFallback: boolean,
  ) => {
    const slotKey = `${face.weight ?? 400}:${face.style ?? "normal"}`;
    let bucket = byFamilyBuckets.get(familyKey);
    if (bucket === undefined) {
      bucket = new Map<string, FontFace>();
      byFamilyBuckets.set(familyKey, bucket);
    }
    if (!bucket.has(slotKey)) {
      bucket.set(slotKey, face);
    }
    if (
      !fallbackBySlot.has(slotKey) ||
      (isPreferredFallback && !fallbackPreferredSlots.has(slotKey))
    ) {
      fallbackBySlot.set(slotKey, face);
      if (isPreferredFallback) {
        fallbackPreferredSlots.add(slotKey);
      }
    }
  };

  for (const filePath of fontFiles) {
    const raw = new Uint8Array(await readFile(filePath));
    const ext = extname(filePath).toLowerCase();
    if (ext === ".ttc" || ext === ".otc") {
      let collectionFamilies: string[] = [];
      try {
        const opened = fontkit.create(Buffer.from(raw)) as {
          readonly fonts?: readonly { readonly familyName?: unknown }[];
        };
        if (Array.isArray(opened.fonts)) {
          collectionFamilies = [
            ...new Set(
              opened.fonts
                .map((member) =>
                  typeof member.familyName === "string"
                    ? member.familyName.trim()
                    : "",
                )
                .filter((name) => name !== ""),
            ),
          ];
        }
      } catch {
        continue;
      }
      for (const fam of collectionFamilies) {
        try {
          const unpackedFont = fromFontCollection(raw, fam, {
            substitute: true,
          });
          const faces = unpackedFont.faces ?? [];
          if (faces.length === 0) continue;
          const key = fam.toLowerCase();
          for (const face of faces) {
            registerFamilyFace(key, face, supportsCjkFallback(face.data));
          }
        } catch {
          continue;
        }
      }
      continue;
    }
    let familyName: string | undefined;
    let weight = 400;
    let style: FontFaceStyle = "normal";
    try {
      const opened = fontkit.create(Buffer.from(raw));
      if (
        "fvar" in opened &&
        opened["fvar"] !== undefined &&
        opened["fvar"] !== null
      ) {
        continue;
      }
      if ("familyName" in opened && typeof opened.familyName === "string") {
        familyName = opened.familyName.trim();
      }
      if ("OS/2" in opened && isRecord(opened["OS/2"])) {
        const weightClass = opened["OS/2"]["usWeightClass"];
        if (
          typeof weightClass === "number" &&
          Number.isInteger(weightClass) &&
          weightClass >= 1 &&
          weightClass <= 1000
        ) {
          weight = weightClass;
        }
        const fsSelection = opened["OS/2"]["fsSelection"];
        const isItalicFromFs =
          (isRecord(fsSelection) &&
            (Boolean(fsSelection["italic"]) ||
              Boolean(fsSelection["oblique"]))) ||
          (typeof fsSelection === "number" && (fsSelection & 1) !== 0);
        const isItalicFromAngle =
          "italicAngle" in opened &&
          typeof opened.italicAngle === "number" &&
          opened.italicAngle !== 0;
        if (isItalicFromFs || isItalicFromAngle) {
          style = "italic";
        }
      } else if (
        "italicAngle" in opened &&
        typeof opened.italicAngle === "number" &&
        opened.italicAngle !== 0
      ) {
        style = "italic";
      }
    } catch {
      continue;
    }

    const face: FontFace = { weight, style, data: raw };
    if (familyName !== undefined && familyName !== "") {
      const key = familyName.toLowerCase();
      registerFamilyFace(key, face, supportsCjkFallback(raw));
    } else {
      const slotKey = `${face.weight ?? 400}:${face.style ?? "normal"}`;
      if (!fallbackBySlot.has(slotKey)) {
        fallbackBySlot.set(slotKey, face);
      }
    }
  }

  const byFamily = new Map<string, readonly FontFace[]>(
    [...byFamilyBuckets.entries()].map(([key, bucket]) => [
      key,
      [...bucket.values()],
    ]),
  );
  const fallback = [...fallbackBySlot.values()];

  return { byFamily, fallback };
}

function createSubstituteFont(
  familyName: string,
  pool: LoadedFontPool,
  cache: Map<string, Font>,
): Font {
  const cached = cache.get(familyName);
  if (cached !== undefined) return cached;
  const faces = pool.byFamily.get(familyName.toLowerCase()) ?? pool.fallback;
  const created = font(familyName, faces, { substitute: true });
  cache.set(familyName, created);
  return created;
}

function replaceThemeFonts(
  theme: Theme,
  pool: LoadedFontPool,
  cache: Map<string, Font>,
): Theme {
  const major = theme.fonts?.major;
  const minor = theme.fonts?.minor;
  const majorEastAsian = theme.fonts?.majorEastAsian;
  const minorEastAsian = theme.fonts?.minorEastAsian;
  const nextMajor =
    typeof major === "string"
      ? createSubstituteFont(major, pool, cache)
      : major;
  const nextMinor =
    typeof minor === "string"
      ? createSubstituteFont(minor, pool, cache)
      : minor;
  const nextMajorEa =
    typeof majorEastAsian === "string"
      ? createSubstituteFont(majorEastAsian, pool, cache)
      : majorEastAsian;
  const nextMinorEa =
    typeof minorEastAsian === "string"
      ? createSubstituteFont(minorEastAsian, pool, cache)
      : minorEastAsian;
  if (
    nextMajor === major &&
    nextMinor === minor &&
    nextMajorEa === majorEastAsian &&
    nextMinorEa === minorEastAsian
  ) {
    return theme;
  }
  return defineTheme({
    name: theme.name,
    colors: theme.colors,
    fonts: {
      ...(nextMajor === undefined ? {} : { major: nextMajor }),
      ...(nextMinor === undefined ? {} : { minor: nextMinor }),
      ...(nextMajorEa === undefined ? {} : { majorEastAsian: nextMajorEa }),
      ...(nextMinorEa === undefined ? {} : { minorEastAsian: nextMinorEa }),
    },
    ...(theme.formatScheme === undefined
      ? {}
      : { formatScheme: theme.formatScheme }),
  });
}

function replacePlainFonts(
  value: unknown,
  pool: LoadedFontPool,
  cache: Map<string, Font>,
): unknown {
  if (value === null || typeof value !== "object") return value;
  if (value instanceof Uint8Array || value instanceof URL) return value;
  if (value instanceof CommonNode) return value;
  if (isTheme(value)) return replaceThemeFonts(value, pool, cache);
  if (Array.isArray(value)) {
    let changed = false;
    const mapped = value.map((item) => {
      const next = replacePlainFonts(item, pool, cache);
      if (next !== item) changed = true;
      return next;
    });
    return changed ? mapped : value;
  }
  if (!isRecord(value)) return value;

  let changed = false;
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value)) {
    if (
      (key === "fontFamily" ||
        key === "eastAsianFontFamily" ||
        key === "major" ||
        key === "minor" ||
        key === "majorEastAsian" ||
        key === "minorEastAsian") &&
      typeof entry === "string" &&
      entry.trim() !== "" &&
      !isThemeFontSlot(entry)
    ) {
      out[key] = createSubstituteFont(entry, pool, cache);
      changed = true;
    } else {
      const next = replacePlainFonts(entry, pool, cache);
      if (next !== entry) changed = true;
      out[key] = next;
    }
  }
  return changed ? out : value;
}

function applySubstituteFontsToDocument(
  root: CommonNode,
  pool: LoadedFontPool,
  cache: Map<string, Font>,
  visited = new Set<CommonNode>(),
): void {
  if (pool.fallback.length === 0 || visited.has(root)) return;
  visited.add(root);

  const updates: Record<string, unknown> = {};
  let hasUpdate = false;
  let styleUpdate: Record<string, unknown> | undefined;
  for (const [key, value] of Object.entries(root.props)) {
    if (key === "children") continue;
    if (value instanceof CommonNode) {
      applySubstituteFontsToDocument(value, pool, cache, visited);
      continue;
    }
    if (key === "theme" && isTheme(value)) {
      const nextTheme = replaceThemeFonts(value, pool, cache);
      if (nextTheme !== value) {
        updates[key] = nextTheme;
        hasUpdate = true;
      }
      continue;
    }
    const nextValue = replacePlainFonts(value, pool, cache);
    if (nextValue !== value) {
      if (key === "style" && isRecord(nextValue)) {
        styleUpdate = nextValue;
      } else {
        updates[key] = nextValue;
        hasUpdate = true;
      }
    }
  }
  if (styleUpdate !== undefined) {
    root.style(styleUpdate);
  }
  if (hasUpdate) {
    root.set(updates);
  }
  for (const child of root.children) {
    if (child instanceof CommonNode) {
      applySubstituteFontsToDocument(child, pool, cache, visited);
    }
  }
}

async function writeCodegenFiles(
  targetDir: string,
  files: CodegenResult["files"],
): Promise<void> {
  for (const [relPath, content] of Object.entries(files)) {
    const absPath = join(targetDir, relPath);
    await mkdir(dirname(absPath), { recursive: true });
    await writeFile(absPath, content);
  }
}

function serializeWorkerPayload(
  payload: z.infer<typeof workerStageSchema>,
): string {
  return JSON.stringify(payload);
}

async function runSingleSampleWorker(params: {
  readonly file: string;
  readonly caseDir: string;
  readonly stageFile: string;
  readonly fontsDir?: string | undefined;
}): Promise<void> {
  const { file, caseDir, stageFile, fontsDir } = params;
  const writeStage = async (
    payload: z.infer<typeof workerStageSchema>,
  ): Promise<void> => {
    await writeFile(stageFile, serializeWorkerPayload(payload));
  };

  let imported: ReturnType<typeof importPptx>;
  const rawBytes = new Uint8Array(await readFile(file));
  try {
    imported = importPptx(rawBytes);
  } catch (error) {
    if (error instanceof ImportError) {
      await writeStage({
        file,
        status: "rejected-typed",
        importErrorCode: error.issue.code,
        codegenStatus: "not-run",
        emitStatus: "not-run",
        paintStatus: "not-run",
      });
      return;
    }
    await writeStage({
      file,
      status: "violated",
      codegenStatus: "not-run",
      emitStatus: "not-run",
      paintStatus: "not-run",
      error: `Invariant 1 uncaught import error: ${error instanceof Error ? error.message : String(error)}`,
    });
    return;
  } finally {
    rawBytes.fill(0);
  }

  let codegenStatus: "passed" | "failed" = "passed";
  let codegenFailure: string | undefined;
  try {
    const vanillaResult = codegen(imported.document, { target: "vanilla" });
    const reactResult = codegen(imported.document, { target: "react" });
    await writeCodegenFiles(join(caseDir, "vanilla"), vanillaResult.files);
    await writeCodegenFiles(join(caseDir, "react"), reactResult.files);
  } catch (error) {
    codegenStatus = "failed";
    codegenFailure =
      error instanceof CodegenError
        ? `Invariant 2 CodegenError: ${error.message}`
        : `Invariant 2 codegen error: ${error instanceof Error ? error.message : String(error)}`;
  }

  let emitStatus: "passed" | "skipped-font-unavailable" | "failed" = "passed";
  let emitFailure: string | undefined;
  try {
    const resolution = resolve(imported.document);
    const emitted = emit(resolution);
    const validated = await validatePptx(emitted.pptx);
    if (validated.errors.length > 0) {
      emitStatus = "failed";
      emitFailure = `Invariant 3 validatePptx errors: ${JSON.stringify(validated.errors)}`;
    }
  } catch (error) {
    if (
      error instanceof BuildError &&
      error.diagnostics.length > 0 &&
      error.diagnostics.every((diag) => diag.code === "font-unavailable")
    ) {
      emitStatus = "skipped-font-unavailable";
    } else {
      emitStatus = "failed";
      emitFailure = `Invariant 3 resolve/emit error: ${error instanceof Error ? error.message : String(error)}`;
    }
  }

  const baseStatus =
    codegenStatus === "failed" || emitStatus === "failed"
      ? "violated"
      : "passed";
  const baseError = [codegenFailure, emitFailure]
    .filter((msg) => msg !== undefined)
    .join("; ");

  // Persist invariants 1-3 before paint so a native crash in paint preserves earlier stage facts.
  await writeStage({
    file,
    status: baseStatus,
    slideCount: imported.report.slideCount,
    importDiagnosticsCount: imported.report.diagnostics.length,
    codegenStatus,
    emitStatus,
    paintStatus: fontsDir === undefined ? "not-run" : "failed",
    ...(baseError === "" ? {} : { error: baseError }),
  });

  if (fontsDir === undefined) return;

  try {
    const pool = await loadTopLevelSubstituteFontPool(fontsDir);
    const fontCache = new Map<string, Font>();
    const freshBytes = new Uint8Array(await readFile(file));
    const paintDoc = importPptx(freshBytes).document;
    freshBytes.fill(0);
    applySubstituteFontsToDocument(paintDoc, pool, fontCache);
    const resolution = resolve(paintDoc);
    const painted = paint(resolution);
    const previewErrors = painted.pages.flatMap((page) => page.errors);
    const paintStatus =
      previewErrors.length > 0 ? "completed-with-page-errors" : "passed";

    await writeStage({
      file,
      status: baseStatus,
      slideCount: imported.report.slideCount,
      importDiagnosticsCount: imported.report.diagnostics.length,
      codegenStatus,
      emitStatus,
      paintStatus,
      previewErrors,
      ...(baseError === "" ? {} : { error: baseError }),
    });
  } catch (error) {
    const paintError = `Invariant 4 paint exception: ${error instanceof Error ? error.message : String(error)}`;
    await writeStage({
      file,
      status: "violated",
      slideCount: imported.report.slideCount,
      importDiagnosticsCount: imported.report.diagnostics.length,
      codegenStatus,
      emitStatus,
      paintStatus: "failed",
      error: baseError === "" ? paintError : `${baseError}; ${paintError}`,
    });
  }
}

async function runIsolatedSample(params: {
  readonly file: string;
  readonly caseDir: string;
  readonly stageFile: string;
  readonly fontsDir?: string | undefined;
  readonly timeoutMs: number;
}): Promise<SampleCheckResult> {
  const scriptPath = fileURLToPath(import.meta.url);
  const args = [
    "--conditions=source",
    scriptPath,
    "--worker",
    "--input",
    params.file,
    "--case-dir",
    params.caseDir,
    "--stage-file",
    params.stageFile,
    ...(params.fontsDir === undefined ? [] : ["--fonts", params.fontsDir]),
  ];

  const childOutcome = await new Promise<{
    readonly exitCode: number | null;
    readonly signal: string | null;
    readonly timedOut: boolean;
    readonly stderr: string;
  }>((resolvePromise) => {
    const child = spawn(process.execPath, args, {
      cwd: packageRoot,
      stdio: ["ignore", "pipe", "pipe"],
    });
    const stderrChunks: Buffer[] = [];
    child.stderr.on("data", (chunk: Buffer) => stderrChunks.push(chunk));

    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      child.kill("SIGKILL");
    }, params.timeoutMs);

    child.once("close", (code, signal) => {
      clearTimeout(timer);
      resolvePromise({
        exitCode: code,
        signal,
        timedOut,
        stderr: Buffer.concat(stderrChunks).toString("utf8").trim(),
      });
    });
  });

  let parsedStage: z.infer<typeof workerStageSchema> | undefined;
  try {
    const raw = await readFile(params.stageFile, "utf8");
    parsedStage = workerStageSchema.parse(JSON.parse(raw));
  } catch {
    parsedStage = undefined;
  }

  if (
    childOutcome.timedOut ||
    childOutcome.signal !== null ||
    childOutcome.exitCode !== 0 ||
    parsedStage === undefined
  ) {
    const failureReason = childOutcome.timedOut
      ? `Sample timed out after ${params.timeoutMs}ms`
      : childOutcome.signal !== null
        ? `Worker terminated by signal ${childOutcome.signal}`
        : childOutcome.exitCode !== 0
          ? `Worker exited with code ${String(childOutcome.exitCode)}${childOutcome.stderr === "" ? "" : `: ${childOutcome.stderr}`}`
          : "Worker did not emit stage report";
    return {
      file: params.file,
      status: "violated",
      ...(parsedStage?.slideCount === undefined
        ? {}
        : { slideCount: parsedStage.slideCount }),
      ...(parsedStage?.importDiagnosticsCount === undefined
        ? {}
        : { importDiagnosticsCount: parsedStage.importDiagnosticsCount }),
      codegenStatus: parsedStage?.codegenStatus ?? "not-run",
      emitStatus: parsedStage?.emitStatus ?? "not-run",
      paintStatus:
        params.fontsDir === undefined
          ? (parsedStage?.paintStatus ?? "not-run")
          : "failed",
      ...(childOutcome.signal === null ? {} : { signal: childOutcome.signal }),
      exitCode: childOutcome.exitCode,
      timedOut: childOutcome.timedOut,
      error: [parsedStage?.error, failureReason]
        .filter((item) => item !== undefined && item !== "")
        .join("; "),
    };
  }

  return {
    file: parsedStage.file,
    status: parsedStage.status,
    ...(parsedStage.importErrorCode === undefined
      ? {}
      : { importErrorCode: parsedStage.importErrorCode }),
    ...(parsedStage.slideCount === undefined
      ? {}
      : { slideCount: parsedStage.slideCount }),
    ...(parsedStage.importDiagnosticsCount === undefined
      ? {}
      : { importDiagnosticsCount: parsedStage.importDiagnosticsCount }),
    ...(parsedStage.codegenStatus === undefined
      ? {}
      : { codegenStatus: parsedStage.codegenStatus }),
    ...(parsedStage.emitStatus === undefined
      ? {}
      : { emitStatus: parsedStage.emitStatus }),
    ...(parsedStage.paintStatus === undefined
      ? {}
      : { paintStatus: parsedStage.paintStatus }),
    ...(parsedStage.previewErrors === undefined
      ? {}
      : {
          previewErrors: parsedStage.previewErrors as readonly PreviewError[],
        }),
    ...(parsedStage.error === undefined ? {} : { error: parsedStage.error }),
  };
}

async function runBatchTscCheck(
  codegenBatchRoot: string,
  hasGeneratedSources: boolean,
): Promise<{ readonly passed: boolean; readonly diagnostics?: string }> {
  if (!hasGeneratedSources) {
    return { passed: true };
  }

  const nodeModulesDir = join(codegenBatchRoot, "node_modules");
  await mkdir(join(nodeModulesDir, "@officedk"), { recursive: true });
  await mkdir(join(nodeModulesDir, "@types"), { recursive: true });
  await symlink(packageRoot, join(nodeModulesDir, "@officedk/pptxdk"), "dir");
  await symlink(
    join(packageRoot, "node_modules/react"),
    join(nodeModulesDir, "react"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@types/react"),
    join(nodeModulesDir, "@types/react"),
    "dir",
  );
  await symlink(
    join(packageRoot, "node_modules/@types/node"),
    join(nodeModulesDir, "@types/node"),
    "dir",
  );

  await writeFile(
    join(codegenBatchRoot, "package.json"),
    JSON.stringify({ type: "module" }),
  );
  await writeFile(
    join(codegenBatchRoot, "tsconfig.json"),
    JSON.stringify({
      compilerOptions: {
        target: "es2024",
        module: "nodenext",
        moduleResolution: "nodenext",
        jsx: "react-jsx",
        jsxImportSource: "react",
        strict: true,
        noEmit: true,
        skipLibCheck: false,
        types: ["node"],
      },
      include: ["cases/**/*.ts", "cases/**/*.tsx"],
    }),
  );

  const tscResult = spawnSync(
    tscBin,
    ["-p", "tsconfig.json", "--pretty", "false"],
    { cwd: codegenBatchRoot, encoding: "utf8" },
  );
  if (tscResult.status === 0) {
    return { passed: true };
  }
  const diagnostics = `${tscResult.stdout}\n${tscResult.stderr}`.trim();
  return { passed: false, diagnostics };
}

/**
 * Executes the decoupled external PPTX corpus verification across a single `.pptx` file or
 * a flat directory of `.pptx` files in isolated child processes with batch strict `tsc`.
 */
export async function runExternalCorpusCheck(
  options: ExternalCorpusOptions,
): Promise<ExternalCorpusReport> {
  const timeoutMs = options.timeoutMs ?? defaultTimeoutMs;
  if (!Number.isInteger(timeoutMs) || timeoutMs <= 0) {
    throw new Error(`Invalid timeoutMs: ${String(timeoutMs)}`);
  }
  const concurrency = options.concurrency ?? defaultConcurrency;
  if (!Number.isInteger(concurrency) || concurrency <= 0) {
    throw new Error(`Invalid concurrency: ${String(concurrency)}`);
  }
  const files = await discoverTopLevelPptxFiles(options.input);

  const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-conditional-corpus-"));
  try {
    const results = Array.from<SampleCheckResult | undefined>({
      length: files.length,
    });
    const workerCount = Math.min(concurrency, files.length);
    let nextIndex = 0;
    await Promise.all(
      Array.from({ length: workerCount }, async () => {
        while (nextIndex < files.length) {
          const index = nextIndex;
          nextIndex += 1;
          const file = files[index]!;
          const caseDir = join(tempRoot, "cases", `s${String(index)}`);
          const stageFile = join(tempRoot, `stage-${String(index)}.json`);
          results[index] = await runIsolatedSample({
            file,
            caseDir,
            stageFile,
            fontsDir: options.fontsDir,
            timeoutMs,
          });
        }
      }),
    );
    const orderedResults = results.map((item) => {
      if (item === undefined) {
        throw new Error("Missing sample result from worker pool");
      }
      return item;
    });

    const codegenPassedCount = orderedResults.filter(
      (item) => item.codegenStatus === "passed",
    ).length;
    const batchTsc = await runBatchTscCheck(tempRoot, codegenPassedCount > 0);

    const typedRejectedCount = orderedResults.filter(
      (item) => item.status === "rejected-typed",
    ).length;
    const importedCount = orderedResults.filter(
      (item) =>
        item.status !== "rejected-typed" && item.slideCount !== undefined,
    ).length;
    const emitPassedCount = orderedResults.filter(
      (item) => item.emitStatus === "passed",
    ).length;
    const emitSkippedFontUnavailableCount = orderedResults.filter(
      (item) => item.emitStatus === "skipped-font-unavailable",
    ).length;
    const paintCompletedCount = orderedResults.filter(
      (item) =>
        item.paintStatus === "passed" ||
        item.paintStatus === "completed-with-page-errors",
    ).length;
    const paintPageErrorsCount = orderedResults.reduce(
      (sum, item) => sum + (item.previewErrors?.length ?? 0),
      0,
    );
    const crashedCount = orderedResults.filter(
      (item) => item.signal !== undefined,
    ).length;
    const timedOutCount = orderedResults.filter(
      (item) => item.timedOut === true,
    ).length;
    const sampleViolations = orderedResults.filter(
      (item) => item.status === "violated",
    ).length;
    const violations = sampleViolations + (batchTsc.passed ? 0 : 1);

    const report: ExternalCorpusReport = {
      inputPath: resolvePath(options.input),
      ...(options.fontsDir === undefined
        ? {}
        : { fontsDir: resolvePath(options.fontsDir) }),
      totalFiles: files.length,
      passed: violations === 0,
      violations,
      summary: {
        importedCount,
        typedRejectedCount,
        codegenPassedCount,
        batchTscPassed: batchTsc.passed,
        emitPassedCount,
        emitSkippedFontUnavailableCount,
        paintCompletedCount,
        paintPageErrorsCount,
        crashedCount,
        timedOutCount,
      },
      ...(batchTsc.diagnostics === undefined
        ? {}
        : { batchTscDiagnostics: batchTsc.diagnostics }),
      results: orderedResults,
    };

    if (options.reportPath !== undefined) {
      const absReport = resolvePath(options.reportPath);
      await mkdir(dirname(absReport), { recursive: true });
      await writeFile(absReport, JSON.stringify(report, null, 2));
    }

    return report;
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

type CliParseResult =
  | {
      readonly mode: "worker";
      readonly file: string;
      readonly caseDir: string;
      readonly stageFile: string;
      readonly fontsDir?: string | undefined;
    }
  | {
      readonly mode: "orchestrator";
      readonly options: ExternalCorpusOptions;
    };

function parseCliArgs(argv: readonly string[]): CliParseResult {
  let isWorker = false;
  let input = process.env["PPTXDK_CONDITIONAL_INPUT"];
  let fontsDir = process.env["PPTXDK_CONDITIONAL_FONTS"];
  let reportPath = process.env["PPTXDK_CONDITIONAL_REPORT"];
  const rawTimeout = process.env["PPTXDK_CONDITIONAL_TIMEOUT_MS"];
  let timeoutMs =
    rawTimeout === undefined || rawTimeout.trim() === ""
      ? undefined
      : Number(rawTimeout);
  const rawConcurrency = process.env["PPTXDK_CONDITIONAL_CONCURRENCY"];
  let concurrency =
    rawConcurrency === undefined || rawConcurrency.trim() === ""
      ? undefined
      : Number(rawConcurrency);
  let caseDir: string | undefined;
  let stageFile: string | undefined;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    if (arg === "--") continue;
    if (arg === "--worker") {
      isWorker = true;
      continue;
    }
    if (arg === "--input") {
      input = argv[++i];
      continue;
    }
    if (arg === "--fonts") {
      fontsDir = argv[++i];
      continue;
    }
    if (arg === "--report") {
      reportPath = argv[++i];
      continue;
    }
    if (arg === "--timeout-ms") {
      timeoutMs = Number(argv[++i]);
      continue;
    }
    if (arg === "--concurrency") {
      concurrency = Number(argv[++i]);
      continue;
    }
    if (arg === "--case-dir") {
      caseDir = argv[++i];
      continue;
    }
    if (arg === "--stage-file") {
      stageFile = argv[++i];
      continue;
    }
    if (!arg.startsWith("--") && input === undefined) {
      input = arg;
      continue;
    }
    throw new Error(`Unknown or unexpected argument: ${arg}`);
  }

  if (isWorker) {
    if (
      input === undefined ||
      caseDir === undefined ||
      stageFile === undefined
    ) {
      throw new Error(
        "Worker mode requires --input, --case-dir, and --stage-file",
      );
    }
    return {
      mode: "worker",
      file: input,
      caseDir,
      stageFile,
      fontsDir,
    };
  }

  if (input === undefined || input.trim() === "") {
    throw new Error(
      "Usage: check-external-corpus --input <pptx-file-or-flat-dir> [--fonts <flat-font-dir>] [--report <report.json>] [--timeout-ms <ms>] [--concurrency <n>]",
    );
  }
  return {
    mode: "orchestrator",
    options: {
      input,
      fontsDir,
      reportPath,
      timeoutMs,
      concurrency,
    },
  };
}

async function main(): Promise<void> {
  let parsed: CliParseResult;
  try {
    parsed = parseCliArgs(process.argv.slice(2));
  } catch (error) {
    process.stderr.write(
      `${error instanceof Error ? error.message : String(error)}\n`,
    );
    process.exitCode = 2;
    return;
  }

  if (parsed.mode === "worker") {
    await runSingleSampleWorker({
      file: parsed.file,
      caseDir: parsed.caseDir,
      stageFile: parsed.stageFile,
      fontsDir: parsed.fontsDir,
    });
    return;
  }

  const report = await runExternalCorpusCheck(parsed.options);
  process.stdout.write(`${JSON.stringify(report.summary)}\n`);
  if (!report.passed) {
    process.stderr.write(
      `External corpus check failed with ${report.violations} invariant violation(s)\n`,
    );
    process.exitCode = 1;
  }
}

const mainFile = process.argv[1];
if (
  mainFile !== undefined &&
  pathToFileURL(resolvePath(mainFile)).href === import.meta.url
) {
  await main();
}
