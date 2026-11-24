import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  mkdir,
  mkdtemp,
  rename,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { validatePptx, type ValidationResult } from "@pptxdk/validation";
import type { BuildOptions, CodegenResult } from "pptxdk";
import { buildFromSource, type BuildResult } from "pptxdk/node";

const packageRoot = fileURLToPath(new URL("../../", import.meta.url));
const tscBin = join(packageRoot, "node_modules/.bin/tsc");

export type GeneratedSourceTypecheckCase = {
  readonly label: string;
  readonly result: CodegenResult;
};

export type RelocatedRebuildOutcome = {
  readonly built: BuildResult;
  readonly file: ValidationResult;
};

/** Writes all relative-path entries of a `CodegenResult` into `targetDir`. */
export async function writeCodegenFiles(
  targetDir: string,
  files: CodegenResult["files"],
): Promise<void> {
  for (const [relPath, content] of Object.entries(files)) {
    const absPath = join(targetDir, relPath);
    await mkdir(dirname(absPath), { recursive: true });
    await writeFile(absPath, content);
  }
}

async function linkConsumerNodeModules(
  nodeModulesDir: string,
  includeTypes: boolean,
): Promise<void> {
  await mkdir(nodeModulesDir, { recursive: true });
  await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
  await symlink(
    join(packageRoot, "node_modules/react"),
    join(nodeModulesDir, "react"),
    "dir",
  );
  if (includeTypes) {
    await mkdir(join(nodeModulesDir, "@types"), { recursive: true });
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
  }
}

/**
 * Batch-checks generated `vanilla` (`deck.ts`) and `react` (`deck.tsx`) outputs in a single
 * isolated consumer directory against published SDK types with strict `tsc --noEmit`.
 */
export async function verifyGeneratedSourcesTypecheck(
  cases: readonly GeneratedSourceTypecheckCase[],
): Promise<void> {
  if (cases.length === 0) return;

  const dir = await mkdtemp(join(tmpdir(), "pptxdk-codegen-tsc-"));
  try {
    await writeFile(
      join(dir, "package.json"),
      JSON.stringify({ type: "module" }),
    );
    await linkConsumerNodeModules(join(dir, "node_modules"), true);
    const caseManifest: string[] = [];

    for (const [index, item] of cases.entries()) {
      const caseDir = join(dir, "cases", `case-${index}`);
      await writeCodegenFiles(caseDir, item.result.files);
      caseManifest.push(`case-${index}: ${item.label} (${item.result.entry})`);
    }

    await writeFile(
      join(dir, "tsconfig.json"),
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
      {
        cwd: dir,
        encoding: "utf8",
      },
    );

    assert.equal(
      tscResult.status,
      0,
      `Consumer strict tsc failed for:\n${caseManifest.join("\n")}\n\n${tscResult.stdout}\n${tscResult.stderr}`,
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

function assertVanillaBuildFromSourceIsolatesReactAndPreview(
  entryPath: string,
): void {
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
      specifier.includes("@pptxdk/preview") ||
      specifier.includes("@resvg/resvg-js")
    ) {
      throw new Error("Vanilla buildFromSource loaded forbidden dependency: " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { buildFromSource } = await import("pptxdk/node");
const built = await buildFromSource(${JSON.stringify(entryPath)});
if (!(built.pptx instanceof Uint8Array) || built.pptx.byteLength === 0) {
  throw new Error("Expected non-empty rebuilt PPTX bytes");
}
process.stdout.write("ok");
`;

  const child = spawnSync(
    process.execPath,
    ["--conditions=source", "--input-type=module", "--eval", script],
    { cwd: packageRoot, encoding: "utf8" },
  );
  assert.equal(
    child.status,
    0,
    `Isolated vanilla buildFromSource check failed:\n${child.stderr}\n${child.stdout}`,
  );
  assert.equal(child.stdout, "ok");
}

/**
 * Writes generated files to an origin directory, relocates them via `rename` to a nested
 * project path unrelated to `process.cwd()`, rebuilds via `buildFromSource`, and validates
 * the resulting PPTX archive via `validatePptx`.
 */
export async function rebuildGeneratedSourceRelocated(
  result: CodegenResult,
  options: {
    readonly verifyVanillaIsolation?: boolean;
    readonly buildOptions?: BuildOptions;
  } = {},
): Promise<RelocatedRebuildOutcome> {
  const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-codegen-rebuild-"));
  try {
    const consumerRoot = join(tempRoot, "consumer");
    await linkConsumerNodeModules(join(consumerRoot, "node_modules"), false);

    const originDir = join(tempRoot, "origin");
    await writeCodegenFiles(originDir, result.files);

    const relocatedParent = join(consumerRoot, "relocated", "nested");
    await mkdir(relocatedParent, { recursive: true });
    const relocatedDir = join(relocatedParent, "project");
    await rename(originDir, relocatedDir);

    const entryPath = join(relocatedDir, result.entry);
    if (
      result.target === "vanilla" &&
      options.verifyVanillaIsolation !== false
    ) {
      assertVanillaBuildFromSourceIsolatesReactAndPreview(entryPath);
    }

    const built = await buildFromSource(entryPath, options.buildOptions);
    const file = await validatePptx(built.pptx);
    return { built, file };
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}
