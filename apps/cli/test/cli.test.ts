import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import {
  mkdir,
  mkdtemp,
  readdir,
  readFile,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { validatePptx } from "@pptxdk/validation";
import { describe, expect, it } from "vitest";

import { runCli } from "../src/index.ts";

const fixture = (name: string) => join(import.meta.dirname, "fixtures", name);
const cli = join(import.meta.dirname, "../dist/cli.js");
const waitForClose = (child: ReturnType<typeof spawn>) =>
  new Promise<number>((resolve, reject) => {
    child.once("error", reject);
    child.once("close", (code) => resolve(code ?? 1));
  });

function runNode(args: readonly string[], cwd: string) {
  const result = spawnSync(process.execPath, [...args], {
    cwd,
    encoding: "utf8",
  });
  if (result.error !== undefined) throw result.error;
  return {
    code: result.status ?? 1,
    stdout: result.stdout,
    stderr: result.stderr,
  };
}

async function runInMemory(
  args: readonly string[],
  cwd: string,
  signal?: AbortSignal,
) {
  let stdout = "";
  let stderr = "";
  const code = await runCli({
    argv: ["node", "pptxdk", ...args],
    cwd,
    ...(signal !== undefined && { signal }),
    write: ({ destination, chunk }) => {
      if (destination === "stdout") stdout += chunk;
      else stderr += chunk;
    },
  });
  return { code, stdout, stderr };
}

describe("pptxdk CLI build", () => {
  it("exposes help, version, and schema contract discovery on pptxdk and build", async () => {
    const rootHelp = await runInMemory(["--help"], import.meta.dirname);
    expect(rootHelp.code).toBe(0);
    expect(rootHelp.stderr).toBe("");
    expect(rootHelp.stdout).toContain("Usage");
    expect(rootHelp.stdout).toContain("pptxdk <command>");
    expect(rootHelp.stdout).toContain("build");
    expect(rootHelp.stdout).toContain("--version");
    expect(rootHelp.stdout).toContain("--schema");

    const buildHelp = await runInMemory(["build", "-h"], import.meta.dirname);
    expect(buildHelp.code).toBe(0);
    expect(buildHelp.stderr).toBe("");
    expect(buildHelp.stdout).toContain(
      "pptxdk build [--output <mode>] <entry>",
    );
    expect(buildHelp.stdout).toContain("--out");
    expect(buildHelp.stdout).toContain("--export");
    expect(buildHelp.stdout).toContain("--target");
    expect(buildHelp.stdout).toContain("--embed-fonts");
    expect(buildHelp.stdout).toContain("--output");
    expect(buildHelp.stdout).toContain('(choices: "full", "subset")');

    const pkg = JSON.parse(
      await readFile(join(import.meta.dirname, "../package.json"), "utf8"),
    ) as { readonly version: string };

    const versionLong = await runInMemory(["--version"], import.meta.dirname);
    expect(versionLong.code).toBe(0);
    expect(versionLong.stderr).toBe("");
    expect(versionLong.stdout).toBe(`${pkg.version}\n`);

    const versionShort = await runInMemory(["-V"], import.meta.dirname);
    expect(versionShort.code).toBe(0);
    expect(versionShort.stdout).toBe(`${pkg.version}\n`);

    const rootSchema = await runInMemory(["--schema"], import.meta.dirname);
    expect(rootSchema.code).toBe(0);
    expect(JSON.parse(rootSchema.stdout)).toMatchObject({
      kind: "directory",
      schemaVersion: "5",
      command: "pptxdk",
      children: [{ command: "build" }],
    });

    const buildSchema = await runInMemory(
      ["build", "--schema"],
      import.meta.dirname,
    );
    expect(buildSchema.code).toBe(0);
    expect(JSON.parse(buildSchema.stdout)).toMatchObject({
      kind: "summary",
      schemaVersion: "5",
      command: "build",
    });

    const buildInputSchema = await runInMemory(
      ["build", "--schema=input"],
      import.meta.dirname,
    );
    expect(buildInputSchema.code).toBe(0);
    expect(JSON.parse(buildInputSchema.stdout)).toMatchObject({
      kind: "input",
      schemaVersion: "5",
      command: "build",
    });
  });

  it("builds default and named exports across working directories with delivery options and validates PPTX", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk cli "));
    try {
      const source = fixture("source deck.tsx");
      const defaultOutput = join(directory, "source deck.pptx");

      // Default output writes <stem>.pptx into current working directory when run outside source dir.
      const defaultResult = runNode([cli, "build", source], directory);
      expect(defaultResult.code).toBe(0);
      expect(defaultResult.stderr).toBe("");
      expect(defaultResult.stdout).toContain(`Built deck: ${defaultOutput}`);
      expect(defaultResult.stdout).toContain("Slides: 1");

      const defaultBytes = new Uint8Array(await readFile(defaultOutput));
      const defaultFile = await validatePptx(defaultBytes);
      expect(defaultFile.errors).toEqual([]);
      expect(defaultFile.slides).toHaveLength(1);
      expect(defaultFile.slides[0]?.objects[0]).toMatchObject({
        name: "From source dependency",
        text: { text: "From source dependency" },
      });

      const namedStructuredOut = join(directory, "nested", "overflow.pptx");
      const namedStructured = await runInMemory(
        [
          "build",
          source,
          "-e",
          "OverflowDeck",
          "-o",
          namedStructuredOut,
          "--embed-fonts",
          "subset",
          "--output",
          "structured",
        ],
        directory,
      );
      expect(namedStructured.code).toBe(0);
      expect(namedStructured.stderr).toBe("");
      expect(JSON.parse(namedStructured.stdout)).toMatchObject({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "data",
        variant: "built",
        data: {
          output: namedStructuredOut,
          slides: 1,
          warnings: [
            {
              path: "deck/slide[overflow]/shape[overflow-shape]",
              code: "overflow",
            },
          ],
        },
      });

      const overflowFile = await validatePptx(
        new Uint8Array(await readFile(namedStructuredOut)),
      );
      expect(overflowFile.errors).toEqual([]);
      expect(overflowFile.slides).toHaveLength(1);
      expect(overflowFile.slides[0]?.objects).toHaveLength(1);

      const textWarning = await runInMemory(
        [
          "build",
          source,
          "--export",
          "OverflowDeck",
          "--out",
          join(directory, "overflow-text.pptx"),
          "--embed-fonts",
          "full",
        ],
        directory,
      );
      expect(textWarning.code).toBe(0);
      expect(textWarning.stdout).toContain(
        "Warning deck/slide[overflow]/shape[overflow-shape] [overflow]:",
      );
      expect(textWarning.stdout).toContain("Built deck:");
      expect(textWarning.stdout).toContain("Slides: 1");

      const baselineOut = join(directory, "baseline.pptx");
      const baseline = runNode(
        [
          cli,
          "build",
          source,
          "--target",
          "baseline-only",
          "--out",
          baselineOut,
        ],
        directory,
      );
      expect(baseline.code).toBe(0);
      const baselineFile = await validatePptx(
        new Uint8Array(await readFile(baselineOut)),
      );
      expect(baselineFile.errors).toEqual([]);
      expect(baselineFile.extensions).toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("returns usage status 2 for missing subcommand or invalid arguments and rejects overwriting source", async () => {
    const missingSubcommand = await runInMemory(
      ["talk.tsx"],
      import.meta.dirname,
    );
    expect(missingSubcommand.code).toBe(2);
    expect(JSON.parse(missingSubcommand.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "pptxdk",
      kind: "usageFailure",
      usage: "pptxdk <command>",
      helpArgv: ["pptxdk", "--help"],
      issues: [{ code: "unknownCommand", command: "talk.tsx", position: 0 }],
    });

    const missingEntry = await runInMemory(["build"], import.meta.dirname);
    expect(missingEntry.code).toBe(2);
    expect(JSON.parse(missingEntry.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      helpArgv: ["pptxdk", "build", "--help"],
      issues: [{ code: "missingRequiredField", field: "entry" }],
    });

    const unknownOption = await runInMemory(
      ["build", "--unknown"],
      import.meta.dirname,
    );
    expect(unknownOption.code).toBe(2);
    expect(JSON.parse(unknownOption.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [{ code: "unknownOption", option: "--unknown", position: 1 }],
    });

    const missingValue = await runInMemory(
      ["build", "--embed-fonts"],
      import.meta.dirname,
    );
    expect(missingValue.code).toBe(2);
    expect(JSON.parse(missingValue.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [
        {
          code: "missingOptionValue",
          field: "embedFonts",
          option: "--embed-fonts",
          position: 1,
        },
      ],
    });

    const invalidChoice = await runInMemory(
      ["build", fixture("source deck.tsx"), "--embed-fonts", "invalid"],
      import.meta.dirname,
    );
    expect(invalidChoice.code).toBe(2);
    expect(JSON.parse(invalidChoice.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [
        {
          code: "invalidFieldChoice",
          field: "embedFonts",
          value: "invalid",
          choices: ["full", "subset"],
        },
      ],
    });

    const invalidTarget = await runInMemory(
      ["build", fixture("source deck.tsx"), "--target", "invalid"],
      import.meta.dirname,
    );
    expect(invalidTarget.code).toBe(2);
    expect(JSON.parse(invalidTarget.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [
        {
          code: "invalidFieldChoice",
          field: "target",
          value: "invalid",
          choices: ["transitional", "baseline-only"],
        },
      ],
    });

    const invalidOutputMode = await runInMemory(
      ["build", fixture("source deck.tsx"), "--output", "xml"],
      import.meta.dirname,
    );
    expect(invalidOutputMode.code).toBe(2);
    expect(JSON.parse(invalidOutputMode.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [{ code: "invalidOutputMode", option: "--output" }],
    });

    const invalidSchemaQuery = await runInMemory(
      ["build", "--schema=unknown"],
      import.meta.dirname,
    );
    expect(invalidSchemaQuery.code).toBe(2);
    expect(JSON.parse(invalidSchemaQuery.stderr)).toMatchObject({
      schemaVersion: "3",
      command: "build",
      kind: "usageFailure",
      issues: [{ code: "invalidSchemaQuery" }],
    });

    const directory = await mkdtemp(join(tmpdir(), "pptxdk cli "));
    try {
      const copiedSource = join(directory, "deck.tsx");
      const sourceContent = await readFile(fixture("source deck.tsx"), "utf8");
      await writeFile(copiedSource, sourceContent);
      const linkedSource = join(directory, "deck-link.tsx");
      await symlink(copiedSource, linkedSource);

      const textOverwrite = await runInMemory(
        ["build", copiedSource, "--out", copiedSource],
        directory,
      );
      expect(textOverwrite.code).toBe(2);
      expect(textOverwrite.stdout).toBe("");
      expect(textOverwrite.stderr).toBe(
        "Output path cannot overwrite deck source\n",
      );
      expect(await readFile(copiedSource, "utf8")).toBe(sourceContent);

      const structuredOverwrite = await runInMemory(
        [
          "build",
          copiedSource,
          "--out",
          linkedSource,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structuredOverwrite.code).toBe(2);
      expect(structuredOverwrite.stdout).toBe("");
      expect(JSON.parse(structuredOverwrite.stderr)).toEqual({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "failure",
        variant: "invalidOutput",
        data: {
          entry: copiedSource,
          output: linkedSource,
        },
      });
      expect(await readFile(copiedSource, "utf8")).toBe(sourceContent);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("reports sourceLoadFailed, buildFailed, writeFailed, and cancelled across text and structured outputs without stale artifacts", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk cli "));
    try {
      const missingOut = join(directory, "missing.pptx");
      const missingEntry = join(directory, "does-not-exist.tsx");
      const textSourceFailure = await runInMemory(
        ["build", missingEntry, "--out", missingOut],
        directory,
      );
      expect(textSourceFailure.code).toBe(1);
      expect(textSourceFailure.stdout).toBe("");
      expect(textSourceFailure.stderr).toContain("Failed to load deck source:");
      expect(existsSync(missingOut)).toBe(false);

      const envLoadFailure = spawnSync(
        process.execPath,
        [cli, "build", fixture("load failure.tsx"), "--out", missingOut],
        {
          cwd: directory,
          encoding: "utf8",
          env: { ...process.env, PPTXDK_TEST_LOAD_FAILURE: "1" },
        },
      );
      expect(envLoadFailure.status).toBe(1);
      expect(envLoadFailure.stdout).toBe("");
      expect(envLoadFailure.stderr).toContain("Failed to load deck source:");
      expect(envLoadFailure.stderr).toContain(
        "Failed to load author source module",
      );
      expect(existsSync(missingOut)).toBe(false);

      const structuredSourceFailure = await runInMemory(
        [
          "build",
          fixture("source deck.tsx"),
          "--export",
          "MissingExport",
          "--out",
          missingOut,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structuredSourceFailure.code).toBe(1);
      expect(structuredSourceFailure.stdout).toBe("");
      expect(JSON.parse(structuredSourceFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "failure",
        variant: "sourceLoadFailed",
        data: {
          message: expect.stringContaining("MissingExport"),
        },
      });
      expect(existsSync(missingOut)).toBe(false);

      const brokenOut = join(directory, "never.pptx");
      const textBuildFailure = runNode(
        [cli, "build", fixture("broken deck.tsx"), "--out", brokenOut],
        directory,
      );
      expect(textBuildFailure.code).toBe(1);
      expect(textBuildFailure.stdout).toBe("");
      expect(textBuildFailure.stderr).toContain("Build failed:");
      expect(textBuildFailure.stderr).toContain(
        "deck/slide[broken]/shape[broken-shape] [invalid-coordinate]:",
      );
      expect(existsSync(brokenOut)).toBe(false);

      const structuredBuildFailure = await runInMemory(
        [
          "build",
          fixture("broken deck.tsx"),
          "--out",
          brokenOut,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structuredBuildFailure.code).toBe(1);
      expect(structuredBuildFailure.stdout).toBe("");
      expect(JSON.parse(structuredBuildFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "failure",
        variant: "buildFailed",
        data: {
          diagnostics: [
            {
              path: "deck/slide[broken]/shape[broken-shape]",
              code: "invalid-coordinate",
              message: expect.any(String),
            },
          ],
        },
      });
      expect(existsSync(brokenOut)).toBe(false);

      // Pre-create --out as a directory so rename fails with EISDIR even when running as root.
      const dirAsOutput = join(directory, "occupied-as-dir.pptx");
      await mkdir(dirAsOutput);

      const textWriteFailure = await runInMemory(
        ["build", fixture("source deck.tsx"), "--out", dirAsOutput],
        directory,
      );
      expect(textWriteFailure.code).toBe(1);
      expect(textWriteFailure.stdout).toBe("");
      expect(textWriteFailure.stderr).toContain(
        `Failed to write PPTX output to ${dirAsOutput}:`,
      );
      expect(await readdir(directory)).not.toEqual(
        expect.arrayContaining([expect.stringMatching(/\.tmp$/u)]),
      );

      const structuredWriteFailure = await runInMemory(
        [
          "build",
          fixture("source deck.tsx"),
          "--out",
          dirAsOutput,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structuredWriteFailure.code).toBe(1);
      expect(structuredWriteFailure.stdout).toBe("");
      expect(JSON.parse(structuredWriteFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "failure",
        variant: "writeFailed",
        data: {
          output: dirAsOutput,
          message: expect.any(String),
        },
      });

      const preAborted = new AbortController();
      preAborted.abort(new Error("aborted before run"));
      const cancelledOut = join(directory, "cancelled.pptx");
      const structuredCancelled = await runInMemory(
        [
          "build",
          fixture("source deck.tsx"),
          "--out",
          cancelledOut,
          "--output",
          "structured",
        ],
        directory,
        preAborted.signal,
      );
      expect(structuredCancelled.code).toBe(130);
      expect(structuredCancelled.stdout).toBe("");
      expect(JSON.parse(structuredCancelled.stderr)).toEqual({
        schemaVersion: "3",
        command: "build",
        outputMode: "structured",
        kind: "failure",
        variant: "cancelled",
        data: {},
      });
      expect(existsSync(cancelledOut)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("connects SIGINT during build and source import to cancelled exit status 130", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk cli "));
    try {
      const renderOutput = join(directory, "cancelled-render.pptx");
      const renderChild = spawn(
        process.execPath,
        [cli, "build", fixture("cancelled deck.tsx"), "--out", renderOutput],
        {
          cwd: directory,
          stdio: ["ignore", "pipe", "pipe"],
        },
      );
      let renderStdout = "";
      let renderStderr = "";
      let renderInterrupted = false;
      renderChild.stdout.setEncoding("utf8");
      renderChild.stderr.setEncoding("utf8");
      renderChild.stdout.on("data", (chunk: string) => {
        renderStdout += chunk;
        if (
          !renderInterrupted &&
          renderStdout.includes("__pptxdk_cancel_ready__")
        ) {
          renderInterrupted = true;
          renderChild.kill("SIGINT");
        }
      });
      renderChild.stderr.on("data", (chunk: string) => {
        renderStderr += chunk;
      });
      const renderCode = await waitForClose(renderChild);

      expect(renderCode).toBe(130);
      expect(renderStderr).toBe("Deck build cancelled\n");
      expect(existsSync(renderOutput)).toBe(false);

      const loadOutput = join(directory, "cancelled-source-load.pptx");
      const loadChild = spawn(
        process.execPath,
        [
          cli,
          "build",
          fixture("source-load-cli-cancelled.tsx"),
          "--out",
          loadOutput,
        ],
        {
          cwd: directory,
          stdio: ["ignore", "pipe", "pipe"],
        },
      );
      let loadStdout = "";
      let loadStderr = "";
      let loadInterrupted = false;
      loadChild.stdout.setEncoding("utf8");
      loadChild.stderr.setEncoding("utf8");
      loadChild.stdout.on("data", (chunk: string) => {
        loadStdout += chunk;
        if (
          !loadInterrupted &&
          loadStdout.includes("__pptxdk_source_load_cancel_ready__")
        ) {
          loadInterrupted = true;
          loadChild.kill("SIGINT");
        }
      });
      loadChild.stderr.on("data", (chunk: string) => {
        loadStderr += chunk;
      });
      const loadCode = await waitForClose(loadChild);

      expect(loadCode).toBe(130);
      expect(loadStderr).toBe("Deck build cancelled\n");
      expect(loadStderr).not.toContain("Failed to load deck source");
      expect(existsSync(loadOutput)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("handles EPIPE on stdout and unexpected stream write failures", async () => {
    const brokenPipeCode = await runCli({
      argv: ["node", "pptxdk", "--help"],
      cwd: import.meta.dirname,
      write: async () => {
        throw Object.assign(new Error("broken pipe"), { code: "EPIPE" });
      },
    });
    expect(brokenPipeCode).toBe(0);

    let diagnosticStderr = "";
    const writeFailureCode = await runCli({
      argv: ["node", "pptxdk", "--help"],
      cwd: import.meta.dirname,
      write: async ({ destination, chunk }) => {
        if (destination === "stdout") {
          throw new Error("disk full");
        }
        diagnosticStderr += chunk;
      },
    });
    expect(writeFailureCode).toBe(70);
    expect(diagnosticStderr).toContain("Command execution failed");
  });

  it("keeps CLI build output identical to buildFromSource and releases source hooks across outcomes", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk cli "));
    try {
      const source = fixture("source deck.tsx");
      const programmaticOut = join(directory, "programmatic.pptx");
      const cliOut = join(directory, "cli.pptx");
      const programmaticScript = [
        'import { pathToFileURL } from "node:url";',
        'import { writeFile } from "node:fs/promises";',
        'import { buildFromSource } from "pptxdk/node";',
        `const result = await buildFromSource(${JSON.stringify(source)}, { exportName: "SourceDeck" });`,
        `await writeFile(${JSON.stringify(programmaticOut)}, result.pptx);`,
        'async function expectHookReleased(path) { const error = await import(pathToFileURL(path).href).then(() => undefined, failure => failure); if (error?.code !== "ERR_UNKNOWN_FILE_EXTENSION") throw new Error(JSON.stringify({ path, error })); }',
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after success.tsx"))});`,
        'process.env.PPTXDK_TEST_LOAD_FAILURE = "1";',
        `const loadFailure = await buildFromSource(${JSON.stringify(fixture("load failure.tsx"))}).then(() => undefined, error => error);`,
        'if (loadFailure?.name !== "SourceLoadError") throw new Error("Source load error was not preserved");',
        "delete process.env.PPTXDK_TEST_LOAD_FAILURE;",
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after load failure.tsx"))});`,
        `const buildFailure = await buildFromSource(${JSON.stringify(fixture("broken deck.tsx"))}).then(() => undefined, error => error);`,
        'if (buildFailure?.name !== "BuildError") throw new Error("Build error was not preserved");',
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after build failure.tsx"))});`,
        "const controller = new AbortController();",
        'const reason = new Error("Programmatic source build cancelled");',
        "globalThis.__pptxdkCancellation = { controller, reason };",
        `const cancellation = await buildFromSource(${JSON.stringify(fixture("cancelled deck.tsx"))}, { signal: controller.signal }).then(() => undefined, error => error);`,
        'if (cancellation !== reason) throw new Error("Cancellation reason was not propagated from public build");',
        "delete globalThis.__pptxdkCancellation;",
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after cancellation.tsx"))});`,
        `const independent = await buildFromSource(${JSON.stringify(fixture("independent deck.tsx"))});`,
        "process.stdout.write(JSON.stringify({ sourceSlide: result.report.slides[0]?.id, independentSlide: independent.report.slides[0]?.id }));",
      ].join("\n");
      const programmatic = runNode(
        ["--input-type=module", "-e", programmaticScript],
        import.meta.dirname,
      );
      expect(programmatic.code).toBe(0);
      expect(programmatic.stdout).toContain('"sourceSlide":"source"');
      expect(programmatic.stdout).toContain('"independentSlide":"independent"');

      const cliResult = runNode(
        [cli, "build", source, "--export", "SourceDeck", "--out", cliOut],
        directory,
      );
      expect(cliResult.code).toBe(0);
      expect(await readFile(cliOut)).toEqual(await readFile(programmaticOut));
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("cancellation during source import terminates the call immediately and isolates concurrent loading", () => {
    const programmaticScript = [
      'import { buildFromSource } from "pptxdk/node";',
      'import { pathToFileURL } from "node:url";',
      `const pendingEntry = ${JSON.stringify(fixture("source-load-pending.tsx"))};`,
      `const concurrentEntry = ${JSON.stringify(fixture("source-load-after-cancel.tsx"))};`,
      `const unloadedEntry = ${JSON.stringify(fixture("unloaded after cancellation.tsx"))};`,
      `const independentEntry = ${JSON.stringify(fixture("independent deck.tsx"))};`,
      'async function expectHookReleased(path) { const error = await import(pathToFileURL(path).href).then(() => undefined, failure => failure); if (error?.code !== "ERR_UNKNOWN_FILE_EXTENSION") throw new Error(JSON.stringify({ path, error })); }',
      "const state = { controller: new AbortController(), started: false, concurrentStarted: false, rejectPending: undefined, allowConcurrentImport: undefined };",
      "globalThis.__pptxdkSourceLoadState = state;",
      "const preController = new AbortController();",
      'const preReason = new Error("Cancelled before source load");',
      "preController.abort(preReason);",
      "const preCancelled = await buildFromSource(pendingEntry, { signal: preController.signal }).then(() => undefined, error => error);",
      'if (preCancelled !== preReason || state.started) throw new Error("Pre-cancellation executed source entrypoint or lost cancellation reason");',
      'const reason = new Error("Cancelled during source import");',
      "const unhandled = [];",
      "const collectUnhandled = error => unhandled.push(error);",
      'process.on("unhandledRejection", collectUnhandled);',
      'const first = buildFromSource(pendingEntry, { signal: state.controller.signal }).then(value => ({ kind: "built", value }), error => ({ kind: "failed", error }));',
      "while (!state.started) await new Promise(setImmediate);",
      "const second = buildFromSource(concurrentEntry);",
      "while (!state.concurrentStarted) await new Promise(setImmediate);",
      "state.controller.abort(reason);",
      'const firstOutcome = await Promise.race([first, new Promise(resolve => setImmediate(() => resolve({ kind: "pending" })))]);',
      'if (firstOutcome.kind === "pending") { state.rejectPending(reason); await first; throw new Error("Source import cancellation did not terminate call in time"); }',
      'if (firstOutcome.kind !== "failed" || firstOutcome.error !== reason) throw new Error("Source import cancellation did not preserve original reason");',
      "state.rejectPending(reason);",
      'if (state.allowConcurrentImport === undefined) throw new Error("Concurrent entrypoint has not reached deferred import stage");',
      "state.allowConcurrentImport();",
      "const secondResult = await second;",
      'if (secondResult.report.slides[0]?.id !== "source-after-cancellation") throw new Error("Cancellation of A affected deferred TSX import of concurrent B");',
      "await new Promise(resolve => setTimeout(resolve, 0));",
      'process.removeListener("unhandledRejection", collectUnhandled);',
      "if (unhandled.length > 0) throw new Error(`Late module rejection was not handled: ${String(unhandled[0])}`);",
      "await expectHookReleased(unloadedEntry);",
      "const independent = await buildFromSource(independentEntry);",
      'if (independent.report.slides[0]?.id !== "independent") throw new Error("Subsequent source build failed after cancellation");',
    ].join("\n");
    const programmatic = runNode(
      ["--input-type=module", "-e", programmaticScript],
      import.meta.dirname,
    );

    expect(
      programmatic.code,
      `${programmatic.stderr}\n${programmatic.stdout}`,
    ).toBe(0);
  });
});
