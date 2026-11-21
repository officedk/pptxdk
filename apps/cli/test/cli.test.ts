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
      children: [{ command: "build" }, { command: "preview" }],
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

describe("pptxdk CLI preview", () => {
  it("exposes preview in CLI help and schema contract and writes all pages by default or to --out-dir", async () => {
    const rootHelp = await runInMemory(["--help"], import.meta.dirname);
    expect(rootHelp.code).toBe(0);
    expect(rootHelp.stdout).toContain("build");
    expect(rootHelp.stdout).toContain("preview");

    const rootSchema = await runInMemory(["--schema"], import.meta.dirname);
    expect(rootSchema.code).toBe(0);
    const parsedRootSchema = JSON.parse(rootSchema.stdout) as {
      readonly children: ReadonlyArray<{ readonly command: string }>;
    };
    expect(parsedRootSchema.children.map((child) => child.command)).toEqual([
      "build",
      "preview",
    ]);

    const previewHelp = await runInMemory(
      ["preview", "-h"],
      import.meta.dirname,
    );
    expect(previewHelp.code).toBe(0);
    expect(previewHelp.stderr).toBe("");
    expect(previewHelp.stdout).toContain(
      "pptxdk preview [--output <mode>] <entry>",
    );
    expect(previewHelp.stdout).toContain("--out-dir");
    expect(previewHelp.stdout).toContain("--export");
    expect(previewHelp.stdout).toContain("--output");

    const previewSchema = await runInMemory(
      ["preview", "--schema"],
      import.meta.dirname,
    );
    expect(previewSchema.code).toBe(0);
    expect(JSON.parse(previewSchema.stdout)).toMatchObject({
      kind: "summary",
      schemaVersion: "5",
      command: "preview",
    });

    const previewOutputSchema = await runInMemory(
      ["preview", "--schema=success:structured"],
      import.meta.dirname,
    );
    expect(previewOutputSchema.code).toBe(0);
    expect(JSON.parse(previewOutputSchema.stdout)).toMatchObject({
      kind: "success",
      schemaVersion: "5",
      command: "preview",
      mode: "structured",
      success: {
        kind: "data",
        variants: {
          previewed: {
            exitCode: 0,
            outputSchema: {
              required: ["outputDir", "pages", "warnings"],
            },
          },
        },
      },
    });

    const previewFailureSchema = await runInMemory(
      ["preview", "--schema=failures:structured"],
      import.meta.dirname,
    );
    expect(previewFailureSchema.code).toBe(0);
    expect(
      Object.keys(
        (
          JSON.parse(previewFailureSchema.stdout) as {
            readonly failures: Record<string, unknown>;
          }
        ).failures,
      ),
    ).toEqual([
      "invalidOutput",
      "pageOutOfRange",
      "sourceLoadFailed",
      "buildFailed",
      "previewFailed",
      "writeFailed",
      "cancelled",
    ]);

    const directory = await mkdtemp(join(tmpdir(), "pptxdk preview "));
    try {
      const source = fixture("preview deck.tsx");
      const defaultOutDir = join(directory, "preview deck-preview");

      const defaultResult = runNode([cli, "preview", source], directory);
      expect(defaultResult.code).toBe(0);
      expect(defaultResult.stderr).toBe("");
      expect(defaultResult.stdout).toContain(
        `Previewed 1 page(s) in ${defaultOutDir}`,
      );
      expect(defaultResult.stdout).toContain(
        `Page 1 (preview): ${join(defaultOutDir, "page-01.png")}`,
      );

      const defaultFiles = await readdir(defaultOutDir);
      expect(defaultFiles).toEqual(["page-01.png"]);

      const { previewFromSource } = await import("pptxdk/node");
      const formal = await previewFromSource(source);
      expect(formal.pages[0]?.png).toBeInstanceOf(Uint8Array);
      expect(await readFile(join(defaultOutDir, "page-01.png"))).toEqual(
        Buffer.from(formal.pages[0]!.png!),
      );

      const customOutDir = join(directory, "custom", "frames");
      const customResult = await runInMemory(
        ["preview", source, "-o", customOutDir, "--output", "structured"],
        directory,
      );
      expect(customResult.code).toBe(0);
      expect(customResult.stderr).toBe("");
      expect(JSON.parse(customResult.stdout)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "data",
        variant: "previewed",
        data: {
          outputDir: customOutDir,
          pages: [
            {
              page: 1,
              id: "preview",
              output: join(customOutDir, "page-01.png"),
            },
          ],
          warnings: [],
        },
      });
      expect(await readdir(customOutDir)).toEqual(["page-01.png"]);
      expect(await readFile(join(customOutDir, "page-01.png"))).toEqual(
        Buffer.from(formal.pages[0]!.png!),
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("builds and previews async vanilla decks, JSX elements, and React Hooks components via --export/-e", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk dual "));
    try {
      const source = fixture("preview deck.tsx");
      const { previewFromSource } = await import("pptxdk/node");
      const cases = [
        {
          exportName: "AsyncVanillaDeck",
          slideId: "async-vanilla",
          shapeName: "Async Vanilla Shape",
          flag: "--export",
        },
        {
          exportName: "JsxElementDeck",
          slideId: "jsx-element",
          shapeName: "JSX Element Shape",
          flag: "-e",
        },
        {
          exportName: "HookComponent",
          slideId: "hook-component",
          shapeName: "Hook: committed",
          flag: "--export",
        },
      ] as const;

      for (const item of cases) {
        const pptxOut = join(directory, `${item.exportName}.pptx`);
        const buildResult = runNode(
          [cli, "build", source, item.flag, item.exportName, "--out", pptxOut],
          directory,
        );
        expect(buildResult.code).toBe(0);
        const validated = await validatePptx(
          new Uint8Array(await readFile(pptxOut)),
        );
        expect(validated.errors).toEqual([]);
        expect(validated.slides).toHaveLength(1);
        expect(validated.slides[0]?.objects[0]).toMatchObject({
          name: item.shapeName,
        });

        const previewDir = join(directory, `${item.exportName}-preview`);
        const previewResult = runNode(
          [
            cli,
            "preview",
            source,
            item.flag,
            item.exportName,
            "--out-dir",
            previewDir,
          ],
          directory,
        );
        expect(previewResult.code).toBe(0);
        expect(previewResult.stderr).toBe("");
        expect(previewResult.stdout).toContain(
          `Page 1 (${item.slideId}): ${join(previewDir, "page-01.png")}`,
        );

        const formal = await previewFromSource(source, {
          exportName: item.exportName,
        });
        expect(await readdir(previewDir)).toEqual(["page-01.png"]);
        expect(await readFile(join(previewDir, "page-01.png"))).toEqual(
          Buffer.from(formal.pages[0]!.png!),
        );
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("pads page filenames based on total slide count and preserves slide ids with spaces and slashes", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk paging "));
    try {
      const source = fixture("multi-slide.ts");

      const specialDir = join(directory, "special");
      const specialResult = await runInMemory(
        [
          "preview",
          source,
          "--export",
          "SpecialIdDeck",
          "--out-dir",
          specialDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(specialResult.code).toBe(0);
      expect(specialResult.stderr).toBe("");
      expect(JSON.parse(specialResult.stdout)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "data",
        variant: "previewed",
        data: {
          outputDir: specialDir,
          pages: [
            {
              page: 1,
              id: "sec 1/intro",
              output: join(specialDir, "page-01.png"),
            },
            {
              page: 2,
              id: "sec 2/summary",
              output: join(specialDir, "page-02.png"),
            },
          ],
          warnings: [],
        },
      });
      expect((await readdir(specialDir)).toSorted()).toEqual([
        "page-01.png",
        "page-02.png",
      ]);

      const hundredDir = join(directory, "hundred");
      const hundredResult = await runInMemory(
        [
          "preview",
          source,
          "-e",
          "HundredPageDeck",
          "-o",
          hundredDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(hundredResult.code).toBe(0);
      const parsedHundred = JSON.parse(hundredResult.stdout) as {
        readonly data: {
          readonly pages: ReadonlyArray<{
            readonly page: number;
            readonly id: string;
            readonly output: string;
          }>;
        };
      };
      expect(parsedHundred.data.pages).toHaveLength(100);
      expect(parsedHundred.data.pages[0]).toEqual({
        page: 1,
        id: "slide-1",
        output: join(hundredDir, "page-001.png"),
      });
      expect(parsedHundred.data.pages[99]).toEqual({
        page: 100,
        id: "slide-100",
        output: join(hundredDir, "page-100.png"),
      });
      const files = (await readdir(hundredDir)).toSorted();
      expect(files).toHaveLength(100);
      expect(files[0]).toBe("page-001.png");
      expect(files[99]).toBe("page-100.png");
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("writes successful sibling pages atomically, preserves stale failed PNGs and unrelated files, and reports previewFailed", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk partial "));
    try {
      const source = fixture("partial-failure-preview.tsx");
      const outDir = join(directory, "preview-out");
      await mkdir(outDir, { recursive: true });
      await writeFile(join(outDir, "page-02.png"), "stale-png");
      await writeFile(join(outDir, "notes.txt"), "keep-me");

      const structured = await runInMemory(
        ["preview", source, "--out-dir", outDir, "--output", "structured"],
        directory,
      );
      expect(structured.code).toBe(1);
      expect(structured.stdout).toBe("");
      expect(JSON.parse(structured.stderr)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "previewFailed",
        data: {
          outputDir: outDir,
          pages: [
            {
              page: 1,
              id: "warn-ok",
              output: join(outDir, "page-01.png"),
            },
            {
              page: 3,
              id: "good-sibling",
              output: join(outDir, "page-03.png"),
            },
          ],
          warnings: [
            {
              path: "deck/slide[warn-ok]/shape[overflow-shape]",
              code: "overflow",
              message: expect.any(String),
            },
          ],
          diagnostics: [
            {
              page: 2,
              path: "deck/slide[bad-image]/image[corrupt-img]",
              message: expect.stringContaining("PNG"),
            },
          ],
        },
      });

      const { previewFromSource } = await import("pptxdk/node");
      const formal = await previewFromSource(source);
      expect(await readFile(join(outDir, "page-01.png"))).toEqual(
        Buffer.from(formal.pages[0]!.png!),
      );
      expect(await readFile(join(outDir, "page-03.png"))).toEqual(
        Buffer.from(formal.pages[2]!.png!),
      );
      expect(await readFile(join(outDir, "page-02.png"), "utf8")).toBe(
        "stale-png",
      );
      expect(await readFile(join(outDir, "notes.txt"), "utf8")).toBe("keep-me");
      expect((await readdir(outDir)).toSorted()).toEqual([
        "notes.txt",
        "page-01.png",
        "page-02.png",
        "page-03.png",
      ]);

      const textResult = runNode(
        [cli, "preview", source, "--out-dir", outDir],
        directory,
      );
      expect(textResult.code).toBe(1);
      expect(textResult.stdout).toBe("");
      expect(textResult.stderr).toContain(
        "Warning deck/slide[warn-ok]/shape[overflow-shape] [overflow]:",
      );
      expect(textResult.stderr).toContain("Preview failed:");
      expect(textResult.stderr).toContain(
        "Page 2 deck/slide[bad-image]/image[corrupt-img]:",
      );
      expect(textResult.stderr).toContain("Written pages:");
      expect(textResult.stderr).toContain(
        `Page 1 (warn-ok): ${join(outDir, "page-01.png")}`,
      );
      expect(textResult.stderr).toContain(
        `Page 3 (good-sibling): ${join(outDir, "page-03.png")}`,
      );
      expect(textResult.stderr).not.toContain("Page 2 (bad-image)");
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("rejects non-directory outDir with invalidOutput, avoids premature mkdir on failures, and reports writeFailed", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk preview errors "));
    try {
      const missingEntryCheck = await runInMemory(["preview"], directory);
      expect(missingEntryCheck.code).toBe(2);
      expect(JSON.parse(missingEntryCheck.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
        kind: "usageFailure",
        helpArgv: ["pptxdk", "preview", "--help"],
        issues: [{ code: "missingRequiredField", field: "entry" }],
      });

      const source = fixture("preview deck.tsx");
      const regularFile = join(directory, "existing-file.txt");
      await writeFile(regularFile, "not a directory");
      const danglingLink = join(directory, "dangling-link");
      await symlink(join(directory, "missing-target"), danglingLink);
      const loopLink = join(directory, "loop-dir");
      await symlink("loop-dir", loopLink);

      const fileOutDir = await runInMemory(
        ["preview", source, "--out-dir", regularFile],
        directory,
      );
      expect(fileOutDir.code).toBe(2);
      expect(fileOutDir.stdout).toBe("");
      expect(fileOutDir.stderr).toBe(
        `Output directory cannot be a non-directory path: ${regularFile}\n`,
      );
      expect(await readFile(regularFile, "utf8")).toBe("not a directory");

      const loopTextOutDir = await runInMemory(
        ["preview", source, "--out-dir", loopLink],
        directory,
      );
      expect(loopTextOutDir.code).toBe(2);
      expect(loopTextOutDir.stdout).toBe("");
      expect(loopTextOutDir.stderr).toBe(
        `Output directory cannot be a non-directory path: ${loopLink}\n`,
      );

      for (const invalidTarget of [
        danglingLink,
        loopLink,
        source,
        join(regularFile, "child"),
      ]) {
        const result = await runInMemory(
          [
            "preview",
            source,
            "--out-dir",
            invalidTarget,
            "--output",
            "structured",
          ],
          directory,
        );
        expect(result.code).toBe(2);
        expect(result.stdout).toBe("");
        expect(JSON.parse(result.stderr)).toEqual({
          schemaVersion: "3",
          command: "preview",
          outputMode: "structured",
          kind: "failure",
          variant: "invalidOutput",
          data: {
            entry: source,
            output: invalidTarget,
          },
        });
      }

      const tooLongOutDir = join(directory, "x".repeat(300));
      const tooLongResult = await runInMemory(
        [
          "preview",
          source,
          "--out-dir",
          tooLongOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(tooLongResult.code).toBe(1);
      expect(tooLongResult.stdout).toBe("");
      expect(JSON.parse(tooLongResult.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "writeFailed",
        data: {
          output: tooLongOutDir,
        },
      });

      const missingSourceOutDir = join(directory, "no-dir-on-source-fail");
      const sourceFailure = await runInMemory(
        [
          "preview",
          source,
          "--export",
          "MissingExport",
          "--out-dir",
          missingSourceOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(sourceFailure.code).toBe(1);
      expect(sourceFailure.stdout).toBe("");
      expect(JSON.parse(sourceFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "sourceLoadFailed",
        data: { message: expect.stringContaining("MissingExport") },
      });
      expect(existsSync(missingSourceOutDir)).toBe(false);

      const brokenOutDir = join(directory, "no-dir-on-build-fail");
      const buildFailure = await runInMemory(
        [
          "preview",
          fixture("broken deck.tsx"),
          "--out-dir",
          brokenOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(buildFailure.code).toBe(1);
      expect(buildFailure.stdout).toBe("");
      expect(JSON.parse(buildFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
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
      expect(existsSync(brokenOutDir)).toBe(false);

      // Pre-create page-01.png as a subdirectory inside a valid outDir so rename fails with EISDIR even as root.
      const writeFailOutDir = join(directory, "write-fail-out");
      const blockingPageDir = join(writeFailOutDir, "page-01.png");
      await mkdir(blockingPageDir, { recursive: true });

      const textWriteFailure = await runInMemory(
        ["preview", source, "--out-dir", writeFailOutDir],
        directory,
      );
      expect(textWriteFailure.code).toBe(1);
      expect(textWriteFailure.stdout).toBe("");
      expect(textWriteFailure.stderr).toContain(
        `Failed to write preview image to ${blockingPageDir}:`,
      );
      expect(await readdir(writeFailOutDir)).toEqual(["page-01.png"]);

      const structuredWriteFailure = await runInMemory(
        [
          "preview",
          source,
          "--out-dir",
          writeFailOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structuredWriteFailure.code).toBe(1);
      expect(structuredWriteFailure.stdout).toBe("");
      expect(JSON.parse(structuredWriteFailure.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "writeFailed",
        data: {
          output: blockingPageDir,
          message: expect.any(String),
        },
      });
      expect(await readdir(writeFailOutDir)).toEqual(["page-01.png"]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("connects pre-aborted signals and SIGINT during preview to cancelled exit status 130 without creating outDir", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk preview cancel "));
    try {
      const preAborted = new AbortController();
      preAborted.abort(new Error("aborted before preview"));
      const preAbortOutDir = join(directory, "pre-abort-preview");
      const structuredCancelled = await runInMemory(
        [
          "preview",
          fixture("preview deck.tsx"),
          "--out-dir",
          preAbortOutDir,
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
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "cancelled",
        data: {},
      });
      expect(existsSync(preAbortOutDir)).toBe(false);

      const renderOutDir = join(directory, "cancelled-render-preview");
      const renderChild = spawn(
        process.execPath,
        [
          cli,
          "preview",
          fixture("cancelled deck.tsx"),
          "--out-dir",
          renderOutDir,
        ],
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
      expect(renderStderr).toBe("Deck preview cancelled\n");
      expect(existsSync(renderOutDir)).toBe(false);

      const loadOutDir = join(directory, "cancelled-load-preview");
      const loadChild = spawn(
        process.execPath,
        [
          cli,
          "preview",
          fixture("source-load-cli-cancelled.tsx"),
          "--out-dir",
          loadOutDir,
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
      expect(loadStderr).toBe("Deck preview cancelled\n");
      expect(loadStderr).not.toContain("Failed to load deck source");
      expect(existsSync(loadOutDir)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("selects deduplicated ascending pages via --pages/-p and documents the option in help and contract schema", async () => {
    const previewHelp = await runInMemory(
      ["preview", "-h"],
      import.meta.dirname,
    );
    expect(previewHelp.code).toBe(0);
    expect(previewHelp.stdout).toContain("--pages");
    expect(previewHelp.stdout).toContain("-p");

    const previewInputSchema = await runInMemory(
      ["preview", "--schema=input"],
      import.meta.dirname,
    );
    expect(previewInputSchema.code).toBe(0);
    expect(JSON.parse(previewInputSchema.stdout)).toMatchObject({
      kind: "input",
      schemaVersion: "5",
      command: "preview",
      fields: expect.arrayContaining([
        expect.objectContaining({
          kind: "valueOption",
          longOption: "--pages",
          shortAlias: "-p",
          key: "pages",
          required: false,
        }),
      ]),
      input: {
        inputSchema: {
          properties: {
            pages: { type: "string", minLength: 1 },
          },
        },
        outputSchema: {
          properties: {
            pages: {
              type: "array",
              minItems: 1,
              items: { type: "integer", exclusiveMinimum: 0 },
            },
          },
        },
      },
    });

    const directory = await mkdtemp(join(tmpdir(), "pptxdk select-pages "));
    try {
      const source = fixture("multi-slide.ts");
      const outDir = join(directory, "selected");

      const result = await runInMemory(
        [
          "preview",
          source,
          "-e",
          "SixPageDeck",
          "-p",
          "5,3-5,1",
          "-o",
          outDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(result.code).toBe(0);
      expect(result.stderr).toBe("");
      expect(JSON.parse(result.stdout)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "data",
        variant: "previewed",
        data: {
          outputDir: outDir,
          pages: [
            { page: 1, id: "six-1", output: join(outDir, "page-01.png") },
            { page: 3, id: "six-3", output: join(outDir, "page-03.png") },
            { page: 4, id: "six-4", output: join(outDir, "page-04.png") },
            { page: 5, id: "six-5", output: join(outDir, "page-05.png") },
          ],
          warnings: [],
        },
      });
      expect((await readdir(outDir)).toSorted()).toEqual([
        "page-01.png",
        "page-03.png",
        "page-04.png",
        "page-05.png",
      ]);

      const { previewFromSource } = await import("pptxdk/node");
      const formal = await previewFromSource(source, {
        exportName: "SixPageDeck",
        pages: [1, 3, 4, 5],
      });
      for (const pageResult of formal.pages) {
        expect(
          await readFile(
            join(
              outDir,
              `page-${String(pageResult.page).padStart(2, "0")}.png`,
            ),
          ),
        ).toEqual(Buffer.from(pageResult.png!));
      }

      const cliOutDir = join(directory, "cli-selected");
      const cliResult = runNode(
        [
          cli,
          "preview",
          source,
          "--export",
          "SixPageDeck",
          "--pages",
          "5,3-5,1",
          "--out-dir",
          cliOutDir,
        ],
        directory,
      );
      expect(cliResult.code).toBe(0);
      expect(cliResult.stderr).toBe("");
      expect(cliResult.stdout).toContain(`Previewed 4 page(s) in ${cliOutDir}`);
      expect((await readdir(cliOutDir)).toSorted()).toEqual([
        "page-01.png",
        "page-03.png",
        "page-04.png",
        "page-05.png",
      ]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("rejects invalid --pages syntax at validation time with exit code 2 without loading source or touching outDir", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk invalid-pages "));
    try {
      const brokenSyntaxEntry = join(directory, "broken-syntax.tsx");
      await writeFile(brokenSyntaxEntry, "export default <Deck << broken;");
      const missingEntry = join(directory, "does-not-exist.tsx");

      const preservedDir = join(directory, "preserved-out");
      await mkdir(preservedDir, { recursive: true });
      await writeFile(join(preservedDir, "page-01.png"), "existing-page-1");
      await writeFile(join(preservedDir, "manifest.json"), '{"keep":true}');

      const nonExistentOutDir = join(directory, "never-created-out");

      const invalidSpecs = [
        "",
        "0",
        "00",
        "5-3",
        "0-2",
        "abc",
        "1.5",
        "-1",
        "1,",
        ",1",
        "1,,2",
        "1, 2",
      ] as const;

      for (const spec of invalidSpecs) {
        const result = await runInMemory(
          ["preview", missingEntry, "-p", spec, "--out-dir", nonExistentOutDir],
          directory,
        );
        expect(result.code).toBe(2);
        expect(result.stdout).toBe("");
        expect(JSON.parse(result.stderr)).toMatchObject({
          schemaVersion: "3",
          command: "preview",
          kind: "usageFailure",
          helpArgv: ["pptxdk", "preview", "--help"],
        });
        expect(existsSync(nonExistentOutDir)).toBe(false);
      }

      const syntaxCases = [
        { entry: missingEntry, spec: "0" },
        { entry: brokenSyntaxEntry, spec: "5-3" },
      ] as const;

      for (const item of syntaxCases) {
        const cliResult = runNode(
          [cli, "preview", item.entry, "-p", item.spec, "-o", preservedDir],
          directory,
        );
        expect(cliResult.code).toBe(2);
        expect(cliResult.stdout).toBe("");
        expect(JSON.parse(cliResult.stderr)).toMatchObject({
          schemaVersion: "3",
          command: "preview",
          kind: "usageFailure",
          helpArgv: ["pptxdk", "preview", "--help"],
        });
      }

      expect((await readdir(preservedDir)).toSorted()).toEqual([
        "manifest.json",
        "page-01.png",
      ]);
      expect(await readFile(join(preservedDir, "page-01.png"), "utf8")).toBe(
        "existing-page-1",
      );
      expect(await readFile(join(preservedDir, "manifest.json"), "utf8")).toBe(
        '{"keep":true}',
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("zero-pads selected page 3 of a 120-page deck to page-003.png and preserves unrelated and unselected files", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk 120-pages "));
    try {
      const source = fixture("multi-slide.ts");
      const outDir = join(directory, "preview-120");
      await mkdir(outDir, { recursive: true });

      const page1Bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
      const page120Bytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 9, 8, 7]);
      const manifestBytes = Buffer.from('{"existing":true}\n', "utf8");
      await writeFile(join(outDir, "page-001.png"), page1Bytes);
      await writeFile(join(outDir, "page-120.png"), page120Bytes);
      await writeFile(join(outDir, "manifest.json"), manifestBytes);

      const result = runNode(
        [
          cli,
          "preview",
          source,
          "-e",
          "OneTwentyPageDeck",
          "-p",
          "3",
          "-o",
          outDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(result.code).toBe(0);
      expect(result.stderr).toBe("");
      expect(JSON.parse(result.stdout)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "data",
        variant: "previewed",
        data: {
          outputDir: outDir,
          pages: [
            {
              page: 3,
              id: "deck-120-slide-3",
              output: join(outDir, "page-003.png"),
            },
          ],
          warnings: [],
        },
      });

      const { previewFromSource } = await import("pptxdk/node");
      const formal = await previewFromSource(source, {
        exportName: "OneTwentyPageDeck",
        pages: [3],
      });
      expect(formal.pages).toHaveLength(1);
      expect(formal.pages[0]?.page).toBe(3);
      expect(formal.pages[0]?.id).toBe("deck-120-slide-3");
      expect(await readFile(join(outDir, "page-003.png"))).toEqual(
        Buffer.from(formal.pages[0]!.png!),
      );

      expect((await readdir(outDir)).toSorted()).toEqual([
        "manifest.json",
        "page-001.png",
        "page-003.png",
        "page-120.png",
      ]);
      expect(await readFile(join(outDir, "page-001.png"))).toEqual(page1Bytes);
      expect(await readFile(join(outDir, "page-120.png"))).toEqual(
        page120Bytes,
      );
      expect(await readFile(join(outDir, "manifest.json"))).toEqual(
        manifestBytes,
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("reports pageOutOfRange with exit code 2 and zero writes when any requested page exceeds totalSlides", async () => {
    const failureSchema = await runInMemory(
      ["preview", "--schema=failures:structured"],
      import.meta.dirname,
    );
    expect(failureSchema.code).toBe(0);
    expect(JSON.parse(failureSchema.stdout)).toMatchObject({
      kind: "failures",
      command: "preview",
      failures: {
        pageOutOfRange: {
          exitCode: 2,
          outputSchema: {
            required: ["requestedPages", "totalSlides"],
          },
        },
      },
    });

    const directory = await mkdtemp(join(tmpdir(), "pptxdk out-of-range "));
    try {
      const source = fixture("multi-slide.ts");
      const missingOutDir = join(directory, "never-created-range");

      const structured = await runInMemory(
        [
          "preview",
          source,
          "-e",
          "SixPageDeck",
          "-p",
          "1,3,8",
          "-o",
          missingOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structured.code).toBe(2);
      expect(structured.stdout).toBe("");
      expect(JSON.parse(structured.stderr)).toEqual({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "pageOutOfRange",
        data: {
          requestedPages: [1, 3, 8],
          totalSlides: 6,
        },
      });
      expect(existsSync(missingOutDir)).toBe(false);

      const existingOutDir = join(directory, "existing-range-out");
      await mkdir(existingOutDir, { recursive: true });
      const oldPage1 = Buffer.from("old-page-1");
      const oldPage3 = Buffer.from("old-page-3");
      const keepFile = Buffer.from("keep-unrelated");
      await writeFile(join(existingOutDir, "page-01.png"), oldPage1);
      await writeFile(join(existingOutDir, "page-03.png"), oldPage3);
      await writeFile(join(existingOutDir, "keep.txt"), keepFile);

      const textResult = runNode(
        [
          cli,
          "preview",
          source,
          "-e",
          "SixPageDeck",
          "-p",
          "5,1,8,3",
          "-o",
          existingOutDir,
        ],
        directory,
      );
      expect(textResult.code).toBe(2);
      expect(textResult.stdout).toBe("");
      expect(textResult.stderr).toBe(
        "Requested pages (1, 3, 5, 8) exceed total slides (6)\n",
      );
      expect((await readdir(existingOutDir)).toSorted()).toEqual([
        "keep.txt",
        "page-01.png",
        "page-03.png",
      ]);
      expect(await readFile(join(existingOutDir, "page-01.png"))).toEqual(
        oldPage1,
      );
      expect(await readFile(join(existingOutDir, "page-03.png"))).toEqual(
        oldPage3,
      );
      expect(await readFile(join(existingOutDir, "keep.txt"))).toEqual(
        keepFile,
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("runs full-deck solve before page selection so unselected build errors still return buildFailed without writing selected pages", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk solve-first "));
    try {
      const source = fixture("broken deck.tsx");
      const missingOutDir = join(directory, "uncreated-build-fail-out");

      const structured = await runInMemory(
        [
          "preview",
          source,
          "-e",
          "UnselectedPageBrokenDeck",
          "-p",
          "1",
          "-o",
          missingOutDir,
          "--output",
          "structured",
        ],
        directory,
      );
      expect(structured.code).toBe(1);
      expect(structured.stdout).toBe("");
      expect(JSON.parse(structured.stderr)).toMatchObject({
        schemaVersion: "3",
        command: "preview",
        outputMode: "structured",
        kind: "failure",
        variant: "buildFailed",
        data: {
          diagnostics: [
            {
              path: "deck/slide[broken-unselected]/shape[broken-unselected-shape]",
              code: "invalid-coordinate",
              message: expect.any(String),
            },
          ],
        },
      });
      expect(existsSync(missingOutDir)).toBe(false);

      const existingOutDir = join(directory, "existing-build-fail-out");
      await mkdir(existingOutDir, { recursive: true });
      const keepBytes = Buffer.from("keep-me");
      await writeFile(join(existingOutDir, "keep.txt"), keepBytes);

      const textResult = runNode(
        [
          cli,
          "preview",
          source,
          "-e",
          "UnselectedPageBrokenDeck",
          "--pages",
          "1",
          "--out-dir",
          existingOutDir,
        ],
        directory,
      );
      expect(textResult.code).toBe(1);
      expect(textResult.stdout).toBe("");
      expect(textResult.stderr).toContain("Build failed:");
      expect(textResult.stderr).toContain(
        "deck/slide[broken-unselected]/shape[broken-unselected-shape] [invalid-coordinate]:",
      );
      expect(await readdir(existingOutDir)).toEqual(["keep.txt"]);
      expect(await readFile(join(existingOutDir, "keep.txt"))).toEqual(
        keepBytes,
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});
