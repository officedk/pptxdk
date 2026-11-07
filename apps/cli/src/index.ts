import { randomUUID } from "node:crypto";
import { mkdir, rename, rm, writeFile } from "node:fs/promises";
import { basename, dirname, extname, resolve } from "node:path";

import { BuildError, targets, type Target } from "@react-pptx/react-pptx";
import { buildFromSource, SourceLoadError } from "@react-pptx/react-pptx/node";

type CliDestination = "stdout" | "stderr";

export type CliRuntime = {
  readonly argv: readonly string[];
  readonly signal?: AbortSignal;
  readonly write: (output: {
    readonly destination: CliDestination;
    readonly chunk: string;
  }) => void;
};

type ParsedArguments = {
  readonly entry: string;
  readonly output?: string;
  readonly exportName?: string;
  readonly target?: Target;
  readonly embedFonts?: "full" | "subset";
  readonly help: boolean;
};

class UsageError extends Error {}

const usage = [
  "Usage: react-pptx <deck.tsx> [options]",
  "Options:",
  "  -o, --out <path>         PPTX output path (defaults to <deck>.pptx in current directory)",
  "  -e, --export <name>      Named component exported from entry (defaults to default)",
  "      --target <name>      Target name exposed by library",
  "      --embed-fonts <mode> Embed full fonts or subset of used glyphs (full or subset)",
  "  -h, --help               Show help",
].join("\n");

function targetByName(name: string): Target | undefined {
  return Object.values(targets).find((target) => target.name === name);
}

function parseArguments(args: readonly string[]): ParsedArguments {
  if (args.includes("--help") || args.includes("-h")) {
    return { entry: "", help: true };
  }

  let entry: string | undefined;
  let output: string | undefined;
  let exportName: string | undefined;
  let target: Target | undefined;
  let embedFonts: "full" | "subset" | undefined;

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index]!;
    if (argument === "--out" || argument === "-o") {
      const value = args[index + 1];
      if (value === undefined || value === "" || value.startsWith("-")) {
        throw new UsageError(`${argument} requires an output path`);
      }
      output = value;
      index += 1;
      continue;
    }
    if (argument === "--export" || argument === "-e") {
      const value = args[index + 1];
      if (value === undefined || value === "" || value.startsWith("-")) {
        throw new UsageError(`${argument} requires an export name`);
      }
      exportName = value;
      index += 1;
      continue;
    }
    if (argument === "--embed-fonts") {
      const value = args[index + 1];
      if (value !== "full" && value !== "subset")
        throw new UsageError("--embed-fonts requires full or subset");
      embedFonts = value;
      index += 1;
      continue;
    }
    if (argument === "--target") {
      const value = args[index + 1];
      if (value === undefined || value === "" || value.startsWith("-")) {
        throw new UsageError(
          "--target requires a target name exposed by library",
        );
      }
      target = targetByName(value);
      if (target === undefined) {
        const names = Object.values(targets)
          .map(({ name }) => name)
          .join(", ");
        throw new UsageError(
          `Unknown target ${JSON.stringify(value)}; available values: ${names}`,
        );
      }
      index += 1;
      continue;
    }
    if (argument.startsWith("-")) {
      throw new UsageError(`Unknown option ${JSON.stringify(argument)}`);
    }
    if (entry !== undefined) {
      throw new UsageError("Single export accepts only one deck entrypoint");
    }
    entry = argument;
  }

  if (entry === undefined)
    throw new UsageError("Missing deck source entrypoint");
  return {
    entry,
    ...(output !== undefined && { output }),
    ...(exportName !== undefined && { exportName }),
    ...(target !== undefined && { target }),
    ...(embedFonts !== undefined && { embedFonts }),
    help: false,
  };
}

function write(
  runtime: CliRuntime,
  destination: CliDestination,
  chunk: string,
): void {
  runtime.write({ destination, chunk });
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function writeArtifact(path: string, bytes: Uint8Array): Promise<void> {
  const directory = dirname(path);
  const temporary = resolve(
    directory,
    `.${basename(path)}.${randomUUID()}.tmp`,
  );
  try {
    await mkdir(directory, { recursive: true });
    await writeFile(temporary, bytes, { flag: "wx" });
    await rename(temporary, path);
  } finally {
    await rm(temporary, { force: true });
  }
}

/** Run a single source export; return terminal content and process exit code via the write seam. */
export async function runCli(runtime: CliRuntime): Promise<number> {
  let parsed: ParsedArguments;
  try {
    parsed = parseArguments(runtime.argv.slice(2));
  } catch (error) {
    write(runtime, "stderr", `${errorMessage(error)}\n${usage}\n`);
    return 2;
  }

  if (parsed.help) {
    write(runtime, "stdout", `${usage}\n`);
    return 0;
  }
  if (runtime.signal?.aborted) {
    write(runtime, "stderr", "Deck export cancelled\n");
    return 130;
  }

  const entry = resolve(parsed.entry);
  const output = resolve(
    parsed.output ?? `${basename(entry, extname(entry))}.pptx`,
  );
  if (entry === output) {
    write(runtime, "stderr", "Output path cannot overwrite deck source\n");
    return 2;
  }

  try {
    const result = await buildFromSource(entry, {
      ...(parsed.exportName !== undefined && {
        exportName: parsed.exportName,
      }),
      ...(parsed.target !== undefined && { target: parsed.target }),
      ...(runtime.signal !== undefined && { signal: runtime.signal }),
      ...(parsed.embedFonts !== undefined && { embedFonts: parsed.embedFonts }),
    });
    await writeArtifact(output, result.pptx);

    for (const diagnostic of result.report.diagnostics) {
      if (diagnostic.severity === "warning") {
        write(
          runtime,
          "stderr",
          `Warning ${diagnostic.path} [${diagnostic.code}]: ${diagnostic.message}\n`,
        );
      }
    }
    write(
      runtime,
      "stdout",
      `Exported deck: ${output}\nSlides: ${result.report.slides.length}\n`,
    );
    return 0;
  } catch (error) {
    if (runtime.signal?.aborted && error === runtime.signal.reason) {
      write(runtime, "stderr", "Deck export cancelled\n");
      return 130;
    }
    if (error instanceof BuildError) {
      write(runtime, "stderr", "Build failed:\n");
      for (const diagnostic of error.diagnostics) {
        write(
          runtime,
          "stderr",
          `  ${diagnostic.path} [${diagnostic.code}]: ${diagnostic.message}\n`,
        );
      }
      return 1;
    }
    if (error instanceof SourceLoadError) {
      write(
        runtime,
        "stderr",
        `Failed to load deck source: ${error.message}\n`,
      );
      return 1;
    }
    write(runtime, "stderr", `Deck export failed: ${errorMessage(error)}\n`);
    return 1;
  }
}
