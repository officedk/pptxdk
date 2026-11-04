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
  "用法：react-pptx <文稿.tsx> [选项]",
  "选项：",
  "  -o, --out <路径>       PPTX 输出路径，默认写入当前目录下的同名 .pptx 文件",
  "  -e, --export <名称>    使用入口导出的具名组件，默认使用 default",
  "      --target <名称>    使用库公开的输出目标名称",
  "      --embed-fonts <模式>  嵌入完整字体 full 或已用字形 subset",
  "  -h, --help             显示帮助",
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
        throw new UsageError(`${argument} 后需要输出路径`);
      }
      output = value;
      index += 1;
      continue;
    }
    if (argument === "--export" || argument === "-e") {
      const value = args[index + 1];
      if (value === undefined || value === "" || value.startsWith("-")) {
        throw new UsageError(`${argument} 后需要导出名称`);
      }
      exportName = value;
      index += 1;
      continue;
    }
    if (argument === "--embed-fonts") {
      const value = args[index + 1];
      if (value !== "full" && value !== "subset")
        throw new UsageError("--embed-fonts 后需要 full 或 subset");
      embedFonts = value;
      index += 1;
      continue;
    }
    if (argument === "--target") {
      const value = args[index + 1];
      if (value === undefined || value === "" || value.startsWith("-")) {
        throw new UsageError("--target 后需要库公开的输出目标名称");
      }
      target = targetByName(value);
      if (target === undefined) {
        const names = Object.values(targets)
          .map(({ name }) => name)
          .join("、");
        throw new UsageError(
          `未知输出目标 ${JSON.stringify(value)}；可用值：${names}`,
        );
      }
      index += 1;
      continue;
    }
    if (argument.startsWith("-")) {
      throw new UsageError(`未知参数 ${JSON.stringify(argument)}`);
    }
    if (entry !== undefined) {
      throw new UsageError("单次导出只接受一个文稿入口");
    }
    entry = argument;
  }

  if (entry === undefined) throw new UsageError("缺少文稿源码入口");
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

/** 运行一次源码导出；通过 write 接缝返回终端内容和进程退出码。 */
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
    write(runtime, "stderr", "文稿导出已取消\n");
    return 130;
  }

  const entry = resolve(parsed.entry);
  const output = resolve(
    parsed.output ?? `${basename(entry, extname(entry))}.pptx`,
  );
  if (entry === output) {
    write(runtime, "stderr", "输出路径不能覆盖文稿源码\n");
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
          `警告 ${diagnostic.path} [${diagnostic.code}]：${diagnostic.message}\n`,
        );
      }
    }
    write(
      runtime,
      "stdout",
      `已导出文稿：${output}\n页面数：${result.report.slides.length}\n`,
    );
    return 0;
  } catch (error) {
    if (runtime.signal?.aborted && error === runtime.signal.reason) {
      write(runtime, "stderr", "文稿导出已取消\n");
      return 130;
    }
    if (error instanceof BuildError) {
      write(runtime, "stderr", "构建失败：\n");
      for (const diagnostic of error.diagnostics) {
        write(
          runtime,
          "stderr",
          `  ${diagnostic.path} [${diagnostic.code}]：${diagnostic.message}\n`,
        );
      }
      return 1;
    }
    if (error instanceof SourceLoadError) {
      write(runtime, "stderr", `文稿源码接入失败：${error.message}\n`);
      return 1;
    }
    write(runtime, "stderr", `文稿导出失败：${errorMessage(error)}\n`);
    return 1;
  }
}
