import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { validatePptx } from "@react-pptx/validation";
import { describe, expect, it } from "vitest";

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

describe("react-pptx 单次导出 CLI", () => {
  it("从带空格的 TSX 路径解析作者依赖并写出公开构建结果", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx cli "));
    try {
      const source = fixture("source deck.tsx");
      const output = join(directory, "programmatic.pptx");
      const baselineOutput = join(directory, "programmatic-baseline.pptx");
      const cliOutput = join(directory, "nested output", "source deck.pptx");
      const cliBaselineOutput = join(
        directory,
        "nested output",
        "baseline.pptx",
      );
      const programmaticScript = [
        'import { pathToFileURL } from "node:url";',
        'import { writeFile } from "node:fs/promises";',
        'import { targets } from "@react-pptx/react-pptx";',
        'import { buildFromSource } from "@react-pptx/react-pptx/node";',
        `const result = await buildFromSource(${JSON.stringify(source)});`,
        `await writeFile(${JSON.stringify(output)}, result.pptx);`,
        `const baselineResult = await buildFromSource(${JSON.stringify(source)}, { target: targets.baselineOnly });`,
        `await writeFile(${JSON.stringify(baselineOutput)}, baselineResult.pptx);`,
        'async function expectHookReleased(path) { const error = await import(pathToFileURL(path).href).then(() => undefined, failure => failure); if (error?.code !== "ERR_UNKNOWN_FILE_EXTENSION") throw new Error(JSON.stringify({ path, error })); }',
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after success.tsx"))});`,
        'process.env.REACT_PPTX_TEST_LOAD_FAILURE = "1";',
        `const loadFailure = await buildFromSource(${JSON.stringify(fixture("load failure.tsx"))}).then(() => undefined, error => error);`,
        'if (loadFailure?.name !== "SourceLoadError") throw new Error("源码解析错误没有被保留");',
        "delete process.env.REACT_PPTX_TEST_LOAD_FAILURE;",
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after load failure.tsx"))});`,
        `const buildFailure = await buildFromSource(${JSON.stringify(fixture("broken deck.tsx"))}).then(() => undefined, error => error);`,
        'if (buildFailure?.name !== "BuildError") throw new Error("构建错误没有被保留");',
        `await expectHookReleased(${JSON.stringify(fixture("unloaded after build failure.tsx"))});`,
        "const controller = new AbortController();",
        'const reason = new Error("程序化源码构建已取消");',
        "globalThis.__reactPptxCancellation = { controller, reason };",
        `const cancellation = await buildFromSource(${JSON.stringify(fixture("cancelled deck.tsx"))}, { signal: controller.signal }).then(() => undefined, error => error);`,
        'if (cancellation !== reason) throw new Error("取消原因未从公开构建传播");',
        "delete globalThis.__reactPptxCancellation;",
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

      const result = runNode(
        [cli, source, "--export", "SourceDeck", "--out", cliOutput],
        directory,
      );

      expect(result.code).toBe(0);
      expect(result.stdout).toContain(cliOutput);
      const bytes = await readFile(cliOutput);
      expect(bytes).toEqual(await readFile(output));

      const file = await validatePptx(new Uint8Array(bytes));
      expect(file.errors).toEqual([]);
      expect(file.slides[0]?.objects[0]).toMatchObject({
        name: "来自源码依赖",
        text: { text: "来自源码依赖" },
      });

      const baseline = runNode(
        [cli, source, "--target", "baseline-only", "--out", cliBaselineOutput],
        directory,
      );
      expect(baseline.code).toBe(0);
      const baselineBytes = await readFile(cliBaselineOutput);
      expect(baselineBytes).toEqual(await readFile(baselineOutput));
      expect(
        (await validatePptx(new Uint8Array(baselineBytes))).extensions,
      ).toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("构建诊断导致失败退出且不写出伪成功制品", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx cli "));
    try {
      const output = join(directory, "never.pptx");
      const result = runNode(
        [cli, fixture("broken deck.tsx"), "--out", output],
        directory,
      );

      expect(result.code).toBe(1);
      expect(result.stderr).toContain("构建失败");
      expect(result.stderr).toContain("invalid-coordinate");
      expect(existsSync(output)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("把处理中断接入公开构建取消并以中断状态退出", async () => {
    const directory = await mkdtemp(join(tmpdir(), "react-pptx cli "));
    try {
      const output = join(directory, "cancelled.pptx");
      const child = spawn(
        process.execPath,
        [cli, fixture("cancelled deck.tsx"), "--out", output],
        {
          cwd: directory,
          stdio: ["ignore", "pipe", "pipe"],
        },
      );
      let stdout = "";
      let stderr = "";
      let interrupted = false;
      child.stdout.setEncoding("utf8");
      child.stderr.setEncoding("utf8");
      child.stdout.on("data", (chunk: string) => {
        stdout += chunk;
        if (!interrupted && stdout.includes("__react_pptx_cancel_ready__")) {
          interrupted = true;
          child.kill("SIGINT");
        }
      });
      child.stderr.on("data", (chunk: string) => {
        stderr += chunk;
      });
      const code = await waitForClose(child);

      expect(code).toBe(130);
      expect(stderr).toContain("已取消");
      expect(existsSync(output)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it("对无效参数返回用法状态", async () => {
    const result = runNode([cli, "--unknown"], import.meta.dirname);

    expect(result.code).toBe(2);
    expect(result.stderr).toContain("用法");
  });
});
