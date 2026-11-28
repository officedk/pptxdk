import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import {
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "@officedk/pptxdk";
import { deck, shape, slide, text } from "@officedk/pptxdk/vanilla";
import { describe, it } from "vitest";

import { runExternalCorpusCheck } from "../scripts/check-external-corpus.ts";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const jostAssetUrl = new URL(
  "./assets/fonts/Jost-400-Book.ttf",
  import.meta.url,
);

describe("L4 decoupled conditional external corpus runner", () => {
  it("verifies a single .pptx file and a flat top-layer directory with substitute fonts while ignoring subdirectories and non-pptx files", async () => {
    const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-l4-runner-test-"));
    try {
      const validDeck = deck({ width: 640, height: 360 }).append(
        slide({ id: "s1", background: "#FFFFFF" }).append(
          shape({
            id: "card",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 160,
              height: 60,
              background: "#2563EB",
            },
          }),
          text("Conditional check", {
            id: "title",
            style: {
              position: "absolute",
              left: 24,
              top: 100,
              width: 300,
              height: 48,
              fontFamily: "Missing System Sans",
              fontSize: 18,
            },
          }),
        ),
      );
      const validBytes = (await build(validDeck)).pptx;

      // 1. Single .pptx file input.
      const singlePptx = join(tempRoot, "single.pptx");
      await writeFile(singlePptx, validBytes);
      const singleReport = await runExternalCorpusCheck({ input: singlePptx });
      assert.equal(singleReport.passed, true);
      assert.equal(singleReport.totalFiles, 1);
      assert.equal(singleReport.violations, 0);
      assert.equal(singleReport.summary.importedCount, 1);
      assert.equal(singleReport.summary.codegenPassedCount, 1);
      assert.equal(singleReport.summary.batchTscPassed, true);
      assert.equal(singleReport.summary.emitPassedCount, 1);

      // 2. Flat directory with valid + typed-reject corrupt pptx + non-pptx + nested subdirectory pptx + flat fonts dir.
      const flatDir = join(tempRoot, "flat-corpus");
      await mkdir(join(flatDir, "nested-subdir"), { recursive: true });
      await writeFile(join(flatDir, "01-valid.pptx"), validBytes);
      await writeFile(
        join(flatDir, "02-corrupt.pptx"),
        new Uint8Array([1, 2, 3, 4]),
      );
      await writeFile(join(flatDir, "ignored.txt"), "not a pptx");
      await writeFile(join(flatDir, "ignored.zip"), validBytes);
      // Nested .pptx inside subdirectory must not be discovered.
      await writeFile(
        join(flatDir, "nested-subdir", "nested.pptx"),
        validBytes,
      );

      const fontsDir = join(tempRoot, "flat-fonts");
      await mkdir(join(fontsDir, "nested-fonts"), { recursive: true });
      await copyFile(jostAssetUrl, join(fontsDir, "Jost-400-Book.ttf"));
      await copyFile(
        jostAssetUrl,
        join(fontsDir, "nested-fonts", "Nested.ttf"),
      );

      const reportPath = join(tempRoot, "out", "report.json");
      const dirReport = await runExternalCorpusCheck({
        input: flatDir,
        fontsDir,
        reportPath,
      });

      assert.equal(dirReport.passed, true);
      assert.equal(dirReport.totalFiles, 2);
      assert.equal(dirReport.violations, 0);
      assert.equal(dirReport.summary.importedCount, 1);
      assert.equal(dirReport.summary.typedRejectedCount, 1);
      assert.equal(dirReport.summary.codegenPassedCount, 1);
      assert.equal(dirReport.summary.batchTscPassed, true);
      assert.equal(dirReport.summary.emitPassedCount, 1);
      assert.equal(dirReport.summary.paintCompletedCount, 1);
      assert.equal(dirReport.results[0]?.status, "passed");
      assert.equal(dirReport.results[1]?.status, "rejected-typed");
      assert.equal(dirReport.results[1]?.importErrorCode, "corrupt-archive");

      const persisted = JSON.parse(await readFile(reportPath, "utf8")) as {
        readonly totalFiles: number;
        readonly passed: boolean;
      };
      assert.equal(persisted.totalFiles, 2);
      assert.equal(persisted.passed, true);
    } finally {
      await rm(tempRoot, { recursive: true, force: true });
    }
  }, 60_000);

  it("rejects non-pptx inputs, empty directories, and fails with exit code 2 when invoked without --input or exit code 1 on timeout", async () => {
    const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-l4-negative-"));
    try {
      const nonPptx = join(tempRoot, "notes.txt");
      await writeFile(nonPptx, "hello");
      await assert.rejects(
        () => runExternalCorpusCheck({ input: nonPptx }),
        /\.pptx extension/u,
      );

      const onlySubdirDir = join(tempRoot, "only-subdir");
      await mkdir(join(onlySubdirDir, "child"), { recursive: true });
      const validBytes = (
        await build(
          deck({ width: 320, height: 180 }).append(slide({ id: "s1" })),
        )
      ).pptx;
      await writeFile(join(onlySubdirDir, "child", "deep.pptx"), validBytes);
      await assert.rejects(
        () => runExternalCorpusCheck({ input: onlySubdirDir }),
        /No top-level \.pptx files found/u,
      );

      // Timeout isolation marks sample violated without crashing the orchestrator.
      const slowPptx = join(tempRoot, "slow.pptx");
      await writeFile(slowPptx, validBytes);
      const timeoutReport = await runExternalCorpusCheck({
        input: slowPptx,
        timeoutMs: 1,
      });
      assert.equal(timeoutReport.passed, false);
      assert.equal(timeoutReport.summary.timedOutCount, 1);
      assert.equal(timeoutReport.violations, 1);

      const missingInputCli = spawnSync(
        process.execPath,
        ["--conditions=source", "scripts/check-external-corpus.ts"],
        { cwd: packageRoot, encoding: "utf8" },
      );
      assert.equal(missingInputCli.status, 2);
    } finally {
      await rm(tempRoot, { recursive: true, force: true });
    }
  }, 30_000);
});
