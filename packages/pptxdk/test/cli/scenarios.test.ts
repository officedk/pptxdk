import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  defineCommandScenarios,
  defineFailureScenarios,
  runCliScenario,
} from "@boringparts/cli-contract/testing";
import { CodegenError, build } from "@officedk/pptxdk";
import { deck, shape, slide } from "@officedk/pptxdk/vanilla";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { cliContract } from "../../src/cli-commands.ts";

const fixture = (name: string) => join(import.meta.dirname, "fixtures", name);

describe("pptxdk CLI contract scenarios", () => {
  let stageDir = "";
  let validPptxPath = "";
  let corruptPptxPath = "";
  let occupiedFilePath = "";
  let buildWriteFailDir = "";
  let previewWriteFailDir = "";
  let codegenWriteFailZipDir = "";

  beforeAll(async () => {
    stageDir = await mkdtemp(join(tmpdir(), "pptxdk-cli-scenarios-"));
    validPptxPath = join(stageDir, "valid.pptx");
    corruptPptxPath = join(stageDir, "corrupt.pptx");
    occupiedFilePath = join(stageDir, "occupied-file");
    buildWriteFailDir = join(stageDir, "occupied-out.pptx");
    previewWriteFailDir = join(stageDir, "preview-write-fail");
    codegenWriteFailZipDir = join(stageDir, "occupied-codegen.zip");

    const { pptx } = await build(
      deck({ width: 240, height: 140 }).append(
        slide({ id: "s1" }).append(
          shape({
            id: "box",
            style: {
              position: "absolute",
              left: 16,
              top: 16,
              width: 80,
              height: 40,
              background: "#2563EB",
            },
          }),
        ),
      ),
    );
    await writeFile(validPptxPath, pptx);
    await writeFile(corruptPptxPath, "not-a-pptx-zip");
    await writeFile(occupiedFilePath, "occupied");
    await mkdir(buildWriteFailDir, { recursive: true });
    await mkdir(join(previewWriteFailDir, "page-01.png"), { recursive: true });
    await mkdir(codegenWriteFailZipDir, { recursive: true });
  });

  afterAll(async () => {
    if (stageDir.length > 0) {
      await rm(stageDir, { recursive: true, force: true });
    }
  });

  it("covers all executable commands with static exhaustiveness", async () => {
    const commandScenarios = defineCommandScenarios(cliContract, {
      build: {
        argv: [
          "build",
          fixture("source deck.tsx"),
          "--out",
          join(stageDir, "built.pptx"),
        ],
        dependencies: { cwd: stageDir },
      },
      preview: {
        argv: [
          "preview",
          fixture("preview deck.tsx"),
          "--out-dir",
          join(stageDir, "preview-out"),
        ],
        dependencies: { cwd: stageDir },
      },
      codegen: {
        argv: [
          "codegen",
          validPptxPath,
          "-t",
          "vanilla",
          "--out-dir",
          join(stageDir, "codegen-out"),
        ],
        dependencies: { cwd: stageDir },
      },
    });

    const expectedVariants = {
      build: "built",
      preview: "previewed",
      codegen: "generated",
    } as const;

    for (const [command, scenario] of Object.entries(commandScenarios) as [
      keyof typeof commandScenarios,
      (typeof commandScenarios)[keyof typeof commandScenarios],
    ][]) {
      const observation = await runCliScenario(cliContract, scenario);
      expect(observation.stderr).toBe("");
      expect(observation.stdout.length).toBeGreaterThan(0);
      expect(observation.termination.exitCode).toBe(0);
      expect(observation.termination.kind).toBe("applicationResult");
      if (observation.termination.kind === "applicationResult") {
        expect(observation.termination.result.kind).toBe("data");
        expect(observation.termination.result.command).toBe(command);
        expect(observation.termination.result.variant).toBe(
          expectedVariants[command],
        );
      }
    }
  });

  it("covers all declared failure variants across all commands with static exhaustiveness", async () => {
    const preAborted = new AbortController();
    preAborted.abort(new Error("cancelled by scenario"));

    const sourceDeck = fixture("source deck.tsx");
    const previewDeck = fixture("preview deck.tsx");
    const brokenDeck = fixture("broken deck.tsx");

    const failureScenarios = defineFailureScenarios(cliContract, {
      build: {
        invalidOutput: {
          argv: ["build", sourceDeck, "--out", sourceDeck],
          dependencies: { cwd: stageDir },
        },
        sourceLoadFailed: {
          argv: [
            "build",
            sourceDeck,
            "--export",
            "MissingExport",
            "--out",
            join(stageDir, "never-build-load.pptx"),
          ],
          dependencies: { cwd: stageDir },
        },
        buildFailed: {
          argv: [
            "build",
            brokenDeck,
            "--out",
            join(stageDir, "never-build-fail.pptx"),
          ],
          dependencies: { cwd: stageDir },
        },
        writeFailed: {
          argv: ["build", sourceDeck, "--out", buildWriteFailDir],
          dependencies: { cwd: stageDir },
        },
        cancelled: {
          argv: [
            "build",
            sourceDeck,
            "--out",
            join(stageDir, "never-build-cancel.pptx"),
          ],
          dependencies: { cwd: stageDir, signal: preAborted.signal },
        },
      },
      preview: {
        invalidOutput: {
          argv: ["preview", previewDeck, "--out-dir", occupiedFilePath],
          dependencies: { cwd: stageDir },
        },
        pageOutOfRange: {
          argv: [
            "preview",
            previewDeck,
            "--pages",
            "99",
            "--out-dir",
            join(stageDir, "never-preview-range"),
          ],
          dependencies: { cwd: stageDir },
        },
        sourceLoadFailed: {
          argv: [
            "preview",
            previewDeck,
            "--export",
            "MissingExport",
            "--out-dir",
            join(stageDir, "never-preview-load"),
          ],
          dependencies: { cwd: stageDir },
        },
        buildFailed: {
          argv: [
            "preview",
            brokenDeck,
            "--out-dir",
            join(stageDir, "never-preview-build"),
          ],
          dependencies: { cwd: stageDir },
        },
        previewFailed: {
          argv: [
            "preview",
            fixture("partial-failure-preview.tsx"),
            "--out-dir",
            join(stageDir, "partial-preview-fail"),
          ],
          dependencies: { cwd: stageDir },
        },
        writeFailed: {
          argv: ["preview", previewDeck, "--out-dir", previewWriteFailDir],
          dependencies: { cwd: stageDir },
        },
        cancelled: {
          argv: [
            "preview",
            previewDeck,
            "--out-dir",
            join(stageDir, "never-preview-cancel"),
          ],
          dependencies: { cwd: stageDir, signal: preAborted.signal },
        },
      },
      codegen: {
        invalidOutput: {
          argv: [
            "codegen",
            validPptxPath,
            "-t",
            "vanilla",
            "--out-zip",
            validPptxPath,
          ],
          dependencies: { cwd: stageDir },
        },
        inputReadFailed: {
          argv: [
            "codegen",
            join(stageDir, "missing.pptx"),
            "-t",
            "vanilla",
            "--out-dir",
            join(stageDir, "never-codegen-read"),
          ],
          dependencies: { cwd: stageDir },
        },
        importFailed: {
          argv: [
            "codegen",
            corruptPptxPath,
            "-t",
            "vanilla",
            "--out-dir",
            join(stageDir, "never-codegen-import"),
          ],
          dependencies: { cwd: stageDir },
        },
        codegenFailed: {
          argv: [
            "codegen",
            validPptxPath,
            "-t",
            "vanilla",
            "--out-dir",
            join(stageDir, "never-codegen-fail"),
          ],
          dependencies: {
            cwd: stageDir,
            codegen: () => {
              throw new CodegenError([
                {
                  path: "deck/slide[s1]/shape[unsupported]",
                  code: "unsupported-declaration",
                  declaration: "shape",
                  message: "Cannot serialize custom declaration node",
                },
              ]);
            },
          },
        },
        writeFailed: {
          argv: [
            "codegen",
            validPptxPath,
            "-t",
            "vanilla",
            "--out-zip",
            codegenWriteFailZipDir,
          ],
          dependencies: { cwd: stageDir },
        },
        cancelled: {
          argv: [
            "codegen",
            validPptxPath,
            "-t",
            "vanilla",
            "--out-dir",
            join(stageDir, "never-codegen-cancel"),
          ],
          dependencies: { cwd: stageDir, signal: preAborted.signal },
        },
      },
    });

    const expectedExitCode = (variant: string): number => {
      if (variant === "invalidOutput" || variant === "pageOutOfRange") return 2;
      if (variant === "cancelled") return 130;
      return 1;
    };

    for (const [command, variants] of Object.entries(failureScenarios)) {
      for (const [variant, scenario] of Object.entries(variants)) {
        const observation = await runCliScenario(cliContract, scenario);
        expect(observation.stdout).toBe("");
        expect(observation.stderr.length).toBeGreaterThan(0);
        expect(observation.termination.exitCode).toBe(
          expectedExitCode(variant),
        );
        expect(observation.termination.kind).toBe("applicationResult");
        if (observation.termination.kind === "applicationResult") {
          expect(observation.termination.result.kind).toBe("failure");
          expect(observation.termination.result.command).toBe(command);
          expect(observation.termination.result.variant).toBe(variant);
        }
      }
    }
  });
});
