import { existsSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { build } from "@officedk/pptxdk";
import { deck, slide, text } from "@officedk/pptxdk/vanilla";
import { strFromU8, unzipSync, zipSync } from "fflate";
import { afterEach, describe, expect, it, vi } from "vitest";

describe("pptxdk CLI codegen CodegenError handling", () => {
  afterEach(() => {
    vi.doUnmock("@officedk/pptxdk");
    vi.resetModules();
  });

  it("maps CodegenError to codegenFailed with declaration location and prior import diagnostics without disk writes", async () => {
    const directory = await mkdtemp(join(tmpdir(), "pptxdk-cli-codegen-err-"));
    try {
      const { pptx } = await build(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "s1" }).append(
            text("Hello", { id: "t1", x: 10, y: 10, width: 120, height: 30 }),
          ),
        ),
      );

      // Inject an unconsumed element so import succeeds partially with a diagnostic before codegen fails.
      const parts = unzipSync(pptx);
      const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!).replace(
        "</p:spTree>",
        '<unknown:widget xmlns:unknown="https://pptxdk.example.com/unknown"/></p:spTree>',
      );
      const partialPath = join(directory, "partial-codegen-fail.pptx");
      await writeFile(
        partialPath,
        zipSync({
          ...parts,
          "ppt/slides/slide1.xml": Buffer.from(slideXml, "utf8"),
        }),
      );

      vi.resetModules();
      vi.doMock("@officedk/pptxdk", async (importOriginal) => {
        const actual =
          await importOriginal<typeof import("@officedk/pptxdk")>();
        return {
          ...actual,
          codegen: () => {
            throw new actual.CodegenError([
              {
                path: "deck/slide[s1]/shape[unsupported]",
                code: "unsupported-declaration",
                declaration: "shape",
                message: "Cannot serialize custom declaration node",
              },
            ]);
          },
        };
      });

      const { runCli } = await import("../../src/cli-commands.ts");

      const outDir = join(directory, "should-not-exist");
      let textStdout = "";
      let textStderr = "";
      const textCode = await runCli({
        argv: [
          "node",
          "pptxdk",
          "codegen",
          partialPath,
          "-t",
          "vanilla",
          "--out-dir",
          outDir,
        ],
        cwd: directory,
        write: ({ destination, chunk }) => {
          if (destination === "stdout") textStdout += chunk;
          else textStderr += chunk;
        },
      });
      expect(textCode).toBe(1);
      expect(textStdout).toBe("");
      expect(textStderr).toContain("ppt/slides/slide1.xml");
      expect(textStderr).toContain("Source generation failed:");
      expect(textStderr).toContain(
        "deck/slide[s1]/shape[unsupported] [unsupported-declaration]: Cannot serialize custom declaration node",
      );
      expect(existsSync(outDir)).toBe(false);

      const outZip = join(directory, "should-not-exist.zip");
      let jsonStdout = "";
      let jsonStderr = "";
      const jsonCode = await runCli({
        argv: [
          "node",
          "pptxdk",
          "codegen",
          partialPath,
          "-t",
          "react",
          "--out-zip",
          outZip,
          "--output",
          "structured",
        ],
        cwd: directory,
        write: ({ destination, chunk }) => {
          if (destination === "stdout") jsonStdout += chunk;
          else jsonStderr += chunk;
        },
      });
      expect(jsonCode).toBe(1);
      expect(jsonStdout).toBe("");
      expect(JSON.parse(jsonStderr)).toMatchObject({
        schemaVersion: "3",
        command: "codegen",
        outputMode: "structured",
        kind: "failure",
        variant: "codegenFailed",
        data: {
          entry: partialPath,
          diagnostics: [
            {
              path: "deck/slide[s1]/shape[unsupported]",
              code: "unsupported-declaration",
              declaration: "shape",
              message: "Cannot serialize custom declaration node",
            },
          ],
          importDiagnostics: [
            {
              location: { part: "ppt/slides/slide1.xml" },
            },
          ],
        },
      });
      expect(existsSync(outZip)).toBe(false);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });
});
