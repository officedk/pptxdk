import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import {
  mkdir,
  mkdtemp,
  readFile,
  rename,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { build, codegen, defineTheme, importPptx, lighter } from "pptxdk";
import { buildFromSource } from "pptxdk/node";
import {
  cell,
  chart,
  createMotionScopeHandle,
  createMorphScope,
  deck,
  defineSlideLayout,
  font,
  group,
  image,
  motionScope,
  notes,
  paragraph,
  path,
  placeholder,
  row,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  step,
  table,
  text,
  after,
  fade,
  transitions,
  wipe,
} from "pptxdk/vanilla";
import { describe, expect, it } from "vitest";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const cliBin = join(packageRoot, "dist/cli.js");

function createSamplePng(): Uint8Array {
  return Uint8Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      "base64",
    ),
  );
}

async function writeUnzippedEntries(
  targetDir: string,
  entries: Readonly<Record<string, Uint8Array>>,
): Promise<void> {
  for (const [relPath, bytes] of Object.entries(entries)) {
    const dest = join(targetDir, ...relPath.split("/"));
    await mkdir(dirname(dest), { recursive: true });
    await writeFile(dest, bytes);
  }
}

describe("pptxdk CLI combined release directory and ZIP relocated rebuild", () => {
  it("exports directory and ZIP for a full combined deck, deletes input PPTX, and rebuilds cleanly from relocated paths", async () => {
    const stageRoot = await mkdtemp(join(tmpdir(), "pptxdk-cli-combined-"));
    const consumerRoot = await mkdtemp(
      join(tmpdir(), "pptxdk-cli-combined-consumer-"),
    );

    try {
      const inputDir = join(stageRoot, "input");
      await mkdir(inputDir, { recursive: true });
      const inputPptx = join(inputDir, "combined.pptx");

      const jostBytes = new Uint8Array(
        await readFile(
          new URL("./fixtures/Jost-400-Book.ttf", import.meta.url),
        ),
      );
      const jostFont = font("Jost*", [{ weight: 400, data: jostBytes }]);
      const sharedPng = createSamplePng();
      const brandTheme = defineTheme({
        name: "CLI Combined Brand",
        colors: { accent1: "#1D4ED8", brandTint: lighter("accent1", 0.3) },
        fonts: { major: jostFont, minor: jostFont },
      });

      const primaryLayout = defineSlideLayout({
        id: "cli-content",
        name: "Duplicate Name",
        type: "obj",
        preserve: true,
        placeholders: {
          title: { type: "title", x: 40, y: 48, width: 560, height: 48 },
          picture: {
            type: "pic",
            explicitType: true,
            idx: 1,
            x: 40,
            y: 112,
            width: 140,
            height: 100,
          },
          table: {
            type: "tbl",
            explicitType: true,
            idx: 2,
            x: 200,
            y: 112,
            width: 300,
            height: 140,
          },
          chart: {
            type: "chart",
            explicitType: true,
            idx: 3,
            x: 520,
            y: 112,
            width: 380,
            height: 220,
          },
          emptySlot: {
            type: "obj",
            explicitType: false,
            idx: 4,
            x: 40,
            y: 230,
            width: 140,
            height: 80,
          },
        },
      });
      const idleLayout = defineSlideLayout({
        id: "cli-idle",
        name: "Duplicate Name",
        type: "obj",
        preserve: true,
      });

      const morphScope = createMorphScope();
      const heroKey = morphScope.handle("hero");
      const leadScope = createMotionScopeHandle();
      const followScope = createMotionScopeHandle();
      const readyCue = leadScope.cue("ready");

      const combinedDeck = deck({
        width: 960,
        height: 540,
        theme: brandTheme,
      }).append(
        slideMaster({
          id: "cli-master-1",
          name: "Brand Master",
          theme: brandTheme,
          background: "#F8FAFC",
          headerFooter: { dateTime: false, footer: true, slideNumber: true },
          footerPlaceholders: { footer: { text: "CLI Combined Footer" } },
          placeholders: {
            ftr: { type: "ftr", idx: 5, x: 40, y: 500, width: 300, height: 24 },
          },
          motion: {
            steps: [
              {
                trigger: "click",
                effects: [
                  {
                    target: { placeholder: "ftr" },
                    effect: wipe({ from: "left", duration: 200 }),
                  },
                ],
              },
            ],
          },
        }).append(
          image({
            id: "master-logo",
            src: sharedPng,
            x: 24,
            y: 12,
            width: 36,
            height: 28,
          }),
          slideLayout(primaryLayout),
          slideLayout(idleLayout),
        ),
        section({ name: "Part One" }).append(
          slide({
            id: "s1",
            layout: primaryLayout,
            footer: "CLI Combined Footer",
            slideNumber: true,
            transition: transitions.fade({ duration: 400 }),
            advanceAfter: 5000,
          }).append(
            step({ n: 1 }).append(
              text(
                [
                  paragraph().append(
                    span("Release ", {
                      style: { color: "#1D4ED8" },
                    }),
                    slideNumber(),
                  ),
                  paragraph("Overview"),
                ],
                {
                  id: "s1-title",
                  placeholder: "title",
                  link: { url: "https://example.com" },
                  enter: fade({ duration: 200, by: "paragraph" }),
                  style: { fontFamily: jostFont },
                },
              ),
              chart({
                id: "s1-chart",
                placeholder: "chart",
                x: 520,
                y: 112,
                width: 380,
                height: 220,
                enter: fade({ duration: 220, by: "series" }),
                style: { fontFamily: jostFont },
                axes: {
                  cat: { kind: "category" },
                  val1: { kind: "value", position: "left" },
                  val2: { kind: "value", position: "right" },
                },
                groups: [
                  {
                    kind: "column",
                    axes: { category: "cat", value: "val1" },
                    data: {
                      kind: "shared-categories",
                      categories: ["Q1", "Q2"],
                    },
                    series: [
                      {
                        name: "Actual",
                        values: [0, 42],
                        trendlines: [{ type: "linear" }],
                        errorBars: { type: "fixed", value: 3 },
                      },
                    ],
                  },
                  {
                    kind: "line",
                    axes: { category: "cat", value: "val2" },
                    data: {
                      kind: "shared-categories",
                      categories: ["Q1", "Q2"],
                    },
                    series: [{ name: "Target", values: [10, 30] }],
                  },
                ],
              }),
            ),
            image({
              id: "s1-pic",
              placeholder: "picture",
              src: sharedPng,
              crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 },
            }),
            table({
              id: "s1-tbl",
              placeholder: "table",
              columns: [100, 100, 100],
              style: { fontFamily: jostFont },
            }).append(
              row({ style: { height: 32 } }).append(
                cell("Plan", { colSpan: 3 }),
              ),
              row({ style: { height: 34 } }).append(
                cell("A", { rowSpan: 2 }),
                cell("10"),
                cell("20"),
              ),
              row({ style: { height: 34 } }).append(cell("30"), cell("40")),
            ),
            placeholder({ id: "s1-empty", placeholder: "emptySlot" }),
            shape({
              id: "s1-morph",
              morphKey: heroKey,
              x: 40,
              y: 340,
              width: 140,
              height: 60,
              fill: "#1D4ED8",
            }),
            notes("CLI slide 1 notes"),
          ),
        ),
        section({ name: "Part Two" }).append(
          slide({
            id: "s2",
            layout: primaryLayout,
            transition: transitions.morph({
              duration: 800,
              fallback: transitions.fade({ duration: 400 }),
            }),
            advanceAfter: 4000,
          }).append(
            shape({
              id: "s2-img-fill",
              x: 40,
              y: 120,
              width: 140,
              height: 90,
              geometry: path("M 0 0 L 140 0 L 140 90 L 0 90 Z", {
                viewBox: [140, 90],
              }),
              style: { backgroundImage: { image: sharedPng } },
            }),
            group({
              id: "s2-outer",
              name: "Outer Group",
            }).append(
              group({
                id: "s2-inner",
                name: "Inner Group",
              }).append(
                shape({
                  id: "s2-inner-shape",
                  x: 230,
                  y: 130,
                  width: 100,
                  height: 60,
                  fill: "#059669",
                }),
              ),
            ),
            shape({
              id: "s2-morph",
              morphKey: heroKey,
              x: 40,
              y: 340,
              width: 200,
              height: 70,
              fill: "#059669",
            }),
            motionScope(leadScope, {
              mode: "sequence",
              publishEnd: readyCue,
            }).append(
              shape({
                id: "s2-lead",
                x: 280,
                y: 340,
                width: 100,
                height: 50,
                fill: "#D97706",
                enter: fade({ duration: 200 }),
              }),
            ),
            motionScope(followScope, {
              startAt: after(readyCue, 100),
            }).append(
              shape({
                id: "s2-follow",
                x: 400,
                y: 340,
                width: 100,
                height: 50,
                fill: "#0F172A",
                enter: wipe({ from: "left", duration: 200 }),
              }),
            ),
            chart({
              id: "s2-ring",
              x: 540,
              y: 300,
              width: 240,
              height: 160,
              style: { fontFamily: jostFont },
              groups: [
                {
                  kind: "doughnut",
                  data: { kind: "independent-categories" },
                  series: [
                    { name: "Empty", categories: [], values: [] },
                    { name: "Filled", categories: ["A", "B"], values: [0, 20] },
                  ],
                },
              ],
            }),
            notes("CLI slide 2 notes"),
          ),
        ),
      );

      const { pptx } = await build(combinedDeck, { embedFonts: "full" });
      await writeFile(inputPptx, pptx);

      const importedCheck = importPptx(new Uint8Array(pptx));
      expect(importedCheck.report.diagnostics).toEqual([]);

      const deliveries: {
        readonly target: "vanilla" | "react";
        readonly dirPath: string;
        readonly zipPath: string;
        readonly entry: string;
      }[] = [];

      for (const target of ["vanilla", "react"] as const) {
        const expected = codegen(importedCheck.document, { target });
        const outDir = join(stageRoot, `${target}-dir`);
        const outZip = join(stageRoot, `${target}.zip`);

        execFileSync(
          process.execPath,
          [cliBin, "codegen", inputPptx, "-t", target, "--out-dir", outDir],
          { cwd: stageRoot, encoding: "utf8" },
        );
        execFileSync(
          process.execPath,
          [cliBin, "codegen", inputPptx, "-t", target, "--out-zip", outZip],
          { cwd: stageRoot, encoding: "utf8" },
        );
        deliveries.push({
          target,
          dirPath: outDir,
          zipPath: outZip,
          entry: expected.entry,
        });
      }

      // Delete original input PPTX before relocating directory and ZIP outputs
      await rm(inputDir, { recursive: true, force: true });
      expect(existsSync(inputPptx)).toBe(false);

      const cliNodeModules = join(packageRoot, "node_modules");
      await mkdir(join(consumerRoot, "node_modules"), { recursive: true });
      await symlink(
        join(cliNodeModules, "pptxdk"),
        join(consumerRoot, "node_modules/pptxdk"),
        "dir",
      );
      await symlink(
        join(cliNodeModules, "react"),
        join(consumerRoot, "node_modules/react"),
        "dir",
      );

      for (const item of deliveries) {
        const movedDir = join(consumerRoot, `moved-${item.target}`);
        const unzippedDir = join(consumerRoot, `unzipped-${item.target}`);
        await rename(item.dirPath, movedDir);
        await writeUnzippedEntries(
          unzippedDir,
          unzipSync(new Uint8Array(await readFile(item.zipPath))),
        );

        for (const relocatedDir of [movedDir, unzippedDir]) {
          const rebuilt = await buildFromSource(
            join(relocatedDir, item.entry),
            { embedFonts: "full" },
          );
          const validated = await validatePptx(rebuilt.pptx);
          expect(validated.errors).toEqual([]);
          expect(
            validated.presentation.sections?.map((sec) => sec.name),
          ).toEqual(["Part One", "Part Two"]);
          expect(validated.masters?.length).toBe(1);
          expect(validated.masters?.[0]?.layouts.length).toBe(2);
          expect(validated.slides.length).toBe(2);
          expect(
            validated.slides[1]?.objects.some(
              (o) =>
                o.kind === "chart" &&
                o.chart?.series.some(
                  (s) => s.name === "Empty" && s.values?.length === 0,
                ),
            ),
          ).toBe(true);
          expect(
            validated.slides[0]?.objects.some(
              (o) => o.kind === "table" && (o.table?.merges?.length ?? 0) >= 2,
            ),
          ).toBe(true);
          expect(validated.slides[1]?.transition?.kind).toBe("morph");
          const parts = unzipSync(rebuilt.pptx);
          expect(strFromU8(parts["ppt/slides/slide2.xml"]!)).toContain(
            "<a:blipFill",
          );
        }
      }
    } finally {
      await rm(stageRoot, { recursive: true, force: true });
      await rm(consumerRoot, { recursive: true, force: true });
    }
  }, 120_000);
});
