import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { mkdir, mkdtemp, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import {
  CodegenError,
  build,
  codegen,
  importPptx,
  type CodegenResult,
} from "pptxdk";
import { previewFromSource, type PreviewResult } from "pptxdk/node";
import {
  Deck,
  Notes,
  P,
  Section,
  Slide,
  SlideNumber,
  Span,
  Text,
  expandDocument,
} from "pptxdk/react";
import {
  deck,
  defineTheme,
  font,
  notes,
  paragraph,
  section,
  slide,
  slideNumber,
  span,
  text,
} from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
  type GeneratedSourceTypecheckCase,
} from "./helpers/codegen-acceptance.ts";

const packageRoot = fileURLToPath(new URL("../", import.meta.url));
const tscBin = join(packageRoot, "node_modules/.bin/tsc");

/**
 * Checks generated sources containing `import.meta.url` asset references under strict `tsc`
 * with `"type": "module"` enabled for `module: "nodenext"` resolution.
 */
async function verifyEsmGeneratedSourcesTypecheck(
  cases: readonly GeneratedSourceTypecheckCase[],
): Promise<void> {
  if (cases.length === 0) return;

  const dir = await mkdtemp(join(tmpdir(), "pptxdk-codegen-esm-tsc-"));
  try {
    const nodeModulesDir = join(dir, "node_modules");
    await mkdir(join(nodeModulesDir, "@types"), { recursive: true });
    await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
    await symlink(
      join(packageRoot, "node_modules/react"),
      join(nodeModulesDir, "react"),
      "dir",
    );
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

    const caseManifest: string[] = [];
    for (const [index, item] of cases.entries()) {
      await writeCodegenFiles(
        join(dir, "cases", `case-${index}`),
        item.result.files,
      );
      caseManifest.push(`case-${index}: ${item.label} (${item.result.entry})`);
    }

    await writeFile(
      join(dir, "package.json"),
      JSON.stringify({
        name: "codegen-esm-consumer",
        private: true,
        type: "module",
      }),
    );
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
      { cwd: dir, encoding: "utf8" },
    );
    assert.equal(
      tscResult.status,
      0,
      `Consumer strict ESM tsc failed for:\n${caseManifest.join("\n")}\n\n${tscResult.stdout}\n${tscResult.stderr}`,
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function previewGeneratedSourceRelocated(
  result: CodegenResult,
): Promise<PreviewResult> {
  const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-codegen-preview-"));
  try {
    const consumerRoot = join(tempRoot, "consumer");
    const nodeModulesDir = join(consumerRoot, "node_modules");
    await mkdir(nodeModulesDir, { recursive: true });
    await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
    await symlink(
      join(packageRoot, "node_modules/react"),
      join(nodeModulesDir, "react"),
      "dir",
    );

    const originDir = join(tempRoot, "origin");
    await writeCodegenFiles(originDir, result.files);
    const relocatedDir = join(consumerRoot, "relocated");
    const { rename } = await import("node:fs/promises");
    await rename(originDir, relocatedDir);

    return await previewFromSource(join(relocatedDir, result.entry));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostItalicBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Italic.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [
  { weight: 400, style: "normal", data: jostBookBytes },
  { weight: 700, style: "normal", data: jostBoldBytes },
]);

const jostItalicFont = font("Jost", [
  { weight: 400, style: "italic", data: jostItalicBytes },
]);

function withSyntheticFontTable(bytes: Uint8Array, tag: string): Uint8Array {
  const source = Buffer.from(bytes);
  const numTables = source.readUInt16BE(4);
  const headerSize = 12 + numTables * 16;
  const nextNumTables = numTables + 1;
  const nextHeaderSize = 12 + nextNumTables * 16;
  const records: {
    tag: string;
    checkSum: number;
    offset: number;
    length: number;
  }[] = [];
  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    records.push({
      tag: source.subarray(off, off + 4).toString("ascii"),
      checkSum: source.readUInt32BE(off + 4),
      offset: source.readUInt32BE(off + 8) + 16,
      length: source.readUInt32BE(off + 12),
    });
  }
  records.push({
    tag,
    checkSum: 0,
    offset: source.length + 16,
    length: 16,
  });
  records.sort((a, b) => (a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0));
  const out = Buffer.alloc(source.length + 32);
  source.copy(out, 0, 0, 12);
  out.writeUInt16BE(nextNumTables, 4);
  const entrySelector = Math.floor(Math.log2(nextNumTables));
  const searchRange = (1 << entrySelector) * 16;
  out.writeUInt16BE(searchRange, 6);
  out.writeUInt16BE(entrySelector, 8);
  out.writeUInt16BE(nextNumTables * 16 - searchRange, 10);
  for (let i = 0; i < records.length; i++) {
    const off = 12 + i * 16;
    out.write(records[i]!.tag, off, 4, "ascii");
    out.writeUInt32BE(records[i]!.checkSum, off + 4);
    out.writeUInt32BE(records[i]!.offset, off + 8);
    out.writeUInt32BE(records[i]!.length, off + 12);
  }
  source.copy(out, nextHeaderSize, headerSize);
  return new Uint8Array(out);
}

const colrJostFont = font(
  "ColrJost",
  [{ weight: 400, data: withSyntheticFontTable(jostBookBytes, "COLR") }],
  { substitute: true },
);

const sbixJostFont = font(
  "SbixJost",
  [{ weight: 400, data: withSyntheticFontTable(jostBookBytes, "sbix") }],
  { substitute: true },
);

const embeddedTheme = defineTheme({
  name: "Editorial Jost Theme",
  fonts: { major: jostFont, minor: jostFont },
});

function createRichTextAndNotesVanillaDeck() {
  return deck({ width: 640, height: 360, theme: embeddedTheme }).append(
    section({ name: "Opening" }).append(
      slide({ id: "rich" }).append(
        text(
          [
            paragraph(
              [
                span("First ", {
                  style: {
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    fontSize: "18pt",
                    letterSpacing: 2,
                    textDecoration: "underline",
                  },
                }),
                span("second", {
                  style: {
                    fontFamily: jostItalicFont,
                    fontStyle: "italic",
                    color: "#DC2626",
                    fontSize: "14pt",
                    textDecoration: "line-through",
                    textTransform: "uppercase",
                    textShadow: {
                      color: "#0F172A",
                      blur: 2,
                      distance: 2,
                      angle: 45,
                    },
                  },
                }),
              ],
              {
                style: {
                  align: "center",
                  lineHeight: "24pt",
                  spacingAfter: 10,
                  lang: "en-US",
                },
              },
            ),
            paragraph(),
            paragraph("Bullet\nline", {
              style: { list: { type: "bullet", character: "+", level: 1 } },
            }),
            paragraph(["Item ", slideNumber()], {
              style: {
                list: {
                  type: "number",
                  style: "arabicPeriod",
                  startAt: 5,
                  level: 2,
                },
              },
            }),
          ],
          {
            id: "rich-card",
            name: "RichCard",
            x: 24,
            y: 24,
            width: 420,
            height: 220,
            link: { url: "https://example.com/spec" },
            style: { fontFamily: "minor", fontSize: "14pt", color: "#0F172A" },
          },
        ),
        text("Jump to details", {
          id: "jump-btn",
          x: 464,
          y: 24,
          width: 152,
          height: 44,
          link: { id: "details" },
          style: { fontFamily: "minor", fontSize: "14pt", color: "#1E3A8A" },
        }),
        notes(
          [
            paragraph("Opening cue"),
            paragraph([
              span("Key point: ", {
                style: { fontWeight: "bold", color: "#D97706" },
              }),
              "keep notes off canvas.",
            ]),
            paragraph(),
            paragraph(["Slide ", slideNumber()], {
              style: { list: { type: "bullet" }, align: "center" },
            }),
          ],
          { id: "rich-notes", style: { fontSize: "13pt" } },
        ),
      ),
      slide({ id: "details", hidden: true }).append(
        text(["Home from page ", slideNumber()], {
          id: "home-box",
          x: 32,
          y: 32,
          width: 260,
          height: 48,
          link: { action: "home" },
          style: { fontFamily: "minor", fontSize: "16pt", color: "#111827" },
        }),
        notes("Hidden slide cue", { id: "details-notes" }),
      ),
    ),
    section({ name: "Appendix" }).append(
      slide({ id: "no-notes-page" }).append(
        text("Canvas without notes", {
          id: "plain-box",
          x: 32,
          y: 32,
          width: 260,
          height: 48,
          style: { fontFamily: "minor", fontSize: "16pt", color: "#334155" },
        }),
      ),
    ),
  );
}

async function createRichTextAndNotesReactDeck() {
  return expandDocument(
    createElement(
      Deck,
      { width: 640, height: 360, theme: embeddedTheme },
      createElement(
        Section,
        { name: "Opening" },
        createElement(
          Slide,
          { id: "rich" },
          createElement(
            Text,
            {
              id: "rich-card",
              name: "RichCard",
              x: 24,
              y: 24,
              width: 420,
              height: 220,
              link: { url: "https://example.com/spec" },
              style: {
                fontFamily: "minor",
                fontSize: "14pt",
                color: "#0F172A",
              },
            },
            createElement(
              P,
              {
                style: {
                  align: "center",
                  lineHeight: "24pt",
                  spacingAfter: 10,
                  lang: "en-US",
                },
              },
              createElement(
                Span,
                {
                  style: {
                    fontWeight: "bold",
                    color: "#1D4ED8",
                    fontSize: "18pt",
                    letterSpacing: 2,
                    textDecoration: "underline",
                  },
                },
                "First ",
              ),
              createElement(
                Span,
                {
                  style: {
                    fontFamily: jostItalicFont,
                    fontStyle: "italic",
                    color: "#DC2626",
                    fontSize: "14pt",
                    textDecoration: "line-through",
                    textTransform: "uppercase",
                    textShadow: {
                      color: "#0F172A",
                      blur: 2,
                      distance: 2,
                      angle: 45,
                    },
                  },
                },
                "second",
              ),
            ),
            createElement(P, null),
            createElement(
              P,
              { style: { list: { type: "bullet", character: "+", level: 1 } } },
              "Bullet\nline",
            ),
            createElement(
              P,
              {
                style: {
                  list: {
                    type: "number",
                    style: "arabicPeriod",
                    startAt: 5,
                    level: 2,
                  },
                },
              },
              "Item ",
              createElement(SlideNumber, null),
            ),
          ),
          createElement(
            Text,
            {
              id: "jump-btn",
              x: 464,
              y: 24,
              width: 152,
              height: 44,
              link: { id: "details" },
              style: {
                fontFamily: "minor",
                fontSize: "14pt",
                color: "#1E3A8A",
              },
            },
            "Jump to details",
          ),
          createElement(
            Notes,
            { id: "rich-notes", style: { fontSize: "13pt" } },
            createElement(P, null, "Opening cue"),
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontWeight: "bold", color: "#D97706" } },
                "Key point: ",
              ),
              "keep notes off canvas.",
            ),
            createElement(P, null),
            createElement(
              P,
              { style: { list: { type: "bullet" }, align: "center" } },
              "Slide ",
              createElement(SlideNumber, null),
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "details", hidden: true },
          createElement(
            Text,
            {
              id: "home-box",
              x: 32,
              y: 32,
              width: 260,
              height: 48,
              link: { action: "home" },
              style: {
                fontFamily: "minor",
                fontSize: "16pt",
                color: "#111827",
              },
            },
            "Home from page ",
            createElement(SlideNumber, null),
          ),
          createElement(
            Notes,
            { id: "details-notes" },
            createElement(P, null, "Hidden slide cue"),
          ),
        ),
      ),
      createElement(
        Section,
        { name: "Appendix" },
        createElement(
          Slide,
          { id: "no-notes-page" },
          createElement(
            Text,
            {
              id: "plain-box",
              x: 32,
              y: 32,
              width: 260,
              height: 48,
              style: {
                fontFamily: "minor",
                fontSize: "16pt",
                color: "#334155",
              },
            },
            "Canvas without notes",
          ),
        ),
      ),
    ),
  );
}

describe("pptxdk rich-text, fields, links, and speaker notes codegen acceptance", () => {
  it("preserves rich text, empty paragraphs, lists, slide numbers, links, embedded fonts, and isolated speaker notes across direct and imported vanilla/react roundtrips", async () => {
    const vanillaDoc = createRichTextAndNotesVanillaDeck();
    const reactDoc = await createRichTextAndNotesReactDeck();

    const originalBuilt = await build(vanillaDoc, { embedFonts: "full" });
    assert.deepEqual(
      originalBuilt.report.diagnostics.filter((d) => d.severity === "error"),
      [],
    );
    const pptxBytesCopy = new Uint8Array(originalBuilt.pptx);
    const imported = importPptx(pptxBytesCopy);
    pptxBytesCopy.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const sources = [
      { label: "direct-vanilla", doc: vanillaDoc },
      { label: "direct-react", doc: reactDoc },
      { label: "imported-pptx", doc: imported.document },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc }) => [
      {
        label: `${label}:vanilla`,
        result: codegen(doc, { target: "vanilla" }),
      },
      { label: `${label}:react`, result: codegen(doc, { target: "react" }) },
    ]);

    await verifyEsmGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );
      assert.equal(file.slides.length, 3, `${item.label} slide count`);
      assert.deepEqual(
        file.presentation.sections?.map((sec) => ({
          name: sec.name,
          slideIds: sec.slideIds,
        })),
        [
          { name: "Opening", slideIds: ["rich", "details"] },
          { name: "Appendix", slideIds: ["no-notes-page"] },
        ],
        `${item.label} sections`,
      );

      const s1 = file.slides[0]!;
      assert.equal(
        s1.objects.length,
        2,
        `${item.label} slide 1 canvas objects`,
      );
      const card = s1.objects[0]!;
      assert.equal(card.name, "RichCard", `${item.label} card name`);
      assert.deepEqual(card.link, { url: "https://example.com/spec" });
      assert.ok(card.text !== undefined);
      assert.equal(card.text.paragraphs.length, 4, `${item.label} para count`);

      const p0 = card.text.paragraphs[0]!;
      assert.equal(p0.alignment, "ctr");
      assert.equal(p0.lineSpacingPoints, 24);
      assert.equal(p0.spacingAfterPoints, 7.5);
      assert.equal(p0.runs.length, 2);
      assert.equal(p0.runs[0]?.text, "First ");
      assert.equal(p0.runs[0]?.bold, true);
      assert.equal(p0.runs[0]?.color, "#1D4ED8");
      assert.equal(p0.runs[0]?.fontSizePt, 18);
      assert.equal(p0.runs[1]?.text, "second");
      assert.equal(p0.runs[1]?.italic, true);
      assert.equal(p0.runs[1]?.color, "#DC2626");
      assert.equal(p0.runs[1]?.fontSizePt, 14);
      assert.equal(p0.runs[1]?.fontFamily, "Jost");

      const p1Empty = card.text.paragraphs[1]!;
      assert.equal(p1Empty.runs.length, 0, `${item.label} empty para kept`);
      assert.equal(p1Empty.text, "");

      const p2Bullet = card.text.paragraphs[2]!;
      assert.deepEqual(p2Bullet.list, { type: "bullet", character: "+" });
      assert.equal(p2Bullet.text, "Bullet\nline");

      const p3Number = card.text.paragraphs[3]!;
      assert.deepEqual(p3Number.list, {
        type: "number",
        style: "arabicPeriod",
        startAt: 5,
      });
      assert.equal(p3Number.runs[1]?.field, "slidenum");
      assert.equal(p3Number.runs[1]?.text, "1");

      assert.deepEqual(s1.objects[1]?.link, { slideId: "details" });

      // Notes isolation and formatting on Slide 1
      assert.ok(s1.notes !== undefined, `${item.label} slide 1 notes`);
      assert.equal(
        s1.notes.text.text,
        "Opening cue\nKey point: keep notes off canvas.\n\nSlide 1",
      );
      assert.equal(s1.notes.text.paragraphs[0]?.runs[0]?.fontSizePt, 13);
      assert.equal(s1.notes.text.paragraphs[1]?.runs[0]?.bold, true);
      assert.equal(s1.notes.text.paragraphs[1]?.runs[0]?.color, "#D97706");
      assert.equal(s1.notes.text.paragraphs[2]?.runs.length, 0);
      assert.equal(s1.notes.text.paragraphs[3]?.alignment, "ctr");
      assert.deepEqual(s1.notes.text.paragraphs[3]?.list, {
        type: "bullet",
        character: "•",
      });
      assert.equal(s1.notes.text.paragraphs[3]?.runs[1]?.field, "slidenum");
      assert.equal(s1.notes.text.paragraphs[3]?.runs[1]?.text, "1");

      // Slide 2: hidden + home link + slideNumber + notes
      const s2 = file.slides[1]!;
      assert.equal(s2.hidden, true);
      assert.deepEqual(s2.objects[0]?.link, { action: "home" });
      assert.equal(
        s2.objects[0]?.text?.paragraphs[0]?.runs[1]?.field,
        "slidenum",
      );
      assert.equal(s2.objects[0]?.text?.paragraphs[0]?.runs[1]?.text, "2");
      assert.equal(s2.notes?.text.text, "Hidden slide cue");

      // Slide 3: no notes
      assert.equal(file.slides[2]?.notes, undefined);

      // Verify slide XML contains zero notes text and preserves list levels / decorations
      const parts = unzipSync(built.pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      assert.doesNotMatch(
        slide1Xml,
        /Opening cue|Key point|keep notes off canvas/u,
      );
      assert.match(slide1Xml, /lvl="1"/u);
      assert.match(slide1Xml, /lvl="2"/u);
      assert.match(slide1Xml, /u="sng"/u);
      assert.match(slide1Xml, /strike="sngStrike"/u);
      assert.match(slide1Xml, /cap="all"/u);
      assert.match(slide1Xml, /spc="150"/u);
      assert.match(slide1Xml, /<a:outerShdw /u);

      // Preview renders cleanly without leaking notes onto canvas
      const painted = await previewGeneratedSourceRelocated(item.result);
      assert.equal(painted.pages.length, 3);
      for (const [pageIdx, page] of painted.pages.entries()) {
        assert.deepEqual(
          page.diagnostics,
          [],
          `${item.label} preview page ${pageIdx + 1}`,
        );
        assert.ok(page.png instanceof Uint8Array);
        assert.equal(page.svg?.includes("Opening cue"), false);
        assert.equal(page.svg?.includes("Hidden slide cue"), false);
      }
    }
  }, 60_000);

  it("allows editing a single Span in generated vanilla and react source while keeping adjacent sibling Span formatting intact", async () => {
    const vanillaDoc = createRichTextAndNotesVanillaDeck();
    const built = await build(vanillaDoc, { embedFonts: "full" });
    const bytesCopy = new Uint8Array(built.pptx);
    const imported = importPptx(bytesCopy);
    bytesCopy.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(imported.document, { target });
      const sourceText = generated.files[generated.entry] as string;

      assert.ok(sourceText.includes("First "));
      assert.ok(sourceText.includes("#1D4ED8"));
      assert.ok(sourceText.includes("second"));
      assert.ok(sourceText.includes("#DC2626"));

      const editedSource = sourceText
        .replace("First ", "Edited ")
        .replace("#1D4ED8", "#059669");
      assert.notEqual(editedSource, sourceText);

      const editedResult = {
        target,
        entry: generated.entry,
        files: {
          ...generated.files,
          [generated.entry]: editedSource,
        },
      } as typeof generated;

      await verifyEsmGeneratedSourcesTypecheck([
        { label: `edited-${target}`, result: editedResult },
      ]);
      const rebuilt = await rebuildGeneratedSourceRelocated(editedResult);
      assert.deepEqual(rebuilt.file.errors, []);

      const editedRuns =
        rebuilt.file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs;
      assert.equal(editedRuns?.[0]?.text, "Edited ");
      assert.equal(editedRuns?.[0]?.color, "#059669");
      assert.equal(editedRuns?.[0]?.bold, true);
      assert.equal(editedRuns?.[0]?.fontSizePt, 18);

      // Adjacent sibling Span retains every original local style and text
      assert.equal(editedRuns?.[1]?.text, "second");
      assert.equal(editedRuns?.[1]?.italic, true);
      assert.equal(editedRuns?.[1]?.color, "#DC2626");
      assert.equal(editedRuns?.[1]?.fontSizePt, 14);
      assert.equal(editedRuns?.[1]?.fontFamily, "Jost");

      // Notes and empty paragraph remain untouched
      assert.equal(
        rebuilt.file.slides[0]?.objects[0]?.text?.paragraphs[1]?.runs.length,
        0,
      );
      assert.equal(
        rebuilt.file.slides[0]?.notes?.text.text,
        "Opening cue\nKey point: keep notes off canvas.\n\nSlide 1",
      );
    }
  }, 60_000);

  it("recovers unembedded font declarations and partial corrupt XML imports from file bytes alone", async () => {
    const unembeddedDeck = deck({ width: 640, height: 360 }).append(
      slide({ id: "units" }).append(
        text(
          [
            paragraph([
              span("32pt custom", {
                style: { fontSize: "32pt", fontFamily: "CustomSans" },
              }),
              span("32px major", {
                style: { fontSize: "32px", fontFamily: "major" },
              }),
            ]),
          ],
          { id: "sizes", x: 24, y: 24, width: 360, height: 160 },
        ),
        notes("File-only notes cue", { id: "file-notes" }),
      ),
    );
    const built = await build(unembeddedDeck);
    const entries = unzipSync(built.pptx);
    const slideXml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    // Inject unsupported @baseline attribute on a third run to trigger a localized import diagnostic
    entries["ppt/slides/slide1.xml"] = strToU8(
      slideXml.replace(
        "</a:p>",
        '<a:r><a:rPr sz="2000" baseline="30000"/><a:t>Sup</a:t></a:r></a:p>',
      ),
    );
    const corruptBytes = zipSync(entries);
    const copy = new Uint8Array(corruptBytes);
    const imported = importPptx(copy);
    copy.fill(0);

    assert.equal(imported.report.diagnostics.length, 1);
    assert.equal(imported.report.diagnostics[0]?.code, "unconsumed-xml");

    const generated = [
      {
        label: "unembedded:vanilla",
        result: codegen(imported.document, { target: "vanilla" }),
      },
      {
        label: "unembedded:react",
        result: codegen(imported.document, { target: "react" }),
      },
    ];
    await verifyGeneratedSourcesTypecheck(generated);

    for (const item of generated) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, []);
      const runs = file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs;
      assert.equal(runs?.[0]?.fontSizePt, 32);
      assert.equal(runs?.[0]?.fontFamily, "CustomSans");
      assert.equal(runs?.[1]?.fontSizePt, 24);
      assert.equal(runs?.[1]?.fontFamily, "Aptos Display");
      assert.equal(runs?.[2]?.text, "Sup");
      assert.equal(file.slides[0]?.notes?.text.text, "File-only notes cue");
    }
  }, 60_000);

  it("preserves preview rejection diagnostics, target text ids, notes preservation, and healthy sibling pages across codegen", async () => {
    const rejectionDeck = deck({ width: 480, height: 270 }).append(
      slide({ id: "text-rejections" }).append(
        text("汉", {
          id: "missing-glyph",
          x: 10,
          y: 10,
          width: 180,
          height: 40,
          style: { fontFamily: jostFont },
        }),
        text(
          [
            paragraph("Item", {
              style: { list: { type: "bullet", character: "汉" } },
            }),
          ],
          {
            id: "missing-bullet-glyph",
            x: 10,
            y: 60,
            width: 180,
            height: 40,
            style: { fontFamily: jostFont },
          },
        ),
        text("Color", {
          id: "colr-font",
          x: 220,
          y: 10,
          width: 180,
          height: 40,
          style: { fontFamily: colrJostFont },
        }),
        text("Bitmap", {
          id: "sbix-font",
          x: 220,
          y: 60,
          width: 180,
          height: 40,
          style: { fontFamily: sbixJostFont },
        }),
        notes("Notes survive preview rejection", { id: "preserved-notes" }),
      ),
      slide({ id: "valid-sibling" }).append(
        text("Valid sibling text", {
          id: "ok-text",
          x: 40,
          y: 60,
          width: 240,
          height: 60,
          style: {
            fontFamily: jostFont,
            fontSize: 22,
            padding: 12,
            verticalAlign: "middle",
            backgroundColor: "#EFF6FF",
            borderColor: "#2563EB",
            borderWidth: 2,
          },
        }),
        notes("Valid page notes", { id: "valid-notes" }),
      ),
    );

    const cases = [
      {
        label: "reject:vanilla",
        result: codegen(rejectionDeck, { target: "vanilla" }),
      },
      {
        label: "reject:react",
        result: codegen(rejectionDeck, { target: "react" }),
      },
    ];
    await verifyEsmGeneratedSourcesTypecheck(cases);

    for (const item of cases) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, []);
      assert.equal(file.slides.length, 2);
      assert.equal(
        file.slides[0]?.notes?.text.text,
        "Notes survive preview rejection",
      );
      assert.equal(file.slides[1]?.notes?.text.text, "Valid page notes");

      const painted = await previewGeneratedSourceRelocated(item.result);
      const p1 = painted.pages[0]!;
      assert.equal(p1.png, undefined);
      assert.equal(p1.diagnostics.length, 4);
      for (const id of [
        "missing-glyph",
        "missing-bullet-glyph",
        "colr-font",
        "sbix-font",
      ]) {
        assert.ok(
          p1.diagnostics.some(
            (d) => d.severity === "error" && d.path.includes(`text[${id}]`),
          ),
          `${item.label}: missing diagnostic for ${id}`,
        );
      }

      const p2 = painted.pages[1]!;
      assert.deepEqual(p2.diagnostics, []);
      assert.ok(p2.png instanceof Uint8Array);
      assert.equal(p2.svg?.includes("Valid page notes"), false);
    }

    // Structural codegen rejection at public SDK entrypoint
    const badNotes = notes(text("illegal", { id: "inner" }), { id: "n-bad" });
    const badDoc = deck().append(slide({ id: "s-bad" }).append(badNotes));
    assert.throws(
      () => codegen(badDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-bad]/notes[n-bad]/text[inner]",
    );
  }, 60_000);
});
