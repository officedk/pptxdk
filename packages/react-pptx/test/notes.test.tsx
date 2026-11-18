import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { posix } from "node:path";

import {
  appendJSX,
  build as buildReact,
  Deck,
  Notes,
  P,
  Slide,
  Span,
  textCapability,
} from "@react-pptx/react-pptx";
import {
  build,
  createDocumentBuild,
  deck,
  box,
  documentCapabilities,
  font,
  notes,
  notesCapability,
  paragraph,
  slide,
  span,
  text,
  type Diagnostic,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
]);

function diagnosticsOf(error: unknown): readonly Diagnostic[] {
  return (
    (error as { readonly diagnostics?: readonly Diagnostic[] }).diagnostics ??
    []
  );
}

async function rejectedDiagnostics(
  operation: Promise<unknown>,
): Promise<readonly Diagnostic[]> {
  try {
    await operation;
  } catch (error) {
    return diagnosticsOf(error);
  }
  assert.fail("Expected the build to fail");
}

function presentationXml(pptx: Uint8Array): string {
  return strFromU8(unzipSync(pptx)["ppt/presentation.xml"]!);
}

describe("rich-text speaker notes", () => {
  it("writes actual notes parts through the command entry and reads the body independently", async () => {
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "intro" }).append(
        text("Projected title", {
          id: "title",
          x: 40,
          y: 40,
          width: 400,
          height: 60,
          style: { fontSize: 48 },
        }),
        notes(
          [
            paragraph("Opening line", { id: "opening" }),
            paragraph([
              span("Emphasis", { style: { fontWeight: "bold" } }),
              " tail",
            ]),
            paragraph(),
          ],
          { id: "speech" },
        ),
      ),
    );

    const result = await build(document);
    const read = await validatePptx(result.pptx);
    assert.deepEqual(read.errors, []);
    const notesRead = read.slides[0]?.notes;
    assert.ok(notesRead);
    assert.equal(notesRead.text.text, "Opening line\nEmphasis tail\n");
    assert.match(notesRead.part, /^ppt\/notesSlides\/notesSlide\d+\.xml$/u);
    assert.match(
      notesRead.masterPart,
      /^ppt\/notesMasters\/notesMaster\d+\.xml$/u,
    );
    assert.notEqual(notesRead.themePart, "ppt/theme/theme1.xml");
    assert.match(notesRead.themePart ?? "", /^ppt\/theme\/theme\d+\.xml$/u);
    assert.deepEqual(notesRead.placeholder, { type: "body", index: 1 });
    assert.equal(notesRead.text.paragraphs[0]?.runs[0]?.fontSizePt, 12);
    assert.equal(notesRead.text.paragraphs[0]?.lineSpacingPoints, undefined);
    assert.equal(notesRead.text.paragraphs[1]?.runs[0]?.bold, true);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Projected title"],
    );
    assert.deepEqual(result.report.slides[0]?.layouts ?? [], []);
    assert.deepEqual(result.report.slides[0]?.capabilities, {
      notes: { text: "Opening line\nEmphasis tail\n", paragraphCount: 3 },
    });
    const parts = unzipSync(result.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.doesNotMatch(slideXml, /Opening line|Emphasis/u);
    assert.match(
      presentationXml(result.pptx),
      /<p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst><p:sldIdLst>/u,
    );
  });

  it("writes an independent notes master theme part while preserving the author theme", async () => {
    const document = deck({
      width: 960,
      height: 540,
      theme: {
        name: "Notes Theme Contract",
        colors: { accent1: "#336699", brand: "#123456" },
        fonts: { major: jost, minor: jost },
      },
    }).append(
      slide({
        id: "notes-theme",
        theme: { colors: { brand: "#ABCDEF" } },
      }).append(
        notes(
          paragraph([
            span("Theme-aware notes", {
              style: { color: "brand" },
            }),
          ]),
        ),
      ),
    );

    const result = await build(document);
    const read = await validatePptx(result.pptx);
    assert.deepEqual(read.errors, []);
    const notesRead = read.slides[0]?.notes;
    assert.ok(notesRead?.themePart);
    const parts = unzipSync(result.pptx);

    const masterRelsPath = posix.join(
      posix.dirname(notesRead.masterPart),
      "_rels",
      `${posix.basename(notesRead.masterPart)}.rels`,
    );
    const masterRels = strFromU8(parts[masterRelsPath]!);
    const masterThemes = [
      ...masterRels.matchAll(/Type="[^"]*\/theme" Target="([^"]+)"\/>/gu),
    ].map((match) => match[1]!);
    assert.equal(masterThemes.length, 1);
    const masterThemePath = posix.join(
      posix.dirname(notesRead.masterPart),
      masterThemes[0]!,
    );
    assert.equal(notesRead.themePart, masterThemePath);

    const slideMasterRels = strFromU8(
      parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"]!,
    );
    const slideThemeTarget = /Type="[^"]*\/theme" Target="([^"]+)"\/>/u.exec(
      slideMasterRels,
    )?.[1];
    assert.ok(slideThemeTarget);
    const slideThemePath = posix.join("ppt/slideMasters", slideThemeTarget);
    assert.notEqual(masterThemePath, slideThemePath);
    assert.ok(parts[masterThemePath]);
    assert.ok(parts[slideThemePath]);

    const theme = strFromU8(parts[masterThemePath]);
    assert.match(theme, /<a:theme[^>]+name="Notes Theme Contract"/u);
    assert.match(theme, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/u);
    assert.match(theme, /<a:minorFont><a:latin typeface="Jost\*"/u);
    assert.equal(theme, strFromU8(parts[slideThemePath]));

    const run = notesRead.text.paragraphs[0]?.runs[0];
    assert.equal(run?.color, "#ABCDEF");
    assert.equal(run?.fontFamily, "Jost*");

    const presentationRels = strFromU8(
      parts["ppt/_rels/presentation.xml.rels"]!,
    );
    assert.doesNotMatch(presentationRels, /Type="[^"]*\/theme"/u);

    const contentTypes = strFromU8(parts["[Content_Types].xml"]!);
    assert.match(
      contentTypes,
      new RegExp(
        `PartName="/${masterThemePath}" ContentType="application/vnd.openxmlformats-officedocument.theme\\+xml"`,
        "u",
      ),
    );
  });

  it("keeps page theme tokens and notes defaults while notes, paragraph, and span styles override in order", async () => {
    const resourceFont = jost;
    const document = deck({
      width: 960,
      height: 540,
      theme: {
        fonts: { major: resourceFont, minor: resourceFont },
        colors: { brand: "#123456" },
      },
    }).append(
      slide({
        id: "themed",
        theme: { colors: { brand: "#ABCDEF" }, fonts: { minor: "Jost" } },
        style: { fontSize: 48, lineHeight: "60pt" },
      }).append(
        notes(
          [
            paragraph("Default"),
            paragraph("Paragraph level", {
              style: {
                fontSize: "18pt",
                align: "center",
                list: { type: "bullet" },
              },
            }),
            paragraph([
              span("Span level", {
                style: { fontSize: "16pt", fontWeight: "bold", color: "dk1" },
              }),
            ]),
          ],
          { id: "speech", style: { fontSize: "20pt", color: "brand" } },
        ),
      ),
    );
    const result = await build(document);
    const read = await validatePptx(result.pptx);
    assert.deepEqual(read.errors, []);
    const notesRead = read.slides[0]!.notes!;
    const paragraphRuns = notesRead.text.paragraphs.map((p) => p.runs[0]);
    assert.equal(paragraphRuns[0]?.fontSizePt, 20);
    assert.equal(paragraphRuns[0]?.color, "#ABCDEF");
    assert.equal(paragraphRuns[0]?.fontFamily, "Jost");
    assert.equal(notesRead.text.paragraphs[0]?.lineSpacingPoints, undefined);
    assert.equal(paragraphRuns[1]?.fontSizePt, 18);
    assert.equal(notesRead.text.paragraphs[1]?.alignment, "ctr");
    assert.deepEqual(notesRead.text.paragraphs[1]?.list, {
      type: "bullet",
      character: "•",
    });
    assert.equal(paragraphRuns[2]?.fontSizePt, 16);
    assert.equal(paragraphRuns[2]?.bold, true);
    assert.equal(paragraphRuns[2]?.color, "#000000");
  });

  it("supports real React composition, handoff editing, and appending Notes JSX to an existing slide", async () => {
    function Opening({ subject }: { readonly subject: string }) {
      return (
        <P id="opening">
          Opening: <Span id="subject">{subject}</Span>
        </P>
      );
    }
    function Speech() {
      return (
        <Notes id="speech" style={{ fontSize: "14pt", color: "accent1" }}>
          <Opening subject="first subject" />
          <P>Second paragraph</P>
        </Notes>
      );
    }
    const reactResult = await buildReact(
      <Deck width={960} height={540}>
        <Slide id="react">
          <Speech />
        </Slide>
      </Deck>,
    );
    const reactRead = await validatePptx(reactResult.pptx);
    assert.deepEqual(reactRead.errors, []);
    assert.equal(
      reactRead.slides[0]?.notes?.text.text,
      "Opening: first subject\nSecond paragraph",
    );

    const page = slide({ id: "appended" });
    const document = deck({ width: 960, height: 540 }).append(page);
    await appendJSX(
      page,
      <Notes id="appended-speech">
        <P>Appended body</P>
      </Notes>,
    );
    const appended = await build(document);
    const appendedRead = await validatePptx(appended.pptx);
    assert.equal(appendedRead.slides[0]?.notes?.text.text, "Appended body");
  });

  it("rejects duplicate notes, illegal children, unknown parameters, missing capability, and duplicate ids", async () => {
    const duplicate = deck().append(
      slide({ id: "duplicate" }).append(notes("one"), notes("two")),
    );
    assert.deepEqual(
      (await rejectedDiagnostics(build(duplicate))).map(({ code }) => code),
      ["duplicate-notes"],
    );

    const hiddenDuplicate = deck().append(
      slide({ id: "hidden-duplicate" }).append(
        box({ style: { display: "none" } }).append(notes("hidden")),
        notes("visible"),
      ),
    );
    assert.deepEqual(
      (await rejectedDiagnostics(build(hiddenDuplicate))).map(
        ({ code }) => code,
      ),
      ["duplicate-notes"],
    );

    const illegal = deck().append(
      slide({ id: "illegal" }).append(notes(text("illegal"))),
    );
    assert.equal(
      (await rejectedDiagnostics(build(illegal)))[0]?.code,
      "unsupported-text-child",
    );

    const Unknown = notes as unknown as (
      content: string,
      props: Record<string, unknown>,
    ) => ReturnType<typeof notes>;
    const unknown = deck().append(
      slide({ id: "unknown" }).append(
        Unknown("x", { id: "n", layout: { x: 1 } }),
      ),
    );
    assert.equal(
      (await rejectedDiagnostics(build(unknown)))[0]?.code,
      "unknown-parameter",
    );

    const trimmed = createDocumentBuild(
      documentCapabilities.filter((module) => module !== notesCapability),
    );
    const missing = deck().append(slide({ id: "missing" }).append(notes("x")));
    assert.equal(
      (await rejectedDiagnostics(trimmed(missing)))[0]?.code,
      "missing-capability",
    );

    const duplicateIds = deck().append(
      slide({ id: "ids" }).append(
        text("canvas", { id: "same", x: 10, y: 10, width: 100, height: 30 }),
        notes("notes", { id: "same" }),
      ),
    );
    assert.equal(
      (await rejectedDiagnostics(build(duplicateIds)))[0]?.code,
      "duplicate-author-id",
    );
  });

  it("diagnoses hidden page-affiliated notes when the Notes capability is trimmed", async () => {
    const document = deck().append(
      slide({ id: "hidden" }).append(
        box({ style: { display: "none" } }).append(notes("hidden notes")),
      ),
    );
    const full = await build(document);
    const fullRead = await validatePptx(full.pptx);
    assert.deepEqual(fullRead.errors, []);
    assert.equal(fullRead.slides[0]?.notes?.text.text, "hidden notes");
    assert.deepEqual(full.report.slides[0]?.objects, []);

    const trimmed = createDocumentBuild(
      documentCapabilities.filter((module) => module !== notesCapability),
    );
    assert.deepEqual(
      (await rejectedDiagnostics(trimmed(document))).map(({ code, path }) => ({
        code,
        path,
      })),
      [
        {
          code: "missing-capability",
          path: "deck/slide[hidden]/box[1]/notes[1]",
        },
      ],
    );
  });

  it("keeps pruning hidden canvas content whose capability is trimmed", async () => {
    const trimmedText = createDocumentBuild(
      documentCapabilities.filter((module) => module !== textCapability),
    );
    const document = deck().append(
      slide({ id: "hidden-canvas" }).append(
        box({ style: { display: "none" } }).append(
          text("hidden ordinary content"),
        ),
      ),
    );
    const result = await trimmedText(document);
    assert.deepEqual(result.report.slides[0]?.objects, []);
    assert.deepEqual(result.report.diagnostics, []);
  });

  it("builds notes from the shared text service when the Text capability is trimmed", async () => {
    const notesOnlyBuild = createDocumentBuild(
      documentCapabilities.filter((module) => module !== textCapability),
    );
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "notes-only" }).append(
        notes([paragraph("Only notes"), paragraph()], { id: "speech" }),
      ),
    );
    const result = await notesOnlyBuild(document);
    const read = await validatePptx(result.pptx);
    assert.deepEqual(read.errors, []);
    assert.equal(read.slides[0]?.notes?.text.text, "Only notes\n");
    assert.deepEqual(result.report.slides[0]?.objects, []);
  });

  it("rebuilds moved and edited notes while the previous build stays fixed", async () => {
    const cue = span("first cue");
    const speech = notes(paragraph([cue]), { id: "speech" });
    const firstPage = slide({ id: "first" }).append(speech);
    const secondPage = slide({ id: "second" });
    const document = deck({ width: 960, height: 540 }).append(
      firstPage,
      secondPage,
    );
    const first = await build(document);
    const firstBytes = new Uint8Array(first.pptx);
    const firstRead = await validatePptx(first.pptx);
    assert.equal(firstRead.slides[0]?.notes?.text.text, "first cue");
    assert.equal(firstRead.slides[1]?.notes, undefined);

    cue.content("revised cue");
    speech.moveTo(secondPage);
    const second = await build(document);
    const secondRead = await validatePptx(second.pptx);
    assert.equal(secondRead.slides[0]?.notes, undefined);
    assert.equal(secondRead.slides[1]?.notes?.text.text, "revised cue");
    assert.deepEqual(new Uint8Array(first.pptx), firstBytes);
    const reread = await validatePptx(first.pptx);
    assert.equal(reread.slides[0]?.notes?.text.text, "first cue");
    assert.equal(reread.slides[1]?.notes, undefined);
  });

  it("adds notes-only font characters to the formal embedded font resources", async () => {
    const document = deck({
      width: 960,
      height: 540,
    }).append(
      slide({ id: "fonts" }).append(
        notes(
          paragraph([
            span("QZQX notes-only glyphs", {
              style: { fontFamily: jost },
            }),
          ]),
        ),
      ),
    );
    const result = await build(document, { embedFonts: "subset" });
    const read = await validatePptx(result.pptx);
    assert.deepEqual(read.errors, []);
    assert.equal(
      read.slides[0]?.notes?.text.paragraphs[0]?.runs[0]?.fontFamily,
      "Jost*",
    );
    const parts = unzipSync(result.pptx);
    const fontParts = Object.keys(parts).filter((part) =>
      /^ppt\/fonts\/font\d+\.fntdata$/u.test(part),
    );
    assert.equal(fontParts.length, 1);
    assert.match(
      strFromU8(parts["ppt/presentation.xml"]!),
      /<p:embeddedFont><p:font typeface="Jost\*"/u,
    );
  });
});
