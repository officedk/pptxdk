import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  font,
  notes,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { Notes } from "@pptxdk/capabilities/notes-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

function pixelAt(
  pngBytes: Uint8Array,
  x: number,
  y: number,
): readonly [number, number, number, number] {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  return [
    decoded.data[offset]!,
    decoded.data[offset + 1]!,
    decoded.data[offset + 2]!,
    decoded.data[offset + 3]!,
  ];
}

export const notesPositiveScenario: CapabilityScenario = {
  id: "notes-rich-text-and-canvas-isolation",
  capability: "notes",
  description:
    "Dual-author Notes scenario pairing rich-text speaker notes (paragraphs, bold/themed span, bullet list, SlideNumber) with Shape canvas pages.",
  visualCheckpoints: [
    "Page 1 (intro-notes): clean #F8FAFC canvas with navy rectangle (#1E3A8A) on the left and amber rectangle (#D97706) on the right; lower half is pure #F8FAFC with zero speaker notes text leaked onto the canvas.",
    "Page 2 (no-notes-page): soft indigo canvas (#EEF2FF) with centered indigo ellipse (#4338CA) and clean margins.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint (x64) and WPS (x86): inspect Slide 1 notesPage to verify all 3 speaker notes lines ('Speaker cue: introduce the architecture.', 'Key takeaway: keep notes off the visible canvas.', 'Slide index 1') and Slide 2 has empty notes.",
    "Edit Slide 1 speaker notes via editNotes, saveAs a new PPTX, close and reopen to verify persistence in PowerPoint and WPS, then validate the re-saved PPTX with validatePptx.",
  ],
  jsx() {
    return createElement(
      Deck,
      {
        width: 640,
        height: 360,
        style: { fontSize: 36, color: "#991B1B" },
        theme: {
          colors: {
            brandNavy: "#1E3A8A",
            highlight: "#F59E0B",
          },
        },
      },
      createElement(
        Slide,
        {
          id: "intro-notes",
          background: "#F8FAFC",
          theme: { colors: { highlight: "#D97706" } },
        },
        createElement(Shape, {
          id: "navy-card",
          x: 48,
          y: 48,
          width: 240,
          height: 110,
          fill: "brandNavy",
        }),
        createElement(Shape, {
          id: "amber-badge",
          x: 336,
          y: 48,
          width: 240,
          height: 110,
          fill: "highlight",
        }),
        createElement(
          Notes,
          { id: "intro-speech" },
          createElement(
            P,
            { id: "cue-line" },
            "Speaker cue: introduce the architecture.",
          ),
          createElement(
            P,
            { id: "takeaway-line" },
            createElement(
              Span,
              { style: { fontWeight: "bold", color: "highlight" } },
              "Key takeaway: ",
            ),
            "keep notes off the visible canvas.",
          ),
          createElement(
            P,
            {
              id: "bullet-line",
              style: { list: { type: "bullet" }, align: "center" },
            },
            "Slide index ",
            createElement(SlideNumber, null),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "no-notes-page", background: "#EEF2FF" },
        createElement(Shape, {
          id: "summary-oval",
          geometry: "ellipse",
          x: 200,
          y: 90,
          width: 240,
          height: 180,
          fill: "#4338CA",
        }),
      ),
    );
  },
  document() {
    return deck({
      width: 640,
      height: 360,
      style: { fontSize: 36, color: "#991B1B" },
      theme: {
        colors: {
          brandNavy: "#1E3A8A",
          highlight: "#F59E0B",
        },
      },
    }).append(
      slide({
        id: "intro-notes",
        background: "#F8FAFC",
        theme: { colors: { highlight: "#D97706" } },
      }).append(
        shape({
          id: "navy-card",
          x: 48,
          y: 48,
          width: 240,
          height: 110,
          fill: "brandNavy",
        }),
        shape({
          id: "amber-badge",
          x: 336,
          y: 48,
          width: 240,
          height: 110,
          fill: "highlight",
        }),
        notes(
          [
            paragraph("Speaker cue: introduce the architecture.", {
              id: "cue-line",
            }),
            paragraph(
              [
                span("Key takeaway: ", {
                  style: { fontWeight: "bold", color: "highlight" },
                }),
                "keep notes off the visible canvas.",
              ],
              { id: "takeaway-line" },
            ),
            paragraph(["Slide index ", slideNumber()], {
              id: "bullet-line",
              style: { list: { type: "bullet" }, align: "center" },
            }),
          ],
          { id: "intro-speech" },
        ),
      ),
      slide({ id: "no-notes-page", background: "#EEF2FF" }).append(
        shape({
          id: "summary-oval",
          geometry: "ellipse",
          x: 200,
          y: 90,
          width: 240,
          height: 180,
          fill: "#4338CA",
        }),
      ),
    );
  },
  assert({ author, resolution, pptx, preview, validation }) {
    const expectedText =
      "Speaker cue: introduce the architecture.\nKey takeaway: keep notes off the visible canvas.\nSlide index 1";

    assert.equal(preview.pages.length, 2, author);
    assert.deepEqual(preview.pages[0]?.diagnostics, [], author);
    assert.deepEqual(preview.pages[1]?.diagnostics, [], author);

    // Report observations: only visible canvas shapes are listed under objects; notes stay page-affiliated.
    const reportSlide1 = resolution.report.slides[0]!;
    assert.equal(reportSlide1.objects.length, 2, author);
    assert.deepEqual(reportSlide1.layouts ?? [], [], author);
    assert.deepEqual(
      reportSlide1.capabilities,
      { notes: { text: expectedText, paragraphCount: 3 } },
      author,
    );
    const reportSlide2 = resolution.report.slides[1]!;
    assert.equal(reportSlide2.objects.length, 1, author);
    assert.equal(reportSlide2.capabilities?.["notes"], undefined, author);

    // Independent PPTX notes validation: content, formatting, master/theme parts, and presentation list.
    const fileNotes = validation.slides[0]?.notes;
    assert.ok(fileNotes, author);
    assert.equal(fileNotes.text.text, expectedText, author);
    assert.match(
      fileNotes.part,
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/u,
      author,
    );
    assert.match(
      fileNotes.masterPart,
      /^ppt\/notesMasters\/notesMaster\d+\.xml$/u,
      author,
    );
    assert.notEqual(fileNotes.themePart, "ppt/theme/theme1.xml", author);
    assert.match(
      fileNotes.themePart ?? "",
      /^ppt\/theme\/theme\d+\.xml$/u,
      author,
    );
    assert.deepEqual(fileNotes.placeholder, { type: "body", index: 1 }, author);
    assert.equal(fileNotes.text.paragraphs[0]?.runs[0]?.fontSizePt, 12, author);
    assert.equal(fileNotes.text.paragraphs[1]?.runs[0]?.bold, true, author);
    assert.equal(
      fileNotes.text.paragraphs[1]?.runs[0]?.color,
      "#D97706",
      author,
    );
    assert.equal(fileNotes.text.paragraphs[2]?.alignment, "ctr", author);
    assert.deepEqual(
      fileNotes.text.paragraphs[2]?.list,
      { type: "bullet", character: "•" },
      author,
    );
    assert.equal(
      fileNotes.text.paragraphs[2]?.runs[1]?.field,
      "slidenum",
      author,
    );
    assert.equal(fileNotes.text.paragraphs[2]?.runs[1]?.text, "1", author);
    assert.equal(validation.slides[1]?.notes, undefined, author);

    const parts = unzipSync(pptx);
    const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.doesNotMatch(
      slide1Xml,
      /Speaker cue|Key takeaway|Slide index/u,
      author,
    );
    assert.match(
      strFromU8(parts["ppt/presentation.xml"]!),
      /<p:notesMasterIdLst><p:notesMasterId r:id="rId\d+"\/><\/p:notesMasterIdLst><p:sldIdLst>/u,
      author,
    );

    // Canvas pixel checks: shapes drawn accurately and empty lower canvas stays pure background.
    const png1 = preview.pages[0].png!;
    assert.deepEqual(pixelAt(png1, 10, 10), [0xf8, 0xfa, 0xfc, 255], author);
    assert.deepEqual(pixelAt(png1, 160, 100), [0x1e, 0x3a, 0x8a, 255], author);
    assert.deepEqual(pixelAt(png1, 450, 100), [0xd9, 0x77, 0x06, 255], author);
    assert.deepEqual(pixelAt(png1, 320, 260), [0xf8, 0xfa, 0xfc, 255], author);

    const png2 = preview.pages[1].png!;
    assert.deepEqual(pixelAt(png2, 10, 10), [0xee, 0xf2, 0xff, 255], author);
    assert.deepEqual(pixelAt(png2, 320, 180), [0x43, 0x38, 0xca, 255], author);
  },
};

export const notesPreviewIsolationScenario: CapabilityScenario = {
  id: "notes-preview-isolation",
  capability: "notes",
  kind: "rejection",
  description:
    "Dual-author scenario verifying Notes never blocks valid slide preview and preserves PPTX notes when another canvas object fails preview.",
  visualCheckpoints: [
    "Page 1 (notes-with-missing-glyph): preview omits PNG due to missing glyph in declared font while keeping Notes intact in PPTX.",
    "Page 2 (notes-with-valid-shape): clean white canvas with emerald rectangle (#059669) and zero notes on canvas.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX and verify both Slide 1 and Slide 2 retain their speaker notes.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 480, height: 270 },
      createElement(
        Slide,
        { id: "notes-with-missing-glyph" },
        createElement(
          Text,
          {
            id: "missing-glyph-text",
            x: 40,
            y: 40,
            width: 160,
            height: 80,
            style: { fontFamily: jostFont, fontSize: 18 },
          },
          "汉",
        ),
        createElement(
          Notes,
          { id: "preserved-notes" },
          createElement(P, null, "Notes survive preview rejection"),
        ),
      ),
      createElement(
        Slide,
        { id: "notes-with-valid-shape" },
        createElement(Shape, {
          id: "emerald-card",
          x: 120,
          y: 60,
          width: 240,
          height: 150,
          fill: "#059669",
        }),
        createElement(
          Notes,
          { id: "valid-notes" },
          createElement(P, null, "Valid page notes"),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 480, height: 270 }).append(
      slide({ id: "notes-with-missing-glyph" }).append(
        text("汉", {
          id: "missing-glyph-text",
          x: 40,
          y: 40,
          width: 160,
          height: 80,
          style: { fontFamily: jostFont, fontSize: 18 },
        }),
        notes("Notes survive preview rejection", { id: "preserved-notes" }),
      ),
      slide({ id: "notes-with-valid-shape" }).append(
        shape({
          id: "emerald-card",
          x: 120,
          y: 60,
          width: 240,
          height: 150,
          fill: "#059669",
        }),
        notes("Valid page notes", { id: "valid-notes" }),
      ),
    );
  },
  assert({ author, preview, validation }) {
    assert.equal(validation.slides.length, 2, author);
    assert.equal(
      validation.slides[0]?.notes?.text.text,
      "Notes survive preview rejection",
      author,
    );
    assert.equal(
      validation.slides[1]?.notes?.text.text,
      "Valid page notes",
      author,
    );
    assert.equal(validation.slides[1]?.objects.length, 1, author);

    assert.equal(preview.pages[0]?.png, undefined, author);
    assert.equal(preview.pages[0]?.diagnostics[0]?.severity, "error", author);
    assert.match(
      preview.pages[0]?.diagnostics[0]?.path ?? "",
      /text\[missing-glyph-text\]/u,
      author,
    );

    assert.ok(preview.pages[1]?.png !== undefined, author);
    assert.deepEqual(preview.pages[1]?.diagnostics, [], author);
    assert.deepEqual(
      pixelAt(preview.pages[1].png, 240, 135),
      [0x05, 0x96, 0x69, 255],
      author,
    );
  },
};
