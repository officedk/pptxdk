import assert from "node:assert/strict";

import {
  deck,
  notes,
  paragraph,
  section,
  shape,
  slide,
  slideNumber,
  span,
} from "@pptxdk/capabilities/document";
import { Notes } from "@pptxdk/capabilities/notes-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { P, SlideNumber, Span } from "@pptxdk/capabilities/text-elements";
import { Deck, Section, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../../../test/mechanism-verification.ts";
import type {
  presentationSectionsMechanism,
  speakerNotesMechanism,
} from "../notes.ts";

export const notesAndSectionsRoundtripScenario: CapabilityScenario = {
  id: "mechanism-notes-and-sections-roundtrip",
  capability: "notes",
  description:
    "Dual-author speaker notes, slide order, hidden slide state, and p14:sectionLst grouping roundtrip.",
  visualCheckpoints: [
    "Page 1, 2 (hidden), and 3 draw clean canvas shapes with zero speaker notes text on slide canvas.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain exact slide order, p14:sectionLst membership, show='false' hidden state, and formatted speaker notes attached to owning slides.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Section,
        { name: "Opening" },
        createElement(
          Slide,
          { id: "intro-notes" },
          createElement(Shape, {
            id: "navy-card",
            style: {
              position: "absolute",
              left: 48,
              top: 48,
              width: 240,
              height: 110,
              background: "#1E3A8A",
            },
          }),
          createElement(
            Notes,
            { id: "intro-speech" },
            createElement(P, null, "Speaker cue: introduce the deck."),
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontWeight: "bold", color: "#D97706" } },
                "Key takeaway: ",
              ),
              "keep notes off the visible canvas.",
            ),
            createElement(
              P,
              { style: { list: { type: "bullet" }, align: "center" } },
              "Slide index ",
              createElement(SlideNumber, null),
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "hidden-notes", hidden: true },
          createElement(Shape, {
            id: "hidden-card",
            style: {
              position: "absolute",
              left: 64,
              top: 64,
              width: 180,
              height: 90,
              background: "#991B1B",
            },
          }),
          createElement(
            Notes,
            { id: "hidden-speech" },
            createElement(P, null, "Hidden slide cue"),
          ),
        ),
      ),
      createElement(
        Section,
        { name: "Appendix" },
        createElement(
          Slide,
          { id: "no-notes" },
          createElement(Shape, {
            id: "indigo-oval",
            geometry: "ellipse",
            style: {
              position: "absolute",
              left: 200,
              top: 90,
              width: 240,
              height: 180,
              background: "#4338CA",
            },
          }),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      section({ name: "Opening" }).append(
        slide({ id: "intro-notes" }).append(
          shape({
            id: "navy-card",
            style: {
              position: "absolute",
              left: 48,
              top: 48,
              width: 240,
              height: 110,
              background: "#1E3A8A",
            },
          }),
          notes(
            [
              paragraph("Speaker cue: introduce the deck."),
              paragraph([
                span("Key takeaway: ", {
                  style: { fontWeight: "bold", color: "#D97706" },
                }),
                "keep notes off the visible canvas.",
              ]),
              paragraph(["Slide index ", slideNumber()], {
                style: { list: { type: "bullet" }, align: "center" },
              }),
            ],
            { id: "intro-speech" },
          ),
        ),
        slide({ id: "hidden-notes", hidden: true }).append(
          shape({
            id: "hidden-card",
            style: {
              position: "absolute",
              left: 64,
              top: 64,
              width: 180,
              height: 90,
              background: "#991B1B",
            },
          }),
          notes("Hidden slide cue", { id: "hidden-speech" }),
        ),
      ),
      section({ name: "Appendix" }).append(
        slide({ id: "no-notes" }).append(
          shape({
            id: "indigo-oval",
            geometry: "ellipse",
            style: {
              position: "absolute",
              left: 200,
              top: 90,
              width: 240,
              height: 180,
              background: "#4338CA",
            },
          }),
        ),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 3);
    assert.equal(validation.presentation.sections?.length, 2);
  },
};

export const notesMechanismVerifications = {
  "speaker-notes": {
    scenario: notesAndSectionsRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 3);

      const expectedNotesText =
        "Speaker cue: introduce the deck.\nKey takeaway: keep notes off the visible canvas.\nSlide index 1";
      const notes1 = validation.slides[0]?.notes;
      assert.ok(notes1, `${stage} slide 1 notes`);
      assert.equal(notes1.text.text, expectedNotesText);
      assert.equal(notes1.text.paragraphs[0]?.runs[0]?.fontSizePt, 12);
      assert.equal(notes1.text.paragraphs[1]?.runs[0]?.bold, true);
      assert.equal(notes1.text.paragraphs[1]?.runs[0]?.color, "#D97706");
      assert.equal(notes1.text.paragraphs[2]?.alignment, "ctr");
      assert.deepEqual(notes1.text.paragraphs[2]?.list, {
        type: "bullet",
        character: "•",
      });
      assert.equal(notes1.text.paragraphs[2]?.runs[1]?.field, "slidenum");
      assert.equal(notes1.text.paragraphs[2]?.runs[1]?.text, "1");

      const notes2 = validation.slides[1]?.notes;
      assert.ok(notes2, `${stage} slide 2 notes`);
      assert.equal(notes2.text.text, "Hidden slide cue");
      assert.equal(validation.slides[2]?.notes, undefined);

      const parts = unzipSync(pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      assert.doesNotMatch(slide1Xml, /Speaker cue|Key takeaway/u);
    },
  } satisfies MechanismVerificationEntry<typeof speakerNotesMechanism>,
  "presentation-sections": {
    scenario: notesAndSectionsRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.deepEqual(
        validation.slides.map((s) => ({
          id: s.id,
          hidden: s.hidden ?? false,
          section: s.section,
        })),
        [
          { id: "intro-notes", hidden: false, section: "Opening" },
          { id: "hidden-notes", hidden: true, section: "Opening" },
          { id: "no-notes", hidden: false, section: "Appendix" },
        ],
      );
      assert.deepEqual(
        validation.presentation.sections?.map((sec) => ({
          name: sec.name,
          slideIds: sec.slideIds,
        })),
        [
          { name: "Opening", slideIds: ["intro-notes", "hidden-notes"] },
          { name: "Appendix", slideIds: ["no-notes"] },
        ],
      );
    },
  } satisfies MechanismVerificationEntry<typeof presentationSectionsMechanism>,
} satisfies ExhaustiveMechanismVerificationMap<
  typeof speakerNotesMechanism | typeof presentationSectionsMechanism
>;
