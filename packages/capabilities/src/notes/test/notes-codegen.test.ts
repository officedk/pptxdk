import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  deck,
  notes,
  paragraph,
  section,
  slide,
  slideNumber,
  span,
  text,
} from "../../document.ts";
import { notesCapability } from "../notes.ts";

describe("notes capability codegen", () => {
  it("generates vanilla and react source for simple, rich-text, and empty speaker notes", () => {
    const doc = deck({ width: 960, height: 540 }).append(
      section({ name: "Main" }).append(
        slide({ id: "s1" }).append(
          text("Canvas title", {
            id: "canvas-title",
            x: 48,
            y: 40,
            width: 400,
            height: 48,
          }),
          notes(
            [
              paragraph("Speaker cue: introduce the architecture."),
              paragraph([
                span("Key takeaway: ", {
                  style: { fontWeight: "bold", color: "#D97706" },
                }),
                "keep notes off the visible canvas.",
              ]),
              paragraph(),
              paragraph(["Slide index ", slideNumber()], {
                style: { list: { type: "bullet" }, align: "center" },
              }),
            ],
            { id: "intro-speech", style: { fontSize: "14pt", color: "dk1" } },
          ),
        ),
        slide({ id: "s2" }).append(
          notes("Single line cue", { id: "s2-speech" }),
        ),
        slide({ id: "s3" }).append(
          notes([paragraph("No props line 1"), paragraph("No props line 2")]),
        ),
        slide({ id: "s4" }).append(notes([], { id: "empty-speech" })),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /notes\(/u);
    assert.match(vanillaSource, /id: "intro-speech"/u);
    assert.match(vanillaSource, /fontSize: "14pt"/u);
    assert.match(vanillaSource, /slideNumber\(\)/u);
    assert.match(
      vanillaSource,
      /notes\("Single line cue", \{ id: "s2-speech" \}\)/u,
    );
    assert.match(vanillaSource, /notes\(\)\.append\(/u);
    assert.match(vanillaSource, /notes\(\[\], \{ id: "empty-speech" \}\)/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /<Notes id="intro-speech"/u);
    assert.match(reactSource, /\{"Key takeaway: "\}/u);
    assert.match(reactSource, /<P \/>/u);
    assert.match(reactSource, /\{"Slide index "\}/u);
    assert.match(reactSource, /<SlideNumber \/>/u);
    assert.match(
      reactSource,
      /<Notes id="s2-speech">Single line cue<\/Notes>/u,
    );
    assert.match(reactSource, /<Notes id="empty-speech" \/>/u);
  });

  it("rejects illegal notes children, unknown props, and trimmed notesCapability with exact paths", () => {
    const generate = createCodegen(documentCapabilities);

    const illegalChildDoc = deck().append(
      slide({ id: "s1" }).append(
        notes(text("illegal", { id: "bad-text" }), { id: "n1" }),
      ),
    );
    assert.throws(
      () => generate(illegalChildDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/notes[n1]/text[bad-text]",
    );

    const unknownNotesFactory = notes as unknown as (
      content: string,
      props: Record<string, unknown>,
    ) => ReturnType<typeof notes>;
    const unknownPropDoc = deck().append(
      slide({ id: "s1" }).append(
        unknownNotesFactory("cue", { id: "n1", layout: { x: 1 } }),
      ),
    );
    assert.throws(
      () => generate(unknownPropDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/notes[n1]@layout",
    );

    const trimmedGenerate = createCodegen(
      documentCapabilities.filter((mod) => mod !== notesCapability),
    );
    const trimmedDoc = deck().append(
      slide({ id: "s1" }).append(notes("trimmed", { id: "n1" })),
    );
    assert.throws(
      () => trimmedGenerate(trimmedDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-declaration" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/notes[n1]",
    );
  });
});
