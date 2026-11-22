import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import {
  CommonNode,
  DeckContainer,
  build,
  emit,
  importPptx,
  resolve,
} from "pptxdk";
import { paragraph } from "pptxdk/vanilla";
import { describe, it } from "vitest";

describe("SDK importPptx for prs-notes.pptx and baseline notesMaster", () => {
  it("imports prs-notes.pptx with zero false loss on notesSlide1.xml and baseline notesMaster placeholders, preserving editable speaker notes", async () => {
    const raw = new Uint8Array(
      readFileSync(new URL("./fixtures/prs-notes.pptx", import.meta.url)),
    );
    const imported = importPptx(raw);
    raw.fill(0);

    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.report.slideCount, 1);

    // Zero diagnostics on notesSlide1.xml (sldImg, body, sldNum) and zero false loss on baseline notesMaster placeholders.
    const notesSlideDiags = imported.report.diagnostics.filter((d) =>
      d.location.part.startsWith("ppt/notesSlides/"),
    );
    assert.deepEqual(notesSlideDiags, []);

    const notesMasterDiags = imported.report.diagnostics.filter((d) =>
      d.location.part.startsWith("ppt/notesMasters/"),
    );
    assert.equal(notesMasterDiags.length, 1);
    assert.equal(notesMasterDiags[0]?.code, "unsupported-mechanism");
    if (notesMasterDiags[0]?.code === "unsupported-mechanism") {
      assert.equal(notesMasterDiags[0].loss?.lost, "custom-print-layout");
      assert.equal(notesMasterDiags[0].location.objectName, "Picture 7");
    }

    const slideNode = imported.document.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide",
    );
    assert.ok(slideNode !== undefined);
    const notesNode = slideNode.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "notes",
    );
    assert.ok(notesNode !== undefined);

    // Edit speaker notes and rebuild through the formal SDK pipeline.
    (notesNode as unknown as { content(...items: unknown[]): void }).content(
      paragraph("Updated from prs-notes.pptx"),
    );

    const rebuilt = await build(imported.document);
    const validated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(
      validated.slides[0]?.notes?.text.text,
      "Updated from prs-notes.pptx",
    );

    const reimported = importPptx(rebuilt.pptx);
    assert.deepEqual(reimported.report.diagnostics, []);
  });

  it("imports layouts.pptx baseline notesMaster1.xml and notesSlide1.xml with zero notes diagnostics", () => {
    const raw = new Uint8Array(
      readFileSync(new URL("./fixtures/layouts.pptx", import.meta.url)),
    );
    const imported = importPptx(raw);
    raw.fill(0);

    const notesDiags = imported.report.diagnostics.filter(
      (d) =>
        d.location.part.startsWith("ppt/notesMasters/") ||
        d.location.part.startsWith("ppt/notesSlides/"),
    );
    assert.deepEqual(notesDiags, []);
    void emit;
    void resolve;
  });
});
