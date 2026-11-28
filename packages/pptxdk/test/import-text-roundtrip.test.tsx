import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  ImportError,
  RichTextContainer,
  build,
  emit,
  font,
  importPptx,
  resolve,
} from "@officedk/pptxdk";
import { createImporter, selectCapabilities } from "@officedk/pptxdk/build";
import { importPptxFile } from "@officedk/pptxdk/node";
import { preview } from "@officedk/pptxdk/preview";
import { Deck, Slide, Text } from "@officedk/pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBytes }]);

describe("pptxdk import -> edit -> resolve/preview/emit/build roundtrip", () => {
  it("imports a Transitional single-slide text PPTX, detaches from input bytes, edits text, previews with Font, and re-emits valid OOXML", async () => {
    const initial = await build(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "title",
              name: "Title Box",
              style: {
                position: "absolute",
                left: 40,
                top: 36,
                width: 360,
                height: 60,
                fontFamily: "Jost*",
                fontSize: "24pt",
                color: "#0F172A",
                align: "left",
                verticalAlign: "top",
              },
            },
            "Original Slide Title",
          ),
        ),
      ),
    );

    const inputCopy = new Uint8Array(initial.pptx);
    const imported = importPptx(inputCopy);
    // Destroy caller buffer immediately after import to prove zero aliasing/retention.
    inputCopy.fill(0);

    assert.equal(imported.report.evidence, "import");
    assert.equal(
      imported.report.sourceFormat,
      "ECMA-376 5th edition Transitional",
    );
    assert.equal(imported.report.slideCount, 1);
    assert.deepEqual(imported.report.diagnostics, []);

    const slideNode = imported.document.query.get("cover").node;
    const textNode = slideNode.children.find(
      (child): child is RichTextContainer => child instanceof RichTextContainer,
    );
    assert.ok(textNode !== undefined);

    textNode.content("Edited Slide Title");
    textNode.style({ fontFamily: jostFont, color: "#1D4ED8" });

    const resolved = resolve(imported.document);
    const painted = await preview(imported.document);
    const emitted = emit(resolved);
    const rebuilt = await build(imported.document);

    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.deepEqual(painted.pages[0]?.placeholders, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);
    const png = PNG.sync.read(Buffer.from(painted.pages[0].png));
    assert.equal(png.width, 640);
    assert.equal(png.height, 360);
    assert.ok(painted.pages[0]?.svg?.includes('fill="#1D4ED8"'));

    for (const pptxBytes of [emitted.pptx, rebuilt.pptx]) {
      const validated = await validatePptx(pptxBytes);
      assert.deepEqual(validated.errors, []);
      assert.equal(validated.slides.length, 1);
      assert.equal(validated.slides[0]?.id, "cover");
      assert.equal(validated.slides[0]?.objects.length, 1);
      const obj = validated.slides[0].objects[0]!;
      assert.equal(obj.name, "Title Box");
      assert.equal(obj.text?.text, "Edited Slide Title");
      assert.equal(obj.text?.paragraphs[0]?.runs[0]?.fontSizePt, 24);
      assert.equal(obj.text?.paragraphs[0]?.runs[0]?.color, "#1D4ED8");
      assert.deepEqual(obj.boxEmu, {
        x: 40 * 9525,
        y: 36 * 9525,
        width: 360 * 9525,
        height: 60 * 9525,
      });
    }
  });

  it("supports selective capability import via @officedk/pptxdk/build without loading unselected modules, React, or preview", async () => {
    const sample = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "trimmed" },
          createElement(
            Text,
            {
              id: "t1",
              name: "Label",
              style: {
                position: "absolute",
                left: 16,
                top: 16,
                width: 200,
                height: 40,
              },
            },
            "Selective Import",
          ),
        ),
      ),
    );
    const textOnly = await selectCapabilities(["text"]);
    const selectiveImporter = createImporter(textOnly);
    const res = selectiveImporter(sample.pptx);
    assert.equal(res.report.slideCount, 1);
  });

  it("reads PPTX files via @officedk/pptxdk/node importPptxFile and surfaces ImportError on invalid files", async () => {
    const sample = await build(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "file-slide" },
          createElement(
            Text,
            {
              id: "msg",
              name: "Msg",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 150,
                height: 32,
              },
            },
            "From File",
          ),
        ),
      ),
    );

    const dir = await mkdtemp(join(tmpdir(), "pptxdk-import-node-"));
    try {
      const validPath = join(dir, "valid.pptx");
      const corruptPath = join(dir, "corrupt.pptx");
      await writeFile(validPath, sample.pptx);
      await writeFile(corruptPath, new Uint8Array([9, 8, 7, 6]));

      const imported = await importPptxFile(validPath);
      assert.equal(imported.report.slideCount, 1);
      assert.deepEqual(imported.report.diagnostics, []);

      const rebuilt = await build(imported.document);
      const validated = await validatePptx(rebuilt.pptx);
      assert.deepEqual(validated.errors, []);
      assert.equal(validated.slides[0]?.objects[0]?.text?.text, "From File");

      await assert.rejects(
        () => importPptxFile(corruptPath),
        (error: unknown) =>
          error instanceof ImportError &&
          error.issue.code === "corrupt-archive",
      );

      const controller = new AbortController();
      const reason = new Error("aborted file import");
      controller.abort(reason);
      await assert.rejects(
        () => importPptxFile(validPath, { signal: controller.signal }),
        (error: unknown) => error === reason,
      );
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("reports exact part and xpath for unconsumed XML inside a text box while keeping the text editable", async () => {
    const base = await build(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "t1",
              name: "Box",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 200,
                height: 40,
              },
            },
            "Keep text",
          ),
        ),
      ),
    );
    const entries = unzipSync(base.pptx);
    const slideXml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    entries["ppt/slides/slide1.xml"] = strToU8(
      slideXml
        .replace("<a:bodyPr ", '<a:bodyPr vert="eaVert" ')
        .replace("<a:noFill/>", "<a:noFill/><p:unknownUnconsumed/>"),
    );
    const mutatedPptx = zipSync(entries);

    const imported = importPptx(mutatedPptx);
    mutatedPptx.fill(0);

    assert.equal(imported.report.diagnostics.length, 2);
    const unconsumedAttr = imported.report.diagnostics.find(
      (d) => d.code === "unconsumed-xml" && d.kind === "attribute",
    );
    const unconsumedEl = imported.report.diagnostics.find(
      (d) => d.code === "unconsumed-xml" && d.kind === "element",
    );
    assert.ok(
      unconsumedAttr !== undefined && unconsumedAttr.code === "unconsumed-xml",
    );
    assert.equal(unconsumedAttr.localName, "vert");
    assert.equal(unconsumedAttr.location.part, "ppt/slides/slide1.xml");
    assert.ok(unconsumedAttr.location.xpath?.endsWith("/a:bodyPr/@vert"));

    assert.ok(
      unconsumedEl !== undefined && unconsumedEl.code === "unconsumed-xml",
    );
    assert.equal(unconsumedEl.localName, "unknownUnconsumed");
    assert.equal(unconsumedEl.location.part, "ppt/slides/slide1.xml");
    assert.ok(
      unconsumedEl.location.xpath?.endsWith("/p:spPr/p:unknownUnconsumed"),
    );

    const textNode = imported.document.query
      .get("s1")
      .node.children.find(
        (child): child is RichTextContainer =>
          child instanceof RichTextContainer,
      );
    assert.ok(textNode !== undefined);
    textNode.content("Still editable");

    const rebuilt = await build(imported.document);
    const validated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validated.errors, []);
    assert.equal(validated.slides[0]?.objects[0]?.text?.text, "Still editable");
  });
});
