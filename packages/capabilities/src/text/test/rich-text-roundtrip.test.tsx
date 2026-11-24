import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import type {
  ParagraphNode,
  ParagraphStyle,
  SpanNode,
  TextStyle,
} from "@pptxdk/capabilities/document";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode, RichTextContainer } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

function nodeStyle(node: CommonNode): ParagraphStyle & TextStyle {
  return node.props["style"] ?? {};
}

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);

describe("Rich text, font size units, and theme font roundtrip", () => {
  it("roundtrips rich paragraphs, two distinct Spans, lists, levels, line breaks, shadows, decorations, and SlideNumber, and keeps Span 2 intact when Span 1 is edited", async () => {
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "rich" },
          createElement(
            Text,
            {
              id: "card",
              name: "RichCard",
              x: 24,
              y: 24,
              width: 420,
              height: 220,
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
                    fontFamily: "major",
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
        ),
      ),
    );

    const initialEmit = emit(resolveDocument(doc));
    const pptxCopy = new Uint8Array(initialEmit.pptx);
    const imported = importDocument(pptxCopy);
    pptxCopy.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);

    const textNode = imported.document.query
      .get("rich")
      .node.children.find(
        (child): child is RichTextContainer =>
          child instanceof RichTextContainer && child.kind === "text",
      );
    assert.ok(textNode !== undefined);
    assert.equal(textNode.children.length, 3);

    const p1 = textNode.children[0] as ParagraphNode;
    assert.ok(p1 instanceof RichTextContainer && p1.kind === "paragraph");
    assert.equal(p1.children.length, 2);

    const span1 = p1.children[0] as SpanNode;
    const span2 = p1.children[1] as SpanNode;
    assert.ok(span1 instanceof RichTextContainer && span1.kind === "span");
    assert.ok(span2 instanceof RichTextContainer && span2.kind === "span");
    assert.equal(nodeStyle(span2).fontFamily, "major");
    assert.equal(nodeStyle(span2).fontSize, "14pt");
    assert.equal(nodeStyle(span2).textDecoration, "line-through");
    assert.equal(nodeStyle(span2).textTransform, "uppercase");

    // First re-emit without edits: validate rich text fidelity independently.
    const roundtripEmit = emit(resolveDocument(imported.document));
    const validation = await validatePptx(roundtripEmit.pptx);
    assert.deepEqual(validation.errors, []);

    const cardObj = validation.slides[0]?.objects[0];
    assert.ok(cardObj?.text !== undefined);
    assert.equal(cardObj.text.paragraphs.length, 3);
    assert.equal(cardObj.text.paragraphs[0]?.alignment, "ctr");
    assert.equal(cardObj.text.paragraphs[0]?.lineSpacingPoints, 24);
    assert.equal(cardObj.text.paragraphs[0]?.spacingAfterPoints, 7.5);
    assert.equal(cardObj.text.paragraphs[0]?.runs.length, 2);
    assert.equal(cardObj.text.paragraphs[0]?.runs[0]?.bold, true);
    assert.equal(cardObj.text.paragraphs[0]?.runs[0]?.color, "#1D4ED8");
    assert.equal(cardObj.text.paragraphs[0]?.runs[0]?.fontSizePt, 18);
    assert.equal(cardObj.text.paragraphs[0]?.runs[1]?.italic, true);
    assert.equal(cardObj.text.paragraphs[0]?.runs[1]?.color, "#DC2626");
    assert.equal(cardObj.text.paragraphs[0]?.runs[1]?.fontSizePt, 14);
    assert.equal(
      cardObj.text.paragraphs[0]?.runs[1]?.fontFamily,
      "Aptos Display",
    );

    assert.deepEqual(cardObj.text.paragraphs[1]?.list, {
      type: "bullet",
      character: "+",
    });
    assert.equal(cardObj.text.paragraphs[1]?.text, "Bullet\nline");
    assert.deepEqual(cardObj.text.paragraphs[2]?.list, {
      type: "number",
      style: "arabicPeriod",
      startAt: 5,
    });
    assert.equal(cardObj.text.paragraphs[2]?.runs[1]?.field, "slidenum");

    const slideXml = strFromU8(
      unzipSync(roundtripEmit.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /lvl="1"/u);
    assert.match(slideXml, /lvl="2"/u);
    assert.match(slideXml, /u="sng"/u);
    assert.match(slideXml, /strike="sngStrike"/u);
    assert.match(slideXml, /cap="all"/u);
    assert.match(slideXml, /spc="150"/u);
    assert.match(slideXml, /<a:outerShdw /u);
    assert.match(slideXml, /typeface="\+mj-lt"/u);

    // Edit Span 1 only and verify Span 2 style & text stay intact.
    span1.content("Edited ");
    span1.style({ color: "#059669" });
    const editedEmit = emit(resolveDocument(imported.document));
    const editedValidation = await validatePptx(editedEmit.pptx);
    assert.deepEqual(editedValidation.errors, []);

    const editedRuns =
      editedValidation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs;
    assert.equal(editedRuns?.[0]?.text, "Edited ");
    assert.equal(editedRuns?.[0]?.color, "#059669");
    assert.equal(editedRuns?.[0]?.bold, true);
    assert.equal(editedRuns?.[1]?.text, "second");
    assert.equal(editedRuns?.[1]?.italic, true);
    assert.equal(editedRuns?.[1]?.color, "#DC2626");
    assert.equal(editedRuns?.[1]?.fontSizePt, 14);
    assert.equal(editedRuns?.[1]?.fontFamily, "Aptos Display");
  });

  it("distinguishes 32pt vs 32px font sizes and minor/major/explicit/omitted font families across roundtrip", async () => {
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "units" },
          createElement(
            Text,
            {
              id: "sizes",
              x: 24,
              y: 24,
              width: 360,
              height: 180,
            },
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontSize: "32pt", fontFamily: "minor" } },
                "32pt minor",
              ),
              createElement(
                Span,
                { style: { fontSize: "32px", fontFamily: "major" } },
                "32px major",
              ),
              createElement(
                Span,
                { style: { fontSize: 16, fontFamily: "CustomSans" } },
                "16px custom",
              ),
            ),
          ),
        ),
      ),
    );

    const initial = emit(resolveDocument(doc));
    const imported = importDocument(new Uint8Array(initial.pptx));
    assert.deepEqual(imported.report.diagnostics, []);

    const pNode = imported.document.query
      .get("units")
      .node.children.find(
        (child): child is RichTextContainer =>
          child instanceof RichTextContainer && child.kind === "text",
      )?.children[0] as ParagraphNode;
    const sPt = pNode.children[0] as SpanNode;
    const sPx = pNode.children[1] as SpanNode;
    const sCustom = pNode.children[2] as SpanNode;

    assert.equal(nodeStyle(sPt).fontSize, "32pt");
    assert.equal(nodeStyle(sPt).fontFamily, "minor");
    assert.equal(nodeStyle(sPx).fontSize, "24pt");
    assert.equal(nodeStyle(sPx).fontFamily, "major");
    assert.equal(nodeStyle(sCustom).fontSize, "12pt");
    assert.equal(nodeStyle(sCustom).fontFamily, "CustomSans");

    const reemitted = emit(resolveDocument(imported.document));
    const validation = await validatePptx(reemitted.pptx);
    assert.deepEqual(validation.errors, []);
    const runs = validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs;
    assert.equal(runs?.[0]?.fontSizePt, 32);
    assert.equal(runs?.[1]?.fontSizePt, 24);
    assert.equal(runs?.[2]?.fontSizePt, 12);
    assert.equal(runs?.[2]?.fontFamily, "CustomSans");

    const slideXml = strFromU8(
      unzipSync(reemitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(slideXml, /sz="3200"/u);
    assert.match(slideXml, /sz="2400"/u);
    assert.match(slideXml, /typeface="\+mn-lt"/u);
    assert.match(slideXml, /typeface="\+mj-lt"/u);
  });

  it("handles external <a:br/>, <a:lnSpc><a:spcPct/></a:lnSpc>, dirty/smtClean, omitted font slots, and unsupported baseline attribute", async () => {
    const base = emit(
      resolveDocument(
        await expandDocument(
          createElement(
            Deck,
            { width: 640, height: 360 },
            createElement(
              Slide,
              { id: "ext" },
              createElement(
                Text,
                { id: "t1", x: 20, y: 20, width: 240, height: 80 },
                "Placeholder",
              ),
            ),
          ),
        ),
      ),
    );
    const entries = unzipSync(base.pptx);
    const slideXml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    const customTxBody = `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="en-US" sz="2000" dirty="0" smtClean="0"/><a:t>Line A</a:t></a:r><a:br/><a:r><a:rPr sz="2000" baseline="30000"/><a:t>Sup</a:t></a:r></a:p></p:txBody>`;
    entries["ppt/slides/slide1.xml"] = strToU8(
      slideXml.replace(/<p:txBody>[\s\S]*<\/p:txBody>/u, customTxBody),
    );

    const imported = importDocument(zipSync(entries));
    const baselineDiag = imported.report.diagnostics.find(
      (d) =>
        d.code === "unconsumed-xml" &&
        d.kind === "attribute" &&
        d.localName === "baseline",
    );
    assert.ok(baselineDiag !== undefined);
    // Only @baseline should be unconsumed; spcPct, br, dirty, smtClean are cleanly handled.
    assert.equal(imported.report.diagnostics.length, 1);

    const textNode = imported.document.query
      .get("ext")
      .node.children.find(
        (child): child is RichTextContainer =>
          child instanceof RichTextContainer && child.kind === "text",
      );
    assert.ok(textNode !== undefined);
    const pNode = textNode.children[0] as ParagraphNode;
    assert.equal(nodeStyle(pNode).lineHeight, "30pt");
    // Omitted <a:latin> on external runs remains omitted rather than forced.
    const firstSpan = pNode.children[0] as SpanNode;
    assert.equal(nodeStyle(firstSpan).fontFamily, undefined);

    const rebuilt = emit(resolveDocument(imported.document));
    const validation = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(validation.slides[0]?.objects[0]?.text?.text, "Line A\nSup");
  });
});
