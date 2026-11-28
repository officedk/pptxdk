import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  paragraph,
  section,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { CommonContainer } from "@pptxdk/core/document";
import { describe, it } from "vitest";

describe("text capability codegen", () => {
  it("generates vanilla and react source for text, paragraph, span, and slide-number including lists, links, empty paragraphs, and whitespace", () => {
    const doc = deck({ width: 960, height: 540 }).append(
      section({ name: "Intro" }).append(
        slide({ id: "s1" }).append(
          text("Simple heading", {
            id: "heading",
            link: { url: "https://example.com/docs" },
            style: {
              position: "absolute",
              left: 48,
              top: 40,
              width: 400,
              height: 48,
              fontSize: 24,
              fontWeight: "bold",
              color: "#111827",
            },
          }),
          text(
            [
              paragraph(
                [
                  span("Bold ", { style: { fontWeight: "bold" } }),
                  span("tail\nline", {
                    style: { fontStyle: "italic", color: "#DC2626" },
                  }),
                  span(" end"),
                ],
                { style: { align: "left" } },
              ),
              paragraph(),
              paragraph(
                ["Page ", slideNumber({ style: { fontWeight: "bold" } })],
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
              ),
              paragraph("Custom bullet", {
                style: {
                  list: { type: "bullet", character: "+", level: 1 },
                },
              }),
            ],
            {
              id: "body",
              link: { id: "s2" },
              style: {
                position: "absolute",
                left: 48,
                top: 110,
                width: 400,
                height: 120,
              },
            },
          ),
        ),
        slide({ id: "s2" }).append(
          text(["Home ", slideNumber()], {
            id: "home-link",
            link: { action: "home" },
            style: {
              position: "absolute",
              left: 48,
              top: 40,
              width: 200,
              height: 40,
            },
          }),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /text\(\s*"Simple heading",\s*\{/u);
    assert.match(
      vanillaSource,
      /link: \{ url: "https:\/\/example\.com\/docs" \}/u,
    );
    assert.match(vanillaSource, /link: \{ id: "s2" \}/u);
    assert.match(vanillaSource, /link: \{ action: "home" \}/u);
    assert.match(vanillaSource, /paragraph\(/u);
    assert.match(vanillaSource, /style: \{ align: "left" \}/u);
    assert.match(
      vanillaSource,
      /span\("Bold ", \{ style: \{ fontWeight: "bold" \} \}\)/u,
    );
    assert.match(
      vanillaSource,
      /span\("tail\\nline",\s*\{\s*style: \{ fontStyle: "italic", color: "#DC2626" \}/u,
    );
    assert.match(vanillaSource, /paragraph\(\)/u);
    assert.match(
      vanillaSource,
      /slideNumber\(\{ style: \{ fontWeight: "bold" \} \}\)/u,
    );
    assert.match(vanillaSource, /slideNumber\(\)/u);
    assert.match(vanillaSource, /startAt: 5/u);
    assert.match(vanillaSource, /character: "\+"/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /<Text[\s\S]*?>\s*Simple heading\s*<\/Text>/u);
    assert.match(
      reactSource,
      /<Span style=\{\{ fontWeight: "bold" \}\}>\{"Bold "\}<\/Span>/u,
    );
    assert.match(reactSource, /\{"tail\\nline"\}/u);
    assert.match(reactSource, /<Span>\{" end"\}<\/Span>/u);
    assert.match(reactSource, /<P \/>/u);
    assert.match(
      reactSource,
      /<SlideNumber style=\{\{ fontWeight: "bold" \}\} \/>/u,
    );
    assert.match(reactSource, /<SlideNumber \/>/u);
    assert.match(reactSource, /link=\{\{ action: "home" \}\}/u);
  });

  it("rejects invalid rich-text children and unsupported props with CodegenError and exact paths", () => {
    const generate = createCodegen(documentCapabilities);

    const badSlideNum = new CommonContainer("slide-number", "slide-number", {
      id: "sn1",
    }).append(paragraph("illegal"));
    const badSlideNumDeck = deck().append(
      slide({ id: "s1" }).append(
        text([paragraph([badSlideNum])], {
          id: "t1",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: 40,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(badSlideNumDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/text[t1]/paragraph[1]/slide-number[sn1]",
    );

    const nestedParagraphDeck = deck().append(
      slide({ id: "s1" }).append(
        text(
          [paragraph([paragraph("nested", { id: "inner" })], { id: "outer" })],
          {
            id: "t1",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 100,
              height: 40,
            },
          },
        ),
      ),
    );
    assert.throws(
      () => generate(nestedParagraphDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/text[t1]/paragraph[outer]/paragraph[inner]",
    );

    const badSpanPropDeck = deck().append(
      slide({ id: "s1" }).append(
        text(
          [
            paragraph([
              span("x", { id: "sp1", align: "left" } as unknown as Record<
                string,
                unknown
              >),
            ]),
          ],
          {
            id: "t1",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 100,
              height: 40,
            },
          },
        ),
      ),
    );
    assert.throws(
      () => generate(badSpanPropDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/text[t1]/paragraph[1]/span[sp1]@align",
    );
  });
});
