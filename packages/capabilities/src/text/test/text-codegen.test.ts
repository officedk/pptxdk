import assert from "node:assert/strict";

import { createCodegen } from "@pptxdk/codegen";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import { deck, paragraph, section, slide, span, text } from "../../document.ts";

describe("text capability codegen (basic)", () => {
  it("generates vanilla and react source for text, paragraph, and span", () => {
    const doc = deck({ width: 960, height: 540 }).append(
      section({ name: "Intro" }).append(
        slide({ id: "s1" }).append(
          text("Simple heading", {
            id: "heading",
            x: 48,
            y: 40,
            width: 400,
            height: 48,
            style: { fontSize: 24, fontWeight: "bold", color: "#111827" },
          }),
          text(
            [
              paragraph(
                [
                  span("Bold ", { style: { fontWeight: "bold" } }),
                  span("tail"),
                ],
                { style: { align: "left" } },
              ),
              paragraph(),
            ],
            {
              id: "body",
              x: 48,
              y: 110,
              width: 400,
              height: 120,
            },
          ),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /text\(\s*"Simple heading",\s*\{/u);
    assert.match(vanillaSource, /paragraph\(/u);
    assert.match(vanillaSource, /style: \{ align: "left" \}/u);
    assert.match(
      vanillaSource,
      /span\("Bold ", \{ style: \{ fontWeight: "bold" \} \}\)/u,
    );
    assert.match(vanillaSource, /span\("tail"\)/u);
    assert.match(vanillaSource, /paragraph\(\)/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /<Text[\s\S]*?>\s*Simple heading\s*<\/Text>/u);
    assert.match(
      reactSource,
      /<Span style=\{\{ fontWeight: "bold" \}\}>\{"Bold "\}<\/Span>/u,
    );
    assert.match(reactSource, /<Span>tail<\/Span>/u);
    assert.match(reactSource, /<P \/>/u);
  });
});
