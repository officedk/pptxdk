import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { createElement } from "react";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  deck,
  defineSlideLayout,
  placeholder,
  slide,
  slideLayout,
  slideMaster,
} from "../../document.ts";
import { Placeholder } from "../placeholder-elements.ts";

describe("placeholder capability codegen", () => {
  it("generates vanilla and react source for bound placeholders and omits redundant nativePlaceholder", async () => {
    const contentLayout = defineSlideLayout({
      id: "content-layout",
      name: "Title and Content",
      type: "obj",
      placeholders: {
        content: {
          type: "obj",
          explicitType: false,
          idx: 1,
          x: 60,
          y: 120,
          width: 600,
          height: 300,
        },
      },
    });

    const vanillaDoc = deck({ width: 960, height: 540 }).append(
      slideMaster({ id: "m1" }).append(slideLayout(contentLayout)),
      slide({ id: "s1", layout: contentLayout }).append(
        placeholder({
          id: "ph-redundant",
          name: "Content Placeholder 1",
          placeholder: "content",
          nativePlaceholder: { type: "obj", explicitType: false, idx: 1 },
        }),
        placeholder({
          id: "ph-custom-prompt",
          placeholder: "content",
          nativePlaceholder: {
            type: "obj",
            explicitType: false,
            idx: 1,
            hasCustomPrompt: true,
          },
        }),
        placeholder({
          id: "ph-orphan",
          nativePlaceholder: { type: "obj", idx: 99 },
          x: 80,
          y: 140,
          width: 240,
          height: 120,
        }),
      ),
    );

    const reactDoc = await expandDocument(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          SlideMaster,
          { id: "m1" },
          createElement(SlideLayout, { layout: contentLayout }),
        ),
        createElement(
          Slide,
          { id: "s1", layout: contentLayout },
          createElement(Placeholder, {
            id: "ph-redundant",
            name: "Content Placeholder 1",
            placeholder: "content",
            nativePlaceholder: { type: "obj", explicitType: false, idx: 1 },
          }),
          createElement(Placeholder, {
            id: "ph-custom-prompt",
            placeholder: "content",
            nativePlaceholder: {
              type: "obj",
              explicitType: false,
              idx: 1,
              hasCustomPrompt: true,
            },
          }),
          createElement(Placeholder, {
            id: "ph-orphan",
            nativePlaceholder: { type: "obj", idx: 99 },
            x: 80,
            y: 140,
            width: 240,
            height: 120,
          }),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    for (const [label, doc] of [
      ["vanilla-input", vanillaDoc],
      ["react-input", reactDoc],
    ] as const) {
      const vanillaOut = generate(doc, { target: "vanilla" });
      const vanillaSource = vanillaOut.files[vanillaOut.entry] as string;
      assert.match(
        vanillaSource,
        /placeholder\(\s*\{\s*id:\s*"ph-redundant",\s*name:\s*"Content Placeholder 1",\s*placeholder:\s*"content",?\s*\},?\s*\)/u,
        `${label} vanilla redundant nativePlaceholder should be omitted`,
      );
      assert.match(
        vanillaSource,
        /id:\s*"ph-custom-prompt"[\s\S]*?nativePlaceholder:\s*\{/u,
        `${label} vanilla differing nativePlaceholder should be preserved`,
      );
      assert.match(
        vanillaSource,
        /id:\s*"ph-orphan"[\s\S]*?nativePlaceholder:\s*\{\s*type:\s*"obj",\s*idx:\s*99,?\s*\}/u,
        `${label} vanilla orphan nativePlaceholder should be preserved`,
      );

      const reactOut = generate(doc, { target: "react" });
      const reactSource = reactOut.files[reactOut.entry] as string;
      assert.match(
        reactSource,
        /<Placeholder\s+id="ph-redundant"\s+name="Content Placeholder 1"\s+placeholder="content"\s*\/>/u,
        `${label} react redundant nativePlaceholder should be omitted`,
      );
      assert.match(
        reactSource,
        /id="ph-custom-prompt"[\s\S]*?nativePlaceholder=\{\{/u,
        `${label} react differing nativePlaceholder should be preserved`,
      );
      assert.match(
        reactSource,
        /id="ph-orphan"[\s\S]*?nativePlaceholder=\{\{\s*type:\s*"obj",\s*idx:\s*99,?\s*\}\}/u,
        `${label} react orphan nativePlaceholder should be preserved`,
      );
    }
  });

  it("throws CodegenError with node path when placeholder has unsupported children", () => {
    const badPlaceholder = new CommonContainer("placeholder", "placeholder", {
      id: "bad-ph",
      placeholder: "content",
    });
    badPlaceholder.append(new CommonNode("text", "text", { id: "inner" }));
    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1" }).append(badPlaceholder),
    );

    const generate = createCodegen(documentCapabilities);
    assert.throws(
      () => generate(doc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/placeholder[bad-ph]",
    );
  });
});
