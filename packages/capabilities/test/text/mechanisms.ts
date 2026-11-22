import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  deck,
  font,
  paragraph,
  slide,
  slideNumber,
  span,
  text,
} from "@pptxdk/capabilities/document";
import { P, SlideNumber, Span, Text } from "@pptxdk/capabilities/text-elements";
import { defineTheme } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type {
  embeddedFontMechanism,
  textBoxMechanism,
  textFieldSlideNumMechanism,
} from "../../src/text/text.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";
import type {
  ExhaustiveMechanismVerificationMap,
  MechanismVerificationEntry,
} from "../mechanism-verification.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../assets/fonts/Jost-700-Bold.ttf", import.meta.url),
    ),
  },
]);

export const textBoxRoundtripScenario: CapabilityScenario = {
  id: "mechanism-text-box-roundtrip",
  capability: "text",
  description:
    "Minimal single-slide Transitional text-box scenario proving dual-author equivalence and import/edit/re-emit roundtrip.",
  visualCheckpoints: [
    "Single-slide white canvas with left-aligned headline text box and centered body text box.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain editable <p:sp txBox='1'> objects with matching geometry and text.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "text-roundtrip" },
        createElement(
          Text,
          {
            id: "headline",
            name: "Headline",
            x: 32,
            y: 28,
            width: 320,
            height: 48,
            style: {
              fontFamily: jostFont,
              fontSize: "24pt",
              color: "#111827",
              align: "left",
              verticalAlign: "top",
            },
          },
          "Imported Headline",
        ),
        createElement(
          Text,
          {
            id: "caption",
            name: "Caption",
            x: 32,
            y: 96,
            width: 320,
            height: 64,
            style: {
              fontFamily: jostFont,
              fontSize: "14pt",
              color: "#334155",
              align: "center",
              verticalAlign: "middle",
              padding: [6, 12, 6, 12],
            },
          },
          "Roundtrip Body Text",
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "text-roundtrip" }).append(
        text("Imported Headline", {
          id: "headline",
          name: "Headline",
          x: 32,
          y: 28,
          width: 320,
          height: 48,
          style: {
            fontFamily: jostFont,
            fontSize: "24pt",
            color: "#111827",
            align: "left",
            verticalAlign: "top",
          },
        }),
        text("Roundtrip Body Text", {
          id: "caption",
          name: "Caption",
          x: 32,
          y: 96,
          width: 320,
          height: 64,
          style: {
            fontFamily: jostFont,
            fontSize: "14pt",
            color: "#334155",
            align: "center",
            verticalAlign: "middle",
            padding: [6, 12, 6, 12],
          },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 1);
    assert.equal(validation.slides[0]?.objects.length, 2);
  },
};

export const textFieldSlideNumRoundtripScenario: CapabilityScenario = {
  id: "mechanism-text-field-slidenum-roundtrip",
  capability: "text",
  description:
    "Dual-author rich-text scenario with embedded font and native <a:fld type='slidenum'> field across 2 slides.",
  visualCheckpoints: [
    "Page 1 shows 'Slide 1'; Page 2 shows 'Slide 2' in Jost* with embedded font roundtrip.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX emit <a:fld type='slidenum'> with valid GUID id and embedded Jost* font.",
  ],
  emitOptions: { embedFonts: "full" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "p1" },
        createElement(
          Text,
          {
            id: "num-1",
            name: "Num1",
            x: 24,
            y: 24,
            width: 240,
            height: 48,
            style: { fontFamily: jostFont, fontSize: "18pt", color: "#0F172A" },
          },
          createElement(P, null, "Slide ", createElement(SlideNumber, null)),
        ),
      ),
      createElement(
        Slide,
        { id: "p2" },
        createElement(
          Text,
          {
            id: "num-2",
            name: "Num2",
            x: 24,
            y: 24,
            width: 240,
            height: 48,
            style: { fontFamily: jostFont, fontSize: "18pt", color: "#0F172A" },
          },
          createElement(P, null, "Slide ", createElement(SlideNumber, null)),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "p1" }).append(
        text([paragraph(["Slide ", slideNumber()])], {
          id: "num-1",
          name: "Num1",
          x: 24,
          y: 24,
          width: 240,
          height: 48,
          style: { fontFamily: jostFont, fontSize: "18pt", color: "#0F172A" },
        }),
      ),
      slide({ id: "p2" }).append(
        text([paragraph(["Slide ", slideNumber()])], {
          id: "num-2",
          name: "Num2",
          x: 24,
          y: 24,
          width: 240,
          height: 48,
          style: { fontFamily: jostFont, fontSize: "18pt", color: "#0F172A" },
        }),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 2);
  },
};

const embeddedTheme = defineTheme({
  fonts: { major: jostFont, minor: jostFont },
});

export const embeddedFontRoundtripScenario: CapabilityScenario = {
  id: "mechanism-embedded-font-roundtrip",
  capability: "text",
  description:
    "Dual-author embedded font subset roundtrip covering theme minor font and explicit bold Span.",
  visualCheckpoints: [
    "Single-slide white canvas with theme minor Jost* regular and explicit Jost* bold runs.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX embed ppt/fonts/*.fntdata via <p:embeddedFontLst> and preserve +mn-lt theme reference.",
  ],
  emitOptions: { embedFonts: "subset" },
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: embeddedTheme },
      createElement(
        Slide,
        { id: "fonts" },
        createElement(
          Text,
          {
            id: "rich-font",
            name: "RichFont",
            x: 32,
            y: 32,
            width: 360,
            height: 64,
            style: { fontFamily: "minor", fontSize: "20pt", color: "#1E293B" },
          },
          createElement(
            P,
            null,
            "Theme ",
            createElement(
              Span,
              { style: { fontFamily: jostFont, fontWeight: "bold" } },
              "Embedded",
            ),
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: embeddedTheme }).append(
      slide({ id: "fonts" }).append(
        text(
          [
            paragraph([
              "Theme ",
              span("Embedded", {
                style: { fontFamily: jostFont, fontWeight: "bold" },
              }),
            ]),
          ],
          {
            id: "rich-font",
            name: "RichFont",
            x: 32,
            y: 32,
            width: 360,
            height: 64,
            style: { fontFamily: "minor", fontSize: "20pt", color: "#1E293B" },
          },
        ),
      ),
    );
  },
  assert({ validation, preview }) {
    assert.equal(validation.slides.length, 1);
    assert.deepEqual(preview.pages[0]?.diagnostics, []);
    assert.ok(preview.pages[0]?.png instanceof Uint8Array);
  },
};

export const textMechanismVerifications = {
  "text-box": {
    scenario: textBoxRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.presentation.widthEmu, 640 * 9525);
      assert.equal(validation.presentation.heightEmu, 360 * 9525);
      assert.equal(validation.slides.length, 1);
      const s = validation.slides[0]!;
      assert.equal(s.id, "text-roundtrip");
      assert.equal(s.objects.length, 2);

      const headline = s.objects[0]!;
      assert.equal(headline.name, "Headline");
      assert.deepEqual(headline.boxEmu, {
        x: 32 * 9525,
        y: 28 * 9525,
        width: 320 * 9525,
        height: 48 * 9525,
      });
      assert.equal(headline.text?.text, "Imported Headline");
      assert.equal(headline.text?.paragraphs[0]?.runs[0]?.fontSizePt, 24);
      assert.equal(headline.text?.paragraphs[0]?.runs[0]?.color, "#111827");

      const caption = s.objects[1]!;
      assert.equal(caption.name, "Caption");
      assert.deepEqual(caption.boxEmu, {
        x: 32 * 9525,
        y: 96 * 9525,
        width: 320 * 9525,
        height: 64 * 9525,
      });
      assert.equal(caption.text?.text, "Roundtrip Body Text");
      assert.equal(caption.text?.paragraphs[0]?.alignment, "ctr");
      assert.equal(caption.text?.paragraphs[0]?.runs[0]?.fontSizePt, 14);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      // Unembedded imported text preserves native semantics while preview reports missing outline font.
      assert.equal(roundtrip.pages.length, 1);
      assert.ok((roundtrip.pages[0]?.diagnostics.length ?? 0) > 0);
    },
  } satisfies MechanismVerificationEntry<typeof textBoxMechanism>,
  "text-field-slidenum": {
    scenario: textFieldSlideNumRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      assert.equal(validation.slides.length, 2);
      assert.equal(
        validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[1]?.field,
        "slidenum",
      );
      assert.equal(
        validation.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[1]?.field,
        "slidenum",
      );
      assert.equal(
        validation.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[1]?.text,
        "2",
      );
    },
    assertPaint({ initial, roundtrip }) {
      for (const preview of [initial, roundtrip]) {
        assert.equal(preview.pages.length, 2);
        for (const page of preview.pages) {
          assert.deepEqual(page.diagnostics, []);
          assert.ok(page.png instanceof Uint8Array);
        }
      }
    },
  } satisfies MechanismVerificationEntry<typeof textFieldSlideNumMechanism>,
  "embedded-font": {
    scenario: embeddedFontRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const entries = unzipSync(pptx);
      const fontEntries = Object.keys(entries).filter(
        (key) => key.startsWith("ppt/fonts/") && key.endsWith(".fntdata"),
      );
      assert.ok(fontEntries.length >= 2, `${stage} expected embedded .fntdata`);
      const presXml = strFromU8(entries["ppt/presentation.xml"]!);
      assert.match(presXml, /embedTrueTypeFonts="1"/u);
      assert.match(presXml, /<p:embeddedFontLst>/u);
      const slideXml = strFromU8(entries["ppt/slides/slide1.xml"]!);
      assert.match(slideXml, /typeface="\+mn-lt"/u);
    },
  } satisfies MechanismVerificationEntry<typeof embeddedFontMechanism>,
} satisfies ExhaustiveMechanismVerificationMap;
