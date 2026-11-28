import assert from "node:assert/strict";

import { CodegenError, build, codegen, importPptx } from "@officedk/pptxdk";
import { createCodegen } from "@officedk/pptxdk/build";
import {
  Deck,
  P,
  Section,
  Slide,
  Span,
  Text,
  expandDocument,
} from "@officedk/pptxdk/react";
import {
  chart,
  deck,
  defineSlideLayout,
  paragraph,
  section,
  shape,
  slide,
  slideLayout,
  slideMaster,
  span,
  text,
} from "@officedk/pptxdk/vanilla";
import { CommonNode } from "@pptxdk/core/document";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

function createSampleVanillaDeck() {
  return deck({ width: 960, height: 540 }).append(
    section({ name: "Opening" }).append(
      slide({ id: "cover" }).append(
        text("Quarterly Review", {
          id: "cover-title",
          style: {
            position: "absolute",
            left: 64,
            top: 80,
            width: 600,
            height: 60,
            fontSize: 28,
            fontWeight: "bold",
            color: "#111827",
          },
        }),
      ),
    ),
    section({ name: "Details" }).append(
      slide({ id: "summary" }).append(
        text(
          [
            paragraph(
              [
                span("Revenue ", { style: { fontWeight: "bold" } }),
                span("grew 18%"),
              ],
              { style: { align: "left" } },
            ),
            paragraph("Next steps"),
          ],
          {
            id: "summary-body",
            style: {
              position: "absolute",
              left: 64,
              top: 100,
              width: 640,
              height: 140,
            },
          },
        ),
      ),
      slide({ id: "closing" }).append(
        text("Thank you", {
          id: "closing-note",
          style: {
            position: "absolute",
            left: 64,
            top: 220,
            width: 480,
            height: 52,
          },
        }),
      ),
    ),
  );
}

async function createSampleReactDeck() {
  return expandDocument(
    createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        Section,
        { name: "Opening" },
        createElement(
          Slide,
          { id: "cover" },
          createElement(
            Text,
            {
              id: "cover-title",
              style: {
                position: "absolute",
                left: 64,
                top: 80,
                width: 600,
                height: 60,
                fontSize: 28,
                fontWeight: "bold",
                color: "#111827",
              },
            },
            "Quarterly Review",
          ),
        ),
      ),
      createElement(
        Section,
        { name: "Details" },
        createElement(
          Slide,
          { id: "summary" },
          createElement(
            Text,
            {
              id: "summary-body",
              style: {
                position: "absolute",
                left: 64,
                top: 100,
                width: 640,
                height: 140,
              },
            },
            createElement(
              P,
              { style: { align: "left" } },
              createElement(
                Span,
                { style: { fontWeight: "bold" } },
                "Revenue ",
              ),
              createElement(Span, null, "grew 18%"),
            ),
            createElement(P, null, "Next steps"),
          ),
        ),
        createElement(
          Slide,
          { id: "closing" },
          createElement(
            Text,
            {
              id: "closing-note",
              style: {
                position: "absolute",
                left: 64,
                top: 220,
                width: 480,
                height: 52,
              },
            },
            "Thank you",
          ),
        ),
      ),
    ),
  );
}

describe("pptxdk public codegen acceptance", () => {
  it("generates typechecked, relocated-rebuildable vanilla and react sources from direct vanilla, expandDocument, and importPptx decks", async () => {
    const vanillaInput = createSampleVanillaDeck();
    const reactInput = await createSampleReactDeck();

    const originalBuilt = await build(vanillaInput);
    const importedResult = importPptx(originalBuilt.pptx);
    assert.deepEqual(importedResult.report.diagnostics, []);

    const sources = [
      { label: "direct-vanilla", doc: vanillaInput },
      { label: "direct-react", doc: reactInput },
      { label: "imported-pptx", doc: importedResult.document },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc }) => {
      const vanillaGen = codegen(doc, { target: "vanilla" });
      const reactGen = codegen(doc, { target: "react" });
      assert.equal(vanillaGen.entry, "deck.ts");
      assert.equal(reactGen.entry, "deck.tsx");
      return [
        { label: `${label}:vanilla`, result: vanillaGen },
        { label: `${label}:react`, result: reactGen },
      ];
    });

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );
      assert.equal(file.slides.length, 3, `${item.label} slide count`);
      assert.deepEqual(
        file.presentation.sections?.map((sec) => ({
          name: sec.name,
          count: sec.slideIds.length,
        })),
        [
          { name: "Opening", count: 1 },
          { name: "Details", count: 2 },
        ],
        `${item.label} sections`,
      );
      assert.deepEqual(
        file.slides.map((s) => s.section),
        ["Opening", "Details", "Details"],
        `${item.label} per-slide section ownership`,
      );
      assert.deepEqual(
        file.slides.map((s) => s.objects.map((obj) => obj.text?.text)),
        [["Quarterly Review"], ["Revenue grew 18%\nNext steps"], ["Thank you"]],
        `${item.label} slide order and text content`,
      );
    }
  }, 60_000);

  it("is deterministic across repeated calls and throws CodegenError with precise path on unexpressible declarations", () => {
    const doc = createSampleVanillaDeck();
    for (const target of ["vanilla", "react"] as const) {
      const first = codegen(doc, { target });
      const second = codegen(doc, { target });
      assert.equal(first.entry, second.entry);
      assert.deepEqual(Object.keys(first.files), Object.keys(second.files));
      for (const key of Object.keys(first.files)) {
        assert.deepEqual(first.files[key], second.files[key]);
      }
    }

    const trimmedCodegen = createCodegen([]);
    const trimmedChartDeck = deck({ width: 960, height: 540 }).append(
      section({ name: "Main" }).append(
        slide({ id: "s-chart" }).append(
          chart({
            id: "sales-chart",
            style: {
              position: "absolute",
              left: 60,
              top: 80,
              width: 400,
              height: 240,
            },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "North", values: [10, 20] }],
              },
            ],
          }),
        ),
      ),
    );

    assert.throws(
      () => trimmedCodegen(trimmedChartDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-declaration" &&
        error.diagnostics[0]?.path === "deck/slide[s-chart]/chart[sales-chart]",
    );

    const unknownDeclDeck = deck({ width: 960, height: 540 }).append(
      section({ name: "Main" }).append(
        slide({ id: "s-unknown" }).append(
          new CommonNode("custom-unsupported", "custom-unsupported", {
            id: "unknown-widget",
          }),
        ),
      ),
    );

    assert.throws(
      () => codegen(unknownDeclDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-declaration" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-unknown]/custom-unsupported[unknown-widget]",
    );

    const invalidNumberDeck = deck({ width: 960, height: 540 }).append(
      slide({ id: "s-bad" }).append(
        text("Bad left", {
          id: "bad-text",
          style: {
            position: "absolute",
            left: Number.NaN,
            top: 20,
            width: 200,
            height: 40,
          },
        }),
      ),
    );

    assert.throws(
      () => codegen(invalidNumberDeck, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s-bad]/text[bad-text]@style.left",
    );
  });

  it("preserves backslashes and HTML entity sequences in JSX string attributes across codegen, typecheck, and relocated rebuild", async () => {
    const escLayout = defineSlideLayout({ id: "esc-layout", name: "Esc" });
    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "m-esc",
        headerFooter: { footer: true, dateTime: false, slideNumber: false },
        footerPlaceholders: { footer: { text: "Default" } },
      }).append(slideLayout(escLayout)),
      section({ name: "Sec C:\\Temp &copy; 2026" }).append(
        slide({
          id: "s-esc",
          layout: escLayout,
          footer: "Path C:\\Temp &copy; 2026",
        }).append(
          shape("C:\\Users\\Alice &lt;tag&gt; &amp; Co.", {
            id: "sh-esc",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 160,
              height: 50,
            },
          }),
        ),
      ),
    );

    const cases = [
      { label: "esc:vanilla", result: codegen(doc, { target: "vanilla" }) },
      { label: "esc:react", result: codegen(doc, { target: "react" }) },
    ];
    await verifyGeneratedSourcesTypecheck(cases);

    for (const item of cases) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, []);
      assert.equal(
        file.presentation.sections?.[0]?.name,
        "Sec C:\\Temp &copy; 2026",
        `${item.label} section name`,
      );
      assert.equal(
        file.slides[0]?.objects[0]?.text?.text,
        "C:\\Users\\Alice &lt;tag&gt; &amp; Co.",
        `${item.label} shape text`,
      );
      assert.equal(
        file.slides[0]?.footerPlaceholders?.[0]?.text,
        "Path C:\\Temp &copy; 2026",
        `${item.label} slide footer`,
      );
    }
  }, 60_000);
});
