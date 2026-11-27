import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { createCodegen } from "@pptxdk/codegen";
import { CommonContainer, CommonNode, defineTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { describe, it } from "vitest";

import {
  patchSfntOs2AndPost,
  replaceSfntNameTable,
} from "../../../test/sfnt-test-helpers.ts";
import { documentCapabilities } from "../../all.ts";
import {
  box,
  cell,
  chart,
  deck,
  defineSlideLayout,
  notes,
  paragraph,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  slideNumber,
  span,
  table,
  text,
} from "../../document.ts";
import { font } from "../font.ts";

const resolveDocument = createResolver(documentCapabilities);
const importDocument = createImporter(documentCapabilities);
const generateCode = createCodegen(documentCapabilities);

const jostBook = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostBold = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
  ),
);
const jostItalic = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Italic.ttf", import.meta.url),
  ),
);
const jostBoldItalic = new Uint8Array(
  readFileSync(
    new URL(
      "../../../test/assets/fonts/Jost-700-BoldItalic.ttf",
      import.meta.url,
    ),
  ),
);

function collectNodesByType(
  root: CommonNode,
  type: string,
  out: CommonNode[] = [],
): CommonNode[] {
  if (root.type === type) out.push(root);
  if (root instanceof CommonContainer) {
    for (const child of root.children) {
      if (child instanceof CommonNode) {
        collectNodesByType(child, type, out);
      }
    }
  }
  return out;
}

function namedFontBytes(
  source: Uint8Array,
  familyName: string,
  subfamily: string,
): Uint8Array {
  return replaceSfntNameTable(source, [
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 1,
      value: familyName,
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 2,
      value: subfamily,
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 4,
      value: `${familyName} ${subfamily}`,
    },
    {
      platformID: 3,
      encodingID: 1,
      languageID: 0x0409,
      nameID: 5,
      value: "Version 1.0",
    },
  ]);
}

describe("Imported embedded font weight preservation & backfill", () => {
  it("restores 1..1000 slot weights on FontFace and backfills exact numeric fontWeight across text, spans, slideNumber, shape, table cells, and notes", async () => {
    const jost350 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 350 }),
      "Jost*",
      "Regular",
    );
    const jost650 = namedFontBytes(
      patchSfntOs2AndPost(jostBold, {
        weightClass: 650,
        bold: true,
      }),
      "Jost*",
      "Bold",
    );
    const jostItalic350 = namedFontBytes(
      patchSfntOs2AndPost(jostItalic, {
        weightClass: 350,
        italic: true,
      }),
      "Jost*",
      "Italic",
    );
    const jostBoldItalic800 = namedFontBytes(
      patchSfntOs2AndPost(jostBoldItalic, {
        weightClass: 800,
        bold: true,
        italic: true,
      }),
      "Jost*",
      "Bold Italic",
    );
    const jostMedium500 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 500, bold: false }),
      "Jost Medium",
      "Regular",
    );

    const fourSlotCustomFont = font("Jost*", [
      { data: jost350 },
      { data: jost650 },
      { data: jostItalic350 },
      { data: jostBoldItalic800 },
    ]);
    const mediumThemeFont = font("Jost Medium", [{ data: jostMedium500 }]);

    const doc = deck({
      width: 800,
      height: 450,
      theme: defineTheme({
        fonts: { major: fourSlotCustomFont, minor: mediumThemeFont },
      }),
    }).append(
      slide({ id: "s1" }).append(
        text(
          [
            paragraph([
              span("Theme500 "),
              span("Demi350 ", {
                style: { fontFamily: fourSlotCustomFont, fontWeight: 350 },
              }),
              span("Semi650 ", {
                style: { fontFamily: fourSlotCustomFont, fontWeight: 650 },
              }),
              span("Ital350 ", {
                style: {
                  fontFamily: fourSlotCustomFont,
                  fontWeight: 350,
                  fontStyle: "italic",
                },
              }),
              span("ExtraBoldItal800 ", {
                style: {
                  fontFamily: fourSlotCustomFont,
                  fontWeight: 800,
                  fontStyle: "italic",
                },
              }),
              slideNumber({ style: { fontFamily: "minor", fontWeight: 500 } }),
            ]),
          ],
          {
            id: "rich-runs",
            name: "rich-runs",
            style: {
              position: "absolute",
              left: 32,
              top: 24,
              width: 680,
              height: 80,
              fontFamily: "minor",
              fontWeight: 500,
              fontSize: 18,
              color: "#0F172A",
            },
          },
        ),
        shape("Shape 350", {
          id: "badge-shape",
          name: "badge-shape",
          style: {
            position: "absolute",
            left: 32,
            top: 120,
            width: 240,
            height: 60,
            background: "#E2E8F0",
            fontFamily: fourSlotCustomFont,
            fontWeight: 350,
            fontSize: 16,
            color: "#0F172A",
          },
        }),
        table({
          id: "weights-table",
          name: "weights-table",
          columns: [220, 260],
          style: {
            position: "absolute",
            left: 32,
            top: 200,
            width: 480,
          },
        }).append(
          row().append(
            cell("Cell 500", {
              style: { fontFamily: "minor", fontWeight: 500, fontSize: 14 },
            }),
            cell([
              paragraph([
                span("Cell350 ", {
                  style: { fontFamily: fourSlotCustomFont, fontWeight: 350 },
                }),
                span("Cell650", {
                  style: { fontFamily: fourSlotCustomFont, fontWeight: 650 },
                }),
              ]),
            ]),
          ),
          row({ style: { height: 40 } }).append(
            cell("Hoisted 500", {
              style: { fontFamily: "minor", fontWeight: 500, fontSize: 14 },
            }),
            cell("Hoisted 350", {
              style: {
                fontFamily: fourSlotCustomFont,
                fontWeight: 350,
                fontSize: 14,
              },
            }),
          ),
        ),
        notes("Speaker notes 350", {
          style: {
            fontFamily: fourSlotCustomFont,
            fontWeight: 350,
            fontSize: 12,
          },
        }),
      ),
    );

    for (const embedFonts of ["full", "subset"] as const) {
      const initialEmit = emit(resolveDocument(doc), { embedFonts });
      const imported = importDocument(new Uint8Array(initialEmit.pptx));
      assert.deepEqual(imported.report.diagnostics, []);

      // Theme fonts preserve exact 1..1000 FontFace weights.
      const importedTheme = imported.document.props["theme"] as {
        readonly fonts?: {
          readonly major?: {
            readonly faces?: readonly {
              readonly weight?: number;
              readonly style?: string;
            }[];
          };
          readonly minor?: {
            readonly faces?: readonly {
              readonly weight?: number;
              readonly style?: string;
            }[];
          };
        };
      };
      assert.deepEqual(
        importedTheme.fonts?.major?.faces?.map((f) => [f.weight, f.style]),
        [
          [350, "normal"],
          [650, "normal"],
          [350, "italic"],
          [800, "italic"],
        ],
      );
      assert.deepEqual(
        importedTheme.fonts?.minor?.faces?.map((f) => [f.weight, f.style]),
        [[500, "normal"]],
      );

      // Rich runs & slideNumber backfill exact slot weights.
      const richRunsNode = collectNodesByType(imported.document, "text")[0];
      assert.ok(richRunsNode !== undefined);
      const spans = collectNodesByType(richRunsNode, "span");
      const slideNumbers = collectNodesByType(richRunsNode, "slide-number");
      assert.equal(spans.length, 5);
      assert.equal(
        (spans[0]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        500,
      );
      assert.equal(
        (spans[1]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );
      assert.equal(
        (spans[2]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        650,
      );
      assert.equal(
        (spans[3]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );
      assert.equal(
        (spans[4]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        800,
      );
      assert.equal(slideNumbers.length, 1);
      assert.equal(
        (slideNumbers[0]!.props["style"] as { fontWeight?: unknown })
          .fontWeight,
        500,
      );

      // Shape backfill.
      const shapeNode = collectNodesByType(imported.document, "shape")[0];
      assert.ok(shapeNode !== undefined);
      assert.equal(
        (shapeNode.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );

      // Table cells backfill (both paragraph/span and hoisted cell.style).
      const tableNode = collectNodesByType(imported.document, "table")[0];
      assert.ok(tableNode !== undefined);
      const cells = collectNodesByType(tableNode, "cell");
      assert.equal(cells.length, 4);
      const cell1Span = collectNodesByType(cells[0]!, "span")[0];
      assert.ok(cell1Span !== undefined);
      assert.equal(
        (cell1Span.props["style"] as { fontWeight?: unknown }).fontWeight,
        500,
      );
      const cell2Spans = collectNodesByType(cells[1]!, "span");
      assert.equal(
        (cell2Spans[0]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );
      assert.equal(
        (cell2Spans[1]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        650,
      );
      assert.equal(
        (cells[2]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        500,
      );
      assert.equal(
        (cells[3]!.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );

      // Notes backfill (on notes paragraph/span).
      const notesNodes = collectNodesByType(imported.document, "notes");
      assert.equal(notesNodes.length, 1);
      const notesSpan = collectNodesByType(notesNodes[0]!, "span")[0];
      assert.ok(notesSpan !== undefined);
      assert.equal(
        (notesSpan.props["style"] as { fontWeight?: unknown }).fontWeight,
        350,
      );

      // Resolve, paint, and codegen (vanilla & react) succeed cleanly with exact weights.
      const reResolved = resolveDocument(imported.document);
      assert.deepEqual(
        reResolved.report.fonts.map((u) => [u.family, u.weight, u.hasResource]),
        [
          ["Jost Medium", 500, true],
          ["Jost*", 350, true],
          ["Jost*", 650, true],
          ["Jost*", 350, true],
          ["Jost*", 800, true],
        ],
      );
      const painted = paint(reResolved);
      assert.deepEqual(painted.pages[0]?.errors, []);

      for (const target of ["vanilla", "react"] as const) {
        const generated = generateCode(imported.document, { target });
        const source = generated.files[generated.entry] as string;
        assert.match(source, /weight:\s*350\b/u);
        assert.match(source, /weight:\s*500\b/u);
        assert.match(source, /weight:\s*650\b/u);
        assert.match(source, /weight:\s*800\b/u);
        assert.match(source, /fontWeight:\s*350\b/u);
        assert.match(source, /fontWeight:\s*500\b/u);
        assert.match(source, /fontWeight:\s*650\b/u);
        assert.match(source, /fontWeight:\s*800\b/u);
      }

      const reEmitted = emit(reResolved, { embedFonts });
      const validation = await validatePptx(reEmitted.pptx);
      assert.deepEqual(validation.errors, []);
    }
  });

  it("backfills exact slot weights across master textStyles, master/layout placeholders, footers, bound placeholders, chart text styles, and mixed regular/italic/600-unlinked slots", async () => {
    const jost350 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 350, bold: false }),
      "Jost*",
      "Regular",
    );
    const jost650 = namedFontBytes(
      patchSfntOs2AndPost(jostBold, { weightClass: 650, bold: true }),
      "Jost*",
      "Bold",
    );
    const jostOblique350 = namedFontBytes(
      patchSfntOs2AndPost(jostItalic, {
        weightClass: 350,
        italic: false,
        oblique: true,
        italicAngle: -11,
      }),
      "Jost*",
      "Oblique",
    );
    const jostMedium500 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 500, bold: false }),
      "Jost Medium",
      "Regular",
    );
    const jostMediumItalic400 = namedFontBytes(
      patchSfntOs2AndPost(jostItalic, { weightClass: 400, italic: true }),
      "Jost Medium",
      "Italic",
    );
    const jostExtraBold800 = namedFontBytes(
      patchSfntOs2AndPost(jostBold, { weightClass: 800, bold: true }),
      "Jost Medium",
      "Bold",
    );
    const semiBoldUnlinked600 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 600, bold: false }),
      "Jost SemiBold",
      "Regular",
    );

    const primaryFont = font("Jost*", [
      { data: jost350 },
      { data: jost650 },
      { data: jostOblique350 },
    ]);
    const mediumMixedFont = font("Jost Medium", [
      { data: jostMedium500 },
      { data: jostMediumItalic400 },
      { data: jostExtraBold800 },
    ]);
    const semiBoldUnlinkedFont = font("Jost SemiBold", [
      { data: semiBoldUnlinked600 },
    ]);

    const masterTheme = defineTheme({
      name: "Master Weights Theme",
      fonts: { major: primaryFont, minor: mediumMixedFont },
    });

    const contentLayout = defineSlideLayout({
      id: "content-weights",
      name: "Content Weights",
      type: "obj",
      preserve: true,
      themeOverride: {
        colors: { accent1: "#1D4ED8" },
        fonts: { minor: semiBoldUnlinkedFont },
      },
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 48,
            top: 28,
            width: 700,
            height: 56,
            fontFamily: "major",
            fontWeight: 350,
            fontSize: "28pt",
            color: "#0F172A",
          },
          listStyles: {
            1: { fontFamily: "major", fontWeight: 350, fontSize: "28pt" },
            2: { fontFamily: "major", fontWeight: 650, fontSize: "20pt" },
          },
        },
        body: {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 48,
            top: 100,
            width: 340,
            height: 220,
            fontFamily: "minor",
            fontWeight: 600,
            fontSize: "14pt",
            color: "#1E293B",
          },
          listStyles: {
            1: { fontFamily: "minor", fontWeight: 600, fontSize: "14pt" },
            2: {
              fontFamily: mediumMixedFont,
              fontWeight: 800,
              fontSize: "12pt",
            },
          },
        },
      },
    });

    const doc = deck({
      width: 800,
      height: 450,
      theme: masterTheme,
    }).append(
      slideMaster({
        id: "master-1",
        name: "Weights Master",
        theme: masterTheme,
        preserve: true,
        textStyles: {
          titleStyle: {
            1: {
              fontFamily: "major",
              fontWeight: 350,
              fontSize: "30pt",
              color: "#0F172A",
            },
            2: { fontFamily: "major", fontWeight: 650, fontSize: "22pt" },
          },
          bodyStyle: {
            1: {
              fontFamily: mediumMixedFont,
              fontWeight: 500,
              fontSize: "15pt",
            },
            2: {
              fontFamily: mediumMixedFont,
              fontWeight: 800,
              fontSize: "13pt",
            },
          },
          otherStyle: {
            1: {
              fontFamily: mediumMixedFont,
              fontWeight: 500,
              fontSize: "10pt",
            },
          },
        },
        placeholders: {
          title: {
            type: "title",
            style: {
              position: "absolute",
              left: 48,
              top: 28,
              width: 700,
              height: 56,
              fontFamily: "major",
              fontWeight: 350,
              fontSize: "28pt",
            },
            listStyles: {
              1: { fontFamily: "major", fontWeight: 350, fontSize: "28pt" },
            },
          },
        },
        footerPlaceholders: {
          dateTime: {
            style: {
              position: "absolute",
              left: 48,
              top: 410,
              width: 180,
              height: 24,
              fontFamily: mediumMixedFont,
              fontWeight: 500,
              fontSize: "10pt",
            },
          },
          footer: {
            style: {
              position: "absolute",
              left: 260,
              top: 410,
              width: 280,
              height: 24,
              fontFamily: primaryFont,
              fontWeight: 350,
              fontSize: "10pt",
            },
          },
          slideNumber: {
            style: {
              position: "absolute",
              left: 600,
              top: 410,
              width: 120,
              height: 24,
              fontFamily: semiBoldUnlinkedFont,
              fontWeight: 600,
              fontSize: "10pt",
            },
          },
        },
      }).append(slideLayout(contentLayout)),
      slide({
        id: "s-bound",
        layout: contentLayout,
        dateTime: { text: "2026-11-27", style: { fontWeight: 800 } },
        footer: {
          text: "Oblique Footer",
          style: { fontStyle: "italic", fontWeight: 350 },
        },
        slideNumber: { style: { fontWeight: 600 } },
      }).append(
        text(
          [
            paragraph([
              span("Bound "),
              span("Title650", { style: { fontWeight: 650 } }),
            ]),
          ],
          { placeholder: "title" },
        ),
        text(
          [
            paragraph([span("Body600")], {
              style: { list: { type: "bullet", level: 0 } },
            }),
            paragraph([span("Level2-800")], {
              style: { list: { type: "bullet", level: 1 } },
            }),
          ],
          { placeholder: "body" },
        ),
        box({
          style: {
            position: "absolute",
            left: 48,
            top: 335,
            width: 340,
            height: 60,
          },
        }).append(
          text(
            [
              paragraph([
                span("Mixed500 "),
                span("Italic400", {
                  style: { fontStyle: "italic", fontWeight: 400 },
                }),
              ]),
            ],
            {
              style: {
                position: "relative",
                width: 340,
                height: 50,
                fontFamily: mediumMixedFont,
                fontWeight: 500,
                fontSize: 14,
              },
            },
          ),
        ),
        chart({
          style: {
            position: "absolute",
            left: 410,
            top: 100,
            width: 350,
            height: 280,
            fontFamily: primaryFont,
            fontWeight: 350,
            fontSize: 11,
          },
          title: {
            text: "Chart Weight 650",
            style: {
              fontFamily: primaryFont,
              fontWeight: 650,
              fontSize: 13,
            },
          },
          axes: {
            category: { kind: "category" },
            value: {
              kind: "value",
              style: {
                fontFamily: mediumMixedFont,
                fontWeight: 500,
                fontSize: 10,
              },
              title: {
                text: "Axis 800",
                style: {
                  fontFamily: mediumMixedFont,
                  fontWeight: 800,
                  fontSize: 11,
                },
              },
            },
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              dataLabels: {
                value: true,
                style: {
                  fontFamily: semiBoldUnlinkedFont,
                  fontWeight: 600,
                  fontSize: 10,
                },
              },
              series: [{ name: "Latency", values: [12, 18] }],
            },
          ],
        }),
      ),
    );

    const emitted = emit(resolveDocument(doc), { embedFonts: "full" });
    const imported = importDocument(new Uint8Array(emitted.pptx));
    assert.deepEqual(imported.report.diagnostics, []);

    const importedMaster = imported.document.children.find(
      (child): child is CommonContainer =>
        child instanceof CommonNode && child.type === "slide-master",
    );
    assert.ok(importedMaster !== undefined);

    // Master textStyles backfilled across titleStyle, bodyStyle, otherStyle.
    const txStyles = importedMaster.props["textStyles"] as {
      readonly titleStyle?: Record<number, { readonly fontWeight?: unknown }>;
      readonly bodyStyle?: Record<number, { readonly fontWeight?: unknown }>;
      readonly otherStyle?: Record<number, { readonly fontWeight?: unknown }>;
    };
    assert.equal(txStyles.titleStyle?.[1]?.fontWeight, 350);
    assert.equal(txStyles.titleStyle?.[2]?.fontWeight, 650);
    assert.equal(txStyles.bodyStyle?.[1]?.fontWeight, 500);
    assert.equal(txStyles.bodyStyle?.[2]?.fontWeight, 800);
    assert.equal(txStyles.otherStyle?.[1]?.fontWeight, 500);

    // Master placeholders & derived footer placeholders backfilled.
    const masterPh = importedMaster.props["placeholders"] as Record<
      string,
      {
        readonly listStyles?: Record<number, { readonly fontWeight?: unknown }>;
      }
    >;
    assert.equal(masterPh["title"]?.listStyles?.[1]?.fontWeight, 350);
    assert.equal(masterPh["dt"]?.listStyles?.[1]?.fontWeight, 500);
    assert.equal(masterPh["ftr"]?.listStyles?.[1]?.fontWeight, 350);
    assert.equal(masterPh["sldNum"]?.listStyles?.[1]?.fontWeight, 600);

    // Layout placeholders & listStyles backfilled.
    const importedLayout = importedMaster.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide-layout",
    );
    assert.ok(importedLayout !== undefined);
    const layoutPh = importedLayout.props["placeholders"] as Record<
      string,
      {
        readonly listStyles?: Record<number, { readonly fontWeight?: unknown }>;
      }
    >;
    assert.equal(layoutPh["title"]?.listStyles?.[1]?.fontWeight, 350);
    assert.equal(layoutPh["title"]?.listStyles?.[2]?.fontWeight, 650);
    assert.equal(layoutPh["body"]?.listStyles?.[1]?.fontWeight, 600);
    assert.equal(layoutPh["body"]?.listStyles?.[2]?.fontWeight, 800);

    // Slide footer overrides backfilled.
    const importedSlide = imported.document.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide",
    );
    assert.ok(importedSlide !== undefined);
    assert.equal(
      (
        importedSlide.props["dateTime"] as {
          readonly style?: { readonly fontWeight?: unknown };
        }
      ).style?.fontWeight,
      800,
    );
    assert.equal(
      (
        importedSlide.props["footer"] as {
          readonly style?: { readonly fontWeight?: unknown };
        }
      ).style?.fontWeight,
      350,
    );
    assert.equal(
      (
        importedSlide.props["slideNumber"] as {
          readonly style?: { readonly fontWeight?: unknown };
        }
      ).style?.fontWeight,
      600,
    );

    // Bound title placeholder span inherits major -> 350 and overrides bold -> 650.
    const slideTexts = collectNodesByType(importedSlide, "text");
    const boundTitle = slideTexts.find(
      (t) => t.props["placeholder"] === "title",
    );
    assert.ok(boundTitle !== undefined);
    const titleSpans = collectNodesByType(boundTitle, "span");
    assert.equal(titleSpans.length, 1);
    assert.equal(
      (titleSpans[0]!.props["style"] as { fontWeight?: unknown }).fontWeight,
      650,
    );

    // Mixed regular=500 / italic=400 box preserves 500 on regular and restores 400 on italic span.
    const freeMixedText = slideTexts.find(
      (t) => t.props["placeholder"] === undefined,
    );
    assert.ok(freeMixedText !== undefined);
    const mixedSpans = collectNodesByType(freeMixedText, "span");
    assert.equal(
      (mixedSpans[0]!.props["style"] as { fontWeight?: unknown }).fontWeight,
      500,
    );
    const mixedItalicWeight = (
      mixedSpans[1]!.props["style"] as { fontWeight?: unknown }
    ).fontWeight;
    assert.ok(mixedItalicWeight === 400 || mixedItalicWeight === "normal");

    // Chart text styles backfilled.
    const importedChart = collectNodesByType(importedSlide, "chart")[0];
    assert.ok(importedChart !== undefined);
    assert.equal(
      (importedChart.props["style"] as { fontWeight?: unknown }).fontWeight,
      350,
    );
    assert.equal(
      (
        importedChart.props["title"] as {
          readonly style?: { readonly fontWeight?: unknown };
        }
      ).style?.fontWeight,
      650,
    );
    const importedAxes = importedChart.props["axes"] as {
      readonly value?: {
        readonly style?: { readonly fontWeight?: unknown };
        readonly title?: {
          readonly style?: { readonly fontWeight?: unknown };
        };
      };
    };
    assert.equal(importedAxes.value?.style?.fontWeight, 500);
    assert.equal(importedAxes.value?.title?.style?.fontWeight, 800);
    const importedGroups = importedChart.props["groups"] as readonly {
      readonly dataLabels?: {
        readonly style?: { readonly fontWeight?: unknown };
      };
    }[];
    assert.equal(importedGroups[0]?.dataLabels?.style?.fontWeight, 600);

    // Rebuild & validate after import + codegen (vanilla & react).
    const reResolved = resolveDocument(imported.document);
    const painted = paint(reResolved);
    assert.deepEqual(painted.pages[0]?.errors, []);
    const reEmitted = emit(reResolved, { embedFonts: "full" });
    const validation = await validatePptx(reEmitted.pptx);
    assert.deepEqual(validation.errors, []);

    for (const target of ["vanilla", "react"] as const) {
      const generated = generateCode(imported.document, { target });
      const source = generated.files[generated.entry] as string;
      assert.match(source, /weight:\s*350\b/u);
      assert.match(source, /weight:\s*500\b/u);
      assert.match(source, /weight:\s*600\b/u);
      assert.match(source, /weight:\s*650\b/u);
      assert.match(source, /weight:\s*800\b/u);
    }
  });

  it("backfills inherited non-default slot weights on empty text, paragraph, cell, and notes nodes during importPptx", async () => {
    const jost350 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 350 }),
      "Jost*",
      "Regular",
    );
    const jostMedium500 = namedFontBytes(
      patchSfntOs2AndPost(jostBook, { weightClass: 500, bold: false }),
      "Jost Medium",
      "Regular",
    );
    const major350 = font("Jost*", [{ data: jost350 }]);
    const minor500 = font("Jost Medium", [{ data: jostMedium500 }]);

    const emptyLayout = defineSlideLayout({
      id: "empty-leaf-layout",
      name: "Empty Leaf Layout",
      type: "obj",
      placeholders: {
        body: {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 24,
            top: 24,
            width: 240,
            height: 36,
            fontFamily: "minor",
            fontWeight: 500,
          },
        },
      },
    });

    const sourceDeck = deck({
      width: 640,
      height: 360,
      theme: defineTheme({
        name: "Empty Leaf Weights",
        fonts: { major: major350, minor: minor500 },
      }),
    }).append(
      slideMaster({ id: "empty-master", name: "Empty Master" }).append(
        slideLayout(emptyLayout),
      ),
      slide({
        id: "empty-leaf-slide",
        layout: emptyLayout,
        placeholders: {
          body: text("", { id: "empty-placeholder-text" }),
        },
      }).append(
        text("", {
          id: "empty-text",
          style: {
            position: "absolute",
            left: 24,
            top: 240,
            width: 200,
            height: 32,
            fontWeight: 500,
          },
        }),
        text(
          [
            paragraph("First", {
              style: { fontFamily: major350, fontWeight: 350 },
            }),
            paragraph(""),
          ],
          {
            id: "multi-para-text",
            style: {
              position: "absolute",
              left: 24,
              top: 72,
              width: 240,
              height: 60,
              fontWeight: 500,
            },
          },
        ),
        table({
          id: "empty-cell-table",
          columns: [140, 140],
          style: {
            position: "absolute",
            left: 24,
            top: 120,
            width: 280,
            fontWeight: 500,
          },
        }).append(
          row({ style: { height: 28 } }).append(cell(""), cell("Filled cell")),
        ),
        notes("", { style: { fontWeight: 500 } }),
      ),
    );

    const emitted = emit(resolveDocument(sourceDeck), { embedFonts: "full" });
    const imported = importDocument(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedSlide = imported.document.children.find(
      (child): child is CommonNode =>
        child instanceof CommonNode && child.type === "slide",
    );
    assert.ok(importedSlide !== undefined);
    const texts = collectNodesByType(importedSlide, "text");
    const freeEmptyText = texts.find(
      (t) => t.props["placeholder"] === undefined && t.children.length === 0,
    );
    assert.ok(freeEmptyText !== undefined);
    assert.equal(
      (
        freeEmptyText.props["style"] as
          | { readonly fontWeight?: unknown }
          | undefined
      )?.fontWeight,
      500,
    );
    const paraText = texts.find((t) =>
      t.children.some((c) => c instanceof CommonNode && c.type === "paragraph"),
    );
    assert.ok(paraText !== undefined);
    const emptyPara = collectNodesByType(paraText, "paragraph")[1];
    assert.ok(emptyPara !== undefined);
    assert.equal(emptyPara.children.length, 0);
    assert.equal(
      (
        emptyPara.props["style"] as
          | { readonly fontWeight?: unknown }
          | undefined
      )?.fontWeight,
      500,
    );
    const cells = collectNodesByType(importedSlide, "cell");
    assert.equal(
      (
        cells[0]!.props["style"] as
          | { readonly fontWeight?: unknown }
          | undefined
      )?.fontWeight,
      500,
    );
    const importedNotes = collectNodesByType(importedSlide, "notes")[0];
    assert.ok(importedNotes !== undefined);
    const notesWeightTarget =
      collectNodesByType(importedNotes, "paragraph")[0] ?? importedNotes;
    assert.equal(
      (
        notesWeightTarget.props["style"] as
          | { readonly fontWeight?: unknown }
          | undefined
      )?.fontWeight,
      500,
    );

    const reEmitted = emit(resolveDocument(imported.document), {
      embedFonts: "full",
    });
    assert.deepEqual((await validatePptx(reEmitted.pptx)).errors, []);

    for (const target of ["vanilla", "react"] as const) {
      const generated = generateCode(imported.document, { target });
      const source = generated.files[generated.entry] as string;
      assert.match(source, /fontWeight:\s*500\b/u);
      assert.match(source, /fontWeight:\s*350\b/u);
    }
  });
});
