import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
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
  span,
  sub,
  sup,
  table,
  text,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import { slideBackgroundCapability } from "@pptxdk/capabilities/slide-background";
import { textCapability } from "@pptxdk/capabilities/text";
import {
  P,
  SlideNumber,
  Span,
  Sub,
  Sup,
  Text,
} from "@pptxdk/capabilities/text-elements";
import { createCodegen } from "@pptxdk/codegen";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode, RichTextContainer } from "@pptxdk/core/document";
import { emit, pxToEmu } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

const resolveText = createResolver([slideBackgroundCapability, textCapability]);

describe("Text & font public contract schemas", () => {
  it("resolves valid textList, paragraphStyle, runStyle, and insets across public JSX and resolver boundaries", async () => {
    const validDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "ok",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 240,
                height: 140,
                padding: 10,
                paddingLeft: 20,
              },
            },
            createElement(
              P,
              {
                style: {
                  align: "center",
                  lineHeight: "24pt",
                  spacingAfter: 8,
                  fontSize: "32pt",
                  fontFamily: "minor",
                  list: { type: "bullet", character: "•", level: 2 },
                },
              },
              createElement(
                Span,
                {
                  style: {
                    fontSize: 32,
                    fontFamily: "major",
                    letterSpacing: 1.5,
                    textDecoration: "underline",
                    textTransform: "uppercase",
                    lang: "en-US",
                  },
                },
                "Run",
              ),
            ),
            createElement(
              P,
              {
                style: {
                  list: {
                    type: "number",
                    style: "arabicPeriod",
                    startAt: 3,
                    level: 1,
                  },
                },
              },
              "Page ",
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );

    const resolution = resolveText(validDoc);
    assert.equal(resolution.slides.length, 1);
  });

  it("maps invalid list, run/paragraph style, insets, SlideNumber children, and font options onto typed errors", async () => {
    const invalidCases = [
      {
        label: "multi-char bullet",
        code: "invalid-style",
        el: createElement(
          P,
          { style: { list: { type: "bullet", character: "ab" } } },
          "bad",
        ),
      },
      {
        label: "zero number startAt",
        code: "invalid-style",
        el: createElement(
          P,
          { style: { list: { type: "number", startAt: 0 } } },
          "bad",
        ),
      },
      {
        label: "verticalAlign on Paragraph",
        code: "unknown-prop",
        el: createElement(
          P,
          { style: { verticalAlign: "top" } as never },
          "bad",
        ),
      },
      {
        label: "align on Span",
        code: "unknown-prop",
        el: createElement(
          P,
          null,
          createElement(
            Span,
            // @ts-expect-error -- align is invalid on Span.style
            { style: { align: "center" } },
            "bad span",
          ),
        ),
      },
      {
        label: "children on SlideNumber",
        code: "invalid-prop",
        el: createElement(SlideNumber, null, "child"),
      },
    ] as const;

    for (const item of invalidCases) {
      const doc = await expandDocument(
        createElement(
          Deck,
          { width: 480, height: 270 },
          createElement(
            Slide,
            { id: "s1" },
            createElement(
              Text,
              {
                id: "bad",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 200,
                  height: 40,
                },
              },
              item.el,
            ),
          ),
        ),
      );
      assert.throws(
        () => resolveText(doc),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === item.code),
        `${item.label} should fail with ${item.code}`,
      );
    }

    const badPaddingDoc = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "s1" },
          createElement(
            Text,
            {
              id: "bad-pad",
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 200,
                height: 40,
                padding: -2,
              },
            },
            "bad padding",
          ),
        ),
      ),
    );
    assert.throws(
      () => resolveText(badPaddingDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((d) => d.code === "invalid-style"),
    );

    assert.throws(
      // @ts-expect-error -- substitute must be boolean
      () => font("Jost*", [], { substitute: "yes" }),
      TypeError,
    );
  });

  it("measures content-sized proportional lineHeight from font metrics, preserves <a:spcPct> on emit, and roundtrips across Text, Shape, Master, and Chart", () => {
    const resolveAll = createResolver(documentCapabilities);
    const importAll = createImporter(documentCapabilities);
    const codegenAll = createCodegen(documentCapabilities);

    const doc = deck({ width: 640, height: 360 }).append(
      slideMaster({
        id: "m1",
        preserve: true,
        textStyles: {
          bodyStyle: {
            1: { lineHeight: "130%" },
          },
        },
      }).append(slideLayout({ name: "L1", preserve: true })),
      slide({ id: "s1" }).append(
        text("Single line baseline", {
          id: "auto-100",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 300,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
          },
        }),
        text("Single line proportional", {
          id: "auto-150",
          style: {
            position: "absolute",
            left: 20,
            top: 80,
            width: 300,
            height: "auto",
            fontFamily: jostFont,
            fontSize: 20,
            lineHeight: "150%",
          },
        }),
        shape("Shape proportional", {
          id: "shp-pct",
          style: {
            position: "absolute",
            left: 20,
            top: 160,
            width: 220,
            height: 60,
            fontFamily: jostFont,
            fontSize: 16,
            lineHeight: "140%",
          },
        }),
        chart({
          id: "cht-pct",
          title: {
            text: "Chart Title",
            style: { fontSize: 14, lineHeight: "125%" },
          },
          style: {
            position: "absolute",
            left: 260,
            top: 140,
            width: 340,
            height: 180,
            fontSize: 12,
            lineHeight: "120%",
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S1", values: [1, 2] }],
            },
          ],
        }),
      ),
    );

    const res = resolveAll(doc);
    const h100 = res.slide("s1").get("auto-100").box!.height;
    const h150 = res.slide("s1").get("auto-150").box!.height;
    assert.ok(
      Math.abs(h150 - h100 * 1.5) < 0.05,
      `Expected 150% height (${h150}) to be 1.5x natural height (${h100})`,
    );

    const emitted = emit(res);
    const parts = unzipSync(emitted.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const masterXml = strFromU8(parts["ppt/slideMasters/slideMaster1.xml"]!);
    const chartXml = strFromU8(parts["ppt/charts/chart1.xml"]!);

    assert.match(slideXml, /<a:lnSpc><a:spcPct val="150000"\/><\/a:lnSpc>/u);
    assert.match(slideXml, /<a:lnSpc><a:spcPct val="140000"\/><\/a:lnSpc>/u);
    assert.match(masterXml, /<a:lnSpc><a:spcPct val="130000"\/><\/a:lnSpc>/u);
    assert.match(chartXml, /<a:lnSpc><a:spcPct val="125000"\/><\/a:lnSpc>/u);
    assert.match(chartXml, /<a:lnSpc><a:spcPct val="120000"\/><\/a:lnSpc>/u);

    const imported = importAll(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedMaster = imported.document.children.find(
      (c): c is CommonNode =>
        c instanceof CommonNode && c.type === "slide-master",
    );
    assert.equal(
      (
        importedMaster?.props["textStyles"] as {
          bodyStyle?: { 1?: { lineHeight?: unknown } };
        }
      )?.bodyStyle?.[1]?.lineHeight,
      "130%",
    );

    const importedSlide = imported.document.query.get("s1").node;
    const importedAuto150 = importedSlide.children.find(
      (c): c is RichTextContainer =>
        c instanceof RichTextContainer && c.props["name"] === "Text 2",
    );
    assert.equal(
      (
        importedAuto150?.props["style"] as Record<string, unknown> | undefined
      )?.["lineHeight"],
      "150%",
    );

    const importedShape = importedSlide.children.find(
      (c): c is RichTextContainer =>
        c instanceof RichTextContainer && c.props["name"] === "Shape 3",
    );
    assert.equal(
      (importedShape?.props["style"] as Record<string, unknown> | undefined)?.[
        "lineHeight"
      ],
      "140%",
    );

    const importedChart = importedSlide.children.find(
      (c): c is CommonNode => c instanceof CommonNode && c.type === "chart",
    );
    assert.equal(
      (
        importedChart?.props["title"] as {
          style?: { lineHeight?: unknown };
        }
      )?.style?.lineHeight,
      "125%",
    );
    assert.equal(
      (importedChart?.props["style"] as Record<string, unknown> | undefined)?.[
        "lineHeight"
      ],
      "120%",
    );

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegenAll(imported.document, { target });
      const entry = generated.files[generated.entry] as string;
      assert.match(entry, /"150%"/u);
      assert.match(entry, /"140%"/u);
      assert.match(entry, /"130%"/u);
      assert.match(entry, /"125%"/u);
      assert.match(entry, /"120%"/u);
    }

    // Rejects zero, negative, out-of-range percentage, and unitless string lineHeight.
    for (const badLineHeight of ["0%", "-20%", "13201%", "1.5"] as const) {
      const badDoc = deck({ width: 320, height: 180 }).append(
        slide({ id: "bad" }).append(
          text(
            [
              paragraph("Bad", {
                style: { lineHeight: badLineHeight as never },
              }),
            ],
            {
              id: "t-bad",
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
        () => resolveAll(badDoc),
        (err: unknown) =>
          err instanceof BuildError &&
          err.diagnostics.some((d) => d.code === "invalid-style"),
      );
    }
  });

  it("supports inline Sup, Sub, and baselineShift across Text/Shape/Cell/Notes/Placeholder with scaling, conflict rejection, and lossless import", async () => {
    const resolveAll = createResolver(documentCapabilities);
    const importAll = createImporter(documentCapabilities);

    const phLayout = defineSlideLayout({
      name: "SupSubLayout",
      placeholders: {
        title: {
          type: "title",
          style: { left: 20, top: 20, width: 400, height: 50 },
        },
      },
    });

    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          SlideMaster,
          { id: "m-sub" },
          createElement(SlideLayout, { layout: phLayout }),
        ),
        createElement(
          Slide,
          {
            id: "s-jsx",
            layout: phLayout,
            placeholders: {
              title: createElement(Sup, null, "1"),
            },
          },
          createElement(
            Text,
            {
              id: "chem",
              style: {
                position: "absolute",
                left: 20,
                top: 90,
                width: 300,
                height: "auto",
                fontFamily: jostFont,
                fontSize: 20,
              },
            },
            "H",
            createElement(Sub, null, "2"),
            "O + E=mc",
            createElement(
              Span,
              { style: { color: "#DC2626" } },
              createElement(Sup, null, "2"),
            ),
            createElement(Span, { style: { baselineShift: "15%" } }, " custom"),
          ),
          createElement(
            Text,
            {
              id: "single-sup",
              style: {
                position: "absolute",
                left: 20,
                top: 150,
                width: "auto",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 24,
              },
            },
            createElement(Sup, null, "WWWW"),
          ),
          createElement(
            Text,
            {
              id: "single-normal",
              style: {
                position: "absolute",
                left: 20,
                top: 200,
                width: "auto",
                height: "auto",
                fontFamily: jostFont,
                fontSize: 24,
              },
            },
            "WWWW",
          ),
        ),
      ),
    );

    const res = resolveAll(jsxDoc);
    const supBox = res.slide("s-jsx").get("single-sup").box!;
    const normalBox = res.slide("s-jsx").get("single-normal").box!;
    // Superscript scales advance width (~0.65x) while keeping natural line height (~1.0x).
    assert.ok(
      supBox.width < normalBox.width * 0.72 &&
        supBox.width > normalBox.width * 0.58,
      `Expected sup width ${supBox.width} to be ~0.65 of ${normalBox.width}`,
    );
    assert.ok(Math.abs(supBox.height - normalBox.height) < 0.01);

    const vanillaDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "m-sub" }).append(slideLayout(phLayout)),
      slide({
        id: "s-van",
        layout: phLayout,
        placeholders: { title: sup("1") },
      }).append(
        text(
          [
            paragraph([
              "x",
              sup("2"),
              sub("i"),
              span("s", { style: { baselineShift: "-12.5%" } }),
            ]),
          ],
          {
            id: "v-text",
            style: {
              position: "absolute",
              left: 20,
              top: 90,
              width: 300,
              height: 50,
              fontFamily: jostFont,
            },
          },
        ),
        shape(["CO", sub("2")], {
          id: "v-shape",
          style: {
            position: "absolute",
            left: 20,
            top: 160,
            width: 200,
            height: 50,
            fontFamily: jostFont,
          },
        }),
        table({
          id: "v-table",
          columns: [160],
          style: { position: "absolute", left: 20, top: 220, width: 160 },
        }).append(
          row({ style: { height: 40 } }).append(
            cell(["m", sup("2")], { style: { fontFamily: jostFont } }),
          ),
        ),
        notes([paragraph(["Note", sup("*")])]),
      ),
    );

    const emitted = emit(resolveAll(vanillaDoc));
    const parts = unzipSync(emitted.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    const notesPath = Object.keys(parts).find((name) =>
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(name),
    );
    assert.ok(
      notesPath !== undefined,
      `Expected notes slide part in ${JSON.stringify(Object.keys(parts))}`,
    );
    const notesXml = strFromU8(parts[notesPath]!);

    assert.match(slideXml, /baseline="30000"/u);
    assert.match(slideXml, /baseline="-25000"/u);
    assert.match(slideXml, /baseline="-12500"/u);
    assert.match(notesXml, /baseline="30000"/u);

    const imported = importAll(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const impSlide = imported.document.query.get("s-van").node;
    const impText = impSlide.children.find(
      (c): c is RichTextContainer =>
        c instanceof RichTextContainer && c.props["name"] === "Text 2",
    );
    const impPara = impText?.children[0] as RichTextContainer;
    assert.equal(impPara.children[1]?.type, "sup");
    assert.equal(
      (
        (impPara.children[1] as RichTextContainer).props["style"] as
          | Record<string, unknown>
          | undefined
      )?.["baselineShift"],
      undefined,
    );
    assert.equal(impPara.children[2]?.type, "sub");
    assert.equal(impPara.children[3]?.type, "span");
    assert.equal(
      (
        (impPara.children[3] as RichTextContainer).props["style"] as Record<
          string,
          unknown
        >
      )?.["baselineShift"],
      "-12.5%",
    );

    // Ensure single-sup in JSX slide is NOT flattened into Text.style!
    const importedJsx = importAll(emit(res).pptx);
    const impSingleSup = importedJsx.document.query
      .get("s-jsx")
      .node.children.find(
        (c): c is RichTextContainer =>
          c instanceof RichTextContainer && c.props["name"] === "Text 3",
      );
    assert.equal(
      (impSingleSup?.props["style"] as Record<string, unknown> | undefined)?.[
        "baselineShift"
      ],
      undefined,
    );
    assert.equal(
      (impSingleSup?.children[0] as RichTextContainer | undefined)?.children[0]
        ?.type,
      "sup",
    );

    // Rejects invalid/conflicting baselineShift.
    for (const badRun of [
      span("zero", { style: { baselineShift: "0%" } }),
      span("negzero", { style: { baselineShift: "-0%" } }),
      span("tiny", { style: { baselineShift: "0.0001%" } }),
      sup("conflict", { style: { baselineShift: "-10%" } }),
      sub("conflict", { style: { baselineShift: "30%" } }),
      sup(sub("nested-conflict")),
    ]) {
      const badDoc = deck({ width: 320, height: 180 }).append(
        slide({ id: "bad" }).append(
          text([paragraph([badRun])], {
            id: "t-bad",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 120,
              height: 40,
            },
          }),
        ),
      );
      assert.throws(
        () => resolveAll(badDoc),
        (err: unknown) =>
          err instanceof BuildError &&
          err.diagnostics.some((d) => d.code === "invalid-style"),
      );
    }

    // Rejects baselineShift on container Text or Paragraph P.
    const badContainerDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "bad-c" }).append(
        text("bad", {
          id: "t-bad-c",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 120,
            height: 40,
            baselineShift: "10%",
          } as never,
        }),
      ),
    );
    assert.throws(
      () => resolveAll(badContainerDoc),
      (err: unknown) =>
        err instanceof BuildError &&
        err.diagnostics.some((d) => d.code === "unknown-prop"),
    );
  });

  it("materializes hanging list indents on non-placeholder Text/Shape/Cell without injecting defaults onto placeholders", () => {
    const resolveAll = createResolver(documentCapabilities);
    const importAll = createImporter(documentCapabilities);

    const bodyLayout = defineSlideLayout({
      name: "ListLayout",
      placeholders: {
        body: {
          type: "body",
          style: { left: 20, top: 20, width: 300, height: 100 },
        },
      },
    });

    const doc = deck({ width: 800, height: 450 }).append(
      slideMaster({ id: "m-list" }).append(slideLayout(bodyLayout)),
      slide({
        id: "s-list",
        layout: bodyLayout,
        placeholders: {
          body: [
            paragraph("Placeholder item", {
              style: { list: { type: "bullet" } },
            }),
          ],
        },
      }).append(
        text(
          [
            paragraph("Bullet L0", {
              style: { list: { type: "bullet", character: "•", level: 0 } },
            }),
            paragraph("Number L1", {
              style: {
                list: {
                  type: "number",
                  style: "arabicPeriod",
                  startAt: 2,
                  level: 1,
                },
              },
            }),
          ],
          {
            id: "free-text",
            style: {
              position: "absolute",
              left: 20,
              top: 130,
              width: 300,
              height: 100,
              fontFamily: jostFont,
              fontSize: 18,
            },
          },
        ),
        shape(
          [
            paragraph("Shape bullet", {
              style: { list: { type: "bullet", character: "-", level: 0 } },
            }),
          ],
          {
            id: "free-shape",
            style: {
              position: "absolute",
              left: 340,
              top: 130,
              width: 240,
              height: 80,
              fontFamily: jostFont,
              fontSize: 18,
            },
          },
        ),
        table({
          id: "free-table",
          columns: [240],
          style: { position: "absolute", left: 20, top: 250, width: 240 },
        }).append(
          row({ style: { height: 60 } }).append(
            cell(
              [
                paragraph("Cell bullet", {
                  style: { list: { type: "bullet", character: "•", level: 0 } },
                }),
              ],
              { style: { fontFamily: jostFont, fontSize: 18 } },
            ),
          ),
        ),
      ),
    );

    const emitted = emit(resolveAll(doc));
    const parts = unzipSync(emitted.pptx);
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);

    // For 18px font: hangingIndent = max(18 * 1.25, 20) = 22.5px, levelIndent = level * 27px.
    const expectedL0MarL = pxToEmu(22.5);
    const expectedL0Indent = -pxToEmu(22.5);
    const expectedL1MarL = pxToEmu(27 + 22.5);
    const expectedL1Indent = -pxToEmu(22.5);

    assert.match(
      slideXml,
      new RegExp(
        `<a:pPr[^>]*marL="${expectedL0MarL}"[^>]*indent="${expectedL0Indent}"[^>]*>[\\s\\S]*?<a:buChar char="•"\\/>`,
        "u",
      ),
    );
    assert.match(
      slideXml,
      new RegExp(
        `<a:pPr[^>]*marL="${expectedL1MarL}"[^>]*indent="${expectedL1Indent}"[^>]*>[\\s\\S]*?<a:buAutoNum`,
        "u",
      ),
    );
    assert.match(
      slideXml,
      new RegExp(
        `<a:pPr[^>]*marL="${expectedL0MarL}"[^>]*indent="${expectedL0Indent}"[^>]*>[\\s\\S]*?<a:buChar char="-"\\/>`,
        "u",
      ),
    );

    // Placeholder paragraph keeps <a:buChar> without materialized marL/indent.
    const phMatch = slideXml.match(
      /<p:ph type="body"[\s\S]*?<a:pPr([^>]*)><a:buChar/u,
    );
    assert.ok(phMatch !== null);
    assert.doesNotMatch(phMatch[1]!, /\b(?:marL|indent)=/u);

    const imported = importAll(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
  });
});
