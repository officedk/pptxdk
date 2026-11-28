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
import {
  A_NS,
  OPC_REL_NS,
  P_NS,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  createImporter,
} from "@pptxdk/import";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
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

  it("recovers multi-code-point and empty <a:buChar> values across Text, Shape, and Master list styles so resolve and codegen succeed", () => {
    const pptx = zipSync({
      "_rels/.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>`,
      ),
      "ppt/presentation.xml": strToU8(
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
      ),
      "ppt/_rels/presentation.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_SLIDE}" Target="slides/slide1.xml"/></Relationships>`,
      ),
      "ppt/theme/theme1.xml": strToU8(
        `<a:theme xmlns:a="${A_NS}" name="T"><a:themeElements><a:clrScheme name="C"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="654321"/></a:accent2><a:accent3><a:srgbClr val="112233"/></a:accent3><a:accent4><a:srgbClr val="445566"/></a:accent4><a:accent5><a:srgbClr val="778899"/></a:accent5><a:accent6><a:srgbClr val="AABBCC"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="F"><a:majorFont><a:latin typeface="Jost*"/></a:majorFont><a:minorFont><a:latin typeface="Jost*"/></a:minorFont></a:fontScheme><a:fmtScheme name="O"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
      ),
      "ppt/slideMasters/slideMaster1.xml": strToU8(
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="M1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483650" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr><a:buChar char="•&#xFE0F;"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles></p:sldMaster>`,
      ),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_THEME}" Target="../theme/theme1.xml"/></Relationships>`,
      ),
      "ppt/slideLayouts/slideLayout1.xml": strToU8(
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="L1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      ),
      "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_MASTER}" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
      ),
      "ppt/slides/slide1.xml": strToU8(
        `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld name="s1"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="T1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="95250"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buChar char="-&gt;"/></a:pPr><a:r><a:rPr sz="1600"/><a:t>Text bullet</a:t></a:r></a:p><a:p><a:pPr><a:buChar char=""/></a:pPr><a:r><a:rPr sz="1600"/><a:t>Empty bullet</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="S1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="95250" y="1143000"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buChar char="ab"/></a:pPr><a:r><a:rPr sz="1600"/><a:t>Shape bullet</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
      ),
      "ppt/slides/_rels/slide1.xml.rels": strToU8(
        `<Relationships xmlns="${OPC_REL_NS}"><Relationship Id="rId1" Type="${REL_SLIDE_LAYOUT}" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
      ),
    });

    const importer = createImporter(documentCapabilities);
    const imported = importer(pptx);
    const bulletDiags = imported.report.diagnostics.filter(
      (d) =>
        d.code === "invalid-xml-value" &&
        d.attributeOrElement === "a:buChar/@char",
    );
    assert.ok(bulletDiags.length >= 3);
    assert.ok(bulletDiags.every((d) => d.disposition === "recovered"));

    const codegen = createCodegen(documentCapabilities);
    const vanilla = codegen(imported.document, {
      target: "vanilla",
    });
    const react = codegen(imported.document, {
      target: "react",
    });
    assert.ok(
      Object.keys(vanilla.files).length > 0 &&
        Object.keys(react.files).length > 0,
    );
  });
});
