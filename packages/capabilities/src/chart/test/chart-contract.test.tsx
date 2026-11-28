import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import { chartCapability } from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as impChart,
  deck as impDeck,
  font as impFont,
  slide as impSlide,
} from "@pptxdk/capabilities/document";
import { font } from "@pptxdk/capabilities/font";
import {
  fade,
  motionCapability,
  motionSchedule,
} from "@pptxdk/capabilities/motion";
import { BuildError, createMorphScope, targets } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { expandDocument } from "@pptxdk/react";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBook = readFileSync(
  new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
);
const jostFont = font("Jost*", [{ weight: 400, data: jostBook }]);
const impJostFont = impFont("Jost*", [{ weight: 400, data: jostBook }]);

describe("chartCapability resolved contract", () => {
  it("resolves chart model once, paints independently without emit, and omits preview plot layout from native chart XML", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
    ]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "overview", background: "#FFFFFF" },
          createElement(Chart, {
            id: "sales",
            name: "SalesChart",
            title: "Quarterly sales",
            legend: "bottom",
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 592,
              height: 320,
              fontFamily: jostFont,
              fontSize: 16,
            },
            groups: [
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [
                  { name: "Actual", values: [18, 32, 24] },
                  { name: "Plan", values: [22, 28, 30] },
                ],
                dataLabels: { value: true },
              },
            ],
          }),
        ),
      ),
    );

    const resolution = resolve(doc);

    const salesLocator = resolution.slide("overview").get("sales");
    assert.equal(salesLocator.kind, "chart");
    assert.deepEqual(salesLocator.box, {
      x: 24,
      y: 20,
      width: 592,
      height: 320,
    });
    assert.deepEqual(salesLocator.bounds, {
      x: 24,
      y: 20,
      width: 592,
      height: 320,
    });
    const observed = resolution.report.slides[0]?.objects.find(
      (obj) => obj.authorId === "sales",
    );
    assert.equal(observed?.name, "SalesChart");

    // Paint runs independently without calling emit.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    const page = painted.pages[0]!;
    assert.ok(page.png !== undefined);
    assert.ok(page.svg?.includes("<path"));
    assert.deepEqual(page.errors, []);

    // Native emit keeps automatic <c:layout/> without preview's computed plot area.
    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);
    const zip = unzipSync(emitted.pptx);
    const chartXml = strFromU8(zip["ppt/charts/chart1.xml"]!);
    assert.ok(chartXml.includes("<c:barChart>"));
    assert.ok(chartXml.includes("<c:layout/>"));
    assert.equal(chartXml.includes("<c:manualLayout>"), false);
    assert.ok(
      Object.keys(zip).some((entry) =>
        entry.startsWith("ppt/embeddings/Microsoft_Excel_Worksheet"),
      ),
    );
  });

  it("validates XML 1.0 characters, text styles, chart colors, and plotArea bounds during resolve before emit", () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
    ]);

    assert.throws(
      () =>
        resolve(
          impDeck({ width: 640, height: 360 }).append(
            impSlide({ id: "s1" }).append(
              impChart({
                id: "bad-name",
                name: "Bad\u0000Name",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                },
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-prop" &&
            path === "deck/slide[s1]/chart[bad-name]",
        ),
    );

    assert.throws(
      () =>
        resolve(
          impDeck({ width: 640, height: 360 }).append(
            impSlide({ id: "s1" }).append(
              impChart({
                id: "bad-category",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                },
                groups: [
                  {
                    kind: "column",
                    data: {
                      kind: "shared-categories",
                      categories: ["Bad\u0000Cat"],
                    },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-prop"),
    );

    assert.throws(
      () =>
        resolve(
          impDeck({ width: 640, height: 360 }).append(
            impSlide({ id: "s1" }).append(
              impChart({
                id: "bad-text-style",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                  fontSize: -12,
                },
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-style"),
    );

    assert.throws(
      () =>
        resolve(
          impDeck({ width: 640, height: 360 }).append(
            impSlide({ id: "s1" }).append(
              impChart({
                id: "bad-series-color",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                },
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1], color: "#ZZZZZZ" }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "invalid-style"),
    );

    assert.throws(
      () =>
        resolve(
          impDeck({ width: 640, height: 360 }).append(
            impSlide({ id: "s1" }).append(
              impChart({
                id: "bad-plot-area",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                },
                plotArea: { left: 100, top: 10, width: 250, height: 100 },
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code, path }) =>
            code === "invalid-prop" &&
            path === "deck/slide[s1]/chart[bad-plot-area]",
        ),
    );
  });

  it("reports c15 extension in resolve, enforces target only in emit, and isolates resolve closures from document mutations", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
    ]);
    const chartNode = impChart({
      id: "doughnut",
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 400,
        height: 260,
      },
      groups: [
        {
          kind: "doughnut",
          data: { kind: "independent-categories" },
          dataLabels: { category: true, percent: true },
          series: [
            {
              name: "Ring 1",
              categories: ["Alpha", "Beta"],
              values: [40, 60],
            },
            {
              name: "Ring 2",
              categories: ["Gamma", "Delta", "Epsilon"],
              values: [20, 30, 50],
            },
          ],
        },
      ],
    });
    const slideNode = impSlide({ id: "rings", background: "#FFFFFF" }).append(
      chartNode,
    );
    const doc = impDeck({ width: 640, height: 360 }).append(slideNode);

    const resolution = resolve(doc);
    assert.deepEqual(resolution.report.extensions, ["c15"]);

    // Mutate the editable document after resolve; resolution must remain isolated.
    chartNode.remove();
    assert.equal(slideNode.children.length, 0);

    // Paint reports unsupported preview diagnostic for doughnut without failing resolve or emit.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    assert.equal(painted.pages[0]?.png, undefined);
    assert.ok(
      painted.pages[0]?.errors.some(
        (diag) =>
          diag.code === "preview-missing-font" &&
          diag.path === "deck/slide[rings]/chart[doughnut]",
      ),
    );

    // Baseline target rejects c15 in emit while transitional target succeeds across repeated emits.
    assert.throws(
      () => emit(resolution, { target: targets.baselineOnly }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "target-violation"),
    );

    const emittedFirst = emit(resolution, { target: targets.transitional });
    const emittedSecond = emit(resolution, { target: targets.transitional });
    assert.deepEqual((await validatePptx(emittedFirst.pptx)).errors, []);
    assert.deepEqual((await validatePptx(emittedSecond.pptx)).errors, []);
    const chartXml = strFromU8(
      unzipSync(emittedFirst.pptx)["ppt/charts/chart1.xml"]!,
    );
    assert.ok(chartXml.includes("<c:doughnutChart>"));
    assert.ok(chartXml.includes("c15:datalabelsRange"));
  });

  it("defers font embedding and missing-glyph checks to emit using resolve-time resource claims", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
    ]);
    const validDoc = impDeck({ width: 640, height: 360 }).append(
      impSlide({ id: "fonts", background: "#FFFFFF" }).append(
        impChart({
          id: "sales",
          title: "Revenue",
          legend: "bottom",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 560,
            height: 300,
            fontFamily: impJostFont,
            fontSize: 16,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [{ name: "Actual", values: [10, 25] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
    );

    const validResolution = resolve(validDoc);
    const emittedPlain = emit(validResolution);
    assert.equal(
      Object.keys(unzipSync(emittedPlain.pptx)).some((entry) =>
        entry.startsWith("ppt/fonts/"),
      ),
      false,
    );

    const emittedSubset = emit(validResolution, { embedFonts: "subset" });
    assert.deepEqual((await validatePptx(emittedSubset.pptx)).errors, []);
    const subsetZip = unzipSync(emittedSubset.pptx);
    assert.ok(
      Object.keys(subsetZip).some(
        (entry) => entry.startsWith("ppt/fonts/") && entry.endsWith(".fntdata"),
      ),
    );

    const painted = paint(validResolution);
    assert.ok(painted.pages[0]?.png !== undefined);

    // Missing glyph warns in resolve, paints cleanly, and fails in emit when font embedding is enabled.
    const missingDoc = impDeck({ width: 640, height: 360 }).append(
      impSlide({ id: "missing", background: "#FFFFFF" }).append(
        impChart({
          id: "cjk-chart",
          title: "中文标题",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 560,
            height: 300,
            fontFamily: impJostFont,
            fontSize: 16,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Series", values: [12, 18] }],
            },
          ],
        }),
      ),
    );

    const missingResolution = resolve(missingDoc);
    assert.ok(
      missingResolution.report.fonts.some(
        (fontUsage) =>
          fontUsage.family === "Jost*" &&
          (fontUsage.missingGlyphs?.length ?? 0) > 0,
      ),
    );
    const missingPaint = paint(missingResolution);
    assert.equal(missingPaint.pages[0]?.png, undefined);
    assert.ok((missingPaint.pages[0]?.errors.length ?? 0) > 0);
    const missingPlainEmit = emit(missingResolution);
    assert.deepEqual((await validatePptx(missingPlainEmit.pptx)).errors, []);

    assert.throws(
      () => emit(missingResolution, { embedFonts: "subset" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );
    assert.throws(
      () => emit(missingResolution, { embedFonts: "full" }),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(({ code }) => code === "font-missing-glyph"),
    );
  });

  it("supports morphKey, click links, and motion annotations across resolve, emit, and paint", async () => {
    const scope = createMorphScope();
    const chartMorph = scope.handle("sales-chart");
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
      motionCapability,
    ]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "first", background: "#FFFFFF" },
          createElement(Chart, {
            id: "chart-1",
            morphKey: chartMorph,
            link: { id: "second" },
            enter: fade(),
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 592,
              height: 320,
              fontFamily: jostFont,
              fontSize: 14,
            },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "Sales", values: [10, 20] }],
              },
            ],
          }),
        ),
        createElement(
          Slide,
          { id: "second", background: "#FFFFFF" },
          createElement(Chart, {
            id: "chart-2",
            morphKey: chartMorph,
            style: {
              position: "absolute",
              left: 24,
              top: 20,
              width: 592,
              height: 320,
              fontFamily: jostFont,
              fontSize: 14,
            },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "Sales", values: [15, 25] }],
              },
            ],
          }),
        ),
      ),
    );

    const resolution = resolve(doc);
    const morphName = resolution.report.slides[0]?.objects[0]?.name;
    assert.ok(morphName?.startsWith("!!"));
    assert.equal(resolution.report.slides[1]?.objects[0]?.name, morphName);
    assert.equal(resolution.report.slides[0]?.objects[0]?.link, "slide:second");
    const motion = motionSchedule(resolution.slide("first"));
    assert.ok(motion !== undefined);
    assert.equal(motion.steps.length, 1);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 2);
    assert.ok(painted.pages[0]?.png !== undefined);
    assert.ok(painted.pages[1]?.png !== undefined);

    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);
    const slide1Xml = strFromU8(
      unzipSync(emitted.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slide1Xml.includes(`name="${morphName}"`));
    assert.ok(slide1Xml.includes("<a:hlinkClick"));
    assert.ok(slide1Xml.includes("<p:timing>"));
  });

  it("materializes default -90° rotation for both primary and secondary vertical axis titles across emit, paint, and import roundtrip", async () => {
    const resolve = createResolver([
      slideBackgroundCapability,
      chartCapability,
    ]);
    const importChart = createImporter([
      slideBackgroundCapability,
      chartCapability,
    ]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "dual-axis", background: "#FFFFFF" },
          createElement(Chart, {
            id: "dual",
            title: "Revenue & Margin",
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 600,
              height: 320,
              fontFamily: jostFont,
              fontSize: 12,
            },
            axes: {
              quarter: { kind: "category", title: "Quarter" },
              revenue: {
                kind: "value",
                position: "left",
                title: "Revenue ($M)",
                min: 0,
                max: 100,
              },
              margin: {
                kind: "value",
                position: "right",
                title: "Operating Margin (%)",
                min: 0,
                max: 1,
                numberFormat: "0%",
              },
            },
            groups: [
              {
                kind: "column",
                axes: { category: "quarter", value: "revenue" },
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [{ name: "Revenue", values: [30, 60, 90] }],
              },
              {
                kind: "line",
                axes: { category: "quarter", value: "margin" },
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [{ name: "Margin", values: [0.2, 0.4, 0.7] }],
              },
            ],
          }),
        ),
      ),
    );

    const resolution = resolve(doc);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    const svg = painted.pages[0]?.svg ?? "";
    const negRotMatches = svg.match(/<g transform="rotate\(-90\b/gu) ?? [];
    const posRotMatches = svg.match(/<g transform="rotate\(90\b/gu) ?? [];
    assert.equal(
      negRotMatches.length,
      2,
      "Both left and right vertical axis titles must render at -90 degrees",
    );
    assert.equal(
      posRotMatches.length,
      0,
      "Right secondary vertical axis title must not render at +90 degrees",
    );

    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx);
    assert.deepEqual(validation.errors, []);
    const chartXml = strFromU8(
      unzipSync(emitted.pptx)["ppt/charts/chart1.xml"]!,
    );
    const rotBodyPrMatches =
      chartXml.match(/<a:bodyPr\b[^>]*\brot="-5400000"[^>]*>/gu) ?? [];
    // Each vertical axis title emits <c:tx><c:rich><a:bodyPr ...> and <c:txPr><a:bodyPr ...> (2 per axis title = 4 total).
    assert.equal(
      rotBodyPrMatches.length,
      4,
      'Both vertical axis titles must materialize rot="-5400000" in OOXML',
    );

    const imported = importChart(emitted.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const importedSlide = imported.document.children.find(
      (child) => "kind" in child && child.kind === "slide",
    );
    const importedChart =
      importedSlide !== undefined && "children" in importedSlide
        ? importedSlide.children.find(
            (child) => "kind" in child && child.kind === "chart",
          )
        : undefined;
    assert.ok(importedChart !== undefined && "style" in importedChart);
    (importedChart as { style(patch: unknown): unknown }).style({
      fontFamily: jostFont,
    });
    const rtResolution = resolve(imported.document);
    const rtPainted = paint(rtResolution);
    assert.deepEqual(rtPainted.pages[0]?.errors, []);
    const rtSvg = rtPainted.pages[0]?.svg ?? "";
    assert.equal(
      (rtSvg.match(/<g transform="rotate\(-90\b/gu) ?? []).length,
      2,
    );
    assert.equal((rtSvg.match(/<g transform="rotate\(90\b/gu) ?? []).length, 0);
    const rtEmitted = emit(rtResolution);
    assert.deepEqual((await validatePptx(rtEmitted.pptx)).errors, []);
  });
});
