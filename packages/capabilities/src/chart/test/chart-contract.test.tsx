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
  observeMotion,
} from "@pptxdk/capabilities/motion";
import { BuildError, createMorphScope, targets } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
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
            x: 24,
            y: 20,
            width: 592,
            height: 320,
            title: "Quarterly sales",
            legend: "bottom",
            style: { fontFamily: jostFont, fontSize: 16 },
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

    const observed = resolution.report.slides[0]?.objects.find(
      (obj) => obj.authorId === "sales",
    );
    assert.ok(observed !== undefined);
    assert.equal(observed.kind, "chart");
    assert.equal(observed.name, "SalesChart");
    assert.deepEqual(observed.box, {
      x: 24,
      y: 20,
      width: 592,
      height: 320,
    });

    // Paint runs independently without calling emit.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 1);
    const page = painted.pages[0]!;
    assert.ok(page.png !== undefined);
    assert.ok(page.svg?.includes("<path"));
    assert.deepEqual(page.diagnostics, []);

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
                x: 10,
                y: 10,
                width: 300,
                height: 200,
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
                x: 10,
                y: 10,
                width: 300,
                height: 200,
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
                x: 10,
                y: 10,
                width: 300,
                height: 200,
                style: { fontSize: -12 },
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
                x: 10,
                y: 10,
                width: 300,
                height: 200,
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
                x: 10,
                y: 10,
                width: 300,
                height: 200,
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
      x: 20,
      y: 20,
      width: 400,
      height: 260,
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
      painted.pages[0]?.diagnostics.some(
        (diag) =>
          diag.severity === "error" &&
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
          x: 20,
          y: 20,
          width: 560,
          height: 300,
          title: "Revenue",
          legend: "bottom",
          style: { fontFamily: impJostFont, fontSize: 16 },
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
          x: 20,
          y: 20,
          width: 560,
          height: 300,
          title: "中文标题",
          style: { fontFamily: impJostFont, fontSize: 16 },
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
      missingResolution.report.diagnostics.some(
        (diag) =>
          diag.severity === "warning" && diag.code === "font-missing-glyph",
      ),
    );
    const missingPaint = paint(missingResolution);
    assert.equal(missingPaint.pages[0]?.png, undefined);
    assert.ok((missingPaint.pages[0]?.diagnostics.length ?? 0) > 0);
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
            x: 24,
            y: 20,
            width: 592,
            height: 320,
            style: { fontFamily: jostFont, fontSize: 14 },
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
            x: 24,
            y: 20,
            width: 592,
            height: 320,
            style: { fontFamily: jostFont, fontSize: 14 },
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
    const motion = observeMotion(resolution.report.slides[0]);
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
});
