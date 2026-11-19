import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Chart } from "@react-pptx/capabilities/chart-elements";
import { chart, deck, font, slide } from "@react-pptx/capabilities/document";
import { Deck, Slide } from "@react-pptx/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { expect } from "vitest";

import type { CapabilityScenario } from "../acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const chartTheme = {
  colors: {
    accent1: "#2563EB",
    accent2: "#F59E0B",
    dk1: "#0F172A",
  },
};

export const chartPositiveScenario: CapabilityScenario = {
  id: "chart-clustered-column-preview",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering single-group finite non-negative clustered column charts with title/untitled top value labels, bottom legend, fractional scale, and native chart/workbook retention.",
  visualCheckpoints: [
    "Page 1 (clustered-column-overview): white #FFFFFF canvas; centered title 'Quarterly sales'; L-shaped dark #0F172A axes with Y-axis ticks 0..40 (step 10); Q1/Q2/Q3 clusters each have blue (#2563EB) Actual and amber (#F59E0B) Plan bars with unclipped value labels above every bar; bottom legend displays blue 'Actual' and amber 'Plan' swatches.",
    "Page 2 (untitled-fractional-columns): light slate #F8FAFC canvas; untitled single-series column chart with categories Alpha/Beta/Gamma and values [0, 0.25, 0.75]; Y-axis ticks 0..0.8; value labels '0', '0.25', '0.75' sit cleanly above baseline/bars without top clipping.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS on Windows and verify both slides render native editable clustered column charts (<c:barChart>, gapWidth=150) backed by embedded Excel workbooks.",
    "Open the embedded Excel data sheet on Slide 1, edit Q2 Actual from 32 to 45, confirm the chart bar and data label update, save as a new copy, and reopen to confirm data/chart persistence.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "clustered-column-overview", background: "#FFFFFF" },
        createElement(Chart, {
          id: "sales",
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
      createElement(
        Slide,
        { id: "untitled-fractional-columns", background: "#F8FAFC" },
        createElement(Chart, {
          id: "fractions",
          x: 24,
          y: 12,
          width: 592,
          height: 328,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 16 },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["Alpha", "Beta", "Gamma"],
              },
              series: [{ name: "Score", values: [0, 0.25, 0.75] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({ id: "clustered-column-overview", background: "#FFFFFF" }).append(
        chart({
          id: "sales",
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
      slide({
        id: "untitled-fractional-columns",
        background: "#F8FAFC",
      }).append(
        chart({
          id: "fractions",
          x: 24,
          y: 12,
          width: 592,
          height: 328,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 16 },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["Alpha", "Beta", "Gamma"],
              },
              series: [{ name: "Score", values: [0, 0.25, 0.75] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);

    const page1 = preview.pages[0]!;
    assert.ok(page1.png !== undefined);
    assert.equal(page1.diagnostics.length, 1);
    assert.equal(page1.diagnostics[0]?.severity, "warning");
    assert.equal(
      page1.diagnostics[0]?.path,
      "deck/slide[clustered-column-overview]/chart[sales]",
    );
    expect(page1.diagnostics[0]?.message).toContain("native");

    // Verify rendered PNG actually contains accent1 (#2563EB) and accent2 (#F59E0B) bar pixels.
    const png1 = PNG.sync.read(Buffer.from(page1.png));
    let blueCount = 0;
    let amberCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount++;
      if (r === 0xf5 && g === 0x9e && b === 0x0b) amberCount++;
    }
    assert.ok(blueCount > 500, `Expected blue bars, got ${blueCount} pixels`);
    assert.ok(
      amberCount > 500,
      `Expected amber bars, got ${amberCount} pixels`,
    );

    const page2 = preview.pages[1]!;
    assert.ok(page2.png !== undefined);
    assert.equal(page2.diagnostics.length, 1);
    assert.equal(page2.diagnostics[0]?.severity, "warning");
    assert.equal(
      page2.diagnostics[0]?.path,
      "deck/slide[untitled-fractional-columns]/chart[fractions]",
    );

    const files = unzipSync(pptx);
    const chart1Xml = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chart1Xml).toContain("<c:barChart>");
    expect(chart1Xml).toContain('<c:grouping val="clustered"/>');
    expect(chart1Xml).toContain("Quarterly sales");
    expect(chart1Xml).toContain('<c:showVal val="1"/>');
    expect(chart1Xml).toContain('<c:legendPos val="b"/>');
    expect(chart1Xml).not.toContain("<c:manualLayout>");

    const chart2Xml = strFromU8(files["ppt/charts/chart2.xml"]!);
    expect(chart2Xml).toContain("<c:barChart>");
    expect(chart2Xml).toContain("Alpha");
    expect(chart2Xml).toContain("0.75");
    expect(chart2Xml).not.toContain("<c:legend>");

    const workbooks = Object.keys(files).filter((name) =>
      name.startsWith("ppt/embeddings/Microsoft_Excel_Worksheet"),
    );
    assert.equal(workbooks.length, 2);
  },
};

export const chartRejectionScenario: CapabilityScenario = {
  id: "chart-preview-rejections",
  capability: "rp:chart",
  description:
    "Dual-author Chart rejection scenario verifying unsupported chart kinds, negative values, and warning-cannot-clear-error isolation while retaining native PPTX parts and unaffected pages.",
  visualCheckpoints: [
    "Page 1 (unsupported-line-with-column-warning): no PNG emitted; located error on chart[trend] and located warning on chart[bars].",
    "Page 2 (negative-column-values): no PNG emitted; located error on chart[neg].",
    "Page 3 (unaffected-column): white canvas with valid clustered column chart 'Healthy' ([12, 20]), blue bars (#2563EB), Y ticks, X labels (H1, H2), and layout warning.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify Slide 1's line & column charts, Slide 2's negative-value column chart, and Slide 3's column chart all remain native charts.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "unsupported-line-with-column-warning", background: "#FFFFFF" },
        createElement(Chart, {
          id: "trend",
          x: 20,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Trend", values: [10, 20] }],
            },
          ],
        }),
        createElement(Chart, {
          id: "bars",
          x: 330,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Bars", values: [8, 16] }],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "negative-column-values", background: "#FFFFFF" },
        createElement(Chart, {
          id: "neg",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Delta", values: [-4, 12] }],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "unaffected-column", background: "#FFFFFF" },
        createElement(Chart, {
          id: "healthy",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Healthy",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["H1", "H2"] },
              series: [{ name: "Revenue", values: [12, 20] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({
        id: "unsupported-line-with-column-warning",
        background: "#FFFFFF",
      }).append(
        chart({
          id: "trend",
          x: 20,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Trend", values: [10, 20] }],
            },
          ],
        }),
        chart({
          id: "bars",
          x: 330,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Bars", values: [8, 16] }],
            },
          ],
        }),
      ),
      slide({ id: "negative-column-values", background: "#FFFFFF" }).append(
        chart({
          id: "neg",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Delta", values: [-4, 12] }],
            },
          ],
        }),
      ),
      slide({ id: "unaffected-column", background: "#FFFFFF" }).append(
        chart({
          id: "healthy",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Healthy",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["H1", "H2"] },
              series: [{ name: "Revenue", values: [12, 20] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 3);

    const page1 = preview.pages[0]!;
    assert.equal(page1.png, undefined);
    assert.equal(page1.svg, undefined);
    expect(page1.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          severity: "error",
          path: "deck/slide[unsupported-line-with-column-warning]/chart[trend]",
        }),
        expect.objectContaining({
          severity: "warning",
          path: "deck/slide[unsupported-line-with-column-warning]/chart[bars]",
        }),
      ]),
    );

    const page2 = preview.pages[1]!;
    assert.equal(page2.png, undefined);
    assert.equal(page2.svg, undefined);
    expect(page2.diagnostics).toEqual([
      expect.objectContaining({
        severity: "error",
        path: "deck/slide[negative-column-values]/chart[neg]",
      }),
    ]);

    const page3 = preview.pages[2]!;
    assert.ok(page3.png !== undefined);
    expect(page3.diagnostics).toEqual([
      expect.objectContaining({
        severity: "warning",
        path: "deck/slide[unaffected-column]/chart[healthy]",
      }),
    ]);

    const files = unzipSync(pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).toContain(
      "<c:lineChart>",
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      "<c:barChart>",
    );
    expect(strFromU8(files["ppt/charts/chart4.xml"]!)).toContain(
      "<c:barChart>",
    );
  },
};
