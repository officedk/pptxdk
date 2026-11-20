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
    assert.deepEqual(page1.diagnostics, []);

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
    assert.deepEqual(page2.diagnostics, []);

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
  kind: "rejection",
  description:
    "Dual-author Chart rejection scenario verifying font-resource limitations (preview-missing-glyph and preview-unsupported-font) isolate failed pages while retaining native PPTX parts and unaffected sibling pages.",
  visualCheckpoints: [
    "Page 1 (missing-glyph-chart): no PNG emitted; located preview-missing-glyph error on chart[cjk-in-latin] while sibling chart[healthy-line] has no warning.",
    "Page 2 (missing-font-chart): no PNG emitted; located preview-unsupported-font error on chart[no-font-file].",
    "Page 3 (unaffected-pie): white canvas with valid pie chart 'Healthy Share' ([40, 60]), blue (#2563EB) and amber (#F59E0B) slices, bottom legend, and zero diagnostics.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify Slide 1, Slide 2, and Slide 3 charts all remain native charts.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "missing-glyph-chart", background: "#FFFFFF" },
        createElement(Chart, {
          id: "cjk-in-latin",
          x: 20,
          y: 20,
          width: 280,
          height: 220,
          title: "汉",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Share", values: [40, 60] }],
            },
          ],
        }),
        createElement(Chart, {
          id: "healthy-line",
          x: 330,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Trend", values: [8, 16] }],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "missing-font-chart", background: "#FFFFFF" },
        createElement(Chart, {
          id: "no-font-file",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Missing Font",
          legend: "bottom",
          style: { fontSize: 14 },
          groups: [
            {
              kind: "doughnut",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Share", values: [40, 60] }],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "unaffected-pie", background: "#FFFFFF" },
        createElement(Chart, {
          id: "healthy",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Healthy Share",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "pie",
              data: { kind: "shared-categories", categories: ["H1", "H2"] },
              series: [{ name: "Revenue", values: [40, 60] }],
              dataLabels: { percent: true, position: "center" },
            },
          ],
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({ id: "missing-glyph-chart", background: "#FFFFFF" }).append(
        chart({
          id: "cjk-in-latin",
          x: 20,
          y: 20,
          width: 280,
          height: 220,
          title: "汉",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Share", values: [40, 60] }],
            },
          ],
        }),
        chart({
          id: "healthy-line",
          x: 330,
          y: 20,
          width: 280,
          height: 220,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Trend", values: [8, 16] }],
            },
          ],
        }),
      ),
      slide({ id: "missing-font-chart", background: "#FFFFFF" }).append(
        chart({
          id: "no-font-file",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Missing Font",
          legend: "bottom",
          style: { fontSize: 14 },
          groups: [
            {
              kind: "doughnut",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Share", values: [40, 60] }],
            },
          ],
        }),
      ),
      slide({ id: "unaffected-pie", background: "#FFFFFF" }).append(
        chart({
          id: "healthy",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          title: "Healthy Share",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "pie",
              data: { kind: "shared-categories", categories: ["H1", "H2"] },
              series: [{ name: "Revenue", values: [40, 60] }],
              dataLabels: { percent: true, position: "center" },
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
    expect(page1.diagnostics).toEqual([
      expect.objectContaining({
        severity: "error",
        path: "deck/slide[missing-glyph-chart]/chart[cjk-in-latin]",
        message: expect.stringMatching(/glyph/i),
      }),
    ]);

    const page2 = preview.pages[1]!;
    assert.equal(page2.png, undefined);
    assert.equal(page2.svg, undefined);
    expect(page2.diagnostics).toEqual([
      expect.objectContaining({
        severity: "error",
        path: "deck/slide[missing-font-chart]/chart[no-font-file]",
        message: expect.stringMatching(/font file/i),
      }),
    ]);

    const page3 = preview.pages[2]!;
    assert.ok(page3.png !== undefined);
    assert.deepEqual(page3.diagnostics, []);

    const files = unzipSync(pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).toContain(
      "<c:barChart>",
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      "<c:lineChart>",
    );
    expect(strFromU8(files["ppt/charts/chart3.xml"]!)).toContain(
      "<c:doughnutChart>",
    );
    expect(strFromU8(files["ppt/charts/chart4.xml"]!)).toContain(
      "<c:pieChart>",
    );
  },
};

export const chartAxesAndCompactScenario: CapabilityScenario = {
  id: "chart-axes-scales-and-compact-layout",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering dual value axes (log10 & reversed percent), top/reversed category axes, major/minor gridlines, tick marks, number formats, axis titles, styled chart title, topRight legend, and compact frame / long label adaptation.",
  visualCheckpoints: [
    "Page 1 (axes-log-reversed-dual): white canvas; bold styled title 'Scale & Dual Axes'; topRight legend; left log10 axis (1..1,000) with dashed gridlines, minor gridlines, out ticks, and rotated 'Log Output' title; right reversed percent axis (0%..100%) with cross ticks and 'Reversed %' title.",
    "Page 2 (top-reversed-category-and-zero): top reversed category axis with dashDot axis line, currency formatted ticks ($0..$100), and zero-value column series rendering cleanly.",
    "Page 3 (compact-and-long-labels): tiny 80x50 chart frame and wide chart with long category labels both render cleanly inside their chart bounds with zero preview errors or synthetic ellipsis glyphs.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify logBase=10, orientation=maxMin, legendPos=tr, gridlines, tick marks, and axis titles preserve native XML fidelity.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "axes-log-reversed-dual", background: "#FFFFFF" },
        createElement(Chart, {
          id: "dual",
          x: 20,
          y: 16,
          width: 600,
          height: 328,
          title: {
            text: "Scale & Dual Axes",
            style: { fontSize: 18, fontWeight: "bold", color: "#0F172A" },
          },
          legend: "topRight",
          style: { fontFamily: jostFont, fontSize: 12 },
          axes: {
            cat: {
              kind: "category",
              position: "bottom",
              majorTickMark: "out",
              tickLabelPosition: "low",
              title: "Quarter",
            },
            logVal: {
              kind: "value",
              position: "left",
              logBase: 10,
              min: 1,
              max: 1000,
              numberFormat: "#,##0",
              gridlines: { color: "#CBD5E1", width: 1, dash: "dash" },
              minorGridlines: true,
              majorTickMark: "out",
              minorTickMark: "in",
              title: { text: "Log Output", style: { color: "#1E293B" } },
            },
            pctVal: {
              kind: "value",
              position: "right",
              min: 0,
              max: 1,
              majorUnit: 0.25,
              minorUnit: 0.05,
              numberFormat: "0%",
              majorTickMark: "cross",
              title: "Share %",
            },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "cat", value: "logVal" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                { name: "Volume", values: [10, 120, 850], color: "#2563EB" },
              ],
            },
            {
              kind: "column",
              axes: { category: "cat", value: "pctVal" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                { name: "Share", values: [0.25, 0.5, 0.75], color: "#F59E0B" },
              ],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "top-reversed-category-and-zero", background: "#FFFFFF" },
        createElement(Chart, {
          id: "topcat",
          x: 20,
          y: 20,
          width: 340,
          height: 310,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 12 },
          axes: {
            cat: {
              kind: "category",
              position: "top",
              reversed: true,
              line: { color: "#334155", width: 1.5, dash: "dashDot" },
              tickLabelPosition: "high",
            },
            val: {
              kind: "value",
              position: "left",
              reversed: true,
              min: 0,
              max: 100,
              majorUnit: 25,
              numberFormat: "$#,##0",
              crosses: "min",
            },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "cat", value: "val" },
              data: {
                kind: "shared-categories",
                categories: ["First", "Second", "Third"],
              },
              series: [{ name: "Spend", values: [25, 60, 90] }],
            },
          ],
        }),
        createElement(Chart, {
          id: "zeros",
          x: 380,
          y: 20,
          width: 240,
          height: 310,
          title: "Zero Scale",
          legend: false,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Empty", values: [0, 0] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "compact-and-long-labels", background: "#FFFFFF" },
        createElement(Chart, {
          id: "tiny",
          x: 20,
          y: 20,
          width: 80,
          height: 50,
          title: "Tiny",
          legend: "bottom",
          style: { fontFamily: jostFont },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["Alpha", "Beta"],
              },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
        createElement(Chart, {
          id: "longlabels",
          x: 120,
          y: 20,
          width: 500,
          height: 310,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 13 },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: [
                  "Extraordinarily Long Category Name Without Any Abbreviation",
                  "Another Very Detailed Segment Label",
                ],
              },
              series: [{ name: "Metric", values: [14, 28] }],
            },
          ],
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({ id: "axes-log-reversed-dual", background: "#FFFFFF" }).append(
        chart({
          id: "dual",
          x: 20,
          y: 16,
          width: 600,
          height: 328,
          title: {
            text: "Scale & Dual Axes",
            style: { fontSize: 18, fontWeight: "bold", color: "#0F172A" },
          },
          legend: "topRight",
          style: { fontFamily: jostFont, fontSize: 12 },
          axes: {
            cat: {
              kind: "category",
              position: "bottom",
              majorTickMark: "out",
              tickLabelPosition: "low",
              title: "Quarter",
            },
            logVal: {
              kind: "value",
              position: "left",
              logBase: 10,
              min: 1,
              max: 1000,
              numberFormat: "#,##0",
              gridlines: { color: "#CBD5E1", width: 1, dash: "dash" },
              minorGridlines: true,
              majorTickMark: "out",
              minorTickMark: "in",
              title: { text: "Log Output", style: { color: "#1E293B" } },
            },
            pctVal: {
              kind: "value",
              position: "right",
              min: 0,
              max: 1,
              majorUnit: 0.25,
              minorUnit: 0.05,
              numberFormat: "0%",
              majorTickMark: "cross",
              title: "Share %",
            },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "cat", value: "logVal" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                { name: "Volume", values: [10, 120, 850], color: "#2563EB" },
              ],
            },
            {
              kind: "column",
              axes: { category: "cat", value: "pctVal" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                { name: "Share", values: [0.25, 0.5, 0.75], color: "#F59E0B" },
              ],
            },
          ],
        }),
      ),
      slide({
        id: "top-reversed-category-and-zero",
        background: "#FFFFFF",
      }).append(
        chart({
          id: "topcat",
          x: 20,
          y: 20,
          width: 340,
          height: 310,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 12 },
          axes: {
            cat: {
              kind: "category",
              position: "top",
              reversed: true,
              line: { color: "#334155", width: 1.5, dash: "dashDot" },
              tickLabelPosition: "high",
            },
            val: {
              kind: "value",
              position: "left",
              reversed: true,
              min: 0,
              max: 100,
              majorUnit: 25,
              numberFormat: "$#,##0",
              crosses: "min",
            },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "cat", value: "val" },
              data: {
                kind: "shared-categories",
                categories: ["First", "Second", "Third"],
              },
              series: [{ name: "Spend", values: [25, 60, 90] }],
            },
          ],
        }),
        chart({
          id: "zeros",
          x: 380,
          y: 20,
          width: 240,
          height: 310,
          title: "Zero Scale",
          legend: false,
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Empty", values: [0, 0] }],
              dataLabels: { value: true },
            },
          ],
        }),
      ),
      slide({ id: "compact-and-long-labels", background: "#FFFFFF" }).append(
        chart({
          id: "tiny",
          x: 20,
          y: 20,
          width: 80,
          height: 50,
          title: "Tiny",
          legend: "bottom",
          style: { fontFamily: jostFont },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["Alpha", "Beta"],
              },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
        chart({
          id: "longlabels",
          x: 120,
          y: 20,
          width: 500,
          height: 310,
          legend: false,
          style: { fontFamily: jostFont, fontSize: 13 },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: [
                  "Extraordinarily Long Category Name Without Any Abbreviation",
                  "Another Very Detailed Segment Label",
                ],
              },
              series: [{ name: "Metric", values: [14, 28] }],
            },
          ],
        }),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 3);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    const files = unzipSync(pptx);
    const chart1Xml = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chart1Xml).toContain('<c:logBase val="10"/>');
    expect(chart1Xml).toContain('<c:legendPos val="tr"/>');
    expect(chart1Xml).toContain("<c:majorGridlines>");
    expect(chart1Xml).toContain("<c:minorGridlines/>");
    expect(chart1Xml).toContain('<c:majorTickMark val="out"/>');
    expect(chart1Xml).toContain('<c:minorTickMark val="in"/>');
    expect(chart1Xml).toContain('<c:majorTickMark val="cross"/>');
    expect(chart1Xml).toContain('formatCode="#,##0"');
    expect(chart1Xml).toContain('formatCode="0%"');
    expect(chart1Xml).not.toContain("<c:manualLayout>");

    const chart2Xml = strFromU8(files["ppt/charts/chart2.xml"]!);
    expect(chart2Xml).toContain('<c:orientation val="maxMin"/>');
    expect(chart2Xml).toContain('formatCode="$#,##0"');

    // Verify topRight legend renders in upper-right quadrant of Page 1.
    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let topRightLegendPixels = 0;
    for (let y = 40; y < 120; y += 1) {
      for (let x = 500; x < 620; x += 1) {
        const idx = (y * png1.width + x) * 4;
        if (
          png1.data[idx]! < 250 ||
          png1.data[idx + 1]! < 250 ||
          png1.data[idx + 2]! < 250
        ) {
          topRightLegendPixels += 1;
        }
      }
    }
    assert.ok(
      topRightLegendPixels > 40,
      `Expected topRight legend pixels, got ${topRightLegendPixels}`,
    );
  },
};

export const chartBarColumnVariantsScenario: CapabilityScenario = {
  id: "chart-bar-column-variants-and-legends",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering column and horizontal bar clustered, stacked, and percent variants with negative values, null slots, gapWidth/overlap, series/point fill & stroke overrides, and all legend positions.",
  visualCheckpoints: [
    "Page 1 (column-stacked-and-percent): top-legend stacked column with mixed positive/negative bars and dashed/dotted overrides alongside left-legend percent stacked column with null gap slot.",
    "Page 2 (horizontal-bar-variants): right-legend clustered horizontal bar with null slot preserving spacing, negative overlap, and point color override alongside bottom-legend stacked horizontal bar with negative & positive segments.",
    "Page 3 (horizontal-percent-bar-topRight-legend): percent-stacked horizontal bar with topRight legend and explicit plotArea.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify barDir='bar'/'col', grouping='clustered'/'stacked'/'percentStacked', gapWidth, overlap, and legend positions match.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "column-stacked-and-percent", background: "#FFFFFF" },
        createElement(Chart, {
          id: "col-stacked",
          x: 16,
          y: 16,
          width: 296,
          height: 328,
          title: "Stacked +/-",
          legend: "top",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              stacking: "stacked",
              gapWidth: 80,
              data: { kind: "shared-categories", categories: ["A", "B", "C"] },
              series: [
                {
                  name: "Gain",
                  color: "#10B981",
                  line: { color: "#0F172A", width: 1.5, dash: "dash" },
                  values: [20, -15, 30],
                  points: {
                    1: {
                      color: "#EF4444",
                      line: { color: "#7F1D1D", width: 2, dash: "dot" },
                    },
                  },
                },
                { name: "Extra", color: "#3B82F6", values: [15, -10, -12] },
              ],
            },
          ],
        }),
        createElement(Chart, {
          id: "col-percent",
          x: 328,
          y: 16,
          width: 296,
          height: 328,
          title: "Percent",
          legend: "left",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              stacking: "percent",
              data: { kind: "shared-categories", categories: ["X", "Y", "Z"] },
              series: [
                { name: "P1", values: [10, null, 30] },
                { name: "P2", values: [30, 20, 10] },
              ],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "horizontal-bar-variants", background: "#FFFFFF" },
        createElement(Chart, {
          id: "bar-clustered",
          x: 16,
          y: 16,
          width: 296,
          height: 328,
          title: "Bar Clustered",
          legend: "right",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              gapWidth: 100,
              overlap: -20,
              data: {
                kind: "shared-categories",
                categories: ["North", "South", "West"],
              },
              series: [
                {
                  name: "S1",
                  values: [12, null, 28],
                  points: { 2: { color: "#EC4899" } },
                },
                { name: "S2", values: [18, 22, 16] },
              ],
              dataLabels: { value: true },
            },
          ],
        }),
        createElement(Chart, {
          id: "bar-stacked",
          x: 328,
          y: 16,
          width: 296,
          height: 328,
          title: "Bar Stacked",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              stacking: "stacked",
              data: {
                kind: "shared-categories",
                categories: ["Jan", "Feb", "Mar"],
              },
              series: [
                { name: "In", values: [25, -10, 15] },
                { name: "Out", values: [15, -20, 25] },
              ],
            },
          ],
        }),
      ),
      createElement(
        Slide,
        { id: "horizontal-percent-bar-topRight-legend", background: "#FFFFFF" },
        createElement(Chart, {
          id: "bar-percent",
          x: 24,
          y: 20,
          width: 592,
          height: 320,
          title: "Bar Percent",
          legend: "topRight",
          plotArea: { left: 72, top: 44, width: 410, height: 236 },
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              stacking: "percent",
              data: {
                kind: "shared-categories",
                categories: ["Team A", "Team B"],
              },
              series: [
                { name: "Done", values: [3, 1] },
                { name: "Todo", values: [1, 3] },
              ],
            },
          ],
        }),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({ id: "column-stacked-and-percent", background: "#FFFFFF" }).append(
        chart({
          id: "col-stacked",
          x: 16,
          y: 16,
          width: 296,
          height: 328,
          title: "Stacked +/-",
          legend: "top",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              stacking: "stacked",
              gapWidth: 80,
              data: { kind: "shared-categories", categories: ["A", "B", "C"] },
              series: [
                {
                  name: "Gain",
                  color: "#10B981",
                  line: { color: "#0F172A", width: 1.5, dash: "dash" },
                  values: [20, -15, 30],
                  points: {
                    1: {
                      color: "#EF4444",
                      line: { color: "#7F1D1D", width: 2, dash: "dot" },
                    },
                  },
                },
                { name: "Extra", color: "#3B82F6", values: [15, -10, -12] },
              ],
            },
          ],
        }),
        chart({
          id: "col-percent",
          x: 328,
          y: 16,
          width: 296,
          height: 328,
          title: "Percent",
          legend: "left",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "column",
              stacking: "percent",
              data: { kind: "shared-categories", categories: ["X", "Y", "Z"] },
              series: [
                { name: "P1", values: [10, null, 30] },
                { name: "P2", values: [30, 20, 10] },
              ],
            },
          ],
        }),
      ),
      slide({ id: "horizontal-bar-variants", background: "#FFFFFF" }).append(
        chart({
          id: "bar-clustered",
          x: 16,
          y: 16,
          width: 296,
          height: 328,
          title: "Bar Clustered",
          legend: "right",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              gapWidth: 100,
              overlap: -20,
              data: {
                kind: "shared-categories",
                categories: ["North", "South", "West"],
              },
              series: [
                {
                  name: "S1",
                  values: [12, null, 28],
                  points: { 2: { color: "#EC4899" } },
                },
                { name: "S2", values: [18, 22, 16] },
              ],
              dataLabels: { value: true },
            },
          ],
        }),
        chart({
          id: "bar-stacked",
          x: 328,
          y: 16,
          width: 296,
          height: 328,
          title: "Bar Stacked",
          legend: "bottom",
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              stacking: "stacked",
              data: {
                kind: "shared-categories",
                categories: ["Jan", "Feb", "Mar"],
              },
              series: [
                { name: "In", values: [25, -10, 15] },
                { name: "Out", values: [15, -20, 25] },
              ],
            },
          ],
        }),
      ),
      slide({
        id: "horizontal-percent-bar-topRight-legend",
        background: "#FFFFFF",
      }).append(
        chart({
          id: "bar-percent",
          x: 24,
          y: 20,
          width: 592,
          height: 320,
          title: "Bar Percent",
          legend: "topRight",
          plotArea: { left: 72, top: 44, width: 410, height: 236 },
          style: { fontFamily: jostFont, fontSize: 12 },
          groups: [
            {
              kind: "bar",
              stacking: "percent",
              data: {
                kind: "shared-categories",
                categories: ["Team A", "Team B"],
              },
              series: [
                { name: "Done", values: [3, 1] },
                { name: "Todo", values: [1, 3] },
              ],
            },
          ],
        }),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 3);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    // Check Page 1 contains green (#10B981) and red override (#EF4444) pixels.
    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let greenPixels = 0;
    let redPixels = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x10 && g === 0xb9 && b === 0x81) greenPixels += 1;
      if (r === 0xef && g === 0x44 && b === 0x44) redPixels += 1;
    }
    assert.ok(greenPixels > 200, `Expected #10B981 pixels, got ${greenPixels}`);
    assert.ok(
      redPixels > 100,
      `Expected #EF4444 override pixels, got ${redPixels}`,
    );

    // Check Page 2 contains pink point override (#EC4899) pixels.
    const png2 = PNG.sync.read(Buffer.from(preview.pages[1]!.png!));
    let pinkPixels = 0;
    for (let i = 0; i < png2.data.length; i += 4) {
      if (
        png2.data[i] === 0xec &&
        png2.data[i + 1] === 0x48 &&
        png2.data[i + 2] === 0x99
      ) {
        pinkPixels += 1;
      }
    }
    assert.ok(
      pinkPixels > 100,
      `Expected #EC4899 point override pixels, got ${pinkPixels}`,
    );

    const files = unzipSync(pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).toContain(
      '<c:grouping val="stacked"/>',
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      '<c:grouping val="percentStacked"/>',
    );
    expect(strFromU8(files["ppt/charts/chart3.xml"]!)).toContain(
      '<c:barDir val="bar"/>',
    );
    expect(strFromU8(files["ppt/charts/chart5.xml"]!)).toContain(
      '<c:legendPos val="tr"/>',
    );
    expect(strFromU8(files["ppt/charts/chart5.xml"]!)).toContain(
      "<c:manualLayout>",
    );
  },
};

const allMarkerSymbols = [
  "none",
  "auto",
  "circle",
  "square",
  "diamond",
  "triangle",
  "plus",
  "x",
  "star",
  "dot",
  "dash",
] as const;

const lineMarkersPoints = Object.fromEntries(
  allMarkerSymbols.map((symbol, idx) => [
    idx,
    {
      marker: {
        symbol,
        size: 7,
        fill: idx % 2 === 0 ? "#2563EB" : "#EC4899",
        line: { color: "#0F172A", width: 1 },
      },
      ...(idx === 5
        ? {
            color: "#EF4444",
            line: { color: "#EF4444", width: 3, dash: "dot" as const },
          }
        : {}),
    },
  ]),
);

const lineAreaPage1Charts = [
  {
    id: "line-all-markers-gap",
    x: 16,
    y: 16,
    width: 390,
    height: 328,
    title: "Smooth Line & All Markers",
    legend: "bottom" as const,
    displayBlanksAs: "gap" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "line" as const,
        data: {
          kind: "shared-categories" as const,
          categories: [
            "M1",
            "M2",
            "M3",
            "M4",
            "M5",
            "M6",
            "M7",
            "M8",
            "M9",
            "M10",
            "M11",
          ],
        },
        series: [
          {
            name: "Symbols",
            color: "#2563EB",
            smooth: true,
            line: { width: 2, dash: "solid" as const },
            marker: { symbol: "circle" as const, size: 6 },
            values: [12, 24, 18, null, 26, 34, 22, 30, 16, 28, 36],
            points: lineMarkersPoints,
          },
          {
            name: "Auto",
            color: "#10B981",
            line: { width: 2, dash: "dash" as const },
            marker: { symbol: "auto" as const, size: 6 },
            values: [8, 14, 22, 18, 15, 20, 28, 24, 20, 14, 18],
          },
        ],
        dataLabels: { value: true, position: "top" as const },
      },
    ],
  },
  {
    id: "line-span",
    x: 416,
    y: 16,
    width: 208,
    height: 158,
    title: "Span Blanks",
    legend: "bottom" as const,
    displayBlanksAs: "span" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "line" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B", "C", "D"],
        },
        series: [
          {
            name: "Span",
            color: "#F59E0B",
            marker: true,
            values: [10, null, 28, 22],
          },
        ],
      },
    ],
  },
  {
    id: "line-zero",
    x: 416,
    y: 186,
    width: 208,
    height: 158,
    title: "Zero Blanks",
    legend: "bottom" as const,
    displayBlanksAs: "zero" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "line" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B", "C", "D"],
        },
        series: [
          {
            name: "Zero",
            color: "#8B5CF6",
            marker: true,
            values: [15, null, 30, 20],
          },
        ],
      },
    ],
  },
] as const;

const lineAreaPage2Charts = [
  {
    id: "line-stacked",
    x: 16,
    y: 16,
    width: 296,
    height: 158,
    title: "Line Stacked",
    legend: "top" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "line" as const,
        stacking: "stacked" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3"],
        },
        series: [
          { name: "S1", color: "#2563EB", marker: true, values: [10, 20, 15] },
          {
            name: "S2",
            color: "#F59E0B",
            marker: { symbol: "square" as const },
            values: [15, null, 25],
          },
        ],
      },
    ],
  },
  {
    id: "line-percent",
    x: 328,
    y: 16,
    width: 296,
    height: 158,
    title: "Line Percent",
    legend: "top" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "line" as const,
        stacking: "percent" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3"],
        },
        series: [
          { name: "P1", color: "#10B981", marker: true, values: [10, 30, 20] },
          { name: "P2", color: "#EC4899", marker: true, values: [30, 10, 20] },
        ],
      },
    ],
  },
  {
    id: "area-standard-midcat",
    x: 16,
    y: 186,
    width: 296,
    height: 158,
    title: "Area Standard",
    legend: "bottom" as const,
    displayBlanksAs: "gap" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "area" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Jan", "Feb", "Mar", "Apr", "May"],
        },
        series: [
          {
            name: "Flow",
            color: "#2563EB",
            line: { color: "#1E3A8A", width: 1.5 },
            values: [14, 28, null, 18, 32],
          },
        ],
        dataLabels: { value: true },
      },
    ],
  },
  {
    id: "area-percent",
    x: 328,
    y: 186,
    width: 296,
    height: 158,
    title: "Area Percent",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "area" as const,
        stacking: "percent" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Jan", "Feb", "Mar"],
        },
        series: [
          { name: "A", color: "#10B981", values: [20, 40, 30] },
          { name: "B", color: "#F59E0B", values: [60, 40, 70] },
        ],
      },
    ],
  },
] as const;

export const chartLineAreaParityScenario: CapabilityScenario = {
  id: "chart-line-area-parity",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering line and area standard/stacked/percent variants, smooth splines, all 11 marker symbols, gap/span/zero blank modes, midCat edge alignment, and series/point overrides.",
  visualCheckpoints: [
    "Page 1 (line-variants-markers-and-blanks): smooth blue line with gap break at M4, all 11 marker symbols, dashed green Auto line, plus Span and Zero blank mini charts.",
    "Page 2 (line-area-stacked-percent-and-midCat): stacked and percent lines on top; edge-to-edge midCat standard area with gap and percent-stacked area on bottom.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify line/area groupings, dispBlanksAs, smooth flags, marker symbols, and midCat crossBetween.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "line-variants-markers-and-blanks", background: "#FFFFFF" },
        ...lineAreaPage1Charts.map((props) => createElement(Chart, props)),
      ),
      createElement(
        Slide,
        { id: "line-area-stacked-percent-and-midCat", background: "#FFFFFF" },
        ...lineAreaPage2Charts.map((props) => createElement(Chart, props)),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({
        id: "line-variants-markers-and-blanks",
        background: "#FFFFFF",
      }).append(...lineAreaPage1Charts.map((props) => chart(props))),
      slide({
        id: "line-area-stacked-percent-and-midCat",
        background: "#FFFFFF",
      }).append(...lineAreaPage2Charts.map((props) => chart(props))),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let blueCount = 0;
    let pinkCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount += 1;
      if (r === 0xec && g === 0x48 && b === 0x99) pinkCount += 1;
    }
    assert.ok(
      blueCount > 100,
      `Expected #2563EB line/marker pixels, got ${blueCount}`,
    );
    assert.ok(
      pinkCount > 30,
      `Expected #EC4899 marker pixels, got ${pinkCount}`,
    );

    const png2 = PNG.sync.read(Buffer.from(preview.pages[1]!.png!));
    let greenCount = 0;
    for (let i = 0; i < png2.data.length; i += 4) {
      if (
        png2.data[i] === 0x10 &&
        png2.data[i + 1] === 0xb9 &&
        png2.data[i + 2] === 0x81
      ) {
        greenCount += 1;
      }
    }
    assert.ok(
      greenCount > 500,
      `Expected #10B981 area pixels, got ${greenCount}`,
    );

    const files = unzipSync(pptx);
    const chart1 = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chart1).toContain("<c:lineChart>");
    expect(chart1).toContain('<c:smooth val="1"/>');
    expect(chart1).toContain('<c:symbol val="diamond"/>');
    expect(chart1).toContain('<c:symbol val="star"/>');
    expect(chart1).toContain('<c:symbol val="dash"/>');
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      '<c:dispBlanksAs val="span"/>',
    );
    expect(strFromU8(files["ppt/charts/chart3.xml"]!)).toContain(
      '<c:dispBlanksAs val="zero"/>',
    );
    const chart6 = strFromU8(files["ppt/charts/chart6.xml"]!);
    expect(chart6).toContain("<c:areaChart>");
    expect(chart6).toContain('<c:crossBetween val="midCat"/>');
    expect(strFromU8(files["ppt/charts/chart7.xml"]!)).toContain(
      '<c:grouping val="percentStacked"/>',
    );
  },
};

const scatterPage1Charts = [
  {
    id: "scatter-smooth-marker-unequal",
    x: 16,
    y: 16,
    width: 300,
    height: 158,
    title: "Smooth + Markers (Unequal XY)",
    legend: "bottom" as const,
    displayBlanksAs: "span" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: "smooth" as const,
        markers: true,
        series: [
          {
            name: "Obs5",
            color: "#2563EB",
            marker: { symbol: "circle" as const, size: 6 },
            x: [1, 2, null, 4, 6],
            y: [4, 12, 8, 9, 18],
            points: {
              3: {
                color: "#EF4444",
                marker: {
                  symbol: "diamond" as const,
                  size: 8,
                  fill: "#EF4444",
                },
              },
            },
          },
          {
            name: "Obs3",
            color: "#10B981",
            line: { dash: "dash" as const, width: 2 },
            marker: { symbol: "triangle" as const, size: 6 },
            x: [1.5, 3.5, 5.5],
            y: [14, 6, 15],
          },
          {
            name: "Empty",
            x: [],
            y: [],
          },
        ],
        dataLabels: { xValue: true, value: true, position: "right" as const },
      },
    ],
  },
  {
    id: "scatter-markers-only",
    x: 324,
    y: 16,
    width: 300,
    height: 158,
    title: "Markers Only (Zero Blanks)",
    legend: "right" as const,
    displayBlanksAs: "zero" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: false as const,
        markers: true,
        series: [
          {
            name: "Cluster",
            color: "#F59E0B",
            marker: { symbol: "square" as const, size: 7 },
            x: [1, null, 3, 4],
            y: [5, 10, 15, 12],
          },
        ],
      },
    ],
  },
  {
    id: "scatter-straight-only",
    x: 16,
    y: 186,
    width: 196,
    height: 158,
    title: "Straight Only",
    legend: false as const,
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: "straight" as const,
        markers: false,
        series: [{ name: "S", color: "#8B5CF6", x: [0, 2, 4], y: [2, 8, 5] }],
      },
    ],
  },
  {
    id: "scatter-straight-markers",
    x: 222,
    y: 186,
    width: 196,
    height: 158,
    title: "Straight + Markers",
    legend: false as const,
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: "straight" as const,
        markers: true,
        series: [
          {
            name: "SM",
            color: "#EC4899",
            marker: { symbol: "x" as const, size: 7 },
            x: [1, 3, 5],
            y: [3, 9, 4],
          },
        ],
      },
    ],
  },
  {
    id: "scatter-smooth-only",
    x: 428,
    y: 186,
    width: 196,
    height: 158,
    title: "Smooth Only",
    legend: false as const,
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: "smooth" as const,
        markers: false,
        series: [
          { name: "SO", color: "#14B8A6", x: [1, 2, 3, 4], y: [2, 10, 4, 12] },
        ],
      },
    ],
  },
] as const;

const scatterPage2Charts = [
  {
    id: "scatter-quadrants",
    x: 16,
    y: 20,
    width: 296,
    height: 320,
    title: "4-Quadrant Scatter",
    legend: "bottom" as const,
    displayBlanksAs: "gap" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: "straight" as const,
        markers: true,
        series: [
          {
            name: "Quad",
            color: "#2563EB",
            x: [-10, -4, null, 4, 10],
            y: [-8, 6, 0, -5, 9],
          },
        ],
      },
    ],
  },
  {
    id: "scatter-log-axes",
    x: 328,
    y: 20,
    width: 296,
    height: 320,
    title: "Log10 XY Scatter",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    axes: {
      x: {
        kind: "value" as const,
        logBase: 10,
        min: 1,
        max: 1000,
        gridlines: true,
      },
      y: {
        kind: "value" as const,
        logBase: 10,
        min: 1,
        max: 100,
        gridlines: true,
      },
    },
    groups: [
      {
        kind: "scatter" as const,
        axes: { x: "x", y: "y" },
        data: { kind: "xy" as const },
        line: "straight" as const,
        markers: true,
        series: [
          {
            name: "PowerLaw",
            color: "#F59E0B",
            marker: { symbol: "star" as const, size: 7 },
            x: [1, 10, 100, 1000],
            y: [2, 8, 32, 90],
          },
        ],
      },
    ],
  },
] as const;

export const chartScatterParityScenario: CapabilityScenario = {
  id: "chart-scatter-parity",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering all 5 scatter marker/line presentations, unequal-length XY series with empty series and null coordinates, gap/span/zero blank handling, 4-quadrant and log10 axes, and X/Y data labels.",
  visualCheckpoints: [
    "Page 1 (scatter-five-modes-and-unequal-xy): 5 scatter charts demonstrating smooth+marker (with empty legend entry and X/Y labels), markers-only, straight-only, straight+markers, and smooth-only.",
    "Page 2 (scatter-log-and-quadrants): 4-quadrant scatter with negative/positive X/Y zero baselines and log10 XY scatter with power-law curve.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify all 5 scatterStyle modes, independent XY columns, and logBase=10 scatter axes.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "scatter-five-modes-and-unequal-xy", background: "#FFFFFF" },
        ...scatterPage1Charts.map((props) => createElement(Chart, props)),
      ),
      createElement(
        Slide,
        { id: "scatter-log-and-quadrants", background: "#FFFFFF" },
        ...scatterPage2Charts.map((props) => createElement(Chart, props)),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({
        id: "scatter-five-modes-and-unequal-xy",
        background: "#FFFFFF",
      }).append(...scatterPage1Charts.map((props) => chart(props))),
      slide({ id: "scatter-log-and-quadrants", background: "#FFFFFF" }).append(
        ...scatterPage2Charts.map((props) => chart(props)),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let blueCount = 0;
    let amberCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount += 1;
      if (r === 0xf5 && g === 0x9e && b === 0x0b) amberCount += 1;
    }
    assert.ok(
      blueCount > 80,
      `Expected #2563EB scatter pixels, got ${blueCount}`,
    );
    assert.ok(
      amberCount > 80,
      `Expected #F59E0B marker pixels, got ${amberCount}`,
    );

    const files = unzipSync(pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).toContain(
      '<c:scatterStyle val="smoothMarker"/>',
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      '<c:scatterStyle val="marker"/>',
    );
    expect(strFromU8(files["ppt/charts/chart3.xml"]!)).toContain(
      '<c:scatterStyle val="line"/>',
    );
    expect(strFromU8(files["ppt/charts/chart4.xml"]!)).toContain(
      '<c:scatterStyle val="lineMarker"/>',
    );
    expect(strFromU8(files["ppt/charts/chart5.xml"]!)).toContain(
      '<c:scatterStyle val="smooth"/>',
    );
    expect(strFromU8(files["ppt/charts/chart7.xml"]!)).toContain(
      '<c:logBase val="10"/>',
    );
  },
};

const comboPage1Chart = {
  id: "combo-col-area-line-dual",
  x: 20,
  y: 16,
  width: 600,
  height: 328,
  title: "Revenue, Baseline & Margin %",
  legend: "bottom" as const,
  style: { fontFamily: jostFont, fontSize: 12 },
  axes: {
    catPrimary: {
      kind: "category" as const,
      position: "bottom" as const,
      title: "Quarter",
    },
    catSecondary: { kind: "category" as const, position: "top" as const },
    valPrimary: {
      kind: "value" as const,
      position: "left" as const,
      min: 0,
      max: 100,
      numberFormat: "$#,##0",
      gridlines: true,
      title: "Revenue",
    },
    valSecondary: {
      kind: "value" as const,
      position: "right" as const,
      min: 0,
      max: 1,
      numberFormat: "0%",
      title: "Margin %",
    },
  },
  groups: [
    {
      kind: "area" as const,
      axes: { category: "catPrimary", value: "valPrimary" },
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1", "Q2", "Q3", "Q4"],
      },
      series: [{ name: "Floor", color: "#93C5FD", values: [20, 25, 30, 28] }],
    },
    {
      kind: "column" as const,
      axes: { category: "catPrimary", value: "valPrimary" },
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1", "Q2", "Q3", "Q4"],
      },
      series: [{ name: "Revenue", color: "#2563EB", values: [45, 68, 74, 88] }],
      dataLabels: { value: true, position: "insideEnd" as const },
    },
    {
      kind: "line" as const,
      axes: { category: "catSecondary", value: "valSecondary" },
      data: {
        kind: "shared-categories" as const,
        categories: ["P1", "P2", "P3", "P4"],
      },
      series: [
        {
          name: "Margin",
          color: "#F59E0B",
          smooth: true,
          marker: { symbol: "diamond" as const, size: 7, fill: "#F59E0B" },
          values: [0.22, 0.48, 0.62, 0.78],
        },
      ],
      dataLabels: { value: true, position: "top" as const },
    },
  ],
};

const comboPage2Chart = {
  id: "combo-bar-scatter-shared-x",
  x: 20,
  y: 16,
  width: 600,
  height: 328,
  title: "Horizontal Bar + Scatter Shared X",
  legend: "right" as const,
  style: { fontFamily: jostFont, fontSize: 12 },
  axes: {
    barCat: { kind: "category" as const, position: "left" as const },
    sharedVal: {
      kind: "value" as const,
      position: "bottom" as const,
      min: 0,
      max: 50,
      gridlines: true,
    },
    scatterY: {
      kind: "value" as const,
      position: "right" as const,
      min: 0,
      max: 10,
    },
  },
  groups: [
    {
      kind: "bar" as const,
      axes: { category: "barCat", value: "sharedVal" },
      data: {
        kind: "shared-categories" as const,
        categories: ["North", "South", "West"],
      },
      series: [{ name: "Volume", color: "#10B981", values: [18, 36, 28] }],
    },
    {
      kind: "scatter" as const,
      axes: { x: "sharedVal", y: "scatterY" },
      data: { kind: "xy" as const },
      line: "straight" as const,
      markers: true,
      series: [
        {
          name: "Target",
          color: "#EF4444",
          marker: { symbol: "circle" as const, size: 7 },
          x: [12, 26, 42],
          y: [3, 7, 9],
        },
      ],
    },
  ],
};

export const chartComboDualAxesScenario: CapabilityScenario = {
  id: "chart-combo-dual-axes-parity",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering multi-group combo layering (area + column + line across primary/secondary category & value axes, and horizontal bar + scatter sharing a value axis) with mixed legend icons and data labels.",
  visualCheckpoints: [
    "Page 1 (combo-area-column-line-dual): layered light-blue area, blue columns with insideEnd labels, amber smooth line with diamonds mapped to right 0%..100% axis, top secondary category axis, and bottom legend showing both swatches and line+marker icon.",
    "Page 2 (combo-bar-scatter-shared-x): green horizontal bars and red straight+circle scatter overlay sharing the bottom 0..50 axis with right scatter Y axis and right legend.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify all 4 axes, multi-group axId bindings, and mixed legend entries render cleanly.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "combo-area-column-line-dual", background: "#FFFFFF" },
        createElement(Chart, comboPage1Chart),
      ),
      createElement(
        Slide,
        { id: "combo-bar-scatter-shared-x", background: "#FFFFFF" },
        createElement(Chart, comboPage2Chart),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({
        id: "combo-area-column-line-dual",
        background: "#FFFFFF",
      }).append(chart(comboPage1Chart)),
      slide({ id: "combo-bar-scatter-shared-x", background: "#FFFFFF" }).append(
        chart(comboPage2Chart),
      ),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let blueCount = 0;
    let amberCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount += 1;
      if (r === 0xf5 && g === 0x9e && b === 0x0b) amberCount += 1;
    }
    assert.ok(
      blueCount > 300,
      `Expected #2563EB column pixels, got ${blueCount}`,
    );
    assert.ok(
      amberCount > 80,
      `Expected #F59E0B line/marker pixels, got ${amberCount}`,
    );

    const png2 = PNG.sync.read(Buffer.from(preview.pages[1]!.png!));
    let greenCount = 0;
    let redCount = 0;
    for (let i = 0; i < png2.data.length; i += 4) {
      const r = png2.data[i]!;
      const g = png2.data[i + 1]!;
      const b = png2.data[i + 2]!;
      if (r === 0x10 && g === 0xb9 && b === 0x81) greenCount += 1;
      if (r === 0xef && g === 0x44 && b === 0x44) redCount += 1;
    }
    assert.ok(
      greenCount > 300,
      `Expected #10B981 bar pixels, got ${greenCount}`,
    );
    assert.ok(
      redCount > 50,
      `Expected #EF4444 scatter pixels, got ${redCount}`,
    );

    const files = unzipSync(pptx);
    const chart1 = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chart1).toContain("<c:areaChart>");
    expect(chart1).toContain("<c:barChart>");
    expect(chart1).toContain("<c:lineChart>");
    const chart2 = strFromU8(files["ppt/charts/chart2.xml"]!);
    expect(chart2).toContain('<c:barDir val="bar"/>');
    expect(chart2).toContain("<c:scatterChart>");
  },
};

const pieDoughnutPage1Charts = [
  {
    id: "pie-full-circle-360",
    x: 16,
    y: 16,
    width: 220,
    height: 158,
    title: "360 Full Pie",
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "pie" as const,
        firstSliceAngle: 45,
        data: { kind: "shared-categories" as const, categories: ["Sole"] },
        series: [
          {
            name: "Total",
            values: [100],
            points: { 0: { color: "#2563EB" } },
          },
        ],
        dataLabels: {
          seriesName: true,
          percent: true,
          separator: " - ",
          position: "center" as const,
          style: { color: "#FFFFFF", fontSize: 10 },
        },
      },
    ],
  },
  {
    id: "pie-exploded-leader-and-3level-labels",
    x: 248,
    y: 16,
    width: 376,
    height: 328,
    title: "Exploded Pie & 3-Level Labels",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "pie" as const,
        firstSliceAngle: 90,
        explosion: 6,
        data: {
          kind: "shared-categories" as const,
          categories: ["Cloud", "License", "Services", "Hardware"],
        },
        series: [
          {
            name: "Mix",
            values: [45, 25, 20, 10],
            points: {
              0: {
                color: "#10B981",
                explosion: 18,
                labelColor: "#065F46",
                dataLabels: { position: "insideEnd" as const, percent: true },
              },
              1: {
                color: "#F59E0B",
                line: { color: "#1E293B", width: 1.5 },
              },
              2: {
                color: "#8B5CF6",
                dataLabels: false,
              },
              3: {
                color: "#EC4899",
                dataLabels: {
                  seriesName: true,
                  category: true,
                  percent: true,
                  separator: " - ",
                  position: "outsideEnd" as const,
                },
              },
            },
          },
        ],
        dataLabels: {
          category: true,
          percent: true,
          position: "outsideEnd" as const,
          leaderLines: { color: "#1E293B", width: 1.5 },
        },
      },
    ],
  },
  {
    id: "pie-bestfit-labels",
    x: 16,
    y: 186,
    width: 220,
    height: 158,
    title: "BestFit Pie",
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "pie" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["Major", "Minor"],
        },
        series: [
          {
            name: "Split",
            values: [80, 20],
            points: {
              0: { color: "#3B82F6" },
              1: { color: "#F97316" },
            },
          },
        ],
        dataLabels: { value: true, position: "bestFit" as const },
      },
    ],
  },
] as const;

const pieDoughnutPage2Charts = [
  {
    id: "doughnut-shared-ring",
    x: 16,
    y: 16,
    width: 296,
    height: 328,
    title: "Shared Doughnut (60% Hole)",
    legend: "right" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "doughnut" as const,
        holeSize: 60,
        firstSliceAngle: 30,
        data: {
          kind: "shared-categories" as const,
          categories: ["North", "South", "West"],
        },
        series: [
          {
            name: "Q4",
            values: [50, 30, 20],
            points: {
              0: { color: "#2563EB" },
              1: { color: "#10B981" },
              2: { color: "#F59E0B" },
            },
          },
        ],
        dataLabels: {
          seriesName: true,
          percent: true,
          separator: "\n",
          style: { fontSize: 9, color: "#FFFFFF" },
        },
      },
    ],
  },
  {
    id: "doughnut-multi-independent",
    x: 328,
    y: 16,
    width: 296,
    height: 328,
    title: "Multi-Ring Independent",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "doughnut" as const,
        holeSize: 35,
        firstSliceAngle: 0,
        data: { kind: "independent-categories" as const },
        series: [
          {
            name: "InnerRing",
            categories: ["Core", "Edge"],
            values: [60, 40],
            points: {
              0: { color: "#1D4ED8" },
              1: { color: "#EC4899" },
            },
          },
          {
            name: "OuterRing",
            categories: ["Alpha", "Beta", "Gamma"],
            values: [30, 45, 25],
            points: {
              0: { color: "#14B8A6" },
              1: { color: "#8B5CF6" },
              2: { color: "#F97316" },
            },
          },
        ],
        dataLabels: { category: true, value: true },
      },
    ],
  },
] as const;

export const chartPieDoughnutAndLabelsParityScenario: CapabilityScenario = {
  id: "chart-pie-doughnut-and-labels-parity",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering 360-degree and exploded pie charts, single-ring shared and multi-ring independent-category doughnut charts, firstSliceAngle, holeSize, leaderLines, and 3-level data labels with seriesName, separator, and position overrides.",
  visualCheckpoints: [
    "Page 1 (pie-variants-and-labels): full 360-degree blue pie with seriesName+percent label, bestFit pie, and exploded multi-slice pie with insideEnd/outsideEnd/deleted labels and dark leaderLines.",
    "Page 2 (doughnut-single-multi-and-independent): left shared 60%-hole doughnut with right legend; right concentric 2-ring independent-category doughnut whose bottom legend displays only the inner ring categories ('Core', 'Edge').",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify pieChart/doughnutChart properties, dLbls separator/showSerName, leaderLines, c15:datalabelsRange, and legendEntry deletions.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "pie-variants-and-labels", background: "#FFFFFF" },
        ...pieDoughnutPage1Charts.map((props) => createElement(Chart, props)),
      ),
      createElement(
        Slide,
        { id: "doughnut-single-multi-and-independent", background: "#FFFFFF" },
        ...pieDoughnutPage2Charts.map((props) => createElement(Chart, props)),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({
        id: "pie-variants-and-labels",
        background: "#FFFFFF",
      }).append(...pieDoughnutPage1Charts.map((props) => chart(props))),
      slide({
        id: "doughnut-single-multi-and-independent",
        background: "#FFFFFF",
      }).append(...pieDoughnutPage2Charts.map((props) => chart(props))),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    expect(preview.pages[0]!.svg).toContain("<polyline");
    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let blueCount = 0;
    let greenCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount += 1;
      if (r === 0x10 && g === 0xb9 && b === 0x81) greenCount += 1;
    }
    assert.ok(blueCount > 500, `Expected #2563EB pie pixels, got ${blueCount}`);
    assert.ok(
      greenCount > 500,
      `Expected #10B981 pie pixels, got ${greenCount}`,
    );

    const png2 = PNG.sync.read(Buffer.from(preview.pages[1]!.png!));
    let pinkCount = 0;
    let purpleCount = 0;
    for (let i = 0; i < png2.data.length; i += 4) {
      const r = png2.data[i]!;
      const g = png2.data[i + 1]!;
      const b = png2.data[i + 2]!;
      if (r === 0xec && g === 0x48 && b === 0x99) pinkCount += 1;
      if (r === 0x8b && g === 0x5c && b === 0xf6) purpleCount += 1;
    }
    assert.ok(
      pinkCount > 200,
      `Expected #EC4899 inner ring pixels, got ${pinkCount}`,
    );
    assert.ok(
      purpleCount > 200,
      `Expected #8B5CF6 outer ring pixels, got ${purpleCount}`,
    );

    const files = unzipSync(pptx);
    const chart1 = strFromU8(files["ppt/charts/chart1.xml"]!);
    const chart2 = strFromU8(files["ppt/charts/chart2.xml"]!);
    const chart4 = strFromU8(files["ppt/charts/chart4.xml"]!);
    const chart5 = strFromU8(files["ppt/charts/chart5.xml"]!);
    expect(chart1).toContain('<c:firstSliceAng val="45"/>');
    expect(chart1).toContain('<c:showSerName val="1"/>');
    expect(chart1).toContain("<c:separator> - </c:separator>");
    expect(chart2).toContain('<c:showLeaderLines val="1"/>');
    expect(chart2).toContain(
      '<c:dLbl><c:idx val="2"/><c:delete val="1"/></c:dLbl>',
    );
    expect(chart4).toContain('<c:holeSize val="60"/>');
    expect(chart4).toContain("<c:separator>&#xA;</c:separator>");
    expect(chart5).toContain("<c15:datalabelsRange>");
    expect(chart5).toContain("<c:legendEntry>");
  },
};

const annotationsPage1Charts = [
  {
    id: "col-poly-moving-avg",
    x: 16,
    y: 16,
    width: 300,
    height: 328,
    title: "Polynomial & Moving Avg",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "column" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["C1", "C2", "C3", "C4", "C5"],
        },
        series: [
          {
            name: "Actual",
            color: "#2563EB",
            values: [4, 9, 17, 29, 44],
            trendlines: [
              {
                type: "polynomial" as const,
                order: 2,
                intercept: 1,
                forward: 1,
                backward: 0.5,
                displayEquation: true,
                displayRSquared: true,
                line: { color: "#EF4444", width: 2, dash: "dash" as const },
              },
              {
                type: "movingAverage" as const,
                period: 3,
                line: { color: "#10B981", width: 2 },
              },
            ],
          },
        ],
      },
    ],
  },
  {
    id: "scatter-exp-log-pow-lin",
    x: 324,
    y: 16,
    width: 300,
    height: 328,
    title: "Linear, Exp, Log & Power",
    legend: "bottom" as const,
    style: { fontFamily: jostFont, fontSize: 11 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: false as const,
        markers: true,
        series: [
          {
            name: "Observations",
            color: "#6366F1",
            x: [1, 2, 3, 4, 5],
            y: [2.2, 4.8, 9.1, 15.5, 24.0],
            trendlines: [
              {
                type: "linear" as const,
                forward: 1.5,
                displayEquation: true,
                line: { color: "#F59E0B", width: 1.75 },
              },
              {
                type: "exponential" as const,
                line: { color: "#EC4899", width: 1.5, dash: "dot" as const },
              },
              {
                type: "logarithmic" as const,
                line: { color: "#14B8A6", width: 1.5 },
              },
              {
                type: "power" as const,
                line: { color: "#8B5CF6", width: 1.5 },
              },
            ],
          },
        ],
      },
    ],
  },
] as const;

const annotationsPage2Charts = [
  {
    id: "col-fixed-pct-err",
    x: 16,
    y: 16,
    width: 300,
    height: 158,
    title: "Fixed+ & Percentage- (NoCap)",
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "column" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B", "C", "D"],
        },
        series: [
          {
            name: "FixedPlus",
            color: "#2563EB",
            values: [10, 20, 15, 25],
            errorBars: {
              type: "fixed" as const,
              value: 4,
              include: "plus" as const,
              line: { color: "#DC2626", width: 1.5 },
            },
          },
          {
            name: "PctNoCap",
            color: "#10B981",
            values: [12, 18, 22, 16],
            errorBars: {
              type: "percentage" as const,
              value: 20,
              include: "minus" as const,
              endCap: false,
              line: { color: "#7C3AED", width: 1.5 },
            },
          },
        ],
      },
    ],
  },
  {
    id: "line-stddev-stderr",
    x: 324,
    y: 16,
    width: 300,
    height: 158,
    title: "StdDev (Mean) & StdErr",
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "line" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["T1", "T2", "T3", "T4"],
        },
        series: [
          {
            name: "StdDev",
            color: "#0EA5E9",
            marker: true,
            values: [10, 20, 30, 40],
            errorBars: {
              type: "standardDeviation" as const,
              value: 1,
              line: { color: "#EA580C", width: 1.5 },
            },
          },
          {
            name: "StdErr",
            color: "#F59E0B",
            marker: true,
            values: [14, 18, 22, 26],
            errorBars: { type: "standardError" as const },
          },
        ],
      },
    ],
  },
  {
    id: "scatter-xy-custom-err",
    x: 16,
    y: 186,
    width: 300,
    height: 158,
    title: "Scatter XY Custom + Fixed",
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        line: false as const,
        markers: true,
        series: [
          {
            name: "XY",
            color: "#2563EB",
            x: [2, 4, null, 8],
            y: [5, 12, 9, 18],
            errorBars: {
              x: {
                type: "custom" as const,
                plus: [0.6, 0.8, null, 1.0],
                minus: [0.4, null, 0.3, 0.7],
                line: { color: "#059669", width: 1.5 },
              },
              y: {
                type: "fixed" as const,
                value: 2.5,
                line: { color: "#DB2777", width: 1.5 },
              },
            },
          },
        ],
      },
    ],
  },
  {
    id: "bar-trendline-and-err",
    x: 324,
    y: 186,
    width: 300,
    height: 158,
    title: "Bar Trendline & Error",
    style: { fontFamily: jostFont, fontSize: 10 },
    groups: [
      {
        kind: "bar" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["N", "S", "W"],
        },
        series: [
          {
            name: "Volume",
            color: "#6366F1",
            values: [12, 22, 30],
            trendlines: [
              {
                type: "linear" as const,
                line: { color: "#EF4444", width: 1.5 },
              },
            ],
            errorBars: {
              type: "fixed" as const,
              value: 3,
              line: { color: "#111827", width: 1.5 },
            },
          },
        ],
      },
    ],
  },
] as const;

export const chartTrendlinesAndErrorBarsParityScenario: CapabilityScenario = {
  id: "chart-trendlines-and-error-bars-parity",
  capability: "rp:chart",
  description:
    "Dual-author Chart scenario covering all 6 trendline types (linear, exponential, logarithmic, polynomial with intercept, power, movingAverage) with forward/backward forecasts and equation/R-squared display, plus all 5 error bar modes (fixed, percentage, standardDeviation, standardError, custom) with include/endCap options and scatter XY bidirectional error bars.",
  visualCheckpoints: [
    "Page 1 (all-six-trendlines): column chart with dashed red polynomial curve (backward=0.5, forward=1, eq+R2) and green movingAverage; scatter chart with linear, exponential, logarithmic, and power curves.",
    "Page 2 (all-five-error-bars-and-xy): column fixed-plus (#DC2626) and percentage-minus without end caps (#7C3AED); line mean-centered stdDev (#EA580C) and stdErr; scatter bidirectional custom X (#059669) and fixed Y (#DB2777) error bars; horizontal bar trendline and error bars.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and verify all 6 trendlineType variants, dispEq/dispRSqr, errValType modes, noEndCap, and independent scatter X/Y error bar channels.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "all-six-trendlines", background: "#FFFFFF" },
        ...annotationsPage1Charts.map((props) => createElement(Chart, props)),
      ),
      createElement(
        Slide,
        { id: "all-five-error-bars-and-xy", background: "#FFFFFF" },
        ...annotationsPage2Charts.map((props) => createElement(Chart, props)),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360, theme: chartTheme }).append(
      slide({ id: "all-six-trendlines", background: "#FFFFFF" }).append(
        ...annotationsPage1Charts.map((props) => chart(props)),
      ),
      slide({
        id: "all-five-error-bars-and-xy",
        background: "#FFFFFF",
      }).append(...annotationsPage2Charts.map((props) => chart(props))),
    );
  },
  assert({ pptx, preview, validation }) {
    assert.deepEqual(validation.errors, []);
    assert.equal(preview.pages.length, 2);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png !== undefined);
      assert.ok(page.svg !== undefined);
    }

    const png1 = PNG.sync.read(Buffer.from(preview.pages[0]!.png!));
    let redCount = 0;
    let greenCount = 0;
    for (let i = 0; i < png1.data.length; i += 4) {
      const r = png1.data[i]!;
      const g = png1.data[i + 1]!;
      const b = png1.data[i + 2]!;
      if (r === 0xef && g === 0x44 && b === 0x44) redCount += 1;
      if (r === 0x10 && g === 0xb9 && b === 0x81) greenCount += 1;
    }
    assert.ok(
      redCount > 40,
      `Expected #EF4444 trendline pixels, got ${redCount}`,
    );
    assert.ok(
      greenCount > 30,
      `Expected #10B981 movingAverage pixels, got ${greenCount}`,
    );

    const png2 = PNG.sync.read(Buffer.from(preview.pages[1]!.png!));
    let errGreenCount = 0;
    let errPinkCount = 0;
    for (let i = 0; i < png2.data.length; i += 4) {
      const r = png2.data[i]!;
      const g = png2.data[i + 1]!;
      const b = png2.data[i + 2]!;
      if (r === 0x05 && g === 0x96 && b === 0x69) errGreenCount += 1;
      if (r === 0xdb && g === 0x27 && b === 0x77) errPinkCount += 1;
    }
    assert.ok(
      errGreenCount > 30,
      `Expected #059669 scatter X error bar pixels, got ${errGreenCount}`,
    );
    assert.ok(
      errPinkCount > 30,
      `Expected #DB2777 scatter Y error bar pixels, got ${errPinkCount}`,
    );

    const files = unzipSync(pptx);
    const chart1 = strFromU8(files["ppt/charts/chart1.xml"]!);
    const chart2 = strFromU8(files["ppt/charts/chart2.xml"]!);
    const chart3 = strFromU8(files["ppt/charts/chart3.xml"]!);
    const chart4 = strFromU8(files["ppt/charts/chart4.xml"]!);
    const chart5 = strFromU8(files["ppt/charts/chart5.xml"]!);
    expect(chart1).toContain('<c:trendlineType val="poly"/>');
    expect(chart1).toContain('<c:trendlineType val="movingAvg"/>');
    expect(chart2).toContain('<c:trendlineType val="linear"/>');
    expect(chart2).toContain('<c:trendlineType val="exp"/>');
    expect(chart2).toContain('<c:trendlineType val="log"/>');
    expect(chart2).toContain('<c:trendlineType val="power"/>');
    expect(chart3).toContain('<c:noEndCap val="1"/>');
    expect(chart4).toContain('<c:errValType val="stdDev"/>');
    expect(chart4).toContain('<c:errValType val="stdErr"/>');
    expect(chart5).toContain('<c:errValType val="cust"/>');
    expect(chart5).toContain('<c:errDir val="x"/>');
    expect(chart5).toContain('<c:errDir val="y"/>');
  },
};
