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
    "Dual-author Chart rejection scenario verifying unsupported chart kinds (line/pie) error isolation while retaining native PPTX parts and unaffected pages.",
  visualCheckpoints: [
    "Page 1 (unsupported-line-with-column): no PNG emitted; located error on chart[trend] while chart[bars] has no warning.",
    "Page 2 (unsupported-pie): no PNG emitted; located error on chart[slices].",
    "Page 3 (unaffected-column): white canvas with valid clustered column chart 'Healthy' ([12, 20]), blue bars (#2563EB), Y ticks, X labels (H1, H2), and zero diagnostics.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS and verify Slide 1's line & column charts, Slide 2's pie chart, and Slide 3's column chart all remain native charts.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360, theme: chartTheme },
      createElement(
        Slide,
        { id: "unsupported-line-with-column", background: "#FFFFFF" },
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
        { id: "unsupported-pie", background: "#FFFFFF" },
        createElement(Chart, {
          id: "slices",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "pie",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Share", values: [40, 60] }],
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
        id: "unsupported-line-with-column",
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
      slide({ id: "unsupported-pie", background: "#FFFFFF" }).append(
        chart({
          id: "slices",
          x: 40,
          y: 20,
          width: 560,
          height: 300,
          style: { fontFamily: jostFont, fontSize: 14 },
          groups: [
            {
              kind: "pie",
              data: { kind: "shared-categories", categories: ["X", "Y"] },
              series: [{ name: "Share", values: [40, 60] }],
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
    expect(page1.diagnostics).toEqual([
      expect.objectContaining({
        severity: "error",
        path: "deck/slide[unsupported-line-with-column]/chart[trend]",
      }),
    ]);

    const page2 = preview.pages[1]!;
    assert.equal(page2.png, undefined);
    assert.equal(page2.svg, undefined);
    expect(page2.diagnostics).toEqual([
      expect.objectContaining({
        severity: "error",
        path: "deck/slide[unsupported-pie]/chart[slices]",
      }),
    ]);

    const page3 = preview.pages[2]!;
    assert.ok(page3.png !== undefined);
    assert.deepEqual(page3.diagnostics, []);

    const files = unzipSync(pptx);
    expect(strFromU8(files["ppt/charts/chart1.xml"]!)).toContain(
      "<c:lineChart>",
    );
    expect(strFromU8(files["ppt/charts/chart2.xml"]!)).toContain(
      "<c:barChart>",
    );
    expect(strFromU8(files["ppt/charts/chart3.xml"]!)).toContain(
      "<c:pieChart>",
    );
    expect(strFromU8(files["ppt/charts/chart4.xml"]!)).toContain(
      "<c:barChart>",
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
              data: { kind: "shared-categories", categories: ["Q1", "Q2", "Q3"] },
              series: [{ name: "Volume", values: [10, 120, 850], color: "#2563EB" }],
            },
            {
              kind: "column",
              axes: { category: "cat", value: "pctVal" },
              data: { kind: "shared-categories", categories: ["Q1", "Q2", "Q3"] },
              series: [{ name: "Share", values: [0.25, 0.5, 0.75], color: "#F59E0B" }],
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
              data: { kind: "shared-categories", categories: ["First", "Second", "Third"] },
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
              data: { kind: "shared-categories", categories: ["Alpha", "Beta"] },
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
              data: { kind: "shared-categories", categories: ["Q1", "Q2", "Q3"] },
              series: [{ name: "Volume", values: [10, 120, 850], color: "#2563EB" }],
            },
            {
              kind: "column",
              axes: { category: "cat", value: "pctVal" },
              data: { kind: "shared-categories", categories: ["Q1", "Q2", "Q3"] },
              series: [{ name: "Share", values: [0.25, 0.5, 0.75], color: "#F59E0B" }],
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
              data: { kind: "shared-categories", categories: ["First", "Second", "Third"] },
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
              data: { kind: "shared-categories", categories: ["Alpha", "Beta"] },
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
        if (png1.data[idx]! < 250 || png1.data[idx + 1]! < 250 || png1.data[idx + 2]! < 250) {
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
              data: { kind: "shared-categories", categories: ["North", "South", "West"] },
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
              data: { kind: "shared-categories", categories: ["Jan", "Feb", "Mar"] },
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
              data: { kind: "shared-categories", categories: ["Team A", "Team B"] },
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
              data: { kind: "shared-categories", categories: ["North", "South", "West"] },
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
              data: { kind: "shared-categories", categories: ["Jan", "Feb", "Mar"] },
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
              data: { kind: "shared-categories", categories: ["Team A", "Team B"] },
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
    assert.ok(redPixels > 100, `Expected #EF4444 override pixels, got ${redPixels}`);

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
    assert.ok(pinkPixels > 100, `Expected #EC4899 point override pixels, got ${pinkPixels}`);

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
