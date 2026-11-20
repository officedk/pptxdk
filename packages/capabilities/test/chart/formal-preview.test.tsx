import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { slideBackgroundCapability } from "@react-pptx/capabilities";
import { chartCapability } from "@react-pptx/capabilities/chart";
import { Chart } from "@react-pptx/capabilities/chart-elements";
import {
  chart,
  deck,
  documentCapabilities,
  font,
  slide,
} from "@react-pptx/capabilities/document";
import { type CapabilityModule, CommonNode } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Chart static preview, layout parity, and textCapability isolation", () => {
  const jostFont = font("Jost*", [
    {
      weight: 400,
      data: readFileSync(
        new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    },
  ]);

  it("renders finite clustered column chart with zero diagnostics, preserves native chart, and works without textCapability", async () => {
    const trimmedPreview = runDocPreview([
      slideBackgroundCapability,
      chartCapability,
    ]);

    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "column-defaults", background: "#FFFFFF" }).append(
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
    );

    const result = await trimmedPreview(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    assert.deepEqual(page.diagnostics, []);

    const files = unzipSync(result.pptx);
    const chartXml = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("Quarterly sales");
    expect(chartXml).toContain('<c:showVal val="1"');
    expect(chartXml).not.toContain("<c:manualLayout>");
    expect(chartXml).not.toContain("<c:majorGridlines");
    assert.ok(
      Object.keys(files).some((name) =>
        name.startsWith("ppt/embeddings/Microsoft_Excel_Worksheet"),
      ),
    );

    // JSX entry produces identical PNG through the full capability set.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "column-defaults", background: "#FFFFFF" },
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
      ),
    );
    const jsxResult = await runDocPreview(documentCapabilities)(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.png, page.png);
  });

  it("handles fontSize units/inheritance, zero/fractional scales, top label inset, gapWidth bar geometry, and rebuild isolation", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const salesChart = chart({
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
    });
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "column-defaults", background: "#FFFFFF" }).append(
        salesChart,
      ),
    );

    // 1. Native default cluster gapWidth=150 in rendered bar widths.
    const baseline = await previewDoc(doc);
    const baselinePng = baseline.pages[0]!.png!;
    const baselineSvg = baseline.pages[0]!.svg!;
    const baselinePptx = baseline.pptx.slice();
    const pixels = PNG.sync.read(Buffer.from(baselinePng));
    const widths: number[] = [];
    let barRun = 0;
    for (let x = 0; x < pixels.width; x++) {
      const offset = (240 * pixels.width + x) * 4;
      const isAccent1 =
        pixels.data[offset] === 79 &&
        pixels.data[offset + 1] === 129 &&
        pixels.data[offset + 2] === 189;
      if (isAccent1) barRun++;
      else if (barRun > 0) {
        widths.push(barRun);
        barRun = 0;
      }
    }
    expect(widths).toHaveLength(3);
    for (const width of widths) {
      expect(width).toBeGreaterThanOrEqual(45);
      expect(width).toBeLessThanOrEqual(56);
    }

    // 2. Numeric, unit-based, and inherited fontSize equivalence + rebuild isolation.
    salesChart.style({ fontSize: 24 });
    const numeric = await previewDoc(doc);
    assert.ok(numeric.pages[0]?.png !== undefined);
    // Earlier baseline artifact remains unchanged after mutation & rebuild.
    assert.deepEqual(baseline.pages[0]?.png, baselinePng);
    assert.equal(baseline.pages[0]?.svg, baselineSvg);
    assert.deepEqual(baseline.pptx, baselinePptx);

    for (const fontSize of ["18pt", "24px"] as const) {
      salesChart.style({ fontSize });
      assert.deepEqual(
        (await previewDoc(doc)).pages[0]?.png,
        numeric.pages[0]?.png,
      );
    }
    salesChart.style({ fontSize: undefined });
    doc.style({ fontSize: 24 });
    assert.deepEqual(
      (await previewDoc(doc)).pages[0]?.png,
      numeric.pages[0]?.png,
    );

    // 3. Zero and fractional values render cleanly without tick overflow.
    doc.style({ fontSize: undefined });
    salesChart.style({ fontSize: 16 });
    for (const values of [
      [0, 0, 0],
      [0.1, 0.3, 0.7],
    ]) {
      salesChart.set({
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A", "B", "C"] },
            series: [{ name: "Values", values }],
            dataLabels: { value: true },
          },
        ],
      });
      const res = await previewDoc(doc);
      assert.deepEqual(res.pages[0]?.diagnostics, []);
      assert.ok(res.pages[0]?.png !== undefined);
    }

    // 4. Untitled chart at top edge (0, 0) keeps highest value label unclipped.
    salesChart.set({
      title: undefined,
      x: 0,
      y: 0,
      width: 640,
      height: 360,
      legend: false,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A"] },
          series: [{ name: "Values", values: [40] }],
          dataLabels: { value: true },
        },
      ],
    });
    const edgeResult = await previewDoc(doc);
    assert.deepEqual(edgeResult.pages[0]?.diagnostics, []);
    assert.ok(edgeResult.pages[0]?.png !== undefined);
    const edgePixels = PNG.sync.read(Buffer.from(edgeResult.pages[0].png));
    let labelPixelsInsideTopBand = 0;
    for (let y = 1; y < 30; y++) {
      for (let x = 150; x < 500; x++) {
        const idx = (y * edgePixels.width + x) * 4;
        if (
          edgePixels.data[idx]! < 100 &&
          edgePixels.data[idx + 1]! < 100 &&
          edgePixels.data[idx + 2]! < 100
        ) {
          labelPixelsInsideTopBand++;
        }
      }
    }
    expect(labelPixelsInsideTopBand).toBeGreaterThan(20);
  });

  it("renders all bar/column variants, axes, titles, legends, compact frames, and long labels with zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const baseProps = {
      id: "parity",
      x: 20,
      y: 20,
      width: 560,
      height: 300,
      style: { fontFamily: jostFont, fontSize: 14 },
    } as const;

    const parityCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof chart>;
    }[] = [
      {
        label: "horizontal-bar",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "bar",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "combo-column-groups",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "Bars", values: [1, 2] }],
            },
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "MoreBars", values: [2, 3] }],
            },
          ],
        }),
      },
      {
        label: "stacked-column",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              stacking: "stacked",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "percent-stacked-column",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              stacking: "percent",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "explicit-gapWidth-and-overlap",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              gapWidth: 80,
              overlap: 10,
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [
                { name: "S1", values: [1, 2] },
                { name: "S2", values: [2, 3] },
              ],
            },
          ],
        }),
      },
      {
        label: "explicit-plotArea",
        node: chart({
          ...baseProps,
          plotArea: { left: 40, top: 20, width: 400, height: 220 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "explicit-axes",
        node: chart({
          ...baseProps,
          axes: {
            category: { kind: "category" },
            value: { kind: "value", max: 50 },
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "explicit-displayBlanksAs",
        node: chart({
          ...baseProps,
          displayBlanksAs: "zero",
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "legend-top-and-topRight",
        node: chart({
          ...baseProps,
          legend: "topRight",
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "rich-styled-title",
        node: chart({
          ...baseProps,
          title: { text: "Styled", style: { fontSize: 18, color: "#1D4ED8" } },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
      },
      {
        label: "negative-and-null-values-with-point-styles",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B", "C"] },
              series: [
                {
                  name: "S",
                  values: [-5, null, 20],
                  color: "#FF0000",
                  dataLabels: { value: true },
                },
              ],
              dataLabels: { value: true, category: true },
            },
          ],
        }),
      },
      {
        label: "omitted-font-size-fallback",
        node: chart({
          id: "parity",
          x: 20,
          y: 20,
          width: 400,
          height: 240,
          style: { fontFamily: jostFont },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
      },
      {
        label: "overflowing-label-adaptive",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: [
                  "An excessively long category name that cannot fit on a single line inside its column slot",
                  "B",
                ],
              },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
      },
      {
        label: "too-small-plot-adaptive",
        node: chart({
          id: "parity",
          x: 20,
          y: 20,
          width: 40,
          height: 40,
          style: { fontFamily: jostFont, fontSize: 16 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", values: [10] }],
            },
          ],
        }),
      },
      {
        label: "line-standard-smooth-markers-blanks-and-point-overrides",
        node: chart({
          ...baseProps,
          legend: "bottom",
          displayBlanksAs: "gap",
          groups: [
            {
              kind: "line",
              data: {
                kind: "shared-categories",
                categories: ["A", "B", "C", "D", "E"],
              },
              series: [
                {
                  name: "Smooth",
                  color: "#2563EB",
                  line: { width: 2, dash: "dash" },
                  smooth: true,
                  marker: { symbol: "diamond", size: 7, fill: "#1D4ED8" },
                  values: [10, 28, null, 18, 32],
                  points: {
                    1: {
                      color: "#EF4444",
                      line: { color: "#EF4444", width: 3 },
                      marker: { symbol: "star", size: 9, fill: "#EF4444" },
                    },
                  },
                },
                {
                  name: "AutoMarker",
                  marker: { symbol: "auto", size: 6 },
                  values: [14, 18, 24, 20, 26],
                },
              ],
              dataLabels: { value: true, position: "top" },
            },
          ],
        }),
      },
      {
        label: "line-stacked-and-percent",
        node: chart({
          ...baseProps,
          legend: "right",
          groups: [
            {
              kind: "line",
              stacking: "percent",
              data: {
                kind: "shared-categories",
                categories: ["P1", "P2", "P3"],
              },
              series: [
                { name: "L1", marker: true, values: [20, null, 30] },
                {
                  name: "L2",
                  marker: { symbol: "triangle" },
                  values: [20, 40, 10],
                },
              ],
            },
          ],
        }),
      },
      {
        label: "area-standard-midCat-and-blanks-gap",
        node: chart({
          ...baseProps,
          legend: "bottom",
          displayBlanksAs: "gap",
          groups: [
            {
              kind: "area",
              data: {
                kind: "shared-categories",
                categories: ["Jan", "Feb", "Mar", "Apr", "May"],
              },
              series: [
                {
                  name: "SeriesA",
                  color: "#2563EB",
                  line: { color: "#1E40AF", width: 2 },
                  values: [12, 24, null, 18, 28],
                },
              ],
              dataLabels: { value: true },
            },
          ],
        }),
      },
      {
        label: "area-stacked-and-percent",
        node: chart({
          ...baseProps,
          legend: "top",
          groups: [
            {
              kind: "area",
              stacking: "stacked",
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                { name: "Base", color: "#10B981", values: [10, 20, 15] },
                { name: "Top", color: "#3B82F6", values: [15, null, 25] },
              ],
            },
          ],
        }),
      },
      {
        label: "scatter-five-modes-unequal-xy-and-empty-series",
        node: chart({
          ...baseProps,
          legend: "bottom",
          displayBlanksAs: "span",
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              line: "smooth",
              markers: true,
              series: [
                {
                  name: "Observed",
                  color: "#2563EB",
                  line: { width: 2, dash: "solid" },
                  marker: { symbol: "circle", size: 6 },
                  x: [1, 2, null, 4, 6],
                  y: [3, 8, 10, 5, 11],
                  points: {
                    3: {
                      color: "#EF4444",
                      marker: { symbol: "diamond", size: 8, fill: "#EF4444" },
                    },
                  },
                },
                {
                  name: "EmptySeries",
                  x: [],
                  y: [],
                },
              ],
              dataLabels: { xValue: true, value: true, position: "right" },
            },
          ],
        }),
      },
      {
        label: "scatter-markers-only-and-log-axes",
        node: chart({
          ...baseProps,
          legend: "right",
          axes: {
            x: { kind: "value", logBase: 10, min: 1, max: 100 },
            y: { kind: "value", min: 0, max: 50 },
          },
          groups: [
            {
              kind: "scatter",
              axes: { x: "x", y: "y" },
              data: { kind: "xy" },
              line: false,
              markers: true,
              series: [
                {
                  name: "LogScatter",
                  marker: { symbol: "triangle", size: 7, fill: "#F59E0B" },
                  x: [1, 10, 100],
                  y: [12, 28, 44],
                },
              ],
            },
          ],
        }),
      },
    ];

    for (const item of parityCases) {
      const result = await previewDoc(
        deck({ width: 640, height: 360 }).append(
          slide({ id: "parity-page" }).append(item.node),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.pptx)).errors,
        [],
        item.label,
      );
      assert.deepEqual(result.pages[0]?.diagnostics, [], item.label);
      assert.ok(result.pages[0]?.png !== undefined, item.label);
      assert.ok(result.pages[0]?.svg !== undefined, item.label);
    }
  });

  it("rejects remaining unsupported chart kinds/features and missing font files with located errors", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const validColumnChart = (id: string) =>
      chart({
        id,
        x: 20,
        y: 20,
        width: 280,
        height: 150,
        style: { fontFamily: jostFont, fontSize: 12 },
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A", "B"] },
            series: [{ name: "S", values: [4, 8] }],
          },
        ],
      });

    const baseProps = {
      id: "bad",
      x: 20,
      y: 20,
      width: 560,
      height: 300,
      style: { fontFamily: jostFont, fontSize: 14 },
    } as const;

    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof chart>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "non-column-pie",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "pie",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, 2] }],
            },
          ],
        }),
        expected: /pie/i,
      },
      {
        label: "series-errorBars",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [
                {
                  name: "S",
                  values: [10, 20],
                  errorBars: { type: "fixed", value: 2 },
                },
              ],
            },
          ],
        }),
        expected: /error bars/i,
      },
      {
        label: "series-trendlines",
        node: chart({
          ...baseProps,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [
                {
                  name: "S",
                  values: [10, 20],
                  trendlines: [{ type: "linear" }],
                },
              ],
            },
          ],
        }),
        expected: /trendlines/i,
      },
      {
        label: "missing-font-file",
        node: chart({
          id: "bad",
          x: 20,
          y: 20,
          width: 400,
          height: 240,
          style: { fontSize: 14 },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
        expected: /font file/i,
      },
    ];

    for (const item of rejectionCases) {
      const result = await previewDoc(
        deck({ width: 640, height: 360 }).append(
          slide({ id: "bad-page" }).append(item.node),
          slide({ id: "good-page" }).append(validColumnChart("ok")),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /chart\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
    }
  });
});
