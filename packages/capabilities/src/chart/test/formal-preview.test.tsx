import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { slideBackgroundCapability } from "@pptxdk/capabilities";
import { documentCapabilities } from "@pptxdk/capabilities/all";
import { chartCapability } from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import { chart, deck, font, slide } from "@pptxdk/capabilities/document";
import { type CapabilityModule, CommonNode } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      resolution,
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
        new URL(
          "../../../test/assets/fonts/Jost-400-Book.ttf",
          import.meta.url,
        ),
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
    );

    const result = await trimmedPreview(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    assert.deepEqual(page.errors, []);

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
    const jsxResult = await runDocPreview(documentCapabilities)(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.png, page.png);
  });

  it("handles fontSize units/inheritance, zero/fractional scales, top label inset, gapWidth bar geometry, and rebuild isolation", async () => {
    const previewDoc = runDocPreview(documentCapabilities);

    const salesChart = chart({
      id: "sales",
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
      assert.deepEqual(res.pages[0]?.errors, []);
      assert.ok(res.pages[0]?.png !== undefined);
    }

    // 4. Untitled chart at top edge (0, 0) keeps highest value label unclipped.
    salesChart.style({
      position: "absolute",
      left: 0,
      top: 0,
      width: 640,
      height: 360,
    });
    salesChart.set({
      title: undefined,
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
    assert.deepEqual(edgeResult.pages[0]?.errors, []);
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
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 560,
        height: 300,
        fontFamily: jostFont,
        fontSize: 14,
      },
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
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 400,
            height: 240,
            fontFamily: jostFont,
          },
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
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 40,
            height: 40,
            fontFamily: jostFont,
            fontSize: 16,
          },
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
      assert.deepEqual(result.pages[0]?.errors, [], item.label);
      assert.ok(result.pages[0]?.png !== undefined, item.label);
      assert.ok(result.pages[0]?.svg !== undefined, item.label);
    }
  });

  it("wires seriesName and separator through resolve, emit, validatePptx, and preview", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "data-label-fields", background: "#FFFFFF" }).append(
        chart({
          id: "col-labels",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 290,
            height: 300,
            fontFamily: jostFont,
            fontSize: 12,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [
                {
                  name: "Rev",
                  values: [12, 24],
                },
              ],
              dataLabels: {
                seriesName: true,
                category: true,
                value: true,
                separator: " - ",
              },
            },
          ],
        }),
        chart({
          id: "scatter-labels",
          style: {
            position: "absolute",
            left: 330,
            top: 20,
            width: 290,
            height: 300,
            fontFamily: jostFont,
            fontSize: 12,
          },
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              series: [
                {
                  name: "Obs",
                  x: [1, 2],
                  y: [5, 10],
                },
              ],
              dataLabels: {
                seriesName: true,
                xValue: true,
                value: true,
                separator: "\n",
              },
            },
          ],
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    const files = unzipSync(result.pptx);
    const chart1 = strFromU8(files["ppt/charts/chart1.xml"]!);
    const chart2 = strFromU8(files["ppt/charts/chart2.xml"]!);
    expect(chart1).toContain('<c:showSerName val="1"/>');
    expect(chart1).toContain("<c:separator> - </c:separator>");
    expect(chart2).toContain('<c:showSerName val="1"/>');
    expect(chart2).toContain("<c:separator>&#xA;</c:separator>");
  });

  it("renders pie and single/multi-ring doughnut charts with 360-degree slice, explosion, leaderLines, and 3-level data labels at zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const pieDoc = deck({ width: 640, height: 360 }).append(
      slide({ id: "pie-doughnut", background: "#FFFFFF" }).append(
        chart({
          id: "pie-single-full",
          title: "Full 360",
          style: {
            position: "absolute",
            left: 16,
            top: 20,
            width: 190,
            height: 150,
            fontFamily: jostFont,
            fontSize: 11,
          },
          groups: [
            {
              kind: "pie",
              firstSliceAngle: 45,
              data: { kind: "shared-categories", categories: ["All"] },
              series: [
                {
                  name: "Only",
                  values: [100],
                  points: { 0: { color: "#2563EB" } },
                },
              ],
              dataLabels: { percent: true, position: "center" },
            },
          ],
        }),
        chart({
          id: "pie-leader-explosion",
          legend: "bottom",
          style: {
            position: "absolute",
            left: 220,
            top: 20,
            width: 200,
            height: 150,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "pie",
              firstSliceAngle: 90,
              explosion: 8,
              data: { kind: "shared-categories", categories: ["A", "B", "C"] },
              series: [
                {
                  name: "Share",
                  values: [50, 30, 20],
                  points: {
                    0: {
                      color: "#10B981",
                      explosion: 18,
                      labelColor: "#FFFFFF",
                      dataLabels: { position: "insideEnd", percent: true },
                    },
                    1: {
                      color: "#F59E0B",
                      dataLabels: false,
                    },
                  },
                },
              ],
              dataLabels: {
                category: true,
                percent: true,
                position: "outsideEnd",
                leaderLines: { color: "#1E293B", width: 1.5 },
              },
            },
          ],
        }),
        chart({
          id: "doughnut-multi-ring",
          legend: "bottom",
          style: {
            position: "absolute",
            left: 432,
            top: 20,
            width: 192,
            height: 300,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "doughnut",
              holeSize: 40,
              firstSliceAngle: 30,
              data: { kind: "independent-categories" },
              series: [
                {
                  name: "Inner",
                  categories: ["R1-A", "R1-B"],
                  values: [40, 60],
                  points: {
                    0: { color: "#3B82F6" },
                    1: { color: "#EC4899" },
                  },
                },
                {
                  name: "Outer",
                  categories: ["R2-X", "R2-Y", "R2-Z"],
                  values: [20, 50, 30],
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
        }),
      ),
    );

    const result = await previewDoc(pieDoc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    assert.ok(result.pages[0]?.svg !== undefined);
    const png = PNG.sync.read(Buffer.from(result.pages[0].png));
    let blueCount = 0;
    let pinkCount = 0;
    for (let i = 0; i < png.data.length; i += 4) {
      const r = png.data[i]!;
      const g = png.data[i + 1]!;
      const b = png.data[i + 2]!;
      if (r === 0x25 && g === 0x63 && b === 0xeb) blueCount += 1;
      if (r === 0xec && g === 0x48 && b === 0x99) pinkCount += 1;
    }
    expect(blueCount).toBeGreaterThan(200);
    expect(pinkCount).toBeGreaterThan(100);
  });

  it("renders all six trendline types, forward/backward extrapolation, equations/R2, and custom strokes at zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "trendlines", background: "#FFFFFF" }).append(
        chart({
          id: "col-trendlines",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 300,
            height: 160,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["C1", "C2", "C3", "C4", "C5"],
              },
              series: [
                {
                  name: "Growth",
                  color: "#2563EB",
                  values: [4, 9, 17, 29, 44],
                  trendlines: [
                    {
                      type: "polynomial",
                      order: 2,
                      intercept: 1,
                      forward: 1,
                      backward: 0.5,
                      displayEquation: true,
                      displayRSquared: true,
                      line: { color: "#EF4444", width: 2, dash: "dash" },
                    },
                    {
                      type: "movingAverage",
                      period: 3,
                      line: { color: "#10B981", width: 2 },
                    },
                  ],
                },
              ],
            },
          ],
        }),
        chart({
          id: "scatter-trendlines",
          style: {
            position: "absolute",
            left: 324,
            top: 16,
            width: 300,
            height: 160,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              line: false,
              markers: true,
              series: [
                {
                  name: "Fit",
                  color: "#6366F1",
                  x: [1, 2, 3, 4, 5],
                  y: [2.2, 4.8, 9.1, 15.5, 24.0],
                  trendlines: [
                    {
                      type: "linear",
                      forward: 1.5,
                      displayEquation: true,
                      line: { color: "#F59E0B", width: 1.75 },
                    },
                    {
                      type: "exponential",
                      line: { color: "#EC4899", width: 1.5, dash: "dot" },
                    },
                    {
                      type: "logarithmic",
                      line: { color: "#14B8A6", width: 1.5 },
                    },
                    {
                      type: "power",
                      line: { color: "#8B5CF6", width: 1.5 },
                    },
                  ],
                },
              ],
            },
          ],
        }),
        chart({
          id: "bar-trendline",
          style: {
            position: "absolute",
            left: 16,
            top: 186,
            width: 300,
            height: 158,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "bar",
              data: {
                kind: "shared-categories",
                categories: ["A", "B", "C", "D"],
              },
              series: [
                {
                  name: "Horizontal",
                  color: "#0EA5E9",
                  values: [10, 18, 25, 36],
                  trendlines: [
                    {
                      type: "linear",
                      line: { color: "#DC2626", width: 2 },
                    },
                  ],
                },
              ],
            },
          ],
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    const png = PNG.sync.read(Buffer.from(result.pages[0].png));
    let redCount = 0;
    let greenCount = 0;
    for (let i = 0; i < png.data.length; i += 4) {
      const r = png.data[i]!;
      const g = png.data[i + 1]!;
      const b = png.data[i + 2]!;
      if (r === 0xef && g === 0x44 && b === 0x44) redCount += 1;
      if (r === 0x10 && g === 0xb9 && b === 0x81) greenCount += 1;
    }
    expect(redCount).toBeGreaterThan(20);
    expect(greenCount).toBeGreaterThan(20);
  });

  it("renders all five error bar modes, include/endCap options, and scatter XY bidirectional error bars at zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "error-bars", background: "#FFFFFF" }).append(
        chart({
          id: "col-err",
          style: {
            position: "absolute",
            left: 16,
            top: 16,
            width: 300,
            height: 160,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: ["A", "B", "C", "D"],
              },
              series: [
                {
                  name: "FixedPlus",
                  color: "#2563EB",
                  values: [10, 20, 15, 25],
                  errorBars: {
                    type: "fixed",
                    value: 4,
                    include: "plus",
                    line: { color: "#DC2626", width: 1.5 },
                  },
                },
                {
                  name: "PctNoCap",
                  color: "#10B981",
                  values: [12, 18, 22, 16],
                  errorBars: {
                    type: "percentage",
                    value: 20,
                    include: "minus",
                    endCap: false,
                    line: { color: "#7C3AED", width: 1.5 },
                  },
                },
              ],
            },
          ],
        }),
        chart({
          id: "line-err-stddev",
          style: {
            position: "absolute",
            left: 324,
            top: 16,
            width: 300,
            height: 160,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "line",
              data: {
                kind: "shared-categories",
                categories: ["T1", "T2", "T3", "T4"],
              },
              series: [
                {
                  name: "StdDev",
                  color: "#0EA5E9",
                  marker: true,
                  values: [10, 20, 30, 40],
                  errorBars: {
                    type: "standardDeviation",
                    value: 1,
                    line: { color: "#EA580C", width: 1.5 },
                  },
                },
                {
                  name: "StdErr",
                  color: "#F59E0B",
                  marker: true,
                  values: [14, 18, 22, 26],
                  errorBars: {
                    type: "standardError",
                  },
                },
              ],
            },
          ],
        }),
        chart({
          id: "scatter-err-xy",
          style: {
            position: "absolute",
            left: 16,
            top: 186,
            width: 300,
            height: 158,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              line: false,
              markers: true,
              series: [
                {
                  name: "Obs",
                  color: "#2563EB",
                  x: [2, 4, null, 8],
                  y: [5, 12, 9, 18],
                  errorBars: {
                    x: {
                      type: "custom",
                      plus: [0.6, 0.8, null, 1.0],
                      minus: [0.4, null, 0.3, 0.7],
                      line: { color: "#059669", width: 1.5 },
                    },
                    y: {
                      type: "fixed",
                      value: 2.5,
                      line: { color: "#DB2777", width: 1.5 },
                    },
                  },
                },
              ],
            },
          ],
        }),
        chart({
          id: "bar-err",
          style: {
            position: "absolute",
            left: 324,
            top: 186,
            width: 300,
            height: 158,
            fontFamily: jostFont,
            fontSize: 10,
          },
          groups: [
            {
              kind: "bar",
              data: { kind: "shared-categories", categories: ["N", "S"] },
              series: [
                {
                  name: "Horiz",
                  color: "#6366F1",
                  values: [15, 25],
                  errorBars: {
                    type: "fixed",
                    value: 3,
                    line: { color: "#111827", width: 1.5 },
                  },
                },
              ],
            },
          ],
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(result.pages[0]?.errors, []);
    assert.ok(result.pages[0]?.png !== undefined);
    const png = PNG.sync.read(Buffer.from(result.pages[0].png));
    let redCount = 0;
    let pinkCount = 0;
    let greenCount = 0;
    for (let i = 0; i < png.data.length; i += 4) {
      const r = png.data[i]!;
      const g = png.data[i + 1]!;
      const b = png.data[i + 2]!;
      if (r === 0xdc && g === 0x26 && b === 0x26) redCount += 1;
      if (r === 0xdb && g === 0x27 && b === 0x77) pinkCount += 1;
      if (r === 0x05 && g === 0x96 && b === 0x69) greenCount += 1;
    }
    expect(redCount).toBeGreaterThan(20);
    expect(pinkCount).toBeGreaterThan(20);
    expect(greenCount).toBeGreaterThan(20);
  });

  it("rejects underlying font resource limitations with located errors", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const validColumnChart = (id: string) =>
      chart({
        id,
        style: {
          position: "absolute",
          left: 20,
          top: 20,
          width: 280,
          height: 150,
          fontFamily: jostFont,
          fontSize: 12,
        },
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A", "B"] },
            series: [{ name: "S", values: [4, 8] }],
          },
        ],
      });

    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof chart>;
      readonly code: "preview-missing-glyph" | "preview-missing-font";
      readonly expected: RegExp;
    }[] = [
      {
        label: "missing-glyph",
        node: chart({
          id: "bad",
          title: "汉",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 400,
            height: 240,
            fontFamily: jostFont,
            fontSize: 14,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
        code: "preview-missing-glyph",
        expected: /glyph/i,
      },
      {
        label: "missing-font-file",
        node: chart({
          id: "bad",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 400,
            height: 240,
            fontSize: 14,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [10, 20] }],
            },
          ],
        }),
        code: "preview-missing-font",
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
      assert.equal(result.pages[0]?.errors[0]?.code, item.code, item.label);
      assert.match(
        result.pages[0]?.errors[0]?.path ?? "",
        /chart\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.errors[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.errors, [], item.label);
    }
  });

  it("renders textDecoration (underline and line-through) and textShadow on chart title and rotated axis titles", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "styled-chart-text", background: "#FFFFFF" }).append(
        chart({
          id: "styled-title-chart",
          title: {
            text: "Underlined Title",
            style: {
              fontSize: 22,
              color: "#1D4ED8",
              textDecoration: "underline",
              textShadow: {
                color: "#0F172A",
                blur: 2,
                distance: 3,
                angle: 90,
              },
            },
          },
          axes: {
            x: { kind: "category" },
            y: {
              kind: "value",
              title: {
                text: "Struck Axis",
                style: {
                  fontSize: 14,
                  color: "#DC2626",
                  textDecoration: "line-through",
                  textShadow: {
                    color: "#0F172A",
                    blur: 1,
                    distance: 2,
                    angle: 45,
                  },
                },
              },
            },
          },
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 600,
            height: 320,
            fontFamily: jostFont,
            fontSize: 12,
          },
          groups: [
            {
              kind: "column",
              axes: { category: "x", value: "y" },
              data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
              series: [{ name: "Actual", values: [25, 45] }],
              dataLabels: {
                value: true,
                style: {
                  textDecoration: "underline",
                  color: "#047857",
                },
              },
            },
          ],
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);
    const svg = page.svg ?? "";
    expect(svg).toContain("<feDropShadow");
    expect(svg).toContain('fill="#1D4ED8"');
    // Decoration lines for underline / line-through are rendered as filled <rect> elements.
    expect(svg).toMatch(/<rect[^>]*fill="#1D4ED8"/);
    expect(svg).toMatch(/<rect[^>]*fill="#DC2626"/);
    expect(svg).toMatch(/<rect[^>]*fill="#047857"/);
  });

  it("reserves non-clipping margins for dual-axis combo charts with rotated vertical titles and multi-digit ticks when plotArea is omitted", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const chartStyle = {
      position: "absolute" as const,
      left: 20,
      top: 20,
      width: 560,
      height: 320,
      fontFamily: jostFont,
      fontSize: 13,
    };
    const axes = {
      quarter: { kind: "category" as const, title: "Quarter" },
      primary: {
        kind: "value" as const,
        position: "left" as const,
        title: "Revenue ($M, Adjusted)",
        min: 0,
        max: 1200,
        majorUnit: 300,
        numberFormat: "#,##0",
      },
      secondary: {
        kind: "value" as const,
        position: "right" as const,
        title: "Operating Margin (%)",
        min: 0,
        max: 1.2,
        majorUnit: 0.3,
        numberFormat: "0%",
      },
    };
    const groups = [
      {
        kind: "column" as const,
        axes: { category: "quarter", value: "primary" },
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3", "Q4"],
        },
        series: [{ name: "Revenue", values: [450, 720, 960, 1150] }],
      },
      {
        kind: "line" as const,
        axes: { category: "quarter", value: "secondary" },
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q2", "Q3", "Q4"],
        },
        series: [{ name: "Margin", values: [0.25, 0.45, 0.8, 1.1] }],
      },
    ];

    const vanillaResult = await previewDoc(
      deck({ width: 640, height: 360 }).append(
        slide({ id: "dual-axis-margins", background: "#FFFFFF" }).append(
          chart({
            id: "dual-combo",
            title: "Quarterly Revenue vs Margin",
            legend: "bottom",
            style: chartStyle,
            axes,
            groups,
          }),
        ),
      ),
    );
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "dual-axis-margins", background: "#FFFFFF" },
          createElement(Chart, {
            id: "dual-combo",
            title: "Quarterly Revenue vs Margin",
            legend: "bottom",
            style: chartStyle,
            axes,
            groups,
          }),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);

    assert.deepEqual((await validatePptx(vanillaResult.pptx)).errors, []);
    const files = unzipSync(vanillaResult.pptx);
    const chartXml = strFromU8(files["ppt/charts/chart1.xml"]!);
    expect(chartXml).not.toContain("<c:manualLayout>");

    const page = vanillaResult.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    assert.deepEqual(jsxResult.pages[0]?.png, page.png);

    // Only the outer chart box clipPath and plotArea clipPath should exist;
    // any additional clipPath indicates tick label or axis title overflow truncation.
    const clipMatches = page.svg.match(/<clipPath\b/gu) ?? [];
    expect(clipMatches).toHaveLength(2);

    // Verify rendered PNG margins and separation on both left and right vertical bands.
    const plotClipMatch = page.svg.match(
      /<clipPath id="rp-chart-clip-[^"]+-2"[^>]*><rect x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"\/>/u,
    );
    assert.ok(plotClipMatch !== null);
    const plotX = Number(plotClipMatch[1]);
    const plotY = Number(plotClipMatch[2]);
    const plotW = Number(plotClipMatch[3]);
    const plotH = Number(plotClipMatch[4]);
    const png = PNG.sync.read(Buffer.from(page.png));
    const yStart = Math.ceil(plotY) + 6;
    const yEnd = Math.floor(plotY + plotH) - 6;
    const columnHasInk = (x: number) => {
      for (let y = yStart; y <= yEnd; y += 1) {
        const idx = (y * png.width + x) * 4;
        if (
          png.data[idx]! < 245 ||
          png.data[idx + 1]! < 245 ||
          png.data[idx + 2]! < 245
        ) {
          return true;
        }
      }
      return false;
    };
    const findInkBands = (fromX: number, toX: number) => {
      const bands: { start: number; end: number }[] = [];
      let start = -1;
      for (let x = fromX; x <= toX; x += 1) {
        if (columnHasInk(x)) {
          if (start < 0) start = x;
        } else if (start >= 0) {
          bands.push({ start, end: x - 1 });
          start = -1;
        }
      }
      if (start >= 0) bands.push({ start, end: toX });
      return bands;
    };

    // Left side: white outer margin before rotated title, >= 2 ink bands (title + ticks) separated by clear gap.
    const leftBands = findInkBands(20, Math.floor(plotX) - 5);
    expect(leftBands.length).toBeGreaterThanOrEqual(2);
    expect(leftBands[0]!.start).toBeGreaterThanOrEqual(22);
    const leftTitleBand = leftBands[0]!;
    const leftTicksStart = leftBands[1]!.start;
    expect(leftTicksStart - leftTitleBand.end).toBeGreaterThanOrEqual(4);

    // Right side: ticks + rotated title separated by clear gap, white outer margin before box right edge (580).
    const rightBands = findInkBands(Math.ceil(plotX + plotW) + 5, 579);
    expect(rightBands.length).toBeGreaterThanOrEqual(2);
    const rightTicksEnd = rightBands[rightBands.length - 2]!.end;
    const rightTitleBand = rightBands[rightBands.length - 1]!;
    expect(rightTitleBand.start - rightTicksEnd).toBeGreaterThanOrEqual(4);
    expect(rightTitleBand.end).toBeLessThanOrEqual(577);
  });

  it("reserves non-clipping left margins for decimal-scale area, line, and scatter charts with rotated vertical titles and omitted fontSize", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const chartStyle = {
      position: "absolute" as const,
      left: 20,
      top: 20,
      width: 560,
      height: 300,
      fontFamily: jostFont,
    };

    const cases = [
      {
        label: "area-decimal-midCat",
        axes: {
          x: { kind: "category" as const, title: "Stage" },
          y: {
            kind: "value" as const,
            position: "left" as const,
            title: "Density (kg/L) and Cumulative Adjustment Across Batches",
            min: 0,
            max: 3.5,
            majorUnit: 0.5,
            numberFormat: "0.0",
          },
        },
        groups: [
          {
            kind: "area" as const,
            axes: { category: "x", value: "y" },
            data: {
              kind: "shared-categories" as const,
              categories: ["Start", "Mid", "End"],
            },
            series: [{ name: "Batch A", values: [0.5, 2.25, 3.4] }],
          },
        ],
      },
      {
        label: "line-decimal",
        axes: {
          x: { kind: "category" as const, title: "Sample" },
          y: {
            kind: "value" as const,
            position: "left" as const,
            title: "Score (0.0 - 3.5)",
            min: 0,
            max: 3.5,
            majorUnit: 0.5,
            numberFormat: "0.0",
          },
        },
        groups: [
          {
            kind: "line" as const,
            axes: { category: "x", value: "y" },
            data: {
              kind: "shared-categories" as const,
              categories: ["S1", "S2", "S3"],
            },
            series: [{ name: "Observed", values: [0.8, 1.9, 3.2] }],
          },
        ],
      },
      {
        label: "scatter-decimal-horizontal-axis-first",
        axes: {
          x: {
            kind: "value" as const,
            title: "Elapsed (h)",
            min: 0,
            max: 10,
            majorUnit: 2,
          },
          y: {
            kind: "value" as const,
            position: "left" as const,
            title: "Concentration (mg/mL)",
            min: 0,
            max: 3.5,
            majorUnit: 0.5,
            numberFormat: "0.0",
          },
        },
        groups: [
          {
            kind: "scatter" as const,
            axes: { x: "x", y: "y" },
            data: { kind: "xy" as const },
            series: [{ name: "Run 1", x: [0, 5, 10], y: [0.4, 2.1, 3.3] }],
          },
        ],
      },
    ] as const;

    for (const item of cases) {
      const vanillaResult = await previewDoc(
        deck({ width: 640, height: 360 }).append(
          slide({ id: item.label, background: "#FFFFFF" }).append(
            chart({
              id: item.label,
              style: chartStyle,
              axes: item.axes,
              groups: item.groups,
            }),
          ),
        ),
      );
      const jsxDoc = await expandDocument(
        createElement(
          Deck,
          { width: 640, height: 360 },
          createElement(
            Slide,
            { id: item.label, background: "#FFFFFF" },
            createElement(Chart, {
              id: item.label,
              style: chartStyle,
              axes: item.axes,
              groups: item.groups,
            }),
          ),
        ),
      );
      const jsxResult = await previewDoc(jsxDoc);

      assert.deepEqual(
        (await validatePptx(vanillaResult.pptx)).errors,
        [],
        item.label,
      );
      const chartXml = strFromU8(
        unzipSync(vanillaResult.pptx)["ppt/charts/chart1.xml"]!,
      );
      expect(chartXml, item.label).not.toContain("<c:manualLayout>");

      const page = vanillaResult.pages[0]!;
      assert.deepEqual(page.errors, [], item.label);
      assert.ok(page.png !== undefined, item.label);
      assert.ok(page.svg !== undefined, item.label);
      assert.deepEqual(jsxResult.pages[0]?.png, page.png, item.label);

      const clipMatches = page.svg.match(/<clipPath\b/gu) ?? [];
      expect(clipMatches, item.label).toHaveLength(2);

      const plotClipMatch = page.svg.match(
        /<clipPath id="rp-chart-clip-[^"]+-2"[^>]*><rect x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"\/>/u,
      );
      assert.ok(plotClipMatch !== null, item.label);
      const plotX = Number(plotClipMatch[1]);
      const plotY = Number(plotClipMatch[2]);
      const plotH = Number(plotClipMatch[4]);
      const png = PNG.sync.read(Buffer.from(page.png));
      const yStart = Math.ceil(plotY) + 6;
      const yEnd = Math.floor(plotY + plotH) - 6;
      const bands: { start: number; end: number }[] = [];
      let start = -1;
      for (let x = 20; x <= Math.floor(plotX) - 5; x += 1) {
        let ink = false;
        for (let y = yStart; y <= yEnd; y += 1) {
          const idx = (y * png.width + x) * 4;
          if (
            png.data[idx]! < 245 ||
            png.data[idx + 1]! < 245 ||
            png.data[idx + 2]! < 245
          ) {
            ink = true;
            break;
          }
        }
        if (ink) {
          if (start < 0) start = x;
        } else if (start >= 0) {
          bands.push({ start, end: x - 1 });
          start = -1;
        }
      }
      if (start >= 0) bands.push({ start, end: Math.floor(plotX) - 5 });
      expect(bands.length, item.label).toBeGreaterThanOrEqual(2);
      expect(bands[0]!.start, item.label).toBeGreaterThanOrEqual(22);
      expect(
        bands[1]!.start - bands[0]!.end,
        item.label,
      ).toBeGreaterThanOrEqual(4);
    }
  });

  it("preserves exact explicit plotArea pixel coordinates and native manualLayout across JSX and Vanilla", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const explicitPlotArea = { left: 96, top: 44, width: 410, height: 210 };
    const chartStyle = {
      position: "absolute" as const,
      left: 20,
      top: 20,
      width: 560,
      height: 320,
      fontFamily: jostFont,
      fontSize: 13,
    };
    const groups = [
      {
        kind: "area" as const,
        data: { kind: "shared-categories" as const, categories: ["A", "B"] },
        series: [{ name: "S", values: [1.5, 3.0] }],
      },
    ];

    const vanillaResult = await previewDoc(
      deck({ width: 640, height: 360 }).append(
        slide({ id: "explicit-plot", background: "#FFFFFF" }).append(
          chart({
            id: "explicit-area",
            style: chartStyle,
            plotArea: explicitPlotArea,
            groups,
          }),
        ),
      ),
    );
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "explicit-plot", background: "#FFFFFF" },
          createElement(Chart, {
            id: "explicit-area",
            style: chartStyle,
            plotArea: explicitPlotArea,
            groups,
          }),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);

    assert.deepEqual((await validatePptx(vanillaResult.pptx)).errors, []);
    const page = vanillaResult.pages[0]!;
    assert.deepEqual(page.errors, []);
    assert.deepEqual(jsxResult.pages[0]?.png, page.png);
    expect(page.svg).toContain(
      '<rect x="116" y="64" width="410" height="210"/>',
    );

    const chartXml = strFromU8(
      unzipSync(vanillaResult.pptx)["ppt/charts/chart1.xml"]!,
    );
    expect(chartXml).toContain("<c:manualLayout>");
    expect(chartXml).toContain('<c:x val="0.171429"/>');
    expect(chartXml).toContain('<c:y val="0.1375"/>');
    expect(chartXml).toContain('<c:w val="0.732143"/>');
    expect(chartXml).toContain('<c:h val="0.65625"/>');
  });
});
