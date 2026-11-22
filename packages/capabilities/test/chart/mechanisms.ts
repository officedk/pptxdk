import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart,
  deck,
  defineSlideLayout,
  font,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { Deck, Slide, SlideLayout, SlideMaster } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";

import type { CapabilityScenario } from "../acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../mechanism-verification.ts";

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

const comboCategories = ["Q1", "Q1", "Q3"] as const;

const comboProps = {
  id: "combo-chart",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  title: "Revenue & Share",
  legend: "bottom" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  axes: {
    period: { kind: "category" as const, title: "Period" },
    amount: {
      kind: "value" as const,
      position: "left" as const,
      min: 0,
      max: 40,
    },
    share: {
      kind: "value" as const,
      position: "right" as const,
      min: 0,
      max: 1,
      numberFormat: "0%",
    },
  },
  groups: [
    {
      kind: "column" as const,
      axes: { category: "period", value: "amount" },
      data: { kind: "shared-categories" as const, categories: comboCategories },
      series: [{ name: "Actual", values: [12, null, 28] }],
    },
    {
      kind: "line" as const,
      axes: { category: "period", value: "share" },
      data: { kind: "shared-categories" as const, categories: comboCategories },
      series: [{ name: "Share", values: [0.35, null, 0.85], smooth: true }],
    },
    {
      kind: "area" as const,
      axes: { category: "period", value: "amount" },
      data: { kind: "shared-categories" as const, categories: comboCategories },
      series: [{ name: "Baseline", values: [8, 16, 20] }],
    },
  ],
};

export const chartComboRoundtripScenario: CapabilityScenario = {
  id: "mechanism-chart-combo-roundtrip",
  capability: "chart",
  description:
    "Multi-group combo chart (column + line + area) with primary/secondary axes and duplicate category labels.",
  visualCheckpoints: [
    "Combo chart paints column, line, and area series with left and right value axes.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain group series ownership, duplicate category order, and primary/secondary c:axId bindings.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "combo-slide" },
        createElement(Chart, comboProps),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "combo-slide" }).append(chart(comboProps)),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 1);
  },
};

const barProps = {
  id: "bar-dup",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  legend: "bottom" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "bar" as const,
      data: {
        kind: "shared-categories" as const,
        categories: ["North", "North", "South"],
      },
      series: [{ name: "Units", values: [0, 15, null] }],
    },
  ],
};

const pieProps = {
  id: "pie-dup",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  legend: "right" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "pie" as const,
      firstSliceAngle: 90,
      data: {
        kind: "shared-categories" as const,
        categories: ["Direct", "Partner", "Partner"],
      },
      series: [{ name: "Mix", values: [40, 35, 25] }],
      dataLabels: { value: true, percent: true, category: true },
    },
  ],
};

const scatterProps = {
  id: "scatter-nulls",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "scatter" as const,
      line: "straight" as const,
      markers: true,
      data: { kind: "xy" as const },
      series: [
        {
          name: "Observations",
          x: [1, null, 3, 4],
          y: [10, 20, null, 40],
        },
      ],
    },
  ],
};

export const chartBarPieScatterRoundtripScenario: CapabilityScenario = {
  id: "mechanism-chart-bar-pie-scatter-roundtrip",
  capability: "chart",
  description:
    "Bar with zero/null/duplicate categories, Pie with duplicate categories, and Scatter with X/Y null slots.",
  visualCheckpoints: [
    "All three slides render distinct bar, pie, and scatter charts.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain barDir=bar, pie firstSliceAngle, and scatter X/Y null slot indices.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(Slide, { id: "bar-slide" }, createElement(Chart, barProps)),
      createElement(Slide, { id: "pie-slide" }, createElement(Chart, pieProps)),
      createElement(
        Slide,
        { id: "scatter-slide" },
        createElement(Chart, scatterProps),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "bar-slide" }).append(chart(barProps)),
      slide({ id: "pie-slide" }).append(chart(pieProps)),
      slide({ id: "scatter-slide" }).append(chart(scatterProps)),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 3);
  },
};

const sharedDoughnutProps = {
  id: "doughnut-shared",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  legend: "right" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "doughnut" as const,
      holeSize: 55,
      data: {
        kind: "shared-categories" as const,
        categories: ["Direct", "Partner", "Partner"],
      },
      series: [
        { name: "Q1", values: [10, 20, 30] },
        { name: "Q2", values: [15, 25, 10] },
      ],
    },
  ],
};

const independentDoughnutProps = {
  id: "doughnut-independent",
  x: 20,
  y: 20,
  width: 600,
  height: 320,
  legend: "right" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "doughnut" as const,
      holeSize: 50,
      data: { kind: "independent-categories" as const },
      series: [
        {
          name: "Ring A",
          categories: ["Direct", "Partner", "Partner"],
          values: [30, 20, null],
        },
        {
          name: "Ring B",
          categories: ["Partner", "Direct"],
          values: [40, 10],
        },
      ],
      dataLabels: { value: true, category: true },
    },
  ],
};

export const chartDoughnutRoundtripScenario: CapabilityScenario = {
  id: "mechanism-chart-doughnut-roundtrip",
  capability: "chart",
  description:
    "Shared-category and independent-category (c15:datalabelsRange) multi-ring doughnut roundtrip.",
  visualCheckpoints: [
    "Both shared-category and independent-category multi-ring doughnut charts paint cleanly.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX preserve shared vs independent doughnut categories and c15:datalabelsRange.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "doughnut-shared-slide" },
        createElement(Chart, sharedDoughnutProps),
      ),
      createElement(
        Slide,
        { id: "doughnut-independent-slide" },
        createElement(Chart, independentDoughnutProps),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "doughnut-shared-slide" }).append(chart(sharedDoughnutProps)),
      slide({ id: "doughnut-independent-slide" }).append(
        chart(independentDoughnutProps),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 2);
  },
};

function assertComboNative(
  validation: Parameters<
    ExhaustiveMechanismVerificationMap["chart-column"]["assertNative"]
  >[0]["validation"],
  stage: string,
): void {
  assert.equal(validation.errors.length, 0, `${stage} validation errors`);
  const ch = validation.slides[0]?.objects[0]?.chart;
  assert.ok(ch, `${stage} missing chart`);
  assert.deepEqual(ch.types, ["barChart", "lineChart", "areaChart"]);
  assert.equal(ch.axes.length, 3);
  assert.deepEqual(
    ch.axes.map(({ kind, position }) => ({ kind, position })),
    [
      { kind: "category", position: "bottom" },
      { kind: "value", position: "left" },
      { kind: "value", position: "right" },
    ],
  );
  assert.deepEqual(
    ch.groups.map(({ axisIds }) => axisIds),
    [
      [111111, 222222],
      [111111, 333333],
      [111111, 222222],
    ],
  );
  assert.deepEqual(
    ch.series.map(({ name, groupIndex, categories, values }) => ({
      name,
      groupIndex,
      categories,
      values,
    })),
    [
      {
        name: "Actual",
        groupIndex: 0,
        categories: ["Q1", "Q1", "Q3"],
        values: [12, null, 28],
      },
      {
        name: "Share",
        groupIndex: 1,
        categories: ["Q1", "Q1", "Q3"],
        values: [0.35, null, 0.85],
      },
      {
        name: "Baseline",
        groupIndex: 2,
        categories: ["Q1", "Q1", "Q3"],
        values: [8, 16, 20],
      },
    ],
  );
  assert.deepEqual(ch.embeddedWorkbook?.series, ch.series);
}

const chartHolderLayout = defineSlideLayout({
  id: "chart-holder-layout",
  name: "Chart Holder",
  placeholders: {
    plot: {
      type: "chart",
      explicitType: true,
      idx: 12,
      x: 50,
      y: 40,
      width: 320,
      height: 200,
      name: "Chart Slot",
    },
  },
});

const chartHolderProps = {
  id: "ph-pie-chart",
  legend: "bottom" as const,
  style: { fontFamily: jostFont, fontSize: 11 },
  groups: [
    {
      kind: "pie" as const,
      data: {
        kind: "shared-categories" as const,
        categories: ["Direct", "Partner"],
      },
      series: [{ name: "Share", values: [60, 40] }],
    },
  ],
};

export const chartPlaceholderRoundtripScenario: CapabilityScenario = {
  id: "mechanism-chart-placeholder-roundtrip",
  capability: "chart",
  description:
    "Dual-author chart-placeholder mechanism scenario verifying inherited bounds and <p:ph> binding roundtrip.",
  visualCheckpoints: [
    "Page 1 paints a pie chart inside inherited layout bounds (50, 40, 320x200) with zero diagnostics.",
  ],
  nativeCheckpoints: [
    "Both initial and roundtrip PPTX retain <p:nvPr><p:ph type='chart' idx='12'/></p:nvPr> and <p:xfrm> matching layout bounds.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        SlideMaster,
        { id: "chart-master", name: "Chart Master" },
        createElement(SlideLayout, { layout: chartHolderLayout }),
      ),
      createElement(Slide, {
        id: "chart-ph-slide",
        layout: chartHolderLayout,
        placeholders: {
          plot: createElement(Chart, chartHolderProps),
        },
      }),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "chart-master", name: "Chart Master" }).append(
        slideLayout(chartHolderLayout),
      ),
      slide({
        id: "chart-ph-slide",
        layout: chartHolderLayout,
        placeholders: {
          plot: chart(chartHolderProps),
        },
      }),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 1);
  },
};

const presentationComboProps = {
  id: "pres-combo",
  x: 16,
  y: 16,
  width: 608,
  height: 328,
  title: {
    text: "Presentation Combo",
    style: {
      fontSize: 16,
      fontWeight: "bold" as const,
      color: "#1D4ED8",
      textDecoration: "underline" as const,
    },
  },
  legend: "topRight" as const,
  displayBlanksAs: "gap" as const,
  style: { fontFamily: jostFont, fontSize: 11, color: "#0F172A" },
  axes: {
    cat: {
      kind: "category" as const,
      position: "top" as const,
      reversed: true,
      majorTickMark: "out" as const,
      tickLabelPosition: "high" as const,
      title: { text: "Period", style: { color: "#334155" } },
      line: { color: "#475569", width: 1.5, dash: "dashDot" as const },
    },
    val: {
      kind: "value" as const,
      position: "left" as const,
      reversed: true,
      min: 0,
      max: 100,
      majorUnit: 25,
      minorUnit: 5,
      numberFormat: "$#,##0",
      crosses: "min" as const,
      gridlines: { color: "#CBD5E1", width: 1, dash: "dash" as const },
      minorGridlines: true,
      style: { fontSize: 10, color: "#334155" },
      title: "Amount",
    },
    cat2: { kind: "category" as const, visible: false },
    val2: {
      kind: "value" as const,
      position: "right" as const,
      min: 0,
      max: 50,
      crossesAt: 2,
      numberFormat: "0.0",
    },
  },
  groups: [
    {
      kind: "column" as const,
      axes: { category: "cat", value: "val" },
      gapWidth: 80,
      overlap: -15,
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1", "Q2", "Q3", "Q4"],
      },
      dataLabels: {
        value: true,
        position: "outsideEnd" as const,
        numberFormat: "#,##0",
        style: { color: "#0F172A", textDecoration: "underline" as const },
      },
      series: [
        {
          name: "Primary",
          color: "#2563EB",
          values: [20, 45, 65, 80],
          points: {
            1: {
              color: "#DC2626",
              labelColor: "#FFFFFF",
              dataLabels: { position: "insideEnd" as const },
            },
            3: { dataLabels: false as const },
          },
          trendlines: [
            {
              type: "polynomial" as const,
              order: 2,
              intercept: 5,
              forward: 1,
              backward: 0.5,
              displayEquation: true,
              displayRSquared: true,
              line: { color: "#EF4444", width: 2, dash: "dash" as const },
            },
            {
              type: "linear" as const,
              forward: 0.5,
              line: { color: "#9333EA", width: 1.5 },
            },
          ],
          errorBars: {
            type: "fixed" as const,
            value: 4,
            include: "plus" as const,
            endCap: false,
            line: { color: "#111827", width: 1.5 },
          },
        },
        {
          name: "Muted",
          color: "#94A3B8",
          values: [15, 30, 40, 50],
          dataLabels: false as const,
        },
      ],
    },
    {
      kind: "line" as const,
      axes: { category: "cat2", value: "val2" },
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1", "Q2", "Q3", "Q4"],
      },
      series: [
        {
          name: "Signal",
          color: "#10B981",
          line: { width: 2, dash: "longDash" as const },
          smooth: true,
          marker: {
            symbol: "diamond" as const,
            size: 7,
            fill: "#10B981",
            line: { color: "#064E3B", width: 1 },
          },
          values: [10, 20, 30, 40],
          points: {
            2: {
              marker: {
                symbol: "triangle" as const,
                size: 10,
                fill: "#F59E0B",
              },
            },
          },
          errorBars: {
            type: "percentage" as const,
            value: 10,
            include: "both" as const,
            line: { color: "#065F46", width: 1 },
          },
        },
      ],
    },
  ],
};

const presentationScatterProps = {
  id: "pres-scatter",
  x: 16,
  y: 16,
  width: 300,
  height: 300,
  plotArea: { left: 48, top: 36, width: 220, height: 220 },
  style: { fontFamily: jostFont, fontSize: 10 },
  axes: {
    x: { kind: "value" as const, min: 0, max: 10, crossesAt: 5 },
    y: {
      kind: "value" as const,
      min: -10,
      max: 10,
      crosses: "autoZero" as const,
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
          name: "XY",
          color: "#6366F1",
          x: [1, 2, 4, 8],
          y: [2, 4, 6, 9],
          trendlines: [
            { type: "exponential" as const, line: { color: "#EC4899" } },
            { type: "logarithmic" as const },
            { type: "power" as const },
            { type: "movingAverage" as const, period: 2 },
          ],
          errorBars: {
            x: {
              type: "custom" as const,
              include: "both" as const,
              plus: [0.5, 0.5, 0.5, 0.5],
              minus: [0.2, 0.2, 0.2, 0.2],
            },
            y: { type: "standardError" as const, include: "both" as const },
          },
        },
        {
          name: "Deviation",
          color: "#0EA5E9",
          x: [2, 4, 6],
          y: [3, 5, 7],
          errorBars: {
            y: {
              type: "standardDeviation" as const,
              value: 1.5,
              include: "minus" as const,
            },
          },
        },
      ],
    },
  ],
};

const presentationPieProps = {
  id: "pres-pie",
  x: 324,
  y: 16,
  width: 300,
  height: 300,
  style: { fontFamily: jostFont, fontSize: 10 },
  groups: [
    {
      kind: "pie" as const,
      firstSliceAngle: 45,
      data: {
        kind: "shared-categories" as const,
        categories: ["Alpha", "Beta", "Gamma"],
      },
      dataLabels: {
        value: true,
        percent: true,
        separator: " | ",
        position: "outsideEnd" as const,
        leaderLines: { color: "#1E293B", width: 1.5, dash: "dot" as const },
      },
      series: [
        {
          name: "Share",
          values: [50, 30, 20],
          points: {
            0: {
              color: "#2563EB",
              explosion: 18,
              labelColor: "#1E3A8A",
            },
          },
        },
      ],
    },
  ],
};

export const chartPresentationRoundtripScenario: CapabilityScenario = {
  id: "mechanism-chart-presentation-roundtrip",
  capability: "chart",
  description:
    "3-level data labels (group/series/point + leaderLines), all 6 trendline types, all 5 error-bar modes + scatter XY, markers, point overrides, and c:layout vs c:manualLayout.",
  emitOptions: { embedFonts: "subset" },
  visualCheckpoints: [
    "Combo slide renders styled title, topRight legend, outsideEnd/insideEnd labels, polynomial/linear trendlines, error bars, and diamond/star markers.",
    "Scatter+Pie slide renders manual plotArea, exponential/log/power/moving-average trendlines, custom/SE/SD error bars, and pie leaderLines.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain label placement/format/colors, all 6 trendline types, all 5 error bar modes, markers/point explosions, and <c:layout/> vs <c:manualLayout>.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "pres-combo-slide" },
        createElement(Chart, presentationComboProps),
      ),
      createElement(
        Slide,
        { id: "pres-scatter-pie-slide" },
        createElement(Chart, presentationScatterProps),
        createElement(Chart, presentationPieProps),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "pres-combo-slide" }).append(chart(presentationComboProps)),
      slide({ id: "pres-scatter-pie-slide" }).append(
        chart(presentationScatterProps),
        chart(presentationPieProps),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.errors.length, 0);
    assert.equal(validation.slides.length, 2);
  },
};

export type ChartMechanismKey =
  | "chart-column"
  | "chart-bar"
  | "chart-line"
  | "chart-area"
  | "chart-pie"
  | "chart-doughnut"
  | "chart-scatter"
  | "chart-placeholder"
  | "chart-data-labels"
  | "chart-trendline"
  | "chart-error-bars"
  | "chart-markers-and-points";

export const chartMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  ChartMechanismKey
> = {
  "chart-placeholder": {
    scenario: chartPlaceholderRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const ch = validation.slides[0]?.objects[0];
      assert.equal(ch?.kind, "chart", `${stage} chart kind`);
      assert.deepEqual(ch?.placeholder, { type: "chart", idx: 12 });
      assert.equal(ch?.hasLocalTransform, true);
      assert.deepEqual(ch?.boxEmu, {
        x: 50 * 9525,
        y: 40 * 9525,
        width: 320 * 9525,
        height: 200 * 9525,
      });
      const slideXml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slideXml, /<p:ph type="chart" idx="12"\/>/u);
    },
    assertPaint({ initial, roundtrip }) {
      assert.equal(initial.pages.length, 1);
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
    },
  },
  "chart-column": {
    scenario: chartComboRoundtripScenario,
    assertNative({ stage, validation }) {
      assertComboNative(validation, stage);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
    },
  },
  "chart-line": {
    scenario: chartComboRoundtripScenario,
    assertNative({ stage, validation }) {
      assertComboNative(validation, stage);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
    },
  },
  "chart-area": {
    scenario: chartComboRoundtripScenario,
    assertNative({ stage, validation }) {
      assertComboNative(validation, stage);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 1);
    },
  },
  "chart-bar": {
    scenario: chartBarPieScatterRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const ch = validation.slides[0]?.objects[0]?.chart;
      assert.ok(ch);
      assert.deepEqual(ch.types, ["barChart"]);
      assert.equal(ch.groups[0]?.direction, "bar");
      assert.deepEqual(ch.series[0]?.categories, ["North", "North", "South"]);
      assert.deepEqual(ch.series[0]?.values, [0, 15, null]);
      assert.deepEqual(ch.embeddedWorkbook?.series, ch.series);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 3);
    },
  },
  "chart-pie": {
    scenario: chartBarPieScatterRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const ch = validation.slides[1]?.objects[0]?.chart;
      assert.ok(ch);
      assert.deepEqual(ch.types, ["pieChart"]);
      assert.equal(ch.groups[0]?.firstSliceAngle, 90);
      assert.deepEqual(ch.series[0]?.categories, [
        "Direct",
        "Partner",
        "Partner",
      ]);
      assert.deepEqual(ch.series[0]?.values, [40, 35, 25]);
      assert.deepEqual(ch.embeddedWorkbook?.series, ch.series);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[1]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 3);
    },
  },
  "chart-scatter": {
    scenario: chartBarPieScatterRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const ch = validation.slides[2]?.objects[0]?.chart;
      assert.ok(ch);
      assert.deepEqual(ch.types, ["scatterChart"]);
      assert.deepEqual(ch.series[0]?.x, [1, null, 3, 4]);
      assert.deepEqual(ch.series[0]?.y, [10, 20, null, 40]);
      assert.deepEqual(ch.embeddedWorkbook?.series, ch.series);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[2]?.diagnostics, []);
      assert.ok(initial.pages[2]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 3);
    },
  },
  "chart-doughnut": {
    scenario: chartDoughnutRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const sharedChart = validation.slides[0]?.objects[0]?.chart;
      assert.ok(sharedChart);
      assert.deepEqual(sharedChart.types, ["doughnutChart"]);
      assert.equal(sharedChart.groups[0]?.holeSize, 55);
      assert.equal(sharedChart.extensions, undefined);
      assert.deepEqual(sharedChart.series[0]?.categories, [
        "Direct",
        "Partner",
        "Partner",
      ]);
      assert.deepEqual(sharedChart.series[1]?.categories, [
        "Direct",
        "Partner",
        "Partner",
      ]);

      const indChart = validation.slides[1]?.objects[0]?.chart;
      assert.ok(indChart);
      assert.deepEqual(indChart.types, ["doughnutChart"]);
      assert.deepEqual(indChart.extensions, ["c15"]);
      assert.deepEqual(indChart.series[0]?.categories, [
        "Direct",
        "Partner",
        "Partner",
      ]);
      assert.deepEqual(indChart.series[0]?.values, [30, 20, null]);
      assert.equal(indChart.series[0]?.showDataLabelsRange, true);
      assert.deepEqual(indChart.series[0]?.datalabelsRange?.categories, [
        "Direct",
        "Partner",
        "Partner",
      ]);
      assert.deepEqual(indChart.series[1]?.categories, ["Partner", "Direct"]);
      assert.deepEqual(indChart.series[1]?.values, [40, 10]);
      assert.equal(indChart.series[1]?.showDataLabelsRange, true);
      assert.deepEqual(indChart.series[1]?.datalabelsRange?.categories, [
        "Partner",
        "Direct",
      ]);
      assert.deepEqual(
        indChart.embeddedWorkbook?.series,
        indChart.series.map(
          ({ groupIndex, index, name, order, categories, values }) => ({
            groupIndex,
            index,
            name,
            order,
            categories,
            values,
          }),
        ),
      );
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(initial.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.ok(initial.pages[1]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 2);
    },
  },
  "chart-data-labels": {
    scenario: chartPresentationRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const combo = validation.slides[0]?.objects[0]?.chart;
      assert.ok(combo, `${stage} missing combo chart`);
      const primaryLabels = combo.seriesDataLabels?.find(
        (entry) => entry.seriesIndex === 0,
      );
      assert.equal(
        primaryLabels?.dataLabels &&
          typeof primaryLabels.dataLabels === "object"
          ? primaryLabels.dataLabels.position
          : undefined,
        "outsideEnd",
      );
      assert.equal(
        primaryLabels?.pointDataLabels?.[1] &&
          typeof primaryLabels.pointDataLabels[1] === "object"
          ? primaryLabels.pointDataLabels[1].position
          : undefined,
        "insideEnd",
      );
      assert.equal(primaryLabels?.pointDataLabels?.[3], false);
      const mutedLabels = combo.seriesDataLabels?.find(
        (entry) => entry.seriesIndex === 1,
      );
      assert.equal(mutedLabels?.dataLabels, false);

      const pie = validation.slides[1]?.objects[1]?.chart;
      assert.ok(pie, `${stage} missing pie chart`);
      const pieLabels = pie.seriesDataLabels?.find(
        (entry) => entry.seriesIndex === 0,
      );
      assert.equal(
        pieLabels?.dataLabels && typeof pieLabels.dataLabels === "object"
          ? pieLabels.dataLabels.leaderLines
          : undefined,
        true,
      );
      assert.equal(
        pieLabels?.dataLabels && typeof pieLabels.dataLabels === "object"
          ? pieLabels.dataLabels.leaderLine?.color
          : undefined,
        "#1E293B",
      );

      const files = unzipSync(pptx);
      const chart1Xml = strFromU8(files["ppt/charts/chart1.xml"]!);
      assert.ok(chart1Xml.includes("<c:layout/>"));
      assert.ok(!chart1Xml.includes("<c:manualLayout>"));
      const chart3Xml = strFromU8(files["ppt/charts/chart3.xml"]!);
      assert.ok(chart3Xml.includes("<c:separator> | </c:separator>"));
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(initial.pages[1]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.ok(initial.pages[1]?.png instanceof Uint8Array);
      assert.ok(roundtrip.pages[1]?.svg?.includes("#1E293B"));
    },
  },
  "chart-trendline": {
    scenario: chartPresentationRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const combo = validation.slides[0]?.objects[0]?.chart;
      assert.ok(combo);
      assert.equal(combo.series[0]?.trendlines?.length, 2);
      assert.equal(combo.series[0]?.trendlines?.[0]?.type, "polynomial");
      assert.equal(combo.series[0]?.trendlines?.[0]?.order, 2);
      assert.equal(combo.series[0]?.trendlines?.[0]?.intercept, 5);
      assert.equal(combo.series[0]?.trendlines?.[0]?.forward, 1);
      assert.equal(combo.series[0]?.trendlines?.[0]?.backward, 0.5);
      assert.equal(combo.series[0]?.trendlines?.[0]?.displayEquation, true);
      assert.equal(combo.series[0]?.trendlines?.[0]?.displayRSquared, true);
      assert.equal(combo.series[0]?.trendlines?.[1]?.type, "linear");
      assert.equal(combo.series[1]?.trendlines, undefined);

      const scatter = validation.slides[1]?.objects[0]?.chart;
      assert.ok(scatter);
      assert.deepEqual(
        scatter.series[0]?.trendlines?.map((t) => t.type),
        ["exponential", "logarithmic", "power", "movingAverage"],
      );
      assert.equal(scatter.series[0]?.trendlines?.[3]?.period, 2);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[0]?.svg?.includes("#EF4444"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("#EF4444"));
      assert.ok(roundtrip.pages[1]?.svg?.includes("#EC4899"));
    },
  },
  "chart-error-bars": {
    scenario: chartPresentationRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const combo = validation.slides[0]?.objects[0]?.chart;
      assert.ok(combo);
      assert.equal(combo.series[0]?.errorBars?.[0]?.type, "fixed");
      assert.equal(combo.series[0]?.errorBars?.[0]?.value, 4);
      assert.equal(combo.series[0]?.errorBars?.[0]?.include, "plus");
      assert.equal(combo.series[0]?.errorBars?.[0]?.endCap, false);
      assert.equal(combo.series[1]?.errorBars, undefined);
      assert.equal(combo.series[2]?.errorBars?.[0]?.type, "percentage");

      const scatter = validation.slides[1]?.objects[0]?.chart;
      assert.ok(scatter);
      assert.equal(scatter.series[0]?.errorBars?.length, 2);
      assert.equal(scatter.series[0]?.errorBars?.[0]?.direction, "x");
      assert.equal(scatter.series[0]?.errorBars?.[0]?.type, "custom");
      assert.equal(scatter.series[0]?.errorBars?.[1]?.direction, "y");
      assert.equal(scatter.series[0]?.errorBars?.[1]?.type, "standardError");
      assert.equal(
        scatter.series[1]?.errorBars?.[0]?.type,
        "standardDeviation",
      );
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[0]?.svg?.includes("#111827"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("#111827"));
      assert.ok(roundtrip.pages[1]?.png instanceof Uint8Array);
    },
  },
  "chart-markers-and-points": {
    scenario: chartPresentationRoundtripScenario,
    assertNative({ stage, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} validation errors`);
      const combo = validation.slides[0]?.objects[0]?.chart;
      assert.ok(combo);
      assert.equal(combo.series[2]?.marker?.symbol, "diamond");
      assert.equal(combo.series[2]?.marker?.size, 7);
      assert.equal(combo.series[2]?.smooth, true);

      const pie = validation.slides[1]?.objects[1]?.chart;
      assert.ok(pie);
      assert.equal(pie.series[0]?.pointExplosions?.[0], 18);
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.deepEqual(roundtrip.pages[1]?.diagnostics, []);
      assert.ok(initial.pages[0]?.svg?.includes("#F59E0B"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("#F59E0B"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("#DC2626"));
    },
  },
};
