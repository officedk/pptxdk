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

export type ChartMechanismKey =
  | "chart-column"
  | "chart-bar"
  | "chart-line"
  | "chart-area"
  | "chart-pie"
  | "chart-doughnut"
  | "chart-scatter"
  | "chart-placeholder";

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
};
