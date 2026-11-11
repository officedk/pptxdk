import assert from "node:assert/strict";

import {
  BuildError,
  build,
  Chart,
  Deck,
  fade,
  Slide,
  type ChartGroup,
  type ChartProps,
} from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

function chartXml(pptx: Uint8Array): string {
  const parts = unzipSync(pptx);
  const chartPath = Object.keys(parts).find((part) =>
    /^ppt\/charts\/chart\d+\.xml$/u.test(part),
  );
  assert.ok(chartPath);
  return strFromU8(parts[chartPath]!);
}

async function buildSingleChart(config: ChartProps) {
  return build(
    <Deck width={960} height={540}>
      <Slide id="combo">
        <Chart {...config} />
      </Slide>
    </Deck>,
  );
}

describe("chart combination and dual-axis topology", () => {
  it("writes independent same-kind axis pairs with their own scale and position", async () => {
    const config: ChartProps = {
      x: 20,
      y: 20,
      width: 880,
      height: 480,
      axes: {
        primaryCategories: { kind: "category", title: "Primary period" },
        primaryValues: {
          kind: "value",
          title: "Primary amount",
          min: 0,
          max: 10,
          majorUnit: 2,
          numberFormat: "0",
        },
        secondaryCategories: {
          kind: "category",
          title: "Secondary period",
        },
        secondaryValues: {
          kind: "value",
          title: "Secondary amount",
          min: 0,
          max: 100,
          majorUnit: 20,
          numberFormat: "0%",
        },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "primaryCategories", value: "primaryValues" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Alpha", values: [1, 9] }],
        },
        {
          kind: "column",
          axes: {
            category: "secondaryCategories",
            value: "secondaryValues",
          },
          data: { kind: "shared-categories", categories: ["X", "Y"] },
          series: [{ name: "Beta", values: [10, 90] }],
        },
      ],
    };
    const result = await buildSingleChart(config);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(chart.types, ["barChart", "barChart"]);
    assert.equal(chart.axes.length, 4);
    assert.deepEqual(
      chart.groups.map(({ axisIds }) => axisIds),
      [
        [111111, 222222],
        [333333, 444444],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, position, crossAxisId, min, max }) => ({
        id,
        kind,
        position,
        crossAxisId,
        ...(min === undefined ? {} : { min }),
        ...(max === undefined ? {} : { max }),
      })),
      [
        {
          id: 111111,
          kind: "category",
          position: "bottom",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          position: "left",
          crossAxisId: 111111,
          min: 0,
          max: 10,
        },
        {
          id: 333333,
          kind: "category",
          position: "top",
          crossAxisId: 444444,
        },
        {
          id: 444444,
          kind: "value",
          position: "right",
          crossAxisId: 333333,
          min: 0,
          max: 100,
        },
      ],
    );
    assert.deepEqual(
      chart.series.map(({ name, index, order, groupIndex, categories }) => ({
        name,
        index,
        order,
        groupIndex,
        categories,
      })),
      [
        {
          name: "Alpha",
          index: 0,
          order: 0,
          groupIndex: 0,
          categories: ["A", "B"],
        },
        {
          name: "Beta",
          index: 1,
          order: 1,
          groupIndex: 1,
          categories: ["X", "Y"],
        },
      ],
    );
    assert.deepEqual(chart.embeddedWorkbook?.series, chart.series);

    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:catAx>/gu) ?? []).length, 2);
    assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 2);
    assert.match(xml, /<c:axPos val="b"\/>/u);
    assert.match(xml, /<c:axPos val="t"\/>/u);
    assert.match(xml, /<c:axPos val="l"\/>/u);
    assert.match(xml, /<c:axPos val="r"\/>/u);
    // Each series reads its own ordered category column and value column.
    assert.match(xml, /Sheet1!\$A\$2:\$A\$3/u);
    assert.match(xml, /Sheet1!\$B\$2:\$B\$3/u);
    assert.match(xml, /Sheet1!\$C\$2:\$C\$3/u);
    assert.match(xml, /Sheet1!\$D\$2:\$D\$3/u);
    assert.equal((xml.match(/<c:crosses val="max"\/>/gu) ?? []).length, 2);
  });

  it("shares one category and value axis across column, line, and area groups", async () => {
    const buildCombo = async (
      order: readonly ("column" | "line" | "area")[],
    ) => {
      const groups = order.map((kind, index) => ({
        kind,
        axes: { category: "period", value: "amount" },
        data: {
          kind: "shared-categories" as const,
          categories: ["Q1", "Q1", "Q3"],
        },
        series: [{ name: kind, values: [1, null, 3 + index] }],
      })) as ChartGroup[];
      return buildSingleChart({
        x: 10,
        y: 10,
        width: 880,
        height: 480,
        axes: {
          period: { kind: "category", title: "Period" },
          amount: { kind: "value", title: "Amount", min: 0, max: 20 },
        },
        groups,
      });
    };

    for (const order of [
      ["column", "line", "area"],
      ["line", "area", "column"],
    ] as const) {
      const result = await buildCombo(order);
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const chart = file.slides[0]?.objects[0]?.chart;
      assert.ok(chart);
      assert.equal(chart.axes.length, 2);
      assert.deepEqual(
        chart.groups.map(({ type, axisIds }) => [type, axisIds]),
        order.map((kind) => [
          kind === "column" ? "barChart" : `${kind}Chart`,
          [111111, 222222],
        ]),
      );
      assert.deepEqual(
        chart.series.map(({ name, index, order: seriesOrder, groupIndex }) => ({
          name,
          index,
          order: seriesOrder,
          groupIndex,
        })),
        order.map((kind, index) => ({
          name: kind,
          index,
          order: index,
          groupIndex: index,
        })),
      );
      assert.deepEqual(
        chart.series.map(({ categories }) => categories),
        order.map(() => ["Q1", "Q1", "Q3"]),
      );
      assert.deepEqual(
        chart.axes.map(({ id, kind, position, crossAxisId }) => ({
          id,
          kind,
          position,
          crossAxisId,
        })),
        [
          {
            id: 111111,
            kind: "category",
            position: "bottom",
            crossAxisId: 222222,
          },
          {
            id: 222222,
            kind: "value",
            position: "left",
            crossAxisId: 111111,
          },
        ],
      );
      const xml = chartXml(result.pptx);
      assert.equal((xml.match(/<c:catAx>/gu) ?? []).length, 1);
      assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 1);
      assert.equal((xml.match(/<c:crossAx val="111111"\/>/gu) ?? []).length, 1);
      assert.equal((xml.match(/<c:crossAx val="222222"\/>/gu) ?? []).length, 1);
    }
  });

  it("projects a shared category axis to a hidden companion for left and right values", async () => {
    const categories = ["A", "A", "B"];
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        period: {
          kind: "category",
          title: "Period",
          style: { fontSize: "9pt", color: "#333333" },
        },
        amount: { kind: "value", position: "left", min: 0, max: 10 },
        share: {
          kind: "value",
          position: "right",
          min: 0,
          max: 1,
          numberFormat: "0%",
        },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "period", value: "amount" },
          data: { kind: "shared-categories", categories },
          series: [{ name: "Amount", values: [1, null, 3] }],
        },
        {
          kind: "line",
          axes: { category: "period", value: "share" },
          data: { kind: "shared-categories", categories },
          series: [{ name: "Share", values: [0.35, null, 0.9] }],
        },
        {
          kind: "area",
          axes: { category: "period", value: "amount" },
          data: { kind: "shared-categories", categories },
          series: [{ name: "Area amount", values: [2, 4, 6] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.equal(chart.axes.length, 4);
    assert.deepEqual(
      chart.groups.map(({ axisIds }) => axisIds),
      [
        [111111, 222222],
        [333333, 444444],
        [111111, 222222],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, position, crossAxisId, crosses }) => ({
        id,
        kind,
        position,
        crossAxisId,
        crosses,
      })),
      [
        {
          id: 111111,
          kind: "category",
          position: "bottom",
          crossAxisId: 222222,
          crosses: "min",
        },
        {
          id: 222222,
          kind: "value",
          position: "left",
          crossAxisId: 111111,
          crosses: undefined,
        },
        {
          id: 333333,
          kind: "category",
          position: "bottom",
          crossAxisId: 444444,
          crosses: "max",
        },
        {
          id: 444444,
          kind: "value",
          position: "right",
          crossAxisId: 333333,
          crosses: undefined,
        },
      ],
    );
    assert.deepEqual(
      chart.series.map(({ categories: domain, values }) => ({
        domain,
        values,
      })),
      [
        { domain: categories, values: [1, null, 3] },
        { domain: categories, values: [0.35, null, 0.9] },
        { domain: categories, values: [2, 4, 6] },
      ],
    );
    assert.deepEqual(chart.categoryFormulas, ["Sheet1!$A$2:$A$4"]);
    assert.deepEqual(
      chart.embeddedWorkbook?.series.map(({ categories: domain, values }) => [
        domain,
        values,
      ]),
      [
        [categories, [1, null, 3]],
        [categories, [0.35, null, 0.9]],
        [categories, [2, 4, 6]],
      ],
    );
    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:catAx>/gu) ?? []).length, 2);
    assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 2);
    assert.equal((xml.match(/<c:delete val="1"\/>/gu) ?? []).length, 1);
    assert.equal((xml.match(/<a:t>Period<\/a:t>/gu) ?? []).length, 1);
    assert.equal((xml.match(/Sheet1!\$A\$2:\$A\$4/gu) ?? []).length, 3);
    assert.match(xml, /<c:crosses val="min"\/>/u);
    assert.match(xml, /<c:crosses val="max"\/>/u);
  });

  it("writes two independent scatter axis pairs with unequal observations", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        bx: { kind: "value", title: "Bottom X", min: 0, max: 10 },
        ly: { kind: "value", title: "Left Y", min: 0, max: 10 },
        tx: { kind: "value", title: "Top X", min: 0, max: 40 },
        ry: { kind: "value", title: "Right Y", min: 0, max: 40 },
      },
      groups: [
        {
          kind: "scatter",
          axes: { x: "bx", y: "ly" },
          data: { kind: "xy" },
          line: "straight",
          series: [
            {
              name: "Primary",
              x: [1, 2, 4],
              y: [3, null, 6],
            },
          ],
        },
        {
          kind: "scatter",
          axes: { x: "tx", y: "ry" },
          data: { kind: "xy" },
          markers: false,
          line: "smooth",
          series: [{ name: "Secondary", x: [10, 20], y: [1, 2] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(chart.types, ["scatterChart", "scatterChart"]);
    assert.deepEqual(
      chart.axes.map(({ id, kind, role, position, crossAxisId }) => ({
        id,
        kind,
        role,
        position,
        crossAxisId,
      })),
      [
        {
          id: 111111,
          kind: "value",
          role: "x",
          position: "bottom",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          role: "y",
          position: "left",
          crossAxisId: 111111,
        },
        {
          id: 333333,
          kind: "value",
          role: "x",
          position: "top",
          crossAxisId: 444444,
        },
        {
          id: 444444,
          kind: "value",
          role: "y",
          position: "right",
          crossAxisId: 333333,
        },
      ],
    );
    assert.deepEqual(
      chart.groups.map(({ axisIds, seriesIndices }) => ({
        axisIds,
        seriesIndices,
      })),
      [
        { axisIds: [111111, 222222], seriesIndices: [0] },
        { axisIds: [333333, 444444], seriesIndices: [1] },
      ],
    );
    assert.deepEqual(
      chart.series.map(({ name, x, y, index, order }) => ({
        name,
        x,
        y,
        index,
        order,
      })),
      [
        {
          name: "Primary",
          x: [1, 2, 4],
          y: [3, null, 6],
          index: 0,
          order: 0,
        },
        {
          name: "Secondary",
          x: [10, 20],
          y: [1, 2],
          index: 1,
          order: 1,
        },
      ],
    );
    assert.deepEqual(chart.embeddedWorkbook?.series, chart.series);
    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:scatterChart>/gu) ?? []).length, 2);
    assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 4);
    for (const pos of ["b", "l", "t", "r"]) {
      assert.match(xml, new RegExp(`<c:axPos val="${pos}"/>`, "u"));
    }
    assert.equal((xml.match(/<c:crosses val="max"\/>/gu) ?? []).length, 2);
  });

  it("shares a vertical value axis between a column group and a scatter group", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        period: { kind: "category" },
        value: { kind: "value", min: 0, max: 10 },
        scatterX: { kind: "value", min: 0, max: 10 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "period", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Bars", values: [2, 8] }],
        },
        {
          kind: "scatter",
          axes: { x: "scatterX", y: "value" },
          data: { kind: "xy" },
          line: "straight",
          series: [{ name: "Dots", x: [1, 2, 3], y: [3, 5, 9] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.equal(chart.axes.length, 3);
    assert.deepEqual(
      chart.groups.map(({ type, axisIds }) => [type, axisIds]),
      [
        ["barChart", [111111, 222222]],
        ["scatterChart", [333333, 222222]],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, role, position, crossAxisId }) => ({
        id,
        kind,
        role,
        position,
        crossAxisId,
      })),
      [
        {
          id: 111111,
          kind: "category",
          role: undefined,
          position: "bottom",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          role: "y",
          position: "left",
          crossAxisId: 111111,
        },
        {
          id: 333333,
          kind: "value",
          role: "x",
          position: "bottom",
          crossAxisId: 222222,
        },
      ],
    );
    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:catAx>/gu) ?? []).length, 1);
    assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 2);
    assert.match(xml, /Sheet1!\$A\$2:\$A\$3/u);
    assert.match(xml, /Sheet1!\$B\$2:\$B\$3/u);
    assert.match(xml, /Sheet1!\$C\$2:\$C\$4/u);
    assert.match(xml, /Sheet1!\$D\$2:\$D\$4/u);
  });

  it("shares a horizontal value axis between a bar group and a scatter group", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        category: { kind: "category" },
        barValue: { kind: "value", min: 0, max: 10 },
        scatterY: { kind: "value", min: 0, max: 10 },
      },
      groups: [
        {
          kind: "bar",
          axes: { category: "category", value: "barValue" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Bars", values: [2, 8] }],
        },
        {
          kind: "scatter",
          axes: { x: "barValue", y: "scatterY" },
          data: { kind: "xy" },
          markers: true,
          series: [{ name: "Dots", x: [1, 2, 3], y: [3, 5, 9] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.equal(chart.axes.length, 3);
    assert.deepEqual(
      chart.groups.map(({ type, direction, axisIds }) => [
        type,
        direction,
        axisIds,
      ]),
      [
        ["barChart", "bar", [111111, 222222]],
        ["scatterChart", undefined, [222222, 333333]],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, role, position, crossAxisId }) => ({
        id,
        kind,
        role,
        position,
        crossAxisId,
      })),
      [
        {
          id: 111111,
          kind: "category",
          role: undefined,
          position: "left",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          role: "x",
          position: "bottom",
          crossAxisId: 111111,
        },
        {
          id: 333333,
          kind: "value",
          role: "y",
          position: "left",
          crossAxisId: 222222,
        },
      ],
    );
  });

  it("keeps independent category domains, duplicate labels, and null slots", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        first: { kind: "category", title: "First" },
        firstValue: { kind: "value", position: "left" },
        second: { kind: "category", title: "Second" },
        secondValue: { kind: "value", position: "right" },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "first", value: "firstValue" },
          data: {
            kind: "shared-categories",
            categories: ["A", "A", "B"],
          },
          series: [{ name: "First", values: [1, null, 3] }],
        },
        {
          kind: "line",
          axes: { category: "second", value: "secondValue" },
          data: {
            kind: "shared-categories",
            categories: ["X", "X", "Y"],
          },
          series: [{ name: "Second", values: [4, 5, null] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(
      chart.series.map(({ categories, values }) => [categories, values]),
      [
        [
          ["A", "A", "B"],
          [1, null, 3],
        ],
        [
          ["X", "X", "Y"],
          [4, 5, null],
        ],
      ],
    );
    const workbookPath = chart.embeddedWorkbook?.part;
    assert.ok(workbookPath);
    const parts = unzipSync(result.pptx);
    const sheet = strFromU8(
      unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.match(sheet, /r="A2"[^>]*><is><t>A<\/t>/u);
    assert.match(sheet, /r="A3"[^>]*><is><t>A<\/t>/u);
    assert.match(sheet, /r="B2"[^>]*><is><t>X<\/t>/u);
    assert.match(sheet, /r="B3"[^>]*><is><t>X<\/t>/u);
    assert.ok(!sheet.includes('r="C3"'));
    assert.ok(sheet.includes('r="D3"'));
    assert.ok(sheet.includes('r="C2"'));
    assert.ok(!sheet.includes('r="D4"'));
  });

  it("diagnoses incompatible axis topology instead of guessing", async () => {
    const categoryData = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const singleSeries = [{ name: "S", values: [1, 2] }];
    const rejectsWithCode = async (props: ChartProps, code: string) => {
      await assert.rejects(buildSingleChart(props), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, code);
        return true;
      });
    };

    await rejectsWithCode(
      {
        x: 10,
        y: 10,
        width: 400,
        height: 300,
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
          {
            kind: "bar",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-chart-axes",
    );
    await rejectsWithCode(
      {
        groups: [
          {
            kind: "column",
            data: categoryData,
            series: singleSeries,
          },
          {
            kind: "line",
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-chart-axes",
    );
    await rejectsWithCode(
      {
        axes: {
          category: { kind: "category", position: "left" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-chart-axes",
    );
    await rejectsWithCode(
      {
        axes: {
          category: { kind: "category", position: "middle" as never },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-axis",
    );
    await rejectsWithCode(
      {
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "missing", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "unknown-chart-axis",
    );
    await rejectsWithCode(
      {
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
          unused: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-chart-axes",
    );
    await rejectsWithCode(
      {
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
          {
            kind: "line",
            axes: { category: "category", value: "value" },
            data: {
              kind: "shared-categories",
              categories: ["A", "C"],
            },
            series: singleSeries,
          },
        ],
      },
      "chart-category-domain-mismatch",
    );
    await rejectsWithCode(
      {
        axes: {
          x: { kind: "value" },
          y: { kind: "value" },
        },
        groups: [
          {
            kind: "scatter",
            axes: { x: "x", y: "y" },
            data: { kind: "xy" },
            series: [{ name: "S", x: [1], y: [2] }],
          },
          {
            kind: "scatter",
            axes: { x: "x", y: "x" },
            data: { kind: "xy" },
            series: [{ name: "T", x: [1], y: [2] }],
          },
        ],
      },
      "invalid-chart-axes",
    );
    await rejectsWithCode(
      {
        axes: {
          x: { kind: "category" },
          y: { kind: "value" },
        },
        groups: [
          {
            kind: "scatter",
            axes: { x: "x", y: "y" },
            data: { kind: "xy" },
            series: [{ name: "S", x: [1], y: [2] }],
          },
        ],
      },
      "invalid-chart-axis-role",
    );
    await rejectsWithCode(
      {
        displayBlanksAs: "span",
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
          {
            kind: "line",
            axes: { category: "category", value: "value" },
            data: categoryData,
            series: singleSeries,
          },
        ],
      },
      "invalid-chart-missing-values",
    );
  });

  it("keeps category and independent XY columns separate in one chart", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        period: { kind: "category" },
        categoryValue: { kind: "value", min: 0, max: 10 },
        scatterX: { kind: "value", min: 0, max: 10 },
        scatterY: { kind: "value", min: 0, max: 10 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "period", value: "categoryValue" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Bars", values: [2, 8] }],
        },
        {
          kind: "scatter",
          axes: { x: "scatterX", y: "scatterY" },
          data: { kind: "xy" },
          series: [{ name: "Dots", x: [1, 2, 3], y: [9, 5, 1] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.equal(chart.axes.length, 4);
    assert.deepEqual(
      chart.groups.map(({ type, axisIds }) => [type, axisIds]),
      [
        ["barChart", [111111, 222222]],
        ["scatterChart", [333333, 444444]],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, role, position, crossAxisId }) => ({
        id,
        kind,
        role,
        position,
        crossAxisId,
      })),
      [
        {
          id: 111111,
          kind: "category",
          role: undefined,
          position: "bottom",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          role: undefined,
          position: "left",
          crossAxisId: 111111,
        },
        {
          id: 333333,
          kind: "value",
          role: "x",
          position: "top",
          crossAxisId: 444444,
        },
        {
          id: 444444,
          kind: "value",
          role: "y",
          position: "right",
          crossAxisId: 333333,
        },
      ],
    );
    const xml = chartXml(result.pptx);
    assert.match(xml, /Sheet1!\$A\$2:\$A\$3/u);
    assert.match(xml, /Sheet1!\$B\$2:\$B\$3/u);
    assert.match(xml, /Sheet1!\$C\$2:\$C\$4/u);
    assert.match(xml, /Sheet1!\$D\$2:\$D\$4/u);
  });

  it("produces the same dual-axis chart from the React and command entries", async () => {
    const config: ChartProps = {
      id: "dual-entry-combo",
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category", title: "Period" },
        amount: { kind: "value", min: 0, max: 10 },
        share: { kind: "value", min: 0, max: 100, numberFormat: "0%" },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "amount" },
          data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
          series: [{ name: "Amount", values: [4, 8] }],
        },
        {
          kind: "line",
          axes: { category: "category", value: "share" },
          data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
          series: [{ name: "Share", values: [40, 80] }],
        },
      ],
    };
    const reactResult = await buildSingleChart(config);
    const documentResult = await buildDocument(
      deck({ width: 960, height: 540 }).append(
        documentSlide({ id: "combo" }).append(documentChart(config)),
      ),
    );
    const reactFile = await validatePptx(reactResult.pptx);
    const documentFile = await validatePptx(documentResult.pptx);
    assert.deepEqual(reactFile.errors, []);
    assert.deepEqual(documentFile.errors, []);
    assert.deepEqual(
      reactFile.slides[0]?.objects[0]?.chart,
      documentFile.slides[0]?.objects[0]?.chart,
    );
    assert.equal(chartXml(reactResult.pptx), chartXml(documentResult.pptx));
  });

  it("reads actual PPTX axis mutations as validator input", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        firstCategory: { kind: "category" },
        firstValue: { kind: "value", min: 0, max: 10 },
        secondCategory: { kind: "category" },
        secondValue: { kind: "value", min: 0, max: 100 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "firstCategory", value: "firstValue" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "First", values: [1, 9] }],
        },
        {
          kind: "line",
          axes: { category: "secondCategory", value: "secondValue" },
          data: { kind: "shared-categories", categories: ["X", "Y"] },
          series: [{ name: "Second", values: [10, 90] }],
        },
      ],
    });
    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const originalXml = strFromU8(parts[chartPath]!);
    const mutate = (edit: (xml: string) => string): Uint8Array => {
      const copy = { ...parts };
      copy[chartPath] = strToU8(edit(originalXml));
      return zipSync(copy);
    };
    const expectError = async (pptx: Uint8Array, code: string) => {
      const file = await validatePptx(pptx);
      assert.ok(
        file.errors.some(({ code: actual }) => actual === code),
        JSON.stringify(file.errors),
      );
    };

    await expectError(
      mutate((xml) => xml.replace('<c:axPos val="b"/>', '<c:axPos val="l"/>')),
      "invalid-chart-axis",
    );
    await expectError(
      mutate((xml) =>
        xml.replace(
          '<c:valAx><c:axId val="222222"/>',
          '<c:valAx><c:axId val="111111"/>',
        ),
      ),
      "invalid-chart-axis",
    );
    await expectError(
      mutate((xml) => {
        const start = xml.indexOf('<c:valAx><c:axId val="222222"/>');
        assert.ok(start >= 0);
        const end = xml.indexOf("</c:valAx>", start);
        assert.ok(end > start);
        const block = xml.slice(start, end);
        return (
          xml.slice(0, start) +
          block.replace(
            '<c:crossAx val="111111"/>',
            '<c:crossAx val="444444"/>',
          ) +
          xml.slice(end)
        );
      }),
      "invalid-chart-axis",
    );
    await expectError(
      mutate((xml) =>
        xml.replace(
          '<c:axId val="111111"/><c:axId val="222222"/>',
          '<c:axId val="222222"/><c:axId val="222222"/>',
        ),
      ),
      "invalid-chart-axis-role",
    );
    await expectError(
      mutate((xml) => xml.replaceAll("222222", "4294967296")),
      "invalid-chart-axis",
    );
    await expectError(
      mutate((xml) =>
        xml.replace(
          '<c:axId val="111111"/><c:axId val="222222"/>',
          '<c:axId val="999999"/><c:axId val="222222"/>',
        ),
      ),
      "unknown-chart-axis",
    );
    await expectError(
      mutate((xml) =>
        xml.replace("</c:chart>", '<c:dispBlanksAs val="span"/></c:chart>'),
      ),
      "invalid-chart-missing-values",
    );

    const shared = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value" },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "First", values: [1, 2] }],
        },
        {
          kind: "line",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Second", values: [3, 4] }],
        },
      ],
    });
    const sharedParts = unzipSync(shared.pptx);
    const sharedXml = strFromU8(sharedParts[chartPath]!);
    const secondCategory = sharedXml.indexOf(
      "<c:cat>",
      sharedXml.indexOf("<c:cat>") + 1,
    );
    assert.ok(secondCategory >= 0);
    const secondCategoryEnd = sharedXml.indexOf("</c:cat>", secondCategory);
    assert.ok(secondCategoryEnd > secondCategory);
    const secondBlock = sharedXml.slice(secondCategory, secondCategoryEnd);
    sharedParts[chartPath] = strToU8(
      sharedXml.slice(0, secondCategory) +
        secondBlock.replace("<c:v>A</c:v>", "<c:v>Z</c:v>") +
        sharedXml.slice(secondCategoryEnd),
    );
    await expectError(zipSync(sharedParts), "chart-category-domain-mismatch");
  });

  it("writes axis style, position, visibility, scale, ticks, crossing, and gridlines per axis", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        firstCategory: {
          kind: "category",
          title: "Cat 1",
          position: "top",
          style: { fontSize: "9pt", color: "#112233" },
          majorTickMark: "cross",
        },
        firstValue: {
          kind: "value",
          title: "Val 1",
          position: "right",
          min: 1,
          max: 100,
          logBase: 10,
          numberFormat: "0.0",
          gridlines: true,
          crosses: "min",
          style: { fontSize: "12pt", color: "#445566" },
          majorTickMark: "out",
          tickLabelPosition: "high",
          reversed: true,
        },
        secondCategory: {
          kind: "category",
          title: "Cat 2",
          position: "bottom",
          visible: false,
        },
        secondValue: {
          kind: "value",
          title: "Val 2",
          position: "left",
          min: 0,
          max: 1000,
          majorUnit: 250,
          numberFormat: "0%",
          minorGridlines: { color: "#CCCCCC", width: 1 },
          majorTickMark: "in",
        },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "firstCategory", value: "firstValue" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "First", values: [1, 9] }],
        },
        {
          kind: "column",
          axes: { category: "secondCategory", value: "secondValue" },
          data: { kind: "shared-categories", categories: ["X", "Y"] },
          series: [{ name: "Second", values: [10, 90] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(
      chart.axes.map(
        ({
          id,
          kind,
          position,
          visible,
          title,
          min,
          max,
          logBase,
          orientation,
          numberFormat,
          majorTickMark,
          tickLabelPosition,
          crosses,
          majorGridlines,
          minorGridlines,
        }) => ({
          id,
          kind,
          position,
          visible,
          title,
          min,
          max,
          logBase,
          orientation,
          numberFormat,
          majorTickMark,
          tickLabelPosition,
          crosses,
          majorGridlines,
          minorGridlines,
        }),
      ),
      [
        {
          id: 111111,
          kind: "category",
          position: "top",
          visible: true,
          title: "Cat 1",
          min: undefined,
          max: undefined,
          logBase: undefined,
          orientation: undefined,
          numberFormat: undefined,
          majorTickMark: "cross",
          tickLabelPosition: undefined,
          crosses: "max",
          majorGridlines: undefined,
          minorGridlines: undefined,
        },
        {
          id: 222222,
          kind: "value",
          position: "right",
          visible: true,
          title: "Val 1",
          min: 1,
          max: 100,
          logBase: 10,
          orientation: "maxMin",
          numberFormat: "0.0",
          majorTickMark: "out",
          tickLabelPosition: "high",
          crosses: "min",
          majorGridlines: true,
          minorGridlines: undefined,
        },
        {
          id: 333333,
          kind: "category",
          position: "bottom",
          visible: false,
          title: "Cat 2",
          min: undefined,
          max: undefined,
          logBase: undefined,
          orientation: undefined,
          numberFormat: undefined,
          majorTickMark: undefined,
          tickLabelPosition: undefined,
          crosses: "min",
          majorGridlines: undefined,
          minorGridlines: undefined,
        },
        {
          id: 444444,
          kind: "value",
          position: "left",
          visible: true,
          title: "Val 2",
          min: 0,
          max: 1000,
          logBase: undefined,
          orientation: undefined,
          numberFormat: "0%",
          majorTickMark: "in",
          tickLabelPosition: undefined,
          crosses: "min",
          majorGridlines: undefined,
          minorGridlines: true,
        },
      ],
    );
    const xml = chartXml(result.pptx);
    for (const text of ["Cat 1", "Val 1", "Cat 2", "Val 2"]) {
      assert.match(xml, new RegExp(`<a:t>${text}</a:t>`, "u"));
    }
    assert.match(xml, /<c:numFmt formatCode="0\.0" sourceLinked="0"\/>/u);
    assert.match(xml, /<c:numFmt formatCode="0%" sourceLinked="0"\/>/u);
    assert.match(xml, /<c:logBase val="10"\/>/u);
    assert.match(xml, /<c:orientation val="maxMin"\/>/u);
    assert.match(xml, /<c:majorTickMark val="out"\/>/u);
    assert.match(xml, /<c:majorTickMark val="in"\/>/u);
    assert.match(xml, /<c:tickLblPos val="high"\/>/u);
    assert.match(xml, /<c:crosses val="max"\/>/u);
    assert.match(xml, /<c:delete val="1"\/>/u);
    assert.equal((xml.match(/<c:majorGridlines(?:\/>|>)/gu) ?? []).length, 1);
    assert.equal((xml.match(/<c:minorGridlines(?:\/>|>)/gu) ?? []).length, 1);
    assert.match(xml, /<a:defRPr sz="900"/u);
    assert.match(xml, /<a:defRPr sz="1200"/u);
    assert.match(xml, /<a:srgbClr val="112233"\/>/u);
    assert.match(xml, /<a:srgbClr val="445566"\/>/u);
  });

  it("applies chart-level blank rendering across every participating group", async () => {
    const data = {
      kind: "shared-categories" as const,
      categories: ["A", "B", "C"],
    };
    const groups = [
      {
        kind: "column" as const,
        axes: { category: "category", value: "value" },
        data,
        series: [{ name: "Column", values: [1, null, 3] }],
      },
      {
        kind: "line" as const,
        axes: { category: "category", value: "value" },
        data,
        series: [{ name: "Line", values: [4, null, 6] }],
      },
      {
        kind: "area" as const,
        axes: { category: "category", value: "value" },
        data,
        series: [{ name: "Area", values: [7, null, 9] }],
      },
    ];
    const config = (displayBlanksAs: "gap" | "zero" | "span"): ChartProps => ({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value" },
      },
      displayBlanksAs,
      groups,
    });
    for (const mode of ["gap", "zero"] as const) {
      const result = await buildSingleChart(config(mode));
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      assert.equal(file.slides[0]?.objects[0]?.chart?.displayBlanksAs, mode);
      assert.match(
        chartXml(result.pptx),
        new RegExp(`<c:dispBlanksAs val="${mode}"/>`, "u"),
      );
    }
    await assert.rejects(buildSingleChart(config("span")), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "invalid-chart-missing-values");
      return true;
    });
  });

  it("keeps series and category motion targets aligned with the shared slot map", async () => {
    const groups = [
      {
        kind: "column" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B", "C"],
        },
        series: [{ name: "Column", values: [1, 2, 3] }],
      },
      {
        kind: "line" as const,
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B", "C"],
        },
        series: [{ name: "Line", values: [4, 5, 6] }],
      },
    ];
    const seriesMotion = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value" },
      },
      enter: fade({ by: "series" }),
      groups: groups.map((group) => ({
        ...group,
        axes: { category: "category", value: "value" },
      })),
    });
    const seriesSlide = strFromU8(
      unzipSync(seriesMotion.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(seriesSlide, /seriesIdx="0"/u);
    assert.match(seriesSlide, /seriesIdx="1"/u);
    assert.match(seriesSlide, /<a:bldChart bld="series">/u);

    const categoryMotion = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value" },
      },
      enter: fade({ by: "category" }),
      groups: groups.map((group) => ({
        ...group,
        axes: { category: "category", value: "value" },
      })),
    });
    const categorySlide = strFromU8(
      unzipSync(categoryMotion.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(categorySlide, /<a:bldChart bld="category">/u);
    assert.match(categorySlide, /categoryIdx="0"/u);
    assert.match(categorySlide, /categoryIdx="2"/u);
  });

  it("rejects shared axis sides and reads a self-crossing axis as invalid", async () => {
    await assert.rejects(
      buildSingleChart({
        axes: {
          firstCategory: { kind: "category", position: "top" },
          firstValue: { kind: "value" },
          secondCategory: { kind: "category", position: "top" },
          secondValue: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "firstCategory", value: "firstValue" },
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "First", values: [1] }],
          },
          {
            kind: "column",
            axes: { category: "secondCategory", value: "secondValue" },
            data: { kind: "shared-categories", categories: ["B"] },
            series: [{ name: "Second", values: [2] }],
          },
        ],
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-chart-axes");
        return true;
      },
    );

    const result = await buildSingleChart({
      axes: {
        category: { kind: "category" },
        value: { kind: "value" },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A"] },
          series: [{ name: "S", values: [1] }],
        },
      ],
    });
    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(parts[chartPath]!);
    const start = xml.indexOf('<c:valAx><c:axId val="222222"/>');
    assert.ok(start >= 0);
    const end = xml.indexOf("</c:valAx>", start);
    const block = xml.slice(start, end);
    parts[chartPath] = strToU8(
      xml.slice(0, start) +
        block.replace(
          '<c:crossAx val="111111"/>',
          '<c:crossAx val="222222"/>',
        ) +
        xml.slice(end),
    );
    const file = await validatePptx(zipSync(parts));
    assert.ok(
      file.errors.some(({ code }) => code === "invalid-chart-axis"),
      JSON.stringify(file.errors),
    );
  });

  it("writes same-kind bar dual axes with vertical categories and horizontal values", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 880,
      height: 480,
      axes: {
        primaryCategory: {
          kind: "category",
          position: "left",
          title: "Primary category",
        },
        primaryValue: {
          kind: "value",
          position: "bottom",
          min: 0,
          max: 20,
        },
        secondaryCategory: {
          kind: "category",
          position: "right",
          title: "Secondary category",
        },
        secondaryValue: {
          kind: "value",
          position: "top",
          min: 0,
          max: 100,
          numberFormat: "0%",
        },
      },
      groups: [
        {
          kind: "bar",
          axes: {
            category: "primaryCategory",
            value: "primaryValue",
          },
          data: { kind: "shared-categories", categories: ["North", "South"] },
          series: [{ name: "Primary", values: [6, 14] }],
        },
        {
          kind: "bar",
          axes: {
            category: "secondaryCategory",
            value: "secondaryValue",
          },
          data: { kind: "shared-categories", categories: ["East", "West"] },
          series: [{ name: "Secondary", values: [45, 80] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(chart.types, ["barChart", "barChart"]);
    assert.deepEqual(
      chart.groups.map(({ direction, axisIds }) => [direction, axisIds]),
      [
        ["bar", [111111, 222222]],
        ["bar", [333333, 444444]],
      ],
    );
    assert.deepEqual(
      chart.axes.map(({ id, kind, position, crossAxisId }) => ({
        id,
        kind,
        position,
        crossAxisId,
      })),
      [
        {
          id: 111111,
          kind: "category",
          position: "left",
          crossAxisId: 222222,
        },
        {
          id: 222222,
          kind: "value",
          position: "bottom",
          crossAxisId: 111111,
        },
        {
          id: 333333,
          kind: "category",
          position: "right",
          crossAxisId: 444444,
        },
        {
          id: 444444,
          kind: "value",
          position: "top",
          crossAxisId: 333333,
        },
      ],
    );
    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:catAx>/gu) ?? []).length, 2);
    assert.equal((xml.match(/<c:valAx>/gu) ?? []).length, 2);
    assert.match(xml, /<c:axPos val="l"\/>/u);
    assert.match(xml, /<c:axPos val="r"\/>/u);
    assert.match(xml, /<c:axPos val="b"\/>/u);
    assert.match(xml, /<c:axPos val="t"\/>/u);
  });

  it("maps a single value-axis position to the horizontal crossing control", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value", position: "right", min: 0, max: 10 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [2, 8] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(
      chart.axes.map(({ kind, position, crosses }) => [
        kind,
        position,
        crosses,
      ]),
      [
        ["category", "bottom", "max"],
        ["value", "right", undefined],
      ],
    );
    const xml = chartXml(result.pptx);
    assert.match(xml, /<c:crosses val="max"\/>/u);
    assert.match(xml, /<c:crosses val="autoZero"\/>/u);
  });

  it("preserves explicit crossing controls while deriving their physical sides", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category", crosses: "max" },
        value: { kind: "value", crossesAt: 1, min: 1, max: 100 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [10, 90] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(
      chart.axes.map(({ kind, position, crosses, crossesAt }) => ({
        kind,
        position,
        crosses,
        crossesAt,
      })),
      [
        {
          kind: "category",
          position: "bottom",
          crosses: "max",
          crossesAt: undefined,
        },
        { kind: "value", position: "right", crosses: undefined, crossesAt: 1 },
      ],
    );
    const xml = chartXml(result.pptx);
    assert.match(xml, /<c:crosses val="max"\/>/u);
    assert.match(xml, /<c:crossesAt val="1"\/>/u);
  });

  it("rejects conflicting sides on shared numeric axes", async () => {
    const categoryData = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const rejects = async (props: ChartProps) => {
      await assert.rejects(buildSingleChart(props), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-chart-axes");
        return true;
      });
    };
    await rejects({
      axes: {
        period: { kind: "category", position: "bottom" },
        value: { kind: "value" },
        scatterX: { kind: "value", position: "top" },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "period", value: "value" },
          data: categoryData,
          series: [{ name: "Bars", values: [1, 2] }],
        },
        {
          kind: "scatter",
          axes: { x: "scatterX", y: "value" },
          data: { kind: "xy" },
          series: [{ name: "Dots", x: [1, 2], y: [3, 4] }],
        },
      ],
    });
    await rejects({
      axes: {
        category: { kind: "category", position: "left" },
        barValue: { kind: "value" },
        scatterY: { kind: "value", position: "right" },
      },
      groups: [
        {
          kind: "bar",
          axes: { category: "category", value: "barValue" },
          data: categoryData,
          series: [{ name: "Bars", values: [1, 2] }],
        },
        {
          kind: "scatter",
          axes: { x: "barValue", y: "scatterY" },
          data: { kind: "xy" },
          series: [{ name: "Dots", x: [1, 2], y: [3, 4] }],
        },
      ],
    });
  });

  it("reads a crossing and axis-position mismatch as a validator error", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value", position: "right", min: 0, max: 10 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [2, 8] }],
        },
      ],
    });
    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(parts[chartPath]!);
    parts[chartPath] = strToU8(
      xml.replace('<c:crosses val="max"/>', '<c:crosses val="min"/>'),
    );
    const file = await validatePptx(zipSync(parts));
    assert.ok(
      file.errors.some(({ code }) => code === "invalid-chart-axis"),
      JSON.stringify(file.errors),
    );
  });

  it("diagnoses an explicit max crossing that conflicts with a reversed position", async () => {
    await assert.rejects(
      buildSingleChart({
        x: 10,
        y: 10,
        width: 640,
        height: 360,
        axes: {
          category: { kind: "category", position: "top" },
          value: {
            kind: "value",
            position: "right",
            reversed: true,
            crosses: "max",
            min: 1,
            max: 100,
          },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "category", value: "value" },
            data: { kind: "shared-categories", categories: ["A", "B"] },
            series: [{ name: "S", values: [10, 90] }],
          },
        ],
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-chart-axes");
        return true;
      },
    );
  });

  it("maps an explicit left side to an endpoint crossing instead of autoZero", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        x: { kind: "value", min: -10, max: 10 },
        y: { kind: "value", position: "left", min: 1, max: 3 },
      },
      groups: [
        {
          kind: "scatter",
          axes: { x: "x", y: "y" },
          data: { kind: "xy" },
          line: "straight",
          series: [{ name: "S", x: [-5, 5], y: [1, 2] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.deepEqual(
      chart.axes.map(({ kind, position, crosses }) => [
        kind,
        position,
        crosses,
      ]),
      [
        ["value", "bottom", "min"],
        ["value", "left", undefined],
      ],
    );
    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c:crosses val="min"\/>/gu) ?? []).length, 1);
    assert.equal((xml.match(/<c:crosses val="autoZero"\/>/gu) ?? []).length, 1);
  });

  it("keeps autoZero for undeclared primary defaults", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        x: { kind: "value", min: -10, max: 10 },
        y: { kind: "value", min: 1, max: 3 },
      },
      groups: [
        {
          kind: "scatter",
          axes: { x: "x", y: "y" },
          data: { kind: "xy" },
          series: [{ name: "S", x: [-5, 5], y: [1, 2] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const xml = chartXml(result.pptx);
    assert.ok(!xml.includes('<c:crosses val="min"/>'));
    assert.ok(!xml.includes('<c:crosses val="max"/>'));
    assert.equal((xml.match(/<c:crosses val="autoZero"\/>/gu) ?? []).length, 2);
  });

  it("does not treat an unknown autoZero as an edge in the reader", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        value: { kind: "value", min: -10, max: 10 },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [-5, 5] }],
        },
      ],
    });
    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(parts[chartPath]!);
    parts[chartPath] = strToU8(
      xml.replace('<c:axPos val="b"/>', '<c:axPos val="t"/>'),
    );
    const file = await validatePptx(zipSync(parts));
    assert.ok(
      !file.errors.some(({ code }) => code === "invalid-chart-axis"),
      JSON.stringify(file.errors),
    );
  });

  it("reads a reversed endpoint crossing against the physical side", async () => {
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category", position: "bottom" },
        value: {
          kind: "value",
          position: "left",
          reversed: true,
          crosses: "max",
          min: 1,
          max: 100,
        },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "category", value: "value" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [10, 90] }],
        },
      ],
    });
    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(parts[chartPath]!);
    parts[chartPath] = strToU8(
      xml.replace('<c:axPos val="b"/>', '<c:axPos val="t"/>'),
    );
    const file = await validatePptx(zipSync(parts));
    assert.ok(
      file.errors.some(({ code }) => code === "invalid-chart-axis"),
      JSON.stringify(file.errors),
    );
  });

  it("projects a shared category axis for line-line groups", async () => {
    const categories = ["A", "A", "B"];
    const result = await buildSingleChart({
      x: 10,
      y: 10,
      width: 640,
      height: 360,
      axes: {
        category: { kind: "category" },
        low: { kind: "value", position: "left", min: 0, max: 10 },
        high: { kind: "value", position: "right", min: 0, max: 100 },
      },
      groups: [
        {
          kind: "line",
          axes: { category: "category", value: "low" },
          data: { kind: "shared-categories", categories },
          series: [{ name: "Low", values: [1, null, 3] }],
        },
        {
          kind: "line",
          axes: { category: "category", value: "high" },
          data: { kind: "shared-categories", categories },
          series: [{ name: "High", values: [10, 20, null] }],
        },
      ],
    });
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.ok(chart);
    assert.equal(chart.axes.length, 4);
    assert.deepEqual(
      chart.axes.map(({ kind, position }) => [kind, position]),
      [
        ["category", "bottom"],
        ["value", "left"],
        ["category", "bottom"],
        ["value", "right"],
      ],
    );
    assert.deepEqual(
      chart.groups.map(({ type, axisIds }) => [type, axisIds]),
      [
        ["lineChart", [111111, 222222]],
        ["lineChart", [333333, 444444]],
      ],
    );
    assert.deepEqual(
      chart.series.map(({ categories: domain, values }) => ({
        domain,
        values,
      })),
      [
        { domain: categories, values: [1, null, 3] },
        { domain: categories, values: [10, 20, null] },
      ],
    );
    assert.deepEqual(chart.categoryFormulas, ["Sheet1!$A$2:$A$4"]);
    assert.deepEqual(
      chart.embeddedWorkbook?.series.map(({ categories: domain, values }) => [
        domain,
        values,
      ]),
      [
        [categories, [1, null, 3]],
        [categories, [10, 20, null]],
      ],
    );
  });

  it("keeps the shared category domain under group reorder", async () => {
    const categories = ["A", "A", "B"];
    const buildOrder = async (order: readonly ("left" | "right")[]) =>
      buildSingleChart({
        x: 10,
        y: 10,
        width: 640,
        height: 360,
        axes: {
          category: { kind: "category" },
          left: { kind: "value", position: "left", min: 0, max: 10 },
          right: { kind: "value", position: "right", min: 0, max: 100 },
        },
        groups: order.map((side) => ({
          kind: "column" as const,
          axes: {
            category: "category",
            value: side,
          },
          data: { kind: "shared-categories" as const, categories },
          series: [
            {
              name: side,
              values: side === "left" ? [1, null, 3] : [10, 20, null],
            },
          ],
        })),
      });
    for (const order of [
      ["left", "right"],
      ["right", "left"],
    ] as const) {
      const result = await buildOrder(order);
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const chart = file.slides[0]?.objects[0]?.chart;
      assert.ok(chart);
      assert.equal(chart.axes.length, 4);
      assert.deepEqual(
        chart.groups.map(({ axisIds }) => axisIds.length),
        [2, 2],
      );
      assert.deepEqual(
        chart.series.map(({ name }) => name),
        order,
      );
      assert.deepEqual(chart.categoryFormulas, ["Sheet1!$A$2:$A$4"]);
      assert.deepEqual(chart.embeddedWorkbook?.categories, categories);
    }
  });
});
