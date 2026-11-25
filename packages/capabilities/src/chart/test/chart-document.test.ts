import assert from "node:assert/strict";

import type { ChartSeries } from "@pptxdk/capabilities/chart";
import { chart, deck, slide } from "@pptxdk/capabilities/document";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { validatePptx, type FileChartSeries } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { build } from "../../../test/test-helpers.ts";

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function chartObject(
  result: Awaited<ReturnType<typeof validatePptx>>,
  index: number,
) {
  const object = result.slides[0]?.objects[index];
  assert.equal(object?.kind, "chart");
  assert.ok(object.chart !== undefined);
  return object.chart;
}

function fileSeries(
  categories: readonly string[],
  series: readonly ChartSeries[],
  groupIndex = 0,
): FileChartSeries[] {
  return series.map((item, index) => ({
    name: item.name,
    values: item.values,
    categories,
    index,
    order: index,
    groupIndex,
  }));
}

const relationshipNamespace =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * Rewrites the workbook emitted by the formal chart writer into the shared-string
 * representation PowerPoint uses for embedded chart workbooks, with one
 * rich-text shared string. The workbook stays a real OOXML package; only the
 * storage encoding of string cells changes.
 */
function sharedStringWorkbook(
  pptx: Uint8Array,
  mutate?: (sheetXml: string) => string,
): Uint8Array {
  const parts = unzipSync(pptx);
  const workbookPart = Object.keys(parts).find((name) =>
    /^ppt\/embeddings\/.*\.xlsx$/u.test(name),
  );
  assert.ok(workbookPart !== undefined);
  const workbook = unzipSync(parts[workbookPart]!);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const originalSheet = strFromU8(workbook[sheetPath]!);
  const sheet = mutate === undefined ? originalSheet : mutate(originalSheet);
  const shared: string[] = [];
  const sharedIndex = (value: string): number => {
    const index = shared.indexOf(value);
    if (index !== -1) return index;
    shared.push(value);
    return shared.length - 1;
  };
  const converted = sheet.replace(
    /<c r="([A-Z]+\d+)" t="inlineStr"><is><t(?: xml:space="preserve")?>([^<]*)<\/t><\/is><\/c>/gu,
    (_match, reference: string, value: string) =>
      `<c r="${reference}" t="s"><v>${sharedIndex(value)}</v></c>`,
  );
  assert.notEqual(converted, sheet);
  workbook[sheetPath] = strToU8(converted);
  const sharedStrings = shared
    .map((value) =>
      value === "Q1"
        ? "<si><r><rPr><b/></rPr><t>Q</t></r><r><t>1</t></r></si>"
        : `<si><t xml:space="preserve">${value}</t></si>`,
    )
    .join("");
  workbook["xl/sharedStrings.xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${shared.length}" uniqueCount="${shared.length}">${sharedStrings}</sst>`,
  );
  workbook["[Content_Types].xml"] = strToU8(
    strFromU8(workbook["[Content_Types].xml"]!).replace(
      "</Types>",
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    ),
  );
  workbook["xl/_rels/workbook.xml.rels"] = strToU8(
    strFromU8(workbook["xl/_rels/workbook.xml.rels"]!).replace(
      "</Relationships>",
      `<Relationship Id="rId3" Type="${relationshipNamespace}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
    ),
  );
  parts[workbookPart] = zipSync(workbook);
  return zipSync(parts);
}

describe("chart authoring through the common document", () => {
  it("writes the six established chart kinds with group-owned series and matching native caches", async () => {
    const cartesianCategories = ["Q1", "Q2", "Q3"];
    const cartesianSeries: ChartSeries[] = [
      { name: "Alpha", values: [10, 20, 30] },
      { name: "Beta", values: [5, null, 15] },
    ];
    const shareCategories = ["Direct", "Partner"];
    const shareSeries: ChartSeries[] = [{ name: "Share", values: [3, 2] }];
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "charts" }).append(
        chart({
          id: "column",
          x: 0,
          y: 0,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "column",
              data: {
                kind: "shared-categories",
                categories: cartesianCategories,
              },
              series: cartesianSeries,
            },
          ],
        }),
        chart({
          id: "bar",
          x: 320,
          y: 0,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "bar",
              stacking: "stacked",
              data: {
                kind: "shared-categories",
                categories: cartesianCategories,
              },
              series: cartesianSeries,
            },
          ],
        }),
        chart({
          id: "line",
          x: 0,
          y: 220,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "line",
              data: {
                kind: "shared-categories",
                categories: cartesianCategories,
              },
              series: cartesianSeries,
            },
          ],
        }),
        chart({
          id: "area",
          x: 320,
          y: 220,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "area",
              data: {
                kind: "shared-categories",
                categories: cartesianCategories,
              },
              series: cartesianSeries,
            },
          ],
        }),
        chart({
          id: "pie",
          x: 640,
          y: 0,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "pie",
              data: {
                kind: "shared-categories",
                categories: shareCategories,
              },
              series: shareSeries,
              dataLabels: { value: true, category: true },
            },
          ],
        }),
        chart({
          id: "doughnut",
          x: 640,
          y: 220,
          width: 300,
          height: 200,
          groups: [
            {
              kind: "doughnut",
              holeSize: 55,
              data: {
                kind: "shared-categories",
                categories: shareCategories,
              },
              series: shareSeries,
            },
          ],
        }),
      ),
    );

    const result = await build(document);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const expected = [
      ["barChart", "col", cartesianCategories, cartesianSeries],
      ["barChart", "bar", cartesianCategories, cartesianSeries],
      ["lineChart", undefined, cartesianCategories, cartesianSeries],
      ["areaChart", undefined, cartesianCategories, cartesianSeries],
      ["pieChart", undefined, shareCategories, shareSeries],
      ["doughnutChart", undefined, shareCategories, shareSeries],
    ] as const;
    assert.equal(file.slides[0]?.objects.length, expected.length);
    expected.forEach(([type, direction, categories, chartSeries], index) => {
      const value = chartObject(file, index);
      const series = fileSeries(categories, chartSeries);
      assert.deepEqual(value.types, [type]);
      assert.equal(value.direction, direction);
      assert.deepEqual(value.series, series);
      assert.deepEqual(value.groups, [
        {
          type,
          ...(direction === undefined ? {} : { direction }),
          ...(index === 0 || index === 1
            ? { grouping: index === 0 ? "clustered" : "stacked" }
            : index === 2 || index === 3
              ? { grouping: "standard" }
              : {}),
          axisIds:
            type === "pieChart" || type === "doughnutChart"
              ? []
              : [111111, 222222],
          seriesIndices: series.map(({ index: seriesIndex }) => seriesIndex),
          ...(type === "pieChart" ? { firstSliceAngle: 0 } : {}),
          ...(type === "doughnutChart"
            ? { firstSliceAngle: 0, holeSize: 55 }
            : {}),
        },
      ]);
      const horizontalBar = direction === "bar";
      assert.deepEqual(
        value.axes,
        type === "pieChart" || type === "doughnutChart"
          ? []
          : [
              {
                id: 111111,
                kind: "category",
                crossAxisId: 222222,
                position: horizontalBar ? "left" : "bottom",
                visible: true,
              },
              {
                id: 222222,
                kind: "value",
                crossAxisId: 111111,
                position: horizontalBar ? "bottom" : "left",
                visible: true,
              },
            ],
      );
      assert.deepEqual(value.embeddedWorkbook?.categories, categories);
      assert.deepEqual(value.embeddedWorkbook?.series, series);
    });
    assert.equal(
      (slideXml(result).match(/<p:graphicFrame>/gu) ?? []).length,
      6,
    );
    assert.equal(
      Object.keys(unzipSync(result.pptx)).filter((path) =>
        /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/u.test(path),
      ).length,
      6,
    );
  });

  it("preserves duplicate category positions, group order, and explicit axis references", async () => {
    const categories = ["Q1", "Q1", "Q3"];
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "group-order" }).append(
        chart({
          id: "group-order-chart",
          x: 20,
          y: 20,
          width: 600,
          height: 360,
          axes: {
            period: { kind: "category" },
            amount: { kind: "value", max: 40, majorUnit: 10 },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "period", value: "amount" },
              data: { kind: "shared-categories", categories },
              series: [
                {
                  name: "North",
                  values: [14, null, 18],
                  points: { 0: { labelColor: "#AA0000" } },
                },
              ],
              dataLabels: {
                value: true,
                style: { color: "#112233" },
              },
            },
            {
              kind: "column",
              axes: { category: "period", value: "amount" },
              data: { kind: "shared-categories", categories },
              series: [
                {
                  name: "South",
                  values: [21, 28, 35],
                  points: { 0: { labelColor: "#00AA00" } },
                },
              ],
              dataLabels: {
                value: true,
                style: { color: "#112233" },
              },
            },
          ],
        }),
      ),
    );

    const result = await build(document);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const value = chartObject(file, 0);
    assert.deepEqual(value.types, ["barChart", "barChart"]);
    assert.deepEqual(
      value.groups.map(({ axisIds, seriesIndices }) => ({
        axisIds,
        seriesIndices,
      })),
      [
        { axisIds: [111111, 222222], seriesIndices: [0] },
        { axisIds: [111111, 222222], seriesIndices: [1] },
      ],
    );
    assert.deepEqual(
      value.series.map(
        ({ name, index, order, groupIndex, categories: domain }) => ({
          name,
          index,
          order,
          groupIndex,
          domain,
        }),
      ),
      [
        {
          name: "North",
          index: 0,
          order: 0,
          groupIndex: 0,
          domain: ["Q1", "Q1", "Q3"],
        },
        {
          name: "South",
          index: 1,
          order: 1,
          groupIndex: 1,
          domain: ["Q1", "Q1", "Q3"],
        },
      ],
    );
    assert.deepEqual(
      value.series.map(({ values }) => values),
      [
        [14, null, 18],
        [21, 28, 35],
      ],
    );
    assert.deepEqual(value.embeddedWorkbook?.series, value.series);

    const chartPath = "ppt/charts/chart1.xml";
    const chartXml = strFromU8(unzipSync(result.pptx)[chartPath]!);
    assert.match(chartXml, /<c:ser><c:idx val="0"\/><c:order val="0"\/>/u);
    assert.match(chartXml, /<c:ser><c:idx val="1"\/><c:order val="1"\/>/u);
    assert.match(chartXml, /Sheet1!\$B\$2:\$B\$4/u);
    assert.match(chartXml, /Sheet1!\$C\$2:\$C\$4/u);
    assert.match(chartXml, /<c:ptCount val="3"\/>/u);
    assert.doesNotMatch(chartXml, /<c:pt idx="1"><c:v>0<\/c:v>/u);
    assert.match(chartXml, /<c:axId val="111111"\/>/u);
    assert.match(chartXml, /<c:axId val="222222"\/>/u);
    assert.equal(
      (chartXml.match(/<c:dLbl><c:idx val="0"\/>/gu) ?? []).length,
      2,
    );
    assert.match(chartXml, /<a:srgbClr val="AA0000"\/>/u);
    assert.match(chartXml, /<a:srgbClr val="00AA00"\/>/u);

    const workbookPart = value.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const worksheet = strFromU8(
      unzipSync(unzipSync(result.pptx)[workbookPart]!)[
        "xl/worksheets/sheet1.xml"
      ]!,
    );
    assert.match(
      worksheet,
      /<c r="A3" t="inlineStr"><is><t>Q1<\/t><\/is><\/c>/u,
    );
    assert.match(worksheet, /<c r="B2"><v>14<\/v><\/c>/u);
    assert.doesNotMatch(worksheet, /<c r="B3">/u);
    assert.match(worksheet, /<c r="C4"><v>35<\/v><\/c>/u);
  });

  it("resolves chart data from formulas and the workbook rather than a fixed A/B rectangle", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [121, 135] },
        { name: "South", values: [90, 105] },
      ],
    };
    const result = await build(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "shared-strings" }).append(
          chart({
            id: "shared-strings-chart",
            x: 0,
            y: 0,
            width: 300,
            height: 200,
            groups: [
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: data.categories,
                },
                series: data.series,
              },
            ],
          }),
        ),
      ),
    );

    const file = await validatePptx(sharedStringWorkbook(result.pptx));
    assert.deepEqual(file.errors, []);
    const chartValue = chartObject(file, 0);
    const expectedSeries = fileSeries(data.categories, data.series);
    assert.deepEqual(chartValue.series, expectedSeries);
    assert.deepEqual(chartValue.embeddedWorkbook?.categories, data.categories);
    assert.deepEqual(chartValue.embeddedWorkbook?.series, expectedSeries);
  });

  it("reads data through actual formulas instead of a fixed workbook rectangle", async () => {
    const categories = ["Q1", "Q2"];
    const series = [
      { name: "North", values: [121, 135] },
      { name: "South", values: [90, 105] },
    ];
    const result = await build(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "formula-layout" }).append(
          chart({
            id: "formula-layout-chart",
            x: 0,
            y: 0,
            width: 300,
            height: 200,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories },
                series,
              },
            ],
          }),
        ),
      ),
    );

    const parts = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    let chartXml = strFromU8(parts[chartPath]!);
    for (const [from, to] of [
      ["Sheet1!$A$2:$A$3", "Sheet1!$D$2:$D$3"],
      ["Sheet1!$B$1", "Sheet1!$E$1"],
      ["Sheet1!$B$2:$B$3", "Sheet1!$E$2:$E$3"],
      ["Sheet1!$C$1", "Sheet1!$F$1"],
      ["Sheet1!$C$2:$C$3", "Sheet1!$F$2:$F$3"],
    ] as const) {
      assert.ok(chartXml.includes(from), from);
      chartXml = chartXml.replaceAll(from, to);
    }
    parts[chartPath] = strToU8(chartXml);

    const workbookPart = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const workbook = unzipSync(parts[workbookPart]!);
    const sheetPath = "xl/worksheets/sheet1.xml";
    const originalSheet = strFromU8(workbook[sheetPath]!);
    const movedSheet = originalSheet.replace(
      /<c r="([ABC])(\d+)"/gu,
      (_match, column: string, row: string) =>
        `<c r="${({ A: "D", B: "E", C: "F" } as const)[column as "A" | "B" | "C"]}${row}"`,
    );
    assert.notEqual(movedSheet, originalSheet);
    workbook[sheetPath] = strToU8(movedSheet);
    parts[workbookPart] = zipSync(workbook);

    const file = await validatePptx(zipSync(parts));
    assert.deepEqual(file.errors, []);
    const observed = chartObject(file, 0);
    assert.deepEqual(observed.series, fileSeries(categories, series));
    assert.deepEqual(observed.embeddedWorkbook?.series, observed.series);
  });

  it("keeps chart workbook mismatches after shared-string resolution", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [121, 135] },
        { name: "South", values: [90, 105] },
      ],
    };
    const result = await build(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "shared-strings-mismatch" }).append(
          chart({
            id: "shared-strings-mismatch-chart",
            x: 0,
            y: 0,
            width: 300,
            height: 200,
            groups: [
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: data.categories,
                },
                series: data.series,
              },
            ],
          }),
        ),
      ),
    );
    const rewritten = sharedStringWorkbook(result.pptx, (sheet) => {
      const mutated = sheet.replace(
        '<c r="B3"><v>135</v></c>',
        '<c r="B3"><v>999</v></c>',
      );
      assert.notEqual(mutated, sheet);
      return mutated;
    });

    const file = await validatePptx(rewritten);
    const mismatches = file.errors.filter(
      ({ code }) => code === "chart-workbook-mismatch",
    );
    assert.deepEqual(file.errors, mismatches);
    assert.deepEqual(
      mismatches.map(({ message }) => message),
      [
        "Chart series North does not match embedded workbook content",
        "Chart series North cache does not match workbook",
      ],
    );
    const chartValue = chartObject(file, 0);
    assert.deepEqual(
      chartValue.embeddedWorkbook?.series.map(({ values }) => values),
      [
        [121, 999],
        [90, 105],
      ],
    );
  });

  it("isolates group data by value and applies explicit updates only to later builds", async () => {
    const categories = ["Q1", "Q2"];
    const series = [{ name: "Original", values: [1, 2] }];
    const node = chart({
      id: "editable-chart",
      x: 0,
      y: 0,
      width: 300,
      height: 200,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories },
          series,
        },
      ],
    });
    assert.ok(node instanceof CommonNode);
    assert.ok(!(node instanceof CommonContainer));
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "chart-edit" }).append(node),
    );

    categories[0] = "Changed";
    series[0]!.name = "Mutated";
    series[0]!.values[0] = 99;

    const first = await build(document);
    const firstChart = chartObject(await validatePptx(first.pptx), 0);
    assert.deepEqual(
      firstChart.series.map(({ name, categories: domain, values }) => ({
        name,
        domain,
        values,
      })),
      [{ name: "Original", domain: ["Q1", "Q2"], values: [1, 2] }],
    );
    assert.deepEqual(firstChart.embeddedWorkbook?.series, firstChart.series);

    node.set({
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["R1"] },
          series: [{ name: "Revised", values: [7] }],
        },
      ],
    });
    const second = await build(document);
    const secondChart = chartObject(await validatePptx(second.pptx), 0);
    assert.deepEqual(
      secondChart.series.map(({ name, categories: domain, values }) => ({
        name,
        domain,
        values,
      })),
      [{ name: "Revised", domain: ["R1"], values: [7] }],
    );
    assert.deepEqual(secondChart.embeddedWorkbook?.series, secondChart.series);
    assert.deepEqual(
      firstChart.series.map(({ name, values }) => ({ name, values })),
      [{ name: "Original", values: [1, 2] }],
    );
    assert.doesNotMatch(slideXml(first), /Revised/u);
    assert.doesNotMatch(slideXml(first), /Changed/u);
  });

  it("rejects holes, undefined, non-finite values, unknown axes, and unsupported domains", async () => {
    const data = { kind: "shared-categories" as const, categories: ["A", "B"] };
    const invalidValues = [
      ["hole", Array<number>(2)],
      ["undefined", [1, undefined]],
      ["nan", [1, Number.NaN]],
      ["infinity", [1, Number.POSITIVE_INFINITY]],
    ] as const;
    for (const [label, values] of invalidValues) {
      await assert.rejects(
        build(
          deck({ width: 400, height: 300 }).append(
            slide({ id: `invalid-${label}` }).append(
              chart({
                id: `invalid-${label}`,
                width: 300,
                height: 200,
                groups: [
                  {
                    kind: "column",
                    data,
                    series: [
                      {
                        name: "S",
                        values: values as unknown as (number | null)[],
                      },
                    ],
                  },
                ],
              }),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-data");
          return true;
        },
      );
    }

    const diagnostics = [
      {
        id: "unknown-axis",
        code: "unknown-chart-axis",
        axes: {
          period: { kind: "category" as const },
          amount: { kind: "value" as const },
        },
        axisRefs: { category: "missing", value: "amount" },
      },
      {
        id: "wrong-axis-role",
        code: "invalid-chart-axis-role",
        axes: {
          period: { kind: "category" as const },
          amount: { kind: "value" as const },
        },
        axisRefs: { category: "amount", value: "period" },
      },
    ] as const;
    for (const item of diagnostics) {
      await assert.rejects(
        build(
          deck({ width: 400, height: 300 }).append(
            slide({ id: item.id }).append(
              chart({
                id: item.id,
                width: 300,
                height: 200,
                axes: item.axes,
                groups: [
                  {
                    kind: "column",
                    axes: item.axisRefs,
                    data,
                    series: [{ name: "S", values: [1, 2] }],
                  },
                ],
              }),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, item.code);
          return true;
        },
      );
    }

    await assert.rejects(
      build(
        deck({ width: 400, height: 300 }).append(
          slide({ id: "independent-domain" }).append(
            chart({
              id: "independent-domain",
              width: 300,
              height: 200,
              groups: [
                {
                  kind: "column",
                  data: {
                    kind: "independent-categories",
                  } as never,
                  series: [{ name: "S", values: [1] }],
                },
              ],
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-data");
        return true;
      },
    );

    await assert.rejects(
      build(
        deck({ width: 400, height: 300 }).append(
          slide({ id: "domain-mismatch" }).append(
            chart({
              id: "domain-mismatch",
              width: 300,
              height: 200,
              groups: [
                {
                  kind: "column",
                  data: { kind: "shared-categories", categories: ["A"] },
                  series: [{ name: "A", values: [1] }],
                },
                {
                  kind: "column",
                  data: { kind: "shared-categories", categories: ["B"] },
                  series: [{ name: "B", values: [2] }],
                },
              ],
            }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(
          error.diagnostics[0]?.code,
          "chart-category-domain-mismatch",
        );
        return true;
      },
    );
  });

  it("keeps chart registration explicit for trimmed builds", async () => {
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "trimmed" }).append(
        chart({
          width: 120,
          height: 80,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", values: [1] }],
            },
          ],
        }),
      ),
    );

    await assert.rejects(
      async () => emit(createResolver([])(document)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(({ code }) => code === "missing-capability"),
        );
        return true;
      },
    );
  });
});
