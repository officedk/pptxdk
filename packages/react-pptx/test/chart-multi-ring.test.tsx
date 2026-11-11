import assert from "node:assert/strict";

import {
  BuildError,
  build,
  Chart,
  Deck,
  fade,
  Slide,
  type ChartProps,
} from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  chart as documentChart,
  deck,
  slide,
} from "@react-pptx/react-pptx/document";
import {
  validatePptx,
  type FileChart,
  type FileChartDataLabel,
} from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

const sharedCategories = ["Direct", "Partner", "Direct"] as const;

const sharedRings: ChartProps = {
  x: 0,
  y: 0,
  width: 720,
  height: 420,
  title: "Shared category rings",
  legend: "right",
  groups: [
    {
      kind: "doughnut",
      data: { kind: "shared-categories", categories: sharedCategories },
      holeSize: 55,
      series: [
        {
          name: "Actual",
          values: [12, null, 18],
          points: { 0: { color: "#C8102E" }, 2: { color: "#1D4E89" } },
        },
        {
          name: "Plan",
          values: [10, 15, 20],
          points: { 0: { color: "#C8102E" }, 2: { color: "#1D4E89" } },
        },
        {
          name: "Forecast",
          values: [11, 16, 19],
          points: { 0: { color: "#C8102E" }, 2: { color: "#1D4E89" } },
        },
      ],
      dataLabels: { value: true, category: true },
    },
  ],
};

const independentRings: ChartProps = {
  x: 0,
  y: 0,
  width: 720,
  height: 420,
  title: "Independent category rings",
  legend: "right",
  groups: [
    {
      kind: "doughnut",
      data: { kind: "independent-categories" },
      holeSize: 55,
      series: [
        {
          name: "Ring A",
          categories: ["Direct", "Partner", "Partner"],
          values: [30, 20, null],
          points: {
            0: { color: "#C8102E" },
            1: { color: "#2E7D32" },
          },
        },
        {
          name: "Ring B",
          categories: ["Partner", "Direct"],
          values: [40, 10],
          points: {
            0: { color: "#1565C0" },
            1: { color: "#EF6C00" },
          },
        },
        {
          name: "Ring C",
          categories: ["Direct", "Partner", "Online", "Partner"],
          values: [5, 6, 7, 8],
          points: {
            0: { color: "#6A1B9A" },
            3: { color: "#00838F" },
          },
        },
      ],
      dataLabels: { value: true, category: true },
    },
  ],
};

function chartFrom(
  file: Awaited<ReturnType<typeof validatePptx>>,
  index = 0,
): FileChart {
  const chart = file.slides[0]?.objects[index]?.chart;
  assert.ok(chart);
  return chart;
}

function labelSettings(
  value: false | FileChartDataLabel | undefined,
): FileChartDataLabel | undefined {
  return value === false ? undefined : value;
}

function chartXml(pptx: Uint8Array, chartNumber: number): string {
  const parts = unzipSync(pptx);
  return strFromU8(parts[`ppt/charts/chart${chartNumber}.xml`]!);
}

function workbookSheet(pptx: Uint8Array, chart: FileChart): string {
  const part = chart.embeddedWorkbook?.part;
  assert.ok(part);
  const workbook = unzipSync(unzipSync(pptx)[part]!);
  return strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
}

describe("multi-ring doughnut charts", () => {
  it("keeps one shared category domain across multiple native rings and their workbook columns", async () => {
    const result = await build(
      <Deck>
        <Slide id="shared-rings">
          <Chart id="shared-rings" {...sharedRings} />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = chartFrom(file);
    assert.deepEqual(chart.types, ["doughnutChart"]);
    assert.deepEqual(
      chart.series.map(({ name, categories, values, index, order }) => ({
        name,
        categories,
        values,
        index,
        order,
      })),
      [
        {
          name: "Actual",
          categories: [...sharedCategories],
          values: [12, null, 18],
          index: 0,
          order: 0,
        },
        {
          name: "Plan",
          categories: [...sharedCategories],
          values: [10, 15, 20],
          index: 1,
          order: 1,
        },
        {
          name: "Forecast",
          categories: [...sharedCategories],
          values: [11, 16, 19],
          index: 2,
          order: 2,
        },
      ],
    );
    assert.deepEqual(chart.groups[0]?.seriesIndices, [0, 1, 2]);
    assert.deepEqual(
      chart.embeddedWorkbook?.series.map(({ name, categories, values }) => ({
        name,
        categories,
        values,
      })),
      chart.series.map(({ name, categories, values }) => ({
        name,
        categories,
        values,
      })),
    );

    const xml = chartXml(result.pptx, 1);
    assert.equal((xml.match(/<c:ser>/gu) ?? []).length, 3);
    assert.equal(
      (xml.match(/\$A\$2:\$A\$4/gu) ?? []).length,
      3,
      "every ring must reference the one shared category range",
    );
    for (const column of ["B", "C", "D"]) {
      assert.ok(xml.includes(`$${column}$2:$${column}$4`), column);
    }
    assert.ok(xml.includes('<c:legendPos val="r"/>'));
    assert.ok(xml.includes('<c:varyColors val="1"/>'));
    assert.equal((xml.match(/val="C8102E"/gu) ?? []).length, 3);
    assert.equal((xml.match(/val="1D4E89"/gu) ?? []).length, 3);
    assert.equal((xml.match(/<c:dPt>/gu) ?? []).length, 6);

    const sheet = workbookSheet(result.pptx, chart);
    assert.ok(
      sheet.includes('<c r="A2" t="inlineStr"><is><t>Direct</t></is></c>'),
    );
    assert.ok(
      sheet.includes('<c r="A3" t="inlineStr"><is><t>Partner</t></is></c>'),
    );
    assert.ok(
      sheet.includes('<c r="A4" t="inlineStr"><is><t>Direct</t></is></c>'),
    );
    assert.ok(!sheet.includes('r="B3"'), "the null slot must stay empty");
    assert.ok(sheet.includes('r="B2"'));
    assert.ok(sheet.includes('r="C4"'));
    assert.ok(sheet.includes('r="D4"'));
  });

  it("keeps independent categories, point counts, labels, and explicit point colors per ring through the command entry", async () => {
    const result = await buildDocument(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "independent-rings" }).append(
          documentChart({
            ...independentRings,
            id: "independent-rings",
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = chartFrom(file);
    assert.deepEqual(chart.types, ["doughnutChart"]);
    assert.deepEqual(
      chart.series.map(
        ({ name, categories, values, index, order, groupIndex }) => ({
          name,
          categories,
          values,
          index,
          order,
          groupIndex,
        }),
      ),
      [
        {
          name: "Ring A",
          categories: ["Direct", "Partner", "Partner"],
          values: [30, 20, null],
          index: 0,
          order: 0,
          groupIndex: 0,
        },
        {
          name: "Ring B",
          categories: ["Partner", "Direct"],
          values: [40, 10],
          index: 1,
          order: 1,
          groupIndex: 0,
        },
        {
          name: "Ring C",
          categories: ["Direct", "Partner", "Online", "Partner"],
          values: [5, 6, 7, 8],
          index: 2,
          order: 2,
          groupIndex: 0,
        },
      ],
    );
    assert.deepEqual(chart.groups[0]?.seriesIndices, [0, 1, 2]);
    assert.equal(chart.seriesDataLabels?.length, 3);
    for (const labels of chart.seriesDataLabels ?? []) {
      assert.deepEqual(labels.dataLabels, {
        value: true,
        category: true,
        percent: false,
      });
    }
    assert.equal(chart.embeddedWorkbook?.series.length, 3);
    for (const [index, workbookSeries] of (
      chart.embeddedWorkbook?.series ?? []
    ).entries()) {
      assert.deepEqual(
        workbookSeries.categories,
        chart.series[index]?.categories,
      );
      assert.deepEqual(workbookSeries.values, chart.series[index]?.values);
    }

    const xml = chartXml(result.pptx, 1);
    assert.equal((xml.match(/<c:ser>/gu) ?? []).length, 3);
    assert.ok(xml.includes("Sheet1!$A$2:$A$4"));
    assert.ok(xml.includes("Sheet1!$B$2:$B$4"));
    assert.ok(xml.includes("Sheet1!$C$2:$C$3"));
    assert.ok(xml.includes("Sheet1!$D$2:$D$3"));
    assert.ok(xml.includes("Sheet1!$E$2:$E$5"));
    assert.ok(xml.includes("Sheet1!$F$2:$F$5"));
    // The same label "Partner" is at different positions with different explicit colors.
    for (const color of [
      "C8102E",
      "2E7D32",
      "1565C0",
      "EF6C00",
      "6A1B9A",
      "00838F",
    ]) {
      assert.ok(xml.includes(`<a:srgbClr val="${color}"/>`), color);
    }
    assert.ok(!xml.includes('<c:v val="undefined"/>'));
    assert.ok(!xml.includes('<c:v val="NaN"/>'));

    const sheet = workbookSheet(result.pptx, chart);
    assert.ok(sheet.includes('r="A2"'));
    assert.ok(sheet.includes('r="C2"'));
    assert.ok(sheet.includes('r="D3"'));
    assert.ok(sheet.includes('r="E5"'));
    assert.ok(!sheet.includes('r="B4"'), "the null slot must stay empty");
    assert.ok(!sheet.includes('r="E6"'));
    assert.ok(!sheet.includes('r="F6"'));
  });

  it("feeds real ring and point counts into the existing chart animation mapping", async () => {
    const result = await build(
      <Deck>
        <Slide id="multi-ring-motion">
          <Chart
            id="series-motion"
            x={0}
            y={0}
            width={360}
            height={240}
            enter={fade({ by: "series", duration: 100 })}
            groups={independentRings.groups ?? []}
          />
          <Chart
            id="category-motion"
            x={0}
            y={260}
            width={360}
            height={240}
            enter={fade({ by: "category", duration: 100 })}
            groups={independentRings.groups ?? []}
          />
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.["motion"] as
      | {
          steps: {
            effects: {
              target: {
                authorId?: string;
                range?: {
                  step: string;
                  seriesIndex: number;
                  categoryIndex: number;
                };
              };
            }[];
          }[];
        }
      | undefined;
    const effects = motion?.steps[0]?.effects ?? [];
    const seriesRanges = effects
      .filter(({ target }) => target.authorId === "series-motion")
      .map(({ target }) => target.range)
      .filter((range) => range?.step === "series");
    const categoryRanges = effects
      .filter(({ target }) => target.authorId === "category-motion")
      .map(({ target }) => target.range)
      .filter((range) => range?.step === "category");
    assert.deepEqual(
      seriesRanges.map((range) => range?.seriesIndex),
      [0, 1, 2],
    );
    assert.deepEqual(
      categoryRanges.map((range) => range?.categoryIndex),
      [0, 1, 2, 3],
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
  });

  it("applies 17 series and point formatting per independent ring", async () => {
    const result = await build(
      <Deck>
        <Slide id="ring-format">
          <Chart
            id="ring-format"
            x={0}
            y={0}
            width={640}
            height={420}
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                holeSize: 60,
                series: [
                  {
                    name: "Ring A",
                    categories: ["Direct", "Partner", "Online"],
                    values: [10, 20, 30],
                    explosion: 12,
                    dataLabels: { value: true, leaderLines: false },
                    points: {
                      1: {
                        color: "#AA0000",
                        explosion: 25,
                        dataLabels: { numberFormat: "0.0%" },
                      },
                      2: { dataLabels: false },
                    },
                  },
                  {
                    name: "Ring B",
                    categories: ["Direct", "Partner"],
                    values: [5, null],
                    dataLabels: false,
                  },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = chartFrom(file);
    assert.equal(chart.series[0]?.explosion, 12);
    assert.deepEqual(chart.series[0]?.pointExplosions, { 1: 25 });
    assert.equal(chart.series[1]?.explosion, undefined);

    const ringALabels = chart.seriesDataLabels?.[0];
    const ringASettings = labelSettings(ringALabels?.dataLabels);
    assert.equal(ringASettings?.value, true);
    assert.equal(ringASettings?.leaderLines, false);
    assert.equal(
      labelSettings(ringALabels?.pointDataLabels?.[1])?.numberFormat,
      "0.0%",
    );
    assert.equal(ringALabels?.pointDataLabels?.[2], false);
    assert.equal(chart.seriesDataLabels?.[1]?.dataLabels, false);

    const xml = chartXml(result.pptx, 1);
    assert.equal((xml.match(/<c:ser>/gu) ?? []).length, 2);
    assert.ok(xml.includes('<c:explosion val="12"/>'));
    assert.ok(xml.includes('<c:explosion val="25"/>'));
    assert.ok(xml.includes('<a:srgbClr val="AA0000"/>'));
    assert.ok(xml.includes('<c:showLeaderLines val="0"/>'));
    assert.ok(!xml.includes("<c:leaderLines"));
    assert.ok(
      xml.includes("Sheet1!$C$2:$C$3") && xml.includes("Sheet1!$D$2:$D$3"),
    );

    for (const [id, code, series] of [
      [
        "ring-explosion-overflow",
        "invalid-chart-option",
        {
          name: "S",
          categories: ["A"],
          values: [1],
          points: { 0: { explosion: 2_147_483_648 } },
        },
      ],
      [
        "ring-leader-lines-without-labels",
        "invalid-data-labels",
        {
          name: "S",
          categories: ["A"],
          values: [1],
          dataLabels: { leaderLines: true },
        },
      ],
    ] as const) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id={id}>
              <Chart
                id={id}
                x={0}
                y={0}
                width={200}
                height={120}
                groups={[
                  {
                    kind: "doughnut",
                    data: { kind: "independent-categories" },
                    series: [series],
                  },
                ]}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          return true;
        },
      );
    }
  });

  it("rejects independent categories outside doughnut and malformed per-ring domains", async () => {
    const base = { x: 0, y: 0, width: 200, height: 120 } as const;
    const invalidCases: {
      readonly id: string;
      readonly code?: string;
      readonly props: ChartProps;
    }[] = [
      {
        id: "pie-independent",
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data: { kind: "independent-categories" },
              series: [{ name: "S", categories: ["A"], values: [1] }],
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "column-independent",
        props: {
          ...base,
          groups: [
            {
              kind: "column",
              data: { kind: "independent-categories" },
              series: [{ name: "S", categories: ["A"], values: [1] }],
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "missing-ring-categories",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [{ name: "S", values: [1] }],
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "ring-category-count-mismatch",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [{ name: "S", categories: ["A"], values: [1, 2] }],
            },
          ],
        },
      },
      {
        id: "group-level-categories",
        code: "unknown-parameter",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: {
                kind: "independent-categories",
                categories: ["A"],
              },
              series: [{ name: "S", categories: ["A"], values: [1] }],
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "ring-point-index",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [
                {
                  name: "S",
                  categories: ["A"],
                  values: [1],
                  points: { 1: { color: "#FF0000" } },
                },
              ],
            },
          ],
        },
      },
      {
        id: "shared-ring-categories",
        code: "unknown-parameter",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", categories: ["A"], values: [1] }],
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "ring-non-finite",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [
                {
                  name: "S",
                  categories: ["A"],
                  values: [Number.POSITIVE_INFINITY],
                },
              ],
            },
          ],
        },
      },
    ];
    for (const { id, code = "invalid-data", props } of invalidCases) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id={id}>
              <Chart id={id} {...props} />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          return true;
        },
      );
    }
  });
});
