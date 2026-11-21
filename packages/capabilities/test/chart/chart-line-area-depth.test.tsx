import assert from "node:assert/strict";

import { type ChartProps } from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@pptxdk/capabilities/document";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../test-helpers.ts";

function chartXml(parts: Unzipped, index: number): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

async function rejectsWithCode(props: ChartProps, code: string): Promise<void> {
  const id = "invalid-line-area-chart";
  await assert.rejects(
    buildDocument(
      deck({ width: 240, height: 160 }).append(
        documentSlide({ id }).append(documentChart({ ...props, id })),
      ),
    ),
    (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, code);
      assert.match(error.diagnostics[0]?.path ?? "", new RegExp(id, "u"));
      return true;
    },
  );
}

describe("line and area depth", () => {
  it("writes native line markers, smooth flags, formats, and shared chart settings", async () => {
    const result = await build(
      <Deck width={640} height={400}>
        <Slide id="line-depth">
          <Chart
            id="line-depth"
            x={24}
            y={24}
            width={560}
            height={340}
            title="Quarterly trend"
            displayBlanksAs="span"
            legend="bottom"
            style={{ fontFamily: "Aptos", fontSize: "12pt", color: "#112233" }}
            axes={{
              category: {
                kind: "category",
                title: "Quarter",
                line: { color: "#999999", width: 1 },
              },
              value: {
                kind: "value",
                title: "Revenue",
                min: 0,
                max: 40,
                majorUnit: 10,
                numberFormat: "#,##0",
                gridlines: { color: "#DDDDDD", width: 1, dash: "dash" },
              },
            }}
            groups={[
              {
                kind: "line",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [
                  {
                    name: "Actual",
                    values: [10, null, 30],
                    color: "#336699",
                    line: { width: 2, dash: "dashDot" },
                    marker: {
                      symbol: "circle",
                      size: 7,
                      fill: "#FF6600",
                      line: { color: "#112233", width: 1.5, dash: "dash" },
                    },
                    smooth: true,
                  },
                  {
                    name: "Plan",
                    values: [12, 18, 25],
                    marker: true,
                    smooth: false,
                    points: {
                      1: {
                        marker: {
                          symbol: "star",
                          size: 12,
                          fill: "#00AA66",
                          line: { color: "#554433", width: 1, dash: "dot" },
                        },
                        line: { color: "#775533", width: 1, dash: "dot" },
                      },
                      2: { marker: false },
                    },
                  },
                ],
                dataLabels: {
                  value: true,
                  numberFormat: "0.0",
                },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.types, ["lineChart"]);
    assert.equal(chart?.grouping, "standard");
    assert.equal(chart?.title, "Quarterly trend");
    assert.equal(chart?.displayBlanksAs, "span");
    assert.deepEqual(
      chart?.series.map(({ name, values }) => [name, values]),
      [
        ["Actual", [10, null, 30]],
        ["Plan", [12, 18, 25]],
      ],
    );
    assert.deepEqual(
      chart?.series.map(({ marker, smooth }) => [marker, smooth]),
      [
        [{ symbol: "circle", size: 7 }, true],
        [{ symbol: "circle" }, undefined],
      ],
    );
    assert.deepEqual(chart?.embeddedWorkbook?.series, chart?.series);

    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.ok(xml.includes("<c:lineChart>"));
    assert.ok(xml.includes('<c:grouping val="standard"/>'));
    assert.ok(xml.includes('<c:marker val="1"/>'));
    assert.ok(xml.includes('<c:symbol val="circle"/>'));
    assert.ok(xml.includes('<c:size val="7"/>'));
    assert.ok(xml.includes('<c:symbol val="star"/>'));
    assert.ok(xml.includes('<c:size val="12"/>'));
    assert.ok(xml.includes('<c:smooth val="1"/>'));
    assert.ok(xml.includes('<c:smooth val="0"/>'));
    assert.ok(xml.includes('<a:srgbClr val="FF6600"/>'));
    assert.ok(xml.includes('<a:srgbClr val="00AA66"/>'));
    assert.ok(xml.includes('<a:srgbClr val="554433"/>'));
    assert.ok(xml.includes('<a:srgbClr val="775533"/>'));
    assert.ok(xml.includes('<a:prstDash val="dashDot"/>'));
    assert.ok(xml.includes('<a:prstDash val="dot"/>'));
    assert.ok(xml.includes("<c:dispBlanksAs"));
    assert.ok(xml.includes("Quarterly trend"));
    assert.ok(xml.includes(">Quarter<"));
    assert.ok(xml.includes(">Revenue<"));
    assert.ok(xml.includes('<c:numFmt formatCode="#,##0" sourceLinked="0"/>'));
    assert.ok(xml.includes("<c:majorGridlines>"));
    assert.ok(xml.includes("0.0"));
    assert.ok(
      xml.includes(
        '<c:idx val="2"/><c:marker><c:symbol val="none"/></c:marker>',
      ),
    );
    const seriesXml = [...xml.matchAll(/<c:ser>(.*?)<\/c:ser>/gsu)].map(
      (match) => match[1]!,
    );
    assert.equal(seriesXml.length, 2);
    for (const series of seriesXml) {
      assert.ok(series.includes('<c:showPercent val="0"/>'));
    }

    const workbookPath = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPath);
    const worksheet = strFromU8(
      unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.ok(worksheet.includes("Q1"));
    assert.ok(worksheet.includes("Actual"));
    assert.ok(!worksheet.includes("<v>0</v>"));
  });

  it("writes standard, stacked, and percent line and area variants while preserving null slots", async () => {
    const categories = ["Q1", "Q2", "Q3"];
    const series = [
      { name: "Actual", values: [10, null, 30] as const },
      { name: "Plan", values: [12, null, 25] as const },
    ];
    const variant = (
      id: string,
      kind: "line" | "area",
      stacking: "stacked" | "percent" | undefined,
      displayBlanksAs: "gap" | "zero" | "span",
      x: number,
      y: number,
    ) =>
      documentChart({
        id,
        x,
        y,
        width: 280,
        height: 150,
        displayBlanksAs,
        groups: [
          {
            kind,
            ...(stacking === undefined ? {} : { stacking }),
            data: { kind: "shared-categories", categories },
            series:
              kind === "line"
                ? series.map((item, index) => ({
                    ...item,
                    marker: { symbol: index === 0 ? "square" : "triangle" },
                    smooth: index === 0,
                  }))
                : series,
            ...(kind === "line"
              ? { dataLabels: { value: true, numberFormat: "0.0" } }
              : {}),
          },
        ],
      });
    const result = await buildDocument(
      deck({ width: 900, height: 520 }).append(
        documentSlide({ id: "variants" }).append(
          variant("line-standard", "line", undefined, "span", 10, 10),
          variant("line-stacked", "line", "stacked", "zero", 310, 10),
          variant("line-percent", "line", "percent", "zero", 610, 10),
          variant("area-standard", "area", undefined, "gap", 10, 180),
          variant("area-stacked", "area", "stacked", "zero", 310, 180),
          variant("area-percent", "area", "percent", "zero", 610, 180),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const charts = file.slides[0]?.objects.map(({ chart }) => chart);
    assert.deepEqual(
      charts?.map((chart) => [chart?.types[0], chart?.grouping]),
      [
        ["lineChart", "standard"],
        ["lineChart", "stacked"],
        ["lineChart", "percentStacked"],
        ["areaChart", "standard"],
        ["areaChart", "stacked"],
        ["areaChart", "percentStacked"],
      ],
    );
    assert.deepEqual(
      charts?.map((chart) => chart?.displayBlanksAs),
      ["span", "zero", "zero", "gap", "zero", "zero"],
    );

    const parts = unzipSync(result.pptx);
    const groupings = ["standard", "stacked", "percentStacked"];
    for (const [index, grouping] of groupings.entries()) {
      const line = chartXml(parts, index + 1);
      const area = chartXml(parts, index + 4);
      assert.ok(line.includes("<c:lineChart>"));
      assert.ok(line.includes(`<c:grouping val="${grouping}"/>`));
      assert.ok(!line.includes("<c:areaChart>"));
      assert.ok(area.includes("<c:areaChart>"));
      assert.ok(area.includes(`<c:grouping val="${grouping}"/>`));
      assert.ok(!area.includes("<c:lineChart>"));
      assert.ok(line.includes("<c:smooth"));
      assert.ok(!area.includes("<c:marker>"));
      assert.ok(!area.includes("<c:smooth"));
      assert.ok(line.includes('<c:symbol val="square"/>'));
      assert.ok(line.includes('<c:symbol val="triangle"/>'));
      assert.ok(line.includes("<c:smooth"));
      assert.ok(
        line.includes(
          '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache>',
        ),
      );
      assert.ok(
        line.includes(
          '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="2"><c:v>25</c:v></c:pt></c:numCache>',
        ),
      );
      for (const chart of [line, area]) {
        assert.ok(chart.includes('<c:dispBlanksAs val="'));
      }
    }
    const worksheets = charts?.map((chart) =>
      strFromU8(
        unzipSync(parts[chart!.embeddedWorkbook!.part]!)[
          "xl/worksheets/sheet1.xml"
        ]!,
      ),
    );
    for (const [index, worksheet] of worksheets?.entries() ?? []) {
      assert.ok(!worksheet.includes('r="B3"'), `chart ${index + 1} B3`);
      assert.ok(!worksheet.includes('r="C3"'), `chart ${index + 1} C3`);
      assert.ok(worksheet.includes('r="B2"'), `chart ${index + 1} B2`);
    }
  });

  it("rejects individual fill colors on area data points with a field diagnostic", async () => {
    const id = "area-point-fill";

    await assert.rejects(
      buildDocument(
        deck({ width: 240, height: 160 }).append(
          documentSlide({ id }).append(
            documentChart({
              id,
              x: 0,
              y: 0,
              width: 220,
              height: 140,
              groups: [
                {
                  kind: "area",
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2", "Q3"],
                  },
                  series: [
                    {
                      name: "Actual",
                      values: [10, 20, 30],
                      points: { 2: { color: "accent2" } },
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
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /Chart\.groups\[0\]\.series\[0\]\.points\[2\]\.color/u,
        );
        assert.equal(
          error.diagnostics[0]?.message,
          "Area chart data points do not support individual fill colors; set the series color instead.",
        );
        return true;
      },
    );
  });

  it("keeps area series fill and supported data point formatting", async () => {
    const result = await build(
      <Deck width={240} height={160}>
        <Slide id="area-point-formatting">
          <Chart
            id="area-point-formatting"
            x={0}
            y={0}
            width={220}
            height={140}
            groups={[
              {
                kind: "area",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                dataLabels: { value: true },
                series: [
                  {
                    name: "Actual",
                    values: [10, 20, 30],
                    color: "#336699",
                    points: {
                      2: {
                        line: { color: "#445566", width: 2 },
                        labelColor: "#112233",
                        dataLabels: { numberFormat: "0.00" },
                      },
                    },
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
    assert.deepEqual(file.slides[0]?.objects[0]?.chart?.types, ["areaChart"]);
    const xml = chartXml(unzipSync(result.pptx), 1);
    assert.ok(xml.includes('<a:srgbClr val="336699"/>'));
    assert.ok(xml.includes('<a:srgbClr val="445566"/>'));
    assert.ok(xml.includes('<a:srgbClr val="112233"/>'));
    assert.ok(xml.includes('<c:idx val="2"/>'));
    assert.ok(xml.includes('<c:numFmt formatCode="0.00" sourceLinked="0"/>'));
  });

  it("diagnoses marker, smooth, and missing-value combinations that the native chart kinds reject", async () => {
    const data = {
      kind: "shared-categories" as const,
      categories: ["Q1", "Q2"],
    };
    const cases: readonly [string, ChartProps][] = [
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "area",
              data,
              series: [{ name: "S", values: [1, 2], marker: true }],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "area",
              data,
              series: [
                { name: "S", values: [1, 2], points: { 0: { marker: true } } },
              ],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "column",
              data,
              series: [{ name: "S", values: [1, 2], marker: true }],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "line",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  marker: { symbol: "picture" as never },
                },
              ],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "line",
              data,
              series: [{ name: "S", values: [1, 2], marker: { size: 1 } }],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "line",
              data,
              series: [{ name: "S", values: [1, 2], marker: { size: 7.5 } }],
            },
          ],
        },
      ],
      [
        "invalid-marker",
        {
          groups: [
            {
              kind: "line",
              data,
              series: [
                { name: "S", values: [1, 2], marker: "circle" as never },
              ],
            },
          ],
        },
      ],
      [
        "invalid-smooth",
        {
          groups: [
            {
              kind: "area",
              data,
              series: [{ name: "S", values: [1, 2], smooth: true }],
            },
          ],
        },
      ],
      [
        "invalid-smooth",
        {
          groups: [
            {
              kind: "column",
              data,
              series: [{ name: "S", values: [1, 2], smooth: true }],
            },
          ],
        },
      ],
      [
        "invalid-smooth",
        {
          groups: [
            {
              kind: "line",
              data,
              series: [{ name: "S", values: [1, 2], smooth: "yes" as never }],
            },
          ],
        },
      ],
      [
        "invalid-chart-missing-values",
        {
          displayBlanksAs: "span",
          groups: [
            {
              kind: "area",
              data,
              series: [{ name: "S", values: [1, null] }],
            },
          ],
        },
      ],
      [
        "invalid-chart-missing-values",
        {
          displayBlanksAs: "gap",
          groups: [
            {
              kind: "line",
              stacking: "percent",
              data,
              series: [{ name: "S", values: [1, null] }],
            },
          ],
        },
      ],
      [
        "invalid-chart-missing-values",
        {
          displayBlanksAs: "span",
          groups: [
            {
              kind: "line",
              stacking: "stacked",
              data,
              series: [{ name: "S", values: [1, null] }],
            },
          ],
        },
      ],
      [
        "invalid-chart-missing-values",
        {
          displayBlanksAs: "gap",
          groups: [
            {
              kind: "area",
              stacking: "percent",
              data,
              series: [{ name: "S", values: [1, null] }],
            },
          ],
        },
      ],
    ];
    for (const [code, props] of cases) {
      await rejectsWithCode(props, code);
    }
  });
  it("reads native marker and smooth values from the actual file and rejects malformed values", async () => {
    const result = await build(
      <Deck>
        <Slide id="line-reader">
          <Chart
            id="line-reader"
            x={0}
            y={0}
            width={240}
            height={160}
            groups={[
              {
                kind: "line",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [
                  {
                    name: "Actual",
                    values: [1, 2],
                    marker: { symbol: "diamond", size: 9 },
                    smooth: true,
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
    assert.deepEqual(file.slides[0]?.objects[0]?.chart?.series[0]?.marker, {
      symbol: "diamond",
      size: 9,
    });
    assert.equal(file.slides[0]?.objects[0]?.chart?.series[0]?.smooth, true);

    const parts = unzipSync(result.pptx);
    const path = "ppt/charts/chart1.xml";
    const malformed = chartXml(parts, 1)
      .replace('<c:size val="9"/>', '<c:size val="1"/>')
      .replace('<c:smooth val="1"/>', '<c:smooth val="maybe"/>');
    const invalid = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(malformed) }),
    );
    assert.ok(
      invalid.errors.some(
        ({ code, message }) =>
          code === "invalid-chart-series-format" &&
          message.includes("marker size"),
      ),
    );
    assert.ok(
      invalid.errors.some(
        ({ code, message }) =>
          code === "invalid-chart-series-format" && message.includes("smooth"),
      ),
    );
  });
  it("writes line label positions and series/point overrides while keeping area labels position-free", async () => {
    const props: ChartProps = {
      id: "line-label-overrides",
      x: 0,
      y: 0,
      width: 360,
      height: 240,
      groups: [
        {
          kind: "line",
          data: { kind: "shared-categories", categories: ["A", "B", "C"] },
          dataLabels: {
            value: true,
            category: true,
            position: "right",
            numberFormat: "#,##0.0",
          },
          series: [
            {
              name: "S0",
              values: [1, 2, 3],
              dataLabels: {
                value: true,
                position: "left",
                numberFormat: "0.0",
              },
              points: {
                1: {
                  dataLabels: {
                    numberFormat: "0.00",
                    position: "bottom",
                  },
                },
                2: { dataLabels: false },
              },
            },
            {
              name: "S1",
              values: [4, 5, 6],
              dataLabels: {
                category: true,
                position: "top",
                numberFormat: "0",
              },
            },
          ],
        },
      ],
    };
    const results = [
      await build(
        <Deck width={480} height={320}>
          <Slide id="react-line-labels">
            <Chart {...props} />
          </Slide>
        </Deck>,
      ),
      await buildDocument(
        deck({ width: 480, height: 320 }).append(
          documentSlide({ id: "document-line-labels" }).append(
            documentChart(props),
          ),
        ),
      ),
    ];
    for (const result of results) {
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const chart = file.slides[0]?.objects[0]?.chart;
      assert.deepEqual(chart?.seriesDataLabels, [
        {
          seriesIndex: 0,
          dataLabels: {
            numberFormat: "0.0",
            percent: false,
            position: "left",
            value: true,
            category: true,
          },
          pointDataLabels: {
            1: {
              numberFormat: "0.00",
              percent: false,
              position: "bottom",
              value: true,
              category: true,
            },
            2: false,
          },
        },
        {
          seriesIndex: 1,
          dataLabels: {
            numberFormat: "0",
            percent: false,
            position: "top",
            value: true,
            category: true,
          },
        },
      ]);
      const xml = chartXml(unzipSync(result.pptx), 1);
      assert.ok(xml.includes('<c:dLblPos val="l"/>'));
      assert.ok(xml.includes('<c:dLblPos val="b"/>'));
      assert.ok(xml.includes('<c:dLblPos val="t"/>'));
      assert.ok(!xml.includes('<c:dLblPos val="r"/>'));
      assert.ok(!xml.includes('<c:dLblPos val="bestFit"/>'));
    }

    const areaResult = await build(
      <Deck width={480} height={320}>
        <Slide id="area-labels">
          <Chart
            id="area-labels"
            x={0}
            y={0}
            width={360}
            height={240}
            groups={[
              {
                kind: "area",
                data: { kind: "shared-categories", categories: ["A", "B"] },
                dataLabels: { value: true, numberFormat: "#,##0.0" },
                series: [
                  {
                    name: "S",
                    values: [1, 2],
                    dataLabels: { numberFormat: "0.0" },
                    points: { 1: { dataLabels: false } },
                  },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const areaFile = await validatePptx(areaResult.pptx);
    assert.deepEqual(areaFile.errors, []);
    assert.deepEqual(areaFile.slides[0]?.objects[0]?.chart?.seriesDataLabels, [
      {
        seriesIndex: 0,
        dataLabels: {
          numberFormat: "0.0",
          percent: false,
          value: true,
          category: false,
        },
        pointDataLabels: { 1: false },
      },
    ]);
    assert.ok(!chartXml(unzipSync(areaResult.pptx), 1).includes("<c:dLblPos"));

    const invalidCases: readonly [ChartProps, string][] = [
      [
        {
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A"] },
              dataLabels: { position: "outsideEnd" },
              series: [{ name: "S", values: [1] }],
            },
          ],
        } as unknown as ChartProps,
        "invalid-data-labels",
      ],
      [
        {
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [
                {
                  name: "S",
                  values: [1],
                  dataLabels: { position: "bestFit" },
                },
              ],
            },
          ],
        },
        "invalid-data-labels",
      ],
      [
        {
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [
                {
                  name: "S",
                  values: [1],
                  points: { 0: { dataLabels: { position: "insideEnd" } } },
                },
              ],
            },
          ],
        },
        "invalid-data-labels",
      ],
      [
        {
          groups: [
            {
              kind: "line",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [
                {
                  name: "S",
                  values: [1],
                  dataLabels: { percent: true },
                },
              ],
            },
          ],
        },
        "invalid-data-labels",
      ],
      [
        {
          groups: [
            {
              kind: "area",
              data: { kind: "shared-categories", categories: ["A"] },
              dataLabels: { position: "left" },
              series: [{ name: "S", values: [1] }],
            },
          ],
        } as unknown as ChartProps,
        "invalid-data-labels",
      ],
    ];
    for (const [caseProps, code] of invalidCases) {
      await rejectsWithCode(caseProps, code);
    }
  });
});
