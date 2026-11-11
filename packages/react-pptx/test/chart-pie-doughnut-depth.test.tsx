import assert from "node:assert/strict";

import {
  BuildError,
  build,
  Chart,
  Deck,
  Slide,
  type ChartProps,
} from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@react-pptx/react-pptx/document";
import { validatePptx, type FileChartDataLabel } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

function chartXml(parts: Unzipped, index: number): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

function labelSettings(
  value: false | FileChartDataLabel | undefined,
): FileChartDataLabel | undefined {
  return value === false ? undefined : value;
}

async function rejectsWithCode(
  props: ChartProps,
  code: string,
  message?: string,
): Promise<void> {
  const id = String(props.id ?? "invalid-pie-doughnut-chart");
  await assert.rejects(
    buildDocument(
      deck({ width: 240, height: 160 }).append(
        documentSlide({ id }).append(documentChart({ ...props, id })),
      ),
    ),
    (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, code, id);
      if (message !== undefined) {
        assert.equal(error.diagnostics[0]?.message, message);
      }
      assert.match(error.diagnostics[0]?.path ?? "", new RegExp(id, "u"));
      return true;
    },
  );
}

function labelCaseChart(
  id: string,
  kind: "pie" | "doughnut",
  groupLabels: unknown,
  seriesLabels: unknown,
): ChartProps {
  return {
    id,
    x: 0,
    y: 0,
    width: 400,
    height: 300,
    groups: [
      {
        kind,
        data: { kind: "shared-categories", categories: ["A", "A", "C"] },
        explosion: 10,
        firstSliceAngle: 30,
        ...(kind === "doughnut" ? { holeSize: 55 } : {}),
        dataLabels: groupLabels,
        series: [
          {
            name: "S",
            values: [1, null, 3],
            explosion: 0,
            dataLabels: seriesLabels,
            points: {
              0: { color: "#FF0000", explosion: 0 },
              1: { color: "#00FF00" },
              2: { color: "#0000FF", explosion: 5 },
            },
          },
        ],
      },
    ],
  } as unknown as ChartProps;
}

async function buildLabelCase(
  id: string,
  props: ChartProps,
  entry: "react" | "document",
) {
  if (entry === "react") {
    return build(
      <Deck>
        <Slide id={id}>
          <Chart {...props} />
        </Slide>
      </Deck>,
    );
  }
  return buildDocument(
    deck().append(documentSlide({ id }).append(documentChart(props))),
  );
}

describe("pie and doughnut depth", () => {
  it("writes native pie options, leader lines, and series/point overrides", async () => {
    const categories = ["Direct", "Partner", "Direct", "Online"];
    const result = await build(
      <Deck width={900} height={520}>
        <Slide id="pie-doughnut-depth">
          <Chart
            id="pie-depth"
            name="pie-depth"
            x={24}
            y={24}
            width={520}
            height={460}
            title={{
              text: "Channel mix",
              style: {
                fontFamily: "Jost",
                fontSize: "15pt",
                fontWeight: 700,
                color: "#112233",
              },
            }}
            legend="top"
            style={{ fontFamily: "Jost", fontSize: "10pt", color: "#112233" }}
            groups={[
              {
                kind: "pie",
                data: { kind: "shared-categories", categories },
                series: [
                  {
                    name: "Channel & mix",
                    values: [42.5, null, 27.5, 30],
                    color: "#336699",
                    line: { color: "#FFFFFF", width: 1, dash: "solid" },
                    explosion: 18,
                    dataLabels: {
                      position: "insideEnd",
                      numberFormat: '#,##0" units"',
                    },
                    points: {
                      0: {
                        color: "#CC3300",
                        line: { color: "#112233", width: 2, dash: "dash" },
                        explosion: 9,
                        labelColor: "#FFFFFF",
                        dataLabels: {
                          position: "outsideEnd",
                          numberFormat: '0.0"K"',
                        },
                      },
                      2: {
                        color: "#227744",
                        explosion: 4,
                        dataLabels: false,
                      },
                      3: {
                        line: { color: "#112233", width: 1, dash: "dot" },
                      },
                    },
                  },
                ],
                explosion: 25.5,
                firstSliceAngle: 359.5,
                dataLabels: {
                  value: true,
                  percent: true,
                  category: true,
                  position: "outsideEnd",
                  numberFormat: "0.0%",
                  leaderLines: {
                    color: "#445566",
                    width: 1.5,
                    dash: "dash",
                  },
                  style: {
                    fontFamily: "Jost",
                    fontSize: "9pt",
                    fontWeight: 700,
                    color: "#112233",
                  },
                },
              },
            ]}
          />
          <Chart
            id="doughnut-depth"
            name="doughnut-depth"
            x={568}
            y={24}
            width={308}
            height={460}
            title="Single ring"
            legend="right"
            style={{ fontFamily: "Jost", fontSize: "10pt", color: "#112233" }}
            groups={[
              {
                kind: "doughnut",
                data: {
                  kind: "shared-categories",
                  categories: ["Alpha", "Beta", "Gamma"],
                },
                series: [
                  {
                    name: "Ring",
                    values: [18, null, 32],
                    color: "#AA8822",
                    line: { color: "#FFFFFF", width: 1 },
                    points: {
                      0: {
                        color: "#227744",
                        line: { color: "#112233", width: 1, dash: "dot" },
                        explosion: 5,
                      },
                      2: { explosion: 7 },
                    },
                  },
                ],
                explosion: 400,
                firstSliceAngle: 45,
                holeSize: 55,
                dataLabels: {
                  value: true,
                  percent: true,
                  numberFormat: "0%",
                  leaderLines: {
                    color: "#224466",
                    width: 0.75,
                    dash: "dot",
                  },
                  style: {
                    fontFamily: "Jost",
                    fontSize: "9pt",
                    color: "#112233",
                  },
                },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const parts = unzipSync(result.pptx);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const pie = file.slides[0]?.objects[0]?.chart;
    const doughnut = file.slides[0]?.objects[1]?.chart;
    assert.equal(pie?.types[0], "pieChart");
    assert.equal(doughnut?.types[0], "doughnutChart");
    assert.equal(pie?.title, "Channel mix");
    assert.equal(doughnut?.title, "Single ring");
    assert.deepEqual(pie?.axes, []);
    assert.deepEqual(doughnut?.axes, []);
    assert.equal(pie?.groups[0]?.firstSliceAngle, 360);
    assert.equal(doughnut?.groups[0]?.firstSliceAngle, 45);
    assert.equal(doughnut?.groups[0]?.holeSize, 55);
    assert.equal(pie?.series[0]?.explosion, 18);
    assert.equal(doughnut?.series[0]?.explosion, 400);
    assert.deepEqual(pie?.series[0]?.pointExplosions, { 0: 9, 2: 4 });
    assert.deepEqual(doughnut?.series[0]?.pointExplosions, { 0: 5, 2: 7 });
    assert.deepEqual(pie?.series[0]?.categories, categories);
    assert.deepEqual(pie?.series[0]?.values, [42.5, null, 27.5, 30]);
    assert.deepEqual(
      pie?.embeddedWorkbook?.series.map(
        ({
          name,
          categories: observedCategories,
          values,
          index,
          order,
          groupIndex,
        }) => ({
          name,
          categories: observedCategories,
          values,
          index,
          order,
          groupIndex,
        }),
      ),
      pie?.series.map(
        ({
          name,
          categories: observedCategories,
          values,
          index,
          order,
          groupIndex,
        }) => ({
          name,
          categories: observedCategories,
          values,
          index,
          order,
          groupIndex,
        }),
      ),
    );

    const pieLabels = pie?.seriesDataLabels?.[0];
    const pieSettings = labelSettings(pieLabels?.dataLabels);
    assert.equal(pieSettings?.value, true);
    assert.equal(pieSettings?.percent, true);
    assert.equal(pieSettings?.category, true);
    assert.equal(pieSettings?.position, "insideEnd");
    assert.equal(pieSettings?.numberFormat, '#,##0" units"');
    assert.equal(pieSettings?.leaderLines, true);
    assert.deepEqual(pieSettings?.leaderLine, {
      color: "#445566",
      widthEmu: 14288,
      dash: "dash",
    });
    const piePointSettings = labelSettings(pieLabels?.pointDataLabels?.[0]);
    assert.equal(piePointSettings?.position, "outsideEnd");
    assert.equal(piePointSettings?.numberFormat, '0.0"K"');
    assert.equal(pieLabels?.pointDataLabels?.[2], false);
    const doughnutSettings = labelSettings(
      doughnut?.seriesDataLabels?.[0]?.dataLabels,
    );
    assert.equal(doughnutSettings?.position, undefined);
    assert.equal(doughnutSettings?.leaderLines, true);
    assert.deepEqual(doughnutSettings?.leaderLine, {
      color: "#224466",
      widthEmu: 7144,
      dash: "dot",
    });

    const pieXml = chartXml(parts, 1);
    const doughnutXml = chartXml(parts, 2);
    assert.ok(pieXml.includes("<c:pieChart>"));
    assert.ok(pieXml.includes("Channel mix"));
    assert.ok(pieXml.includes('<c:legendPos val="t"/>'));
    assert.ok(pieXml.includes('<c:firstSliceAng val="360"/>'));
    assert.ok(pieXml.includes('<c:explosion val="18"/>'));
    assert.ok(!pieXml.includes('<c:explosion val="26"/>'));
    assert.ok(pieXml.includes('<c:dLblPos val="inEnd"/>'));
    assert.ok(pieXml.includes('<c:dLblPos val="outEnd"/>'));
    assert.ok(pieXml.includes('<c:showLeaderLines val="1"/>'));
    assert.ok(pieXml.includes("<c:leaderLines>"));
    assert.ok(pieXml.includes('<a:ln w="14288">'));
    assert.ok(pieXml.includes('<a:prstDash val="dash"/>'));
    assert.ok(
      pieXml.includes('<c:dLbl><c:idx val="2"/><c:delete val="1"/></c:dLbl>'),
    );
    assert.ok(pieXml.includes('<a:srgbClr val="CC3300"/>'));
    assert.ok(pieXml.includes('<a:srgbClr val="227744"/>'));
    assert.ok(!pieXml.includes("<c:catAx"));
    assert.ok(!pieXml.includes("<c:valAx"));
    assert.ok(doughnutXml.includes("<c:doughnutChart>"));
    assert.ok(doughnutXml.includes("Single ring"));
    assert.ok(doughnutXml.includes('<c:legendPos val="r"/>'));
    assert.ok(doughnutXml.includes('<c:firstSliceAng val="45"/>'));
    assert.ok(doughnutXml.includes('<c:holeSize val="55"/>'));
    assert.ok(doughnutXml.includes('<c:explosion val="400"/>'));
    assert.ok(doughnutXml.includes("<c:leaderLines>"));
    assert.ok(!doughnutXml.includes("<c:dLblPos"));
    assert.ok(!doughnutXml.includes("<c:catAx"));
    assert.ok(!doughnutXml.includes("<c:valAx"));

    const pieWorkbook = unzipSync(parts[pie?.embeddedWorkbook?.part ?? ""]!);
    const pieSheet = strFromU8(pieWorkbook["xl/worksheets/sheet1.xml"]!);
    assert.ok(pieSheet.includes("Direct"));
    assert.ok(pieSheet.includes("<v>42.5</v>"));
    assert.ok(!pieSheet.includes("<v>0</v>"));
    assert.ok(!pieSheet.includes('<c r="C2"'));
  });

  it("builds pie depth through the imperative document factory", async () => {
    const chart = documentChart({
      id: "imperative-pie",
      name: "imperative-pie",
      x: 20,
      y: 20,
      width: 520,
      height: 420,
      title: "Imperative pie",
      legend: "bottom",
      groups: [
        {
          kind: "pie",
          data: {
            kind: "shared-categories",
            categories: ["A", "B", "C"],
          },
          series: [
            {
              name: "Share",
              values: [5, 3, 2],
              explosion: 12,
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: "0.0",
              },
              points: {
                1: { explosion: 6 },
              },
            },
          ],
          explosion: 33,
          firstSliceAngle: 90,
          dataLabels: {
            value: true,
            position: "outsideEnd",
            numberFormat: "0.0",
            leaderLines: true,
          },
        },
      ],
    } as ChartProps);
    const result = await buildDocument(
      deck({ width: 640, height: 480 }).append(
        documentSlide({ id: "imperative-pie-slide" }).append(chart),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.ok(xml.includes('<c:firstSliceAng val="90"/>'));
    assert.ok(xml.includes('<c:explosion val="12"/>'));
    assert.ok(xml.includes('<c:explosion val="6"/>'));
    assert.ok(!xml.includes('<c:explosion val="33"/>'));
    assert.ok(xml.includes('<c:showLeaderLines val="1"/>'));
    assert.ok(xml.includes("<c:leaderLines/>"));
    const observed = file.slides[0]?.objects[0]?.chart;
    assert.equal(observed?.groups[0]?.firstSliceAngle, 90);
    assert.equal(observed?.series[0]?.explosion, 12);
    assert.deepEqual(observed?.series[0]?.pointExplosions, { 1: 6 });
    const observedLabels = labelSettings(
      observed?.seriesDataLabels?.[0]?.dataLabels,
    );
    assert.equal(observedLabels?.leaderLines, true);
    assert.equal(observedLabels?.position, "outsideEnd");
  });

  it("keeps group, series, and point data-label hiding and re-declaration", async () => {
    const result = await build(
      <Deck>
        <Slide id="pie-label-overrides">
          <Chart
            id="pie-label-overrides"
            x={0}
            y={0}
            width={420}
            height={300}
            groups={[
              {
                kind: "pie",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B", "C"],
                },
                series: [
                  {
                    name: "S",
                    values: [4, 3, 2],
                    dataLabels: false,
                    points: {
                      0: {
                        dataLabels: {
                          value: true,
                          category: true,
                          position: "insideEnd",
                          numberFormat: "0",
                        },
                      },
                    },
                  },
                ],
                dataLabels: {
                  value: true,
                  position: "outsideEnd",
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
    const labels = file.slides[0]?.objects[0]?.chart?.seriesDataLabels?.[0];
    const hiddenSeriesLabels = labelSettings(labels?.dataLabels);
    assert.equal(hiddenSeriesLabels?.value, false);
    assert.equal(hiddenSeriesLabels?.percent, false);
    assert.equal(hiddenSeriesLabels?.category, false);
    const pointLabels = labelSettings(labels?.pointDataLabels?.[0]);
    assert.equal(pointLabels?.value, true);
    assert.equal(pointLabels?.position, "insideEnd");
    const xml = chartXml(unzipSync(result.pptx), 1);
    assert.ok(xml.includes('<c:dLbl><c:idx val="0"/>'));
    assert.ok(!xml.includes('<c:dLbl><c:idx val="1"/>'));
    assert.ok(!xml.includes('<c:dLbls><c:delete val="1"/></c:dLbls>'));
  });

  it("rejects native leader lines outside pie and doughnut charts", async () => {
    const result = await build(
      <Deck>
        <Slide id="column-leader-line-reader">
          <Chart
            id="column-leader-line-reader"
            x={0}
            y={0}
            width={240}
            height={160}
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B"],
                },
                series: [{ name: "S", values: [1, 2] }],
                dataLabels: { value: true },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const source = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(source[chartPath]!);
    const withLeaderLines = xml.replace(
      "</c:dLbls>",
      '<c:showLeaderLines val="1"/><c:leaderLines/></c:dLbls>',
    );
    assert.notEqual(withLeaderLines, xml);
    const withLeaderLinesFile = await validatePptx(
      zipSync({ ...source, [chartPath]: strToU8(withLeaderLines) }),
    );
    assert.ok(
      withLeaderLinesFile.errors.some(
        ({ code }) => code === "invalid-data-label-leader-lines",
      ),
    );

    const withoutShowLeaderLines = xml.replace(
      "</c:dLbls>",
      "<c:leaderLines/></c:dLbls>",
    );
    const withoutShowFile = await validatePptx(
      zipSync({
        ...source,
        [chartPath]: strToU8(withoutShowLeaderLines),
      }),
    );
    assert.ok(
      withoutShowFile.errors.some(
        ({ code, message }) =>
          code === "invalid-data-label-leader-lines" &&
          message.includes("showLeaderLines"),
      ),
    );
  });

  it("rejects inapplicable pie, doughnut, and label combinations", async () => {
    const data = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const series = [{ name: "S", values: [1, 2] }];
    const base = { x: 0, y: 0, width: 200, height: 120 } as const;
    const cases: {
      readonly id: string;
      readonly code: string;
      readonly message?: string;
      readonly props: ChartProps;
    }[] = [
      {
        id: "pie-axis",
        code: "invalid-chart-axes",
        props: {
          ...base,
          axes: {
            category: { kind: "category" },
            value: { kind: "value" },
          },
          groups: [{ kind: "pie", data, series }],
        },
      },
      {
        id: "pie-stacking",
        code: "invalid-chart-option",
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series,
              stacking: "stacked",
            } as unknown as ChartProps["groups"][number],
          ],
        },
      },
      {
        id: "pie-series-marker",
        code: "invalid-marker",
        message:
          'Series "S".marker is only applicable to line chart series and points',
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  marker: "circle",
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "pie-series-smooth",
        code: "invalid-smooth",
        message: 'Series "S".smooth is only applicable to line chart series',
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  smooth: true,
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "doughnut-label-position",
        code: "invalid-data-labels",
        message:
          "Office does not allow dLblPos for doughnutChart series data labels",
        props: {
          ...base,
          groups: [
            {
              kind: "doughnut",
              data,
              series,
              dataLabels: { value: true, position: "outsideEnd" },
            },
          ],
        } as unknown as ChartProps,
      },
      {
        id: "pie-line-position",
        code: "invalid-data-labels",
        message:
          'Office does not allow dLblPos="l" for pieChart series data labels',
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series,
              dataLabels: { value: true, position: "left" },
            },
          ],
        } as unknown as ChartProps,
      },
      {
        id: "pie-inside-base",
        code: "invalid-data-labels",
        message:
          'Office does not allow dLblPos="inBase" for pieChart series data labels',
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series,
              dataLabels: { value: true, position: "insideBase" },
            },
          ],
        } as unknown as ChartProps,
      },
      {
        id: "column-leader-lines",
        code: "invalid-data-labels",
        message: "leaderLines is only supported for pie and doughnut charts",
        props: {
          ...base,
          groups: [
            {
              kind: "column",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  dataLabels: { value: true, leaderLines: true },
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "pie-point-leader-lines",
        code: "unknown-parameter",
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  points: {
                    0: {
                      dataLabels: {
                        value: true,
                        leaderLines: true,
                      },
                    },
                  },
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "column-series-explosion",
        code: "invalid-chart-option",
        message:
          'Series "S" explosion is only supported for pie and doughnut charts',
        props: {
          ...base,
          groups: [
            {
              kind: "column",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  explosion: 10,
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "column-point-explosion",
        code: "invalid-chart-option",
        message:
          'Series "S" data point 0 explosion is only supported for pie and doughnut charts',
        props: {
          ...base,
          groups: [
            {
              kind: "column",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  points: { 0: { explosion: 10 } },
                } as unknown as ChartProps["groups"][number]["series"][number],
              ],
            },
          ],
        },
      },
      {
        id: "pie-leader-lines-without-labels",
        code: "invalid-data-labels",
        message:
          "leaderLines requires at least one visible data label value, percent, or category",
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series,
              dataLabels: { leaderLines: true },
            },
          ],
        },
      },
      {
        id: "pie-leader-lines-series-hides-value",
        code: "invalid-data-labels",
        message:
          "leaderLines requires at least one visible data label value, percent, or category",
        props: {
          ...base,
          groups: [
            {
              kind: "pie",
              data,
              series: [
                {
                  name: "S",
                  values: [1, 2],
                  dataLabels: { value: false },
                },
              ],
              dataLabels: { leaderLines: true },
            },
          ],
        },
      },
    ];

    for (const { id, code, message, props } of cases) {
      await rejectsWithCode({ ...props, id }, code, message);
    }
  });

  it("writes effective leader-line declarations across group and series overrides", async () => {
    type LabelCase = {
      readonly id: string;
      readonly groupLabels: unknown;
      readonly seriesLabels: unknown;
      readonly verify: (
        labels: false | FileChartDataLabel | undefined,
        xml: string,
        caseId: string,
      ) => void;
    };
    const cases: readonly LabelCase[] = [
      {
        id: "series-overrides-leader-lines",
        groupLabels: { leaderLines: true },
        seriesLabels: { value: true, leaderLines: false },
        verify(labels, xml, caseId) {
          const settings = labelSettings(labels);
          assert.equal(settings?.value, true, caseId);
          assert.equal(settings?.leaderLines, false, caseId);
          assert.ok(xml.includes('<c:showVal val="1"/>'), caseId);
          assert.ok(xml.includes('<c:showLeaderLines val="0"/>'), caseId);
          assert.ok(!xml.includes("<c:leaderLines"), caseId);
        },
      },
      {
        id: "series-hides-group-labels",
        groupLabels: { leaderLines: true },
        seriesLabels: false,
        verify(labels, xml, caseId) {
          assert.equal(labels, false, caseId);
          assert.ok(
            xml.includes('<c:dLbls><c:delete val="1"/></c:dLbls>'),
            caseId,
          );
          assert.ok(!xml.includes("showLeaderLines"), caseId);
        },
      },
      {
        id: "false-only-without-visible-labels",
        groupLabels: { leaderLines: false },
        seriesLabels: undefined,
        verify(labels, xml, caseId) {
          const settings = labelSettings(labels);
          assert.equal(settings?.value, false, caseId);
          assert.equal(settings?.category, false, caseId);
          assert.equal(settings?.percent, false, caseId);
          assert.equal(settings?.leaderLines, undefined, caseId);
          assert.ok(xml.includes('<c:showVal val="0"/>'), caseId);
          assert.ok(!xml.includes("showLeaderLines"), caseId);
        },
      },
      {
        id: "visible-group-label-series-disables-leader-lines",
        groupLabels: { value: true, leaderLines: true },
        seriesLabels: { leaderLines: false },
        verify(labels, xml, caseId) {
          const settings = labelSettings(labels);
          assert.equal(settings?.value, true, caseId);
          assert.equal(settings?.leaderLines, false, caseId);
          assert.ok(xml.includes('<c:showVal val="1"/>'), caseId);
          assert.ok(xml.includes('<c:showLeaderLines val="0"/>'), caseId);
          assert.ok(!xml.includes("<c:leaderLines"), caseId);
        },
      },
      {
        id: "visible-group-label-leader-line-style",
        groupLabels: {
          value: true,
          leaderLines: { color: "#112233", width: 2, dash: "dot" },
        },
        seriesLabels: undefined,
        verify(labels, xml, caseId) {
          const settings = labelSettings(labels);
          assert.equal(settings?.value, true, caseId);
          assert.equal(settings?.leaderLines, true, caseId);
          assert.deepEqual(
            settings?.leaderLine,
            { color: "#112233", widthEmu: 19050, dash: "dot" },
            caseId,
          );
          assert.ok(xml.includes('<c:showLeaderLines val="1"/>'), caseId);
          assert.ok(xml.includes("<c:leaderLines>"), caseId);
          assert.ok(xml.includes('<a:ln w="19050">'), caseId);
        },
      },
    ];

    for (const kind of ["pie", "doughnut"] as const) {
      for (const entry of ["react", "document"] as const) {
        for (const current of cases) {
          const id = `${kind}-${current.id}-${entry}`;
          const props = labelCaseChart(
            id,
            kind,
            current.groupLabels,
            current.seriesLabels,
          );
          const result = await buildLabelCase(id, props, entry);
          const file = await validatePptx(result.pptx);
          assert.deepEqual(file.errors, [], id);
          current.verify(
            file.slides[0]?.objects[0]?.chart?.seriesDataLabels?.[0]
              ?.dataLabels,
            chartXml(unzipSync(result.pptx), 1),
            id,
          );
        }
      }
    }
  });

  it("keeps inherited leader lines when a series adds a visible label", async () => {
    for (const kind of ["pie", "doughnut"] as const) {
      for (const entry of ["react", "document"] as const) {
        const id = `${kind}-series-adds-visible-label-${entry}`;
        const props = labelCaseChart(
          id,
          kind,
          { leaderLines: true },
          { value: true },
        );
        const result = await buildLabelCase(id, props, entry);
        const file = await validatePptx(result.pptx);
        assert.deepEqual(file.errors, [], id);
        const settings = labelSettings(
          file.slides[0]?.objects[0]?.chart?.seriesDataLabels?.[0]?.dataLabels,
        );
        assert.equal(settings?.value, true, id);
        assert.equal(settings?.leaderLines, true, id);
        const xml = chartXml(unzipSync(result.pptx), 1);
        assert.ok(xml.includes('<c:showLeaderLines val="1"/>'), id);
        assert.ok(xml.includes("<c:leaderLines"), id);
      }
    }
  });

  it("keeps group label inheritance across multiple series overrides", async () => {
    const result = await build(
      <Deck>
        <Slide id="multi-series-label-inheritance">
          <Chart
            id="multi-series-label-inheritance"
            x={0}
            y={0}
            width={400}
            height={300}
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B"],
                },
                dataLabels: { value: true, position: "outsideEnd" },
                series: [
                  { name: "Inherited", values: [1, 2] },
                  { name: "Hidden", values: [2, 3], dataLabels: false },
                  {
                    name: "Override",
                    values: [3, 4],
                    dataLabels: { position: "insideEnd", numberFormat: "0.0" },
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
    const labels = file.slides[0]?.objects[0]?.chart?.seriesDataLabels;
    assert.equal(labelSettings(labels?.[0]?.dataLabels)?.value, true);
    assert.equal(
      labelSettings(labels?.[0]?.dataLabels)?.position,
      "outsideEnd",
    );
    assert.equal(labels?.[1]?.dataLabels, false);
    assert.equal(labelSettings(labels?.[2]?.dataLabels)?.value, true);
    assert.equal(labelSettings(labels?.[2]?.dataLabels)?.position, "insideEnd");
    assert.equal(labelSettings(labels?.[2]?.dataLabels)?.numberFormat, "0.0");
    const xml = chartXml(unzipSync(result.pptx), 1);
    assert.equal(xml.match(/<c:showVal val="1"\/>/gu)?.length, 2);
    assert.ok(xml.includes('<c:dLbls><c:delete val="1"/></c:dLbls>'));
  });

  it("rejects showLeaderLines without a visible sibling label", async () => {
    const result = await build(
      <Deck>
        <Slide id="reader-leader-lines-without-labels">
          <Chart
            id="reader-leader-lines-without-labels"
            x={0}
            y={0}
            width={240}
            height={160}
            groups={[
              {
                kind: "pie",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B"],
                },
                series: [{ name: "S", values: [1, 2] }],
                dataLabels: { value: false },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const source = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const xml = strFromU8(source[chartPath]!);
    assert.ok(xml.includes('<c:showVal val="0"/>'));

    for (const [name, value] of [
      ["false", "0"],
      ["true", "1"],
    ] as const) {
      const mutated = xml.replace(
        '<c:showBubbleSize val="0"/>',
        `<c:showBubbleSize val="0"/><c:showLeaderLines val="${value}"/>`,
      );
      assert.notEqual(mutated, xml);
      const file = await validatePptx(
        zipSync({ ...source, [chartPath]: strToU8(mutated) }),
      );
      assert.ok(
        file.errors.some(
          ({ code, message }) =>
            code === "invalid-data-label-leader-lines" &&
            message.includes("visible"),
        ),
        name,
      );
    }

    const withVisibleSibling = xml
      .replace('<c:showSerName val="0"/>', '<c:showSerName val="1"/>')
      .replace(
        '<c:showBubbleSize val="0"/>',
        '<c:showBubbleSize val="0"/><c:showLeaderLines val="0"/>',
      );
    const visibleSiblingFile = await validatePptx(
      zipSync({
        ...source,
        [chartPath]: strToU8(withVisibleSibling),
      }),
    );
    assert.deepEqual(visibleSiblingFile.errors, []);
  });
});
