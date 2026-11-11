import assert from "node:assert/strict";

import {
  BuildError,
  build,
  Chart,
  Deck,
  fade,
  Slide,
  type ChartProps,
  type ScatterChartGroup,
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

async function rejectsWithCode(props: ChartProps, code: string): Promise<void> {
  const id = "invalid-scatter-chart";
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

function fileLabel(
  value: false | FileChartDataLabel | undefined,
): FileChartDataLabel {
  assert.ok(value !== false && value !== undefined);
  return value;
}

function xyGroup(
  series: ScatterChartGroup["series"],
  extra: Omit<ScatterChartGroup, "kind" | "data" | "series"> = {},
): ScatterChartGroup {
  return {
    kind: "scatter",
    data: { kind: "xy" },
    series,
    ...extra,
  };
}

async function buildInEntry(
  entry: "document" | "react",
  props: ChartProps,
): Promise<Uint8Array> {
  const id = "scatter-entry";
  if (entry === "document") {
    return (
      await buildDocument(
        deck({ width: 480, height: 320 }).append(
          documentSlide({ id }).append(documentChart({ ...props, id })),
        ),
      )
    ).pptx;
  }
  return (
    await build(
      <Deck width={480} height={320}>
        <Slide id={id}>
          <Chart {...props} id={id} />
        </Slide>
      </Deck>,
    )
  ).pptx;
}

async function rejectsInBothEntries(
  props: ChartProps,
  code: string,
): Promise<void> {
  for (const entry of ["document", "react"] as const) {
    await assert.rejects(
      buildInEntry(entry, props),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, code);
        return true;
      },
      `${entry} entry should reject with ${code}`,
    );
  }
}

describe("scatter depth", () => {
  it("writes independent unequal XY observations through the React entry", async () => {
    const props: ChartProps = {
      x: 10,
      y: 10,
      width: 420,
      height: 260,
      groups: [
        xyGroup([
          {
            name: "Observed",
            x: [1, 2.5, 4, 8],
            y: [3, null, 6, 9],
            color: "#336699",
            points: {
              1: { marker: { symbol: "star", size: 11 } },
              2: { color: "#AA5500" },
            },
          },
          {
            name: "Short",
            x: [1.5, null, 6],
            y: [2, 5, 7],
          },
        ]),
      ],
    };
    const result = await build(
      <Deck width={640} height={400}>
        <Slide id="scatter">
          <Chart {...props} />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.types, ["scatterChart"]);
    assert.deepEqual(
      chart?.series.map(({ name, x, y }) => [name, x, y]),
      [
        ["Observed", [1, 2.5, 4, 8], [3, null, 6, 9]],
        ["Short", [1.5, null, 6], [2, 5, 7]],
      ],
    );
    assert.deepEqual(chart?.groups[0]?.seriesIndices, [0, 1]);
    assert.deepEqual(
      chart?.axes.map(({ kind, role }) => [kind, role]),
      [
        ["value", "x"],
        ["value", "y"],
      ],
    );
    assert.deepEqual(
      chart?.embeddedWorkbook?.series.map(({ x, y }) => [x, y]),
      [
        [
          [1, 2.5, 4, 8],
          [3, null, 6, 9],
        ],
        [
          [1.5, null, 6],
          [2, 5, 7],
        ],
      ],
    );

    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.ok(xml.includes("<c:scatterChart>"));
    assert.ok(xml.includes('<c:scatterStyle val="marker"/>'));
    assert.ok(xml.includes("<c:xVal>"));
    assert.ok(xml.includes("<c:yVal>"));
    assert.ok(!xml.includes("<c:cat>"));
    assert.ok(!xml.includes("<c:val>"));
    assert.ok(xml.includes('<c:ptCount val="4"/>'));
    assert.ok(xml.includes('<c:ptCount val="3"/>'));

    const workbookPath = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPath);
    const worksheet = strFromU8(
      unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.ok(worksheet.includes("Observed"));
    assert.ok(worksheet.includes("Short"));
    assert.ok(worksheet.includes('r="B2"'));
    assert.ok(worksheet.includes('r="C4"'));
    assert.ok(!worksheet.includes('r="C3"'));
    assert.ok(worksheet.includes('r="B3"'));
    assert.ok(worksheet.includes('r="E3"'));
    assert.ok(!worksheet.includes('r="D3"'));
  });

  it("writes the five marker and line presentations through the imperative entry", async () => {
    const variant = (
      id: string,
      line: ScatterChartGroup["line"],
      markers: boolean,
      x: number,
      title: string,
    ) =>
      documentChart({
        id,
        x,
        y: 10,
        width: 300,
        height: 220,
        title,
        legend: "bottom",
        axes: {
          horizontal: {
            kind: "value",
            title: "Load",
            min: 0,
            max: 10,
            majorUnit: 2,
            numberFormat: "0.0",
            gridlines: { color: "#DDDDDD", width: 1, dash: "dash" },
            style: { fontSize: "9pt", color: "#333333" },
          },
          vertical: {
            kind: "value",
            title: "Latency",
            min: 1,
            max: 20,
            logBase: 10,
            numberFormat: "#,##0",
            gridlines: true,
          },
        },
        groups: [
          xyGroup(
            [
              {
                name: "North",
                x: [1, 2, 4, 8],
                y: [2, 4, 8, 16],
                ...(markers
                  ? { marker: { symbol: "circle", size: 8, fill: "#336699" } }
                  : {}),
              },
              {
                name: "South",
                x: [1, 3],
                y: [5, 10],
                ...(markers ? { marker: { symbol: "square", size: 7 } } : {}),
              },
            ],
            {
              markers,
              axes: { x: "horizontal", y: "vertical" },
              dataLabels: {
                value: true,
                position: "right",
                numberFormat: "0.0",
              },
              ...(line === undefined ? {} : { line }),
            },
          ),
        ],
      });
    const result = await buildDocument(
      deck({ width: 1600, height: 300 }).append(
        documentSlide({ id: "presentations" }).append(
          variant("marker-only", undefined, true, 10, "Markers"),
          variant("line", "straight", false, 320, "Line"),
          variant("line-marker", "straight", true, 630, "Line + markers"),
          variant("smooth", "smooth", false, 940, "Smooth"),
          variant("smooth-marker", "smooth", true, 1250, "Smooth + markers"),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const charts = file.slides[0]?.objects.map(({ chart }) => chart);
    assert.deepEqual(
      charts?.map((chart) => [chart?.types[0], chart?.groups[0]?.scatterStyle]),
      [
        ["scatterChart", "marker"],
        ["scatterChart", "line"],
        ["scatterChart", "lineMarker"],
        ["scatterChart", "smooth"],
        ["scatterChart", "smoothMarker"],
      ],
    );
    const xAxis = charts?.[0]?.axes.find(({ role }) => role === "x");
    const yAxis = charts?.[0]?.axes.find(({ role }) => role === "y");
    assert.deepEqual(
      [xAxis?.kind, xAxis?.title, xAxis?.numberFormat, xAxis?.min, xAxis?.max],
      ["value", "Load", "0.0", 0, 10],
    );
    assert.deepEqual(
      [yAxis?.kind, yAxis?.title, yAxis?.logBase, yAxis?.numberFormat],
      ["value", "Latency", 10, "#,##0"],
    );
    assert.equal(charts?.[0]?.series[0]?.marker?.symbol, "circle");
    assert.equal(charts?.[0]?.series[1]?.marker?.symbol, "square");

    const parts = unzipSync(result.pptx);
    const styles = ["marker", "line", "lineMarker", "smooth", "smoothMarker"];
    for (const [index, style] of styles.entries()) {
      const xml = chartXml(parts, index + 1);
      assert.ok(xml.includes(`<c:scatterStyle val="${style}"/>`));
      assert.ok(xml.includes("<c:xVal>"));
      assert.ok(xml.includes("<c:yVal>"));
      assert.ok(xml.includes(">Load<"));
      assert.ok(xml.includes(">Latency<"));
      assert.ok(xml.includes('<c:numFmt formatCode="0.0" sourceLinked="0"/>'));
      assert.ok(xml.includes("<c:majorGridlines"));
      if (style.startsWith("smooth")) {
        assert.ok(xml.includes('<c:smooth val="1"/>'));
      }
      if (style === "line" || style === "smooth") {
        assert.ok(xml.includes('<c:symbol val="none"/>'));
      }
    }
    assert.ok(!chartXml(parts, 1).includes("<c:smooth"));
    assert.ok(chartXml(parts, 3).includes('<c:smooth val="0"/>'));
  });

  it("keeps missing X and missing Y in their original slots with point formatting", async () => {
    const result = await buildDocument(
      deck({ width: 640, height: 400 }).append(
        documentSlide({ id: "missing" }).append(
          documentChart({
            id: "missing",
            x: 10,
            y: 10,
            width: 500,
            height: 300,
            groups: [
              xyGroup([
                {
                  name: "Duplicates",
                  x: [10, 3, 10, null, 1],
                  y: [1, 2, null, 4, 5],
                  points: {
                    0: { marker: { symbol: "star", size: 10 } },
                    2: { color: "#AA5500" },
                    3: { marker: { symbol: "diamond", size: 9 } },
                  },
                },
                {
                  name: "Short",
                  x: [2, null],
                  y: [3, 6],
                },
              ]),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.x, [10, 3, 10, null, 1]);
    assert.deepEqual(chart?.series[0]?.y, [1, 2, null, 4, 5]);
    assert.deepEqual(chart?.series[1]?.x, [2, null]);
    assert.deepEqual(chart?.series[1]?.y, [3, 6]);

    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.ok(xml.includes('<c:ptCount val="5"/>'));
    assert.ok(
      xml.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>10</c:v></c:pt><c:pt idx="4"><c:v>1</c:v></c:pt></c:numCache>',
      ),
    );
    assert.ok(
      xml.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:numCache>',
      ),
    );
    assert.ok(xml.includes('<c:idx val="0"/><c:marker><c:symbol val="star"/>'));
    assert.ok(xml.includes('<a:srgbClr val="AA5500"/>'));
    assert.ok(
      xml.includes('<c:idx val="3"/><c:marker><c:symbol val="diamond"/>'),
    );

    const workbookPath = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPath);
    const worksheet = strFromU8(
      unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.ok(worksheet.includes('r="B4"'));
    assert.ok(!worksheet.includes('r="B5"'));
    assert.ok(!worksheet.includes('r="C4"'));
    assert.ok(worksheet.includes('r="C5"'));
    assert.ok(worksheet.includes('r="B6"'));
    assert.ok(worksheet.includes('r="C6"'));
  });

  it("preserves all-missing and all-zero observations without fabricating values", async () => {
    const result = await buildDocument(
      deck({ width: 480, height: 300 }).append(
        documentSlide({ id: "all-missing-and-zero" }).append(
          documentChart({
            id: "all-missing-and-zero",
            x: 10,
            y: 10,
            width: 400,
            height: 240,
            groups: [
              xyGroup([
                {
                  name: "All null",
                  x: [null, null],
                  y: [null, null],
                },
                {
                  name: "All zero",
                  x: [0, 0],
                  y: [0, 0],
                },
              ]),
            ],
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.x, [null, null]);
    assert.deepEqual(chart?.series[0]?.y, [null, null]);
    assert.deepEqual(chart?.series[1]?.x, [0, 0]);
    assert.deepEqual(chart?.series[1]?.y, [0, 0]);
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.ok(
      xml.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/></c:numCache>',
      ),
    );
    assert.ok(
      xml.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numCache>',
      ),
    );
    const workbookPath = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPath);
    const worksheet = strFromU8(
      unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.ok(!worksheet.includes('r="B2"'));
    assert.ok(!worksheet.includes('r="C2"'));
    assert.ok(worksheet.includes('r="D2"'));
    assert.ok(worksheet.includes('r="E2"'));
    assert.ok(worksheet.includes('r="D3"'));
    assert.ok(worksheet.includes('r="E3"'));
  });

  it("applies blank rendering only where the participating groups draw a line", async () => {
    const lineGroup = xyGroup(
      [{ name: "Line", x: [1, 2, 3], y: [10, null, 30] }],
      { line: "straight", markers: false },
    );
    const markerGroup = xyGroup([
      { name: "Marker", x: [1, 2, 3], y: [10, null, 30] },
    ]);
    const matrix: readonly {
      readonly mode: "gap" | "zero" | "span";
      readonly groups: readonly ScatterChartGroup[];
    }[] = [
      { mode: "gap", groups: [markerGroup] },
      { mode: "zero", groups: [markerGroup] },
      { mode: "gap", groups: [lineGroup] },
      { mode: "zero", groups: [lineGroup] },
      { mode: "span", groups: [lineGroup] },
      {
        mode: "span",
        groups: [
          lineGroup,
          xyGroup([{ name: "Smooth", x: [2, 3], y: [20, 40] }], {
            line: "smooth",
          }),
        ],
      },
    ];
    for (const entry of ["document", "react"] as const) {
      for (const { mode, groups } of matrix) {
        const pptx = await buildInEntry(entry, {
          x: 10,
          y: 10,
          width: 320,
          height: 220,
          displayBlanksAs: mode,
          groups,
        });
        const file = await validatePptx(pptx);
        assert.deepEqual(file.errors, []);
        const chart = file.slides[0]?.objects[0]?.chart;
        assert.equal(chart?.displayBlanksAs, mode);
        assert.deepEqual(chart?.series[0]?.x, [1, 2, 3]);
        assert.deepEqual(chart?.series[0]?.y, [10, null, 30]);
        const parts = unzipSync(pptx);
        const xml = chartXml(parts, 1);
        assert.ok(xml.includes(`<c:dispBlanksAs val="${mode}"/>`));
        assert.ok(
          xml.includes(
            '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache>',
          ),
        );
        const workbookPath = chart?.embeddedWorkbook?.part;
        assert.ok(workbookPath);
        const worksheet = strFromU8(
          unzipSync(parts[workbookPath]!)["xl/worksheets/sheet1.xml"]!,
        );
        assert.ok(worksheet.includes('r="B3"'));
        assert.ok(worksheet.includes('r="C2"'));
        assert.ok(worksheet.includes('r="C4"'));
        assert.ok(!worksheet.includes('r="C3"'));
      }
    }
  });

  it("rejects span when any participating scatter group has no connecting line", async () => {
    const base = {
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      displayBlanksAs: "span" as const,
    };
    const lineGroup = xyGroup([{ name: "Line", x: [1, 2], y: [3, 4] }], {
      line: "straight",
    });
    const markerGroup = xyGroup([{ name: "Marker", x: [1, 2], y: [3, 4] }]);
    await rejectsInBothEntries(
      { ...base, groups: [markerGroup] },
      "invalid-chart-missing-values",
    );
    await rejectsInBothEntries(
      { ...base, groups: [lineGroup, markerGroup] },
      "invalid-chart-missing-values",
    );
    await rejectsInBothEntries(
      { ...base, groups: [markerGroup, lineGroup] },
      "invalid-chart-missing-values",
    );
  });

  it("writes scatter X and Y label content through group, series, and point inheritance", async () => {
    const props: ChartProps = {
      x: 10,
      y: 10,
      width: 340,
      height: 240,
      groups: [
        xyGroup(
          [
            { name: "Both", x: [1, 2, 3], y: [10, 20, 30] },
            {
              name: "Y only",
              x: [4, 5],
              y: [40, 50],
              dataLabels: { xValue: false },
            },
            {
              name: "X only",
              x: [6, 7],
              y: [60, 70],
              dataLabels: { value: false },
            },
            {
              name: "Point overrides",
              x: [8, 9, 10],
              y: [80, 90, 100],
              points: {
                0: { dataLabels: { xValue: false } },
                1: { dataLabels: { value: false } },
                2: { dataLabels: { numberFormat: "0", position: "top" } },
              },
            },
            {
              name: "Recovered point",
              x: [13, 14],
              y: [15, 16],
              dataLabels: false,
              points: { 1: { dataLabels: { xValue: true } } },
            },
            {
              name: "Missing coordinates",
              x: [null, 12],
              y: [13, null],
              dataLabels: { value: true, xValue: true },
            },
          ],
          {
            dataLabels: {
              value: true,
              xValue: true,
              position: "right",
              numberFormat: "0.0",
            },
          },
        ),
      ],
    };
    for (const entry of ["document", "react"] as const) {
      const pptx = await buildInEntry(entry, props);
      const file = await validatePptx(pptx);
      assert.deepEqual(file.errors, []);
      const chart = file.slides[0]?.objects[0]?.chart;
      const labelSettings = (index: number) =>
        chart?.seriesDataLabels?.find((labels) => labels.seriesIndex === index);
      const both = fileLabel(labelSettings(0)?.dataLabels);
      assert.deepEqual(
        [both.value, both.xValue, both.position, both.numberFormat],
        [true, true, "right", "0.0"],
      );
      const yOnly = fileLabel(labelSettings(1)?.dataLabels);
      assert.deepEqual([yOnly.value, yOnly.xValue], [true, false]);
      const xOnly = fileLabel(labelSettings(2)?.dataLabels);
      assert.deepEqual([xOnly.value, xOnly.xValue], [false, true]);
      const overrides = labelSettings(3)?.pointDataLabels;
      assert.deepEqual(Object.keys(overrides ?? {}), ["0", "1", "2"]);
      const pointZero = fileLabel(overrides?.[0]);
      assert.deepEqual([pointZero.value, pointZero.xValue], [true, false]);
      const pointOne = fileLabel(overrides?.[1]);
      assert.deepEqual([pointOne.value, pointOne.xValue], [false, true]);
      const override = fileLabel(overrides?.[2]);
      assert.deepEqual(
        [
          override.value,
          override.xValue,
          override.position,
          override.numberFormat,
        ],
        [true, true, "top", "0"],
      );
      const recovered = fileLabel(labelSettings(4)?.dataLabels);
      assert.deepEqual([recovered.value, recovered.xValue], [false, false]);
      assert.equal(
        fileLabel(labelSettings(4)?.pointDataLabels?.[1]).xValue,
        true,
      );
      const missing = fileLabel(labelSettings(5)?.dataLabels);
      assert.deepEqual([missing.value, missing.xValue], [true, true]);
      assert.deepEqual(chart?.series[5]?.x, [null, 12]);
      assert.deepEqual(chart?.series[5]?.y, [13, null]);

      const xml = chartXml(unzipSync(pptx), 1);
      const seriesBlocks = xml.split("<c:ser>").slice(1);
      assert.ok(seriesBlocks[0]!.includes('<c:showVal val="1"/>'));
      assert.ok(seriesBlocks[0]!.includes('<c:showCatName val="1"/>'));
      assert.ok(seriesBlocks[1]!.includes('<c:showVal val="1"/>'));
      assert.ok(seriesBlocks[1]!.includes('<c:showCatName val="0"/>'));
      assert.ok(seriesBlocks[2]!.includes('<c:showVal val="0"/>'));
      assert.ok(seriesBlocks[2]!.includes('<c:showCatName val="1"/>'));
      const pointBlocks = seriesBlocks[3]!.split("<c:dLbl>").slice(1);
      assert.ok(pointBlocks[0]!.includes('<c:idx val="0"/>'));
      assert.ok(pointBlocks[0]!.includes('<c:showCatName val="0"/>'));
      assert.ok(pointBlocks[1]!.includes('<c:idx val="1"/>'));
      assert.ok(pointBlocks[1]!.includes('<c:showVal val="0"/>'));
      assert.ok(pointBlocks[2]!.includes('<c:idx val="2"/>'));
      assert.ok(pointBlocks[2]!.includes('<c:showCatName val="1"/>'));
      assert.ok(
        pointBlocks[2]!.includes('<c:numFmt formatCode="0" sourceLinked="0"/>'),
      );
      assert.ok(pointBlocks[2]!.includes('<c:dLblPos val="t"/>'));
      assert.ok(seriesBlocks[4]!.includes("<c:dLbls>"));
      const recoveredBlocks = seriesBlocks[4]!.split("<c:dLbl>").slice(1);
      assert.ok(recoveredBlocks[0]!.includes('<c:idx val="1"/>'));
      assert.ok(recoveredBlocks[0]!.includes('<c:showCatName val="1"/>'));
    }
  });

  it("diagnoses invalid and inapplicable scatter declarations", async () => {
    const categoryData = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const base: ChartProps = {
      x: 0,
      y: 0,
      width: 200,
      height: 120,
      groups: [xyGroup([{ name: "S", x: [1], y: [2] }])],
    };
    const cases: readonly [string, ChartProps][] = [
      [
        "invalid-data",
        { ...base, groups: [xyGroup([{ name: "S", x: [1, NaN], y: [2, 3] }])] },
      ],
      [
        "invalid-data",
        {
          ...base,
          groups: [xyGroup([{ name: "S", x: [1, Infinity], y: [2, 3] }])],
        },
      ],
      [
        "invalid-data",
        { ...base, groups: [xyGroup([{ name: "S", x: [1, 2], y: [3] }])] },
      ],
      [
        "invalid-data",
        {
          ...base,
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              series: [{ name: "S", x: "1,2", y: [3, 4] } as never],
            },
          ],
        },
      ],
      [
        "invalid-data",
        { ...base, groups: [xyGroup([{ name: "S", x: [], y: [] }])] },
      ],
      [
        "invalid-data",
        {
          ...base,
          groups: [
            {
              kind: "scatter",
              data: categoryData,
              series: [{ name: "S", x: [1], y: [2] }],
            } as never,
          ],
        },
      ],
      [
        "invalid-data",
        {
          ...base,
          groups: [
            {
              kind: "column",
              data: { kind: "xy" },
              series: [{ name: "S", values: [1] }],
            } as never,
          ],
        },
      ],
      ["invalid-chart-missing-values", { ...base, displayBlanksAs: "span" }],
      [
        "invalid-chart-option",
        {
          ...base,
          groups: [
            {
              kind: "scatter",
              stacking: "stacked",
              data: { kind: "xy" },
              series: [{ name: "S", x: [1], y: [2] }],
            } as never,
          ],
        },
      ],
      [
        "invalid-scatter-presentation",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              line: false,
              markers: false,
            }),
          ],
        },
      ],
      [
        "invalid-marker",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2], marker: true }], {
              line: "straight",
              markers: false,
            }),
          ],
        },
      ],
      [
        "invalid-marker",
        {
          ...base,
          groups: [
            xyGroup(
              [
                {
                  name: "S",
                  x: [1],
                  y: [2],
                  points: { 0: { marker: true } },
                },
              ],
              { line: "straight", markers: false },
            ),
          ],
        },
      ],
      [
        "invalid-marker",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2], marker: { size: 1 } }]),
          ],
        },
      ],
      [
        "invalid-line",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2], line: { width: 2 } }]),
          ],
        },
      ],
      [
        "invalid-data-labels",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              dataLabels: { value: true, category: true } as never,
            }),
          ],
        },
      ],
      [
        "invalid-data-labels",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              dataLabels: { percent: true } as never,
            }),
          ],
        },
      ],
      [
        "invalid-data-labels",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              dataLabels: { value: true, position: "bestFit" } as never,
            }),
          ],
        },
      ],
      [
        "invalid-chart-axis-role",
        {
          ...base,
          axes: { x: { kind: "category" }, y: { kind: "value" } },
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              axes: { x: "x", y: "y" },
            }),
          ],
        },
      ],
      [
        "invalid-chart-axes",
        {
          ...base,
          axes: {
            x: { kind: "value" },
            y: { kind: "value" },
          },
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              axes: { x: "x", y: "x" },
            }),
          ],
        },
      ],
      [
        "invalid-chart-axes",
        {
          ...base,
          axes: {
            horizontal: { kind: "value" },
            vertical: { kind: "value" },
          },
        },
      ],
      [
        "unknown-chart-axis",
        {
          ...base,
          axes: {
            horizontal: { kind: "value" },
            vertical: { kind: "value" },
          },
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              axes: { x: "horizontal", y: "missing" },
            }),
          ],
        },
      ],
      [
        "unsupported-chart-group-combination",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }]),
            {
              kind: "column",
              data: categoryData,
              series: [{ name: "C", values: [1, 2] }],
            },
          ],
        },
      ],
      [
        "invalid-data",
        {
          ...base,
          groups: [
            xyGroup([
              {
                name: "S",
                x: [1],
                y: [2],
                points: { 3: { color: "#000000" } },
              },
            ]),
          ],
        },
      ],
      [
        "invalid-scatter-line",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              line: "curved" as never,
            }),
          ],
        },
      ],
      [
        "invalid-scatter-markers",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2] }], {
              markers: "yes" as never,
            }),
          ],
        },
      ],
    ];
    for (const [code, props] of cases) {
      await rejectsWithCode(props, code);
    }
    await rejectsInBothEntries(
      {
        ...base,
        groups: [
          xyGroup([{ name: "S", x: [1], y: [2] }], {
            dataLabels: { xValue: "yes" } as never,
          }),
        ],
      },
      "invalid-data-labels",
    );
  });

  it("rejects X value labels outside scatter charts in both entries", async () => {
    await rejectsInBothEntries(
      {
        x: 10,
        y: 10,
        width: 320,
        height: 220,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "S", values: [1] }],
            dataLabels: { xValue: true },
          } as never,
        ],
      },
      "invalid-data-labels",
    );
  });

  it("builds scatter series and category batch motion from original observation slots", async () => {
    const result = await build(
      <Deck width={480} height={300}>
        <Slide id="scatter-motion">
          <Chart
            id="scatter-series-motion"
            x={10}
            y={10}
            width={220}
            height={220}
            enter={fade({ by: "series", duration: 400 })}
            groups={[
              xyGroup([
                { name: "North", x: [1, 2, 4], y: [2, 4, 8] },
                { name: "South", x: [1, 3], y: [5, 10] },
              ]),
            ]}
          />
          <Chart
            id="scatter-category-motion"
            x={240}
            y={10}
            width={220}
            height={220}
            enter={fade({ by: "category", duration: 400 })}
            groups={[
              xyGroup([
                { name: "North", x: [1, 2, 4], y: [2, 4, 8] },
                { name: "South", x: [1, 3], y: [5, 10] },
              ]),
            ]}
          />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes('seriesIdx="0"'));
    assert.ok(slideXml.includes('seriesIdx="1"'));
    assert.ok(slideXml.includes('bldStep="series"'));
    assert.ok(slideXml.includes('<a:bldChart bld="series">'));
    assert.ok(slideXml.includes('categoryIdx="0"'));
    assert.ok(slideXml.includes('categoryIdx="2"'));
    assert.ok(slideXml.includes('bldStep="category"'));
    assert.ok(slideXml.includes('<a:bldChart bld="category">'));
  });

  it("diagnoses marker-only span from actual scatter XML independently", async () => {
    const lineGroup = xyGroup([{ name: "Line", x: [1, 2], y: [3, 4] }], {
      line: "straight",
    });
    const markerGroup = xyGroup([{ name: "Marker", x: [1, 2], y: [3, 4] }]);
    const linePptx = await buildInEntry("document", {
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      displayBlanksAs: "span",
      groups: [lineGroup],
    });
    assert.deepEqual((await validatePptx(linePptx)).errors, []);

    const markerPptx = await buildInEntry("document", {
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      groups: [markerGroup],
    });
    const markerParts = unzipSync(markerPptx);
    const path = "ppt/charts/chart1.xml";
    const withSpan = (parts: Unzipped) =>
      zipSync({
        ...parts,
        [path]: strToU8(
          chartXml(parts, 1).replace(
            "</c:chart>",
            '<c:dispBlanksAs val="span"/></c:chart>',
          ),
        ),
      });
    const markerFile = await validatePptx(withSpan(markerParts));
    assert.ok(
      markerFile.errors.some(
        ({ code }) => code === "invalid-chart-missing-values",
      ),
    );

    const mixedPptx = await buildInEntry("document", {
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      groups: [lineGroup, markerGroup],
    });
    const mixedFile = await validatePptx(withSpan(unzipSync(mixedPptx)));
    assert.ok(
      mixedFile.errors.some(
        ({ code }) => code === "invalid-chart-missing-values",
      ),
    );

    const markerGap = await buildInEntry("document", {
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      displayBlanksAs: "gap",
      groups: [markerGroup],
    });
    assert.deepEqual((await validatePptx(markerGap)).errors, []);
  });

  it("reads scatter caches independently and rejects malformed native scatter XML", async () => {
    const result = await buildDocument(
      deck({ width: 360, height: 240 }).append(
        documentSlide({ id: "scatter-reader" }).append(
          documentChart({
            id: "scatter-reader",
            x: 10,
            y: 10,
            width: 300,
            height: 200,
            groups: [xyGroup([{ name: "S", x: [1, 2], y: [3, 4] }])],
          }),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const path = "ppt/charts/chart1.xml";
    const original = chartXml(parts, 1);

    const badStyle = original.replace(
      '<c:scatterStyle val="marker"/>',
      '<c:scatterStyle val="bogus"/>',
    );
    const styleErrors = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(badStyle) }),
    );
    assert.ok(
      styleErrors.errors.some(
        ({ code }) => code === "invalid-chart-scatter-style",
      ),
    );

    const badPoint = original.replace("<c:v>1</c:v>", "<c:v>NaN</c:v>");
    const pointErrors = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(badPoint) }),
    );
    assert.ok(
      pointErrors.errors.some(({ code }) => code === "invalid-chart-cache"),
    );

    const categorySeries = original.replace(
      "<c:xVal>",
      '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/></c:strCache></c:strRef></c:cat><c:xVal>',
    );
    const categoryErrors = await validatePptx(
      zipSync({ ...parts, [path]: strToU8(categorySeries) }),
    );
    assert.ok(
      categoryErrors.errors.some(({ code }) => code === "invalid-chart-series"),
    );
  });
});
