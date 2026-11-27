import assert from "node:assert/strict";

import {
  type ChartProps,
  type ScatterChartGroup,
} from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@pptxdk/capabilities/document";
import { fade } from "@pptxdk/capabilities/motion";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx, type FileChartDataLabel } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

function chartXml(parts: Unzipped, index: number): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

function scatterAxisNumberFormat(xml: string, direction: "x" | "y"): string {
  const scatter = xml
    .split("<c:scatterChart>")[1]
    ?.split("</c:scatterChart>")[0];
  assert.ok(scatter);
  const axisId = [...scatter.matchAll(/<c:axId val="(\d+)"\/>/gu)][
    direction === "x" ? 0 : 1
  ]?.[1];
  assert.ok(axisId);
  const axis = [...xml.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/gu)]
    .map((match) => match[1]!)
    .find((value) => value.includes(`<c:axId val="${axisId}"/>`));
  assert.ok(axis);
  const numberFormat = axis.match(
    /<c:numFmt formatCode="([^"]+)" sourceLinked="([01])"\/>/u,
  );
  assert.ok(numberFormat);
  return `${numberFormat[1]}/${numberFormat[2]}`;
}

function workbookCellXml(
  workbook: Unzipped,
  reference: string,
): string | undefined {
  const worksheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
  return [...worksheet.matchAll(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/gu)]
    .map((match) => match[0])
    .find((cell) => cell.match(/(?:^|\s)r="([^"]+)"/u)?.[1] === reference);
}

function workbookCellNumberFormat(
  workbook: Unzipped,
  reference: string,
): string | undefined {
  const cell = workbookCellXml(workbook, reference);
  if (cell === undefined) return undefined;
  const styleIndex = Number(cell.match(/\bs="(\d+)"/u)?.[1] ?? 0);
  const styles = strFromU8(workbook["xl/styles.xml"]!);
  const cellXfs = styles.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/u)?.[1];
  assert.ok(cellXfs);
  const xf = [...cellXfs.matchAll(/<xf\b([^>]*)\/>/gu)][styleIndex]?.[1];
  assert.ok(xf);
  const numberFormatId = xf.match(/\bnumFmtId="(\d+)"/u)?.[1] ?? "0";
  if (numberFormatId === "0") return "General";
  return [...styles.matchAll(/<numFmt\b([^>]*)\/>/gu)]
    .map((match) => match[1]!)
    .find((attributes) => attributes.includes(`numFmtId="${numberFormatId}"`))
    ?.match(/\bformatCode="([^"]+)"/u)?.[1];
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
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 420,
        height: 260,
      },
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
        style: {
          position: "absolute",
          left: x,
          top: 10,
          width: 300,
          height: 220,
        },
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

  it("disables connecting lines per series for marker-only scatter groups", async () => {
    const series: ScatterChartGroup["series"] = [
      {
        name: "Observed",
        x: [1, 2.5, 4, 8],
        y: [3, 6, 9, 12],
        color: "#336699",
        marker: { symbol: "circle", size: 8 },
        points: {
          1: { marker: { symbol: "star", size: 11 }, color: "#AA5500" },
        },
      },
      {
        name: "Short",
        x: [1.5, 6, 9],
        y: [2, 5, 7],
        marker: { symbol: "square", size: 7 },
      },
    ];
    const chart = (
      id: string,
      x: number,
      extra: Pick<ScatterChartGroup, "line" | "markers"> = {},
    ) => {
      const chartSeries =
        extra.markers === false
          ? series.map((item) => ({
              name: item.name,
              x: item.x,
              y: item.y,
              ...(item.color === undefined ? {} : { color: item.color }),
            }))
          : series;
      return documentChart({
        id,
        style: {
          position: "absolute",
          left: x,
          top: 10,
          width: 300,
          height: 220,
        },
        groups: [xyGroup(chartSeries, extra)],
      });
    };
    const result = await buildDocument(
      deck({ width: 1280, height: 260 }).append(
        documentSlide({ id: "marker-only-lines" }).append(
          chart("omitted", 10),
          chart("explicit-false", 330, { line: false }),
          chart("straight-markers", 650, { line: "straight" }),
          chart("smooth-no-markers", 970, { line: "smooth", markers: false }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const charts = file.slides[0]?.objects.map(({ chart }) => chart);
    assert.deepEqual(
      charts?.map((chart) => chart?.groups[0]?.scatterStyle),
      ["marker", "marker", "lineMarker", "smooth"],
    );
    for (const chart of charts?.slice(0, 2) ?? []) {
      assert.deepEqual(chart?.series[0]?.x, [1, 2.5, 4, 8]);
      assert.deepEqual(chart?.series[0]?.y, [3, 6, 9, 12]);
      assert.deepEqual(chart?.series[1]?.x, [1.5, 6, 9]);
      assert.deepEqual(chart?.series[0]?.marker, {
        symbol: "circle",
        size: 8,
      });
      assert.deepEqual(chart?.series[1]?.marker, {
        symbol: "square",
        size: 7,
      });
    }

    const parts = unzipSync(result.pptx);
    const noLine = "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>";
    for (const index of [1, 2]) {
      const xml = chartXml(parts, index);
      const seriesXml = xml.split("<c:ser>").slice(1);
      assert.equal(seriesXml.length, 2);
      for (const [seriesIndex, markerSeriesXml] of seriesXml.entries()) {
        assert.ok(
          markerSeriesXml.includes(noLine),
          `chart ${index} series ${seriesIndex} should disable its line`,
        );
        assert.ok(
          markerSeriesXml.indexOf(noLine) <
            markerSeriesXml.indexOf("<c:marker>"),
        );
      }
      assert.ok(
        xml.includes('<c:idx val="1"/><c:marker><c:symbol val="star"/>'),
      );
      assert.ok(xml.includes('<a:srgbClr val="AA5500"/>'));
    }
    for (const index of [3, 4]) {
      const xml = chartXml(parts, index);
      assert.ok(!xml.includes(noLine));
      assert.ok(xml.includes("<c:spPr><a:ln "));
    }
    assert.ok(chartXml(parts, 4).includes('<c:smooth val="1"/>'));
  });

  it("keeps missing X and missing Y in their original slots with point formatting", async () => {
    const result = await buildDocument(
      deck({ width: 640, height: 400 }).append(
        documentSlide({ id: "missing" }).append(
          documentChart({
            id: "missing",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 500,
              height: 300,
            },
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
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 400,
              height: 240,
            },
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
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 320,
            height: 220,
          },
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
      style: {
        position: "absolute" as const,
        left: 10,
        top: 10,
        width: 320,
        height: 220,
      },
      displayBlanksAs: "span" as const,
    };
    const lineGroup = xyGroup([{ name: "Line", x: [1, 2], y: [3, 4] }], {
      line: "straight",
    });
    const markerGroup = xyGroup([{ name: "Marker", x: [1, 2], y: [3, 4] }]);
    await rejectsInBothEntries(
      { ...base, groups: [markerGroup] },
      "invalid-prop",
    );
    await rejectsInBothEntries(
      { ...base, groups: [lineGroup, markerGroup] },
      "invalid-prop",
    );
    await rejectsInBothEntries(
      { ...base, groups: [markerGroup, lineGroup] },
      "invalid-prop",
    );
  });

  it("writes scatter X and Y label content through group, series, and point inheritance", async () => {
    const props: ChartProps = {
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 340,
        height: 240,
      },
      groups: [
        xyGroup(
          [
            {
              name: "Both",
              x: [1, 2, 3, 4],
              y: [10, 20, 30, 40],
              points: {
                2: { dataLabels: { xValue: false } },
                3: { dataLabels: { numberFormat: "0.00" } },
              },
            },
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
                2: { dataLabels: { numberFormat: "0.00", position: "top" } },
              },
            },
            {
              name: "Recovered point",
              x: [13, 14],
              y: [15, 16],
              dataLabels: false,
              points: {
                0: { dataLabels: { xValue: true, numberFormat: "0.00" } },
                1: { dataLabels: { xValue: true } },
              },
            },
            {
              name: "Missing coordinates",
              x: [null, 0],
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
        [
          both.value,
          both.xValue,
          both.position,
          both.numberFormat,
          both.percent,
        ],
        [true, true, "right", "0.0", false],
      );
      const bothPoints = labelSettings(0)?.pointDataLabels;
      assert.deepEqual(Object.keys(bothPoints ?? {}), ["2", "3"]);
      assert.equal(fileLabel(bothPoints?.[2]).xValue, false);
      assert.equal(fileLabel(bothPoints?.[3]).numberFormat, "0.00");
      const yOnly = fileLabel(labelSettings(1)?.dataLabels);
      assert.deepEqual(
        [yOnly.value, yOnly.xValue, yOnly.percent],
        [true, false, false],
      );
      const xOnly = fileLabel(labelSettings(2)?.dataLabels);
      assert.deepEqual(
        [xOnly.value, xOnly.xValue, xOnly.percent],
        [false, true, false],
      );
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
          override.percent,
        ],
        [true, true, "top", "0.00", false],
      );
      const recovered = fileLabel(labelSettings(4)?.dataLabels);
      assert.deepEqual([recovered.value, recovered.xValue], [false, false]);
      assert.equal(
        fileLabel(labelSettings(4)?.pointDataLabels?.[0]).numberFormat,
        "0.00",
      );
      assert.equal(
        fileLabel(labelSettings(4)?.pointDataLabels?.[0]).xValue,
        true,
      );
      assert.equal(
        fileLabel(labelSettings(4)?.pointDataLabels?.[1]).xValue,
        true,
      );
      const missing = fileLabel(labelSettings(5)?.dataLabels);
      assert.deepEqual([missing.value, missing.xValue], [true, true]);
      assert.deepEqual(chart?.series[5]?.x, [null, 0]);
      assert.deepEqual(chart?.series[5]?.y, [13, null]);

      const parts = unzipSync(pptx);
      const xml = chartXml(parts, 1);
      const seriesBlocks = xml.split("<c:ser>").slice(1);
      assert.ok(seriesBlocks[0]!.includes('<c:showVal val="1"/>'));
      assert.ok(seriesBlocks[0]!.includes('<c:showCatName val="1"/>'));
      assert.ok(seriesBlocks[0]!.includes('<c:showPercent val="0"/>'));
      assert.ok(seriesBlocks[1]!.includes('<c:showVal val="1"/>'));
      assert.ok(seriesBlocks[1]!.includes('<c:showCatName val="0"/>'));
      assert.ok(seriesBlocks[1]!.includes('<c:showPercent val="0"/>'));
      assert.ok(seriesBlocks[2]!.includes('<c:showVal val="0"/>'));
      assert.ok(seriesBlocks[2]!.includes('<c:showCatName val="1"/>'));
      assert.ok(seriesBlocks[2]!.includes('<c:showPercent val="0"/>'));
      const pointBlocks = seriesBlocks[3]!.split("<c:dLbl>").slice(1);
      assert.ok(pointBlocks[0]!.includes('<c:idx val="0"/>'));
      assert.ok(pointBlocks[0]!.includes('<c:showCatName val="0"/>'));
      assert.ok(pointBlocks[1]!.includes('<c:idx val="1"/>'));
      assert.ok(pointBlocks[1]!.includes('<c:showVal val="0"/>'));
      assert.ok(pointBlocks[2]!.includes('<c:idx val="2"/>'));
      assert.ok(pointBlocks[2]!.includes('<c:showCatName val="1"/>'));
      assert.ok(
        pointBlocks[2]!.includes(
          '<c:numFmt formatCode="0.00" sourceLinked="0"/>',
        ),
      );
      assert.ok(pointBlocks[2]!.includes('<c:showPercent val="0"/>'));
      assert.ok(pointBlocks[2]!.includes('<c:dLblPos val="t"/>'));
      assert.ok(seriesBlocks[4]!.includes("<c:dLbls>"));
      const recoveredBlocks = seriesBlocks[4]!.split("<c:dLbl>").slice(1);
      assert.ok(recoveredBlocks[0]!.includes('<c:idx val="0"/>'));
      assert.ok(recoveredBlocks[0]!.includes('<c:showCatName val="1"/>'));
      assert.ok(
        recoveredBlocks[0]!.includes(
          '<c:numFmt formatCode="0.00" sourceLinked="0"/>',
        ),
      );

      const xCache = (seriesIndex: number) =>
        seriesBlocks[seriesIndex]!.split("<c:xVal>")[1]!.split("</c:xVal>")[0]!;
      assert.ok(xCache(0).includes("<c:formatCode>0.0</c:formatCode>"));
      assert.ok(
        xCache(0).includes(
          '<c:pt idx="2" formatCode="General"><c:v>3</c:v></c:pt>',
        ),
      );
      assert.ok(
        xCache(0).includes(
          '<c:pt idx="3" formatCode="0.00"><c:v>4</c:v></c:pt>',
        ),
      );
      assert.ok(xCache(0).includes("<c:f>Sheet1!$B$2:$B$5</c:f>"));
      assert.ok(xCache(1).includes("<c:formatCode>General</c:formatCode>"));
      assert.ok(xCache(3).includes("<c:formatCode>0.0</c:formatCode>"));
      assert.ok(
        xCache(3).includes(
          '<c:pt idx="0" formatCode="General"><c:v>8</c:v></c:pt>',
        ),
      );
      assert.ok(
        xCache(3).includes(
          '<c:pt idx="2" formatCode="0.00"><c:v>10</c:v></c:pt>',
        ),
      );
      assert.ok(xCache(4).includes("<c:formatCode>General</c:formatCode>"));
      assert.ok(
        xCache(4).includes(
          '<c:pt idx="0" formatCode="0.00"><c:v>13</c:v></c:pt>',
        ),
      );
      assert.ok(xCache(5).includes("<c:formatCode>0.0</c:formatCode>"));
      assert.ok(xCache(5).includes("<c:f>Sheet1!$L$2:$L$3</c:f>"));
      assert.ok(xCache(5).includes('<c:ptCount val="2"/>'));
      assert.ok(!xCache(5).includes('<c:pt idx="0"'));
      assert.ok(xCache(5).includes('<c:pt idx="1"><c:v>0</c:v></c:pt>'));
      const yCache = seriesBlocks[0]!
        .split("<c:yVal>")[1]!
        .split("</c:yVal>")[0]!;
      assert.ok(yCache.includes("<c:formatCode>General</c:formatCode>"));

      const workbookPath = chart?.embeddedWorkbook?.part;
      assert.ok(workbookPath);
      const workbook = unzipSync(parts[workbookPath]!);
      assert.equal(workbookCellNumberFormat(workbook, "B2"), "0.0");
      assert.equal(workbookCellNumberFormat(workbook, "B4"), "General");
      assert.equal(workbookCellNumberFormat(workbook, "B5"), "0.00");
      assert.equal(workbookCellNumberFormat(workbook, "D2"), "General");
      assert.equal(workbookCellNumberFormat(workbook, "C2"), "General");
      assert.equal(workbookCellNumberFormat(workbook, "H2"), "General");
      assert.equal(workbookCellNumberFormat(workbook, "H4"), "0.00");
      assert.equal(workbookCellNumberFormat(workbook, "J2"), "0.00");
      assert.equal(workbookCellNumberFormat(workbook, "J3"), "General");
      assert.equal(workbookCellNumberFormat(workbook, "L2"), "0.0");
      assert.equal(workbookCellNumberFormat(workbook, "L3"), "0.0");
      assert.equal(workbookCellXml(workbook, "L2"), '<c r="L2" s="1"/>');
      assert.equal(
        workbookCellXml(workbook, "L3"),
        '<c r="L3" s="1"><v>0</v></c>',
      );
      assert.deepEqual(chart?.series[5]?.x, [null, 0]);
      assert.deepEqual(chart?.series[5]?.y, [13, null]);
      assert.deepEqual(chart?.series[0]?.x, [1, 2, 3, 4]);
      assert.deepEqual(chart?.embeddedWorkbook?.series[0]?.x, [1, 2, 3, 4]);
      assert.deepEqual(chart?.embeddedWorkbook?.series[5]?.x, [null, 0]);
      assert.deepEqual(chart?.embeddedWorkbook?.series[5]?.y, [13, null]);
    }
  });

  it("isolates only projected scatter X formats from the default X axis", async () => {
    const projectedSeries = {
      name: "Projected X",
      x: [1, 2],
      y: [3, 4],
      dataLabels: { value: false, xValue: true, numberFormat: "0.0" },
    };
    const projected: ChartProps = {
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 340,
        height: 240,
      },
      groups: [xyGroup([projectedSeries])],
    };
    const explicit: ChartProps = {
      ...projected,
      axes: {
        horizontal: { kind: "value", numberFormat: "0.00" },
        vertical: { kind: "value", numberFormat: "0.000" },
      },
      groups: [
        xyGroup([projectedSeries], {
          axes: { x: "horizontal", y: "vertical" },
        }),
      ],
    };
    const noProjection: ChartProps = {
      ...projected,
      groups: [
        xyGroup([
          {
            name: "General X",
            x: [1, 2],
            y: [3, 4],
            dataLabels: { value: false, xValue: true },
          },
        ]),
      ],
    };

    for (const entry of ["document", "react"] as const) {
      const projectedXml = chartXml(
        unzipSync(await buildInEntry(entry, projected)),
        1,
      );
      assert.equal(scatterAxisNumberFormat(projectedXml, "x"), "General/0");
      assert.equal(scatterAxisNumberFormat(projectedXml, "y"), "General/1");

      const explicitXml = chartXml(
        unzipSync(await buildInEntry(entry, explicit)),
        1,
      );
      assert.equal(scatterAxisNumberFormat(explicitXml, "x"), "0.00/0");
      assert.equal(scatterAxisNumberFormat(explicitXml, "y"), "0.000/0");

      const noProjectionXml = chartXml(
        unzipSync(await buildInEntry(entry, noProjection)),
        1,
      );
      assert.equal(scatterAxisNumberFormat(noProjectionXml, "x"), "General/1");
      assert.equal(scatterAxisNumberFormat(noProjectionXml, "y"), "General/1");
    }
  });

  it("diagnoses invalid and inapplicable scatter declarations", async () => {
    const categoryData = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const base: ChartProps = {
      style: { position: "absolute", left: 0, top: 0, width: 200, height: 120 },
      groups: [xyGroup([{ name: "S", x: [1], y: [2] }])],
    };
    const cases: readonly [string, ChartProps][] = [
      [
        "invalid-prop",
        { ...base, groups: [xyGroup([{ name: "S", x: [1, NaN], y: [2, 3] }])] },
      ],
      [
        "invalid-prop",
        {
          ...base,
          groups: [xyGroup([{ name: "S", x: [1, Infinity], y: [2, 3] }])],
        },
      ],
      [
        "invalid-prop",
        { ...base, groups: [xyGroup([{ name: "S", x: [1, 2], y: [3] }])] },
      ],
      [
        "invalid-prop",
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
        "invalid-prop",
        { ...base, groups: [xyGroup([{ name: "S", x: [], y: [3] }])] },
      ],
      [
        "invalid-prop",
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
        "invalid-prop",
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
      ["invalid-prop", { ...base, displayBlanksAs: "span" }],
      [
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2], marker: { size: 1 } }]),
          ],
        },
      ],
      [
        "invalid-prop",
        {
          ...base,
          groups: [
            xyGroup([{ name: "S", x: [1], y: [2], line: { width: 2 } }]),
          ],
        },
      ],
      [
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
        {
          ...base,
          axes: {
            horizontal: { kind: "value" },
            vertical: { kind: "value" },
          },
        },
      ],
      [
        "broken-reference",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
        "invalid-prop",
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
      "invalid-prop",
    );
  });

  it("rejects X value labels outside scatter charts in both entries", async () => {
    await rejectsInBothEntries(
      {
        style: {
          position: "absolute",
          left: 10,
          top: 10,
          width: 320,
          height: 220,
        },
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "S", values: [1] }],
            dataLabels: { xValue: true },
          } as never,
        ],
      },
      "invalid-prop",
    );
  });

  it("builds scatter series and category batch motion from original observation slots", async () => {
    const result = await build(
      <Deck width={480} height={300}>
        <Slide id="scatter-motion">
          <Chart
            id="scatter-series-motion"
            style={{
              position: "absolute",
              left: 10,
              top: 10,
              width: 220,
              height: 220,
            }}
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
            style={{
              position: "absolute",
              left: 240,
              top: 10,
              width: 220,
              height: 220,
            }}
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
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 320,
        height: 220,
      },
      displayBlanksAs: "span",
      groups: [lineGroup],
    });
    assert.deepEqual((await validatePptx(linePptx)).errors, []);

    const markerPptx = await buildInEntry("document", {
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 320,
        height: 220,
      },
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
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 320,
        height: 220,
      },
      groups: [lineGroup, markerGroup],
    });
    const mixedFile = await validatePptx(withSpan(unzipSync(mixedPptx)));
    assert.ok(
      mixedFile.errors.some(
        ({ code }) => code === "invalid-chart-missing-values",
      ),
    );

    const markerGap = await buildInEntry("document", {
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 320,
        height: 220,
      },
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
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 300,
              height: 200,
            },
            groups: [xyGroup([{ name: "S", x: [1, 2], y: [3, 4] }])],
          }),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const workbookPart =
      file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const workbook = unzipSync(parts[workbookPart]!);
    const worksheetPath = "xl/worksheets/sheet1.xml";
    const worksheet = strFromU8(workbook[worksheetPath]!);
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

    const badSharedString = worksheet.replace(
      '<c r="B2"><v>1</v></c>',
      '<c r="B2" t="s"><v>999</v></c>',
    );
    assert.notEqual(badSharedString, worksheet);
    const sharedStringErrors = await validatePptx(
      zipSync({
        ...parts,
        [workbookPart]: zipSync({
          ...workbook,
          [worksheetPath]: strToU8(badSharedString),
        }),
      }),
    );
    assert.ok(
      sharedStringErrors.errors.some(
        ({ code, message }) =>
          code === "invalid-embedded-workbook-value" &&
          message.includes("B2 shared string index is invalid"),
      ),
    );

    const badNumber = worksheet.replace(
      '<c r="B2"><v>1</v></c>',
      '<c r="B2"><v>not-a-number</v></c>',
    );
    assert.notEqual(badNumber, worksheet);
    const numberErrors = await validatePptx(
      zipSync({
        ...parts,
        [workbookPart]: zipSync({
          ...workbook,
          [worksheetPath]: strToU8(badNumber),
        }),
      }),
    );
    assert.ok(
      numberErrors.errors.some(
        ({ code, message }) =>
          code === "invalid-embedded-workbook-value" &&
          message.includes("B2 is not a finite number"),
      ),
    );
  });
});
