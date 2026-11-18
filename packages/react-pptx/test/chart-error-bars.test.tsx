import assert from "node:assert/strict";

import {
  BuildError,
  build,
  Chart,
  Deck,
  Slide,
  type ChartProps,
  type ChartGroup,
  type ChartSeries,
  type ScatterChartGroup,
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

function chartParts(pptx: Uint8Array): {
  readonly chart: (index: number) => string;
  readonly workbook: (part: string) => string;
} {
  const parts = unzipSync(pptx);
  return {
    chart: (index) => strFromU8(parts[`ppt/charts/chart${index}.xml`]!),
    workbook: (part) =>
      strFromU8(unzipSync(parts[part]!)["xl/worksheets/sheet1.xml"]!),
  };
}

async function buildInEntry(
  entry: "document" | "react",
  props: ChartProps,
): Promise<Uint8Array> {
  const id = "error-bars-entry";
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

function columnChart(
  series: readonly ChartSeries[],
  kind: "column" | "bar" | "line" = "column",
): ChartProps {
  return {
    x: 0,
    y: 0,
    width: 400,
    height: 260,
    groups: [
      {
        kind,
        data: {
          kind: "shared-categories",
          categories: ["North", "South", "East"],
        },
        series,
      } as ChartGroup,
    ],
  };
}

async function rejectsWithCode(props: ChartProps, code: string): Promise<void> {
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

function areaSeries(file: Awaited<ReturnType<typeof validatePptx>>) {
  return file.slides[0]?.objects[0]?.chart?.series[0];
}

function stackedSeries(file: Awaited<ReturnType<typeof validatePptx>>) {
  return file.slides[1]?.objects[0]?.chart?.series[0];
}

describe("native chart error bars", () => {
  it("writes a fixed category error bar without forcing an invalid direction", async () => {
    const props = columnChart([
      {
        name: "Revenue",
        values: [1, 2, 3],
        errorBars: { type: "fixed", value: 0.5 },
      },
    ]);
    const result = await build(
      <Deck width={480} height={320}>
        <Slide id="error-bars">
          <Chart {...props} id="error-bars" />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const series = file.slides[0]?.objects[0]?.chart?.series[0];
    assert.deepEqual(series?.errorBars, [
      { direction: "y", include: "both", type: "fixed", value: 0.5 },
    ]);

    const xml = chartParts(result.pptx).chart(1);
    assert.ok(!xml.includes("<c:errDir"));
    assert.ok(xml.includes("<c:errBarType"));
    assert.ok(xml.includes("<c:errValType"));
    assert.ok(xml.includes("<c:val"));
  });

  it("keeps bar direction horizontal and line direction vertical with caps and stroke", async () => {
    const bar = columnChart(
      [
        {
          name: "Revenue",
          values: [1, 2, 3],
          errorBars: {
            type: "fixed",
            value: 2,
            endCap: false,
            line: { color: "#112233", width: 2, dash: "dash" },
          },
        },
      ],
      "bar",
    );
    const line = columnChart(
      [
        {
          name: "Revenue",
          values: [1, 2, 3],
          errorBars: { type: "percentage", value: 10, include: "plus" },
        },
      ],
      "line",
    );
    const result = await build(
      <Deck width={480} height={640}>
        <Slide id="bar">
          <Chart {...bar} id="bar" />
        </Slide>
        <Slide id="line">
          <Chart {...line} id="line" />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.chart?.series[0]?.errorBars, [
      {
        direction: "x",
        include: "both",
        type: "fixed",
        endCap: false,
        value: 2,
        line: { color: "#112233", widthEmu: 19_050, dash: "dash" },
      },
    ]);
    assert.deepEqual(file.slides[1]?.objects[0]?.chart?.series[0]?.errorBars, [
      {
        direction: "y",
        declaredDirection: "y",
        include: "plus",
        type: "percentage",
        value: 10,
      },
    ]);
    const parts = chartParts(result.pptx);
    const barXml = parts.chart(1);
    assert.ok(!barXml.includes("<c:errDir"));
    assert.ok(barXml.includes('<c:noEndCap val="1"/>'));
    assert.ok(barXml.includes('<a:ln w="19050">'));
    assert.ok(barXml.includes('<a:prstDash val="dash"/>'));
    const lineXml = parts.chart(2);
    assert.ok(lineXml.includes('<c:errDir val="y"/>'));
    assert.ok(!lineXml.includes("<c:noEndCap"));
    assert.ok(lineXml.includes('<c:errBarType val="plus"/>'));
    assert.ok(lineXml.includes('<c:errValType val="percentage"/>'));
  });

  it("writes independent scatter X and Y custom channels without collapsing null slots", async () => {
    const group: ScatterChartGroup = {
      kind: "scatter",
      data: { kind: "xy" },
      series: [
        {
          name: "Observed",
          x: [1, 2, 7],
          y: [3, 5, null],
          errorBars: {
            x: { type: "fixed", value: 0.5 },
            y: {
              type: "custom",
              plus: [0.4, null, 0.8],
              minus: [0.1, 0.2, 0.3],
              endCap: false,
              line: { color: "#446688", width: 2, dash: "dash" },
            },
          },
        },
        {
          name: "Short",
          x: [1.5, 6],
          y: [2, 7],
          errorBars: { y: { type: "standardDeviation", value: 2 } },
        },
      ],
    };
    const props: ChartProps = {
      x: 0,
      y: 0,
      width: 420,
      height: 260,
      groups: [group],
    };
    const result = await build(
      <Deck width={480} height={320}>
        <Slide id="scatter">
          <Chart {...props} id="scatter" />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.errorBars, [
      {
        direction: "x",
        declaredDirection: "x",
        include: "both",
        type: "fixed",
        value: 0.5,
      },
      {
        direction: "y",
        declaredDirection: "y",
        include: "both",
        type: "custom",
        endCap: false,
        line: { color: "#446688", widthEmu: 19_050, dash: "dash" },
        plus: {
          source: "reference",
          formula: "Sheet1!$F$2:$F$4",
          values: [0.4, null, 0.8],
        },
        minus: {
          source: "reference",
          formula: "Sheet1!$G$2:$G$4",
          values: [0.1, 0.2, 0.3],
        },
      },
    ]);
    assert.deepEqual(chart?.series[1]?.errorBars, [
      {
        direction: "y",
        declaredDirection: "y",
        include: "both",
        type: "standardDeviation",
        value: 2,
      },
    ]);
    assert.deepEqual(
      chart?.embeddedWorkbook?.series[0]?.errorBars?.[1]?.plus?.values,
      [0.4, null, 0.8],
    );
    assert.equal(
      chart?.embeddedWorkbook?.series[0]?.errorBars?.[1]?.plus?.formula,
      "Sheet1!$F$2:$F$4",
    );

    const parts = chartParts(result.pptx);
    const xml = parts.chart(1);
    assert.ok(xml.includes('<c:errDir val="x"/>'));
    assert.ok(xml.includes('<c:errDir val="y"/>'));
    assert.ok(xml.includes("<c:plus>"));
    assert.ok(xml.includes("<c:minus>"));
    assert.ok(xml.includes('<c:ptCount val="3"/>'));
    assert.ok(xml.includes('<c:pt idx="0">'));
    assert.ok(xml.includes('<c:pt idx="2">'));
    assert.ok(!xml.includes('<c:pt idx="1"><c:v>0.4'));
    const workbook = chart?.embeddedWorkbook?.part;
    assert.ok(workbook);
    const sheet = parts.workbook(workbook);
    assert.ok(sheet.includes("Observed Y +"));
    assert.ok(sheet.includes("Observed Y -"));
    assert.ok(sheet.includes('r="F2"'));
    assert.ok(sheet.includes('r="G4"'));
    assert.ok(!sheet.includes('r="F3"'));
  });

  it("keeps declared error values after the source arrays are mutated", async () => {
    const plus = [0.4, 0.5, 0.6];
    const minus = [0.1, 0.2, 0.3];
    const series: ChartSeries = {
      name: "Snapshotted",
      values: [1, 2, 3],
      errorBars: { type: "custom", plus, minus },
    };
    const node = documentChart(columnChart([series]));
    const document = deck({ width: 480, height: 320 }).append(
      documentSlide({ id: "snapshot" }).append(node),
    );
    plus[0] = 99;
    minus[2] = 99;
    const first = await buildDocument(document);
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.plus
        ?.values,
      [0.4, 0.5, 0.6],
    );
    assert.deepEqual(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.minus
        ?.values,
      [0.1, 0.2, 0.3],
    );

    node.set({
      groups: [
        {
          kind: "column",
          data: {
            kind: "shared-categories",
            categories: ["North", "South", "East"],
          },
          series: [
            {
              name: "Snapshotted",
              values: [1, 2, 3],
              errorBars: { type: "fixed", value: 7 },
            },
          ],
        },
      ],
    });
    const second = await buildDocument(document);
    const secondFile = await validatePptx(second.pptx);
    assert.deepEqual(
      secondFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars,
      [{ direction: "y", include: "both", type: "fixed", value: 7 }],
    );
    assert.deepEqual(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.plus
        ?.values,
      [0.4, 0.5, 0.6],
    );
  });

  it("rejects inapplicable charts and invalid error-bar declarations", async () => {
    const base = columnChart([{ name: "Only", values: [1, 2, 3] }]);
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "pie",
            data: {
              kind: "shared-categories",
              categories: ["A", "B", "C"],
            },
            series: [
              {
                name: "Only",
                values: [1, 2, 3],
                errorBars: { type: "fixed", value: 1 },
              },
            ],
          },
        ],
      },
      "invalid-error-bars",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "scatter",
            data: { kind: "xy" },
            series: [
              {
                name: "Only",
                x: [1, 2],
                y: [1, 2],
                errorBars: {},
              },
            ],
          },
        ],
      },
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "custom", plus: [1], minus: [1, 2, 3] },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "custom", plus: [1, 2, -3], minus: [1, 2, 3] },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: {
            type: "custom",
            include: "plus",
            minus: [1, 2, 3],
          } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "standardError", value: 2 } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "fixed" } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "fixed", value: 1, plus: [1, 1, 1] } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: {
            type: "fixed",
            value: 1,
            x: { type: "fixed", value: 1 },
          } as never,
        },
      ]),
      "unknown-parameter",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "fixed", value: -1 },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "percentage", value: Number.NaN },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: {
            type: "custom",
            plus: Array<number | null>(3),
            minus: [1, 2, 3],
          },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: {
            type: "custom",
            plus: [1, Number.POSITIVE_INFINITY, 3],
            minus: [1, 2, 3],
          },
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "custom", plus: [1, 2, 3] } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "fixed", value: 1, include: "sideways" } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "bogus", value: 1 } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: { type: "fixed", value: 1, endCap: "yes" } as never,
        },
      ]),
      "invalid-error-bars",
    );
    await rejectsWithCode(
      columnChart([
        {
          name: "Only",
          values: [1, 2, 3],
          errorBars: {
            type: "fixed",
            value: 1,
            line: { width: 0 },
          },
        },
      ]),
      "invalid-style",
    );
  });

  it("writes standard error, single-sided custom bars, and stacked groups", async () => {
    const area: ChartProps = {
      x: 0,
      y: 0,
      width: 420,
      height: 260,
      groups: [
        {
          kind: "area",
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q2", "Q3"],
          },
          series: [
            {
              name: "Area",
              values: [1, 2, 3],
              errorBars: { type: "standardError", include: "plus" },
            },
          ],
        },
      ],
    };
    const stacked: ChartProps = {
      x: 0,
      y: 0,
      width: 420,
      height: 260,
      groups: [
        {
          kind: "column",
          stacking: "stacked",
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q2", "Q3"],
          },
          series: [
            {
              name: "Stacked",
              values: [1, 2, 3],
              errorBars: {
                type: "custom",
                include: "minus",
                minus: [0.1, 0.2, 0.3],
              },
            },
          ],
        },
      ],
    };
    const result = await build(
      <Deck width={480} height={640}>
        <Slide id="area">
          <Chart {...area} id="area" />
        </Slide>
        <Slide id="stacked">
          <Chart {...stacked} id="stacked" />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(areaSeries(file)?.errorBars, [
      {
        direction: "y",
        include: "plus",
        type: "standardError",
      },
    ]);
    assert.deepEqual(stackedSeries(file)?.errorBars, [
      {
        direction: "y",
        include: "minus",
        type: "custom",
        minus: {
          source: "reference",
          formula: "Sheet1!$C$2:$C$4",
          values: [0.1, 0.2, 0.3],
        },
      },
    ]);
    const parts = chartParts(result.pptx);
    const areaXml = parts.chart(1);
    assert.ok(areaXml.includes('<c:errValType val="stdErr"/>'));
    assert.ok(!areaXml.includes("<c:val val="));
    assert.ok(areaXml.includes('<c:errBarType val="plus"/>'));
    const stackedXml = parts.chart(2);
    assert.ok(stackedXml.includes('<c:grouping val="stacked"/>'));
    assert.ok(stackedXml.includes("<c:minus>"));
    assert.ok(!stackedXml.includes("<c:plus>"));
  });

  it("reads literal custom values and flags corrupted actual-file error bars", async () => {
    const props = columnChart([
      {
        name: "Observed",
        values: [1, 2, 3],
        errorBars: {
          type: "custom",
          plus: [0.4, null, 0.8],
          minus: [0.1, 0.2, 0.3],
        },
      },
    ]);
    const result = await build(
      <Deck width={480} height={320}>
        <Slide id="literal">
          <Chart {...props} id="literal" />
        </Slide>
      </Deck>,
    );
    const original = unzipSync(result.pptx);
    const chartPath = "ppt/charts/chart1.xml";
    const originalXml = strFromU8(original[chartPath]!);

    const literal = originalXml.replace(
      /<c:plus><c:numRef>.*?<\/c:numRef><\/c:plus>/u,
      '<c:plus><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>0.4</c:v></c:pt><c:pt idx="2"><c:v>0.8</c:v></c:pt></c:numLit></c:plus>',
    );
    assert.notEqual(literal, originalXml);
    const literalParts = unzipSync(result.pptx);
    literalParts[chartPath] = strToU8(literal);
    const literalFile = await validatePptx(zipSync(literalParts));
    assert.deepEqual(literalFile.errors, []);
    assert.deepEqual(
      literalFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.plus,
      {
        source: "literal",
        values: [0.4, null, 0.8],
      },
    );

    const duplicateStart = originalXml.indexOf("<c:errBars>");
    const duplicateEnd =
      originalXml.indexOf("</c:errBars>") + "</c:errBars>".length;
    const duplicate =
      originalXml.slice(0, duplicateEnd) +
      originalXml.slice(duplicateStart, duplicateEnd) +
      originalXml.slice(duplicateEnd);
    const duplicateParts = unzipSync(result.pptx);
    duplicateParts[chartPath] = strToU8(duplicate);
    const duplicateFile = await validatePptx(zipSync(duplicateParts));
    assert.ok(
      duplicateFile.errors.some(
        ({ code, message }) =>
          code === "invalid-chart-error-bars" &&
          message.includes("more than one"),
      ),
    );

    const mismatched = originalXml.replace(
      '<c:pt idx="0"><c:v>0.4</c:v></c:pt>',
      '<c:pt idx="0"><c:v>9.4</c:v></c:pt>',
    );
    assert.notEqual(mismatched, originalXml);
    const mismatchParts = unzipSync(result.pptx);
    mismatchParts[chartPath] = strToU8(mismatched);
    const mismatchFile = await validatePptx(zipSync(mismatchParts));
    assert.ok(
      mismatchFile.errors.some(
        ({ code }) => code === "chart-workbook-mismatch",
      ),
    );

    const savedAsStandardError = originalXml.replace(
      '<c:errValType val="cust"/>',
      '<c:errValType val="stdErr"/>',
    );
    assert.notEqual(savedAsStandardError, originalXml);
    const savedParts = unzipSync(result.pptx);
    savedParts[chartPath] = strToU8(savedAsStandardError);
    const savedFile = await validatePptx(zipSync(savedParts));
    assert.deepEqual(savedFile.errors, []);
    assert.equal(
      savedFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.type,
      "standardError",
    );
    assert.ok(
      savedFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0]?.plus,
    );

    const negative = originalXml.replace(
      '<c:pt idx="0"><c:v>0.4</c:v></c:pt>',
      '<c:pt idx="0"><c:v>-0.4</c:v></c:pt>',
    );
    const negativeParts = unzipSync(result.pptx);
    negativeParts[chartPath] = strToU8(negative);
    const negativeFile = await validatePptx(zipSync(negativeParts));
    assert.ok(
      negativeFile.errors.some(
        ({ code, message }) =>
          code === "invalid-chart-error-bars" &&
          message.includes("non-negative"),
      ),
    );
  });

  it("leaves charts without error bars unchanged", async () => {
    const props = columnChart([{ name: "Plain", values: [1, 2, 3] }]);
    const result = await build(
      <Deck width={480} height={320}>
        <Slide id="plain">
          <Chart {...props} id="plain" />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.chart?.series[0]?.errorBars,
      undefined,
    );
    assert.ok(!chartParts(result.pptx).chart(1).includes("<c:errBars"));
  });

  it("builds line custom error bars for both and single sides through both entries", async () => {
    const props = (include: "both" | "plus" | "minus"): ChartProps => ({
      x: 0,
      y: 0,
      width: 400,
      height: 260,
      groups: [
        {
          kind: "line",
          data: {
            kind: "shared-categories",
            categories: ["Alpha", "Alpha", "Beta"],
          },
          series: [
            {
              name: "Observed",
              values: [1, null, 3],
              errorBars:
                include === "plus"
                  ? {
                      type: "custom",
                      include: "plus",
                      plus: [0.4, null, 0.8],
                      endCap: false,
                      line: { color: "#112233", width: 2, dash: "dash" },
                    }
                  : include === "minus"
                    ? {
                        type: "custom",
                        include: "minus",
                        minus: [0.1, 0.2, 0.3],
                        endCap: false,
                        line: { color: "#112233", width: 2, dash: "dash" },
                      }
                    : {
                        type: "custom",
                        plus: [0.4, null, 0.8],
                        minus: [0.1, 0.2, 0.3],
                        endCap: false,
                        line: { color: "#112233", width: 2, dash: "dash" },
                      },
            },
          ],
        },
      ],
    });
    for (const entry of ["document", "react"] as const) {
      for (const include of ["both", "plus", "minus"] as const) {
        const pptx = await buildInEntry(entry, props(include));
        const file = await validatePptx(pptx);
        assert.deepEqual(
          file.errors,
          [],
          `${entry} ${include} line custom error bars should validate`,
        );
        const chart = file.slides[0]?.objects[0]?.chart;
        const bar = chart?.series[0]?.errorBars?.[0];
        assert.ok(bar, `${entry} ${include} should expose an error bar`);
        assert.equal(bar.direction, "y");
        assert.equal(bar.declaredDirection, "y");
        assert.equal(bar.include, include);
        assert.equal(bar.type, "custom");
        assert.equal(bar.endCap, false);
        assert.deepEqual(bar.line, {
          color: "#112233",
          widthEmu: 19_050,
          dash: "dash",
        });
        if (include === "minus") {
          assert.equal(bar.plus, undefined);
        } else {
          assert.deepEqual(bar.plus, {
            source: "reference",
            formula: "Sheet1!$C$2:$C$4",
            values: [0.4, null, 0.8],
          });
          assert.deepEqual(
            chart?.embeddedWorkbook?.series[0]?.errorBars?.[0]?.plus?.values,
            [0.4, null, 0.8],
          );
        }
        if (include === "plus") {
          assert.equal(bar.minus, undefined);
        } else {
          assert.deepEqual(bar.minus, {
            source: "reference",
            formula:
              include === "both" ? "Sheet1!$D$2:$D$4" : "Sheet1!$C$2:$C$4",
            values: [0.1, 0.2, 0.3],
          });
          assert.deepEqual(
            chart?.embeddedWorkbook?.series[0]?.errorBars?.[0]?.minus?.values,
            [0.1, 0.2, 0.3],
          );
        }
        const parts = chartParts(pptx);
        const xml = parts.chart(1);
        assert.ok(xml.includes('<c:errDir val="y"/>'));
        assert.equal(xml.includes("<c:plus>"), include !== "minus");
        assert.equal(xml.includes("<c:minus>"), include !== "plus");
        const workbook = chart?.embeddedWorkbook?.part;
        assert.ok(workbook);
        const sheet = parts.workbook(workbook);
        assert.equal(sheet.includes("Observed +"), include !== "minus");
        assert.equal(sheet.includes("Observed -"), include !== "plus");
        if (include === "both") {
          assert.ok(sheet.includes('r="C2"'));
          assert.ok(sheet.includes('r="C4"'));
          assert.ok(!sheet.includes('r="C3"'));
          assert.ok(sheet.includes('r="D2"'));
          assert.ok(sheet.includes('r="D3"'));
          assert.ok(sheet.includes('r="D4"'));
        } else if (include === "plus") {
          assert.ok(sheet.includes('r="C2"'));
          assert.ok(sheet.includes('r="C4"'));
          assert.ok(!sheet.includes('r="C3"'));
          assert.ok(!sheet.includes('r="D2"'));
        } else {
          assert.ok(sheet.includes('r="C2"'));
          assert.ok(sheet.includes('r="C3"'));
          assert.ok(sheet.includes('r="C4"'));
          assert.ok(!sheet.includes('r="D2"'));
        }
      }
    }
  });

  it("builds the same declarations through the document and React entries", async () => {
    const props = columnChart([
      {
        name: "Revenue",
        values: [1, 2, 3],
        errorBars: {
          type: "custom",
          plus: [0.4, null, 0.8],
          minus: [0.1, 0.2, 0.3],
          include: "both",
          line: { color: "#336699", width: 1 },
        },
      },
    ]);
    const documentFile = await validatePptx(
      await buildInEntry("document", props),
    );
    const reactFile = await validatePptx(await buildInEntry("react", props));
    assert.deepEqual(documentFile.errors, []);
    assert.deepEqual(reactFile.errors, []);
    assert.deepEqual(
      reactFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars,
      documentFile.slides[0]?.objects[0]?.chart?.series[0]?.errorBars,
    );
  });

  it("combines series trendlines with custom error bars through both entries", async () => {
    const line: ChartProps = {
      x: 0,
      y: 0,
      width: 420,
      height: 260,
      groups: [
        {
          kind: "line",
          data: {
            kind: "shared-categories",
            categories: ["Alpha", "Alpha", "Beta"],
          },
          series: [
            {
              name: "Observed",
              values: [1, null, 3],
              trendlines: [
                {
                  type: "linear",
                  displayEquation: true,
                  displayRSquared: true,
                },
              ],
              errorBars: {
                type: "custom",
                plus: [0.4, null, 0.8],
                minus: [0.1, 0.2, 0.3],
              },
            },
          ],
        },
      ],
    };
    const scatter: ChartProps = {
      x: 0,
      y: 280,
      width: 420,
      height: 260,
      groups: [
        {
          kind: "scatter",
          data: { kind: "xy" },
          line: "straight",
          series: [
            {
              name: "Sample",
              x: [1, 2, null, 4],
              y: [2, null, 4, 8],
              trendlines: [{ type: "linear" }],
              errorBars: {
                x: {
                  type: "custom",
                  plus: [0.2, null, 0.4, 0.5],
                  minus: [0.1, 0.2, 0.3, 0.4],
                },
                y: {
                  type: "custom",
                  include: "plus",
                  plus: [0.5, null, 0.7, 0.9],
                },
              },
            },
          ],
        },
      ],
    };
    for (const entry of ["document", "react"] as const) {
      let pptx: Uint8Array;
      if (entry === "document") {
        pptx = (
          await buildDocument(
            deck({ width: 480, height: 640 }).append(
              documentSlide({ id: "combo-line" }).append(
                documentChart({ ...line, id: "combo-line" }),
              ),
              documentSlide({ id: "combo-scatter" }).append(
                documentChart({ ...scatter, id: "combo-scatter" }),
              ),
            ),
          )
        ).pptx;
      } else {
        pptx = (
          await build(
            <Deck width={480} height={640}>
              <Slide id="combo-line">
                <Chart {...line} id="combo-line" />
              </Slide>
              <Slide id="combo-scatter">
                <Chart {...scatter} id="combo-scatter" />
              </Slide>
            </Deck>,
          )
        ).pptx;
      }
      const file = await validatePptx(pptx);
      assert.deepEqual(file.errors, [], `${entry} schema and reader errors`);
      const lineChart = file.slides[0]?.objects[0]?.chart;
      const lineSeries = lineChart?.series[0];
      assert.deepEqual(lineSeries?.values, [1, null, 3]);
      assert.equal(lineSeries?.trendlines?.[0]?.type, "linear");
      assert.equal(lineSeries?.trendlines?.[0]?.displayEquation, true);
      assert.equal(lineSeries?.trendlines?.[0]?.displayRSquared, true);
      const lineBar = lineSeries?.errorBars?.[0];
      assert.equal(lineBar?.direction, "y");
      assert.equal(lineBar?.type, "custom");
      assert.deepEqual(lineBar?.plus, {
        source: "reference",
        formula: "Sheet1!$C$2:$C$4",
        values: [0.4, null, 0.8],
      });
      assert.deepEqual(lineBar?.minus, {
        source: "reference",
        formula: "Sheet1!$D$2:$D$4",
        values: [0.1, 0.2, 0.3],
      });
      assert.deepEqual(
        lineChart?.embeddedWorkbook?.series[0]?.errorBars?.[0]?.plus?.values,
        [0.4, null, 0.8],
      );
      assert.deepEqual(
        lineChart?.embeddedWorkbook?.series[0]?.errorBars?.[0]?.minus?.values,
        [0.1, 0.2, 0.3],
      );
      const scatterChart = file.slides[1]?.objects[0]?.chart;
      const scatterSeries = scatterChart?.series[0];
      assert.deepEqual(scatterSeries?.x, [1, 2, null, 4]);
      assert.deepEqual(scatterSeries?.y, [2, null, 4, 8]);
      assert.equal(scatterSeries?.trendlines?.[0]?.type, "linear");
      assert.deepEqual(
        scatterSeries?.errorBars?.map(({ direction }) => direction),
        ["x", "y"],
      );
      const [xBar, yBar] = scatterSeries?.errorBars ?? [];
      assert.deepEqual(xBar?.plus, {
        source: "reference",
        formula: "Sheet1!$D$2:$D$5",
        values: [0.2, null, 0.4, 0.5],
      });
      assert.deepEqual(xBar?.minus, {
        source: "reference",
        formula: "Sheet1!$E$2:$E$5",
        values: [0.1, 0.2, 0.3, 0.4],
      });
      assert.deepEqual(yBar?.plus, {
        source: "reference",
        formula: "Sheet1!$F$2:$F$5",
        values: [0.5, null, 0.7, 0.9],
      });
      assert.equal(yBar?.minus, undefined);
      assert.deepEqual(
        scatterChart?.embeddedWorkbook?.series[0]?.errorBars?.[1]?.plus?.values,
        [0.5, null, 0.7, 0.9],
      );
      const parts = chartParts(pptx);
      for (const [index, dataMarker] of [
        [1, "<c:cat>"],
        [2, "<c:xVal>"],
      ] as const) {
        const xml = parts.chart(index);
        const trendline = xml.indexOf("<c:trendline>");
        const errBars = xml.indexOf("<c:errBars>");
        const data = xml.indexOf(dataMarker);
        const labels = xml.indexOf("<c:dLbls>");
        assert.ok(
          trendline >= 0 && trendline < errBars && errBars < data,
          `${entry} chart${index} must serialize trendline before error bars before data`,
        );
        assert.ok(
          labels === -1 || labels < trendline,
          `${entry} chart${index} must serialize labels before trendline`,
        );
      }
      const lineWorkbook = lineChart?.embeddedWorkbook?.part;
      assert.ok(lineWorkbook);
      const lineSheet = parts.workbook(lineWorkbook);
      assert.ok(lineSheet.includes('r="C2"'));
      assert.ok(!lineSheet.includes('r="C3"'));
      assert.ok(lineSheet.includes('r="C4"'));
      assert.ok(lineSheet.includes('r="D2"'));
      assert.ok(lineSheet.includes('r="D3"'));
      assert.ok(lineSheet.includes('r="D4"'));
      const scatterWorkbook = scatterChart?.embeddedWorkbook?.part;
      assert.ok(scatterWorkbook);
      const scatterSheet = parts.workbook(scatterWorkbook);
      assert.ok(scatterSheet.includes('r="D2"'));
      assert.ok(!scatterSheet.includes('r="D3"'));
      assert.ok(scatterSheet.includes('r="D4"'));
      assert.ok(!scatterSheet.includes('r="F3"'));
      assert.ok(scatterSheet.includes('r="F5"'));
    }
  });
});
