import assert from "node:assert/strict";

import type { ChartGroup } from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

type ChartSeriesObservation = {
  readonly name: string;
  readonly categories?: readonly string[];
  readonly values?: readonly (number | null)[];
  readonly x?: readonly (number | null)[];
  readonly y?: readonly (number | null)[];
};

const columnZero: readonly ChartGroup[] = [
  {
    kind: "column",
    data: { kind: "shared-categories", categories: [] },
    series: [{ name: "Empty", values: [] }],
  },
];

const independentRings: readonly ChartGroup[] = [
  {
    kind: "doughnut",
    data: { kind: "independent-categories" },
    series: [
      { name: "Empty", categories: [], values: [] },
      { name: "Filled", categories: ["A", "B"], values: [2, 3] },
    ],
  },
];

const mixedScatter: readonly ChartGroup[] = [
  {
    kind: "scatter",
    data: { kind: "xy" },
    series: [
      { name: "Empty", x: [], y: [] },
      { name: "Filled", x: [1], y: [2] },
    ],
  },
];

async function buildChart(groups: readonly ChartGroup[]): Promise<{
  readonly pptx: Uint8Array;
  readonly chart: NonNullable<ReturnType<typeof chartFromResult>>;
}> {
  const result = await buildDocument(
    deck({ width: 800, height: 600 }).append(
      documentSlide({ id: "zero-category-column" }).append(
        documentChart({
          id: "zero-category-column",
          x: 10,
          y: 10,
          width: 760,
          height: 560,
          groups,
        }),
      ),
    ),
  );
  const file = await validatePptx(result.pptx);
  assert.deepEqual(file.errors, []);
  const chart = chartFromResult(file);
  assert.ok(chart);
  return { pptx: result.pptx, chart };
}

function chartFromResult(
  file: Awaited<ReturnType<typeof validatePptx>>,
):
  | NonNullable<
      NonNullable<
        Awaited<
          ReturnType<typeof validatePptx>
        >["slides"][number]["objects"][number]
      >["chart"]
    >
  | undefined {
  return file.slides[0]?.objects[0]?.chart;
}

async function withWorkbookCell(
  pptx: Uint8Array,
  reference: string,
  cellXml: string,
): Promise<Uint8Array> {
  const file = await validatePptx(pptx);
  const workbookPart =
    file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.part;
  assert.ok(workbookPart);
  const parts = unzipSync(pptx);
  const workbook = unzipSync(parts[workbookPart]!);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheet = strFromU8(workbook[sheetPath]!);
  const row = Number(/(\d+)$/u.exec(reference)?.[1]);
  assert.ok(Number.isSafeInteger(row) && row >= 2);
  const rowTag = `<row r="${row}">`;
  const nextSheet = sheet.includes(rowTag)
    ? sheet.replace(rowTag, `${rowTag}${cellXml}`)
    : sheet.replace(
        "</sheetData>",
        `<row r="${row}">${cellXml}</row></sheetData>`,
      );
  return zipSync({
    ...parts,
    [workbookPart]: zipSync({
      ...workbook,
      [sheetPath]: strToU8(nextSheet),
    }),
  });
}

function categoryOnlyLabel(): string {
  return '<c r="A2" t="inlineStr"><is><t>New category</t></is></c>';
}

function valueOnlyCell(): string {
  return '<c r="B2"><v>9</v></c>';
}

function assertMismatch(file: Awaited<ReturnType<typeof validatePptx>>): void {
  assert.ok(
    file.errors.some(
      ({ code, part }) =>
        code === "chart-workbook-mismatch" && part === "ppt/charts/chart1.xml",
    ),
    JSON.stringify(file.errors),
  );
}

function assertSeries(
  series: readonly ChartSeriesObservation[],
  expected: readonly ChartSeriesObservation[],
): void {
  assert.deepEqual(
    series.map(({ name, categories, values, x, y }) =>
      categories === undefined ? { name, x, y } : { name, categories, values },
    ),
    expected,
  );
}

describe("generated zero-source category column ownership", () => {
  it("keeps the unmodified category source valid", async () => {
    const result = await buildChart(columnZero);
    assert.deepEqual(result.chart.series[0]?.categories, []);
  });

  it("detects a new label in the owned category column", async () => {
    const source = await buildChart(columnZero);
    const result = await validatePptx(
      await withWorkbookCell(source.pptx, "A2", categoryOnlyLabel()),
    );
    assertMismatch(result);
    assert.deepEqual(
      result.slides[0]?.objects[0]?.chart?.series[0]?.categories,
      [],
    );
    assert.deepEqual(
      result.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.categories,
      [],
    );
  });

  it("keeps the existing numeric value-column mismatch", async () => {
    const source = await buildChart(columnZero);
    const result = await validatePptx(
      await withWorkbookCell(source.pptx, "B2", valueOnlyCell()),
    );
    assertMismatch(result);
  });

  it("keeps independent-ring sources valid when the zero ring has no label", async () => {
    const result = await buildChart(independentRings);
    assertSeries(result.chart.series, [
      { name: "Empty", categories: [], values: [] },
      { name: "Filled", categories: ["A", "B"], values: [2, 3] },
    ]);
  });

  it("detects the zero ring's owned category column without touching the filled ring", async () => {
    const source = await buildChart(independentRings);
    const result = await validatePptx(
      await withWorkbookCell(source.pptx, "A2", categoryOnlyLabel()),
    );
    assertMismatch(result);
    assertSeries(
      result.slides[0]?.objects[0]?.chart
        ?.series as readonly ChartSeriesObservation[],
      [
        { name: "Empty", categories: [], values: [] },
        { name: "Filled", categories: ["A", "B"], values: [2, 3] },
      ],
    );
  });

  it("keeps the independent-ring numeric value-column mismatch", async () => {
    const source = await buildChart(independentRings);
    const result = await validatePptx(
      await withWorkbookCell(source.pptx, "B2", valueOnlyCell()),
    );
    assertMismatch(result);
    assert.deepEqual(
      result.slides[0]?.objects[0]?.chart?.series[1]?.values,
      [2, 3],
    );
  });

  it("does not treat the scatter A column as owned by the zero series", async () => {
    const source = await buildChart(mixedScatter);
    const result = await validatePptx(
      await withWorkbookCell(source.pptx, "A2", categoryOnlyLabel()),
    );
    assert.deepEqual(result.errors, []);
    assertSeries(
      result.slides[0]?.objects[0]?.chart
        ?.series as readonly ChartSeriesObservation[],
      [
        { name: "Empty", x: [], y: [] },
        { name: "Filled", x: [1], y: [2] },
      ],
    );
  });
});
