import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CommonContainer,
  CommonNode,
  RichTextContainer,
} from "@react-pptx/core/document";
import {
  BuildError,
  build,
  cell,
  chart,
  createDocumentBuild,
  deck,
  font,
  paragraph,
  row,
  slide,
  span,
  table,
  type ChartData,
  type TableColumn,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

const notoSansScRegular = new Uint8Array(
  readFileSync(
    new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
  ),
);

const metricFont = font(
  "Table Metric",
  [{ weight: 400, data: notoSansScRegular }],
  { substitute: true },
);

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

describe("table and chart authoring through the common document", () => {
  it("creates an editable native table, preserves rich text, spans, styles, and measured tracks, then edits it", async () => {
    const columns: TableColumn[] = ["auto", "1fr"];
    const mergedStyle = { backgroundColor: "#DDEEFF" };
    const merged = cell(
      paragraph().append(
        span("North").style({ fontWeight: "bold", color: "#AA0000" }),
        " Q1",
      ),
      { colSpan: 2, style: mergedStyle },
    );
    const revenue = cell("Revenue");
    const value = cell("42");
    const firstRow = row({ style: { height: 30 } }).append(merged);
    const secondRow = row().append(revenue, value);
    assert.ok(merged instanceof RichTextContainer);
    assert.ok(firstRow instanceof CommonContainer);
    const tableNode = table({
      testId: "document-table",
      x: 16,
      y: 20,
      width: 360,
      columns,
      style: { fontFamily: metricFont, fontSize: 12 },
    }).append(firstRow, secondRow);
    assert.ok(tableNode instanceof CommonContainer);
    const document = deck({ width: 960, height: 540 }).append(
      slide({ slideId: "table" }).append(tableNode),
    );

    // Values are received at construction time; later edits to the caller's inputs do not leak in.
    columns.push("1fr");
    mergedStyle.backgroundColor = "#0000FF";

    const first = await build(document);
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(firstFile.errors, []);
    const tableObject = firstFile.slides[0]?.objects[0];
    assert.equal(tableObject?.kind, "table");
    const tableObservation = tableObject?.table;
    assert.ok(tableObservation !== undefined);
    assert.deepEqual(tableObservation.cells, [
      ["North Q1", ""],
      ["Revenue", "42"],
    ]);
    assert.deepEqual(tableObservation.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
    ]);
    assert.equal(tableObservation.columnWidthsEmu.length, 2);
    const tableWidth = tableObservation.columnWidthsEmu.reduce(
      (sum, width) => sum + width,
      0,
    );
    assert.ok(Math.abs(tableWidth - 360 * 9525) <= 1);
    assert.equal(tableObservation.rowHeightsEmu[0], 30 * 9525);
    assert.ok((tableObservation.rowHeightsEmu[1] ?? 0) > 0);

    const mergedText = tableObservation.cellTexts.find(
      ({ row: cellRow, column }) => cellRow === 0 && column === 0,
    )?.text;
    assert.equal(mergedText?.text, "North Q1");
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.text, "North");
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.bold, true);
    assert.equal(mergedText?.paragraphs[0]?.runs[0]?.color, "#AA0000");
    assert.match(slideXml(first), /DDEEFF/u);
    assert.doesNotMatch(slideXml(first), /0000FF/u);

    merged.content("Updated ", span("cell").style({ fontStyle: "italic" }));
    value.style({ backgroundColor: "#FFFF00" });
    secondRow.style({ height: 40 });

    const second = await build(document);
    const secondFile = await validatePptx(second.pptx);
    assert.deepEqual(secondFile.errors, []);
    const updatedTable = secondFile.slides[0]?.objects[0]?.table;
    assert.deepEqual(updatedTable?.cells, [
      ["Updated cell", ""],
      ["Revenue", "42"],
    ]);
    assert.equal(updatedTable?.rowHeightsEmu[1], 40 * 9525);
    assert.match(slideXml(second), /Updated /u);
    assert.match(slideXml(second), />cell</u);
    assert.match(slideXml(second), /FFFF00/u);
    assert.match(slideXml(second), /i="1"/u);
    assert.doesNotMatch(slideXml(first), /Updated /u);
    assert.deepEqual(firstFile, await validatePptx(first.pptx));
  });

  it("writes the six established chart kinds with matching native caches and embedded workbooks", async () => {
    const cartesian: ChartData = {
      categories: ["Q1", "Q2", "Q3"],
      series: [
        { name: "Alpha", values: [10, 20, 30] },
        { name: "Beta", values: [5, null, 15] },
      ],
    };
    const share: ChartData = {
      categories: ["Direct", "Partner"],
      series: [{ name: "Share", values: [3, 2] }],
    };
    const document = deck({ width: 960, height: 540 }).append(
      slide({ slideId: "charts" }).append(
        chart({
          testId: "column",
          kind: "column",
          x: 0,
          y: 0,
          width: 300,
          height: 200,
          data: cartesian,
        }),
        chart({
          testId: "bar",
          kind: "bar",
          x: 320,
          y: 0,
          width: 300,
          height: 200,
          data: cartesian,
          stacking: "stacked",
        }),
        chart({
          testId: "line",
          kind: "line",
          x: 0,
          y: 220,
          width: 300,
          height: 200,
          data: cartesian,
        }),
        chart({
          testId: "area",
          kind: "area",
          x: 320,
          y: 220,
          width: 300,
          height: 200,
          data: cartesian,
        }),
        chart({
          testId: "pie",
          kind: "pie",
          x: 640,
          y: 0,
          width: 300,
          height: 200,
          data: share,
          dataLabels: { value: true, category: true },
        }),
        chart({
          testId: "doughnut",
          kind: "doughnut",
          x: 640,
          y: 220,
          width: 300,
          height: 200,
          data: share,
          holeSize: 55,
        }),
      ),
    );

    const result = await build(document);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);

    const expected = [
      ["barChart", "col", cartesian],
      ["barChart", "bar", cartesian],
      ["lineChart", undefined, cartesian],
      ["areaChart", undefined, cartesian],
      ["pieChart", undefined, share],
      ["doughnutChart", undefined, share],
    ] as const;
    assert.equal(file.slides[0]?.objects.length, expected.length);
    expected.forEach(([type, direction, data], index) => {
      const value = chartObject(file, index);
      const series = data.series.map((item) => ({
        ...item,
        categories: data.categories,
      }));
      assert.deepEqual(value.types, [type]);
      assert.equal(value.direction, direction);
      assert.deepEqual(value.series, series);
      assert.deepEqual(value.embeddedWorkbook?.categories, data.categories);
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

  it("resolves shared-string and rich-text chart workbook cells through the workbook relationship", async () => {
    const data: ChartData = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [121, 135] },
        { name: "South", values: [90, 105] },
      ],
    };
    const result = await build(
      deck({ width: 960, height: 540 }).append(
        slide({ slideId: "shared-strings" }).append(
          chart({
            testId: "shared-strings-chart",
            kind: "column",
            x: 0,
            y: 0,
            width: 300,
            height: 200,
            data,
          }),
        ),
      ),
    );

    const file = await validatePptx(sharedStringWorkbook(result.pptx));
    assert.deepEqual(file.errors, []);
    const chartValue = chartObject(file, 0);
    const expectedSeries = data.series.map((item) => ({
      ...item,
      categories: data.categories,
    }));
    assert.deepEqual(chartValue.series, expectedSeries);
    assert.deepEqual(chartValue.embeddedWorkbook?.categories, data.categories);
    assert.deepEqual(chartValue.embeddedWorkbook?.series, expectedSeries);
  });

  it("keeps chart workbook mismatches after shared-string resolution", async () => {
    const data: ChartData = {
      categories: ["Q1", "Q2"],
      series: [
        { name: "North", values: [121, 135] },
        { name: "South", values: [90, 105] },
      ],
    };
    const result = await build(
      deck({ width: 960, height: 540 }).append(
        slide({ slideId: "shared-strings-mismatch" }).append(
          chart({
            testId: "shared-strings-mismatch-chart",
            kind: "column",
            x: 0,
            y: 0,
            width: 300,
            height: 200,
            data,
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
    assert.deepEqual(chartValue.embeddedWorkbook?.series, [
      { name: "North", categories: ["Q1", "Q2"], values: [121, 999] },
      { name: "South", categories: ["Q1", "Q2"], values: [90, 105] },
    ]);
  });

  it("isolates nested chart data and applies explicit updates only to later builds", async () => {
    const data = {
      categories: ["Q1", "Q2"],
      series: [{ name: "Original", values: [1, 2] }],
    };
    const node = chart({
      testId: "editable-chart",
      kind: "column",
      x: 0,
      y: 0,
      width: 300,
      height: 200,
      data,
    });
    assert.ok(node instanceof CommonNode);
    assert.ok(!(node instanceof CommonContainer));
    const document = deck({ width: 960, height: 540 }).append(
      slide({ slideId: "chart-edit" }).append(node),
    );

    data.categories[0] = "Changed";
    data.series[0]!.name = "Mutated";
    data.series[0]!.values[0] = 99;

    const first = await build(document);
    const firstChart = chartObject(await validatePptx(first.pptx), 0);
    assert.deepEqual(firstChart.series, [
      { name: "Original", categories: ["Q1", "Q2"], values: [1, 2] },
    ]);
    assert.deepEqual(firstChart.embeddedWorkbook?.series, firstChart.series);

    node.set({
      data: {
        categories: ["R1"],
        series: [{ name: "Revised", values: [7] }],
      },
    });
    const second = await build(document);
    const secondChart = chartObject(await validatePptx(second.pptx), 0);
    assert.deepEqual(secondChart.series, [
      { name: "Revised", categories: ["R1"], values: [7] },
    ]);
    assert.deepEqual(secondChart.embeddedWorkbook?.series, secondChart.series);
    assert.deepEqual(firstChart.series, [
      { name: "Original", categories: ["Q1", "Q2"], values: [1, 2] },
    ]);
    assert.doesNotMatch(slideXml(first), /Revised/u);
    assert.doesNotMatch(slideXml(first), /Changed/u);
  });

  it("keeps table and chart registration explicit for trimmed builds", async () => {
    const document = deck({ width: 960, height: 540 }).append(
      slide({ slideId: "trimmed" }).append(
        table({ columns: [100] }).append(
          row({ style: { height: 20 } }).append(cell("Trimmed table")),
        ),
        chart({
          kind: "column",
          width: 120,
          height: 80,
          data: {
            categories: ["A"],
            series: [{ name: "S", values: [1] }],
          },
        }),
      ),
    );

    await assert.rejects(
      createDocumentBuild([])(document),
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
