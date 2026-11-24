import assert from "node:assert/strict";

import { type ChartGroup, type ChartProps } from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@pptxdk/capabilities/document";
import { BuildError, CommonNode } from "@pptxdk/core";
import { Deck, expandDocument, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

function chartXml(parts: Unzipped, index = 1): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

function sheetXml(parts: Unzipped, workbookPart: string): string {
  return strFromU8(
    unzipSync(parts[workbookPart]!)["xl/worksheets/sheet1.xml"]!,
  );
}

function categoryChart(groups: ChartProps["groups"]): ChartProps {
  return {
    id: "empty-category",
    x: 24,
    y: 24,
    width: 560,
    height: 360,
    title: { text: "Empty category chart" },
    legend: "bottom",
    style: { fontFamily: "Arial", fontSize: "12pt", color: "#112233" },
    groups,
  };
}

describe("editable empty charts", () => {
  it("keeps a declared category chart native when categories and values are empty", async () => {
    const result = await build(
      <Deck width={640} height={480}>
        <Slide id="empty-category">
          <Chart
            {...categoryChart([
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [
                  {
                    name: "North",
                    values: [],
                    color: "#336699",
                  },
                ],
              },
            ])}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects[0];
    assert.equal(object?.kind, "chart");
    assert.equal(object?.chart?.series.length, 1);
    assert.deepEqual(object?.chart?.series[0]?.categories, []);
    assert.deepEqual(object?.chart?.series[0]?.values, []);
    assert.equal(object?.chart?.series[0]?.name, "North");
    assert.deepEqual(object?.chart?.embeddedWorkbook?.series[0]?.values, []);
    assert.deepEqual(object?.boxEmu, {
      x: 24 * 9525,
      y: 24 * 9525,
      width: 560 * 9525,
      height: 360 * 9525,
    });

    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts);
    assert.ok(xml.includes("<c:ser>"));
    assert.ok(xml.includes("<c:tx>"));
    assert.ok(!xml.includes("<c:cat>"));
    assert.ok(!xml.includes("<c:val>"));
    assert.ok(!xml.includes("<c:numCache>"));
    assert.ok(!xml.includes("$A$2:$A$1"));
    const workbookPart = object?.chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(parts, workbookPart);
    assert.ok(sheet.includes("North"));
    assert.ok(!sheet.includes('r="B2"'));
  });

  it("keeps independent XY observations empty without substituting category data", async () => {
    const result = await build(
      <Deck width={640} height={480}>
        <Slide id="empty-scatter">
          <Chart
            id="empty-scatter"
            x={24}
            y={24}
            width={560}
            height={360}
            title="Empty scatter"
            groups={[
              {
                kind: "scatter",
                data: { kind: "xy" },
                line: "straight",
                markers: true,
                series: [
                  {
                    name: "Observations",
                    x: [],
                    y: [],
                    color: "#AA5500",
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
    const object = file.slides[0]?.objects[0];
    assert.equal(object?.kind, "chart");
    assert.deepEqual(object?.chart?.series[0]?.x, []);
    assert.deepEqual(object?.chart?.series[0]?.y, []);
    assert.equal(object?.chart?.series[0]?.categories, undefined);
    assert.deepEqual(object?.chart?.embeddedWorkbook?.series[0]?.x, []);
    assert.deepEqual(object?.chart?.embeddedWorkbook?.series[0]?.y, []);

    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts);
    assert.ok(xml.includes("<c:scatterChart>"));
    assert.ok(!xml.includes("<c:xVal>"));
    assert.ok(!xml.includes("<c:yVal>"));
    assert.ok(!xml.includes("<c:cat>"));
    assert.ok(!xml.includes("<c:val>"));
    assert.ok(!xml.includes("<c:numCache>"));
    const workbookPart = object?.chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(parts, workbookPart);
    assert.ok(sheet.includes("Observations"));
    assert.ok(!sheet.includes('r="B2"'));
    assert.ok(!sheet.includes('r="C2"'));
  });

  it("keeps every existing category kind native with no observations", async () => {
    const kinds = ["column", "bar", "line", "area", "pie", "doughnut"] as const;
    for (const kind of kinds) {
      const result = await build(
        <Deck width={320} height={240}>
          <Slide id={`empty-${kind}`}>
            <Chart
              id={`empty-${kind}`}
              x={16}
              y={16}
              width={288}
              height={180}
              groups={[
                {
                  kind,
                  data: { kind: "shared-categories", categories: [] },
                  series: [{ name: `${kind} series`, values: [] }],
                } as ChartGroup,
              ]}
            />
          </Slide>
        </Deck>,
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, [], kind);
      const chart = file.slides[0]?.objects[0]?.chart;
      assert.equal(file.slides[0]?.objects[0]?.kind, "chart", kind);
      assert.deepEqual(chart?.series[0]?.values, [], kind);
      assert.deepEqual(chart?.series[0]?.categories, [], kind);
      const xml = chartXml(unzipSync(result.pptx));
      if (kind === "pie" || kind === "doughnut") {
        assert.ok(!xml.includes("<c:catAx"), kind);
        assert.ok(!xml.includes("<c:valAx"), kind);
      } else {
        assert.ok(xml.includes("<c:catAx"), kind);
        assert.ok(xml.includes("<c:valAx"), kind);
      }
      assert.ok(!xml.includes("<c:cat>"), kind);
      assert.ok(!xml.includes("<c:val>"), kind);
      assert.ok(!xml.includes("<c:numCache>"), kind);
    }
  });

  it("keeps same-kind groups with empty category domains separate", async () => {
    const result = await build(
      <Deck width={640} height={480}>
        <Slide id="empty-multi-group">
          <Chart
            id="empty-multi-group"
            x={24}
            y={24}
            width={560}
            height={360}
            title="Two empty groups"
            legend="bottom"
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [
                  { name: "North", values: [], color: "#336699" },
                  { name: "South", values: [], color: "#AA5500" },
                ],
                gapWidth: 70,
              },
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [{ name: "East", values: [], color: "#117744" }],
                gapWidth: 20,
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.equal(chart?.groups.length, 2);
    assert.deepEqual(
      chart?.series.map(({ name, groupIndex, values }) => ({
        name,
        groupIndex,
        values,
      })),
      [
        { name: "North", groupIndex: 0, values: [] },
        { name: "South", groupIndex: 0, values: [] },
        { name: "East", groupIndex: 1, values: [] },
      ],
    );
    assert.deepEqual(
      chart?.embeddedWorkbook?.series.map(({ name }) => name),
      ["North", "South", "East"],
    );
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts);
    assert.equal(xml.match(/<c:ser>/gu)?.length, 3);
    assert.ok(xml.includes('<a:srgbClr val="336699"/>'));
    assert.ok(xml.includes('<a:srgbClr val="AA5500"/>'));
    assert.ok(xml.includes('<a:srgbClr val="117744"/>'));
    assert.ok(!xml.includes("<c:cat>"));
    assert.ok(!xml.includes("<c:val>"));
    assert.ok(!xml.includes("<c:numCache>"));
    const workbookPart = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(parts, workbookPart);
    for (const name of ["North", "South", "East"]) {
      assert.ok(sheet.includes(name), name);
    }
    assert.ok(!sheet.includes('r="A2"'));
    assert.ok(!sheet.includes('r="B2"'));
    assert.ok(!sheet.includes('r="C2"'));
  });

  it("keeps independent and shared multi-ring zero observations", async () => {
    const result = await build(
      <Deck width={960} height={540}>
        <Slide id="multi-ring-empty">
          <Chart
            id="independent-rings"
            x={20}
            y={20}
            width={440}
            height={260}
            title="Independent rings"
            legend="right"
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                series: [
                  {
                    name: "Empty ring",
                    categories: [],
                    values: [],
                    color: "#336699",
                  },
                  {
                    name: "Filled ring",
                    categories: ["A", "B"],
                    values: [1, 2],
                    color: "#AA5500",
                  },
                ],
                dataLabels: { value: true, category: true },
              },
            ]}
          />
          <Chart
            id="shared-rings"
            x={500}
            y={20}
            width={440}
            height={260}
            title="Shared rings"
            legend="right"
            groups={[
              {
                kind: "doughnut",
                data: { kind: "shared-categories", categories: [] },
                series: [
                  { name: "Ring one", values: [], color: "#117744" },
                  { name: "Ring two", values: [], color: "#80558C" },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const independent = file.slides[0]?.objects[0]?.chart;
    const shared = file.slides[0]?.objects[1]?.chart;
    assert.deepEqual(
      independent?.series.map(({ name, categories, values }) => ({
        name,
        categories,
        values,
      })),
      [
        { name: "Empty ring", categories: [], values: [] },
        { name: "Filled ring", categories: ["A", "B"], values: [1, 2] },
      ],
    );
    assert.deepEqual(
      shared?.series.map(({ name, categories, values }) => ({
        name,
        categories,
        values,
      })),
      [
        { name: "Ring one", categories: [], values: [] },
        { name: "Ring two", categories: [], values: [] },
      ],
    );
    const parts = unzipSync(result.pptx);
    const independentXml = chartXml(parts, 1);
    const sharedXml = chartXml(parts, 2);
    assert.equal((independentXml.match(/<c:ser>/gu) ?? []).length, 2);
    assert.equal((independentXml.match(/<c:cat>/gu) ?? []).length, 1);
    assert.equal((independentXml.match(/<c:val>/gu) ?? []).length, 1);
    assert.ok(!independentXml.includes("$A$2:$A$1"));
    assert.equal((sharedXml.match(/<c:cat>/gu) ?? []).length, 0);
    assert.equal((sharedXml.match(/<c:val>/gu) ?? []).length, 0);
    assert.ok(!sharedXml.includes("<c:numCache>"));
    const independentWorkbook = independent?.embeddedWorkbook?.part;
    assert.ok(independentWorkbook);
    const independentSheet = sheetXml(parts, independentWorkbook);
    assert.ok(independentSheet.includes("Empty ring"));
    assert.ok(independentSheet.includes("Filled ring"));
    assert.ok(!independentSheet.includes('r="B2"'));
    assert.ok(independentSheet.includes('r="D2"'));
    const sharedWorkbook = shared?.embeddedWorkbook?.part;
    assert.ok(sharedWorkbook);
    const sharedSheet = sheetXml(parts, sharedWorkbook);
    assert.ok(sharedSheet.includes("Ring one"));
    assert.ok(sharedSheet.includes("Ring two"));
    assert.ok(!sharedSheet.includes('r="A2"'));
    assert.ok(!sharedSheet.includes('r="B2"'));
  });

  it("keeps a mixed zero group and non-empty group on independent axes", async () => {
    const axes: NonNullable<ChartProps["axes"]> = {
      cat1: { kind: "category" },
      val1: { kind: "value", min: 0, max: 4 },
      cat2: { kind: "category" },
      val2: { kind: "value", min: 0, max: 10 },
    };
    const initialGroups: readonly ChartGroup[] = [
      {
        kind: "column",
        axes: { category: "cat1", value: "val1" },
        data: { kind: "shared-categories", categories: [] },
        series: [{ name: "Empty column", values: [], color: "#336699" }],
      },
      {
        kind: "line",
        axes: { category: "cat2", value: "val2" },
        data: { kind: "shared-categories", categories: ["A", "B"] },
        series: [{ name: "Filled line", values: [3, 4], color: "#AA5500" }],
      },
    ];
    const node = documentChart({
      id: "combo-empty-fill",
      x: 20,
      y: 20,
      width: 600,
      height: 360,
      axes,
      title: "Mixed empty and filled",
      groups: initialGroups,
    });
    const document = deck({ width: 640, height: 420 }).append(
      documentSlide({ id: "combo-empty-fill" }).append(node),
    );
    const emptyResult = await buildDocument(document);
    const emptyFile = await validatePptx(emptyResult.pptx);
    assert.deepEqual(emptyFile.errors, []);
    const emptyChart = emptyFile.slides[0]?.objects[0]?.chart;
    assert.deepEqual(emptyChart?.series[0]?.values, []);
    assert.deepEqual(emptyChart?.series[1]?.values, [3, 4]);
    const emptyParts = unzipSync(emptyResult.pptx);
    const emptyXml = chartXml(emptyParts);
    assert.ok(!emptyXml.includes("$A$2:$A$1"));
    assert.equal((emptyXml.match(/<c:cat>/gu) ?? []).length, 1);
    assert.equal((emptyXml.match(/<c:val>/gu) ?? []).length, 1);
    const emptyPptx = emptyResult.pptx.slice();
    const workbookPart = emptyChart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const workbook = unzipSync(emptyParts[workbookPart]!);
    const sheetPath = "xl/worksheets/sheet1.xml";
    const ownedCategoryLabel = strFromU8(workbook[sheetPath]!).replace(
      "</sheetData>",
      '<row r="4"><c r="A4" t="inlineStr"><is><t>Owner only</t></is></c></row></sheetData>',
    );
    const mutated = await validatePptx(
      zipSync({
        ...emptyParts,
        [workbookPart]: zipSync({
          ...workbook,
          [sheetPath]: strToU8(ownedCategoryLabel),
        }),
      }),
    );
    assert.ok(
      mutated.errors.some(
        ({ code, part }) =>
          code === "chart-workbook-mismatch" &&
          part === "ppt/charts/chart1.xml",
      ),
      JSON.stringify(mutated.errors),
    );
    assert.deepEqual(emptyResult.pptx, emptyPptx);

    node.set({
      groups: [
        {
          kind: "column",
          axes: { category: "cat1", value: "val1" },
          data: { kind: "shared-categories", categories: ["X", "Y"] },
          series: [{ name: "Empty column", values: [1, 2], color: "#336699" }],
        },
        {
          kind: "line",
          axes: { category: "cat2", value: "val2" },
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "Filled line", values: [3, 4], color: "#AA5500" }],
        },
      ],
    });
    const filledResult = await buildDocument(document);
    const filledFile = await validatePptx(filledResult.pptx);
    assert.deepEqual(filledFile.errors, []);
    const filledChart = filledFile.slides[0]?.objects[0]?.chart;
    assert.deepEqual(filledChart?.series[0]?.values, [1, 2]);
    assert.deepEqual(filledChart?.series[1]?.values, [3, 4]);
    assert.deepEqual(emptyResult.pptx, emptyPptx);
  });

  it("keeps an empty XY series alongside unequal non-empty observations", async () => {
    const result = await build(
      <Deck width={640} height={480}>
        <Slide id="scatter-empty-mixed">
          <Chart
            id="scatter-empty-mixed"
            x={20}
            y={20}
            width={600}
            height={360}
            title="Scatter empty and filled"
            groups={[
              {
                kind: "scatter",
                data: { kind: "xy" },
                line: "straight",
                markers: true,
                series: [
                  {
                    name: "Empty observations",
                    x: [],
                    y: [],
                    color: "#336699",
                  },
                  {
                    name: "Filled observations",
                    x: [1, 2, 3],
                    y: [4, 5, 6],
                    color: "#AA5500",
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
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.x, []);
    assert.deepEqual(chart?.series[0]?.y, []);
    assert.deepEqual(chart?.series[1]?.x, [1, 2, 3]);
    assert.deepEqual(chart?.series[1]?.y, [4, 5, 6]);
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts);
    assert.equal((xml.match(/<c:xVal>/gu) ?? []).length, 1);
    assert.equal((xml.match(/<c:yVal>/gu) ?? []).length, 1);
    assert.ok(!xml.includes("$B$2:$B$1"));
    assert.ok(!xml.includes("$C$2:$C$1"));
    const workbookPart = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(parts, workbookPart);
    assert.ok(sheet.includes("Empty observations"));
    assert.ok(sheet.includes("Filled observations"));
    assert.ok(!sheet.includes('r="B2"'));
    assert.ok(!sheet.includes('r="C2"'));
    assert.ok(sheet.includes('r="D2"'));
    assert.ok(sheet.includes('r="E2"'));
    assert.ok(sheet.includes('r="D4"'));
    assert.ok(sheet.includes('r="E4"'));
  });

  it("keeps empty, all-missing, and all-zero values as distinct declarations", async () => {
    const chartProps = (
      id: string,
      categories: readonly string[],
      values: readonly (number | null)[],
    ): ChartProps => ({
      id,
      x: 0,
      y: 0,
      width: 200,
      height: 120,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories },
          series: [{ name: id, values }],
        },
      ],
    });
    const result = await build(
      <Deck width={640} height={160}>
        <Slide id="empty-null-zero">
          <Chart {...chartProps("empty", [], [])} />
          <Chart {...chartProps("missing", ["A", "B"], [null, null])} />
          <Chart {...chartProps("zero", ["A", "B"], [0, 0])} />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const charts = file.slides[0]?.objects.map((object) => object.chart);
    assert.deepEqual(
      charts?.map((chart) => chart?.series[0]?.values),
      [[], [null, null], [0, 0]],
    );
    assert.deepEqual(
      charts?.map((chart) => chart?.embeddedWorkbook?.series[0]?.values),
      [[], [null, null], [0, 0]],
    );

    const parts = unzipSync(result.pptx);
    const empty = chartXml(parts, 1);
    const missing = chartXml(parts, 2);
    const zero = chartXml(parts, 3);
    assert.ok(!empty.includes("<c:cat>"));
    assert.ok(!empty.includes("<c:val>"));
    assert.ok(!empty.includes("<c:numCache>"));
    assert.ok(
      missing.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/></c:numCache>',
      ),
    );
    assert.ok(
      zero.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numCache>',
      ),
    );
    const sheet = (index: number) => {
      const workbookPart = charts?.[index - 1]?.embeddedWorkbook?.part;
      assert.ok(workbookPart);
      return sheetXml(parts, workbookPart);
    };
    assert.ok(!sheet(1).includes('r="B2"'));
    assert.ok(!sheet(2).includes('r="B2"'));
    assert.ok(!sheet(2).includes('r="B3"'));
    assert.ok(sheet(3).includes('<c r="B2"><v>0</v></c>'));
    assert.ok(sheet(3).includes('<c r="B3"><v>0</v></c>'));
  });
  it("fills a command-authoring chart node and rebuilds consistent native data", async () => {
    const node = documentChart({
      id: "empty-command",
      x: 24,
      y: 24,
      width: 560,
      height: 360,
      title: "Command fill",
      legend: "bottom",
      style: { fontFamily: "Arial", fontSize: "12pt", color: "#112233" },
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: [] },
          series: [{ name: "North", values: [], color: "#336699" }],
        },
      ],
    });
    const document = deck({ width: 640, height: 480 }).append(
      documentSlide({ id: "fill-command" }).append(node),
    );

    const emptyResult = await buildDocument(document);
    const emptyFile = await validatePptx(emptyResult.pptx);
    assert.deepEqual(emptyFile.errors, []);
    assert.deepEqual(
      emptyFile.slides[0]?.objects[0]?.chart?.series[0]?.values,
      [],
    );
    const emptyPptx = emptyResult.pptx.slice();
    const emptyBox = emptyFile.slides[0]?.objects[0]?.boxEmu;

    node.set({
      groups: [
        {
          kind: "column",
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q2"],
          },
          series: [{ name: "North", values: [14, 18], color: "#336699" }],
        },
      ],
    });
    const filledResult = await buildDocument(document);
    const filledFile = await validatePptx(filledResult.pptx);
    assert.deepEqual(filledFile.errors, []);
    const filledChart = filledFile.slides[0]?.objects[0]?.chart;
    assert.deepEqual(filledChart?.series[0]?.categories, ["Q1", "Q2"]);
    assert.deepEqual(filledChart?.series[0]?.values, [14, 18]);
    assert.deepEqual(
      filledChart?.embeddedWorkbook?.series[0]?.values,
      [14, 18],
    );
    assert.deepEqual(filledFile.slides[0]?.objects[0]?.boxEmu, emptyBox);

    const filledParts = unzipSync(filledResult.pptx);
    const filledXml = chartXml(filledParts);
    assert.ok(
      filledXml.includes(
        '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>14</c:v></c:pt><c:pt idx="1"><c:v>18</c:v></c:pt></c:numCache>',
      ),
    );
    assert.ok(filledXml.includes('typeface="Arial"'));
    const workbookPart = filledChart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(filledParts, workbookPart);
    assert.ok(sheet.includes('<c r="B2"><v>14</v></c>'));
    assert.ok(sheet.includes('<c r="B3"><v>18</v></c>'));
    assert.deepEqual(emptyResult.pptx, emptyPptx);
    const emptyXml = chartXml(unzipSync(emptyResult.pptx));
    assert.ok(!emptyXml.includes("<c:cat>"));
    assert.ok(!emptyXml.includes("<c:val>"));
    assert.ok(!emptyXml.includes("<c:numCache>"));
  });

  it("fills a React handoff chart node and rebuilds the edited document", async () => {
    const expanded = await expandDocument(
      <Deck width={640} height={480}>
        <Slide id="fill-react">
          <Chart
            id="empty-react"
            x={24}
            y={24}
            width={560}
            height={360}
            title="React fill"
            style={{ fontFamily: "Arial" }}
            groups={[
              {
                kind: "line",
                data: { kind: "shared-categories", categories: [] },
                series: [{ name: "Trend", values: [], color: "#117744" }],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const findChart = (parent: CommonNode): CommonNode | undefined => {
      for (const child of parent.children) {
        if (!(child instanceof CommonNode)) continue;
        if (child.kind === "chart") return child;
        const found = findChart(child);
        if (found !== undefined) return found;
      }
      return undefined;
    };
    const node = findChart(expanded);
    assert.ok(node instanceof CommonNode);
    assert.equal(node.props["id"], "empty-react");

    const emptyResult = await buildDocument(expanded);
    const emptyFile = await validatePptx(emptyResult.pptx);
    assert.deepEqual(emptyFile.errors, []);
    assert.deepEqual(
      emptyFile.slides[0]?.objects[0]?.chart?.series[0]?.values,
      [],
    );
    const emptyPptx = emptyResult.pptx.slice();

    node.set({
      groups: [
        {
          kind: "line",
          data: { kind: "shared-categories", categories: ["Jan", "Feb"] },
          series: [{ name: "Trend", values: [10, 30], color: "#117744" }],
        },
      ],
    });
    const filledResult = await buildDocument(expanded);
    const filledFile = await validatePptx(filledResult.pptx);
    assert.deepEqual(filledFile.errors, []);
    const filledChart = filledFile.slides[0]?.objects[0]?.chart;
    assert.deepEqual(filledChart?.series[0]?.categories, ["Jan", "Feb"]);
    assert.deepEqual(filledChart?.series[0]?.values, [10, 30]);
    assert.deepEqual(
      filledChart?.embeddedWorkbook?.series[0]?.values,
      [10, 30],
    );
    assert.ok(
      chartXml(unzipSync(filledResult.pptx)).includes('typeface="Arial"'),
    );
    assert.deepEqual(emptyResult.pptx, emptyPptx);
  });
  it("preserves data and axis diagnostics instead of treating every empty declaration as valid", async () => {
    const base = {
      id: "invalid-empty",
      x: 0,
      y: 0,
      width: 240,
      height: 160,
    } as const;
    const rejectsWithCode = async (
      props: ChartProps,
      code: string,
    ): Promise<void> => {
      await assert.rejects(
        buildDocument(
          deck({ width: 320, height: 240 }).append(
            documentSlide({ id: props.id ?? "invalid-empty" }).append(
              documentChart(props),
            ),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          return true;
        },
      );
    };

    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: [] },
            series: [],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: [] },
            series: [{ name: "S", values: [1] }],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "S", values: [] }],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "scatter",
            data: { kind: "xy" },
            series: [{ name: "S", x: [], y: [1] }],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "scatter",
            data: { kind: "xy" },
            series: [{ name: "S", x: [0], y: [] }],
          },
        ],
      },
      "invalid-data",
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
                name: "S",
                x: [],
                y: [],
                points: { 0: { color: "#336699" } },
              },
            ],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "S", values: [Number.NaN] }],
          },
        ],
      },
      "invalid-data",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category" },
          value: { kind: "value" },
        },
        groups: [
          {
            kind: "column",
            axes: { category: "missing", value: "value" },
            data: { kind: "shared-categories", categories: [] },
            series: [{ name: "S", values: [] }],
          },
        ],
      },
      "unknown-chart-axis",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: { x: { kind: "value" } },
        groups: [
          {
            kind: "scatter",
            data: { kind: "xy" },
            series: [{ name: "S", x: [], y: [] }],
          },
        ],
      },
      "invalid-chart-axes",
    );
  });
  it("keeps invalid empty-cache references and workbook mismatches as reader errors", async () => {
    const result = await build(
      <Deck width={320} height={240}>
        <Slide id="reader-empty">
          <Chart
            id="reader-empty"
            x={16}
            y={16}
            width={288}
            height={180}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [{ name: "Reader", values: [] }],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const source = unzipSync(result.pptx);
    const sourceFile = await validatePptx(result.pptx);
    assert.deepEqual(sourceFile.errors, []);
    const chartPath = "ppt/charts/chart1.xml";
    const workbookPart =
      sourceFile.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const chartPart = strFromU8(source[chartPath]!);
    const revalidate = async (chart: string) =>
      validatePptx(
        zipSync({
          ...source,
          [chartPath]: strToU8(chart),
        }),
      );
    const withSeriesChild = (child: string) =>
      chartPart.replace("</c:ser>", `${child}</c:ser>`);

    const oldInverseRange = await revalidate(
      withSeriesChild(
        '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$1</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache></c:numRef></c:val>',
      ),
    );
    assert.ok(
      oldInverseRange.errors.some(
        ({ code }) => code === "invalid-chart-formula",
      ),
      JSON.stringify(oldInverseRange.errors),
    );

    const singleCategorySource = await revalidate(
      withSeriesChild(
        '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$1</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:cat>',
      ),
    );
    assert.ok(
      singleCategorySource.errors.some(
        ({ code }) => code === "invalid-chart-series",
      ),
      JSON.stringify(singleCategorySource.errors),
    );

    const withCategoryRange = await revalidate(
      withSeriesChild(
        '<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:datalabelsRange><c15:f>Sheet1!$A$2:$A$1</c15:f><c15:dlblRangeCache><c:ptCount val="0"/></c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst>',
      ),
    );
    assert.ok(
      withCategoryRange.errors.some(
        ({ code }) => code === "invalid-chart-data-label-range",
      ),
      JSON.stringify(withCategoryRange.errors),
    );

    const scatterResult = await build(
      <Deck width={320} height={240}>
        <Slide id="reader-scatter-single-source">
          <Chart
            id="reader-scatter-single-source"
            x={16}
            y={16}
            width={288}
            height={180}
            groups={[
              {
                kind: "scatter",
                data: { kind: "xy" },
                series: [{ name: "Scatter", x: [1], y: [2] }],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const scatterParts = unzipSync(scatterResult.pptx);
    const scatterPath = "ppt/charts/chart1.xml";
    const scatterPart = strFromU8(scatterParts[scatterPath]!);
    const scatterWithoutX = scatterPart.replace(
      /<c:xVal>[\s\S]*?<\/c:xVal>/u,
      "",
    );
    const singleScatterSource = await validatePptx(
      zipSync({
        ...scatterParts,
        [scatterPath]: strToU8(scatterWithoutX),
      }),
    );
    assert.ok(
      singleScatterSource.errors.some(
        ({ code }) => code === "invalid-chart-series",
      ),
      JSON.stringify(singleScatterSource.errors),
    );

    const workbookParts = unzipSync(source[workbookPart]!);
    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheetWithData = strFromU8(workbookParts[sheetPath]!).replace(
      "</sheetData>",
      '<row r="2"><c r="B2"><v>1</v></c></row></sheetData>',
    );
    const withWorkbookData = await validatePptx(
      zipSync({
        ...source,
        [workbookPart]: zipSync({
          ...workbookParts,
          [sheetPath]: strToU8(sheetWithData),
        }),
      }),
    );
    assert.ok(
      withWorkbookData.errors.some(
        ({ code }) => code === "chart-workbook-mismatch",
      ),
      JSON.stringify(withWorkbookData.errors),
    );

    const withoutName = strFromU8(workbookParts[sheetPath]!).replace(
      /<c r="B1" t="inlineStr"><is><t>Reader<\/t><\/is><\/c>/u,
      "",
    );
    const missingName = await validatePptx(
      zipSync({
        ...source,
        [workbookPart]: zipSync({
          ...workbookParts,
          [sheetPath]: strToU8(withoutName),
        }),
      }),
    );
    assert.ok(
      missingName.errors.some(
        ({ code }) => code === "incomplete-embedded-workbook",
      ),
      JSON.stringify(missingName.errors),
    );
  });
});
