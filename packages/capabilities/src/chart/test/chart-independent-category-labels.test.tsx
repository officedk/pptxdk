import assert from "node:assert/strict";

import { type ChartProps } from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  deck,
  slide,
} from "@pptxdk/capabilities/document";
import { BuildError, targets } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import {
  validatePptx,
  type FileChart,
  type FileChartDataLabel,
} from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

const independentProps: ChartProps = {
  style: { position: "absolute", left: 0, top: 0, width: 720, height: 420 },
  legend: "right",
  groups: [
    {
      kind: "doughnut",
      data: { kind: "independent-categories" },
      series: [
        {
          name: "Ring A",
          categories: ["Direct", "Partner", "Partner"],
          values: [30, 20, null],
        },
        {
          name: "Ring B",
          categories: ["Partner", "Direct"],
          values: [40, 10],
        },
        {
          name: "Ring C",
          categories: ["Direct", "Partner", "Online", "Partner"],
          values: [5, 6, 7, 8],
        },
      ],
      dataLabels: { value: true, category: true },
    },
  ],
};

const sharedProps: ChartProps = {
  style: { position: "absolute", left: 0, top: 0, width: 720, height: 420 },
  legend: "right",
  groups: [
    {
      kind: "doughnut",
      data: { kind: "shared-categories", categories: sharedCategories() },
      series: [
        { name: "One", values: [1, 2, 3] },
        { name: "Two", values: [3, 2, 1] },
      ],
      dataLabels: { value: true, category: true },
    },
  ],
};

function sharedCategories(): readonly string[] {
  return ["Direct", "Partner", "Online"];
}

function chartXml(pptx: Uint8Array, index = 1): string {
  return strFromU8(unzipSync(pptx)[`ppt/charts/chart${index}.xml`]!);
}

function labelSettings(
  value: false | FileChartDataLabel | undefined,
): FileChartDataLabel | undefined {
  return value === false ? undefined : value;
}

function chartFrom(file: Awaited<ReturnType<typeof validatePptx>>): FileChart {
  const chart = file.slides[0]?.objects.find(
    ({ kind }) => kind === "chart",
  )?.chart;
  assert.ok(chart);
  return chart;
}

async function buildIndependent(
  props: ChartProps = independentProps,
): Promise<Awaited<ReturnType<typeof build>>> {
  return build(
    <Deck>
      <Slide id="independent-c15">
        <Chart id="independent-c15" {...props} />
      </Slide>
    </Deck>,
  );
}

describe("independent category labels", () => {
  it("writes a c15 category field for every independent ring and trailing legend deletion", async () => {
    const result = await buildIndependent();
    assert.deepEqual(result.report.target.extensions, ["p14", "p159", "c15"]);
    assert.deepEqual(result.report.extensions, ["c15"]);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, ["c15"]);
    const withC15 = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159", "c15"],
    });
    assert.deepEqual(withC15.errors, []);
    const withoutC15 = await validatePptx(result.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.ok(
      withoutC15.errors.some(({ code }) => code === "forbidden-extension"),
    );
    const chart = chartFrom(file);
    assert.deepEqual(chart.extensions, ["c15"]);
    assert.deepEqual(
      chart.series.map(
        ({
          name,
          categories,
          values,
          index,
          order,
          datalabelsRange,
          showDataLabelsRange,
        }) => ({
          name,
          categories,
          values,
          index,
          order,
          range: datalabelsRange,
          show: showDataLabelsRange,
        }),
      ),
      [
        {
          name: "Ring A",
          categories: ["Direct", "Partner", "Partner"],
          values: [30, 20, null],
          index: 0,
          order: 0,
          range: {
            formula: "Sheet1!$A$2:$A$4",
            categories: ["Direct", "Partner", "Partner"],
            source: "series-category",
          },
          show: true,
        },
        {
          name: "Ring B",
          categories: ["Partner", "Direct"],
          values: [40, 10],
          index: 1,
          order: 1,
          range: {
            formula: "Sheet1!$C$2:$C$3",
            categories: ["Partner", "Direct"],
            source: "series-category",
          },
          show: true,
        },
        {
          name: "Ring C",
          categories: ["Direct", "Partner", "Online", "Partner"],
          values: [5, 6, 7, 8],
          index: 2,
          order: 2,
          range: {
            formula: "Sheet1!$E$2:$E$5",
            categories: ["Direct", "Partner", "Online", "Partner"],
            source: "series-category",
          },
          show: true,
        },
      ],
    );
    for (const labels of chart.seriesDataLabels ?? []) {
      const settings = labelSettings(labels.dataLabels);
      assert.equal(settings?.category, false);
      assert.equal(settings?.value, true);
    }

    const xml = chartXml(result.pptx);
    assert.equal((xml.match(/<c15:datalabelsRange>/gu) ?? []).length, 3);
    assert.equal(
      (xml.match(/<c15:showDataLabelsRange val="1"\/>/gu) ?? []).length,
      3,
    );
    assert.equal((xml.match(/<c:showCatName val="0"\/>/gu) ?? []).length, 3);
    assert.equal((xml.match(/<c:showCatName val="1"\/>/gu) ?? []).length, 0);
    assert.equal(
      (xml.match(/\{02D57815-91ED-43cb-92C2-25804820EDAC\}/gu) ?? []).length,
      3,
    );
    assert.equal(
      (xml.match(/\{CE6537A1-D6FC-4f65-9D91-7224C49458BB\}/gu) ?? []).length,
      3,
    );
    assert.equal((xml.match(/<c:legendEntry>/gu) ?? []).length, 1);
    assert.ok(
      xml.includes(
        '<c:legendEntry><c:idx val="3"/><c:delete val="1"/></c:legendEntry>',
      ),
    );
    assert.ok(xml.includes('<c:legendPos val="r"/>'));
    assert.ok(!xml.includes("<c:dPt>"));
    assert.ok(!xml.includes("<c:tx><c:strRef><c:f>Sheet1!$C$2</c:f>"));
    assert.ok(!xml.includes("<c:tx><c:strRef><c:f>Sheet1!$E$2</c:f>"));

    const workbook = unzipSync(
      unzipSync(result.pptx)[chart.embeddedWorkbook?.part ?? ""]!,
    );
    const sheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
    assert.ok(
      sheet.includes('<c r="A2" t="inlineStr"><is><t>Direct</t></is></c>'),
    );
    assert.ok(
      sheet.includes('<c r="C2" t="inlineStr"><is><t>Partner</t></is></c>'),
    );
    assert.ok(
      sheet.includes('<c r="E4" t="inlineStr"><is><t>Online</t></is></c>'),
    );
    assert.ok(!sheet.includes('r="B4"'));
  });

  it("deletes every trailing legend entry beyond the first ring category count", async () => {
    const result = await build(
      <Deck>
        <Slide id="legend-tail">
          <Chart
            id="legend-tail"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 640,
              height: 420,
            }}
            legend="right"
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                series: [
                  {
                    name: "First",
                    categories: ["Direct", "Partner"],
                    values: [1, 2],
                  },
                  {
                    name: "Second",
                    categories: ["Direct", "Partner", "Online", "Partner"],
                    values: [1, 2, 3, 4],
                  },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const xml = chartXml(result.pptx);
    assert.deepEqual(
      [
        ...xml.matchAll(
          /<c:legendEntry><c:idx val="(\d+)"\/><c:delete val="1"\/><\/c:legendEntry>/gu,
        ),
      ].map((match) => Number(match[1])),
      [2, 3],
    );
    assert.ok(!xml.includes('<c:legendEntry><c:idx val="0"/>'));
    assert.ok(!xml.includes('<c:legendEntry><c:idx val="1"/>'));
  });

  it("keeps shared-category doughnut bytes free of c15 and trailing legend deletions", async () => {
    const result = await build(
      <Deck>
        <Slide id="shared-no-c15">
          <Chart id="shared-no-c15" {...sharedProps} />
        </Slide>
      </Deck>,
    );
    assert.deepEqual(result.report.extensions, []);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, []);
    const chart = chartFrom(file);
    assert.equal(chart.extensions, undefined);
    const xml = chartXml(result.pptx);
    assert.ok(!xml.includes("datalabelsRange"));
    assert.ok(!xml.includes("showDataLabelsRange"));
    assert.ok(!xml.includes("c15:"));
    assert.ok(!xml.includes("<c:legendEntry>"));
    assert.equal((xml.match(/<c:showCatName val="1"\/>/gu) ?? []).length, 2);
  });

  it("keeps point category overrides and 17 formats while the series range is shared", async () => {
    const result = await build(
      <Deck>
        <Slide id="point-c15">
          <Chart
            id="point-c15"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 720,
              height: 420,
            }}
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                series: [
                  {
                    name: "Ring A",
                    categories: ["Direct", "Partner", "Online"],
                    values: [10, 20, 30],
                    explosion: 12,
                    dataLabels: { leaderLines: true },
                    points: {
                      0: { dataLabels: { category: false } },
                      1: {
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
                    points: {
                      0: { dataLabels: { value: true, category: true } },
                    },
                  },
                ],
                dataLabels: { value: true, category: true, leaderLines: true },
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
    assert.equal(chart.series[0]?.datalabelsRange?.source, "series-category");
    assert.equal(chart.series[1]?.datalabelsRange?.source, "series-category");
    assert.equal(chart.series[0]?.showDataLabelsRange, true);
    assert.equal(chart.series[1]?.showDataLabelsRange, false);
    assert.deepEqual(chart.series[0]?.pointShowDataLabelsRange, {
      0: false,
      1: true,
    });
    assert.deepEqual(chart.series[1]?.pointShowDataLabelsRange, { 0: true });
    const labelsA = chart.seriesDataLabels?.[0];
    const labelsASettings = labelSettings(labelsA?.dataLabels);
    assert.equal(labelsASettings?.value, true);
    assert.equal(labelsASettings?.category, false);
    assert.equal(labelsASettings?.leaderLines, true);
    assert.equal(labelsA?.pointDataLabels?.[2], false);
    const labelsB = chart.seriesDataLabels?.[1];
    const labelsBSeries =
      labelsB?.dataLabels === false ? undefined : labelsB?.dataLabels;
    assert.equal(labelsBSeries?.category, false);
    assert.equal(labelsBSeries?.value, false);
    assert.equal(
      (labelsB?.pointDataLabels?.[0] as { category?: boolean } | undefined)
        ?.category,
      false,
    );
    assert.equal(chart.series[1]?.pointShowDataLabelsRange?.[0], true);
    const xml = chartXml(result.pptx);
    assert.ok(xml.includes('<c15:showDataLabelsRange val="1"/>'));
    assert.ok(xml.includes('<c15:showDataLabelsRange val="0"/>'));
    assert.ok(xml.includes('<c:explosion val="12"/>'));
    assert.ok(xml.includes('<c:explosion val="25"/>'));
    assert.ok(xml.includes("<c:showLeaderLines"));
  });

  it("accepts the range through the command entry and rejects targets that omit c15", async () => {
    const imperative = await buildDocument(
      deck({ width: 720, height: 420 }).append(
        slide({ id: "command-c15" }).append(
          documentChart({
            ...independentProps,
            id: "command-c15",
          }),
        ),
      ),
    );
    const file = await validatePptx(imperative.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.extensions, ["c15"]);
    assert.equal(
      chartFrom(file).series[0]?.datalabelsRange?.formula,
      "Sheet1!$A$2:$A$4",
    );

    for (const target of [
      targets.baselineOnly,
      { ...targets.transitional, extensions: ["p14", "p159"] as const },
    ]) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id="forbidden-c15">
              <Chart id="forbidden-c15" {...independentProps} />
            </Slide>
          </Deck>,
          { target },
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.ok(
            error.diagnostics.some(({ code }) => code === "target-violation"),
          );
          return true;
        },
      );
    }
  });

  it("reads missing field, missing cache and custom ranges without treating them as category", async () => {
    const result = await buildIndependent();
    const parts = unzipSync(result.pptx);
    const original = strFromU8(parts["ppt/charts/chart1.xml"]!);

    const missingRange = original.replace(
      /<c:extLst><c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}"[\s\S]*?<\/c:ext><\/c:extLst>/u,
      "",
    );
    assert.notEqual(missingRange, original);
    const missingRangeFile = await validatePptx(
      zipSync({
        ...parts,
        "ppt/charts/chart1.xml": strToU8(missingRange),
      }),
    );
    assert.ok(
      missingRangeFile.errors.some(
        ({ code }) => code === "invalid-chart-data-label-range",
      ),
    );

    const missingCache = original.replace(
      /<c15:dlblRangeCache>[\s\S]*?<\/c15:dlblRangeCache>/u,
      "",
    );
    assert.notEqual(missingCache, original);
    const missingCacheFile = await validatePptx(
      zipSync({
        ...parts,
        "ppt/charts/chart1.xml": strToU8(missingCache),
      }),
    );
    assert.ok(
      missingCacheFile.errors.some(
        ({ code }) => code === "invalid-chart-data-label-range",
      ),
    );

    const customRange = original
      .replace(
        "<c15:f>Sheet1!$A$2:$A$4</c15:f>",
        "<c15:f>Sheet1!$C$2:$C$3</c15:f>",
      )
      .replace(
        /<c15:dlblRangeCache><c:ptCount val="3"\/>[\s\S]*?<\/c15:dlblRangeCache>/u,
        '<c15:dlblRangeCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Partner</c:v></c:pt><c:pt idx="1"><c:v>Direct</c:v></c:pt></c15:dlblRangeCache>',
      );
    assert.notEqual(customRange, original);
    const customFile = await validatePptx(
      zipSync({
        ...parts,
        "ppt/charts/chart1.xml": strToU8(customRange),
      }),
    );
    assert.deepEqual(customFile.errors, []);
    const customChart = chartFrom(customFile);
    assert.equal(customChart.series[0]?.datalabelsRange?.source, "custom");
    assert.deepEqual(customChart.series[0]?.datalabelsRange?.categories, [
      "Partner",
      "Direct",
    ]);

    const pointResult = await build(
      <Deck>
        <Slide id="point-field">
          <Chart
            id="point-field"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 320,
              height: 240,
            }}
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                series: [
                  {
                    name: "Ring",
                    categories: ["Direct"],
                    values: [1],
                    points: { 0: { dataLabels: { category: true } } },
                  },
                ],
                dataLabels: { value: true, category: true },
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const pointParts = unzipSync(pointResult.pptx);
    const pointXml = strFromU8(pointParts["ppt/charts/chart1.xml"]!);
    const unknownPoint = pointXml.replace(
      '<c:dLbl><c:idx val="0"/>',
      '<c:dLbl><c:idx val="99"/>',
    );
    assert.notEqual(unknownPoint, pointXml);
    const unknownFile = await validatePptx(
      zipSync({
        ...pointParts,
        "ppt/charts/chart1.xml": strToU8(unknownPoint),
      }),
    );
    assert.ok(
      unknownFile.errors.some(
        ({ code, message }) =>
          code === "invalid-chart-data-label-range" &&
          message.includes("unknown point"),
      ),
    );
  });
});
