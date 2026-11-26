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
import { strFromU8, unzipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

function chartXml(parts: Unzipped, index: number): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

async function rejectsWithCode(
  props: ChartProps,
  code: string,
  id = "invalid-depth-chart",
): Promise<void> {
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

describe("column and bar depth", () => {
  it("writes native titles, labels, number formats, grid styles, and series/point formats", async () => {
    const result = await build(
      <Deck width={640} height={400}>
        <Slide id="chart-depth">
          <Chart
            id="depth"
            title={{
              text: "Quarterly revenue",
              style: {
                fontSize: "16pt",
                fontWeight: 700,
                color: "#1F3A5F",
                align: "center",
                lineHeight: "20pt",
              },
            }}
            displayBlanksAs="zero"
            legend="bottom"
            style={{
              position: "absolute",
              left: 24,
              top: 24,
              width: 560,
              height: 340,
              fontFamily: "Aptos",
              fontSize: "12pt",
              color: "#112233",
            }}
            plotArea={{ left: 72, top: 64, width: 456, height: 220 }}
            axes={{
              category: {
                kind: "category",
                title: "Quarter",
                line: { color: "#999999", width: 1 },
              },
              value: {
                kind: "value",
                title: {
                  text: "Revenue",
                  style: { fontSize: "12pt", fontWeight: 700 },
                },
                min: 0,
                max: 40,
                majorUnit: 10,
                minorUnit: 5,
                numberFormat: "#,##0",
                line: false,
                gridlines: { color: "#CCCCCC", width: 1, dash: "dash" },
                minorGridlines: { color: "#EEEEEE", width: 0.5, dash: "dot" },
              },
            }}
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [
                  {
                    name: "North",
                    values: [10, null, 30],
                    color: "#336699",
                    line: { color: "#112233", width: 2, dash: "dashDot" },
                    points: {
                      0: {
                        color: "#FF6600",
                        line: { color: "#006600", width: 1.5, dash: "dot" },
                        labelColor: "#FFFFFF",
                      },
                    },
                  },
                  { name: "South", values: [12, 18, 20] },
                ],
                gapWidth: 70,
                overlap: -10,
                dataLabels: {
                  value: true,
                  category: true,
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
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts, 1);
    assert.equal((xml.match(/<c:title>/gu) ?? []).length, 3);
    assert.ok(xml.includes('<c:autoTitleDeleted val="0"/>'));
    assert.ok(xml.includes("Quarterly revenue"));
    assert.ok(xml.includes('<a:pPr algn="ctr">'));
    assert.ok(xml.includes('<a:lnSpc><a:spcPts val="2000"/></a:lnSpc>'));
    assert.ok(xml.includes(">Quarter<"));
    assert.ok(xml.includes(">Revenue<"));
    assert.ok(xml.includes('<c:legendPos val="b"/>'));
    assert.ok(xml.includes('<c:gapWidth val="70"/>'));
    assert.ok(xml.includes('<c:overlap val="-10"/>'));
    assert.ok(xml.includes('<c:dLblPos val="outEnd"/>'));
    assert.ok(xml.includes('<c:numFmt formatCode="0.0" sourceLinked="0"/>'));
    assert.ok(xml.includes('<c:numFmt formatCode="#,##0" sourceLinked="0"/>'));
    assert.ok(xml.includes('<c:minorUnit val="5"/>'));
    assert.ok(xml.includes("<c:minorGridlines>"));
    assert.ok(xml.includes('<a:prstDash val="dash"/>'));
    assert.ok(xml.includes('<a:prstDash val="dashDot"/>'));
    assert.ok(xml.includes('<a:prstDash val="dot"/>'));
    assert.ok(xml.includes('<a:ln w="19050">'));
    assert.ok(xml.includes('<a:ln w="14288">'));
    assert.ok(xml.includes('<c:dispBlanksAs val="zero"/>'));

    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.ok(!slideXml.includes("Quarterly revenue"));
    assert.ok(!slideXml.includes(">Revenue<"));

    const chart = file.slides[0]?.objects[0]?.chart;
    assert.equal(chart?.title, "Quarterly revenue");
    assert.equal(chart?.displayBlanksAs, "zero");
    assert.equal(
      chart?.axes.find(({ kind }) => kind === "category")?.title,
      "Quarter",
    );
    assert.equal(
      chart?.axes.find(({ kind }) => kind === "value")?.title,
      "Revenue",
    );
    assert.equal(
      chart?.axes.find(({ kind }) => kind === "value")?.minorUnit,
      5,
    );
    assert.deepEqual(chart?.series[0]?.values, [10, null, 30]);
  });

  it("writes clustered, stacked, and percent variants in both directions", async () => {
    const data = {
      kind: "shared-categories" as const,
      categories: ["A", "B"],
    };
    const series = [
      { name: "One", values: [3, 7] },
      { name: "Two", values: [5, 2] },
    ];
    const variants = [
      { id: "column-clustered", kind: "column", stacking: undefined },
      { id: "column-stacked", kind: "column", stacking: "stacked" },
      { id: "column-percent", kind: "column", stacking: "percent" },
      { id: "bar-clustered", kind: "bar", stacking: undefined },
      { id: "bar-stacked", kind: "bar", stacking: "stacked" },
      { id: "bar-percent", kind: "bar", stacking: "percent" },
    ] as const;
    const result = await build(
      <Deck width={660} height={420}>
        <Slide id="chart-variants">
          {variants.map((variant, index) => (
            <Chart
              key={variant.id}
              id={variant.id}
              style={{
                position: "absolute",
                left: 12 + (index % 3) * 214,
                top: 12 + Math.floor(index / 3) * 196,
                width: 200,
                height: 180,
              }}
              displayBlanksAs={variant.stacking === undefined ? "gap" : "zero"}
              groups={[
                {
                  kind: variant.kind,
                  data,
                  series,
                  ...(variant.stacking === undefined
                    ? {}
                    : { stacking: variant.stacking }),
                  gapWidth: 60,
                  ...(variant.id === "column-stacked" ? { overlap: 40 } : {}),
                },
              ]}
            />
          ))}
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const expected = [
      ["col", "clustered"],
      ["col", "stacked"],
      ["col", "percentStacked"],
      ["bar", "clustered"],
      ["bar", "stacked"],
      ["bar", "percentStacked"],
    ] as const;
    for (const [index, [direction, grouping]] of expected.entries()) {
      const xml = chartXml(parts, index + 1);
      assert.ok(xml.includes(`<c:barDir val="${direction}"/>`));
      assert.ok(xml.includes(`<c:grouping val="${grouping}"/>`));
      assert.ok(xml.includes('<c:gapWidth val="60"/>'));
      const group = file.slides[0]?.objects[index]?.chart?.groups[0];
      assert.equal(group?.direction, direction);
      assert.equal(group?.grouping, grouping);
    }
    assert.ok(
      chartXml(parts, 2).includes('<c:overlap val="40"/>'),
      "explicit stacked overlap must not be silently replaced",
    );
    assert.ok(chartXml(parts, 3).includes('<c:overlap val="100"/>'));
    assert.ok(chartXml(parts, 5).includes('<c:overlap val="100"/>'));
    assert.ok(chartXml(parts, 6).includes('<c:overlap val="100"/>'));
  });

  it("writes axis direction, ticks, crossings, and logarithmic value rules", async () => {
    const result = await build(
      <Deck width={480} height={300}>
        <Slide id="axis-depth">
          <Chart
            id="axis-depth"
            style={{
              position: "absolute",
              left: 16,
              top: 16,
              width: 440,
              height: 260,
            }}
            title="Log revenue"
            axes={{
              category: {
                kind: "category",
                title: "Quarter",
                crosses: "max",
                majorTickMark: "in",
                tickLabelPosition: "high",
              },
              value: {
                kind: "value",
                title: "Revenue",
                reversed: true,
                logBase: 10,
                min: 1,
                max: 1000,
                majorTickMark: "out",
                minorTickMark: "cross",
                tickLabelPosition: "low",
                crossesAt: 1,
              },
            }}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A", "B"] },
                series: [{ name: "S", values: [10, 100] }],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const xml = chartXml(unzipSync(result.pptx), 1);
    assert.ok(xml.includes('<c:logBase val="10"/>'));
    assert.equal(
      (xml.match(/<c:orientation val="maxMin"\/>/gu) ?? []).length,
      1,
    );
    assert.ok(xml.includes('<c:min val="1"/>'));
    assert.ok(xml.includes('<c:max val="1000"/>'));
    assert.ok(xml.includes('<c:majorTickMark val="out"/>'));
    assert.ok(xml.includes('<c:minorTickMark val="cross"/>'));
    assert.ok(xml.includes('<c:tickLblPos val="low"/>'));
    assert.ok(xml.includes('<c:majorTickMark val="in"/>'));
    assert.ok(xml.includes('<c:tickLblPos val="high"/>'));
    assert.ok(xml.includes('<c:crosses val="max"/>'));
    assert.ok(xml.includes('<c:crossesAt val="1"/>'));
    assert.ok(!xml.includes('<c:crosses val="autoZero"/><c:crossesAt'));

    const chart = file.slides[0]?.objects[0]?.chart;
    const category = chart?.axes.find(({ kind }) => kind === "category");
    const value = chart?.axes.find(({ kind }) => kind === "value");
    assert.equal(category?.title, "Quarter");
    assert.equal(category?.majorTickMark, "in");
    assert.equal(category?.tickLabelPosition, "high");
    assert.equal(category?.crosses, undefined);
    assert.equal(category?.crossesAt, 1);
    assert.equal(value?.title, "Revenue");
    assert.equal(value?.orientation, "maxMin");
    assert.equal(value?.logBase, 10);
    assert.equal(value?.min, 1);
    assert.equal(value?.max, 1000);
    assert.equal(value?.majorTickMark, "out");
    assert.equal(value?.minorTickMark, "cross");
    assert.equal(value?.tickLabelPosition, "low");
    assert.equal(value?.crosses, "max");
    assert.equal(value?.crossesAt, undefined);
  });

  it("diagnoses inapplicable or invalid column/bar depth options", async () => {
    const base = {
      style: { position: "absolute", left: 0, top: 0, width: 200, height: 120 },
      groups: [
        {
          kind: "column" as const,
          data: { kind: "shared-categories" as const, categories: ["A"] },
          series: [{ name: "S", values: [1] }],
        },
      ],
    } satisfies ChartProps;

    await rejectsWithCode({ ...base, displayBlanksAs: "span" }, "invalid-prop");
    await rejectsWithCode(
      {
        ...base,
        displayBlanksAs: "gap",
        groups: [
          {
            kind: "line",
            data: { kind: "shared-categories", categories: ["A"] },
            series: [{ name: "S", values: [1] }],
            stacking: "stacked",
          },
        ],
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category", logBase: 10 } as never,
          value: { kind: "value" },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category" },
          value: { kind: "value", logBase: 1 },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category" },
          value: { kind: "value", logBase: 10, min: 0 },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category" },
          value: { kind: "value", crosses: "max", crossesAt: 1 },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category", crossesAt: 1.5 },
          value: { kind: "value" },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category" },
          value: { kind: "value", majorTickMark: "sideways" as never },
        },
      },
      "invalid-prop",
    );
    await rejectsWithCode({ ...base, title: "" }, "invalid-prop");
    await rejectsWithCode(
      {
        ...base,
        axes: {
          category: { kind: "category", title: "" },
          value: { kind: "value" },
        },
      },
      "invalid-prop",
    );
  });

  it("keeps null slots and workbook data while writing explicit blank rendering", async () => {
    const result = await build(
      <Deck width={500} height={300}>
        <Slide id="missing-depth">
          <Chart
            id="missing-column"
            style={{
              position: "absolute",
              left: 10,
              top: 10,
              width: 220,
              height: 180,
            }}
            displayBlanksAs="zero"
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B", "C"],
                },
                series: [{ name: "S", values: [5, null, 15] }],
                stacking: "stacked",
              },
            ]}
          />
          <Chart
            id="missing-bar"
            style={{
              position: "absolute",
              left: 250,
              top: 10,
              width: 220,
              height: 180,
            }}
            displayBlanksAs="gap"
            groups={[
              {
                kind: "bar",
                data: {
                  kind: "shared-categories",
                  categories: ["A", "B", "C"],
                },
                series: [{ name: "S", values: [null, 25, 75] }],
                stacking: "percent",
              },
            ]}
          />
        </Slide>
      </Deck>,
    );

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const columnXml = chartXml(parts, 1);
    const barXml = chartXml(parts, 2);
    assert.ok(columnXml.includes('<c:dispBlanksAs val="zero"/>'));
    assert.ok(barXml.includes('<c:dispBlanksAs val="gap"/>'));
    assert.deepEqual(file.slides[0]?.objects[0]?.chart?.series[0]?.values, [
      5,
      null,
      15,
    ]);
    assert.deepEqual(file.slides[0]?.objects[1]?.chart?.series[0]?.values, [
      null,
      25,
      75,
    ]);
    assert.equal(file.slides[0]?.objects[0]?.chart?.displayBlanksAs, "zero");
    assert.equal(file.slides[0]?.objects[1]?.chart?.displayBlanksAs, "gap");
    assert.deepEqual(
      file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.series[0]?.values,
      [5, null, 15],
    );
    assert.deepEqual(
      file.slides[0]?.objects[1]?.chart?.embeddedWorkbook?.series[0]?.values,
      [null, 25, 75],
    );
    const workbookPart =
      file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.part ?? "";
    const worksheet = strFromU8(
      unzipSync(parts[workbookPart]!)["xl/worksheets/sheet1.xml"]!,
    );
    assert.ok(!worksheet.includes("<v></v>"));
  });

  it("overrides series and point label display, position, and number format through both authoring entries", async () => {
    const props: ChartProps = {
      id: "label-overrides",
      style: { position: "absolute", left: 0, top: 0, width: 360, height: 240 },
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A", "B", "C"] },
          dataLabels: {
            value: true,
            category: true,
            position: "outsideEnd",
            numberFormat: "0.00",
          },
          series: [
            {
              name: "Override",
              values: [10, 20, 30],
              dataLabels: { value: true, numberFormat: "0.0" },
              points: {
                1: {
                  dataLabels: { numberFormat: "0.000", position: "insideEnd" },
                },
                2: { dataLabels: false },
              },
            },
            { name: "Hidden", values: [5, 6, 7], dataLabels: false },
            {
              name: "Category",
              values: [1, 2, 3],
              dataLabels: { value: false, category: true, numberFormat: "0" },
            },
          ],
        },
      ],
    };
    const results = [
      await build(
        <Deck width={480} height={320}>
          <Slide id="react-label-overrides">
            <Chart {...props} />
          </Slide>
        </Deck>,
      ),
      await buildDocument(
        deck({ width: 480, height: 320 }).append(
          documentSlide({ id: "document-label-overrides" }).append(
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
            value: true,
            category: true,
            percent: false,
            position: "outsideEnd",
            numberFormat: "0.0",
          },
          pointDataLabels: {
            1: {
              value: true,
              category: true,
              percent: false,
              numberFormat: "0.000",
              position: "insideEnd",
            },
            2: false,
          },
        },
        { seriesIndex: 1, dataLabels: false },
        {
          seriesIndex: 2,
          dataLabels: {
            value: false,
            category: true,
            percent: false,
            position: "outsideEnd",
            numberFormat: "0",
          },
        },
      ]);

      const xml = chartXml(unzipSync(result.pptx), 1);
      const seriesXml = [...xml.matchAll(/<c:ser>(.*?)<\/c:ser>/gsu)].map(
        (match) => match[1]!,
      );
      assert.equal(seriesXml.length, 3);
      assert.ok(
        seriesXml[0]!.includes('<c:numFmt formatCode="0.0" sourceLinked="0"/>'),
      );
      assert.ok(
        seriesXml[0]!.includes(
          '<c:dLbl><c:idx val="1"/><c:numFmt formatCode="0.000" sourceLinked="0"/>',
        ),
      );
      assert.ok(seriesXml[0]!.includes('<c:dLblPos val="inEnd"/>'));
      assert.ok(
        seriesXml[0]!.includes(
          '<c:dLbl><c:idx val="2"/><c:delete val="1"/></c:dLbl>',
        ),
      );
      assert.ok(
        seriesXml[1]!.includes('<c:dLbls><c:delete val="1"/></c:dLbls>'),
      );
      assert.ok(seriesXml[2]!.includes('<c:showVal val="0"/>'));
      assert.ok(seriesXml[2]!.includes('<c:showCatName val="1"/>'));
      assert.ok(
        seriesXml[2]!.includes('<c:numFmt formatCode="0" sourceLinked="0"/>'),
      );
      assert.ok(seriesXml[0]!.includes('<c:showPercent val="0"/>'));
      assert.ok(seriesXml[2]!.includes('<c:showPercent val="0"/>'));
      assert.ok(!xml.includes('<c:showPercent val="1"/>'));
    }
  });

  it("diagnoses percent data labels outside pie and doughnut charts", async () => {
    const data = {
      kind: "shared-categories" as const,
      categories: ["A"],
    };
    const base = {
      style: { position: "absolute", left: 0, top: 0, width: 200, height: 120 },
    } as const;
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data,
            series: [{ name: "S", values: [1] }],
            dataLabels: { percent: true },
          },
        ],
      },
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "bar",
            data,
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
      "invalid-prop",
    );
    await rejectsWithCode(
      {
        ...base,
        groups: [
          {
            kind: "column",
            data,
            series: [
              {
                name: "S",
                values: [1],
                points: { 0: { dataLabels: { percent: true } } },
              },
            ],
          },
        ],
      },
      "invalid-prop",
    );
  });
});
