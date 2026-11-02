import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  build,
  Chart,
  Deck,
  fade,
  font,
  Slide,
  spin,
  Step,
  type ChartProps,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

describe("原生坐标系图表", () => {
  it("登记值轴实际字体资源并检测与图表文字的冲突", async () => {
    const jost = new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    );
    const notoSans = new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    );
    const primaryFont = font(
      "Collision Sans",
      [{ weight: 400, style: "normal", data: jost }],
      { substitute: true },
    );
    const sameFont = font(
      "Collision Sans",
      [{ weight: 400, style: "normal", data: jost }],
      { substitute: true },
    );
    const conflictingFont = font(
      "Collision Sans",
      [{ weight: 400, style: "normal", data: notoSans }],
      { substitute: true },
    );
    const result = await build(
      <Deck>
        <Slide slideId="value-axis-font">
          <Chart
            testId="value-axis-font"
            kind="column"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            style={{ fontFamily: primaryFont }}
            valueAxis={{ style: { fontFamily: sameFont } }}
          />
        </Slide>
      </Deck>,
    );
    const chartPath = Object.keys(unzipSync(result.pptx)).find((path) =>
      /^ppt\/charts\/chart\d+\.xml$/u.test(path),
    );
    assert.ok(chartPath);
    const chartXml = strFromU8(unzipSync(result.pptx)[chartPath]!);
    assert.ok(chartXml.includes("<c:valAx"));
    assert.ok(chartXml.includes('<a:latin typeface="Collision Sans"/>'));
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="value-axis-font-conflict">
            <Chart
              testId="value-axis-font-conflict"
              kind="column"
              x={0}
              y={0}
              width={200}
              height={120}
              data={{
                categories: ["A"],
                series: [{ name: "S", values: [1] }],
              }}
              style={{ fontFamily: primaryFont }}
              valueAxis={{ style: { fontFamily: conflictingFont } }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "conflicting-resource");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /chart\[value-axis-font-conflict\]/u,
        );
        return true;
      },
    );
  });

  it("保留基础动画、点击步与禁用声明，并引用 graphicFrame 对象", async () => {
    const result = await build(
      <Deck>
        <Slide slideId="chart-motion">
          <Chart
            testId="entering-chart"
            kind="column"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            enter={fade({ duration: 700 })}
            emphasis={spin({ duration: 250, from: 180 })}
          />
          <Step n={3}>
            <Chart
              testId="click-chart"
              kind="line"
              x={220}
              y={0}
              width={200}
              height={120}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              enter={fade({ duration: 300 })}
            />
          </Step>
          <Chart
            testId="disabled-chart"
            kind="area"
            x={440}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            enter={fade()}
            motion={false}
          />
        </Slide>
      </Deck>,
    );
    const [entering, clicking, disabled] =
      result.report.slides[0]?.objects ?? [];
    assert.deepEqual(result.report.slides[0]?.capabilities?.["motion"], {
      steps: [
        {
          trigger: "auto",
          effects: [
            {
              target: {
                id: entering?.id,
                name: "entering-chart",
                testId: "entering-chart",
              },
              kind: "entrance",
              preset: "fade",
              durationMs: 700,
              delayMs: 0,
              startMs: 0,
              endMs: 700,
            },
            {
              target: {
                id: entering?.id,
                name: "entering-chart",
                testId: "entering-chart",
              },
              kind: "emphasis",
              preset: "spin",
              durationMs: 250,
              delayMs: 0,
              from: 180,
              startMs: 0,
              endMs: 250,
            },
          ],
        },
        {
          trigger: "click",
          number: 3,
          effects: [
            {
              target: {
                id: clicking?.id,
                name: "click-chart",
                testId: "click-chart",
              },
              kind: "entrance",
              preset: "fade",
              durationMs: 300,
              delayMs: 0,
              startMs: 0,
              endMs: 300,
            },
          ],
        },
      ],
    });

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.animations?.map(({ trigger, effects }) => ({
        trigger,
        effects: effects.map(({ targetId, preset }) => [targetId, preset]),
      })),
      [
        {
          trigger: "auto",
          effects: [
            [entering?.id, "fade"],
            [entering?.id, "spin"],
          ],
        },
        {
          trigger: "click",
          effects: [[clicking?.id, "fade"]],
        },
      ],
    );
    const slideXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.ok(slideXml.includes(`<p:spTgt spid="${entering?.id}">`));
    assert.ok(slideXml.includes(`<p:spTgt spid="${clicking?.id}">`));
    assert.ok(!slideXml.includes(`<p:spTgt spid="${disabled?.id}">`));
    assert.ok(!slideXml.includes("<p:bldP"));

    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="invalid-chart-motion">
            <Chart
              testId="invalid-motion"
              kind="column"
              x={0}
              y={0}
              width={200}
              height={120}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              enter={fade({ duration: 0 })}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-duration");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /chart\[invalid-motion\]/u,
        );
        return true;
      },
    );
  });

  it("将柱形图、图表缓存和嵌入工作簿写成一致的可编辑数据", async () => {
    const result = await build(
      <Deck width={800} height={600}>
        <Slide slideId="quarterly-sales">
          <Chart
            testId="sales"
            kind="column"
            x={32}
            y={48}
            width={560}
            height={360}
            data={{
              categories: ["一季度", "二季度"],
              series: [
                { name: "北区", values: [14, null], color: "#336699" },
                { name: "南区", values: [21, 28] },
              ],
            }}
            legend="right"
            gapWidth={80}
            overlap={-25}
            valueAxis={{ max: 40, majorUnit: 10, numberFormat: "#,##0" }}
            plotArea={{ left: 56, top: 24, width: 440, height: 290 }}
          />
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.slides[0]?.objects[0], {
      id: 2,
      kind: "chart",
      name: "sales",
      testId: "sales",
      box: { x: 32, y: 48, width: 560, height: 360 },
    });

    const parts = unzipSync(result.pptx);
    const chartPath = Object.keys(parts).find((path) =>
      /^ppt\/charts\/chart\d+\.xml$/u.test(path),
    );
    assert.ok(chartPath);
    const chartXml = strFromU8(parts[chartPath]!);
    assert.ok(chartXml.includes("<c:barChart>"));
    assert.ok(chartXml.includes('<c:barDir val="col"/>'));
    assert.ok(chartXml.includes('<c:gapWidth val="80"/>'));
    assert.ok(chartXml.includes('<c:overlap val="-25"/>'));
    assert.ok(chartXml.includes('<c:max val="40"/>'));
    assert.ok(chartXml.includes('<c:majorUnit val="10"/>'));
    assert.ok(chartXml.includes("<c:strCache>"));
    assert.ok(chartXml.includes("北区"));
    assert.ok(chartXml.includes("南区"));
    assert.ok(chartXml.includes("一季度"));
    assert.ok(chartXml.includes("二季度"));
    assert.ok(chartXml.includes("<c:numCache>"));
    assert.ok(chartXml.includes("<c:v>14</c:v>"));
    assert.ok(chartXml.includes("<c:v>21</c:v>"));
    assert.ok(chartXml.includes("<c:v>28</c:v>"));

    const chartRelationships = strFromU8(
      parts[`ppt/charts/_rels/${chartPath.slice("ppt/charts/".length)}.rels`]!,
    );
    const workbookName = chartRelationships.match(
      /Target="\.\.\/embeddings\/([^"]+)"/u,
    )?.[1];
    assert.ok(workbookName);
    const workbook = unzipSync(parts[`ppt/embeddings/${workbookName}`]!);
    const worksheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
    assert.ok(worksheet.includes("一季度"));
    assert.ok(worksheet.includes("二季度"));
    assert.ok(worksheet.includes("北区"));
    assert.ok(worksheet.includes("南区"));
    assert.ok(worksheet.includes("<v>14</v>"));
    assert.ok(worksheet.includes("<v>21</v>"));
    assert.ok(worksheet.includes("<v>28</v>"));

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(file.slides[0]?.objects[0]?.chart, {
      types: ["barChart"],
      direction: "col",
      grouping: "clustered",
      series: [
        {
          name: "北区",
          categories: ["一季度", "二季度"],
          values: [14, null],
        },
        {
          name: "南区",
          categories: ["一季度", "二季度"],
          values: [21, 28],
        },
      ],
      embeddedWorkbook: {
        part: "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
        sheetName: "Sheet1",
        categories: ["一季度", "二季度"],
        series: [
          {
            name: "北区",
            categories: ["一季度", "二季度"],
            values: [14, null],
          },
          {
            name: "南区",
            categories: ["一季度", "二季度"],
            values: [21, 28],
          },
        ],
      },
    });
  });

  it("逐类写出柱形、条形、折线和面积图及其嵌入工作簿", async () => {
    const data = {
      categories: ["FY1", "FY2", "FY3"],
      series: [
        {
          name: "Primary",
          values: [6, 0, null],
          color: "#336699",
          points: { 1: { color: "#FF6600", labelColor: "#FFFFFF" } },
        },
        { name: "Secondary", values: [2, 4, 8] },
      ],
    } as const;
    const result = await build(
      <Deck width={640} height={480}>
        <Slide slideId="chart-types">
          <Chart
            testId="column"
            kind="column"
            x={0}
            y={0}
            width={300}
            height={200}
            data={data}
            stacking="stacked"
            gapWidth={40}
            overlap={-10}
            legend="top"
            dataLabels={{
              value: true,
              category: true,
              position: "outsideEnd",
              style: {
                fontSize: "12pt",
                lineHeight: "20pt",
                verticalAlign: "middle",
              },
            }}
            categoryAxis={{
              visible: true,
              line: { color: "#112233", width: 1 },
              gridlines: true,
            }}
            valueAxis={{
              min: 0,
              max: 10,
              majorUnit: 2,
              numberFormat: "0.0",
              gridlines: { color: "#888888", width: 0.5 },
              line: false,
            }}
            plotArea={{ left: 24, top: 16, width: 260, height: 160 }}
          />
          <Chart
            testId="bar"
            kind="bar"
            x={0}
            y={210}
            width={300}
            height={200}
            data={data}
            stacking="percent"
            gapWidth={100}
            dataLabels={{
              percent: true,
              value: true,
              category: true,
              position: "insideEnd",
            }}
          />
          <Chart
            testId="line"
            kind="line"
            x={310}
            y={0}
            width={300}
            height={200}
            data={data}
            stacking="stacked"
            dataLabels={{ value: true, category: true }}
            style={{ fontFamily: "Aptos", fontSize: "12pt", color: "#224466" }}
          />
          <Chart
            testId="area"
            kind="area"
            x={310}
            y={210}
            width={300}
            height={200}
            data={data}
            stacking="percent"
            legend={false}
          />
        </Slide>
        <Slide slideId="second-page">
          <Chart
            testId="second-page-chart"
            kind="column"
            x={0}
            y={0}
            width={200}
            height={120}
            data={data}
          />
        </Slide>
      </Deck>,
    );

    const parts = unzipSync(result.pptx);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const charts = file.slides[0]?.objects.map(({ chart }) => chart);
    const expectedTypes = ["barChart", "barChart", "lineChart", "areaChart"];
    for (const [index, expectedType] of expectedTypes.entries()) {
      assert.deepEqual(charts?.[index]?.types, [expectedType]);
      assert.deepEqual(charts?.[index]?.series, [
        {
          name: "Primary",
          categories: ["FY1", "FY2", "FY3"],
          values: [6, 0, null],
        },
        {
          name: "Secondary",
          categories: ["FY1", "FY2", "FY3"],
          values: [2, 4, 8],
        },
      ]);
      assert.deepEqual(
        charts?.[index]?.embeddedWorkbook?.series,
        charts?.[index]?.series,
      );
      const workbookPath = charts?.[index]?.embeddedWorkbook?.part;
      assert.ok(workbookPath);
      const workbook = unzipSync(parts[workbookPath]!);
      const worksheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
      for (const value of [
        "FY1",
        "FY2",
        "FY3",
        "Primary",
        "Secondary",
        "<v>6</v>",
        "<v>8</v>",
      ])
        assert.ok(worksheet.includes(value), `${expectedType}: ${value}`);
    }
    assert.equal(charts?.[0]?.direction, "col");
    assert.equal(charts?.[1]?.direction, "bar");
    assert.equal(charts?.[0]?.grouping, "stacked");
    assert.equal(charts?.[1]?.grouping, "percentStacked");
    const columnXml = strFromU8(parts["ppt/charts/chart1.xml"]!);
    assert.ok(columnXml.includes('<c:legendPos val="t"/>'));
    assert.ok(columnXml.includes('<c:gapWidth val="40"/>'));
    assert.ok(columnXml.includes('<c:overlap val="100"/>'));
    assert.ok(columnXml.includes('<c:max val="10"/>'));
    assert.ok(columnXml.includes('<c:majorUnit val="2"/>'));
    assert.ok(columnXml.includes('<a:ln w="9525">'));
    assert.ok(columnXml.includes('<a:ln w="4763">'));
    assert.ok(
      columnXml.includes('<c:numFmt formatCode="0.0" sourceLinked="0"/>'),
    );
    assert.ok(columnXml.includes('<a:srgbClr val="888888"/>'));
    assert.ok(columnXml.includes('<a:srgbClr val="FFFFFF"/>'));
    assert.ok(columnXml.includes("<c:dLbl>"));
    assert.ok(columnXml.includes('<c:dLblPos val="outEnd"/>'));
    assert.ok(columnXml.includes('<a:bodyPr anchor="ctr"/>'));
    assert.ok(columnXml.includes('<a:lnSpc><a:spcPts val="2000"/></a:lnSpc>'));
    assert.ok(
      strFromU8(parts["ppt/charts/chart2.xml"]!).includes(
        '<c:showPercent val="1"/>',
      ),
    );
    assert.ok(
      strFromU8(parts["ppt/charts/chart3.xml"]!).includes("<c:lineChart>"),
    );
    assert.ok(
      strFromU8(parts["ppt/charts/chart4.xml"]!).includes("<c:areaChart>"),
    );
    assert.equal(
      file.slides[1]?.objects[0]?.chart?.embeddedWorkbook?.part,
      "ppt/embeddings/Microsoft_Excel_Worksheet5.xlsx",
    );
    assert.deepEqual(
      file.slides[1]?.objects[0]?.chart?.embeddedWorkbook?.series,
      [
        {
          name: "Primary",
          categories: ["FY1", "FY2", "FY3"],
          values: [6, 0, null],
        },
        {
          name: "Secondary",
          categories: ["FY1", "FY2", "FY3"],
          values: [2, 4, 8],
        },
      ],
    );
  });

  it("报告损坏的图表到工作簿部件关系", async () => {
    const result = await build(
      <Deck>
        <Slide slideId="invalid-workbook-relation">
          <Chart
            testId="missing-workbook-relation"
            kind="line"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
          />
        </Slide>
      </Deck>,
    );
    const parts = unzipSync(result.pptx);
    const relationshipPath = "ppt/charts/_rels/chart1.xml.rels";
    const relationshipXml = strFromU8(parts[relationshipPath]!);
    parts[relationshipPath] = strToU8(
      relationshipXml.replace(
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"',
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"',
      ),
    );

    const file = await validatePptx(zipSync(parts));
    assert.ok(
      file.errors.some(
        ({ code, part }) =>
          code === "invalid-workbook-relationship" &&
          part === "ppt/charts/chart1.xml",
      ),
    );
  });

  it("发现图表公式引用了嵌入工作簿中不存在的工作表", async () => {
    const result = await build(
      <Deck>
        <Slide slideId="renamed-chart-sheet">
          <Chart
            testId="renamed-chart-sheet"
            kind="line"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
          />
        </Slide>
      </Deck>,
    );
    const parts = unzipSync(result.pptx);
    const workbookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
    const workbook = unzipSync(parts[workbookPath]!);
    const workbookXml = strFromU8(workbook["xl/workbook.xml"]!);
    assert.ok(workbookXml.includes('name="Sheet1"'));
    workbook["xl/workbook.xml"] = strToU8(
      workbookXml.replace('name="Sheet1"', 'name="ActuallyOther"'),
    );
    parts[workbookPath] = zipSync(workbook);

    const file = await validatePptx(zipSync(parts));
    assert.equal(
      file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.sheetName,
      "ActuallyOther",
    );
    assert.ok(
      file.errors.some(
        ({ code, part }) =>
          code === "invalid-chart-formula" && part === "ppt/charts/chart1.xml",
      ),
      JSON.stringify(file.errors),
    );
  });

  it("轴线和网格线的换算宽度必须落在 DrawingML 范围内", async () => {
    for (const { testId, valueAxis } of [
      { testId: "axis-line-overflow", valueAxis: { line: { width: 1e308 } } },
      {
        testId: "axis-gridline-overflow",
        valueAxis: { gridlines: { width: 20_116_801 / 9525 } },
      },
    ]) {
      await assert.rejects(
        build(
          <Deck>
            <Slide slideId={testId}>
              <Chart
                testId={testId}
                kind="column"
                x={0}
                y={0}
                width={200}
                height={120}
                data={{
                  categories: ["A"],
                  series: [{ name: "S", values: [1] }],
                }}
                valueAxis={valueAxis}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-style");
          assert.match(
            error.diagnostics[0]?.path ?? "",
            new RegExp(`chart\\[${testId}\\]`, "u"),
          );
          return true;
        },
      );
    }

    const maximumWidth = 20_116_800 / 9525;
    const boundaryResult = await build(
      <Deck>
        <Slide slideId="maximum-axis-line-width">
          <Chart
            testId="maximum-axis-line-width"
            kind="column"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            valueAxis={{ line: { width: maximumWidth } }}
          />
        </Slide>
      </Deck>,
    );
    const boundaryChart = strFromU8(
      unzipSync(boundaryResult.pptx)["ppt/charts/chart1.xml"]!,
    );
    assert.ok(boundaryChart.includes('<a:ln w="20116800">'));
    assert.deepEqual((await validatePptx(boundaryResult.pptx)).errors, []);
  });

  it("按每个分类下标拒绝稀疏系列值", async () => {
    const sparseValues = Array<number>(2);
    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="sparse-chart-data">
            <Chart
              testId="sparse-values"
              kind="column"
              x={0}
              y={0}
              width={200}
              height={120}
              data={{
                categories: ["A", "B"],
                series: [{ name: "S", values: sparseValues }],
              }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-data");
        assert.match(
          error.diagnostics[0]?.path ?? "",
          /chart\[sparse-values\]/u,
        );
        return true;
      },
    );
  });

  it("拒绝分类轴没有写出语义的数值参数", async () => {
    for (const [option, value] of [
      ["min", 20],
      ["max", 40],
      ["majorUnit", 5],
      ["numberFormat", "0.000"],
    ] as const) {
      await assert.rejects(
        build(
          <Deck>
            <Slide slideId={`category-axis-${option}`}>
              <Chart
                testId="category-axis-option"
                kind="column"
                x={0}
                y={0}
                width={200}
                height={120}
                data={{
                  categories: ["A"],
                  series: [{ name: "S", values: [1] }],
                }}
                categoryAxis={{ [option]: value }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-axis");
          assert.match(
            error.diagnostics[0]?.path ?? "",
            /chart\[category-axis-option\]/u,
          );
          return true;
        },
      );
    }
  });

  it("以定位构建错误拒绝无效数据和不适用参数组合", async () => {
    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="invalid-data">
            <Chart
              testId="mismatched"
              kind="column"
              x={0}
              y={0}
              width={200}
              height={120}
              data={{
                categories: ["A", "B"],
                series: [{ name: "S", values: [1] }],
              }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.deepEqual(
          error.diagnostics.map(({ code, path }) => ({ code, path })),
          [
            {
              code: "invalid-data",
              path: "deck/slide[invalid-data]/chart[mismatched]",
            },
          ],
        );
        return true;
      },
    );

    const invalidLineProps: ChartProps = {
      kind: "line",
      x: 0,
      y: 0,
      width: 100,
      height: 100,
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      // @ts-expect-error -- gapWidth 只适用于柱形和条形图
      gapWidth: 40,
    };
    const invalidLine = <Chart {...invalidLineProps} />;
    await assert.rejects(
      build(
        <Deck>
          <Slide slideId="invalid-combination">{invalidLine}</Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-chart-option");
        assert.match(error.diagnostics[0]?.path ?? "", /chart\[1\]/u);
        return true;
      },
    );
  });
});
