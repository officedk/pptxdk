import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  build,
  Chart,
  chartCapability,
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
import { describe, it, vi } from "vitest";

describe("native Cartesian charts", () => {
  it("inherits shared text visual styles across charts, axes, and data labels", async () => {
    const result = await build(
      <Deck>
        <Slide id="chart-text-visual">
          <Chart
            id="chart-text-visual"
            kind="column"
            x={0}
            y={0}
            width={240}
            height={160}
            data={{
              categories: ["ßeta"],
              series: [{ name: "Series", values: [1] }],
            }}
            style={{
              fontFamily: "Aptos",
              fontSize: "12pt",
              letterSpacing: 1.25,
              textTransform: "uppercase",
              textDecoration: "underline",
              textShadow: { color: "accent3", blur: 2 },
            }}
            categoryAxis={{
              style: {
                letterSpacing: 2,
                textTransform: "uppercase",
                textDecoration: "underline",
                textShadow: { color: "accent4", blur: 1 },
              },
            }}
            dataLabels={{
              category: true,
              style: {
                letterSpacing: 1.5,
                textTransform: "uppercase",
                textDecoration: "underline",
                textShadow: { color: "accent5", blur: 1 },
              },
            }}
          />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    const parts = unzipSync(result.pptx);
    const chartPath = Object.keys(parts).find((part) =>
      /^ppt\/charts\/chart\d+\.xml$/u.test(part),
    );
    assert.ok(chartPath);
    const chartXml = strFromU8(parts[chartPath]!);
    assert.match(chartXml, /spc="94"/u);
    assert.match(chartXml, /spc="150"/u);
    assert.match(chartXml, /cap="all"/u);
    assert.equal((chartXml.match(/u="sng"/gu) ?? []).length, 3);
    assert.equal((chartXml.match(/<a:outerShdw/gu) ?? []).length, 3);
  });

  it("registers actual font resources for value axes and detects conflicts with chart text", async () => {
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
        <Slide id="value-axis-font">
          <Chart
            id="value-axis-font"
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
          <Slide id="value-axis-font-conflict">
            <Chart
              id="value-axis-font-conflict"
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

  it("preserves basic motion, click steps, and disabled declarations, referencing graphicFrame objects", async () => {
    const result = await build(
      <Deck>
        <Slide id="chart-motion">
          <Chart
            id="entering-chart"
            name="entering-chart"
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
              id="click-chart"
              name="click-chart"
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
            id="disabled-chart"
            name="disabled-chart"
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
                authorId: "entering-chart",
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
                authorId: "entering-chart",
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
                authorId: "click-chart",
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
          <Slide id="invalid-chart-motion">
            <Chart
              id="invalid-motion"
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

  it("writes column charts, chart cache, and embedded workbooks as consistent editable data", async () => {
    const result = await build(
      <Deck width={800} height={600}>
        <Slide id="quarterly-sales">
          <Chart
            id="sales"
            name="sales"
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
      authorId: "sales",
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

  it("writes column, bar, line, and area charts and their embedded workbooks by category", async () => {
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
        <Slide id="chart-types">
          <Chart
            id="column"
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
              position: "insideBase",
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
            id="bar"
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
            id="line"
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
            id="area"
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
        <Slide id="second-page">
          <Chart
            id="second-page-chart"
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
    assert.ok(columnXml.includes('<c:dLblPos val="inBase"/>'));
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

  it("writes native pie and doughnut charts while preserving applicable styles and workbook data", async () => {
    const categories = ["线上 & <门店> 😀", "合作门店", "空值"];
    const pieSeries = {
      name: "渠道 & mix",
      values: [72.5, 27.5, null],
      points: {
        0: { color: "#CC3300", labelColor: "#FFFFFF" },
        1: { color: "#336699" },
        2: { color: "#999999" },
      },
    } as const;
    const doughnutSeries = {
      name: "Q1",
      values: [10, null, 30],
      points: {
        0: { color: "#227744" },
        2: { color: "#AA8822", labelColor: "#FFFFFF" },
      },
    } as const;
    const result = await build(
      <Deck width={800} height={600}>
        <Slide id="pie-and-doughnut">
          <Chart
            id="pie"
            name="pie"
            kind="pie"
            x={20}
            y={20}
            width={350}
            height={350}
            data={{ categories, series: [pieSeries] }}
            explosion={25.5}
            firstSliceAngle={359.5}
            enter={fade({ duration: 500 })}
            legend="top"
            dataLabels={{
              value: true,
              percent: true,
              category: true,
              position: "outsideEnd",
              numberFormat: "0.0%",
              style: {
                fontFamily: "Aptos",
                fontSize: "12pt",
                color: "#112233",
              },
            }}
          />
          <Chart
            id="doughnut"
            name="doughnut"
            kind="doughnut"
            x={410}
            y={20}
            width={350}
            height={350}
            data={{ categories, series: [doughnutSeries] }}
            explosion={400}
            firstSliceAngle={360}
            holeSize={89.5}
            legend={false}
            dataLabels={{
              value: true,
              category: true,
            }}
          />
        </Slide>
      </Deck>,
    );

    const parts = unzipSync(result.pptx);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chartObjects = result.report.slides[0]?.objects ?? [];
    assert.deepEqual(
      chartObjects.map(({ kind, name, authorId }) => [kind, name, authorId]),
      [
        ["chart", "pie", "pie"],
        ["chart", "doughnut", "doughnut"],
      ],
    );
    assert.deepEqual(
      file.slides[0]?.objects.map(({ chart }) => chart?.types),
      [["pieChart"], ["doughnutChart"]],
    );
    assert.deepEqual(file.slides[0]?.objects[0]?.chart?.series, [
      { name: "渠道 & mix", categories, values: [72.5, 27.5, null] },
    ]);
    assert.deepEqual(
      file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.series,
      [{ name: "渠道 & mix", categories, values: [72.5, 27.5, null] }],
    );
    assert.deepEqual(file.slides[0]?.objects[1]?.chart?.series, [
      { name: "Q1", categories, values: [10, null, 30] },
    ]);
    assert.deepEqual(
      file.slides[0]?.objects[1]?.chart?.embeddedWorkbook?.series,
      [{ name: "Q1", categories, values: [10, null, 30] }],
    );

    const pieXml = strFromU8(parts["ppt/charts/chart1.xml"]!);
    assert.ok(pieXml.includes("<c:pieChart>"));
    assert.ok(pieXml.includes('<c:explosion val="26"/>'));
    assert.ok(pieXml.includes('<c:firstSliceAng val="360"/>'));
    assert.ok(pieXml.includes('<c:legendPos val="t"/>'));
    assert.ok(pieXml.includes('<c:dLblPos val="outEnd"/>'));
    assert.ok(
      pieXml.includes('<c:numFmt formatCode="0.0%" sourceLinked="0"/>'),
    );
    assert.ok(pieXml.includes('<c:showVal val="1"/>'));
    assert.ok(pieXml.includes('<c:showPercent val="1"/>'));
    assert.ok(pieXml.includes('<c:showCatName val="1"/>'));
    assert.ok(pieXml.includes('<a:latin typeface="Aptos"/>'));
    assert.ok(pieXml.includes('<a:srgbClr val="CC3300"/>'));
    assert.ok(pieXml.includes('<a:srgbClr val="FFFFFF"/>'));
    assert.ok(pieXml.includes("线上 &amp; &lt;门店&gt; 😀"));
    assert.ok(!pieXml.includes("<c:catAx"));
    assert.ok(!pieXml.includes("<c:valAx"));
    const slideXml = strFromU8(parts["ppt/slides/slide1.xml"]!);
    assert.ok(slideXml.includes(`<p:spTgt spid="${chartObjects[0]?.id}">`));
    assert.ok(!slideXml.includes("<p:bldP"));

    const doughnutXml = strFromU8(parts["ppt/charts/chart2.xml"]!);
    assert.ok(doughnutXml.includes("<c:doughnutChart>"));
    assert.ok(doughnutXml.includes('<c:explosion val="400"/>'));
    assert.ok(doughnutXml.includes('<c:firstSliceAng val="360"/>'));
    assert.ok(doughnutXml.includes('<c:holeSize val="90"/>'));
    assert.ok(!doughnutXml.includes("<c:dLblPos"));
    assert.ok(doughnutXml.includes('<c:showVal val="1"/>'));
    assert.ok(doughnutXml.includes('<c:showCatName val="1"/>'));
    assert.ok(!doughnutXml.includes("<c:legend>"));
    assert.ok(!doughnutXml.includes("<c:catAx"));
    assert.ok(!doughnutXml.includes("<c:valAx"));

    const pieWorkbook = unzipSync(
      parts[file.slides[0]?.objects[0]?.chart?.embeddedWorkbook?.part ?? ""]!,
    );
    const pieSheet = strFromU8(pieWorkbook["xl/worksheets/sheet1.xml"]!);
    for (const value of [
      "线上 &amp; &lt;门店&gt; 😀",
      "合作门店",
      "空值",
      "渠道 &amp; mix",
      "<v>72.5</v>",
      "<v>27.5</v>",
    ])
      assert.ok(pieSheet.includes(value), value);
    assert.ok(!pieSheet.includes("<v></v>"));
  });

  it("restricts final explosion values of pie and doughnut charts under current formal validation profile", async () => {
    const data = {
      categories: ["A", "B"],
      series: [{ name: "S", values: [1, 2] }],
    } as const;
    const acceptedValues = [
      ["zero", 0, 0],
      ["common", 400, 400],
      ["maximum", 2_147_483_647, 2_147_483_647],
      ["rounds-to-maximum", 2_147_483_647.49, 2_147_483_647],
    ] as const;
    const expectedValues = acceptedValues.map(([, , rounded]) => rounded);
    const kinds = ["pie", "doughnut"] as const;
    const result = await build(
      <Deck>
        <Slide id="explosion-supported-range">
          {kinds.flatMap((kind) =>
            acceptedValues.map(([label, explosion]) => (
              <Chart
                key={`${kind}-${label}`}
                id={`${kind}-${label}`}
                kind={kind}
                x={0}
                y={0}
                width={200}
                height={120}
                data={data}
                explosion={explosion}
              />
            )),
          )}
        </Slide>
      </Deck>,
    );

    const parts = unzipSync(result.pptx);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides[0]?.objects.map(({ chart }) => chart?.types),
      [
        ["pieChart"],
        ["pieChart"],
        ["pieChart"],
        ["pieChart"],
        ["doughnutChart"],
        ["doughnutChart"],
        ["doughnutChart"],
        ["doughnutChart"],
      ],
    );
    const chartPaths = Object.keys(parts)
      .filter((path) => /^ppt\/charts\/chart\d+\.xml$/u.test(path))
      .sort(
        (left, right) =>
          Number(left.match(/chart(\d+)/u)?.[1]) -
          Number(right.match(/chart(\d+)/u)?.[1]),
      );
    assert.deepEqual(
      chartPaths.map((path) => {
        const xml = strFromU8(parts[path]!);
        return Number(xml.match(/<c:explosion val="(\d+)"\/>/u)?.[1]);
      }),
      [...expectedValues, ...expectedValues],
    );

    const rejectedValues = [
      ["half-rounds-overflow", 2_147_483_647.5],
      ["first-unrepresentable", 2_147_483_648],
      ["unsigned-integer-maximum", 4_294_967_295],
    ] as const;
    for (const kind of kinds) {
      for (const [label, explosion] of rejectedValues) {
        const testId = `${kind}-explosion-${label}`;
        await assert.rejects(
          build(
            <Deck>
              <Slide id={testId}>
                <Chart
                  id={testId}
                  kind={kind}
                  x={0}
                  y={0}
                  width={200}
                  height={120}
                  data={data}
                  explosion={explosion}
                />
              </Slide>
            </Deck>,
          ),
          (error: unknown) => {
            assert.ok(error instanceof BuildError);
            assert.equal(error.diagnostics[0]?.code, "invalid-chart-option");
            assert.match(
              error.diagnostics[0]?.path ?? "",
              new RegExp(`chart\\[${testId}\\]`, "u"),
            );
            return true;
          },
        );
      }
    }
  });

  it("rejects invalid combinations of pie series, values, and parameters", async () => {
    const base = {
      x: 0,
      y: 0,
      width: 200,
      height: 120,
      data: {
        categories: ["A", "B"],
        series: [{ name: "S", values: [1, 2] }],
      },
    } as const;
    const twoSeries = {
      ...base.data,
      series: [
        { name: "S", values: [1, 2] },
        { name: "T", values: [3, 4] },
      ],
    } as const;
    const invalidCases: {
      readonly id: string;
      readonly code: string;
      readonly message?: string;
      readonly props: ChartProps;
    }[] = [
      {
        id: "pie-series-count",
        code: "invalid-data",
        props: { ...base, kind: "pie", data: twoSeries },
      },
      {
        id: "doughnut-series-count",
        code: "invalid-data",
        props: { ...base, kind: "doughnut", data: twoSeries },
      },
      {
        id: "pie-negative-value",
        code: "invalid-data",
        props: {
          ...base,
          kind: "pie",
          data: {
            categories: ["A", "B"],
            series: [{ name: "S", values: [1, -2] }],
          },
        },
      },
      {
        id: "pie-angle-overflow",
        code: "invalid-chart-option",
        props: { ...base, kind: "pie", firstSliceAngle: 360.1 },
      },
      {
        id: "doughnut-hole-underflow",
        code: "invalid-chart-option",
        props: { ...base, kind: "doughnut", holeSize: 9.9 },
      },
      {
        id: "doughnut-hole-overflow",
        code: "invalid-chart-option",
        props: { ...base, kind: "doughnut", holeSize: 90.1 },
      },
      {
        id: "pie-explosion-underflow",
        code: "invalid-chart-option",
        props: { ...base, kind: "pie", explosion: -1 },
      },
      {
        id: "pie-explosion-overflow",
        code: "invalid-chart-option",
        props: { ...base, kind: "pie", explosion: 4_294_967_296 },
      },
      {
        id: "pie-hole-size",
        code: "invalid-chart-option",
        props: {
          ...base,
          kind: "pie",
          holeSize: 50,
        } as unknown as ChartProps,
      },
      {
        id: "doughnut-stacking",
        code: "invalid-chart-option",
        props: {
          ...base,
          kind: "doughnut",
          stacking: "stacked",
        } as unknown as ChartProps,
      },
      {
        id: "doughnut-label-position",
        code: "invalid-data-labels",
        message:
          "Office does not allow dLblPos for doughnutChart series data labels",
        props: {
          ...base,
          kind: "doughnut",
          dataLabels: { value: true, position: "bestFit" },
        } as unknown as ChartProps,
      },
      {
        id: "pie-inside-base-label-position",
        code: "invalid-data-labels",
        message:
          'Office does not allow dLblPos="inBase" for pieChart series data labels',
        props: {
          ...base,
          kind: "pie",
          dataLabels: { value: true, position: "insideBase" },
        } as unknown as ChartProps,
      },
      {
        id: "column-best-fit-label-position",
        code: "invalid-data-labels",
        message:
          'Office does not allow dLblPos="bestFit" for barChart series data labels',
        props: {
          ...base,
          kind: "column",
          dataLabels: { value: true, position: "bestFit" },
        } as unknown as ChartProps,
      },
      {
        id: "stacked-column-outside-end-label-position",
        code: "invalid-data-labels",
        message:
          'Office does not allow dLblPos="outEnd" for stacked barChart series data labels',
        props: {
          ...base,
          kind: "column",
          stacking: "stacked",
          dataLabels: { value: true, position: "outsideEnd" },
        } as unknown as ChartProps,
      },
      {
        id: "line-explicit-data-label-position",
        code: "invalid-data-labels",
        message:
          "Chart generation does not expose explicit data label positions for line charts",
        props: {
          ...base,
          kind: "line",
          dataLabels: { value: true, position: "center" },
        } as unknown as ChartProps,
      },
    ];

    for (const { id, code, message, props } of invalidCases) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id={id}>
              <Chart id={id} {...props} />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, code);
          if (message !== undefined) {
            assert.equal(error.diagnostics[0]?.message, message);
          }
          assert.match(
            error.diagnostics[0]?.path ?? "",
            new RegExp(`chart\\[${id}\\]`, "u"),
          );
          return true;
        },
      );
    }
  });

  it("validates data label positions by chart group and series ancestor", async () => {
    const result = await build(
      <Deck>
        <Slide id="chart-label-position-context">
          <Chart
            id="line-label-position"
            kind="line"
            x={0}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            dataLabels={{ value: true }}
          />
          <Chart
            id="clustered-label-position"
            kind="column"
            x={220}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            dataLabels={{ value: true, position: "outsideEnd" }}
          />
          <Chart
            id="doughnut-label-position"
            kind="doughnut"
            x={440}
            y={0}
            width={200}
            height={120}
            data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
            dataLabels={{ value: true }}
          />
        </Slide>
      </Deck>,
    );
    const originalParts = unzipSync(result.pptx);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const linePath = "ppt/charts/chart1.xml";
    const clusteredPath = "ppt/charts/chart2.xml";
    const doughnutPath = "ppt/charts/chart3.xml";
    const lineXml = strFromU8(originalParts[linePath]!);
    const clusteredXml = strFromU8(originalParts[clusteredPath]!);
    const doughnutXml = strFromU8(originalParts[doughnutPath]!);
    const lineWithValidPositions = lineXml
      .replace("<c:showLegendKey", '<c:dLblPos val="l"/><c:showLegendKey')
      .replace(
        "</c:ser><c:marker",
        '</c:ser><c:dLbls><c:dLblPos val="bestFit"/></c:dLbls><c:marker',
      );
    assert.notEqual(lineWithValidPositions, lineXml);
    const validParts = unzipSync(result.pptx);
    validParts[linePath] = strToU8(lineWithValidPositions);
    assert.deepEqual((await validatePptx(zipSync(validParts))).errors, []);

    const lineWithInvalidSeriesPosition = lineXml.replace(
      "<c:showLegendKey",
      '<c:dLblPos val="bestFit"/><c:showLegendKey',
    );
    assert.notEqual(lineWithInvalidSeriesPosition, lineXml);
    const invalidLineParts = unzipSync(result.pptx);
    invalidLineParts[linePath] = strToU8(lineWithInvalidSeriesPosition);
    const invalidLine = await validatePptx(zipSync(invalidLineParts));
    assert.ok(
      invalidLine.errors.some(
        ({ code, part }) =>
          code === "invalid-data-label-position" && part === linePath,
      ),
    );

    const stackedWithInvalidPosition = clusteredXml.replace(
      '<c:grouping val="clustered"/>',
      '<c:grouping val="stacked"/>',
    );
    assert.notEqual(stackedWithInvalidPosition, clusteredXml);
    const invalidStackedParts = unzipSync(result.pptx);
    invalidStackedParts[clusteredPath] = strToU8(stackedWithInvalidPosition);
    const invalidStacked = await validatePptx(zipSync(invalidStackedParts));
    assert.ok(
      invalidStacked.errors.some(
        ({ code, part }) =>
          code === "invalid-data-label-position" && part === clusteredPath,
      ),
    );

    const doughnutWithInvalidPosition = doughnutXml.replace(
      "<c:showLegendKey",
      '<c:dLblPos val="ctr"/><c:showLegendKey',
    );
    assert.notEqual(doughnutWithInvalidPosition, doughnutXml);
    const invalidDoughnutParts = unzipSync(result.pptx);
    invalidDoughnutParts[doughnutPath] = strToU8(doughnutWithInvalidPosition);
    const invalidDoughnut = await validatePptx(zipSync(invalidDoughnutParts));
    assert.ok(
      invalidDoughnut.errors.some(
        ({ code, part }) =>
          code === "invalid-data-label-position" && part === doughnutPath,
      ),
    );
  });

  it("reports broken chart-to-workbook part relationships", async () => {
    const result = await build(
      <Deck>
        <Slide id="invalid-workbook-relation">
          <Chart
            id="missing-workbook-relation"
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

  it("detects chart formulas referencing non-existent worksheets in embedded workbooks", async () => {
    const result = await build(
      <Deck>
        <Slide id="renamed-chart-sheet">
          <Chart
            id="renamed-chart-sheet"
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

  it("requires converted widths of axis lines and grid lines to fall within DrawingML range", async () => {
    for (const { id, valueAxis } of [
      { id: "axis-line-overflow", valueAxis: { line: { width: 1e308 } } },
      {
        id: "axis-gridline-overflow",
        valueAxis: { gridlines: { width: 20_116_801 / 9525 } },
      },
    ]) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id={id}>
              <Chart
                id={id}
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
            new RegExp(`chart\\[${id}\\]`, "u"),
          );
          return true;
        },
      );
    }

    const maximumWidth = 20_116_800 / 9525;
    const boundaryResult = await build(
      <Deck>
        <Slide id="maximum-axis-line-width">
          <Chart
            id="maximum-axis-line-width"
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

  it("rejects sparse series values across category indices", async () => {
    const sparseValues = Array<number>(2);
    await assert.rejects(
      build(
        <Deck>
          <Slide id="sparse-chart-data">
            <Chart
              id="sparse-values"
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

  it("rejects numeric parameters without emission semantics on category axes", async () => {
    for (const [option, value] of [
      ["min", 20],
      ["max", 40],
      ["majorUnit", 5],
      ["numberFormat", "0.000"],
    ] as const) {
      await assert.rejects(
        build(
          <Deck>
            <Slide id={`category-axis-${option}`}>
              <Chart
                id="category-axis-option"
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

  it("rejects invalid data and inapplicable parameter combinations with located build errors", async () => {
    const invalidDoughnutLabels: ChartProps = {
      kind: "doughnut",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      // @ts-expect-error -- doughnut chart groups do not allow explicit label positions
      dataLabels: { position: "center" },
    };
    // @ts-expect-error -- pie chart groups do not allow insideBase labels
    const invalidPieLabels: ChartProps = {
      kind: "pie",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      dataLabels: { position: "insideBase" },
    };
    // @ts-expect-error -- bestFit labels are unsupported for bar chart groups
    const invalidColumnLabels: ChartProps = {
      kind: "column",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      dataLabels: { position: "bestFit" },
    };
    // @ts-expect-error -- stacked bar chart groups do not allow outsideEnd labels
    const invalidStackedColumnLabels: ChartProps = {
      kind: "column",
      stacking: "stacked",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      dataLabels: { position: "outsideEnd" },
    };
    const invalidLineLabels: ChartProps = {
      kind: "line",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      // @ts-expect-error -- line chart data label positions are not exposed by Chart
      dataLabels: { position: "center" },
    };
    const invalidAreaLabels: ChartProps = {
      kind: "area",
      data: { categories: ["A"], series: [{ name: "S", values: [1] }] },
      // @ts-expect-error -- area chart groups do not allow explicit label positions
      dataLabels: { position: "insideEnd" },
    };
    void [
      invalidDoughnutLabels,
      invalidPieLabels,
      invalidColumnLabels,
      invalidStackedColumnLabels,
      invalidLineLabels,
      invalidAreaLabels,
    ];

    await assert.rejects(
      build(
        <Deck>
          <Slide id="invalid-data">
            <Chart
              id="mismatched"
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
      // @ts-expect-error -- gapWidth only applies to column and bar charts
      gapWidth: 40,
    };
    const invalidLine = <Chart {...invalidLineProps} />;
    await assert.rejects(
      build(
        <Deck>
          <Slide id="invalid-combination">{invalidLine}</Slide>
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

  it("executes chartCapability.parse exactly once per Chart node during build and reuses parsed model", async () => {
    const parseSpy = vi.spyOn(chartCapability, "parse");
    const compileSpy = vi.spyOn(chartCapability, "compile");

    try {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="single-chart-parse">
            <Chart
              id="piloted-chart"
              kind="column"
              x={50}
              y={50}
              width={400}
              height={300}
              data={{
                categories: ["Q1", "Q2"],
                series: [{ name: "Sales", values: [100, 150] }],
              }}
              style={{ fontSize: 14, color: "#333333" }}
            />
          </Slide>
        </Deck>,
      );

      assert.equal(
        parseSpy.mock.calls.length,
        1,
        `Expected chartCapability.parse to be called exactly 1 time, got ${parseSpy.mock.calls.length}`,
      );

      const parsedResult = parseSpy.mock.results[0]?.value;
      assert.ok(parsedResult);
      assert.equal(parsedResult.contract.nativeKind, "graphicFrame");
      assert.equal(parsedResult.parsed.kind, "column");
      assert.equal(parsedResult.parsed.data.categories.length, 2);

      assert.equal(compileSpy.mock.calls.length, 1);
      assert.equal(
        compileSpy.mock.calls[0]?.[0]?.data,
        parsedResult.parsed.data,
      );

      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
    } finally {
      parseSpy.mockRestore();
      compileSpy.mockRestore();
    }
  });

  it("rejects textWrap, unknown style properties, and unknown parameters on Chart during parse", async () => {
    // 1. textWrap rejection
    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="chart-text-wrap">
            <Chart
              kind="column"
              width={200}
              height={150}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              // @ts-expect-error -- textWrap is explicitly rejected on Chart
              style={{ textWrap: "wrap" }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-style");
        return true;
      },
    );

    // 2. unknown style property rejection
    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="chart-unknown-style">
            <Chart
              kind="column"
              width={200}
              height={150}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              // @ts-expect-error -- testing unknown style property
              style={{ unknownProp: "forbidden" }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-style");
        return true;
      },
    );

    // 3. unknown parameter rejection
    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="chart-unknown-param">
            <Chart
              kind="column"
              width={200}
              height={150}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              // @ts-expect-error -- testing unknown parameter
              unknownParameter="forbidden"
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "unknown-parameter");
        return true;
      },
    );

    // 4. padding rejection
    await assert.rejects(
      build(
        <Deck width={960} height={540}>
          <Slide id="chart-padding">
            <Chart
              kind="column"
              width={200}
              height={150}
              data={{ categories: ["A"], series: [{ name: "S", values: [1] }] }}
              style={{ padding: 10 }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-layout-context");
        return true;
      },
    );
  });
});
