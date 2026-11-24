import assert from "node:assert/strict";

import {
  type ChartProps,
  type ChartSeries,
  type ChartTrendline,
} from "@pptxdk/capabilities/chart";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@pptxdk/capabilities/document";
import { BuildError } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx, type FileChartTrendline } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";

function chartXml(pptx: Uint8Array, index = 1): string {
  return strFromU8(unzipSync(pptx)[`ppt/charts/chart${index}.xml`]!);
}

async function buildInEntry(
  entry: "document" | "react",
  props: ChartProps,
): Promise<Uint8Array> {
  const id = "trendline-chart";
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

function categoryChart(series: readonly ChartSeries[]): ChartProps {
  return {
    x: 10,
    y: 10,
    width: 420,
    height: 260,
    groups: [
      {
        kind: "column",
        data: {
          kind: "shared-categories",
          categories: ["Q1", "Q2", "Q3", "Q4", "Q5"],
        },
        series,
      },
    ],
  };
}

function scatterChart(
  series: Extract<ChartProps["groups"][number], { kind: "scatter" }>["series"],
): ChartProps {
  return {
    x: 10,
    y: 10,
    width: 420,
    height: 260,
    groups: [
      {
        kind: "scatter",
        data: { kind: "xy" },
        line: "straight",
        series,
      },
    ],
  };
}

function trendlines(
  value: readonly FileChartTrendline[] | undefined,
): readonly FileChartTrendline[] {
  assert.ok(value !== undefined);
  return value;
}

describe("trendlines", () => {
  it("writes and reads a linear trendline through both authoring entries", async () => {
    const props = categoryChart([
      {
        name: "Revenue",
        values: [10, 14, 13, 19, 24],
        trendlines: [
          {
            type: "linear",
            intercept: 2,
            forward: 1.5,
            backward: 0.5,
            displayEquation: true,
            displayRSquared: true,
            line: { color: "#225577", width: 1.5, dash: "dash" },
          },
        ],
      },
    ]);
    for (const entry of ["document", "react"] as const) {
      const pptx = await buildInEntry(entry, props);
      const file = await validatePptx(pptx);
      assert.deepEqual(file.errors, [], entry);
      const series = file.slides[0]?.objects[0]?.chart?.series[0];
      const line = trendlines(series?.trendlines)[0];
      assert.ok(line !== undefined);
      assert.equal(line.type, "linear");
      assert.equal(line.intercept, 2);
      assert.equal(line.forward, 1.5);
      assert.equal(line.backward, 0.5);
      assert.equal(line.displayEquation, true);
      assert.equal(line.displayRSquared, true);
      assert.deepEqual(line.line, {
        color: "#225577",
        widthEmu: 14288,
        dash: "dash",
      });
      assert.deepEqual(series?.values, [10, 14, 13, 19, 24]);

      const xml = chartXml(pptx);
      assert.match(
        xml,
        /<c:trendline><c:spPr><a:ln w="14288"><a:solidFill>[\s\S]*?<c:trendlineType val="linear"\/><c:forward val="1.5"\/><c:backward val="0.5"\/><c:intercept val="2"\/><c:dispRSqr val="1"\/><c:dispEq val="1"\/><\/c:trendline>/,
      );
    }
  });

  it("accepts multiple trendlines owned by one series", async () => {
    const pptx = await buildInEntry(
      "document",
      categoryChart([
        {
          name: "Revenue",
          values: [10, 14, 13, 19, 24],
          trendlines: [
            { type: "linear" },
            { type: "polynomial", order: 2, displayEquation: true },
          ],
        },
      ]),
    );
    const file = await validatePptx(pptx);
    assert.deepEqual(file.errors, []);
    const lines = trendlines(
      file.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
    );
    assert.deepEqual(
      lines.map(({ type, order }) => [type, order]),
      [
        ["linear", undefined],
        ["polynomial", 2],
      ],
    );
  });

  it("preserves category labels, null value slots, and series order", async () => {
    const pptx = await buildInEntry(
      "document",
      categoryChart([
        {
          name: "A",
          values: [10, null, 13, 19, 24],
          trendlines: [{ type: "linear" }],
        },
        {
          name: "B",
          values: [1, 2, 3, 4, 5],
          trendlines: [{ type: "linear" }],
        },
      ]),
    );
    const file = await validatePptx(pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.values, [10, null, 13, 19, 24]);
    assert.deepEqual(chart?.series[1]?.values, [1, 2, 3, 4, 5]);
    assert.deepEqual(
      chart?.series.map(({ index, order }) => [index, order]),
      [
        [0, 0],
        [1, 1],
      ],
    );
    assert.deepEqual(chart?.embeddedWorkbook?.series[0]?.values, [
      10,
      null,
      13,
      19,
      24,
    ]);
  });

  it("writes all six native trendline types and their applicable parameters", async () => {
    const category = await buildInEntry(
      "document",
      categoryChart([
        {
          name: "Types",
          values: [2, 4, 8, 16, 32],
          trendlines: [
            { type: "linear", intercept: 1 },
            { type: "logarithmic" },
            { type: "polynomial", order: 3, intercept: 0 },
            { type: "movingAverage", period: 2, forward: 1, backward: 0.5 },
          ],
        },
      ]),
    );
    const categoryFile = await validatePptx(category);
    assert.deepEqual(categoryFile.errors, []);
    assert.deepEqual(
      trendlines(
        categoryFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      ).map(({ type, order, period }) => [type, order, period]),
      [
        ["linear", undefined, undefined],
        ["logarithmic", undefined, undefined],
        ["polynomial", 3, undefined],
        ["movingAverage", undefined, 2],
      ],
    );

    const scatter = await buildInEntry(
      "document",
      scatterChart([
        {
          name: "Types",
          x: [1, 2, 3, 4, 5],
          y: [2, 4, 8, 16, 32],
          trendlines: [
            { type: "exponential", intercept: 2 },
            { type: "power" },
            { type: "movingAverage", period: 2 },
          ],
        },
      ]),
    );
    const scatterFile = await validatePptx(scatter);
    assert.deepEqual(scatterFile.errors, []);
    assert.deepEqual(
      trendlines(
        scatterFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      ).map(({ type, intercept, period }) => [type, intercept, period]),
      [
        ["exponential", 2, undefined],
        ["power", undefined, undefined],
        ["movingAverage", undefined, 2],
      ],
    );
    assert.deepEqual(
      scatterFile.slides[0]?.objects[0]?.chart?.series[0]?.x,
      [1, 2, 3, 4, 5],
    );
    assert.deepEqual(
      scatterFile.slides[0]?.objects[0]?.chart?.series[0]?.y,
      [2, 4, 8, 16, 32],
    );
  });

  it("diagnoses inapplicable trendline structures and parameters", async () => {
    const base = categoryChart([{ name: "S", values: [1, 2, 3, 4, 5] }]);
    const cases: readonly [string, ChartProps][] = [
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "polynomial" } as never],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "polynomial", order: 7 } as never],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "linear", period: 2 } as never],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "power", intercept: 1 } as never],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "movingAverage", period: 5 } as never],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [
              {
                type: "movingAverage",
                period: 2,
                displayEquation: true,
              } as never,
            ],
          },
        ]),
      ],
      [
        "invalid-trendline",
        {
          ...base,
          groups: [
            {
              kind: "column",
              stacking: "stacked",
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3", "Q4", "Q5"],
              },
              series: [
                {
                  name: "S",
                  values: [1, 2, 3, 4, 5],
                  trendlines: [{ type: "linear" }],
                },
              ],
            },
          ],
        },
      ],
      [
        "invalid-trendline",
        {
          ...base,
          groups: [
            {
              kind: "pie",
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3", "Q4", "Q5"],
              },
              series: [
                {
                  name: "S",
                  values: [1, 2, 3, 4, 5],
                  trendlines: [{ type: "linear" }],
                },
              ],
            },
          ],
        },
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "linear", forward: 0 }],
          },
        ]),
      ],
      [
        "invalid-trendline",
        categoryChart([
          {
            name: "S",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "linear", forward: 1.25 }],
          },
        ]),
      ],
    ];
    for (const [code, props] of cases) {
      await rejectsInBothEntries(props, code);
    }
  });

  it("diagnoses exponential, logarithmic, and power value domains", async () => {
    const cases: readonly ChartProps[] = [
      scatterChart([
        {
          name: "Exp",
          x: [1, 2, 3],
          y: [1, 0, 3],
          trendlines: [{ type: "exponential" }],
        },
      ]),
      scatterChart([
        {
          name: "Log",
          x: [1, 0, 3],
          y: [1, 2, 3],
          trendlines: [{ type: "logarithmic" }],
        },
      ]),
      scatterChart([
        {
          name: "Power X",
          x: [1, -2, 3],
          y: [1, 2, 3],
          trendlines: [{ type: "power" }],
        },
      ]),
      scatterChart([
        {
          name: "Power Y",
          x: [1, 2, 3],
          y: [1, 2, -3],
          trendlines: [{ type: "power" }],
        },
      ]),
    ];
    for (const props of cases) {
      await rejectsInBothEntries(props, "invalid-trendline-domain");
    }
  });

  it("reads malformed native trendline XML independently", async () => {
    const result = await buildDocument(
      deck({ width: 360, height: 240 }).append(
        documentSlide({ id: "reader" }).append(
          documentChart({
            ...categoryChart([
              {
                name: "S",
                values: [1, 2, 3, 4, 5],
                trendlines: [{ type: "polynomial", order: 2 }],
              },
            ]),
            id: "reader",
          }),
        ),
      ),
    );
    const parts = unzipSync(result.pptx);
    const path = "ppt/charts/chart1.xml";
    const original = chartXml(result.pptx);

    const expectIssue = async (xml: string, code: string) => {
      const file = await validatePptx(
        zipSync({ ...parts, [path]: strToU8(xml) }),
      );
      assert.ok(
        file.errors.some((error) => error.code === code),
        `expected ${code}; got ${file.errors.map(({ code }) => code).join(",")}`,
      );
    };

    await expectIssue(
      original.replace(
        '<c:trendlineType val="poly"/>',
        '<c:trendlineType val="bogus"/>',
      ),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original.replace('<c:order val="2"/>', '<c:order val="7"/>'),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original.replace('<c:order val="2"/>', ""),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original.replace(
        '<c:trendlineType val="poly"/>',
        '<c:trendlineType val="linear"/>',
      ),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original.replace(
        '<c:order val="2"/>',
        '<c:period val="2"/><c:order val="2"/>',
      ),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original
        .replace(
          '<c:trendlineType val="poly"/>',
          '<c:trendlineType val="power"/>',
        )
        .replace('<c:order val="2"/>', '<c:intercept val="1"/>'),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original.replace(
        '<c:order val="2"/>',
        '<c:forward val="0"/><c:order val="2"/>',
      ),
      "invalid-chart-trendline",
    );
    await expectIssue(
      original
        .replace(
          '<c:trendlineType val="poly"/>',
          '<c:trendlineType val="movingAvg"/>',
        )
        .replace(
          '<c:order val="2"/>',
          '<c:period val="2"/><c:intercept val="0"/>',
        ),
      "invalid-chart-trendline",
    );
  });

  it("accepts ignored moving-average display markers in real PPTX XML", async () => {
    for (const entry of ["document", "react"] as const) {
      const pptx = await buildInEntry(
        entry,
        categoryChart([
          {
            name: "Average",
            values: [1, 2, 3, 4, 5],
            trendlines: [{ type: "movingAverage", period: 2 }],
          },
        ]),
      );
      const parts = unzipSync(pptx);
      const path = "ppt/charts/chart1.xml";
      const original = strFromU8(parts[path]!);
      const anchor = '<c:trendlineType val="movingAvg"/><c:period val="2"/>';
      assert.ok(original.includes(anchor));
      for (const name of ["dispRSqr", "dispEq"] as const) {
        for (const value of ["0", "1"] as const) {
          const xml = original.replace(
            anchor,
            `${anchor}<c:${name} val="${value}"/>`,
          );
          assert.notEqual(xml, original);
          const file = await validatePptx(
            zipSync({ ...parts, [path]: strToU8(xml) }),
          );
          assert.deepEqual(file.errors, [], `${entry} ${name}=${value}`);
          const line = trendlines(
            file.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
          )[0];
          assert.equal(line?.type, "movingAverage");
          assert.equal(
            name === "dispEq" ? line?.displayEquation : line?.displayRSquared,
            value === "1",
            `${entry} ${name}=${value} observation`,
          );
        }
      }
      const both = original.replace(
        anchor,
        `${anchor}<c:dispRSqr val="1"/><c:dispEq val="0"/>`,
      );
      const bothFile = await validatePptx(
        zipSync({ ...parts, [path]: strToU8(both) }),
      );
      assert.deepEqual(bothFile.errors, [], `${entry} both markers`);
      const bothLine = trendlines(
        bothFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      )[0];
      assert.deepEqual(
        [bothLine?.displayRSquared, bothLine?.displayEquation],
        [true, false],
      );

      // The ignored display markers do not relax a real moving-average order error.
      const badOrder = original.replace(
        anchor,
        '<c:trendlineType val="movingAvg"/><c:order val="2"/><c:period val="2"/>',
      );
      const badOrderFile = await validatePptx(
        zipSync({ ...parts, [path]: strToU8(badOrder) }),
      );
      assert.ok(
        badOrderFile.errors.some(
          ({ code }) => code === "invalid-chart-trendline",
        ),
        `${entry} moving-average order must stay invalid`,
      );
    }
  });

  it("keeps charts without analysis declarations unchanged", async () => {
    const pptx = await buildInEntry(
      "document",
      categoryChart([{ name: "S", values: [1, 2, 3, 4, 5] }]),
    );
    const file = await validatePptx(pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(
      file.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      undefined,
    );
    assert.ok(!chartXml(pptx).includes("<c:trendline>"));
  });

  it("captures the declared trendline array at build time and rebuilds from updates", async () => {
    const declared: ChartTrendline[] = [{ type: "linear" }];
    const props = categoryChart([
      { name: "S", values: [1, 2, 3, 4, 5], trendlines: declared },
    ]);
    const first = await buildInEntry("document", props);
    declared[0] = { type: "polynomial", order: 2 };
    const firstFile = await validatePptx(first);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(
      trendlines(
        firstFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      )[0]?.type,
      "linear",
    );
    const second = await buildInEntry("document", props);
    const secondFile = await validatePptx(second);
    assert.deepEqual(secondFile.errors, []);
    assert.deepEqual(
      trendlines(
        secondFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      ).map(({ type, order }) => [type, order]),
      [["polynomial", 2]],
    );
  });

  it("uses half-category forecasts and free scatter X forecasts", async () => {
    const category = await buildInEntry(
      "document",
      categoryChart([
        {
          name: "Category",
          values: [1, 2, 3, 4, 5],
          trendlines: [{ type: "linear", forward: 0.5, backward: 2 }],
        },
      ]),
    );
    const categoryFile = await validatePptx(category);
    assert.deepEqual(categoryFile.errors, []);
    const categoryLine = trendlines(
      categoryFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
    )[0];
    assert.deepEqual([categoryLine?.forward, categoryLine?.backward], [0.5, 2]);

    const scatter = await buildInEntry(
      "document",
      scatterChart([
        {
          name: "Scatter",
          x: [1, 2, 3, 4, 5],
          y: [2, 4, 8, 16, 32],
          trendlines: [{ type: "linear", forward: 0.25 }],
        },
      ]),
    );
    const scatterFile = await validatePptx(scatter);
    assert.deepEqual(scatterFile.errors, []);
    assert.equal(
      trendlines(
        scatterFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
      )[0]?.forward,
      0.25,
    );
  });

  it("writes explicit false display flags and reads them back", async () => {
    const pptx = await buildInEntry(
      "document",
      categoryChart([
        {
          name: "S",
          values: [1, 2, 3, 4, 5],
          trendlines: [
            { type: "linear", displayEquation: false, displayRSquared: false },
          ],
        },
      ]),
    );
    const file = await validatePptx(pptx);
    assert.deepEqual(file.errors, []);
    const line = trendlines(
      file.slides[0]?.objects[0]?.chart?.series[0]?.trendlines,
    )[0];
    assert.deepEqual(
      [line?.displayEquation, line?.displayRSquared],
      [false, false],
    );
    assert.ok(chartXml(pptx).includes('<c:dispRSqr val="0"/>'));
    assert.ok(chartXml(pptx).includes('<c:dispEq val="0"/>'));
  });

  it("reads trendlines under stacked and axisless groups as invalid", async () => {
    const built = async (props: ChartProps) =>
      unzipSync(await buildInEntry("document", props));
    const stacked = await built({
      ...categoryChart([{ name: "S", values: [1, 2, 3] }]),
      groups: [
        {
          kind: "column",
          stacking: "stacked",
          data: {
            kind: "shared-categories",
            categories: ["A", "B", "C"],
          },
          series: [{ name: "S", values: [1, 2, 3] }],
        },
      ],
    });
    const stackedXml = strFromU8(stacked["ppt/charts/chart1.xml"]!).replace(
      "<c:cat>",
      '<c:trendline><c:trendlineType val="linear"/></c:trendline><c:cat>',
    );
    const stackedFile = await validatePptx(
      zipSync({
        ...stacked,
        "ppt/charts/chart1.xml": strToU8(stackedXml),
      }),
    );
    assert.ok(
      stackedFile.errors.some(({ code }) => code === "invalid-chart-trendline"),
    );

    const pie = await built({
      x: 10,
      y: 10,
      width: 320,
      height: 220,
      groups: [
        {
          kind: "pie",
          data: {
            kind: "shared-categories",
            categories: ["A", "B", "C"],
          },
          series: [{ name: "S", values: [1, 2, 3] }],
        },
      ],
    });
    const pieXml = strFromU8(pie["ppt/charts/chart1.xml"]!).replace(
      "<c:cat>",
      '<c:trendline><c:trendlineType val="linear"/></c:trendline><c:cat>',
    );
    const pieFile = await validatePptx(
      zipSync({ ...pie, "ppt/charts/chart1.xml": strToU8(pieXml) }),
    );
    assert.ok(
      pieFile.errors.some(({ code }) => code === "invalid-chart-trendline"),
    );
  });
});
