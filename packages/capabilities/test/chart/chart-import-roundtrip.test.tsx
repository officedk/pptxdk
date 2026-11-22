import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  type ChartNode,
  type ChartProps,
  type DoughnutChartGroup,
  type IndependentCategoryData,
} from "@pptxdk/capabilities/chart";
import { chart, deck, font, slide } from "@pptxdk/capabilities/document";
import { BuildError, defineTheme } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import type { CommonContainer, CommonNode } from "@pptxdk/core/document";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const resolveDocument = createResolver(documentCapabilities);
const importPptx = createImporter(documentCapabilities);

function buildChartDeck(props: ChartProps): Uint8Array {
  const doc = deck({ width: 640, height: 360 }).append(
    slide({ id: "s1" }).append(chart(props)),
  );
  return emit(resolveDocument(doc)).pptx;
}

function slideAt(document: CommonNode, index: number): CommonContainer {
  const container = document as CommonContainer;
  const slides = container.children.filter(
    (c) => "kind" in c && c.kind === "slide",
  );
  return slides[index] as CommonContainer;
}

function chartAt(
  document: CommonNode,
  slideIndex: number,
): {
  readonly node: ChartNode;
  readonly props: ChartProps;
} {
  const node = slideAt(document, slideIndex).children.find(
    (c) => "kind" in c && c.kind === "chart",
  ) as ChartNode;
  return { node, props: node.props as unknown as ChartProps };
}

describe("chart import, editing, and native-23-absent roundtrip", () => {
  it("imports, detaches, edits shared duplicate categories & secondary axes, and paints with regenerated workbook", async () => {
    const initialPptx = buildChartDeck({
      id: "combo",
      x: 20,
      y: 20,
      width: 600,
      height: 320,
      title: "Combo Overview",
      legend: "bottom",
      style: { fontFamily: jostFont, fontSize: 11 },
      axes: {
        period: { kind: "category", title: "Period" },
        amount: { kind: "value", position: "left", min: 0, max: 50 },
        share: {
          kind: "value",
          position: "right",
          min: 0,
          max: 1,
          numberFormat: "0%",
        },
      },
      groups: [
        {
          kind: "column",
          axes: { category: "period", value: "amount" },
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q1", "Q3"],
          },
          series: [{ name: "Actual", values: [12, null, 28] }],
        },
        {
          kind: "line",
          axes: { category: "period", value: "share" },
          data: {
            kind: "shared-categories",
            categories: ["Q1", "Q1", "Q3"],
          },
          series: [{ name: "Rate", values: [0.25, 0.5, 0.75] }],
        },
      ],
    });

    const inputCopy = new Uint8Array(initialPptx);
    const imported = importPptx(inputCopy);
    inputCopy.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);
    const { node: chartNode, props } = chartAt(imported.document, 0);
    assert.equal(chartNode.kind, "chart");
    assert.equal(props.groups.length, 2);

    const editedCategories = ["H1", "H1", "H2"] as const;
    const [colGroup, lineGroup] = props.groups;
    assert.ok(colGroup?.kind === "column");
    assert.ok(lineGroup?.kind === "line");
    const colAxes = colGroup.axes!;
    const lineAxes = lineGroup.axes!;
    assert.ok("category" in colAxes && "value" in colAxes);
    assert.ok("category" in lineAxes && "value" in lineAxes);

    const nextAxes = { ...props.axes };
    const secondaryValAxis = nextAxes[lineAxes.value];
    assert.ok(secondaryValAxis?.kind === "value");
    nextAxes[lineAxes.value] = { ...secondaryValAxis, max: 2 };

    chartNode
      .set({
        axes: nextAxes,
        groups: [
          {
            ...colGroup,
            data: { kind: "shared-categories", categories: editedCategories },
            series: [{ name: "Actual", values: [18, 0, 35] }],
          },
          {
            ...lineGroup,
            data: { kind: "shared-categories", categories: editedCategories },
            series: [{ name: "Rate", values: [0.4, null, 1.2] }],
          },
        ],
      })
      .style({ fontFamily: jostFont, fontSize: 11 });

    const resolution = resolveDocument(imported.document);
    const { pptx } = emit(resolution);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);

    const validation = await validatePptx(pptx);
    assert.deepEqual(validation.errors, []);
    const fileChart = validation.slides[0]?.objects[0]?.chart;
    assert.ok(fileChart);
    assert.deepEqual(fileChart.types, ["barChart", "lineChart"]);
    assert.deepEqual(
      fileChart.series.map(({ name, groupIndex, categories, values }) => ({
        name,
        groupIndex,
        categories,
        values,
      })),
      [
        {
          name: "Actual",
          groupIndex: 0,
          categories: ["H1", "H1", "H2"],
          values: [18, 0, 35],
        },
        {
          name: "Rate",
          groupIndex: 1,
          categories: ["H1", "H1", "H2"],
          values: [0.4, null, 1.2],
        },
      ],
    );
    assert.equal(fileChart.axes[2]?.max, 2);
    assert.deepEqual(fileChart.embeddedWorkbook?.series, fileChart.series);
  });

  it("imports, edits independent doughnut ring categories (c15:datalabelsRange), and paints cleanly", async () => {
    const initialPptx = buildChartDeck({
      id: "rings",
      x: 20,
      y: 20,
      width: 600,
      height: 320,
      legend: "right",
      style: { fontFamily: jostFont, fontSize: 11 },
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
          ],
          dataLabels: { value: true, category: true },
        },
      ],
    });

    const inputCopy = new Uint8Array(initialPptx);
    const imported = importPptx(inputCopy);
    inputCopy.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const { node: chartNode, props } = chartAt(imported.document, 0);
    const rawGroup = props.groups[0]!;
    assert.ok(
      rawGroup.kind === "doughnut" &&
        rawGroup.data.kind === "independent-categories",
    );
    const group = rawGroup as Extract<
      DoughnutChartGroup,
      { readonly data: IndependentCategoryData }
    >;
    chartNode
      .set({
        groups: [
          {
            ...group,
            series: [
              {
                ...group.series[0]!,
                categories: ["Direct", "Referral", "Referral"],
                values: [50, 25, 25],
              },
              group.series[1]!,
            ],
          },
        ],
      })
      .style({ fontFamily: jostFont, fontSize: 11 });

    const resolution = resolveDocument(imported.document);
    const { pptx } = emit(resolution);
    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);

    const validation = await validatePptx(pptx);
    assert.deepEqual(validation.errors, []);
    const fileChart = validation.slides[0]?.objects[0]?.chart;
    assert.ok(fileChart);
    assert.deepEqual(fileChart.extensions, ["c15"]);
    assert.deepEqual(fileChart.series[0]?.categories, [
      "Direct",
      "Referral",
      "Referral",
    ]);
    assert.deepEqual(fileChart.series[0]?.values, [50, 25, 25]);
    assert.deepEqual(fileChart.series[0]?.datalabelsRange?.categories, [
      "Direct",
      "Referral",
      "Referral",
    ]);
    assert.equal(fileChart.series[0]?.showDataLabelsRange, true);
    assert.deepEqual(fileChart.series[1]?.categories, ["Partner", "Direct"]);
    assert.deepEqual(
      fileChart.embeddedWorkbook?.series,
      fileChart.series.map(
        ({ groupIndex, index, name, order, categories, values }) => ({
          groupIndex,
          index,
          name,
          order,
          categories,
          values,
        }),
      ),
    );
  });

  it("imports all 6 native-23-absent PowerPoint & WPS fixtures cleanly, distinguishes [] vs null vs 0, and fills empty series after input zeroing", async () => {
    const emptyFiles = [
      "23-absent-pp-untouched.pptx",
      "23-absent-pp-filled-cells.pptx",
      "23-absent-wps-untouched.pptx",
      "23-absent-wps-filled-cells.pptx",
    ] as const;

    for (const file of emptyFiles) {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`../assets/native-23-absent/${file}`, import.meta.url),
        ),
      );
      const imported = importPptx(bytes);
      bytes.fill(0);

      // Ignore presentation/master-header boilerplate diagnostics while strictly requiring zero diagnostics across all chart and slide parts.
      const chartAndSlideDiagnostics = imported.report.diagnostics.filter(
        (d) =>
          d.location.part.startsWith("ppt/charts/") ||
          /^ppt\/slides\/slide\d+\.xml$/u.test(d.location.part),
      );
      assert.deepEqual(
        chartAndSlideDiagnostics,
        [],
        `${file} must import with zero chart or slide diagnostics`,
      );
      assert.equal(imported.report.slideCount, 10);

      // Slide 1: empty column ([] vs [])
      const colEmpty = chartAt(imported.document, 0).props.groups[0]!;
      assert.equal(colEmpty.kind, "column");
      if (colEmpty.kind === "column") {
        assert.deepEqual(colEmpty.data.categories, []);
        assert.deepEqual(colEmpty.series[0]?.values, []);
      }

      // Slide 2: missing values column (["Q1","Q2","Q3"] vs [null, null, null])
      const colMissing = chartAt(imported.document, 1).props.groups[0]!;
      if (colMissing.kind === "column") {
        assert.deepEqual(colMissing.data.categories, ["Q1", "Q2", "Q3"]);
        assert.deepEqual(colMissing.series[0]?.values, [null, null, null]);
      }

      // Slide 3: zero values column (["Q1","Q2","Q3"] vs [0, 0, 0])
      const colZero = chartAt(imported.document, 2).props.groups[0]!;
      if (colZero.kind === "column") {
        assert.deepEqual(colZero.data.categories, ["Q1", "Q2", "Q3"]);
        assert.deepEqual(colZero.series[0]?.values, [0, 0, 0]);
      }

      // Slide 4: empty multi-group column
      const colMulti = chartAt(imported.document, 3).props.groups;
      assert.equal(colMulti.length, 2);
      assert.equal(colMulti[0]?.series.length, 2);
      assert.equal(colMulti[1]?.series.length, 1);

      // Slide 5: pie empty
      const pieEmpty = chartAt(imported.document, 4).props.groups[0]!;
      assert.equal(pieEmpty.kind, "pie");
      if (pieEmpty.kind === "pie") {
        assert.deepEqual(pieEmpty.data.categories, []);
        assert.deepEqual(pieEmpty.series[0]?.values, []);
      }

      // Slide 6: doughnut empty
      const doughnutEmpty = chartAt(imported.document, 5).props.groups[0]!;
      assert.equal(doughnutEmpty.kind, "doughnut");

      // Slide 7: scatter empty ([] vs [])
      const scatterEmpty = chartAt(imported.document, 6).props.groups[0]!;
      assert.equal(scatterEmpty.kind, "scatter");
      if (scatterEmpty.kind === "scatter") {
        assert.deepEqual(scatterEmpty.series[0]?.x, []);
        assert.deepEqual(scatterEmpty.series[0]?.y, []);
      }

      // Slide 8: scatter missing ([null, null, null])
      const scatterMissing = chartAt(imported.document, 7).props.groups[0]!;
      if (scatterMissing.kind === "scatter") {
        assert.deepEqual(scatterMissing.series[0]?.x, [null, null, null]);
        assert.deepEqual(scatterMissing.series[0]?.y, [null, null, null]);
      }

      // Slide 9: scatter zero ([0, 0, 0])
      const scatterZero = chartAt(imported.document, 8).props.groups[0]!;
      if (scatterZero.kind === "scatter") {
        assert.deepEqual(scatterZero.series[0]?.x, [0, 0, 0]);
        assert.deepEqual(scatterZero.series[0]?.y, [0, 0, 0]);
      }

      // Slide 10: empty multi-series line chart
      const lineMulti = chartAt(imported.document, 9).props.groups;
      assert.equal(lineMulti.length, 1);
      assert.equal(lineMulti[0]?.kind, "line");
      assert.equal(lineMulti[0]?.series.length, 2);
    }

    // Rebound range fixtures (PP & WPS)
    for (const file of [
      "23-absent-pp-filled-range.pptx",
      "23-absent-wps-series-ui.pptx",
    ] as const) {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`../assets/native-23-absent/${file}`, import.meta.url),
        ),
      );
      const imported = importPptx(bytes);
      bytes.fill(0);
      assert.deepEqual(
        imported.report.diagnostics.filter(
          (d) =>
            d.location.part.startsWith("ppt/charts/") ||
            /^ppt\/slides\/slide\d+\.xml$/u.test(d.location.part),
        ),
        [],
        file,
      );

      const s1Col = chartAt(imported.document, 0).props.groups[0]!;
      assert.equal(s1Col.kind, "column");
      if (s1Col.kind === "column") {
        assert.deepEqual(s1Col.data.categories, ["A", "B"]);
        assert.deepEqual(s1Col.series[0]?.values, [3, 7]);
      }

      if (file === "23-absent-pp-filled-range.pptx") {
        const s5Pie = chartAt(imported.document, 4).props.groups[0]!;
        if (s5Pie.kind === "pie") {
          assert.deepEqual(s5Pie.data.categories, ["One", "Two"]);
          assert.deepEqual(s5Pie.series[0]?.values, [2, 5]);
        }
        const s7Scatter = chartAt(imported.document, 6).props.groups[0]!;
        if (s7Scatter.kind === "scatter") {
          assert.deepEqual(s7Scatter.series[0]?.x, [2, 5]);
          assert.deepEqual(s7Scatter.series[0]?.y, [4, 1]);
        }
      }
    }

    // Fill empty series from both untouched PP and WPS fixtures after input zeroing
    for (const untouched of [
      "23-absent-pp-untouched.pptx",
      "23-absent-wps-untouched.pptx",
    ] as const) {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`../assets/native-23-absent/${untouched}`, import.meta.url),
        ),
      );
      const imported = importPptx(bytes);
      bytes.fill(0);

      // Attach a monochrome TTF theme font so all 10 slides paint cleanly with zero diagnostics.
      imported.document.set({
        theme: defineTheme({
          fonts: { major: jostFont, minor: jostFont },
        }),
      });

      const { node: colChart, props: colProps } = chartAt(imported.document, 0);
      const colGroup = colProps.groups[0]!;
      assert.ok(colGroup.kind === "column");
      colChart.set({
        groups: [
          {
            ...colGroup,
            data: {
              kind: "shared-categories",
              categories: ["Jan", "Jan", "Mar"],
            },
            series: [{ ...colGroup.series[0]!, values: [10, null, 25] }],
          },
        ],
      });

      const { node: pieChart, props: pieProps } = chartAt(imported.document, 4);
      const pieGroup = pieProps.groups[0]!;
      assert.ok(pieGroup.kind === "pie");
      pieChart.set({
        groups: [
          {
            ...pieGroup,
            data: { kind: "shared-categories", categories: ["A", "B"] },
            series: [{ ...pieGroup.series[0]!, values: [60, 40] }],
          },
        ],
      });

      const { node: scatterChart, props: scatterProps } = chartAt(
        imported.document,
        6,
      );
      const scatterGroup = scatterProps.groups[0]!;
      assert.ok(scatterGroup.kind === "scatter");
      scatterChart.set({
        groups: [
          {
            ...scatterGroup,
            series: [
              {
                ...scatterGroup.series[0]!,
                x: [1, null, 3],
                y: [2, 4, null],
              },
            ],
          },
        ],
      });

      const resolution = resolveDocument(imported.document);
      const { pptx } = emit(resolution);
      const painted = paint(resolution, { pages: [1, 5, 7] });
      for (const page of painted.pages) {
        assert.deepEqual(
          page.diagnostics,
          [],
          `${untouched} page ${page.page}`,
        );
        assert.ok(page.png instanceof Uint8Array);
      }

      const validation = await validatePptx(pptx);
      assert.deepEqual(validation.errors, [], untouched);
      const s1Chart = validation.slides[0]?.objects[0]?.chart;
      assert.deepEqual(s1Chart?.series[0]?.categories, ["Jan", "Jan", "Mar"]);
      assert.deepEqual(s1Chart?.series[0]?.values, [10, null, 25]);
      assert.deepEqual(s1Chart?.embeddedWorkbook?.series, s1Chart?.series);

      const s7Chart = validation.slides[6]?.objects[0]?.chart;
      assert.deepEqual(s7Chart?.series[0]?.x, [1, null, 3]);
      assert.deepEqual(s7Chart?.series[0]?.y, [2, 4, null]);
      assert.deepEqual(s7Chart?.embeddedWorkbook?.series, s7Chart?.series);
    }
  });

  it("isolates missing/corrupt chart relationships and malformed chart XML while keeping sibling content", () => {
    const basePptx = buildChartDeck({
      id: "healthy",
      x: 10,
      y: 10,
      width: 400,
      height: 240,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A"] },
          series: [{ name: "S", values: [5] }],
        },
      ],
    });
    const parts = unzipSync(basePptx);

    // 1. Corrupt chart XML
    const corruptXml = zipSync({
      ...parts,
      "ppt/charts/chart1.xml": strToU8("<c:chartSpace><broken"),
    });
    const corruptResult = importPptx(corruptXml);
    assert.ok(
      corruptResult.report.diagnostics.some(
        (d) =>
          d.code === "corrupt-part" &&
          d.location.part === "ppt/charts/chart1.xml",
      ),
    );
    assert.equal(
      slideAt(corruptResult.document, 0).children.filter(
        (c) => "kind" in c && c.kind === "chart",
      ).length,
      0,
    );

    // 2. Missing chart target part
    const { "ppt/charts/chart1.xml": _removed, ...withoutChart } = parts;
    const missingPart = zipSync(withoutChart);
    const missingResult = importPptx(missingPart);
    assert.ok(
      missingResult.report.diagnostics.some(
        (d) =>
          d.code === "missing-target-part" &&
          d.targetPart === "ppt/charts/chart1.xml",
      ),
    );

    // 3. Broken chart relationship on slide
    const slideRels = strFromU8(
      parts["ppt/slides/_rels/slide1.xml.rels"]!,
    ).replace(/Id="rId\d+"(?=[^>]*chart1\.xml)/u, 'Id="rId99"');
    const brokenRel = zipSync({
      ...parts,
      "ppt/slides/_rels/slide1.xml.rels": strToU8(slideRels),
    });
    const brokenResult = importPptx(brokenRel);
    assert.ok(
      brokenResult.report.diagnostics.some(
        (d) => d.code === "broken-relationship" && d.location.slideIndex === 0,
      ),
    );
  });

  it("enforces cross-group shared category domain and axis topology invariants in domain operations", () => {
    assert.throws(
      () =>
        resolveDocument(
          deck({ width: 640, height: 360 }).append(
            slide({ id: "mismatch" }).append(
              chart({
                id: "mismatch-chart",
                x: 10,
                y: 10,
                width: 400,
                height: 240,
                axes: {
                  cat: { kind: "category" },
                  val: { kind: "value" },
                },
                groups: [
                  {
                    kind: "column",
                    axes: { category: "cat", value: "val" },
                    data: {
                      kind: "shared-categories",
                      categories: ["A", "B"],
                    },
                    series: [{ name: "S1", values: [1, 2] }],
                  },
                  {
                    kind: "line",
                    axes: { category: "cat", value: "val" },
                    data: {
                      kind: "shared-categories",
                      categories: ["A", "C"],
                    },
                    series: [{ name: "S2", values: [3, 4] }],
                  },
                ],
              }),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          ({ code }) => code === "chart-category-domain-mismatch",
        ),
    );
  });
});
