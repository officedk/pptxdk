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
import { BuildError, DeckContainer, defineTheme, isTheme } from "@pptxdk/core";
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
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
  {
    weight: 700,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-700-Bold.ttf", import.meta.url),
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
          new URL(`./fixtures/native-23-absent/${file}`, import.meta.url),
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
          new URL(`./fixtures/native-23-absent/${file}`, import.meta.url),
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
          new URL(`./fixtures/native-23-absent/${untouched}`, import.meta.url),
        ),
      );
      const imported = importPptx(bytes);
      bytes.fill(0);

      // Attach a monochrome TTF theme font so all 10 slides paint cleanly with zero diagnostics.
      const fontFallbackTheme = defineTheme({
        fonts: { major: jostFont, minor: jostFont },
      });
      imported.document.set({ theme: fontFallbackTheme });
      if (imported.document instanceof DeckContainer) {
        for (const master of imported.document.masters) {
          if (isTheme(master.theme)) {
            master.theme.set({ fonts: { major: jostFont, minor: jostFont } });
          }
        }
      }

      const bindChartFont = (props: ChartProps) => {
        const mapStyle = (style: ChartProps["style"] | undefined) =>
          style === undefined ? undefined : { ...style, fontFamily: jostFont };
        const mapTitle = (title: ChartProps["title"]) => {
          if (title === undefined || typeof title === "string") return title;
          const nextStyle = mapStyle(title.style);
          return nextStyle === undefined
            ? { text: title.text }
            : { ...title, style: nextStyle };
        };
        const mapAxes = (axes: ChartProps["axes"]) => {
          if (axes === undefined) return undefined;
          const entries = Object.entries(axes).map(([key, ax]) => {
            const nextStyle = mapStyle(ax.style);
            const nextTitle = mapTitle(ax.title);
            return [
              key,
              {
                ...ax,
                ...(nextStyle === undefined ? {} : { style: nextStyle }),
                ...(nextTitle === undefined ? {} : { title: nextTitle }),
              },
            ];
          });
          return Object.fromEntries(entries) as NonNullable<ChartProps["axes"]>;
        };
        const nextTitle = mapTitle(props.title);
        const nextAxes = mapAxes(props.axes);
        return {
          ...(nextTitle === undefined ? {} : { title: nextTitle }),
          ...(nextAxes === undefined ? {} : { axes: nextAxes }),
        };
      };

      const { node: colChart, props: colProps } = chartAt(imported.document, 0);
      const colGroup = colProps.groups[0]!;
      assert.ok(colGroup.kind === "column");
      colChart
        .set({
          ...bindChartFont(colProps),
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
        })
        .style({ ...colProps.style, fontFamily: jostFont });

      const { node: pieChart, props: pieProps } = chartAt(imported.document, 4);
      const pieGroup = pieProps.groups[0]!;
      assert.ok(pieGroup.kind === "pie");
      pieChart
        .set({
          ...bindChartFont(pieProps),
          groups: [
            {
              ...pieGroup,
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ ...pieGroup.series[0]!, values: [60, 40] }],
            },
          ],
        })
        .style({ ...pieProps.style, fontFamily: jostFont });

      const { node: scatterChart, props: scatterProps } = chartAt(
        imported.document,
        6,
      );
      const scatterGroup = scatterProps.groups[0]!;
      assert.ok(scatterGroup.kind === "scatter");
      scatterChart
        .set({
          ...bindChartFont(scatterProps),
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
        })
        .style({ ...scatterProps.style, fontFamily: jostFont });

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

  it("roundtrips chart presentation, 3-level labels, point overrides, trendlines, error bars, and axis crossings without leakage", async () => {
    const initialDoc = deck({
      width: 640,
      height: 360,
      theme: defineTheme({
        colors: { accent1: "#2563EB", accent2: "#F59E0B", dk1: "#0F172A" },
        fonts: { major: jostFont, minor: jostFont },
      }),
    }).append(
      slide({ id: "presentation-cartesian" }).append(
        chart({
          id: "col-line-rich",
          x: 20,
          y: 20,
          width: 600,
          height: 320,
          title: {
            text: "Rich Presentation",
            style: { fontSize: 16, fontWeight: "bold", color: "#1D4ED8" },
          },
          legend: "topRight",
          displayBlanksAs: "gap",
          style: { fontFamily: jostFont, fontSize: 11, color: "#0F172A" },
          axes: {
            cat: {
              kind: "category",
              position: "top",
              reversed: true,
              majorTickMark: "out",
              tickLabelPosition: "high",
              title: { text: "Period", style: { color: "#334155" } },
              line: { color: "#475569", width: 1.5, dash: "dashDot" },
            },
            val: {
              kind: "value",
              position: "left",
              reversed: true,
              min: 0,
              max: 100,
              majorUnit: 25,
              minorUnit: 5,
              numberFormat: "$#,##0",
              crosses: "min",
              gridlines: { color: "#CBD5E1", width: 1, dash: "dash" },
              minorGridlines: true,
              style: { fontSize: 10, color: "#334155" },
              title: "Amount",
            },
            cat2: {
              kind: "category",
              visible: false,
            },
            val2: {
              kind: "value",
              position: "right",
              min: 0,
              max: 50,
              crossesAt: 2,
              numberFormat: "0.0",
            },
          },
          groups: [
            {
              kind: "column",
              axes: { category: "cat", value: "val" },
              gapWidth: 80,
              overlap: -15,
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3", "Q4"],
              },
              dataLabels: {
                value: true,
                position: "outsideEnd",
                numberFormat: "#,##0",
                style: { color: "#0F172A", textDecoration: "underline" },
              },
              series: [
                {
                  name: "Primary",
                  color: "#2563EB",
                  values: [20, 45, 65, 80],
                  points: {
                    1: {
                      color: "#DC2626",
                      labelColor: "#FFFFFF",
                      dataLabels: { position: "insideEnd" },
                    },
                    3: { dataLabels: false },
                  },
                  trendlines: [
                    {
                      type: "polynomial",
                      order: 2,
                      intercept: 5,
                      forward: 1,
                      backward: 0.5,
                      displayEquation: true,
                      displayRSquared: true,
                      line: { color: "#EF4444", width: 2, dash: "dash" },
                    },
                  ],
                  errorBars: {
                    type: "fixed",
                    value: 4,
                    include: "plus",
                    endCap: false,
                    line: { color: "#111827", width: 1.5 },
                  },
                },
                {
                  name: "Muted",
                  color: "#94A3B8",
                  values: [15, 30, 40, 50],
                  dataLabels: false,
                },
              ],
            },
            {
              kind: "line",
              axes: { category: "cat2", value: "val2" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3", "Q4"],
              },
              series: [
                {
                  name: "Signal",
                  color: "#10B981",
                  line: { width: 2, dash: "longDash" },
                  smooth: true,
                  marker: {
                    symbol: "diamond",
                    size: 7,
                    fill: "#10B981",
                    line: { color: "#064E3B", width: 1 },
                  },
                  values: [10, 20, 30, 40],
                  points: {
                    2: {
                      marker: { symbol: "star", size: 10, fill: "#F59E0B" },
                    },
                  },
                },
              ],
            },
          ],
        }),
      ),
      slide({ id: "presentation-scatter-pie" }).append(
        chart({
          id: "scatter-rich",
          x: 16,
          y: 16,
          width: 300,
          height: 300,
          plotArea: { left: 48, top: 36, width: 220, height: 220 },
          style: { fontFamily: jostFont, fontSize: 10 },
          axes: {
            x: { kind: "value", min: 0, max: 10, crossesAt: 5 },
            y: { kind: "value", min: -10, max: 10, crosses: "autoZero" },
          },
          groups: [
            {
              kind: "scatter",
              axes: { x: "x", y: "y" },
              data: { kind: "xy" },
              line: "straight",
              markers: true,
              series: [
                {
                  name: "XY",
                  color: "#6366F1",
                  x: [1, 2, 4, 8],
                  y: [2, 4, 6, 9],
                  trendlines: [
                    { type: "exponential", line: { color: "#EC4899" } },
                    { type: "logarithmic" },
                    { type: "power" },
                    { type: "movingAverage", period: 2 },
                  ],
                  errorBars: {
                    x: {
                      type: "custom",
                      include: "both",
                      plus: [0.5, 0.5, 0.5, 0.5],
                      minus: [0.2, 0.2, 0.2, 0.2],
                    },
                    y: { type: "standardError", include: "both" },
                  },
                },
              ],
            },
          ],
        }),
        chart({
          id: "pie-rich",
          x: 324,
          y: 16,
          width: 300,
          height: 300,
          style: { fontFamily: jostFont, fontSize: 10 },
          groups: [
            {
              kind: "pie",
              firstSliceAngle: 45,
              explosion: 12,
              data: {
                kind: "shared-categories",
                categories: ["Alpha", "Beta", "Gamma"],
              },
              dataLabels: {
                value: true,
                percent: true,
                separator: " | ",
                position: "outsideEnd",
                leaderLines: { color: "#1E293B", width: 1.5, dash: "dot" },
              },
              series: [
                {
                  name: "Share",
                  values: [50, 30, 20],
                  points: {
                    0: {
                      color: "#2563EB",
                      explosion: 18,
                      labelColor: "#1E3A8A",
                    },
                  },
                },
              ],
            },
          ],
        }),
      ),
    );

    const initialEmit = emit(resolveDocument(initialDoc), {
      embedFonts: "subset",
    });
    const inputCopy = new Uint8Array(initialEmit.pptx);
    const imported = importPptx(inputCopy);
    inputCopy.fill(0);

    assert.deepEqual(imported.report.diagnostics, []);

    // Chart 1 (cartesian combo): verify presentation options, labels, trendlines, error bars, and crossing survived import
    const { node: colLineNode, props: colLineProps } = chartAt(
      imported.document,
      0,
    );
    assert.equal(colLineProps.legend, "topRight");
    assert.equal(colLineProps.plotArea, undefined);
    assert.ok(typeof colLineProps.title === "object");
    assert.equal(colLineProps.title.text, "Rich Presentation");
    assert.equal(colLineProps.title.style?.color, "#1D4ED8");
    assert.equal(colLineProps.title.style?.fontWeight, "bold");

    const valAxis = colLineProps.axes?.["value"];
    assert.ok(valAxis?.kind === "value");
    assert.equal(valAxis.reversed, true);
    assert.equal(valAxis.crosses, "min");
    assert.equal(valAxis.numberFormat, "$#,##0");
    assert.equal(valAxis.minorUnit, 5);

    const colGroup = colLineProps.groups[0]!;
    assert.ok(colGroup.kind === "column");
    assert.equal(colGroup.gapWidth, 80);
    assert.equal(colGroup.overlap, -15);
    assert.ok(
      typeof colGroup.dataLabels === "object" &&
        colGroup.dataLabels.style?.textDecoration === "underline",
    );
    const primarySeries = colGroup.series[0]!;
    assert.equal(primarySeries.points?.[1]?.color, "#DC2626");
    assert.equal(primarySeries.points?.[1]?.labelColor, "#FFFFFF");
    assert.equal(primarySeries.points?.[3]?.dataLabels, false);
    assert.equal(primarySeries.trendlines?.length, 1);
    assert.equal(primarySeries.trendlines?.[0]?.type, "polynomial");
    assert.equal(primarySeries.errorBars?.type, "fixed");
    assert.equal(primarySeries.errorBars?.endCap, false);
    assert.equal(colGroup.series[1]?.dataLabels, false);

    // Edit one point override on Primary series and verify sibling series/group stay intact
    colLineNode.set({
      groups: [
        {
          ...colGroup,
          series: [
            {
              ...primarySeries,
              points: {
                ...primarySeries.points,
                1: {
                  ...primarySeries.points?.[1],
                  color: "#16A34A",
                },
              },
            },
            colGroup.series[1],
          ],
        },
        colLineProps.groups[1]!,
      ],
    });

    // Chart 2 (scatter): verify manualLayout, interior crossesAt & autoZero, trendlines, and XY error bars
    const { props: scatterProps } = chartAt(imported.document, 1);
    assert.ok(scatterProps.plotArea !== undefined);
    const scatterGroup = scatterProps.groups[0]!;
    assert.ok(scatterGroup.kind === "scatter");
    assert.equal(scatterGroup.series[0]?.trendlines?.length, 4);
    assert.equal(scatterGroup.series[0]?.errorBars?.x?.type, "custom");
    assert.equal(scatterGroup.series[0]?.errorBars?.y?.type, "standardError");

    const resolution = resolveDocument(imported.document);
    const reemitted = emit(resolution, { embedFonts: "subset" });
    const painted = paint(resolution);
    for (const page of painted.pages) {
      assert.deepEqual(page.diagnostics, []);
      assert.ok(page.png instanceof Uint8Array);
    }

    const validation = await validatePptx(reemitted.pptx);
    assert.deepEqual(validation.errors, []);

    const chart1Xml = strFromU8(
      unzipSync(reemitted.pptx)["ppt/charts/chart1.xml"]!,
    );
    assert.ok(chart1Xml.includes("<c:layout/>"));
    assert.ok(!chart1Xml.includes("<c:manualLayout>"));

    const fileColLine = validation.slides[0]?.objects[0]?.chart;
    assert.ok(fileColLine);
    assert.equal(fileColLine.series[0]?.trendlines?.length, 1);
    assert.equal(fileColLine.series[0]?.trendlines?.[0]?.type, "polynomial");
    assert.equal(fileColLine.series[0]?.errorBars?.[0]?.type, "fixed");
    assert.equal(fileColLine.series[0]?.errorBars?.[0]?.endCap, false);
    assert.equal(fileColLine.series[1]?.trendlines, undefined);
    assert.equal(fileColLine.series[1]?.errorBars, undefined);
    assert.equal(
      fileColLine.seriesDataLabels?.find((entry) => entry.seriesIndex === 1)
        ?.dataLabels,
      false,
    );

    const fileScatter = validation.slides[1]?.objects[0]?.chart;
    assert.ok(fileScatter);
    assert.equal(fileScatter.series[0]?.trendlines?.length, 4);
    assert.equal(fileScatter.series[0]?.errorBars?.length, 2);
  });

  it("reports unconsumed-xml when plotArea, legend, or series spPr contains unsupported visual fills instead of swallowing them", () => {
    const pptx = buildChartDeck({
      id: "unconsumed-sppr",
      x: 20,
      y: 20,
      width: 600,
      height: 320,
      legend: "bottom",
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [{ name: "S", values: [1, 2] }],
        },
      ],
    });
    const entries = unzipSync(pptx);
    const baseChartXml = strFromU8(entries["ppt/charts/chart1.xml"]!);

    const gradFill =
      '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>';

    const mutatedPlotArea = baseChartXml.replace(
      "<c:plotArea>",
      `<c:plotArea><c:spPr>${gradFill}</c:spPr>`,
    );
    const importedPlotArea = importPptx(
      zipSync({
        ...entries,
        "ppt/charts/chart1.xml": strToU8(mutatedPlotArea),
      }),
    );
    assert.ok(
      importedPlotArea.report.diagnostics.some(
        (d) =>
          d.code === "unconsumed-xml" &&
          d.location.xpath?.includes("plotArea") &&
          d.localName === "gradFill",
      ),
    );

    const mutatedLegend = baseChartXml.replace(
      "</c:legend>",
      `<c:spPr>${gradFill}</c:spPr></c:legend>`,
    );
    const importedLegend = importPptx(
      zipSync({ ...entries, "ppt/charts/chart1.xml": strToU8(mutatedLegend) }),
    );
    assert.ok(
      importedLegend.report.diagnostics.some(
        (d) =>
          d.code === "unconsumed-xml" &&
          d.location.xpath?.includes("legend") &&
          (d.localName === "spPr" || d.localName === "gradFill"),
      ),
    );

    const mutatedSeries = baseChartXml.replace(
      "</c:ser>",
      `<c:spPr>${gradFill}</c:spPr></c:ser>`,
    );
    const importedSeries = importPptx(
      zipSync({ ...entries, "ppt/charts/chart1.xml": strToU8(mutatedSeries) }),
    );
    assert.ok(
      importedSeries.report.diagnostics.some(
        (d) =>
          d.code === "unconsumed-xml" &&
          d.location.xpath?.includes("ser") &&
          d.localName === "gradFill",
      ),
    );
  });
});
