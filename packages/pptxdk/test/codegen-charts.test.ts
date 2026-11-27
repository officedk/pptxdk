import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  build,
  codegen,
  importPptx,
  type CodegenResult,
} from "@officedk/pptxdk";
import {
  Chart,
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  defineSlideLayout,
  expandDocument,
  lighter,
} from "@officedk/pptxdk/react";
import {
  chart,
  deck,
  font,
  slide,
  slideLayout,
  slideMaster,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  previewGeneratedSourceRelocated,
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
} from "./helpers/codegen-acceptance.ts";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [
  { weight: 400, style: "normal", data: jostBookBytes },
  { weight: 700, style: "normal", data: jostBoldBytes },
]);

const chartPlaceholderLayout = defineSlideLayout({
  id: "chart-layout",
  name: "Chart Layout",
  placeholders: {
    title: {
      type: "title",
      style: {
        position: "absolute",
        left: 32,
        top: 16,
        width: 896,
        height: 36,
      },
    },
    chart: {
      type: "chart",
      explicitType: true,
      idx: 2,
      style: {
        position: "absolute",
        left: 32,
        top: 60,
        width: 520,
        height: 300,
      },
    },
  },
});

function countPaintedPixelsInBox(
  pngBytes: Uint8Array,
  region: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
): number {
  const png = PNG.sync.read(Buffer.from(pngBytes));
  const minX = Math.max(0, Math.floor(region.x));
  const minY = Math.max(0, Math.floor(region.y));
  const maxX = Math.min(png.width - 1, Math.ceil(region.x + region.width));
  const maxY = Math.min(png.height - 1, Math.ceil(region.y + region.height));
  let painted = 0;
  for (let y = minY; y <= maxY; y += 1) {
    for (let x = minX; x <= maxX; x += 1) {
      const idx = (y * png.width + x) * 4;
      const r = png.data[idx]!;
      const g = png.data[idx + 1]!;
      const b = png.data[idx + 2]!;
      const a = png.data[idx + 3]!;
      if (a > 0 && (r < 245 || g < 245 || b < 245)) {
        painted += 1;
      }
    }
  }
  return painted;
}

function createDirectChartsVanillaDeck() {
  return deck({ width: 960, height: 540 }).append(
    slideMaster({ id: "charts-master", name: "Charts Master" }).append(
      slideLayout(chartPlaceholderLayout),
    ),
    slide({ id: "charts-overview", layout: chartPlaceholderLayout }).append(
      chart({
        id: "combo-chart",
        name: "Combo Overview",
        placeholder: "chart",
        nativePlaceholder: { type: "chart", idx: 2 },
        link: { url: "https://example.com/report" },
        title: {
          text: "Revenue & Margin",
          style: {
            fontFamily: jostFont,
            fontSize: 16,
            fontWeight: "bold",
            color: "#0F172A",
          },
        },
        displayBlanksAs: "gap",
        legend: "bottom",
        plotArea: { left: 48, top: 44, width: 420, height: 210 },
        style: {
          position: "absolute",
          left: 32,
          top: 60,
          width: 520,
          height: 300,
          fontFamily: jostFont,
          fontSize: 11,
          color: lighter("accent1", 0.4),
        },
        axes: {
          period: { kind: "category", title: "Period" },
          amount: { kind: "value", position: "left", min: 0, max: 100 },
          margin: {
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
              categories: ["Q1", "Q2", "Q3"],
            },
            series: [
              {
                name: "ZeroAndNull",
                values: [0, null, 75],
                points: {
                  0: { color: "#1D4ED8", dataLabels: { value: true } },
                },
                trendlines: [
                  {
                    type: "linear",
                    displayEquation: true,
                    displayRSquared: true,
                  },
                ],
                errorBars: {
                  type: "custom",
                  include: "both",
                  plus: [5, null, 8],
                  minus: [2, 0, 4],
                },
              },
            ],
          },
          {
            kind: "line",
            axes: { category: "period", value: "margin" },
            data: {
              kind: "shared-categories",
              categories: ["Q1", "Q2", "Q3"],
            },
            series: [
              {
                name: "Rate",
                smooth: true,
                values: [0.15, 0.25, 0.4],
              },
            ],
          },
        ],
      }),
      chart({
        id: "ring-chart",
        name: "Ring Breakdown",
        style: {
          position: "absolute",
          left: 580,
          top: 60,
          width: 340,
          height: 240,
          fontFamily: jostFont,
          fontSize: 11,
        },
        groups: [
          {
            kind: "doughnut",
            data: { kind: "independent-categories" },
            series: [
              { name: "EmptyRing", categories: [], values: [] },
              { name: "FilledRing", categories: ["A", "B"], values: [0, 40] },
            ],
          },
        ],
      }),
      chart({
        id: "scatter-chart",
        name: "Scatter Observations",
        style: {
          position: "absolute",
          left: 32,
          top: 376,
          width: 440,
          height: 148,
          fontFamily: jostFont,
          fontSize: 11,
        },
        groups: [
          {
            kind: "scatter",
            data: { kind: "xy" },
            series: [
              { name: "EmptyXY", x: [], y: [] },
              {
                name: "MeasuredXY",
                x: [0, 2],
                y: [null, 5],
                errorBars: {
                  x: { type: "fixed", value: 0.5 },
                  y: { type: "standardError", include: "both" },
                },
              },
            ],
          },
        ],
      }),
    ),
  );
}

async function createDirectChartsReactDeck() {
  return expandDocument(
    createElement(
      Deck,
      { width: 960, height: 540 },
      createElement(
        SlideMaster,
        { id: "charts-master", name: "Charts Master" },
        createElement(SlideLayout, { layout: chartPlaceholderLayout }),
      ),
      createElement(
        Slide,
        { id: "charts-overview", layout: chartPlaceholderLayout },
        createElement(Chart, {
          id: "combo-chart",
          name: "Combo Overview",
          placeholder: "chart",
          nativePlaceholder: { type: "chart", idx: 2 },
          link: { url: "https://example.com/report" },
          title: {
            text: "Revenue & Margin",
            style: {
              fontFamily: jostFont,
              fontSize: 16,
              fontWeight: "bold",
              color: "#0F172A",
            },
          },
          displayBlanksAs: "gap",
          legend: "bottom",
          plotArea: { left: 48, top: 44, width: 420, height: 210 },
          style: {
            position: "absolute",
            left: 32,
            top: 60,
            width: 520,
            height: 300,
            fontFamily: jostFont,
            fontSize: 11,
            color: lighter("accent1", 0.4),
          },
          axes: {
            period: { kind: "category", title: "Period" },
            amount: { kind: "value", position: "left", min: 0, max: 100 },
            margin: {
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
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                {
                  name: "ZeroAndNull",
                  values: [0, null, 75],
                  points: {
                    0: { color: "#1D4ED8", dataLabels: { value: true } },
                  },
                  trendlines: [
                    {
                      type: "linear",
                      displayEquation: true,
                      displayRSquared: true,
                    },
                  ],
                  errorBars: {
                    type: "custom",
                    include: "both",
                    plus: [5, null, 8],
                    minus: [2, 0, 4],
                  },
                },
              ],
            },
            {
              kind: "line",
              axes: { category: "period", value: "margin" },
              data: {
                kind: "shared-categories",
                categories: ["Q1", "Q2", "Q3"],
              },
              series: [
                {
                  name: "Rate",
                  smooth: true,
                  values: [0.15, 0.25, 0.4],
                },
              ],
            },
          ],
        }),
        createElement(Chart, {
          id: "ring-chart",
          name: "Ring Breakdown",
          style: {
            position: "absolute",
            left: 580,
            top: 60,
            width: 340,
            height: 240,
            fontFamily: jostFont,
            fontSize: 11,
          },
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [
                { name: "EmptyRing", categories: [], values: [] },
                { name: "FilledRing", categories: ["A", "B"], values: [0, 40] },
              ],
            },
          ],
        }),
        createElement(Chart, {
          id: "scatter-chart",
          name: "Scatter Observations",
          style: {
            position: "absolute",
            left: 32,
            top: 376,
            width: 440,
            height: 148,
            fontFamily: jostFont,
            fontSize: 11,
          },
          groups: [
            {
              kind: "scatter",
              data: { kind: "xy" },
              series: [
                { name: "EmptyXY", x: [], y: [] },
                {
                  name: "MeasuredXY",
                  x: [0, 2],
                  y: [null, 5],
                  errorBars: {
                    x: { type: "fixed", value: 0.5 },
                    y: { type: "standardError", include: "both" },
                  },
                },
              ],
            },
          ],
        }),
      ),
    ),
  );
}

function assertDualAxisComboChartXml(
  chartXml: string,
  label: string,
): {
  readonly barValueAxId: string;
  readonly lineValueAxId: string;
} {
  const barMatch = /<c:barChart>([\s\S]*?)<\/c:barChart>/u.exec(chartXml);
  const lineMatch = /<c:lineChart>([\s\S]*?)<\/c:lineChart>/u.exec(chartXml);
  assert.ok(barMatch, `${label} barChart element`);
  assert.ok(lineMatch, `${label} lineChart element`);

  const barAxIds = [...barMatch[1]!.matchAll(/<c:axId val="(\d+)"\/>/gu)].map(
    (m) => m[1]!,
  );
  const lineAxIds = [...lineMatch[1]!.matchAll(/<c:axId val="(\d+)"\/>/gu)].map(
    (m) => m[1]!,
  );
  assert.equal(barAxIds.length, 2, `${label} barChart axIds`);
  assert.equal(lineAxIds.length, 2, `${label} lineChart axIds`);
  assert.equal(
    barAxIds[0],
    lineAxIds[0],
    `${label} shared category axis id between bar and line groups`,
  );
  assert.notEqual(
    barAxIds[1],
    lineAxIds[1],
    `${label} distinct primary and secondary value axis ids`,
  );

  const valAxes = [
    ...chartXml.matchAll(/<c:valAx>([\s\S]*?)<\/c:valAx>/gu),
  ].map((m) => m[1]!);
  const barValAx = valAxes.find((ax) =>
    ax.includes(`<c:axId val="${barAxIds[1]}"/>`),
  );
  const lineValAx = valAxes.find((ax) =>
    ax.includes(`<c:axId val="${lineAxIds[1]}"/>`),
  );
  assert.ok(barValAx?.includes('<c:axPos val="l"/>'), `${label} left valAx`);
  assert.ok(lineValAx?.includes('<c:axPos val="r"/>'), `${label} right valAx`);

  return {
    barValueAxId: barAxIds[1]!,
    lineValueAxId: lineAxIds[1]!,
  };
}

function readWorkbookSheetXml(
  parts: Readonly<Record<string, Uint8Array>>,
  workbookPart: string,
): string {
  const workbookArchive = unzipSync(parts[workbookPart]!);
  const sheetBytes = workbookArchive["xl/worksheets/sheet1.xml"];
  assert.ok(sheetBytes, `Missing sheet1.xml in ${workbookPart}`);
  return strFromU8(sheetBytes);
}

describe("pptxdk native charts codegen acceptance", () => {
  it("generates typechecked, relocated-rebuildable, and previewable vanilla and react sources for direct and imported charts with dual axes, three-state data, trendlines, and errorBars", async () => {
    const vanillaInput = createDirectChartsVanillaDeck();
    const reactInput = await createDirectChartsReactDeck();

    const originalBuilt = await build(vanillaInput, { embedFonts: "full" });
    const originalBytes = new Uint8Array(originalBuilt.pptx);
    const imported = importPptx(originalBytes);
    originalBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const sources = [
      { label: "direct-vanilla", doc: vanillaInput },
      { label: "direct-react", doc: reactInput },
      { label: "imported-pptx", doc: imported.document },
    ] as const;

    const generatedCases = sources.flatMap(({ label, doc }) => {
      const vanillaGen = codegen(doc, { target: "vanilla" });
      const reactGen = codegen(doc, { target: "react" });
      return [
        { label: `${label}:vanilla`, result: vanillaGen },
        { label: `${label}:react`, result: reactGen },
      ];
    });

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(
        file.errors,
        [],
        `${item.label} validation errors: ${JSON.stringify(file.errors)}`,
      );

      const parts = unzipSync(built.pptx);
      const objects = file.slides[0]!.objects;

      const comboObj = objects.find((obj) => obj.name === "Combo Overview");
      assert.ok(comboObj?.chart, `${item.label} Combo Overview`);
      assert.deepEqual(comboObj.placeholder, { type: "chart", idx: 2 });
      assert.deepEqual(
        comboObj.chart.series.map((s) => ({
          name: s.name,
          categories: s.categories,
          values: s.values,
        })),
        [
          {
            name: "ZeroAndNull",
            categories: ["Q1", "Q2", "Q3"],
            values: [0, null, 75],
          },
          {
            name: "Rate",
            categories: ["Q1", "Q2", "Q3"],
            values: [0.15, 0.25, 0.4],
          },
        ],
        `${item.label} combo chart series`,
      );
      assert.deepEqual(
        comboObj.chart.embeddedWorkbook?.series.map((s) => ({
          name: s.name,
          categories: s.categories,
          values: s.values,
        })),
        [
          {
            name: "ZeroAndNull",
            categories: ["Q1", "Q2", "Q3"],
            values: [0, null, 75],
          },
          {
            name: "Rate",
            categories: ["Q1", "Q2", "Q3"],
            values: [0.15, 0.25, 0.4],
          },
        ],
        `${item.label} combo embedded workbook series`,
      );

      const comboChartXml = strFromU8(parts["ppt/charts/chart1.xml"]!);
      assertDualAxisComboChartXml(comboChartXml, item.label);
      assert.ok(comboChartXml.includes('<c:dispBlanksAs val="gap"/>'));
      assert.ok(comboChartXml.includes("<c:trendline>"));
      assert.ok(comboChartXml.includes("<c:errBars>"));

      const comboWorkbookSheet = readWorkbookSheetXml(
        parts,
        comboObj.chart.embeddedWorkbook.part,
      );
      assert.ok(comboWorkbookSheet.includes("<v>0</v>"));
      assert.ok(comboWorkbookSheet.includes("<v>75</v>"));
      assert.ok(comboWorkbookSheet.includes("<v>0.15</v>"));

      const ringObj = objects.find((obj) => obj.name === "Ring Breakdown");
      assert.ok(ringObj?.chart, `${item.label} Ring Breakdown`);
      assert.deepEqual(
        ringObj.chart.series.map((s) => ({
          name: s.name,
          categories: s.categories,
          values: s.values,
        })),
        [
          { name: "EmptyRing", categories: [], values: [] },
          { name: "FilledRing", categories: ["A", "B"], values: [0, 40] },
        ],
      );
      assert.deepEqual(
        ringObj.chart.embeddedWorkbook?.series.map((s) => ({
          name: s.name,
          categories: s.categories,
          values: s.values,
        })),
        [
          { name: "EmptyRing", categories: [], values: [] },
          { name: "FilledRing", categories: ["A", "B"], values: [0, 40] },
        ],
      );

      const scatterObj = objects.find(
        (obj) => obj.name === "Scatter Observations",
      );
      assert.ok(scatterObj?.chart, `${item.label} Scatter Observations`);
      assert.deepEqual(
        scatterObj.chart.series.map((s) => ({
          name: s.name,
          x: s.x,
          y: s.y,
        })),
        [
          { name: "EmptyXY", x: [], y: [] },
          { name: "MeasuredXY", x: [0, 2], y: [null, 5] },
        ],
      );
      assert.deepEqual(
        scatterObj.chart.embeddedWorkbook?.series.map((s) => ({
          name: s.name,
          x: s.x,
          y: s.y,
        })),
        [
          { name: "EmptyXY", x: [], y: [] },
          { name: "MeasuredXY", x: [0, 2], y: [null, 5] },
        ],
      );

      const painted = await previewGeneratedSourceRelocated(item.result);
      const page = painted.pages[0]!;
      assert.ok(page.png instanceof Uint8Array, `${item.label} png`);
      assert.ok(typeof page.svg === "string", `${item.label} svg`);
      assert.deepEqual(page.errors, [], `${item.label} preview errors`);

      for (const chartName of [
        "Combo Overview",
        "Ring Breakdown",
        "Scatter Observations",
      ]) {
        const bounds = page.nodes.find((n) => n.name === chartName)?.bounds;
        assert.ok(bounds !== undefined, `${item.label} ${chartName} bounds`);
        assert.ok(
          countPaintedPixelsInBox(page.png, bounds) > 100,
          `${item.label} ${chartName} painted pixels`,
        );
      }
    }
  }, 90_000);

  it("edits one series data points in an independent source copy and updates its cache and workbook while preserving sibling series and dual-axis bindings", async () => {
    const vanillaInput = createDirectChartsVanillaDeck();
    const builtOriginal = await build(vanillaInput, { embedFonts: "full" });
    const originalBytes = new Uint8Array(builtOriginal.pptx);
    const imported = importPptx(originalBytes);
    originalBytes.fill(0);
    assert.deepEqual(imported.report.diagnostics, []);

    const sourceCases = [
      {
        label: "direct:vanilla",
        result: codegen(vanillaInput, { target: "vanilla" }),
      },
      {
        label: "direct:react",
        result: codegen(vanillaInput, { target: "react" }),
      },
      {
        label: "imported:vanilla",
        result: codegen(imported.document, { target: "vanilla" }),
      },
      {
        label: "imported:react",
        result: codegen(imported.document, { target: "react" }),
      },
    ] as const;

    const initialPreview = await previewGeneratedSourceRelocated(
      sourceCases[0].result,
    );
    assert.deepEqual(initialPreview.pages[0]?.errors, []);
    const initialComboNode = initialPreview.pages[0]?.nodes.find(
      (n) => n.name === "Combo Overview",
    );
    assert.ok(initialComboNode?.bounds);
    const initialComboPixels = countPaintedPixelsInBox(
      initialPreview.pages[0].png!,
      initialComboNode.bounds,
    );

    const editedCases = sourceCases.map(({ label, result }) => {
      const source = result.files[result.entry];
      assert.equal(typeof source, "string");
      if (typeof source !== "string") {
        throw new Error(`Expected string source for ${label}`);
      }
      assert.ok(
        source.includes("values: [0, null, 75]"),
        `${label} contains original ZeroAndNull values`,
      );
      const updated = source.replace(
        "values: [0, null, 75]",
        "values: [12, 48, 90]",
      );
      assert.notEqual(updated, source, `${label} edit applied`);
      return {
        label: `${label}:edited`,
        result: {
          ...result,
          files: {
            ...result.files,
            [result.entry]: updated,
          },
        },
      };
    });

    await verifyGeneratedSourcesTypecheck(editedCases);

    for (const item of editedCases) {
      const tempRoot = await mkdtemp(
        join(tmpdir(), "pptxdk-codegen-cht-edit-"),
      );
      try {
        // Confirm edits live in an independent directory without touching earlier outputs.
        await writeCodegenFiles(tempRoot, item.result.files);
        const diskSource = await readFile(
          join(tempRoot, item.result.entry),
          "utf8",
        );
        assert.ok(diskSource.includes("values: [12, 48, 90]"));
      } finally {
        await rm(tempRoot, { recursive: true, force: true });
      }

      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(file.errors, [], `${item.label} validation errors`);

      const comboObj = file.slides[0]!.objects.find(
        (obj) => obj.name === "Combo Overview",
      );
      assert.ok(comboObj?.chart, `${item.label} Combo Overview`);
      assert.deepEqual(
        comboObj.chart.series.map((s) => ({
          name: s.name,
          values: s.values,
        })),
        [
          { name: "ZeroAndNull", values: [12, 48, 90] },
          { name: "Rate", values: [0.15, 0.25, 0.4] },
        ],
        `${item.label} updated chart cache and preserved sibling series`,
      );
      assert.deepEqual(
        comboObj.chart.embeddedWorkbook?.series.map((s) => ({
          name: s.name,
          values: s.values,
        })),
        [
          { name: "ZeroAndNull", values: [12, 48, 90] },
          { name: "Rate", values: [0.15, 0.25, 0.4] },
        ],
        `${item.label} updated embedded workbook and preserved sibling series`,
      );

      const parts = unzipSync(built.pptx);
      const comboXml = strFromU8(parts["ppt/charts/chart1.xml"]!);
      assertDualAxisComboChartXml(comboXml, item.label);
      const comboSheet = readWorkbookSheetXml(
        parts,
        comboObj.chart.embeddedWorkbook.part,
      );
      assert.ok(comboSheet.includes("<v>12</v>"), `${item.label} workbook 12`);
      assert.ok(comboSheet.includes("<v>48</v>"), `${item.label} workbook 48`);
      assert.ok(comboSheet.includes("<v>90</v>"), `${item.label} workbook 90`);
      assert.ok(
        comboSheet.includes("<v>0.15</v>"),
        `${item.label} workbook sibling 0.15`,
      );

      const painted = await previewGeneratedSourceRelocated(item.result);
      assert.deepEqual(
        painted.pages[0]?.errors,
        [],
        `${item.label} preview errors`,
      );
      const comboPreviewNode = painted.pages[0]?.nodes.find(
        (n) => n.name === "Combo Overview",
      );
      assert.ok(comboPreviewNode?.bounds);
      const editedComboPixels = countPaintedPixelsInBox(
        painted.pages[0].png!,
        comboPreviewNode.bounds,
      );
      assert.ok(
        editedComboPixels > initialComboPixels,
        `${item.label} adding Q1/Q2 bar values should paint more pixels in Combo Overview (${editedComboPixels} > ${initialComboPixels})`,
      );
    }
  }, 90_000);

  it("imports all 6 native-23-absent PowerPoint and WPS fixtures, preserves original diagnostics vs zero-error rebuilt archives, and retains empty vs rebound data", async () => {
    type Native23AbsentFixture = {
      readonly file: string;
      readonly sha256: string;
      readonly kind:
        | "untouched"
        | "filled-cells"
        | "pp-rebound"
        | "wps-rebound";
    };

    const fixtures: readonly Native23AbsentFixture[] = [
      {
        file: "23-absent-pp-untouched.pptx",
        sha256:
          "72912ea68f0b376146ee8e3d85cf402b154c98fdb5789320b76f979811d0407a",
        kind: "untouched",
      },
      {
        file: "23-absent-pp-filled-cells.pptx",
        sha256:
          "f6bb4bcdce2067aacdbd750a40fe1e81e279b730a7cd91f2131131797ba19865",
        kind: "filled-cells",
      },
      {
        file: "23-absent-pp-filled-range.pptx",
        sha256:
          "d37156b2bd8bd67c8bba12e3eb41adaa62f3bfde6185fa319b793a22db0a10e1",
        kind: "pp-rebound",
      },
      {
        file: "23-absent-wps-untouched.pptx",
        sha256:
          "fc80e861b7092fcfb3fb23ae6f799186eefdaead9bf83d0080552cbc16ae0696",
        kind: "untouched",
      },
      {
        file: "23-absent-wps-filled-cells.pptx",
        sha256:
          "ac8b373aff1ab9284fb88c74518d5c1442ee29e77a5d1214e595bb9f18e24954",
        kind: "filled-cells",
      },
      {
        file: "23-absent-wps-series-ui.pptx",
        sha256:
          "4787cd57af84b24f977a6f69e3433e144b7158521b3325bfe63b499a4af13001",
        kind: "wps-rebound",
      },
    ];

    const emptyTargets = [
      { slideIndex: 0, part: "ppt/charts/chart1.xml", name: "North" },
      { slideIndex: 4, part: "ppt/charts/chart5.xml", name: "Share" },
      { slideIndex: 6, part: "ppt/charts/chart7.xml", name: "Observations" },
    ] as const;

    const cases: {
      readonly fixture: Native23AbsentFixture;
      readonly target: "vanilla" | "react";
      readonly label: string;
      readonly result: CodegenResult;
    }[] = [];

    for (const fixture of fixtures) {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(
            `./fixtures/native-23-absent/${fixture.file}`,
            import.meta.url,
          ),
        ),
      );
      assert.equal(
        createHash("sha256").update(bytes).digest("hex"),
        fixture.sha256,
        `${fixture.file} SHA-256`,
      );

      if (fixture.kind === "filled-cells") {
        // Confirm the original fixture carries chart-workbook-mismatch before codegen/rebuild.
        const originalValidation = await validatePptx(bytes);
        const originalParts = unzipSync(bytes);
        for (const target of emptyTargets) {
          assert.ok(
            originalValidation.errors.some(
              (err) =>
                err.code === "chart-workbook-mismatch" &&
                err.part === target.part,
            ),
            `${fixture.file} original ${target.part} must report chart-workbook-mismatch`,
          );
          const workbookPart =
            originalValidation.slides[target.slideIndex]?.objects[0]?.chart
              ?.embeddedWorkbook?.part;
          assert.ok(workbookPart);
          const sheet = readWorkbookSheetXml(originalParts, workbookPart);
          assert.ok(
            sheet.includes('r="B2"') && sheet.includes('r="B3"'),
            `${fixture.file} original workbook has unbound B2/B3 cells`,
          );
        }
      }

      const imported = importPptx(bytes);
      bytes.fill(0);
      assert.deepEqual(
        imported.report.diagnostics.filter(
          (d) =>
            d.location.part.startsWith("ppt/charts/") ||
            /^ppt\/slides\/slide\d+\.xml$/u.test(d.location.part),
        ),
        [],
        `${fixture.file} import diagnostics`,
      );

      for (const target of ["vanilla", "react"] as const) {
        cases.push({
          fixture,
          target,
          label: `${fixture.file}:${target}`,
          result: codegen(imported.document, { target }),
        });
      }
    }

    await verifyGeneratedSourcesTypecheck(
      cases.map(({ label, result }) => ({ label, result })),
    );

    for (const item of cases) {
      const { built, file } = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(
        file.errors,
        [],
        `${item.label} rebuilt archive must have 0 validation errors, got: ${JSON.stringify(file.errors)}`,
      );
      const parts = unzipSync(built.pptx);

      // Slide 2, 3, 8, 9 verify null vs 0 fidelity across all 6 native samples.
      const colMissing = file.slides[1]?.objects[0]?.chart?.series[0];
      assert.deepEqual(colMissing?.categories, ["Q1", "Q2", "Q3"], item.label);
      assert.deepEqual(colMissing?.values, [null, null, null], item.label);

      const colZero = file.slides[2]?.objects[0]?.chart?.series[0];
      assert.deepEqual(colZero?.categories, ["Q1", "Q2", "Q3"], item.label);
      assert.deepEqual(colZero?.values, [0, 0, 0], item.label);

      const scatterMissing = file.slides[7]?.objects[0]?.chart?.series[0];
      assert.deepEqual(scatterMissing?.x, [null, null, null], item.label);
      assert.deepEqual(scatterMissing?.y, [null, null, null], item.label);

      const scatterZero = file.slides[8]?.objects[0]?.chart?.series[0];
      assert.deepEqual(scatterZero?.x, [0, 0, 0], item.label);
      assert.deepEqual(scatterZero?.y, [0, 0, 0], item.label);

      if (
        item.fixture.kind === "untouched" ||
        item.fixture.kind === "filled-cells"
      ) {
        for (const target of emptyTargets) {
          const chartData = file.slides[target.slideIndex]?.objects[0]?.chart;
          assert.ok(chartData, `${item.label} ${target.part}`);
          assert.equal(chartData.series.length, 1);
          assert.equal(chartData.series[0]?.name, target.name);
          if (target.slideIndex === 6) {
            assert.deepEqual(chartData.series[0]?.x, [], item.label);
            assert.deepEqual(chartData.series[0]?.y, [], item.label);
            assert.deepEqual(
              chartData.embeddedWorkbook?.series[0]?.x,
              [],
              item.label,
            );
            assert.deepEqual(
              chartData.embeddedWorkbook?.series[0]?.y,
              [],
              item.label,
            );
          } else {
            assert.deepEqual(chartData.series[0]?.categories, [], item.label);
            assert.deepEqual(chartData.series[0]?.values, [], item.label);
            assert.deepEqual(
              chartData.embeddedWorkbook?.series[0]?.categories,
              [],
              item.label,
            );
            assert.deepEqual(
              chartData.embeddedWorkbook?.series[0]?.values,
              [],
              item.label,
            );
          }
          const chartXml = strFromU8(parts[target.part]!);
          assert.ok(chartXml.includes(`<c:v>${target.name}</c:v>`), item.label);
          assert.ok(!chartXml.includes("<c:numCache>"), item.label);
          const rebuiltSheet = readWorkbookSheetXml(
            parts,
            chartData.embeddedWorkbook.part,
          );
          assert.ok(
            !rebuiltSheet.includes('r="B2"') &&
              !rebuiltSheet.includes('r="B3"'),
            `${item.label} ${target.part} rebuilt workbook must not retain unbound B2/B3 cells`,
          );
        }
      } else if (item.fixture.kind === "pp-rebound") {
        const s1Chart = file.slides[0]?.objects[0]?.chart;
        assert.equal(s1Chart?.series[0]?.name, "North");
        assert.deepEqual(s1Chart?.series[0]?.categories, ["A", "B"]);
        assert.deepEqual(s1Chart?.series[0]?.values, [3, 7]);
        assert.deepEqual(s1Chart?.embeddedWorkbook?.series[0]?.categories, [
          "A",
          "B",
        ]);
        assert.deepEqual(s1Chart?.embeddedWorkbook?.series[0]?.values, [3, 7]);

        const s5Chart = file.slides[4]?.objects[0]?.chart;
        assert.equal(s5Chart?.series[0]?.name, "Share");
        assert.deepEqual(s5Chart?.series[0]?.categories, ["One", "Two"]);
        assert.deepEqual(s5Chart?.series[0]?.values, [2, 5]);
        assert.deepEqual(s5Chart?.embeddedWorkbook?.series[0]?.categories, [
          "One",
          "Two",
        ]);
        assert.deepEqual(s5Chart?.embeddedWorkbook?.series[0]?.values, [2, 5]);

        const s7Chart = file.slides[6]?.objects[0]?.chart;
        assert.equal(s7Chart?.series[0]?.name, "Y");
        assert.deepEqual(s7Chart?.series[0]?.x, [2, 5]);
        assert.deepEqual(s7Chart?.series[0]?.y, [4, 1]);
        assert.deepEqual(s7Chart?.embeddedWorkbook?.series[0]?.x, [2, 5]);
        assert.deepEqual(s7Chart?.embeddedWorkbook?.series[0]?.y, [4, 1]);
      } else {
        // 23-absent-wps-series-ui.pptx
        const s1Chart = file.slides[0]?.objects[0]?.chart;
        assert.equal(s1Chart?.series[0]?.name, "North");
        assert.deepEqual(s1Chart?.series[0]?.categories, ["A", "B"]);
        assert.deepEqual(s1Chart?.series[0]?.values, [3, 7]);
        assert.deepEqual(s1Chart?.embeddedWorkbook?.series[0]?.categories, [
          "A",
          "B",
        ]);
        assert.deepEqual(s1Chart?.embeddedWorkbook?.series[0]?.values, [3, 7]);

        const s5Chart = file.slides[4]?.objects[0]?.chart;
        assert.equal(s5Chart?.series[0]?.name, "Share");
        assert.deepEqual(s5Chart?.series[0]?.categories, []);
        assert.deepEqual(s5Chart?.series[0]?.values, []);

        const s7Chart = file.slides[6]?.objects[0]?.chart;
        assert.equal(s7Chart?.series[0]?.name, "Observations");
        assert.deepEqual(s7Chart?.series[0]?.x, []);
        assert.deepEqual(s7Chart?.series[0]?.y, []);
      }
    }
  }, 120_000);
});
