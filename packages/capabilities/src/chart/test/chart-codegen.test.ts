import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { chartCapability } from "@pptxdk/capabilities/chart";
import { chart, deck, font, slide, text } from "@pptxdk/capabilities/document";
import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { createMorphScope } from "@pptxdk/core/capability";
import { CommonContainer } from "@pptxdk/core/document";
import { createImporter } from "@pptxdk/import";
import { describe, it } from "vitest";

describe("chart capability codegen", () => {
  it("generates vanilla and react source preserving groups, axes, three-state data ([] vs null vs 0), trendlines, errorBars, styles, and placeholders", () => {
    const morphScope = createMorphScope();
    const chartFont = font("ChartSans");

    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "charts-overview" }).append(
        chart({
          id: "combo-chart",
          name: "Combo Overview",
          placeholder: "chart",
          nativePlaceholder: { type: "chart", idx: 2 },
          link: { url: "https://example.com/report" },
          title: {
            text: "Revenue & Margin",
            style: { fontSize: 16, fontWeight: "bold", color: "#0F172A" },
          },
          displayBlanksAs: "gap",
          legend: "bottom",
          plotArea: { left: 48, top: 44, width: 420, height: 220 },
          style: {
            position: "absolute",
            left: 40,
            top: 48,
            width: 520,
            height: 320,
            fontFamily: chartFont,
            fontSize: 11,
            color: { slot: "accent1", lumMod: 0.6, lumOff: 0.4 },
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
          id: "empty-named-series",
          style: {
            position: "absolute",
            left: 580,
            top: 48,
            width: 340,
            height: 320,
          },
          groups: [
            {
              kind: "doughnut",
              data: { kind: "independent-categories" },
              series: [
                { name: "EmptyRing", categories: [], values: [] },
                { name: "FilledRing", categories: ["A"], values: [0] },
              ],
            },
          ],
        }),
        chart({
          id: "empty-scatter",
          morphKey: morphScope.handle("sales-chart"),
          style: {
            position: "absolute",
            left: 40,
            top: 380,
            width: 300,
            height: 140,
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

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry];
    assert.equal(typeof vanillaSource, "string");
    if (typeof vanillaSource !== "string") return;

    assert.match(vanillaSource, /chart\(\s*\{/u);
    assert.match(vanillaSource, /placeholder: "chart"/u);
    assert.match(
      vanillaSource,
      /nativePlaceholder: \{ type: "chart", idx: 2 \}/u,
    );
    assert.match(vanillaSource, /morphScope\.handle\("sales-chart"\)/u);
    assert.match(vanillaSource, /lighter\("accent1", 0\.4\)/u);
    assert.match(vanillaSource, /displayBlanksAs: "gap"/u);
    assert.match(vanillaSource, /legend: "bottom"/u);
    assert.match(vanillaSource, /values: \[0, null, 75\]/u);
    assert.match(vanillaSource, /plus: \[5, null, 8\]/u);
    assert.match(vanillaSource, /minus: \[2, 0, 4\]/u);
    assert.match(
      vanillaSource,
      /\{ name: "EmptyRing", categories: \[\], values: \[\] \}/u,
    );
    assert.match(
      vanillaSource,
      /\{ name: "FilledRing", categories: \["A"\], values: \[0\] \}/u,
    );
    assert.match(vanillaSource, /\{ name: "EmptyXY", x: \[\], y: \[\] \}/u);
    assert.match(vanillaSource, /x: \[0, 2\]/u);
    assert.match(vanillaSource, /y: \[null, 5\]/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry];
    assert.equal(typeof reactSource, "string");
    if (typeof reactSource !== "string") return;

    assert.match(reactSource, /<Chart\b/u);
    assert.match(reactSource, /id="combo-chart"/u);
    assert.match(reactSource, /placeholder="chart"/u);
    assert.match(reactSource, /values: \[0, null, 75\]/u);
    assert.match(
      reactSource,
      /\{ name: "EmptyRing", categories: \[\], values: \[\] \}/u,
    );
    assert.match(reactSource, /\{ name: "EmptyXY", x: \[\], y: \[\] \}/u);
  });

  it("throws CodegenError with exact paths on children, unknown props, invalid chart structure, non-finite values, and trimmed capability", () => {
    const generate = createCodegen(documentCapabilities);

    const withChildren = new CommonContainer("chart", "chart", {
      id: "bad-children",
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 300,
        height: 200,
      },
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A"] },
          series: [{ name: "S", values: [1] }],
        },
      ],
    }).append(text("unexpected"));
    assert.throws(
      () =>
        generate(deck().append(slide({ id: "s1" }).append(withChildren)), {
          target: "vanilla",
        }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/chart[bad-children]",
    );

    const unknownPropChart = (
      chart as unknown as (
        props: Record<string, unknown>,
      ) => ReturnType<typeof chart>
    )({
      id: "unknown-prop",
      style: {
        position: "absolute",
        left: 10,
        top: 10,
        width: 300,
        height: 200,
      },
      bogus: true,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A"] },
          series: [{ name: "S", values: [1] }],
        },
      ],
    });
    assert.throws(
      () =>
        generate(deck().append(slide({ id: "s1" }).append(unknownPropChart)), {
          target: "react",
        }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/chart[unknown-prop]@bogus",
    );

    const nanDoc = deck().append(
      slide({ id: "s1" }).append(
        chart({
          id: "nan-chart",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 300,
            height: 200,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1, Number.NaN] }],
            },
          ],
        }),
      ),
    );
    assert.throws(
      () => generate(nanDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/chart[nan-chart]@groups[0].series[0].values[1]",
    );

    const mismatchDoc = deck().append(
      slide({ id: "s1" }).append(
        chart({
          id: "mismatch-chart",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 300,
            height: 200,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A", "B"] },
              series: [{ name: "S", values: [1] }],
            },
          ],
        }),
      ),
    );
    assert.throws(
      () => generate(mismatchDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path.startsWith(
          "deck/slide[s1]/chart[mismatch-chart]@groups[0]",
        ),
    );

    const trimmedGenerate = createCodegen(
      documentCapabilities.filter((mod) => mod !== chartCapability),
    );
    assert.throws(
      () =>
        trimmedGenerate(
          deck().append(
            slide({ id: "s1" }).append(
              chart({
                id: "trimmed",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 300,
                  height: 200,
                },
                groups: [
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: ["A"] },
                    series: [{ name: "S", values: [1] }],
                  },
                ],
              }),
            ),
          ),
          { target: "vanilla" },
        ),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-declaration" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/chart[trimmed]",
    );
  });

  it("imports all 6 native-23-absent PowerPoint and WPS fixtures and emits both targets preserving empty, missing, zero, and rebound data", () => {
    const importPptx = createImporter(documentCapabilities);
    const generate = createCodegen(documentCapabilities);

    for (const file of [
      "23-absent-pp-untouched.pptx",
      "23-absent-pp-filled-cells.pptx",
      "23-absent-wps-untouched.pptx",
      "23-absent-wps-filled-cells.pptx",
    ] as const) {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`./fixtures/native-23-absent/${file}`, import.meta.url),
        ),
      );
      const imported = importPptx(bytes);
      bytes.fill(0);

      for (const target of ["vanilla", "react"] as const) {
        const result = generate(imported.document, { target });
        const source = result.files[result.entry];
        assert.equal(typeof source, "string");
        if (typeof source !== "string") continue;
        assert.match(source, /name: "North"/u, `${file} ${target}`);
        assert.match(source, /categories: \[\]/u, `${file} ${target}`);
        assert.match(source, /values: \[\]/u, `${file} ${target}`);
        assert.match(
          source,
          /values: \[null, null, null\]/u,
          `${file} ${target}`,
        );
        assert.match(source, /values: \[0, 0, 0\]/u, `${file} ${target}`);
        assert.match(
          source,
          /name: "Observations",\s*x: \[\],\s*y: \[\]/u,
          `${file} ${target}`,
        );
        assert.match(source, /x: \[null, null, null\]/u, `${file} ${target}`);
        assert.match(source, /x: \[0, 0, 0\]/u, `${file} ${target}`);
      }
    }

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

      for (const target of ["vanilla", "react"] as const) {
        const result = generate(imported.document, { target });
        const source = result.files[result.entry];
        assert.equal(typeof source, "string");
        if (typeof source !== "string") continue;
        assert.match(source, /categories: \["A", "B"\]/u, `${file} ${target}`);
        assert.match(source, /values: \[3, 7\]/u, `${file} ${target}`);
        if (file === "23-absent-pp-filled-range.pptx") {
          assert.match(
            source,
            /categories: \["One", "Two"\]/u,
            `${file} ${target}`,
          );
          assert.match(source, /values: \[2, 5\]/u, `${file} ${target}`);
          assert.match(
            source,
            /name: "Y",\s*x: \[2, 5\],\s*y: \[4, 1\]/u,
            `${file} ${target}`,
          );
        } else {
          assert.match(
            source,
            /name: "Observations",\s*x: \[\],\s*y: \[\]/u,
            `${file} ${target}`,
          );
        }
      }
    }
  });
});
