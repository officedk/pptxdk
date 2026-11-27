import assert from "node:assert/strict";

import {
  after,
  Chart,
  fade,
  MotionScope,
  motionSchedule,
  Shape,
  useMotionScope,
  type ResolvedMotionEffect,
} from "@pptxdk/capabilities";
import {
  chart as documentChart,
  deck,
  slide,
} from "@pptxdk/capabilities/document";
import { Deck, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

import { build, build as buildDocument } from "../../../test/test-helpers.ts";
import type { ResolvedMotion } from "../motion-api.ts";

const orders = [
  "series",
  "category",
  "seriesElements",
  "categoryElements",
] as const;

const sharedCategories = ["Direct", "Partner", "Direct"] as const;

function rewriteSlideXml(
  pptx: Uint8Array,
  rewrite: (xml: string) => string,
): Uint8Array {
  const path = "ppt/slides/slide1.xml";
  const parts = unzipSync(pptx);
  const xml = strFromU8(parts[path]!);
  const next = rewrite(xml);
  assert.notEqual(next, xml);
  return zipSync({ ...parts, [path]: strToU8(next) });
}

const sharedRingGroups = [
  {
    kind: "doughnut" as const,
    data: { kind: "shared-categories" as const, categories: sharedCategories },
    holeSize: 55,
    series: [
      { name: "Actual", values: [12, null, 18] },
      { name: "Plan", values: [10, 15, 20] },
      { name: "Forecast", values: [11, 16, 19] },
    ],
  },
];

const independentRingGroups = [
  {
    kind: "doughnut" as const,
    data: { kind: "independent-categories" as const },
    holeSize: 55,
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
  },
];

const sharedCombo = {
  axes: {
    period: { kind: "category" as const },
    amount: {
      kind: "value" as const,
      position: "left" as const,
      min: 0,
      max: 10,
    },
    share: {
      kind: "value" as const,
      position: "right" as const,
      min: 0,
      max: 100,
      numberFormat: "0%",
    },
  },
  groups: [
    {
      kind: "column" as const,
      axes: { category: "period", value: "amount" },
      data: {
        kind: "shared-categories" as const,
        categories: ["A", "B", "C"],
      },
      series: [{ name: "Amount", values: [2, 6, 4] }],
    },
    {
      kind: "line" as const,
      axes: { category: "period", value: "share" },
      data: {
        kind: "shared-categories" as const,
        categories: ["A", "B", "C"],
      },
      series: [{ name: "Share", values: [30, 70, 45] }],
    },
  ],
};

const mixedDomainCombo = {
  axes: {
    period: { kind: "category" as const },
    categoryValue: { kind: "value" as const, min: 0, max: 10 },
    scatterX: { kind: "value" as const, min: 0, max: 5 },
    scatterY: { kind: "value" as const, min: 0, max: 10 },
  },
  groups: [
    {
      kind: "column" as const,
      axes: { category: "period", value: "categoryValue" },
      data: {
        kind: "shared-categories" as const,
        categories: ["A", "B"],
      },
      series: [{ name: "Bars", values: [2, 8] }],
    },
    {
      kind: "scatter" as const,
      axes: { x: "scatterX", y: "scatterY" },
      data: { kind: "xy" as const },
      series: [
        { name: "Dots long", x: [1, 2, 3], y: [9, 5, 1] },
        { name: "Dots short", x: [1, 4], y: [2, 7] },
      ],
    },
  ],
};

function chartEffects(
  motion: ResolvedMotion,
  authorId: string,
): readonly ResolvedMotionEffect[] {
  return (
    motion.steps[0]?.effects.filter(
      ({ target }) => target.authorId === authorId,
    ) ?? []
  );
}

function chartRanges(
  motion: ResolvedMotion,
  authorId: string,
): readonly unknown[] {
  return chartEffects(motion, authorId).map(({ target }) => target.range);
}

describe("chart animation topology integration", () => {
  it("expands shared multi-ring categories in all four orders", async () => {
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="shared-ring-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`shared-ring-${by}`}
              name={`shared-ring-${by}`}
              groups={sharedRingGroups}
              legend="bottom"
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
        { step: "series", seriesIndex: 2, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        ...[0, 1, 2].flatMap((seriesIndex) =>
          [0, 1, 2].map((categoryIndex) => ({
            step: "ptInSeries",
            seriesIndex,
            categoryIndex,
          })),
        ),
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        ...[0, 1, 2].flatMap((categoryIndex) =>
          [0, 1, 2].map((seriesIndex) => ({
            step: "ptInCategory",
            seriesIndex,
            categoryIndex,
          })),
        ),
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `shared-ring-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const by of orders) {
      const object = file.slides[0]?.objects.find(
        ({ name }) => name === `shared-ring-${by}`,
      );
      assert.ok(object?.chart !== undefined);
      assert.deepEqual(
        object.chart.series.map(({ index, categories, values }) => ({
          index,
          categories,
          values,
        })),
        [
          {
            index: 0,
            categories: [...sharedCategories],
            values: [12, null, 18],
          },
          { index: 1, categories: [...sharedCategories], values: [10, 15, 20] },
          { index: 2, categories: [...sharedCategories], values: [11, 16, 19] },
        ],
      );
      const fileEffects = file.slides[0]?.animations?.[0]?.effects.filter(
        ({ targetId }) => targetId === object.id,
      );
      const build = {
        series: "series",
        category: "category",
        seriesElements: "seriesEl",
        categoryElements: "categoryEl",
      } as const;
      assert.deepEqual(
        fileEffects?.map(({ targetRange, nativeBuild }) => ({
          targetRange,
          nativeBuild,
        })),
        expected[by].map((range) => ({
          targetRange: { type: "chart", ...range },
          nativeBuild: build[by],
        })),
        by,
      );
    }
  });

  it("expands independent rings of 3, 2, and 4 slots without compressing indices", async () => {
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="independent-ring-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`independent-ring-${by}`}
              name={`independent-ring-${by}`}
              groups={independentRingGroups}
              legend="bottom"
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
        { step: "series", seriesIndex: 2, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
        { step: "category", seriesIndex: -4, categoryIndex: 3 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 2 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 2 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 3 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 2 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 2 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 3 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `independent-ring-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const by of orders) {
      const object = file.slides[0]?.objects.find(
        ({ name }) => name === `independent-ring-${by}`,
      );
      assert.ok(object?.chart !== undefined);
      assert.deepEqual(
        object.chart.series.map(({ index, categories, values }) => ({
          index,
          categories,
          values,
        })),
        [
          {
            index: 0,
            categories: ["Direct", "Partner", "Partner"],
            values: [30, 20, null],
          },
          {
            index: 1,
            categories: ["Partner", "Direct"],
            values: [40, 10],
          },
          {
            index: 2,
            categories: ["Direct", "Partner", "Online", "Partner"],
            values: [5, 6, 7, 8],
          },
        ],
      );
    }
  });

  it("completes every independent-ring point before a sequenced successor starts", async () => {
    function RingScopeDeck() {
      const source = useMotionScope();
      const follower = useMotionScope();
      return (
        <Deck width={520} height={360}>
          <Slide id="ring-scope">
            <MotionScope scope={source}>
              <Chart
                id="ring-scope-chart"
                groups={independentRingGroups}
                enter={fade({ by: "categoryElements", duration: 100 })}
                style={{
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 500,
                  height: 220,
                }}
              />
            </MotionScope>
            <MotionScope scope={follower} startAt={after(source.end, 25)}>
              <Shape
                id="ring-scope-follower"
                enter={fade({ duration: 50 })}
                style={{
                  position: "absolute",
                  left: 10,
                  top: 245,
                  width: 60,
                  height: 40,
                  background: "#1D4ED8",
                }}
              />
            </MotionScope>
          </Slide>
        </Deck>
      );
    }

    const result = await build(<RingScopeDeck />);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    assert.deepEqual(
      chartEffects(motion, "ring-scope-chart").map(({ startMs, endMs }) => [
        startMs,
        endMs,
      ]),
      [
        [0, 100],
        [100, 200],
        [200, 300],
        [300, 400],
        [400, 500],
        [500, 600],
        [600, 700],
        [700, 800],
        [800, 900],
        [900, 1000],
      ],
    );
    assert.deepEqual(
      motion.steps[0]?.effects
        .filter(({ target }) => target.authorId === "ring-scope-follower")
        .map(({ startMs, endMs }) => [startMs, endMs]),
      [[1025, 1075]],
    );

    const query = motionSchedule(result.report.slides[0]!);
    const queryEffects = query?.steps[0]?.effects ?? [];
    const last = queryEffects
      .filter(({ target }) => target.authorId === "ring-scope-chart")
      .at(-1);
    assert.deepEqual([last?.startMs, last?.endMs], [900, 1000]);
    const follower = queryEffects.find(
      ({ target }) => target.authorId === "ring-scope-follower",
    );
    assert.deepEqual([follower?.startMs, follower?.endMs], [1025, 1075]);
  });

  it("expands a shared-category three-axis two-group combo in all four orders", async () => {
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="shared-combo-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`shared-combo-${by}`}
              name={`shared-combo-${by}`}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
              {...sharedCombo}
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        ...[0, 1].flatMap((seriesIndex) =>
          [0, 1, 2].map((categoryIndex) => ({
            step: "ptInSeries",
            seriesIndex,
            categoryIndex,
          })),
        ),
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        ...[0, 1, 2].flatMap((categoryIndex) =>
          [0, 1].map((seriesIndex) => ({
            step: "ptInCategory",
            seriesIndex,
            categoryIndex,
          })),
        ),
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `shared-combo-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "shared-combo-seriesElements",
    );
    assert.ok(object?.chart !== undefined);
    assert.equal(object.chart.axes.length, 3);
    assert.deepEqual(
      object.chart.series.map(({ index, categories, values }) => ({
        index,
        categories,
        values,
      })),
      [
        { index: 0, categories: ["A", "B", "C"], values: [2, 6, 4] },
        { index: 1, categories: ["A", "B", "C"], values: [30, 70, 45] },
      ],
    );
  });

  it("expands mixed category and XY domains by global series slots", async () => {
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="mixed-domain-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`mixed-domain-${by}`}
              name={`mixed-domain-${by}`}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
              {...mixedDomainCombo}
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
        { step: "series", seriesIndex: 2, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 2 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 1 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 0, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 2 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `mixed-domain-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "mixed-domain-seriesElements",
    );
    assert.ok(object?.chart !== undefined);
    assert.equal(object.chart.axes.length, 4);
    assert.deepEqual(
      object.chart.series.map(({ index, categories, x, y, values }) => ({
        index,
        ...(categories === undefined ? {} : { categories }),
        ...(x === undefined ? {} : { x }),
        ...(y === undefined ? {} : { y }),
        ...(values === undefined ? {} : { values }),
      })),
      [
        { index: 0, categories: ["A", "B"], values: [2, 8] },
        {
          index: 1,
          x: [1, 2, 3],
          y: [9, 5, 1],
        },
        {
          index: 2,
          x: [1, 4],
          y: [2, 7],
        },
      ],
    );
  });
  it("keeps all-zero chart orders limited to the real background target", async () => {
    const emptyGroups = [
      {
        kind: "column" as const,
        data: { kind: "shared-categories" as const, categories: [] },
        series: [
          { name: "Empty North", values: [], color: "#1D4ED8" },
          { name: "Empty South", values: [], color: "#C2410C" },
        ],
      },
    ];
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="all-zero-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`all-zero-${by}`}
              name={`all-zero-${by}`}
              groups={emptyGroups}
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
      ],
      category: [{ step: "gridLegend", seriesIndex: -3, categoryIndex: -3 }],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `all-zero-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
      assert.deepEqual(
        chartEffects(motion, `all-zero-${by}`).map(({ startMs, endMs }) => [
          startMs,
          endMs,
        ]),
        expected[by].map((_, index) => [index * 100, (index + 1) * 100]),
        by,
      );
    }

    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    for (const by of orders) {
      const object = file.slides[0]?.objects.find(
        ({ name }) => name === `all-zero-${by}`,
      );
      assert.ok(object?.chart !== undefined);
      assert.deepEqual(
        object.chart.series.map(({ name, values }) => ({ name, values })),
        [
          { name: "Empty North", values: [] },
          { name: "Empty South", values: [] },
        ],
      );
      const build = {
        series: "series",
        category: "category",
        seriesElements: "seriesEl",
        categoryElements: "categoryEl",
      } as const;
      assert.deepEqual(
        file.slides[0]?.animations?.[0]?.effects
          .filter(({ targetId }) => targetId === object.id)
          .map(({ targetRange, nativeBuild }) => ({
            targetRange,
            nativeBuild,
          })),
        expected[by].map((range) => ({
          targetRange: { type: "chart", ...range },
          nativeBuild: build[by],
        })),
        by,
      );
    }
  });

  it("keeps all-zero build targets valid only for the background", async () => {
    const result = await build(
      <Deck width={480} height={300}>
        <Slide id="all-zero-negative">
          <Chart
            id="all-zero-negative-chart"
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [{ name: "Empty", values: [] }],
              },
            ]}
            enter={fade({ by: "seriesElements", duration: 100 })}
            style={{
              position: "absolute",
              left: 10,
              top: 10,
              width: 460,
              height: 280,
            }}
          />
        </Slide>
      </Deck>,
    );
    const source = await validatePptx(result.pptx);
    assert.deepEqual(source.errors, []);
    const sourceXml = strFromU8(
      unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
    );
    const buildElement = sourceXml.match(
      /<p:bldGraphic\b[\s\S]*?<\/p:bldGraphic>/u,
    )?.[0];
    assert.ok(buildElement !== undefined);
    const groupMatch = buildElement.match(/grpId="(\d+)"/u);
    assert.ok(groupMatch !== null);
    const groupId = Number(groupMatch[1]);
    const expectMutationError = async (
      label: string,
      mutate: (xml: string) => string,
      message: RegExp,
    ): Promise<void> => {
      const read = await validatePptx(rewriteSlideXml(result.pptx, mutate));
      assert.ok(
        read.errors.some(
          (error) =>
            error.code === "invalid-animation-target-range" &&
            message.test(error.message),
        ) ||
          read.errors.some(
            (error) =>
              error.code === "invalid-chart-build-group" &&
              message.test(error.message),
          ),
        `${label}: ${JSON.stringify(read.errors)}`,
      );
    };

    await expectMutationError(
      "point into zero series",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="-3" categoryIdx="-3" bldStep="gridLegend"',
          'seriesIdx="0" categoryIdx="0" bldStep="ptInSeries"',
        ),
      /exceeds points in native chart series/u,
    );
    await expectMutationError(
      "background sentinel",
      (xml) =>
        xml.replaceAll(
          'seriesIdx="-3" categoryIdx="-3"',
          'seriesIdx="-4" categoryIdx="-3"',
        ),
      /native -3 sentinel/u,
    );
    await expectMutationError(
      "invalid build declaration",
      (xml) =>
        xml.replace(
          buildElement,
          buildElement.replace('bld="seriesEl"', 'bld="bogus"'),
        ),
      /supported native build mode/u,
    );
    await expectMutationError(
      "step mismatch",
      (xml) => xml.replaceAll('bldStep="gridLegend"', 'bldStep="ptInCategory"'),
      /does not match native chart build mode/u,
    );
    await expectMutationError(
      "build group",
      (xml) =>
        xml.replace(
          buildElement,
          buildElement.replace(`grpId="${groupId}"`, 'grpId="777"'),
        ),
      /actual chart animation effect/u,
    );
  });

  it("completes an all-zero point expansion at its background duration", async () => {
    function ZeroScopeDeck() {
      const source = useMotionScope();
      const follower = useMotionScope();
      return (
        <Deck width={520} height={360}>
          <Slide id="zero-scope">
            <MotionScope scope={source}>
              <Chart
                id="zero-scope-chart"
                groups={[
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: [] },
                    series: [{ name: "Empty", values: [] }],
                  },
                ]}
                enter={fade({ by: "categoryElements", duration: 100 })}
                style={{
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 500,
                  height: 200,
                }}
              />
            </MotionScope>
            <MotionScope scope={follower} startAt={after(source.end, 25)}>
              <Shape
                id="zero-scope-follower"
                enter={fade({ duration: 50 })}
                style={{
                  position: "absolute",
                  left: 10,
                  top: 220,
                  width: 60,
                  height: 40,
                  background: "#1D4ED8",
                }}
              />
            </MotionScope>
          </Slide>
        </Deck>
      );
    }

    const result = await build(<ZeroScopeDeck />);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    assert.deepEqual(
      chartEffects(motion, "zero-scope-chart").map(({ startMs, endMs }) => [
        startMs,
        endMs,
      ]),
      [[0, 100]],
    );
    assert.deepEqual(
      motion.steps[0]?.effects
        .filter(({ target }) => target.authorId === "zero-scope-follower")
        .map(({ startMs, endMs }) => [startMs, endMs]),
      [[125, 175]],
    );
    const query = motionSchedule(result.report.slides[0]!);
    const queryEffects = query?.steps[0]?.effects ?? [];
    const last = queryEffects
      .filter(({ target }) => target.authorId === "zero-scope-chart")
      .at(-1);
    assert.deepEqual([last?.startMs, last?.endMs], [0, 100]);
    const follower = queryEffects.find(
      ({ target }) => target.authorId === "zero-scope-follower",
    );
    assert.deepEqual([follower?.startMs, follower?.endMs], [125, 175]);
  });

  it("keeps mixed independent ring counts without renumbering series", async () => {
    const groups = [
      {
        kind: "doughnut" as const,
        data: { kind: "independent-categories" as const },
        series: [
          { name: "Empty ring", categories: [], values: [] },
          { name: "Two points", categories: ["A", "B"], values: [1, 2] },
          {
            name: "Three points",
            categories: ["A", "B", "C"],
            values: [3, 4, 5],
          },
        ],
      },
    ];
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="mixed-ring-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`mixed-ring-${by}`}
              name={`mixed-ring-${by}`}
              groups={groups}
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
        { step: "series", seriesIndex: 2, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 2 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 2 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `mixed-ring-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "mixed-ring-seriesElements",
    );
    assert.ok(object?.chart !== undefined);
    assert.deepEqual(
      object.chart.series.map(({ index, categories, values }) => ({
        index,
        categories,
        values,
      })),
      [
        { index: 0, categories: [], values: [] },
        { index: 1, categories: ["A", "B"], values: [1, 2] },
        { index: 2, categories: ["A", "B", "C"], values: [3, 4, 5] },
      ],
    );
  });

  it("keeps mixed XY counts without renumbering global series", async () => {
    const groups = [
      {
        kind: "scatter" as const,
        data: { kind: "xy" as const },
        series: [
          { name: "Empty", x: [], y: [] },
          { name: "Two points", x: [1, 2], y: [3, 4] },
          { name: "Three points", x: [5, 6, 7], y: [8, 9, 10] },
        ],
      },
    ];
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="mixed-xy-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`mixed-xy-${by}`}
              name={`mixed-xy-${by}`}
              groups={groups}
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
        { step: "series", seriesIndex: 2, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
        { step: "category", seriesIndex: -4, categoryIndex: 2 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInSeries", seriesIndex: 2, categoryIndex: 2 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 1 },
        { step: "ptInCategory", seriesIndex: 2, categoryIndex: 2 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `mixed-xy-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "mixed-xy-seriesElements",
    );
    assert.ok(object?.chart !== undefined);
    assert.deepEqual(
      object.chart.series.map(({ index, x, y }) => ({ index, x, y })),
      [
        { index: 0, x: [], y: [] },
        { index: 1, x: [1, 2], y: [3, 4] },
        { index: 2, x: [5, 6, 7], y: [8, 9, 10] },
      ],
    );
  });

  it("keeps one empty and one non-empty category axis domain", async () => {
    const groups = [
      {
        kind: "column" as const,
        axes: { category: "emptyCategories", value: "emptyValues" },
        data: { kind: "shared-categories" as const, categories: [] },
        series: [{ name: "Empty", values: [] }],
      },
      {
        kind: "line" as const,
        axes: { category: "fullCategories", value: "fullValues" },
        data: {
          kind: "shared-categories" as const,
          categories: ["A", "B"],
        },
        series: [{ name: "Filled", values: [1, 2] }],
      },
    ];
    const axes = {
      emptyCategories: {
        kind: "category" as const,
        position: "bottom" as const,
      },
      emptyValues: {
        kind: "value" as const,
        position: "left" as const,
        min: 0,
        max: 10,
      },
      fullCategories: { kind: "category" as const, position: "top" as const },
      fullValues: {
        kind: "value" as const,
        position: "right" as const,
        min: 0,
        max: 10,
      },
    };
    const result = await build(
      <Deck width={1000} height={620}>
        <Slide id="mixed-domain-axis-orders">
          {orders.map((by, index) => (
            <Chart
              key={by}
              id={`mixed-axis-${by}`}
              name={`mixed-axis-${by}`}
              axes={axes}
              groups={groups}
              enter={fade({ by, duration: 100 })}
              style={{
                position: "absolute",
                left: 10 + (index % 2) * 500,
                top: 10 + Math.floor(index / 2) * 305,
                width: 480,
                height: 280,
              }}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    const expected = {
      series: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "series", seriesIndex: 0, categoryIndex: -4 },
        { step: "series", seriesIndex: 1, categoryIndex: -4 },
      ],
      category: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "category", seriesIndex: -4, categoryIndex: 0 },
        { step: "category", seriesIndex: -4, categoryIndex: 1 },
      ],
      seriesElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInSeries", seriesIndex: 1, categoryIndex: 1 },
      ],
      categoryElements: [
        { step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 0 },
        { step: "ptInCategory", seriesIndex: 1, categoryIndex: 1 },
      ],
    } as const;
    for (const by of orders) {
      assert.deepEqual(
        chartRanges(motion, `mixed-axis-${by}`),
        expected[by].map((range) => ({ type: "chart", ...range })),
        by,
      );
    }
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects.find(
      ({ name }) => name === "mixed-axis-seriesElements",
    );
    assert.ok(object?.chart !== undefined);
    assert.equal(object.chart.axes.length, 4);
    assert.deepEqual(
      object.chart.series.map(({ index, categories, values }) => ({
        index,
        categories,
        values,
      })),
      [
        { index: 0, categories: [], values: [] },
        { index: 1, categories: ["A", "B"], values: [1, 2] },
      ],
    );
  });
  it("builds the same all-zero background target from the command entry", async () => {
    const document = deck({ width: 480, height: 300 }).append(
      slide({ id: "zero-command" }).append(
        documentChart({
          id: "zero-command-chart",
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: [] },
              series: [{ name: "Empty", values: [] }],
            },
          ],
          enter: fade({ by: "categoryElements", duration: 100 }),
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 460,
            height: 260,
          },
        }),
      ),
    );
    const result = await buildDocument(document);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as ResolvedMotion;
    assert.deepEqual(chartRanges(motion, "zero-command-chart"), [
      { type: "chart", step: "gridLegend", seriesIndex: -3, categoryIndex: -3 },
    ]);
    assert.deepEqual(
      chartEffects(motion, "zero-command-chart").map(({ startMs, endMs }) => [
        startMs,
        endMs,
      ]),
      [[0, 100]],
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const object = file.slides[0]?.objects[0];
    assert.equal(object?.kind, "chart");
    assert.deepEqual(
      file.slides[0]?.animations?.[0]?.effects.map(
        ({ nativeBuild, targetRange }) => ({ nativeBuild, targetRange }),
      ),
      [
        {
          nativeBuild: "categoryEl",
          targetRange: {
            type: "chart",
            step: "gridLegend",
            seriesIndex: -3,
            categoryIndex: -3,
          },
        },
      ],
    );
  });
});
