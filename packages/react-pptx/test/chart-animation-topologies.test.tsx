import assert from "node:assert/strict";

import {
  after,
  build,
  Chart,
  Deck,
  fade,
  MotionScope,
  observeMotion,
  Shape,
  Slide,
  useMotionScope,
  type MotionEffectObservation,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { describe, it } from "vitest";

const orders = [
  "series",
  "category",
  "seriesElements",
  "categoryElements",
] as const;

const sharedCategories = ["Direct", "Partner", "Direct"] as const;

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
  motion: MotionObservation,
  authorId: string,
): readonly MotionEffectObservation[] {
  return (
    motion.steps[0]?.effects.filter(
      ({ target }) => target.authorId === authorId,
    ) ?? []
  );
}

function chartRanges(
  motion: MotionObservation,
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
              x={10 + (index % 2) * 500}
              y={10 + Math.floor(index / 2) * 305}
              width={480}
              height={280}
              groups={sharedRingGroups}
              legend="bottom"
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
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
              x={10 + (index % 2) * 500}
              y={10 + Math.floor(index / 2) * 305}
              width={480}
              height={280}
              groups={independentRingGroups}
              legend="bottom"
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
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
                x={10}
                y={10}
                width={500}
                height={220}
                groups={independentRingGroups}
                enter={fade({ by: "categoryElements", duration: 100 })}
              />
            </MotionScope>
            <MotionScope scope={follower} startAt={after(source.end, 25)}>
              <Shape
                id="ring-scope-follower"
                x={10}
                y={245}
                width={60}
                height={40}
                fill="#1D4ED8"
                enter={fade({ duration: 50 })}
              />
            </MotionScope>
          </Slide>
        </Deck>
      );
    }

    const result = await build(<RingScopeDeck />);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
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

    const query = observeMotion(result.report.slides[0]!);
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
              x={10 + (index % 2) * 500}
              y={10 + Math.floor(index / 2) * 305}
              width={480}
              height={280}
              {...sharedCombo}
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
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
              x={10 + (index % 2) * 500}
              y={10 + Math.floor(index / 2) * 305}
              width={480}
              height={280}
              {...mixedDomainCombo}
              enter={fade({ by, duration: 100 })}
            />
          ))}
        </Slide>
      </Deck>,
    );
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
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
});
