import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

type ExpectedSeries = {
  readonly name: string;
  readonly categories?: readonly string[];
  readonly values?: readonly (number | null)[];
  readonly x?: readonly (number | null)[];
  readonly y?: readonly (number | null)[];
};

type NativeFixture = {
  readonly file: string;
  readonly sha256: string;
  readonly expected: readonly ExpectedSeries[];
};

const fixtures: readonly NativeFixture[] = [
  {
    file: "23-absent-pp-untouched.pptx",
    sha256: "72912ea68f0b376146ee8e3d85cf402b154c98fdb5789320b76f979811d0407a",
    expected: [
      { name: "North", categories: [], values: [] },
      { name: "Share", categories: [], values: [] },
      { name: "Observations", x: [], y: [] },
    ],
  },
  {
    file: "23-absent-wps-untouched.pptx",
    sha256: "fc80e861b7092fcfb3fb23ae6f799186eefdaead9bf83d0080552cbc16ae0696",
    expected: [
      { name: "North", categories: [], values: [] },
      { name: "Share", categories: [], values: [] },
      { name: "Observations", x: [], y: [] },
    ],
  },
];

const filledCellsFixtures: readonly NativeFixture[] = [
  {
    file: "23-absent-pp-filled-cells.pptx",
    sha256: "f6bb4bcdce2067aacdbd750a40fe1e81e279b730a7cd91f2131131797ba19865",
    expected: [
      { name: "North", categories: [], values: [] },
      { name: "Share", categories: [], values: [] },
      { name: "Observations", x: [], y: [] },
    ],
  },
  {
    file: "23-absent-wps-filled-cells.pptx",
    sha256: "ac8b373aff1ab9284fb88c74518d5c1442ee29e77a5d1214e595bb9f18e24954",
    expected: [
      { name: "North", categories: [], values: [] },
      { name: "Share", categories: [], values: [] },
      { name: "Observations", x: [], y: [] },
    ],
  },
];

const filledRangeFixture: NativeFixture = {
  file: "23-absent-pp-filled-range.pptx",
  sha256: "d37156b2bd8bd67c8bba12e3eb41adaa62f3bfde6185fa319b793a22db0a10e1",
  expected: [
    { name: "North", categories: ["A", "B"], values: [3, 7] },
    { name: "Share", categories: ["One", "Two"], values: [2, 5] },
    { name: "Y", x: [2, 5], y: [4, 1] },
  ],
};

const targets = [
  { slideIndex: 0, part: "ppt/charts/chart1.xml" },
  { slideIndex: 4, part: "ppt/charts/chart5.xml" },
  { slideIndex: 6, part: "ppt/charts/chart7.xml" },
] as const;

function assertSeries(
  observed: {
    readonly name: string;
    readonly categories?: readonly string[];
    readonly values?: readonly (number | null)[];
    readonly x?: readonly (number | null)[];
    readonly y?: readonly (number | null)[];
  },
  expected: ExpectedSeries,
): void {
  assert.equal(observed.name, expected.name);
  if (expected.categories !== undefined) {
    assert.deepEqual(observed.categories, expected.categories);
    assert.deepEqual(observed.values, expected.values);
    assert.equal(observed.x, undefined);
    assert.equal(observed.y, undefined);
  } else {
    assert.equal(observed.categories, undefined);
    assert.equal(observed.values, undefined);
    assert.deepEqual(observed.x, expected.x);
    assert.deepEqual(observed.y, expected.y);
  }
}

describe("native saved zero-data-source series", () => {
  for (const fixture of fixtures) {
    it(`reads ${fixture.file} without injecting observations`, async () => {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`./assets/native-23-absent/${fixture.file}`, import.meta.url),
        ),
      );
      assert.equal(
        createHash("sha256").update(bytes).digest("hex"),
        fixture.sha256,
      );
      const file = await validatePptx(bytes);
      assert.deepEqual(
        file.errors.filter(
          ({ code, part }) =>
            code === "invalid-chart-cache" &&
            targets.some(({ part: targetPart }) => targetPart === part),
        ),
        [],
      );
      const parts = unzipSync(bytes);
      for (const [index, target] of targets.entries()) {
        const object = file.slides[target.slideIndex]?.objects[0];
        assert.equal(object?.kind, "chart");
        assert.ok(object?.chart);
        assert.equal(object.chart.series.length, 1);
        assertSeries(object.chart.series[0]!, fixture.expected[index]!);
        const workbookSeries = object.chart.embeddedWorkbook?.series;
        assert.equal(workbookSeries?.length, 1);
        assertSeries(workbookSeries[0]!, fixture.expected[index]!);

        const xml = strFromU8(parts[target.part]!);
        assert.equal((xml.match(/<c:ser>/gu) ?? []).length, 1, target.part);
        assert.ok(
          xml.includes(`<c:v>${fixture.expected[index]!.name}</c:v>`),
          target.part,
        );
        assert.ok(!xml.includes("<c:numCache>"), target.part);
        if (target.part === "ppt/charts/chart7.xml") {
          assert.ok(!xml.includes("<c:xVal>"), target.part);
          assert.ok(!xml.includes("<c:yVal>"), target.part);
        } else {
          assert.ok(!xml.includes("<c:cat>"), target.part);
          assert.ok(!xml.includes("<c:val>"), target.part);
        }
      }
    });
  }
});

describe("native filled workbook without series binding", () => {
  for (const fixture of filledCellsFixtures) {
    it(`keeps ${fixture.file} zero-source inconsistent with its filled workbook`, async () => {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`./assets/native-23-absent/${fixture.file}`, import.meta.url),
        ),
      );
      assert.equal(
        createHash("sha256").update(bytes).digest("hex"),
        fixture.sha256,
      );
      const file = await validatePptx(bytes);
      const parts = unzipSync(bytes);
      for (const [index, target] of targets.entries()) {
        const object = file.slides[target.slideIndex]?.objects[0];
        assert.equal(object?.kind, "chart");
        assert.ok(object?.chart);
        assert.equal(object.chart.series.length, 1);
        assertSeries(object.chart.series[0]!, fixture.expected[index]!);
        const workbookSeries = object.chart.embeddedWorkbook?.series;
        assert.equal(workbookSeries?.length, 1);
        assertSeries(workbookSeries[0]!, fixture.expected[index]!);
        assert.ok(
          file.errors.some(
            ({ code, part }) =>
              code === "chart-workbook-mismatch" && part === target.part,
          ),
          `${fixture.file} ${target.part} must report mismatch`,
        );
        const workbookPart = object.chart.embeddedWorkbook?.part;
        assert.ok(workbookPart);
        const workbook = unzipSync(parts[workbookPart]!);
        const sheet = strFromU8(workbook["xl/worksheets/sheet1.xml"]!);
        assert.ok(
          sheet.includes('r="B2"') && sheet.includes('r="B3"'),
          `${target.part} workbook must contain filled cells`,
        );
      }
    });
  }
});

describe("native filled workbook with a full rebound range", () => {
  it("reads the actual two categories and XY observations", async () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL(
          `./assets/native-23-absent/${filledRangeFixture.file}`,
          import.meta.url,
        ),
      ),
    );
    assert.equal(
      createHash("sha256").update(bytes).digest("hex"),
      filledRangeFixture.sha256,
    );
    const file = await validatePptx(bytes);
    for (const [index, target] of targets.entries()) {
      const object = file.slides[target.slideIndex]?.objects[0];
      assert.equal(object?.kind, "chart");
      assert.ok(object?.chart);
      assert.equal(object.chart.series.length, 1);
      assertSeries(
        object.chart.series[0]!,
        filledRangeFixture.expected[index]!,
      );
      const workbookSeries = object.chart.embeddedWorkbook?.series;
      assert.equal(workbookSeries?.length, 1);
      assertSeries(workbookSeries[0]!, filledRangeFixture.expected[index]!);
      assert.deepEqual(
        file.errors.filter(
          ({ code, part }) =>
            part === target.part && code !== "invalid-data-label-leader-lines",
        ),
        [],
      );
    }
  });
});

describe("WPS native series UI binding", () => {
  it("reads the saved A/B categories and 3/7 values on p1", async () => {
    const bytes = new Uint8Array(
      readFileSync(
        new URL(
          "./assets/native-23-absent/23-absent-wps-series-ui.pptx",
          import.meta.url,
        ),
      ),
    );
    assert.equal(
      createHash("sha256").update(bytes).digest("hex"),
      "4787cd57af84b24f977a6f69e3433e144b7158521b3325bfe63b499a4af13001",
    );
    const file = await validatePptx(bytes);
    const object = file.slides[0]?.objects[0];
    assert.equal(object?.kind, "chart");
    assert.ok(object?.chart);
    assertSeries(object.chart.series[0]!, {
      name: "North",
      categories: ["A", "B"],
      values: [3, 7],
    });
    const workbookSeries = object.chart.embeddedWorkbook?.series;
    assert.equal(workbookSeries?.length, 1);
    assertSeries(workbookSeries[0]!, {
      name: "North",
      categories: ["A", "B"],
      values: [3, 7],
    });
    assert.deepEqual(
      file.errors.filter(
        ({ code, part }) =>
          part === "ppt/charts/chart1.xml" &&
          (code === "invalid-chart-cache" ||
            code === "chart-workbook-mismatch"),
      ),
      [],
    );
  });
});
