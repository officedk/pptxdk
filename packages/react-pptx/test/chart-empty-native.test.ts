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
