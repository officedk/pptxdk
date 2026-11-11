import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { validatePptx } from "@react-pptx/validation";
import { describe, it } from "vitest";

type CategoryFixture = {
  readonly file: string;
  readonly sha256: string;
  readonly mismatch: boolean;
  readonly check?: (
    series: readonly {
      readonly name: string;
      readonly categories?: readonly string[];
      readonly values?: readonly (number | null)[];
      readonly x?: readonly (number | null)[];
      readonly y?: readonly (number | null)[];
    }[],
  ) => void;
};

const fixtures: readonly CategoryFixture[] = [
  {
    file: "category-source.pptx",
    sha256: "cb729891318efbd58683444828dabf4127160b7abe3cbcbfa9a34bc68fa62476",
    mismatch: false,
  },
  {
    file: "category-category-only.pptx",
    sha256: "d924f88f8dbc8cf40ecde9b4eeded6226d6559481a5c68c507c745a445dfc51e",
    mismatch: true,
  },
  {
    file: "category-value-only.pptx",
    sha256: "7aea0bfb480147eb5495649e6212ef849cf741ba6f701b9bd239498f3e138406",
    mismatch: true,
  },
  {
    file: "rings-source.pptx",
    sha256: "24e65124c1de0681ac6f98e9fa5eb98c6d1cb4b5fe9fce87461e181abd6db17e",
    mismatch: false,
    check: (series) => {
      assert.deepEqual(series[0]?.categories, []);
      assert.deepEqual(series[1]?.categories, ["A", "B"]);
      assert.deepEqual(series[1]?.values, [2, 3]);
    },
  },
  {
    file: "rings-category-only.pptx",
    sha256: "5c84d66ae2423f0a44031acdacc4013bb1e6497e7202001f85a7f979df1a9af4",
    mismatch: true,
    check: (series) => {
      assert.deepEqual(series[0]?.categories, []);
      assert.deepEqual(series[1]?.categories, ["A", "B"]);
      assert.deepEqual(series[1]?.values, [2, 3]);
    },
  },
  {
    file: "rings-value-only.pptx",
    sha256: "8648544fdf5a97baa44ff45950f3a4a967e3a65229d63b5f0f4f17dd0833c20f",
    mismatch: true,
    check: (series) => {
      assert.deepEqual(series[0]?.categories, []);
      assert.deepEqual(series[1]?.categories, ["A", "B"]);
    },
  },
  {
    file: "scatter-category-only.pptx",
    sha256: "add98f8b0c3a26de7b0897b968eb0f32bd6cfff59c6a8acb1bffc1ecc90a3f9e",
    mismatch: false,
    check: (series) => {
      assert.deepEqual(series[0]?.x, []);
      assert.deepEqual(series[0]?.y, []);
      assert.deepEqual(series[1]?.x, [1]);
      assert.deepEqual(series[1]?.y, [2]);
    },
  },
];

describe("native zero-source category column ownership", () => {
  for (const fixture of fixtures) {
    it(`checks the owning category column in ${fixture.file}`, async () => {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(
            `./assets/native-23-zero-category/${fixture.file}`,
            import.meta.url,
          ),
        ),
      );
      assert.equal(
        createHash("sha256").update(bytes).digest("hex"),
        fixture.sha256,
      );
      const result = await validatePptx(bytes);
      const chart = result.slides[0]?.objects[0]?.chart;
      assert.ok(chart);
      fixture.check?.(chart.series);
      const targetMismatch = result.errors.filter(
        ({ code, part }) =>
          code === "chart-workbook-mismatch" &&
          part === "ppt/charts/chart1.xml",
      );
      assert.equal(targetMismatch.length, fixture.mismatch ? 1 : 0);
      if (!fixture.mismatch) {
        assert.deepEqual(result.errors, []);
      }
    });
  }
});
