import assert from "node:assert/strict";

import { build } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { it } from "vitest";

import { DeepSeekSystemsDeck } from "../src/index.ts";

it("full deck preserves 14 slides, native charts and tables, morph continuity, click animations, and valid slide bounds", async () => {
  const result = await build(<DeepSeekSystemsDeck />, {
    embedFonts: "full",
  });
  assert.deepEqual(result.report.diagnostics, []);

  const file = await validatePptx(result.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  assert.deepEqual(file.errors, []);
  assert.equal(file.slides.length, 14);
  assert.deepEqual(
    file.slides.map((slide) => slide.id),
    [
      "opening",
      "training-cost",
      "mla-moe-efficiency",
      "aux-loss-free-mtp",
      "fp8-kernels",
      "dualpipe",
      "mpft-deepep",
      "three-fs",
      "pd-disaggregation",
      "decode-pipeline",
      "production-metrics",
      "r1-rl-distill",
      "frontier-benchmarks",
      "open-infra-synthesis",
    ],
  );

  for (const slide of file.slides) {
    for (const object of slide.objects) {
      if (object.boxEmu === undefined) continue;
      assert.ok(
        object.boxEmu.x >= 0,
        `${slide.id}/${object.name} overflows left boundary`,
      );
      assert.ok(
        object.boxEmu.y >= 0,
        `${slide.id}/${object.name} overflows top boundary`,
      );
      assert.ok(
        object.boxEmu.x + object.boxEmu.width <= 1280 * 9525,
        `${slide.id}/${object.name} overflows right boundary`,
      );
      assert.ok(
        object.boxEmu.y + object.boxEmu.height <= 720 * 9525,
        `${slide.id}/${object.name} overflows bottom boundary`,
      );
    }
  }

  // Verify all native charts and their embedded Excel workbooks are synchronized.
  const charts = file.slides.flatMap((slide) =>
    slide.objects.flatMap((object) =>
      object.chart === undefined
        ? []
        : [{ slideId: slide.id, chart: object.chart }],
    ),
  );
  assert.equal(charts.length, 12);
  for (const { slideId, chart } of charts) {
    assert.ok(
      chart.series.length > 0,
      `Chart on ${slideId} must contain at least one series`,
    );
    assert.deepEqual(
      chart.embeddedWorkbook?.series,
      chart.series,
      `Embedded workbook on ${slideId} must match chart series`,
    );
  }

  // Verify native tables exist on the three architectural comparison slides.
  for (const slideId of ["training-cost", "dualpipe", "pd-disaggregation"]) {
    const slide = file.slides.find(({ id }) => id === slideId)!;
    assert.ok(
      slide.objects.some((object) => object.kind === "table"),
      `Slide ${slideId} should contain a native editable table`,
    );
  }

  // Verify Morph transition and shared !! continuity identities between Slide 09 and Slide 10.
  const pdSlide = file.slides.find(({ id }) => id === "pd-disaggregation")!;
  const decodeSlide = file.slides.find(({ id }) => id === "decode-pipeline")!;
  assert.equal(decodeSlide.transition?.kind, "morph");
  assert.ok((decodeSlide.transition?.durationMs ?? 0) > 0);
  const pdMorphNames = new Set(
    pdSlide.objects
      .filter(({ name }) => name.startsWith("!!"))
      .map(({ name }) => name),
  );
  const decodeMorphNames = new Set(
    decodeSlide.objects
      .filter(({ name }) => name.startsWith("!!"))
      .map(({ name }) => name),
  );
  assert.equal(pdMorphNames.size, 5);
  assert.deepEqual(decodeMorphNames, pdMorphNames);

  // Verify intentional click animations on the 4 interactive slides.
  for (const slideId of [
    "aux-loss-free-mtp",
    "dualpipe",
    "decode-pipeline",
    "r1-rl-distill",
  ]) {
    const slide = file.slides.find(({ id }) => id === slideId)!;
    assert.ok(
      (slide.animations?.length ?? 0) >= 1,
      `Slide ${slideId} should have click-triggered animation steps`,
    );
    for (const step of slide.animations ?? []) {
      assert.equal(step.trigger, "click");
      assert.ok(step.effects.length > 0);
    }
  }
}, 60_000);
