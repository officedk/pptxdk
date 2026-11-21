import assert from "node:assert/strict";

import { validatePptx } from "@pptxdk/validation";
import { emit, expandDocument, resolve } from "@react-pptx/react-pptx";
import { paint } from "@react-pptx/react-pptx/preview";
import { it } from "vitest";

import { DeepSeekSystemsDeck } from "../src/index.ts";

function assertValidPng1280x720(png: Uint8Array, label: string): void {
  const pngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  assert.ok(png.byteLength > 24, `${label}: PNG too small`);
  assert.deepEqual(
    [...png.subarray(0, 8)],
    pngSignature,
    `${label}: invalid PNG signature`,
  );
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  assert.equal(view.getUint32(16, false), 1280, `${label}: PNG width`);
  assert.equal(view.getUint32(20, false), 720, `${label}: PNG height`);
}

it("full deck preserves 14 slides, zero-diagnostic static preview, native charts and tables, morph continuity, click animations, and valid slide bounds", async () => {
  const resolution = resolve(await expandDocument(<DeepSeekSystemsDeck />));
  const result = emit(resolution, {
    embedFonts: "full",
  });
  assert.deepEqual(resolution.report.diagnostics, []);
  assert.deepEqual(result.report.diagnostics, []);

  const painted = paint(resolution);
  assert.equal(painted.pages.length, 14);
  for (const page of painted.pages) {
    assert.deepEqual(page.diagnostics, [], `preview page ${page.page}`);
    assert.ok(page.png !== undefined, `preview page ${page.page} png`);
    assertValidPng1280x720(page.png, `preview page ${page.page}`);
  }

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

  const openingSlide = file.slides.find(({ id }) => id === "opening")!;
  const kpiAuthorIds = Array.from(
    { length: 4 },
    (_, index) => `opening-kpi-${index + 1}`,
  );
  const kpiGroups = kpiAuthorIds.map((authorId) => {
    const groupId = result.report.slides
      .find(({ id }) => id === "opening")
      ?.objects.find(
        (object) =>
          object.kind === "group" &&
          object.members?.some(({ authorId: value }) => value === authorId),
      )?.id;
    const groups = openingSlide.objects.filter(
      (object) => object.kind === "group" && object.id === groupId,
    );
    assert.equal(groups.length, 1, `${authorId} should be one native group`);

    const [group] = groups;
    assert.ok(group);
    const cardTexts =
      group.members
        ?.filter((member) => member.text !== undefined)
        .map((member) => member.text?.text) ?? [];
    assert.equal(cardTexts.length, 3, `${authorId} should group all card text`);
    assert.ok(
      cardTexts.every((text) => (text?.trim().length ?? 0) > 0),
      `${authorId} should keep the card text non-empty`,
    );
    assert.equal(
      group.members?.filter(
        (member) => member.kind === "shape" && member.text === undefined,
      ).length,
      2,
      `${authorId} should group its background and accent`,
    );

    return group.id;
  });
  const kpiEffects = (openingSlide.animations ?? []).flatMap((step) =>
    step.effects
      .filter((effect) => kpiGroups.includes(effect.targetId))
      .map((effect) => ({ trigger: step.trigger, effect })),
  );
  assert.equal(kpiEffects.length, 4);
  assert.deepEqual(
    kpiEffects
      .map(({ effect }) => effect.targetId)
      .sort((left, right) => left - right),
    [...kpiGroups].sort((left, right) => left - right),
  );
  assert.deepEqual(
    kpiEffects
      .map(({ trigger, effect }) => [
        trigger,
        effect.preset,
        effect.durationMs,
        effect.delayMs,
      ])
      .sort((left, right) => Number(left[3]) - Number(right[3])),
    [
      ["auto", "floatUp", 420, 0],
      ["auto", "floatUp", 420, 120],
      ["auto", "floatUp", 420, 240],
      ["auto", "floatUp", 420, 360],
    ],
  );
  assert.equal(
    openingSlide.animations?.some((step) => step.trigger === "click"),
    false,
    "KPI entrances should not add a click step",
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
