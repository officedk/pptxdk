import assert from "node:assert/strict";

import {
  emit,
  motionSchedule,
  resolve,
  type ResolvedObject,
  type TextMetrics,
} from "@officedk/pptxdk";
import { paint } from "@officedk/pptxdk/preview";
import { expandDocument } from "@officedk/pptxdk/react";
import { deck, slide } from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { it } from "vitest";

import { createDualPipeSchedule } from "../src/components/dualpipe-schedule.ts";
import { fetchGitHubProjects } from "../src/github-projects.ts";
import { DeepSeekSystemsDeck } from "../src/index.ts";
import { theme } from "../src/theme.ts";

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

function flattenObjects(objects: readonly ResolvedObject[]): ResolvedObject[] {
  return objects.flatMap((object) => [
    object,
    ...(object.members === undefined ? [] : flattenObjects(object.members)),
  ]);
}

function assertMeasuredTextFits(
  metrics: TextMetrics | undefined,
  label: string,
): void {
  assert.ok(metrics !== undefined, `${label}: expected textMetrics`);
  assert.equal(
    metrics.status,
    "measured",
    `${label}: expected measured textMetrics, got ${metrics.status}`,
  );
  assert.equal(
    metrics.fits,
    true,
    `${label}: expected text to fit (available=${JSON.stringify(metrics.available)}, content=${JSON.stringify(metrics.content)})`,
  );
}

it("full deck preserves 15 slides, zero-error static preview, native charts and tables, morph continuity, click animations, measured text fit, and valid slide bounds", async () => {
  const projects = new Map([
    ["FlashMLA", 0],
    ["DeepEP", 500],
    ["DeepGEMM", 1200],
  ]);
  const requests: string[] = [];
  const fetchProject: typeof fetch = async (input) => {
    assert.ok(typeof input === "string");
    const url = new URL(input);
    const name = url.pathname.split("/").at(-1)!;
    assert.equal(url.href, `https://api.github.com/repos/deepseek-ai/${name}`);
    assert.ok(projects.has(name));
    requests.push(name);
    return Response.json({
      name,
      html_url: `https://github.com/deepseek-ai/${name}`,
      stargazers_count: projects.get(name),
    });
  };
  const resolution = resolve(
    await expandDocument(<DeepSeekSystemsDeck fetchProject={fetchProject} />),
  );
  assert.deepEqual(requests, [...projects.keys()]);
  assert.deepEqual(resolution.report.dimensions, { width: 1280, height: 720 });
  assert.ok(resolution.report.fonts.length >= 2);
  for (const fontUsage of resolution.report.fonts) {
    assert.equal(
      fontUsage.hasResource,
      true,
      `font ${fontUsage.family} (${fontUsage.weight}) must have embedded resource`,
    );
    assert.equal(
      fontUsage.missingGlyphs,
      undefined,
      `font ${fontUsage.family} (${fontUsage.weight}) has missing glyphs: ${fontUsage.missingGlyphs?.join(", ")}`,
    );
  }

  for (const slide of resolution.slides) {
    assert.deepEqual(slide.dimensions, { width: 1280, height: 720 });
    for (const object of flattenObjects(slide.objects)) {
      const bounds = object.bounds ?? object.box;
      if (bounds !== undefined) {
        assert.ok(
          bounds.x >= 0,
          `${slide.id}/${object.path} overflows left boundary (${bounds.x})`,
        );
        assert.ok(
          bounds.y >= 0,
          `${slide.id}/${object.path} overflows top boundary (${bounds.y})`,
        );
        assert.ok(
          bounds.x + bounds.width <= slide.dimensions.width,
          `${slide.id}/${object.path} overflows right boundary (${bounds.x + bounds.width})`,
        );
        assert.ok(
          bounds.y + bounds.height <= slide.dimensions.height,
          `${slide.id}/${object.path} overflows bottom boundary (${bounds.y + bounds.height})`,
        );
      }
      if (
        (object.kind === "text" || object.kind === "shape") &&
        object.text !== undefined &&
        object.text.trim() !== ""
      ) {
        assertMeasuredTextFits(
          object.textMetrics,
          `${slide.id}/${object.path} (${JSON.stringify(object.text)})`,
        );
      }
      if (object.table !== undefined) {
        for (const cell of object.table.cells) {
          if (cell.text.trim() !== "") {
            assertMeasuredTextFits(
              cell.textMetrics,
              `${slide.id}/${cell.path} (${JSON.stringify(cell.text)})`,
            );
          }
        }
      }
    }
  }

  const result = emit(resolution, {
    embedFonts: "full",
  });
  assert.equal(resolution.slides.length, 15);
  assert.equal(result.report.slides.length, 15);

  const painted = paint(resolution);
  assert.equal(painted.pages.length, 15);
  for (const page of painted.pages) {
    assert.deepEqual(page.errors, [], `preview page ${page.page} errors`);
    assert.deepEqual(
      page.placeholders,
      [],
      `preview page ${page.page} placeholders`,
    );
    assert.ok(page.png !== undefined, `preview page ${page.page} png`);
    assertValidPng1280x720(page.png, `preview page ${page.page}`);
  }

  const file = await validatePptx(result.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  assert.deepEqual(file.errors, []);
  assert.equal(file.slides.length, 15);
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
      "github-projects",
      "open-infra-synthesis",
    ],
  );

  const githubPage = file.slides.find(({ id }) => id === "github-projects")!;
  const githubChart = githubPage.objects.find(
    ({ chart }) => chart !== undefined,
  )?.chart;
  assert.deepEqual(githubChart?.series[0]?.categories, [...projects.keys()]);
  assert.deepEqual(githubChart?.series[0]?.values, [...projects.values()]);
  for (const name of projects.keys()) {
    assert.ok(
      githubPage.objects.some(
        ({ text, link }) =>
          text?.text === name &&
          link &&
          "url" in link &&
          link.url === `https://github.com/deepseek-ai/${name}`,
      ),
    );
  }

  const openingSlide = file.slides.find(({ id }) => id === "opening")!;
  const openingLocator = resolution.slide("opening");
  const kpiAuthorIds = Array.from(
    { length: 4 },
    (_, index) => `opening-kpi-${index + 1}`,
  );
  const kpiGroups = kpiAuthorIds.map((authorId) => {
    const groupId = openingLocator.objects.find(
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
  const openingMotion = motionSchedule(openingLocator);
  assert.ok(openingMotion !== undefined);
  assert.equal(openingMotion.steps.length, 1);
  assert.equal(openingMotion.steps[0]?.trigger, "auto");
  assert.deepEqual(
    openingMotion.steps[0]?.effects
      .map((effect) => [
        effect.preset,
        effect.durationMs,
        effect.start.state === "known" ? effect.start.ms : -1,
      ])
      .sort((left, right) => Number(left[2]) - Number(right[2])),
    [
      ["floatUp", 420, 0],
      ["floatUp", 420, 120],
      ["floatUp", 420, 240],
      ["floatUp", 420, 360],
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
  assert.equal(charts.length, 13);
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

  // Verify native tables exist on the three architectural comparison slides and all rows/cells are measured.
  for (const [slideId, tableId, expectedRows, expectedCols] of [
    ["training-cost", "training-cost-table", 5, 5],
    ["dualpipe", "dualpipe-table", 5, 5],
    ["pd-disaggregation", "pd-comparison-table", 6, 4],
  ] as const) {
    const resolvedTable = resolution.slide(slideId).get(tableId)
      .objects[0]?.table;
    assert.ok(
      resolvedTable !== undefined,
      `Slide ${slideId} should expose ResolvedTable for ${tableId}`,
    );
    assert.equal(resolvedTable.rows.length, expectedRows);
    assert.equal(resolvedTable.columns.length, expectedCols);
    assert.equal(resolvedTable.cells.length, expectedRows * expectedCols);

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

  // Verify intentional click animations on the 4 interactive slides via motionSchedule and emitted PPTX.
  for (const slideId of [
    "aux-loss-free-mtp",
    "dualpipe",
    "decode-pipeline",
    "r1-rl-distill",
  ]) {
    const schedule = motionSchedule(resolution.slide(slideId));
    assert.ok(
      (schedule?.steps.length ?? 0) >= 1,
      `Slide ${slideId} motionSchedule should have click-triggered steps`,
    );
    for (const step of schedule?.steps ?? []) {
      assert.equal(step.trigger, "click");
      assert.ok(step.effects.length > 0);
    }

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

  // Verify the vanilla DualPipe schedule diagram preserves exact slide-space bounds and emitted shapes.
  const dualpipeResolved = resolution.slide("dualpipe").objects;
  const expectedLabels = [
    { text: "正向数据流 (Rank 0→P)", x: 68, y: 180, width: 124 },
    { text: "镜像数据流 (Rank P→0)", x: 68, y: 226, width: 124 },
  ] as const;
  for (const expected of expectedLabels) {
    const match = dualpipeResolved.find(
      (object) => object.kind === "text" && object.text === expected.text,
    );
    assert.ok(match !== undefined, `Missing DualPipe label ${expected.text}`);
    assert.equal(match.box?.x, expected.x);
    assert.equal(match.box?.y, expected.y);
    assert.equal(match.box?.width, expected.width);
    assertMeasuredTextFits(match.textMetrics, `dualpipe/${expected.text}`);
  }

  const expectedBlocks = [
    { text: "F0", x: 196, y: 172, width: 50, height: 34 },
    { text: "F1", x: 252, y: 172, width: 50, height: 34 },
    { text: "F&B (双向重叠)", x: 308, y: 172, width: 124, height: 34 },
    { text: "F&B (双向重叠)", x: 438, y: 172, width: 124, height: 34 },
    { text: "B1", x: 568, y: 172, width: 52, height: 34 },
    { text: "W1", x: 626, y: 172, width: 48, height: 34 },
    { text: "F'0", x: 196, y: 218, width: 50, height: 34 },
    { text: "F&B (双向重叠)", x: 252, y: 218, width: 124, height: 34 },
    { text: "F&B (双向重叠)", x: 382, y: 218, width: 124, height: 34 },
    { text: "B'0", x: 512, y: 218, width: 52, height: 34 },
    { text: "B'1", x: 570, y: 218, width: 52, height: 34 },
    { text: "W'0", x: 628, y: 218, width: 46, height: 34 },
  ] as const;
  const resolvedBlocks = dualpipeResolved.filter(
    (object) =>
      object.kind === "shape" &&
      (object.box?.y === 172 || object.box?.y === 218) &&
      object.box.x >= 196,
  );
  assert.deepEqual(
    resolvedBlocks.map((object) => ({
      text: object.text,
      x: object.box?.x,
      y: object.box?.y,
      width: object.box?.width,
      height: object.box?.height,
    })),
    expectedBlocks,
  );

  const dualpipeFile = file.slides.find(({ id }) => id === "dualpipe")!;
  for (const expected of expectedBlocks) {
    assert.ok(
      dualpipeFile.objects.some(
        (object) =>
          object.text?.text === expected.text &&
          object.boxEmu?.x === expected.x * 9525 &&
          object.boxEmu?.y === expected.y * 9525 &&
          object.boxEmu?.width === expected.width * 9525 &&
          object.boxEmu?.height === expected.height * 9525,
      ),
      `Missing emitted DualPipe block ${expected.text} at (${expected.x}, ${expected.y})`,
    );
  }
}, 60_000);

it("renders the vanilla DualPipe schedule diagram cleanly without React at custom panel origins", () => {
  const standalone = deck({ theme }).append(
    slide({ id: "dualpipe-standalone" }).append(
      createDualPipeSchedule({ x: 100, y: 50 }),
    ),
  );
  const resolution = resolve(standalone);
  const objects = resolution.slide("dualpipe-standalone").objects;
  assert.equal(objects.length, 14);
  for (const object of objects) {
    assertMeasuredTextFits(
      object.textMetrics,
      `dualpipe-standalone/${object.path}`,
    );
  }
  assert.deepEqual(
    objects[0]?.box && { x: objects[0].box.x, y: objects[0].box.y },
    {
      x: 120,
      y: 122,
    },
  );
  assert.deepEqual(objects[1]?.box, { x: 248, y: 114, width: 50, height: 34 });
  assert.deepEqual(objects[13]?.box, { x: 680, y: 160, width: 46, height: 34 });
});

it("identifies the repository and HTTP status when fetching projects fails", async () => {
  await assert.rejects(
    fetchGitHubProjects(async (input) => {
      assert.ok(typeof input === "string");
      const name = new URL(input).pathname.split("/").at(-1);
      if (name === "DeepEP") return new Response(null, { status: 403 });
      return Response.json({
        name,
        html_url: `https://github.com/deepseek-ai/${name}`,
        stargazers_count: 1,
      });
    }),
    /deepseek-ai\/DeepEP.*403/u,
  );
});
