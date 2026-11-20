import assert from "node:assert/strict";

import { emit, expandDocument, resolve } from "@react-pptx/react-pptx";
import { paint } from "@react-pptx/react-pptx/preview";
import { validatePptx } from "@react-pptx/validation";
import { it } from "vitest";

import { GenerativeAiHistoryDeck } from "../src/index.ts";

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

it("full deck preserves 14 slides, zero-diagnostic static preview, editable chart, and valid object boundaries", async () => {
  const resolution = resolve(await expandDocument(<GenerativeAiHistoryDeck />));
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
      "timeline",
      "prediction",
      "vae-gan",
      "transformer",
      "scale",
      "diffusion",
      "alignment",
      "systems",
      "frontier",
      "synthesis",
      "questions",
      "sources-1",
      "sources-2",
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

  const flowchartObjects = file.slides.find(
    (slide) => slide.id === "vae-gan",
  )!.objects;
  for (const label of ["编码器", "解码器", "生成器 G", "真实样本"]) {
    assert.ok(
      flowchartObjects.some(
        (object) =>
          object.text?.text === label &&
          object.geometry?.startsWith("flowChart"),
      ),
      `${label} should use native flowchart geometry and embedded text`,
    );
  }
  assert.equal(
    flowchartObjects.filter((object) => object.geometry === "rightArrow")
      .length,
    7,
  );

  const chart = file.slides
    .find((slide) => slide.id === "scale")
    ?.objects.find((object) => object.chart !== undefined)?.chart;
  assert.ok(chart);
  assert.deepEqual(chart.types, ["barChart"]);
  assert.deepEqual(chart.series, [
    {
      name: "参数量（十亿）",
      categories: ["GPT-1 · 2018", "GPT-2 · 2019", "GPT-3 · 2020"],
      values: [0.117, 1.5, 175],
      index: 0,
      order: 0,
      groupIndex: 0,
    },
  ]);
  assert.deepEqual(chart.embeddedWorkbook?.series, chart.series);
}, 60_000);
