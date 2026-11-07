import assert from "node:assert/strict";

import { build } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { it } from "vitest";

import { GenerativeAiHistoryDeck } from "../src/index.ts";

it("完整文稿保留十四页、可编辑图表与有效对象边界", async () => {
  const result = await build(<GenerativeAiHistoryDeck />, {
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
      assert.ok(object.boxEmu.x >= 0, `${slide.id}/${object.name} 超出左边界`);
      assert.ok(object.boxEmu.y >= 0, `${slide.id}/${object.name} 超出上边界`);
      assert.ok(
        object.boxEmu.x + object.boxEmu.width <= 1280 * 9525,
        `${slide.id}/${object.name} 超出右边界`,
      );
      assert.ok(
        object.boxEmu.y + object.boxEmu.height <= 720 * 9525,
        `${slide.id}/${object.name} 超出下边界`,
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
      `${label} 应使用原生流程图几何并内嵌文字`,
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
    },
  ]);
  assert.deepEqual(chart.embeddedWorkbook?.series, chart.series);
}, 60_000);
