import assert from "node:assert/strict";

import { build } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { it } from "vitest";

import { ShowcaseDeck } from "../src/index.ts";

it("地图六次点击累积变色，同期出现地区说明，四页画廊保留连续性与直接导航", async () => {
  const result = await build(<ShowcaseDeck />, { embedFonts: "full" });
  // 画廊的其他面板有意放在离屏画布上，不能因此忽略装不下的文字。
  assert.deepEqual(
    result.report.diagnostics.filter(({ code }) => code !== "overflow"),
    [],
  );
  assert.ok(
    result.report.diagnostics
      .filter(({ code }) => code === "overflow")
      .every(({ path }) => path.includes("gallery-")),
  );
  const file = await validatePptx(result.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  assert.deepEqual(file.errors, []);
  for (const [id, titles] of [
    ["grouping", ["分组", "对齐", "一致性"]],
    ["reveal", ["先给出中心观点", "再补充理解它的依据"]],
  ] as const) {
    const page = file.slides.find((slide) => slide.id === id)!;
    assert.deepEqual(
      page.objects.flatMap(({ text }) =>
        text && titles.some((title) => title === text.text) ? [text.text] : [],
      ),
      titles,
    );
    assert.deepEqual(
      page.animations?.map(({ effects }) =>
        effects.map(({ durationMs, delayMs }) => [durationMs, delayMs]),
      ),
      titles.map(() => [
        [280, 0],
        [350, 280],
      ]),
    );
    assert.equal(
      new Set(
        page.animations?.flatMap(({ effects }) =>
          effects.map(({ targetId }) => targetId),
        ),
      ).size,
      titles.length * 2,
    );
  }
  const calendar = file.slides.find(({ id }) => id === "calendar")!;
  assert.equal(calendar.animations?.length, 3);
  for (const [index, step] of (calendar.animations ?? []).entries()) {
    const color = step.effects.find(({ preset }) => preset === "fillColor")!;
    const caption = step.effects.find(({ preset }) => preset === "fade")!;
    const date = calendar.objects.find(({ id }) => id === color.targetId)!;
    const note = calendar.objects.find(({ id }) => id === caption.targetId)!;
    assert.equal(date.text?.text, String([4, 12, 23][index]));
    assert.equal(note.name, `event-${[4, 12, 23][index]}`);
    assert.equal(color.delayMs, caption.delayMs);
  }
  const world = file.slides.find(({ id }) => id === "world")!;
  assert.deepEqual(
    world.animations?.map(({ trigger, effects }) => [
      trigger,
      effects.filter(({ preset }) => preset === "fillColor").length,
    ]),
    [
      ["click", 123],
      ["click", 33],
      ["click", 54],
      ["click", 42],
      ["click", 119],
      ["click", 24],
    ],
  );
  const targets = new Set<number>();
  for (const step of world.animations ?? []) {
    assert.equal(
      step.effects.filter(({ preset }) => preset === "fade").length,
      1,
    );
    for (const effect of step.effects) {
      assert.equal(effect.delayMs, 0);
      if (effect.preset !== "fillColor") continue;
      assert.equal(effect.durationMs, 650);
      assert.equal(effect.fromColor, "#E4ECE8");
      assert.notEqual(effect.toColor, effect.fromColor);
      assert.ok(!targets.has(effect.targetId));
      targets.add(effect.targetId);
    }
  }
  const gallery = file.slides.filter(({ id }) => id.startsWith("gallery-"));
  assert.equal(gallery.length, 4);
  const identities = gallery[0]!.objects
    .filter(({ name }) => name.startsWith("!!"))
    .map(({ name }) => name);
  assert.equal(identities.length, 9);
  for (const page of gallery) {
    assert.equal(page.transition?.kind, "morph");
    assert.equal(page.transition?.durationMs, 1200);
    assert.deepEqual(
      page.objects
        .filter(({ name }) => name.startsWith("!!"))
        .map(({ name }) => name),
      identities,
    );
    assert.deepEqual(
      page.objects.flatMap(({ link }) =>
        link && "slideId" in link ? [link.slideId] : [],
      ),
      [
        "gallery-contrast",
        "gallery-process",
        "gallery-hierarchy",
        "gallery-space",
      ],
    );
  }
}, 30_000);
