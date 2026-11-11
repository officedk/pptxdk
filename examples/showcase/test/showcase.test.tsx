import assert from "node:assert/strict";

import { build, Deck, Slide } from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { it } from "vitest";

import { Progress } from "../src/components/progress.tsx";
import { ShowcaseDeck } from "../src/index.ts";
import { CardsPage } from "../src/templates/cards-page.tsx";
import { theme } from "../src/theme.ts";

it("map and calendar synchronize click behaviors, four-page gallery preserves continuity and direct navigation", async () => {
  const result = await build(<ShowcaseDeck />, { embedFonts: "full" });
  assert.deepEqual(result.report.diagnostics, []);
  const file = await validatePptx(result.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  assert.deepEqual(file.errors, []);

  for (const [id, titles] of [
    ["grouping", ["Lorem ipsum", "Dolor sit", "Amet consectetur"]],
    ["reveal", ["Lorem ipsum", "Dolor sit amet"]],
  ] as const) {
    const page = file.slides.find((slide) => slide.id === id)!;
    assert.equal(page.animations?.length, titles.length);
    const targets = new Set<number>();
    for (const [index, step] of (page.animations ?? []).entries()) {
      assert.equal(step.trigger, "click");
      assert.ok(step.effects.every(({ kind }) => kind === "entrance"));
      const revealedText = step.effects.map((effect) => {
        const object = page.objects.find(({ id }) => id === effect.targetId)!;
        assert.ok(!targets.has(object.id));
        targets.add(object.id);
        return object.text?.text;
      });
      assert.equal(revealedText[0], titles[index]);
      assert.equal(revealedText.length, 2);
      assert.ok(revealedText[1]?.startsWith("Lorem ipsum"));
    }
  }

  const calendar = file.slides.find(({ id }) => id === "calendar")!;
  assert.equal(calendar.animations?.length, 3);
  for (const [index, step] of (calendar.animations ?? []).entries()) {
    const color = step.effects.find(({ preset }) => preset === "fillColor")!;
    const caption = step.effects.find(({ kind }) => kind === "entrance")!;
    const date = calendar.objects.find(({ id }) => id === color.targetId)!;
    const note = calendar.objects.find(({ id }) => id === caption.targetId)!;
    const noteAuthorId = result.report.slides
      .find(({ id }) => id === "calendar")
      ?.objects.find(({ id }) => id === note.id)?.authorId;
    assert.equal(step.trigger, "click");
    assert.equal(date.text?.text, String([4, 12, 23][index]).padStart(2, "0"));
    assert.equal(noteAuthorId, `event-${[4, 12, 23][index]}`);
    assert.equal(color.delayMs, caption.delayMs);
  }

  const world = file.slides.find(({ id }) => id === "world")!;
  assert.equal(world.animations?.length, 6);
  const targets = new Set<number>();
  for (const step of world.animations ?? []) {
    assert.equal(step.trigger, "click");
    const changes = step.effects.filter(({ preset }) => preset === "fillColor");
    const captions = step.effects.filter(({ kind }) => kind === "entrance");
    assert.ok(changes.length > 0);
    assert.equal(captions.length, 1);
    assert.equal(captions[0]!.delayMs, changes[0]!.delayMs);
    for (const effect of changes) {
      assert.notEqual(effect.toColor, effect.fromColor);
      assert.ok(!targets.has(effect.targetId));
      targets.add(effect.targetId);
    }
  }

  const gallery = file.slides.filter(({ id }) => id.startsWith("gallery-"));
  assert.equal(gallery.length, 4);
  const identities = new Set(
    gallery[0]!.objects
      .filter(({ name }) => name.startsWith("!!"))
      .map(({ name }) => name),
  );
  assert.ok(identities.size >= 4);
  for (const page of gallery) {
    assert.equal(page.transition?.kind, "morph");
    assert.ok(page.transition.durationMs! > 0);
    const chart = page.objects
      .flatMap(({ members }) => members ?? [])
      .find(({ chart }) => chart !== undefined)?.chart;
    assert.ok(chart);
    assert.deepEqual(chart.types, ["barChart"]);
    assert.deepEqual(chart.series, [
      { name: "Lorem ipsum", categories: ["A", "B"], values: [84, 56] },
    ]);
    assert.deepEqual(chart.embeddedWorkbook?.series, chart.series);
    assert.deepEqual(
      new Set(
        page.objects
          .filter(({ name }) => name.startsWith("!!"))
          .map(({ name }) => name),
      ),
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

it("cards template preserves layout and independent reveal across content replacements, additions, removals, and reordering", async () => {
  const first = {
    title: "Lorem ipsum",
    body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
  };
  const second = {
    title: "Dolor sit",
    body: "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
  };
  const third = {
    title: "Amet consectetur",
    body: "Ut enim ad minim veniam, quis nostrud exercitation.",
  };
  const changed = {
    ...first,
    title: "A considerably longer title still wraps naturally",
    body: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
  };
  const variants = [
    [first, second],
    [third, changed, second],
  ];
  const result = await build(
    <Deck
      width={1280}
      height={720}
      theme={theme}
      style={{ fontFamily: "major", fontSize: 24 }}
    >
      {variants.map((items, index) => (
        <CardsPage
          key={index}
          id={`reuse-${index}`}
          title="Lorem ipsum dolor sit"
          subtitle="Lorem ipsum dolor sit amet, consectetur adipiscing elit."
          items={items}
        />
      ))}
    </Deck>,
  );
  assert.deepEqual(result.report.diagnostics, []);
  const file = await validatePptx(result.pptx);
  assert.deepEqual(file.errors, []);
  const bodyWidths: number[] = [];
  for (const [index, page] of file.slides.entries()) {
    let previousRight = 0;
    assert.equal(page.animations?.length, variants[index]!.length);
    const ids = new Set<number>();
    for (const [position, step] of (page.animations ?? []).entries()) {
      const item = variants[index]![position]!;
      const title = page.objects.find(
        ({ id }) => id === step.effects[0]?.targetId,
      )!;
      const body = page.objects.find(
        ({ id }) => id === step.effects[1]?.targetId,
      )!;
      assert.equal(title.text?.text, item.title);
      assert.equal(body.text?.text, item.body);
      assert.ok(step.effects.every(({ kind }) => kind === "entrance"));
      for (const object of [title, body]) {
        assert.ok(!ids.has(object.id));
        ids.add(object.id);
        assert.ok(object.boxEmu.x >= previousRight);
        assert.ok(object.boxEmu.y + object.boxEmu.height <= 720 * 9525);
        assert.ok(object.boxEmu.x + object.boxEmu.width <= 1280 * 9525);
      }
      assert.equal(body.boxEmu.x, title.boxEmu.x);
      assert.ok(body.boxEmu.y >= title.boxEmu.y + title.boxEmu.height);
      previousRight = title.boxEmu.x + title.boxEmu.width;
      if (position === 0) bodyWidths.push(body.boxEmu.width);
    }
  }
  assert.ok(bodyWidths[0]! > bodyWidths[1]!);
});

it("progress module receives position and width from slide while preserving internal proportions and click order", async () => {
  const frames = [
    { x: 80, y: 120, width: 1120 },
    { x: 280, y: 140, width: 560 },
  ];
  const result = await build(
    <Deck width={1280} height={720} theme={theme}>
      {frames.map((frame, index) => (
        <Slide key={index} id={`progress-${index}`}>
          <Progress {...frame} completed={3} total={4} />
        </Slide>
      ))}
    </Deck>,
  );
  assert.deepEqual(result.report.diagnostics, []);
  const file = await validatePptx(result.pptx);
  assert.deepEqual(file.errors, []);
  for (const [index, page] of file.slides.entries()) {
    const frame = frames[index]!;
    const progressId = result.report.slides[index]?.objects.find(
      ({ authorId }) => authorId === "progress-example",
    )?.id;
    const progress = page.objects.find(({ id }) => id === progressId)!;
    const bounds = progress.boxEmu;
    assert.equal(page.animations?.length, 2);
    assert.ok(Math.abs(bounds.x - frame.x * 9525) <= 9525);
    assert.ok(Math.abs(bounds.y - frame.y * 9525) <= 9525);
    assert.ok(Math.abs(bounds.width - frame.width * 9525) <= 9525);
    assert.ok(
      Math.abs(bounds.height - (frame.width / 1120) * 468 * 9525) <= 9525,
    );
  }
});
