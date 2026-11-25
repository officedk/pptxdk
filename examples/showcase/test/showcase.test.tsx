import assert from "node:assert/strict";

import { validatePptx } from "@pptxdk/validation";
import { unzipSync } from "fflate";
import { build, emit, resolve } from "pptxdk";
import { paint } from "pptxdk/preview";
import { Deck, expandDocument, Slide } from "pptxdk/react";
import { it } from "vitest";

import { Progress } from "../src/components/progress.tsx";
import { createShowcaseDocument, ShowcaseDeck } from "../src/index.ts";
import { CardsPage } from "../src/templates/cards-page.tsx";
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

it("builds and paints a byte-equivalent mixed showcase entry with a factory comparison page", async () => {
  const jsxResolution = resolve(await expandDocument(<ShowcaseDeck />));
  const factoryResolution = resolve(await createShowcaseDocument());
  const jsx = emit(jsxResolution, { embedFonts: "full" });
  const factory = emit(factoryResolution, { embedFonts: "full" });
  const jsxPreview = paint(jsxResolution);
  const factoryPreview = paint(factoryResolution);

  assert.deepEqual(jsxResolution.report.diagnostics, []);
  assert.deepEqual(factoryResolution.report.diagnostics, []);
  assert.deepEqual(jsx.report.diagnostics, []);
  assert.deepEqual(factory.report.diagnostics, []);
  assert.deepEqual(factory.report, jsx.report);

  assert.equal(jsxPreview.pages.length, 16);
  assert.equal(factoryPreview.pages.length, 16);
  for (let i = 0; i < 16; i += 1) {
    const jsxPage = jsxPreview.pages[i]!;
    const factoryPage = factoryPreview.pages[i]!;
    assert.deepEqual(jsxPage.diagnostics, [], `jsx page ${i + 1}`);
    assert.deepEqual(factoryPage.diagnostics, [], `factory page ${i + 1}`);
    assert.ok(jsxPage.png !== undefined, `jsx page ${i + 1} png`);
    assert.ok(factoryPage.png !== undefined, `factory page ${i + 1} png`);
    assertValidPng1280x720(jsxPage.png, `jsx page ${i + 1}`);
    assertValidPng1280x720(factoryPage.png, `factory page ${i + 1}`);
    assert.equal(
      Buffer.compare(factoryPage.png, jsxPage.png),
      0,
      `PNG differs on page ${i + 1}`,
    );
  }
  const jsxFile = await validatePptx(jsx.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  const factoryFile = await validatePptx(factory.pptx, {
    allowedExtensions: ["p14", "p159"],
  });
  assert.deepEqual(jsxFile.errors, []);
  assert.deepEqual(factoryFile.errors, []);
  assert.deepEqual(factoryFile.slides, jsxFile.slides);
  assert.deepEqual(
    factoryFile.slides.map(({ id }) => id),
    [
      "opening",
      "hierarchy",
      "grouping",
      "reveal",
      "comparison",
      "process",
      "progress",
      "calendar",
      "devices",
      "world",
      "gallery-contrast",
      "gallery-process",
      "gallery-hierarchy",
      "gallery-space",
      "combination",
      "closing",
    ],
  );
  const comparison = factoryFile.slides.find(({ id }) => id === "comparison")!;
  const group = comparison.objects.find(({ kind }) => kind === "group");
  assert.ok(group);
  const clickSteps = comparison.animations ?? [];
  assert.deepEqual(
    clickSteps.map(({ trigger }) => trigger),
    ["click", "click"],
  );
  const cardEffects = clickSteps[0]!.effects;
  const cardText = cardEffects.map(
    ({ targetId }) =>
      comparison.objects.find(({ id }) => id === targetId)?.text?.text,
  );
  assert.deepEqual(cardText, [
    "Lorem ipsum",
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.",
  ]);
  assert.ok(
    comparison.animationContainers?.some(({ kind }) => kind === "sequence"),
  );
  const groupEffect = clickSteps[1]!.effects.find(
    ({ targetId }) => targetId === group.id,
  );
  assert.deepEqual(
    [
      groupEffect?.kind,
      groupEffect?.preset,
      groupEffect?.durationMs,
      groupEffect?.from,
      group.members?.length,
    ],
    ["entrance", "wipe", 600, "left", 9],
  );

  const jsxParts = unzipSync(jsx.pptx);
  const factoryParts = unzipSync(factory.pptx);
  const partNames = Object.keys(jsxParts).sort();
  assert.deepEqual(Object.keys(factoryParts).sort(), partNames);
  for (const partName of partNames) {
    assert.deepEqual(
      factoryParts[partName],
      jsxParts[partName],
      `PPTX part differs: ${partName}`,
    );
  }
}, 30_000);

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
      {
        name: "Lorem ipsum",
        categories: ["A", "B"],
        values: [84, 56],
        index: 0,
        order: 0,
        groupIndex: 0,
      },
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
