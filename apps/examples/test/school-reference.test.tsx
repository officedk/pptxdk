import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";

import type { MotionObservation } from "@react-pptx/react-pptx";
import { buildFromSource } from "@react-pptx/react-pptx/node";
import { validatePptx, type FileObject } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

function flattenObjects(objects: readonly FileObject[]): FileObject[] {
  return objects.flatMap((object) => [
    object,
    ...flattenObjects(object.members ?? []),
  ]);
}

function slideText(objects: readonly FileObject[]): string {
  return flattenObjects(objects)
    .flatMap((object) => object.text?.text ?? [])
    .join("\n");
}

describe("School 参考页面", () => {
  it("保留两页内容与相对布局，并让各任务高亮跟随任务标记起点", async () => {
    const result = await buildFromSource(
      new URL("../src/index.ts", import.meta.url),
      { exportName: "schoolReferenceDeck" },
    );
    const file = await validatePptx(result.pptx);

    assert.equal(result.report.evidence, "generation");
    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 2);
    assert.equal(result.report.slides.length, 2);

    const serviceLayouts = result.report.slides[0]?.layouts ?? [];
    const agendaLayouts = result.report.slides[1]?.layouts ?? [];
    const layout = (values: typeof serviceLayouts, testId: string) => {
      const found = values.find((item) => item.testId === testId);
      assert.ok(found, `缺少布局 ${testId}`);
      return found.box;
    };

    assert.deepEqual(
      [
        layout(serviceLayouts, "school-services-row").x,
        layout(serviceLayouts, "school-services-row").y,
        layout(serviceLayouts, "school-services-row").width,
      ],
      [61, 280, 1155],
    );
    assert.deepEqual(
      [
        layout(serviceLayouts, "school-service-first").x,
        layout(serviceLayouts, "school-service-first").width,
        layout(serviceLayouts, "school-service-second").x,
        layout(serviceLayouts, "school-service-second").width,
        layout(serviceLayouts, "school-service-third").x,
        layout(serviceLayouts, "school-service-third").width,
      ],
      [61, 385, 446, 385, 831, 385],
    );
    assert.deepEqual(
      [
        layout(agendaLayouts, "school-calendar").x,
        layout(agendaLayouts, "school-calendar").y,
        layout(agendaLayouts, "school-calendar").width,
        layout(agendaLayouts, "school-calendar").height,
        layout(agendaLayouts, "school-timeline").x,
        layout(agendaLayouts, "school-timeline").y,
        layout(agendaLayouts, "school-timeline").width,
        layout(agendaLayouts, "school-timeline").height,
      ],
      [0, 0, 640, 720, 640, 0, 640, 720],
    );
    const servicesSlide = file.slides[0];
    const agendaSlide = file.slides[1];
    assert.ok(servicesSlide);
    assert.ok(agendaSlide);

    const servicesFileObjects = flattenObjects(servicesSlide.objects);
    const serviceFileBox = (name: string) => {
      const object = servicesFileObjects.find(
        (candidate) => candidate.name === name,
      );
      assert.ok(object, `独立读取结果缺少对象 ${name}`);
      return {
        object,
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      };
    };
    const caption = serviceFileBox("school-service-first-caption");
    const title = serviceFileBox("school-service-first-title");
    const body = serviceFileBox("school-service-first-body");
    assert.ok(Math.abs(title.x - caption.x) < 0.01);
    assert.ok(Math.abs(title.width - caption.width) < 0.01);
    assert.ok(title.width >= 217.019);
    assert.equal(title.object.text?.paragraphs[0]?.alignment, "ctr");
    assert.ok(Math.abs(body.y - title.y - 42.71) < 0.01);

    const agendaFileObjects = flattenObjects(agendaSlide.objects);
    const fileBox = (name: string) => {
      const object = agendaFileObjects.find(
        (candidate) => candidate.name === name,
      );
      assert.ok(object, `独立读取结果缺少对象 ${name}`);
      return {
        x: object.boxEmu.x / 9525,
        y: object.boxEmu.y / 9525,
        width: object.boxEmu.width / 9525,
        height: object.boxEmu.height / 9525,
      };
    };
    const month = fileBox("school-calendar-month");
    assert.ok(month.width >= 212.333);
    assert.ok(Math.abs(month.x + month.width / 2 - 320) < 0.01);
    assert.equal(
      agendaFileObjects.find(
        (candidate) => candidate.name === "school-calendar-month",
      )?.text?.paragraphs[0]?.alignment,
      "ctr",
    );
    const weekdays = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"].map(
      (weekday) => fileBox(`school-weekday-${weekday}`),
    );
    const sunday = weekdays[0];
    const saturday = weekdays[6];
    const firstDay = fileBox("school-day-1-label");
    assert.ok(sunday);
    assert.ok(saturday);
    assert.ok(Math.abs(firstDay.x - 348.652) < 5);
    assert.ok(Math.abs(firstDay.y - 319.567) < 5);
    assert.ok(month.y + month.height < sunday.y);
    assert.ok(sunday.x > 60 && sunday.x < 90);
    assert.ok(sunday.y > 260 && sunday.y < 290);
    assert.ok(Math.abs(sunday.x + (7 * sunday.width) / 2 - 320) < 16);
    const calendarPanel = layout(agendaLayouts, "school-calendar");
    assert.ok(sunday.x >= calendarPanel.x);
    assert.ok(
      saturday.x + saturday.width <= calendarPanel.x + calendarPanel.width,
    );
    const lastDay = fileBox("school-day-31-label");
    assert.ok(
      lastDay.y + lastDay.height <= calendarPanel.y + calendarPanel.height,
    );
    for (const [column, weekday] of weekdays.entries()) {
      assert.ok(weekday);
      assert.ok(Math.abs(weekday.x - (sunday.x + column * sunday.width)) < 1);
      assert.ok(Math.abs(weekday.y - sunday.y) < 1);
      assert.ok(Math.abs(weekday.width - sunday.width) < 0.1);
    }
    assert.ok(Math.abs(firstDay.x - (sunday.x + 4 * sunday.width)) < 1);
    assert.ok(Math.abs(firstDay.y - (sunday.y + sunday.height)) < 1);

    const octoberFirstColumn = 4;
    for (let day = 1; day <= 31; day += 1) {
      const label = fileBox(`school-day-${day}-label`);
      const calendarOffset = octoberFirstColumn + day - 1;
      const weekday = weekdays[calendarOffset % 7];
      assert.ok(weekday);
      assert.ok(Math.abs(label.x - weekday.x) < 1);
      assert.ok(
        Math.abs(
          label.y -
            (sunday.y + sunday.height + Math.floor(calendarOffset / 7) * 42.6),
        ) < 1,
      );
    }
    for (const day of [7, 17, 19]) {
      const label = fileBox(`school-day-${day}-label`);
      const highlight = fileBox(`school-day-${day}-highlight`);
      assert.ok(
        Math.abs(
          label.x + label.width / 2 - (highlight.x + highlight.width / 2),
        ) < 1,
      );
      assert.ok(
        Math.abs(
          label.y + label.height / 2 - (highlight.y + highlight.height / 2),
        ) < 1,
      );
    }

    const servicesText = slideText(servicesSlide.objects);
    const agendaText = slideText(agendaSlide.objects);
    assert.ok(servicesText.includes("Our Services"));
    assert.ok(servicesText.includes("Your Text"));
    assert.ok(
      servicesText.includes(
        "Here You can Add Some Brief Text That will Explain Your Tittle",
      ),
    );
    assert.ok(agendaText.includes("Agenda Slide"));
    assert.ok(agendaText.includes("October"));
    assert.ok(agendaText.includes("October 7, 2020"));
    assert.ok(agendaText.includes("October 17, 2020"));
    assert.ok(agendaText.includes("October 19, 2020"));
    assert.equal(servicesText.split("Your Text").length - 1, 3);
    assert.equal(
      agendaText.split("Here You can Add Some Brief Text").length - 1,
      3,
    );
    assert.equal(
      flattenObjects([...servicesSlide.objects, ...agendaSlide.objects]).filter(
        ({ kind }) => kind === "image",
      ).length,
      0,
    );

    const serviceMotion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.equal(serviceMotion.steps.length, 3);
    assert.deepEqual(
      serviceMotion.steps.map(({ trigger }) => trigger),
      ["click", "click", "click"],
    );
    const serviceStaggerScopes = (serviceMotion.scopes ?? []).filter(
      ({ mode, eachMs, members }) =>
        mode === "stagger" && eachMs === 250 && members.length === 6,
    );
    assert.equal(serviceStaggerScopes.length, 3);
    assert.equal(new Set(serviceStaggerScopes.map(({ id }) => id)).size, 3);
    for (const [index, id] of ["first", "second", "third"].entries()) {
      const step = serviceMotion.steps[index];
      assert.ok(step);
      const cardEffects = step.effects;
      assert.deepEqual(
        cardEffects
          .filter(({ target }) =>
            target.testId?.startsWith(`school-service-${id}-tile-`),
          )
          .map(({ startMs }) => startMs),
        [0, 250, 500, 750, 1000],
      );
      const check = cardEffects.find(
        ({ target }) => target.testId === `school-service-${id}-check`,
      );
      const caption = cardEffects.find(
        ({ target }) => target.testId === `school-service-${id}-caption`,
      );
      assert.ok(check);
      assert.ok(caption);
      assert.equal(check.startMs, 1250);
      assert.equal(caption.startMs, 1750);
    }

    const motion =
      agendaSlide &&
      (result.report.slides[1]?.capabilities?.["motion"] as MotionObservation);
    assert.ok(motion);
    assert.deepEqual(
      motion.steps.map(({ trigger, number }) => [trigger, number]),
      [
        ["click", 1],
        ["click", 2],
        ["click", 3],
      ],
    );
    const taskMarkerScopes = (motion.scopes ?? []).filter(
      ({ mode, eachMs, members }) =>
        mode === "stagger" && eachMs === 250 && members.length === 2,
    );
    assert.equal(taskMarkerScopes.length, 3);
    assert.equal(new Set(taskMarkerScopes.map(({ id }) => id)).size, 3);
    assert.equal(
      new Set(taskMarkerScopes.map(({ parentId }) => parentId)).size,
      3,
    );
    for (const [index, [id, day]] of [
      ["first", 7],
      ["second", 17],
      ["third", 19],
    ].entries()) {
      const step = motion.steps[index];
      assert.ok(step);
      const effectFor = (testId: string) =>
        step.effects.find(({ target }) => target.testId === testId);
      const marker = effectFor(`school-task-${id}-ring`);
      const highlight = effectFor(`school-day-${day}-highlight`);
      const dot = effectFor(`school-task-${id}-dot`);
      const text = effectFor(`school-task-${id}-text`);
      assert.ok(marker, `缺少任务 ${id} 的标记动画`);
      assert.ok(highlight, `缺少日期 ${day} 的高亮动画`);
      assert.ok(dot);
      assert.ok(text);
      assert.equal(highlight.startMs, marker.startMs);
      assert.equal(dot.startMs, (marker.startMs ?? 0) + 250);
      assert.equal(text.startMs, dot.endMs);
      if (index > 0) {
        const connector = effectFor(`school-task-${id}-connector`);
        assert.ok(connector);
        assert.equal(connector.startMs, 0);
        assert.equal(marker.startMs, 250);
      } else {
        assert.equal(marker.startMs, 0);
      }
    }

    const servicesXml = unzipSync(result.pptx)["ppt/slides/slide1.xml"];
    const agendaXml = unzipSync(result.pptx)["ppt/slides/slide2.xml"];
    const themeXml = unzipSync(result.pptx)["ppt/theme/theme1.xml"];
    assert.ok(servicesXml);
    assert.ok(agendaXml);
    assert.ok(themeXml);
    assert.doesNotMatch(strFromU8(servicesXml), /<p:pic(?:\s|>)/u);
    assert.doesNotMatch(strFromU8(agendaXml), /<p:pic(?:\s|>)/u);
    assert.match(strFromU8(servicesXml), /typeface="Tw Cen MT"/u);
    assert.match(strFromU8(themeXml), /typeface="Montserrat"/u);
    assert.ok((agendaSlide.animations?.length ?? 0) > 0);

    const directory = join(
      import.meta.dirname,
      "../test-results/school-reference",
    );
    await mkdir(directory, { recursive: true });
    await writeFile(join(directory, "school-reference.pptx"), result.pptx);
    await writeFile(
      join(directory, "school-reference-report.json"),
      `${JSON.stringify(result.report, null, 2)}\n`,
    );
    await writeFile(
      join(directory, "school-reference-reader.json"),
      `${JSON.stringify(file, null, 2)}\n`,
    );
    await writeFile(
      join(directory, "school-reference.sha256"),
      `${createHash("sha256").update(result.pptx).digest("hex")}  school-reference.pptx\n`,
    );
  });

  it("从正式 dist 入口读取包内度量字体并生成相同两类页面", async () => {
    const result = await buildFromSource(
      new URL("../dist/index.js", import.meta.url),
      { exportName: "schoolReferenceDeck" },
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 2);
    assert.ok(
      slideText(file.slides[0]?.objects ?? []).includes("Our Services"),
    );
    assert.ok(
      slideText(file.slides[1]?.objects ?? []).includes("Agenda Slide"),
    );
  });
});
