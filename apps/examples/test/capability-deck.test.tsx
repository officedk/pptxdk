import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { mkdir, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";

import type { MotionObservation } from "@react-pptx/react-pptx";
import { buildFromSource } from "@react-pptx/react-pptx/node";
import { validatePptx, type FileObject } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import * as fontkit from "fontkit";
import { describe, it } from "vitest";

function flattenObjects(objects: readonly FileObject[]): FileObject[] {
  return objects.flatMap((object) => [
    object,
    ...flattenObjects(object.members ?? []),
  ]);
}

function fileObject(objects: readonly FileObject[], name: string): FileObject {
  const result = flattenObjects(objects).find((object) => object.name === name);
  assert.ok(result, `PPTX 读回结果缺少对象 ${name}`);
  return result;
}

function slideText(objects: readonly FileObject[]): string {
  return flattenObjects(objects)
    .flatMap((object) => object.text?.text ?? [])
    .join("\n");
}

function embeddedFonts(parts: Record<string, Uint8Array>) {
  return Object.entries(parts)
    .filter(
      ([part]) => part.startsWith("ppt/fonts/") && part.endsWith(".fntdata"),
    )
    .map(([part, bytes]) => {
      const view = new DataView(
        bytes.buffer,
        bytes.byteOffset,
        bytes.byteLength,
      );
      const payloadLength = view.getUint32(4, true);
      assert.equal(view.getUint32(8, true), 0x00020002);
      assert.ok(payloadLength > 0 && payloadLength <= bytes.byteLength);
      const payload = bytes.subarray(bytes.byteLength - payloadLength);
      const parsed = fontkit.create(Buffer.from(payload));
      assert.equal(parsed.type, "TTF");
      return {
        part,
        weight: view.getUint32(28, true),
        payload,
        parsed,
      };
    });
}

describe("原创能力册", () => {
  it("公开源码保留原生对象、实际内容、布局、链接与编排关系", async () => {
    const result = await buildFromSource(
      new URL("../src/index.ts", import.meta.url),
      { exportName: "capabilityDeck" },
    );
    const file = await validatePptx(result.pptx);
    const parts = unzipSync(result.pptx);
    const allObjects = file.slides.flatMap(({ objects }) =>
      flattenObjects(objects),
    );

    assert.equal(result.report.evidence, "generation");
    assert.deepEqual(result.report.diagnostics, []);
    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.slideCount, 12);
    assert.deepEqual(
      new Set(allObjects.map(({ kind }) => kind)),
      new Set(["shape", "image", "group", "table", "chart"]),
    );

    const picture = fileObject(
      file.slides[0]?.objects ?? [],
      "cover-observatory-image",
    );
    assert.equal(picture.kind, "image");
    assert.ok(picture.image);
    const imageBytes = parts[picture.image.part];
    assert.ok(imageBytes && imageBytes.byteLength > 10_000);

    const group = fileObject(file.slides[1]?.objects ?? [], "native-group");
    assert.equal(group.kind, "group");
    assert.deepEqual(
      group.members?.map(({ kind }) => kind),
      ["shape", "shape"],
    );

    const table = allObjects.find(({ kind }) => kind === "table");
    assert.ok(table?.table);
    assert.ok(
      table.table.merges.some(
        ({ row, column, colSpan }) =>
          row === 0 && column === 0 && colSpan === 2,
      ),
    );
    assert.ok(
      table.table.merges.some(
        ({ row, column, rowSpan }) =>
          row === 1 && column === 2 && rowSpan === 2,
      ),
    );
    assert.ok(table.table.cells.flat().includes("需求验证"));
    const tableXml = strFromU8(parts["ppt/slides/slide3.xml"]!);
    assert.match(tableXml, /gridSpan="2"/u);
    assert.match(tableXml, /rowSpan="2"/u);

    const chartTypes = allObjects
      .filter(({ kind }) => kind === "chart")
      .map(({ chart }) => chart?.types[0])
      .sort((left, right) => left!.localeCompare(right!));
    assert.deepEqual(chartTypes, [
      "areaChart",
      "barChart",
      "barChart",
      "doughnutChart",
      "lineChart",
      "pieChart",
    ]);
    assert.deepEqual(
      allObjects
        .filter(({ chart }) => chart?.types[0] === "barChart")
        .map(({ chart }) => chart?.direction)
        .sort((left, right) => left!.localeCompare(right!)),
      ["bar", "col"],
    );
    assert.ok(
      allObjects
        .filter(({ kind }) => kind === "chart")
        .every(({ chart }) => chart?.embeddedWorkbook !== undefined),
    );
    const chartXml = Object.entries(parts)
      .filter(
        ([part]) =>
          part.startsWith("ppt/charts/chart") && part.endsWith(".xml"),
      )
      .map(([, bytes]) => strFromU8(bytes))
      .join("\n");
    assert.ok(chartXml.includes("万元"));
    assert.match(
      chartXml,
      /<c:(?:pie|doughnut)Chart>[\s\S]*?<c:showPercent val="1"/u,
    );

    const list = fileObject(file.slides[5]?.objects ?? [], "native-list-text");
    assert.deepEqual(
      list.text?.paragraphs.map(({ list: item }) => item?.type),
      ["bullet", "bullet", "number", "number"],
    );
    const uppercase = fileObject(
      file.slides[5]?.objects ?? [],
      "transformed-uppercase",
    );
    assert.equal(uppercase.text?.text, "native typography");
    assert.match(strFromU8(parts["ppt/slides/slide6.xml"]!), /cap="all"/u);
    assert.ok(slideText(file.slides[5]?.objects ?? []).includes("中文排版"));

    const sourceSlideNumbers = file.slides.map(
      (slide) =>
        flattenObjects(slide.objects)
          .flatMap(({ text }) => text?.paragraphs ?? [])
          .flatMap(({ runs }) => runs)
          .filter(({ field }) => field === "slidenum").length,
    );
    assert.deepEqual(
      sourceSlideNumbers,
      Array.from({ length: 12 }, () => 1),
    );
    assert.deepEqual(
      fileObject(file.slides[0]?.objects ?? [], "cover-browse-button").link,
      { slideId: "objects" },
    );
    assert.equal(
      allObjects.filter(({ link }) => link !== undefined && "action" in link)
        .length,
      12,
    );

    const layoutItems = result.report.slides[6]?.layouts ?? [];
    const layoutFor = (testId: string) => {
      const layout = layoutItems.find((item) => item.testId === testId);
      assert.ok(layout, `缺少布局 ${testId}`);
      return layout.box;
    };
    const layoutRow = layoutFor("capability-layout-row");
    const firstCard = layoutFor("layout-card-one");
    const secondCard = layoutFor("layout-card-two");
    const thirdCard = layoutFor("layout-card-three");
    assert.deepEqual(layoutRow, { x: 64, y: 196, width: 1152, height: 366 });
    assert.ok(Math.abs(secondCard.x - firstCard.x - firstCard.width - 24) < 1);
    assert.ok(Math.abs(thirdCard.x - secondCard.x - secondCard.width - 24) < 1);
    assert.ok(Math.abs(firstCard.width - secondCard.width) < 1);
    const localThemeSlide = strFromU8(parts["ppt/slides/slide7.xml"]!);
    assert.ok(localThemeSlide.includes("C45D3F"));

    const transitions = file.slides.map((slide) => slide.transition?.kind);
    for (const kind of ["fade", "strips", "push", "flip", "morph"] as const) {
      assert.ok(transitions.includes(kind), `缺少 ${kind} 转场`);
    }
    assert.deepEqual(file.slides[7]?.transition?.kind, "flip");
    assert.deepEqual(
      [file.slides[4]?.transition?.kind, file.slides[4]?.transition?.by],
      ["morph", "word"],
    );
    assert.deepEqual(
      [file.slides[9]?.transition?.kind, file.slides[9]?.transition?.by],
      ["morph", "object"],
    );
    assert.ok(file.extensions.includes("p14"));
    assert.ok(file.extensions.includes("p159"));
    assert.match(strFromU8(parts["ppt/slides/slide8.xml"]!), /p14:flip/u);
    assert.match(strFromU8(parts["ppt/slides/slide5.xml"]!), /p159:morph/u);
    assert.match(strFromU8(parts["ppt/slides/slide10.xml"]!), /name="!!/u);

    const finiteMotion = result.report.slides[10]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const finiteEffects = finiteMotion.steps.flatMap(({ effects }) => effects);
    const bodyEffects = finiteEffects.filter(({ target }) =>
      ["finite-body-step-one", "finite-body-step-two"].includes(
        target.testId ?? "",
      ),
    );
    assert.equal(bodyEffects.length, 2);
    assert.deepEqual(
      finiteMotion.steps
        .filter(({ trigger }) => trigger === "click")
        .map(({ number }) => number),
      [1, 2],
    );
    const bodyIds = new Set(bodyEffects.map(({ target }) => target.id));
    const finiteBodyScope = finiteMotion.scopes?.find(({ members }) =>
      members.some(({ kind, id }) => kind === "object" && bodyIds.has(id)),
    );
    assert.equal(finiteBodyScope?.completion, "known");

    const breathingGroup = fileObject(
      file.slides[10]?.objects ?? [],
      "breathing-native-group",
    );
    assert.deepEqual(
      breathingGroup.members?.map(({ kind }) => kind),
      ["shape", "shape"],
    );
    const breathingIds = new Set(breathingGroup.members?.map(({ id }) => id));
    const breathingFileEffects = (file.slides[10]?.animations ?? [])
      .flatMap(({ effects }) => effects)
      .filter(({ targetId }) => breathingIds.has(targetId));
    assert.deepEqual(
      breathingFileEffects.map(({ repeatCount, completion }) => [
        repeatCount,
        completion,
      ]),
      [
        ["forever", "no-natural-completion"],
        ["forever", "no-natural-completion"],
      ],
    );

    const unknownMotion = result.report.slides[11]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    const unknownEffects = unknownMotion.steps.flatMap(
      ({ effects }) => effects,
    );
    const wordEffect = unknownEffects.find(
      ({ target }) => target.testId === "word-by-word-reveal",
    );
    const paragraphEffects = unknownEffects.filter(
      ({ target }) => target.testId === "paragraph-by-paragraph-reveal",
    );
    const paragraphEffect = paragraphEffects[0];
    const followerEffect = unknownEffects.find(
      ({ target }) => target.testId === "wait-for-all-completion",
    );
    assert.deepEqual(
      [wordEffect?.by, wordEffect?.completion],
      ["word", "unknown"],
    );
    assert.deepEqual(
      paragraphEffects.map(({ by, startMs, endMs }) => [by, startMs, endMs]),
      [
        ["paragraph", 0, 340],
        ["paragraph", 340, 680],
      ],
    );
    const allScope = unknownMotion.scopes?.find(
      ({ completion, members }) =>
        completion === "unknown" &&
        members.some(
          ({ kind, id }) => kind === "object" && id === wordEffect?.target.id,
        ) &&
        members.some(
          ({ kind, id }) =>
            kind === "object" && id === paragraphEffect?.target.id,
        ),
    );
    assert.equal(allScope?.completion, "unknown");
    assert.deepEqual(
      [followerEffect?.startAfter?.point, followerEffect?.startAfter?.offsetMs],
      ["end", 0],
    );
    const unknownSlide = file.slides[11];
    assert.ok(unknownSlide);
    const paragraphText = fileObject(
      unknownSlide.objects,
      "paragraph-by-paragraph-reveal",
    );
    const wholeText = fileObject(
      unknownSlide.objects,
      "wait-for-all-completion",
    );
    assert.ok(
      unknownSlide.animationContainers?.some(
        ({ endSync }) => endSync === "all",
      ),
    );
    const unknownFileEffects = (unknownSlide.animations ?? []).flatMap(
      ({ effects }) => effects,
    );
    assert.ok(
      unknownFileEffects.some(
        ({ by, completion }) => by === "word" && completion === "unknown",
      ),
    );
    assert.deepEqual(
      unknownFileEffects
        .filter(
          ({ nativeBuild, targetId }) =>
            nativeBuild === "paragraph" && targetId === paragraphText.id,
        )
        .map(({ targetRange, delayMs }) => [
          targetRange?.type,
          targetRange?.type === "paragraph" ? targetRange.start : undefined,
          delayMs,
        ]),
      [
        ["paragraph", 0, 0],
        ["paragraph", 1, 340],
      ],
    );
    const wholeTextEffect = unknownFileEffects.find(
      ({ targetId }) => targetId === wholeText.id,
    );
    assert.ok(wholeTextEffect);
    assert.deepEqual(
      [
        wholeTextEffect.targetRange,
        wholeTextEffect.nativeBuild,
        wholeTextEffect.nativeBuildGroupId,
        wholeTextEffect.preset,
        wholeText.text?.paragraphs.length,
      ],
      [undefined, "paragraph", 1, "fade", 1],
    );
    assert.deepEqual(
      wholeTextEffect.nativeBehaviors?.map(
        ({ behavior, targetRange, filter }) => [
          behavior,
          targetRange?.type === "paragraph" ? targetRange.start : undefined,
          filter,
        ],
      ),
      [
        ["set", 0, undefined],
        ["animEffect", 0, "fade"],
      ],
    );
    assert.match(
      strFromU8(parts["ppt/slides/slide12.xml"]!),
      /<p:rtn val="all"/u,
    );
  }, 60_000);

  it("从源码与正式 dist 链生成三种字体模式并保存同次证据", async () => {
    const outputDirectory = join(
      import.meta.dirname,
      "../test-results/capability-deck",
    );
    await rm(outputDirectory, { recursive: true, force: true });
    await mkdir(outputDirectory, { recursive: true });

    for (const [entry, mode] of [
      ["source", "default"],
      ["source", "full"],
      ["source", "subset"],
      ["dist", "default"],
      ["dist", "full"],
      ["dist", "subset"],
    ] as const) {
      const result = await buildFromSource(
        new URL(
          entry === "source" ? "../src/index.ts" : "../dist/index.js",
          import.meta.url,
        ),
        {
          exportName: "capabilityDeck",
          ...(mode === "default" ? {} : { embedFonts: mode }),
        },
      );
      const file = await validatePptx(result.pptx);
      const parts = unzipSync(result.pptx);
      const fonts = embeddedFonts(parts);
      const allObjects = file.slides.flatMap(({ objects }) =>
        flattenObjects(objects),
      );

      assert.deepEqual(result.report.diagnostics, []);
      assert.deepEqual(file.errors, []);
      assert.equal(file.presentation.slideCount, 12);
      if (mode === "default") {
        assert.equal(fonts.length, 0);
      } else {
        const jostFaces = fonts.filter(
          ({ parsed }) => parsed.familyName === "Jost*",
        );
        assert.deepEqual(
          jostFaces
            .map(({ weight }) => weight)
            .sort((left, right) => left - right),
          [400, 700],
        );
        const jostRegular = jostFaces.find(({ weight }) => weight === 400);
        const jostBold = jostFaces.find(({ weight }) => weight === 700);
        assert.ok(jostRegular && jostBold, "实际嵌入资源缺少 Jost 字重");
        for (const character of "FIELD NOTES / 2026") {
          assert.ok(
            jostRegular.parsed.hasGlyphForCodePoint(character.codePointAt(0)!),
            `Jost 常规字重缺少页脚所需字形 ${character}`,
          );
        }
        for (const character of "NATIVE TYPOGRAPHYDESIGN SYSTEM") {
          assert.ok(
            jostBold.parsed.hasGlyphForCodePoint(character.codePointAt(0)!),
            `Jost 粗体缺少实际大写文字所需字形 ${character}`,
          );
        }
        const cjk = fonts.find(
          ({ parsed }) => parsed.familyName === "ZCOOL XiaoWei",
        );
        assert.ok(cjk, "实际嵌入资源缺少 ZCOOL XiaoWei");
        const cjkText = allObjects
          .flatMap(({ text, table, chart }) => [
            ...(text?.paragraphs.flatMap(({ runs, list }) => {
              const cjkRuns = runs.filter(
                ({ fontFamily }) => fontFamily === "ZCOOL XiaoWei",
              );
              return [
                ...cjkRuns.map(({ text: runText }) => runText),
                ...(cjkRuns.length > 0 && list?.type === "bullet"
                  ? [list.character]
                  : []),
                ...(cjkRuns.length > 0 &&
                list?.type === "number" &&
                list.startAt !== undefined
                  ? [String(list.startAt)]
                  : []),
              ];
            }) ?? []),
            ...(table?.cellTexts.flatMap(({ text: cellText }) =>
              cellText.paragraphs.flatMap(({ runs }) =>
                runs
                  .filter(({ fontFamily }) => fontFamily === "ZCOOL XiaoWei")
                  .map(({ text: runText }) => runText),
              ),
            ) ?? []),
            ...(chart?.series.flatMap(({ name, categories, values }) => [
              name,
              ...categories,
              ...values.map((value) => String(value ?? "")),
            ]) ?? []),
          ])
          .join("");
        const chartXml = Object.entries(parts)
          .filter(
            ([part]) =>
              part.startsWith("ppt/charts/chart") && part.endsWith(".xml"),
          )
          .map(([, bytes]) => strFromU8(bytes))
          .join("\n");
        const chartFormatText = [...chartXml.matchAll(/formatCode="([^"]*)"/gu)]
          .map(([, format]) => format?.replaceAll("&quot;", '"') ?? "")
          .filter((format) => format !== "General")
          .map((format) => format.replace(/"([^"]*)"/gu, "$1"))
          .join("");
        const chartDisplayText = `${chartFormatText}${chartXml.includes('<c:showPercent val="1"') ? "%" : ""}`;
        const cjkCharacters = new Set(Array.from(cjkText + chartDisplayText));
        for (const character of cjkCharacters) {
          assert.ok(
            cjk.parsed.hasGlyphForCodePoint(character.codePointAt(0)!),
            `ZCOOL XiaoWei 缺少本次 PPTX 实际内容所需字形 ${character}`,
          );
        }
        assert.ok(chartXml.includes("万元"));
        if (mode === "full") {
          const source = readFileSync(
            new URL(
              "../src/assets/fonts/ZCOOLXiaoWei_400Regular.ttf",
              import.meta.url,
            ),
          );
          assert.deepEqual(cjk.payload, new Uint8Array(source));
        } else {
          assert.ok(cjk.payload.byteLength < 6_302_244);
          assert.equal(
            cjk.parsed.hasGlyphForCodePoint("Z".codePointAt(0)!),
            false,
          );
        }
      }

      const prefix = `${entry}-${mode}-capability-deck`;
      await writeFile(join(outputDirectory, `${prefix}.pptx`), result.pptx);
      await writeFile(
        join(outputDirectory, `${prefix}-report.json`),
        `${JSON.stringify(result.report, null, 2)}\n`,
      );
      await writeFile(
        join(outputDirectory, `${prefix}-reader.json`),
        `${JSON.stringify(file, null, 2)}\n`,
      );
      const digest = createHash("sha256").update(result.pptx).digest("hex");
      await writeFile(
        join(outputDirectory, `${prefix}.sha256`),
        `${digest}  ${prefix}.pptx\n`,
      );
    }
  }, 60_000);
});
