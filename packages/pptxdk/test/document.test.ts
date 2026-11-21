import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { getEventListeners } from "node:events";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  build,
  createMorphScope,
  font,
  type ChartSeries,
  type ChartTrendline,
  type LinearGradient,
} from "pptxdk";
import {
  box,
  chart,
  deck,
  paragraph,
  slide,
  slideNumber,
  span,
  text,
} from "pptxdk/vanilla";
import { describe, it } from "vitest";

const jostBook = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

describe("public document authoring entry", () => {
  it("reads chart series getters when taking declaration values", async () => {
    let reads = 0;
    let trendlineTypeReads = 0;
    const sourceValues = [1, 2];
    const sourceTrendline = {
      get type(): "linear" {
        trendlineTypeReads += 1;
        return "linear";
      },
    } satisfies ChartTrendline;
    const sourceSeries = {
      name: "Observed",
      get values() {
        reads += 1;
        return sourceValues;
      },
      trendlines: [sourceTrendline],
    } satisfies ChartSeries;
    const node = chart({
      id: "getter-series",
      x: 0,
      y: 0,
      width: 400,
      height: 260,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [sourceSeries],
        },
      ],
    });
    const dataSeries = {
      name: "Data property",
      values: [7, 8],
    } satisfies ChartSeries;
    const dataNode = chart({
      id: "data-series",
      x: 480,
      y: 0,
      width: 400,
      height: 260,
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [dataSeries],
        },
      ],
    });
    sourceValues[0] = 9;
    sourceSeries.name = "Changed outside";
    dataSeries.values[0] = 70;

    const document = deck({ width: 960, height: 320 }).append(
      slide({ id: "getter-values" }).append(node, dataNode),
    );
    const result = await build(document);
    const firstFile = await validatePptx(result.pptx);
    const firstPptx = new Uint8Array(result.pptx);
    const firstReport = structuredClone(result.report);
    const firstQuery = structuredClone(
      result.query.getSlide("getter-values").get("getter-series").observe(),
    );

    assert.equal(reads, 1);
    assert.equal(trendlineTypeReads, 1);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.name,
      "Observed",
    );
    assert.deepEqual(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.values,
      [1, 2],
    );
    assert.equal(
      firstFile.slides[0]?.objects[0]?.chart?.series[0]?.trendlines?.[0]?.type,
      "linear",
    );
    assert.deepEqual(
      firstFile.slides[0]?.objects[1]?.chart?.series[0]?.values,
      [7, 8],
    );

    let updateReads = 0;
    const updatedValues = [3, 4];
    const updatedSeries = {
      name: "Updated",
      get values() {
        updateReads += 1;
        return updatedValues;
      },
    } satisfies ChartSeries;
    node.set({
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [updatedSeries],
        },
      ],
    });
    updatedValues[0] = 8;
    updatedSeries.name = "Changed after set";
    const updated = await build(document);
    const updatedFile = await validatePptx(updated.pptx);

    assert.equal(updateReads, 1);
    assert.deepEqual(updatedFile.errors, []);
    assert.equal(
      updatedFile.slides[0]?.objects[0]?.chart?.series[0]?.name,
      "Updated",
    );
    assert.deepEqual(
      updatedFile.slides[0]?.objects[0]?.chart?.series[0]?.values,
      [3, 4],
    );
    assert.deepEqual(result.pptx, firstPptx);
    assert.deepEqual(result.report, firstReport);
    assert.deepEqual(
      result.query.getSlide("getter-values").get("getter-series").observe(),
      firstQuery,
    );
  });

  it("takes nested getter and class style values at declaration time", async () => {
    let getterReads = 0;
    let sourceColor = "#AA0000";
    const getterStyle = {
      get color() {
        getterReads += 1;
        return sourceColor;
      },
      fontSize: 18,
    };
    class StyleConfig {
      color = "#00AA00";
      fontSize = 18;
    }
    const classStyle = new StyleConfig();
    const document = deck({ width: 480, height: 320 }).append(
      slide({ id: "style-values" }).append(
        text("Getter style", {
          id: "getter-style",
          x: 0,
          y: 0,
          width: 300,
          height: 40,
          style: getterStyle,
        }),
        text("Class style", {
          id: "class-style",
          x: 0,
          y: 50,
          width: 300,
          height: 40,
          style: classStyle,
        }),
      ),
    );
    sourceColor = "#0000FF";
    classStyle.color = "#0000FF";
    const result = await build(document);
    const file = await validatePptx(result.pptx);

    assert.equal(getterReads, 1);
    assert.deepEqual(file.errors, []);
    const getterRun = file.slides[0]?.objects.find(({ text: value }) =>
      value?.text.includes("Getter style"),
    )?.text?.paragraphs[0]?.runs[0];
    const classRun = file.slides[0]?.objects.find(({ text: value }) =>
      value?.text.includes("Class style"),
    )?.text?.paragraphs[0]?.runs[0];
    assert.equal(getterRun?.color, "#AA0000");
    assert.equal(classRun?.color, "#00AA00");
  });

  it("accepts record-shaped class props as rich-text factory configuration", async () => {
    class TextOptions {
      id = "class-props";
      x = 0;
      y = 0;
      width = 300;
      height = 40;
      style = { color: "#AA0000" };
    }

    const document = deck({ width: 480, height: 320 }).append(
      slide({ id: "class-props" }).append(
        text("Hello", new TextOptions()),
        text("Hello", {
          id: "plain-props",
          x: 0,
          y: 50,
          width: 300,
          height: 40,
          style: { color: "#AA0000" },
        }),
      ),
    );
    const result = await build(document);
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    const helloTexts = file.slides[0]?.objects.filter(
      ({ text: value }) => value?.text === "Hello",
    );
    assert.equal(helloTexts?.length, 2);
    assert.deepEqual(
      helloTexts?.map(
        ({ text: value }) => value?.paragraphs[0]?.runs[0]?.color,
      ),
      ["#AA0000", "#AA0000"],
    );
  });

  it("creates common nodes, builds them through the formal build, and keeps old build results immutable", async () => {
    const title = text("Initial title").set({
      x: 0,
      y: 0,
      width: 720,
      height: 40,
    });
    assert.equal(title.style({ fontSize: 28, color: "#123456" }), title);
    const body = text("Body copy").set({
      x: 0,
      y: 0,
      width: 720,
      height: 32,
    });
    const card = box()
      .style({
        display: "flex",
        flexDirection: "column",
        gap: 12,
        padding: 16,
        width: 800,
      })
      .append(title, body);
    const cover = slide({ id: "cover" }).append(card);
    const document = deck({ width: 960, height: 540 }).append(cover);

    const first = await build(document);
    assert.equal(first.report.slides[0]?.id, "cover");
    assert.deepEqual(
      first.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Initial title", "Body copy"],
    );
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);

    const firstSlideXml = slideXml(first);
    assert.match(firstSlideXml, /Initial title/u);

    title.content("Revised title");
    const second = await build(document);
    assert.deepEqual(
      second.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Revised title", "Body copy"],
    );
    assert.deepEqual(
      first.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Initial title", "Body copy"],
    );
    assert.doesNotMatch(firstSlideXml, /Revised title/u);
    assert.match(slideXml(second), /Revised title/u);
  });

  it("enforces single parents, explicit same-document moves, removal, and document ownership", async () => {
    const first = box({ id: "first" }).style({
      backgroundColor: "#FFFFFF",
      width: 120,
      height: 40,
    });
    const second = box({ id: "second" }).style({
      backgroundColor: "#FFFFFF",
      width: 120,
      height: 40,
    });
    const slideA = slide({ id: "a" }).append(first);
    const slideB = slide({ id: "b" }).append(second);
    const document = deck({ width: 960, height: 540 }).append(slideA, slideB);
    const otherSlide = slide({ id: "other" });
    const other = deck({ width: 960, height: 540 }).append(otherSlide);

    assert.throws(() => slideA.append(first), /already attached/u);
    assert.throws(
      () => slideB.append(first),
      /already attached|another document/u,
    );
    assert.throws(() => document.append(slideA), /already attached/u);

    slideA.remove();
    assert.throws(() => other.append(slideA), /another document/u);
    document.append(slideA);

    first.remove();
    assert.throws(() => other.append(first), /another document/u);
    assert.throws(() => first.moveTo(otherSlide), /another document/u);
    slideA.append(first);

    const moved = box({ id: "moved" }).style({
      backgroundColor: "#FFFFFF",
      width: 80,
      height: 30,
    });
    slideA.append(moved);
    assert.equal(moved.moveTo(slideB), moved);
    const movedResult = await build(document);
    const movedSlideB = movedResult.report.slides.find(({ id }) => id === "b");
    const movedSlideA = movedResult.report.slides.find(({ id }) => id === "a");
    assert.ok(
      movedSlideB?.objects.some(({ authorId: testId }) => testId === "moved"),
    );
    assert.ok(
      movedSlideA?.objects.every(({ authorId: testId }) => testId !== "moved"),
    );

    moved.moveTo(slideA, first);
    const reordered = await build(document);
    assert.deepEqual(
      reordered.report.slides
        .find(({ id }) => id === "a")
        ?.objects.map(({ authorId: testId }) => testId),
      ["moved", "first"],
    );

    const parent = box();
    const child = box();
    parent.append(child);
    assert.throws(() => parent.moveTo(child), /own descendant/u);
    assert.throws(() => moved.moveTo(slideA, box()), /direct child/u);
    moved.remove();
    assert.throws(() => moved.moveTo(otherSlide), /another document/u);
    assert.equal(slideA.append(moved), slideA);
    assert.deepEqual(
      (await build(document)).report.slides
        .find(({ id }) => id === "a")
        ?.objects.map(({ authorId: testId }) => testId),
      ["first", "moved"],
    );
  });

  it("copies style, theme, font, and byte inputs while preserving local style inheritance", async () => {
    const sourceFontBytes = new Uint8Array(jostBook);
    const themeFont = font(
      "Document Metric",
      [{ weight: 400, data: sourceFontBytes }],
      { substitute: true },
    );
    const colors = { accent1: "#112233" };
    const stops = [
      { color: "#AA0000", position: 0 },
      { color: "#00AA00", position: 1 },
    ];
    const gradient: LinearGradient = {
      type: "linear-gradient",
      angle: 0,
      stops,
    };
    const boxStyle = {
      backgroundImage: gradient,
      width: 600,
      height: 200,
      color: "#C00000",
      fontSize: 20,
    };
    const parent = box({ style: boxStyle });
    const child = text("Measured body").style({
      width: 300,
      height: "auto",
    });
    const linkTarget = { url: "https://example.test/original" };
    child.set({ link: linkTarget });
    parent.append(child);
    const document = deck({
      width: 960,
      height: 540,
      theme: { colors, fonts: { minor: themeFont } },
    }).append(slide({ id: "values" }).append(parent));

    sourceFontBytes[0] = 0;
    themeFont.faces![0]!.data![0] = 0;
    colors.accent1 = "#445566";
    stops[0]!.color = "#0000FF";
    boxStyle.color = "#0000FF";
    linkTarget.url = "https://example.test/mutated";

    const first = await build(document);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const textObject = first.report.slides[0]?.objects.find(
      ({ kind }) => kind === "text",
    );
    assert.ok(textObject !== undefined);
    assert.ok(
      (textObject.box?.height ?? 0) > 0,
      "content-sized text should use the copied font bytes",
    );
    assert.equal(textObject.link, "url:https://example.test/original");
    const firstXml = slideXml(first);
    assert.match(firstXml, /AA0000/u);
    assert.doesNotMatch(firstXml, /0000FF/u);
    assert.match(firstXml, /C00000/u);
    const themeXml = strFromU8(unzipSync(first.pptx)["ppt/theme/theme1.xml"]!);
    assert.match(themeXml, /112233/u);
    assert.doesNotMatch(themeXml, /445566/u);
    assert.match(themeXml, /Document Metric/u);

    child.style({ color: "#00AA00" });
    assert.match(slideXml(await build(document)), /00AA00/u);
    child.style({ color: undefined });
    assert.match(slideXml(await build(document)), /C00000/u);

    const replacement: LinearGradient = {
      type: "linear-gradient",
      angle: 90,
      stops: [
        { color: "#BBCCDD", position: 0 },
        { color: "#DDCCBB", position: 1 },
      ],
    };
    parent.style({ backgroundImage: replacement });
    const replaced = slideXml(await build(document));
    assert.match(replaced, /BBCCDD/u);
    assert.doesNotMatch(replaced, /AA0000/u);

    parent.style({ padding: [40, 10, 10, 10] });
    const padded = await build(document);
    const paddedText = padded.report.slides[0]?.objects.find(
      ({ kind }) => kind === "text",
    );
    parent.style({ padding: [5, 10, 10, 10] });
    const tighter = await build(document);
    const tighterText = tighter.report.slides[0]?.objects.find(
      ({ kind }) => kind === "text",
    );
    assert.ok(
      (paddedText?.box?.y ?? 0) > (tighterText?.box?.y ?? Number.MAX_VALUE),
      "padding arrays should replace rather than merge",
    );

    const exposed = document.props["theme"] as {
      readonly fonts: {
        readonly minor: {
          readonly faces?: readonly { readonly data?: Uint8Array }[];
        };
      };
    };
    const exposedBytes = exposed.fonts.minor.faces?.[0]?.data;
    assert.ok(exposedBytes !== undefined);
    exposedBytes[0] = 0;
    assert.ok(Object.isFrozen(document.props));
    assert.ok(Object.isFrozen(document.children));
    assert.deepEqual(
      (await validatePptx((await build(document)).pptx)).errors,
      [],
    );
  });

  it("isolates author style objects at creation and update like the React cleanup handoff", async () => {
    const creationStyle = { color: "#FF0000" };
    const title = text("Owned input", { style: creationStyle }).set({
      x: 0,
      y: 0,
      width: 200,
      height: 40,
    });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "style-alias" }).append(title),
    );

    // Simulates React useLayoutEffect cleanup mutating the caller's original style object.
    creationStyle.color = "#0000FF";
    const created = await build(document);
    assert.match(slideXml(created), /FF0000/u);
    assert.doesNotMatch(slideXml(created), /0000FF/u);

    const updateStyle = { color: "#00AA00" };
    title.style(updateStyle);
    updateStyle.color = "#0000FF";
    const updated = await build(document);
    assert.match(slideXml(updated), /00AA00/u);
    assert.doesNotMatch(slideXml(updated), /0000FF/u);
  });

  it("uses measured text at the build call time and keeps the cancellation contract", async () => {
    const themeFont = font(
      "Document Metric",
      [{ weight: 400, data: new Uint8Array(jostBook) }],
      { substitute: true },
    );
    const title = text("Short").style({
      width: 120,
      height: "auto",
      fontSize: 20,
    });
    const document = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: themeFont } },
    }).append(slide({ id: "timing" }).append(title));

    const firstPromise = build(document);
    title.content("A much longer heading that wraps onto several lines");
    const first = await firstPromise;
    const firstHeight = first.report.slides[0]?.objects[0]?.box?.height ?? 0;
    assert.match(slideXml(first), /Short/u);
    assert.doesNotMatch(slideXml(first), /much longer/u);

    const second = await build(document);
    const secondHeight = second.report.slides[0]?.objects[0]?.box?.height ?? 0;
    assert.ok(secondHeight > firstHeight);
    assert.match(slideXml(second), /much longer/u);

    title.content("Third revision");
    assert.match(slideXml(second), /much longer/u);
    assert.doesNotMatch(slideXml(second), /Third revision/u);
    assert.equal(first.query.getSlide("timing").observation.id, "timing");

    const controller = new AbortController();
    const reason = new Error("caller cancelled");
    controller.abort(reason);
    await assert.rejects(
      build(document, { signal: controller.signal }),
      (error: unknown) => error === reason,
    );
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);

    const late = new AbortController();
    const completed = await build(document, { signal: late.signal });
    late.abort(new Error("late cancellation"));
    assert.equal(completed.report.slides[0]?.id, "timing");
    assert.equal(getEventListeners(late.signal, "abort").length, 0);
  });

  it("builds paragraphs, spans, slide-number fields, and repeated morph handles", async () => {
    const scope = createMorphScope();
    const hero = scope.handle("hero");
    const morphStyle = {
      backgroundColor: "#FFFFFF",
      width: 80,
      height: 30,
    };
    const numbered = text(
      paragraph("One"),
      paragraph([span("Two").style({ color: "#C00000" }), slideNumber()]),
    ).set({ x: 0, y: 0, width: 400, height: 80 });
    const footer = text("Page ", slideNumber()).set({
      x: 0,
      y: 100,
      width: 200,
      height: 24,
    });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "first" }).append(box({ morphKey: hero, style: morphStyle })),
      slide({ id: "second" }).append(
        box({ morphKey: hero, style: morphStyle }),
        numbered,
        footer,
      ),
    );

    const result = await build(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.equal(
      result.report.slides[1]?.objects.find(({ kind }) => kind === "text")
        ?.text,
      "One\nTwo2",
    );
    assert.ok(
      result.report.slides[1]?.objects.some(({ text }) => text === "Page 2"),
    );
    assert.match(slideXml(result, 2), /type="slidenum"/u);
    const names = result.report.slides.flatMap(({ objects }) =>
      objects.filter(({ kind }) => kind === "shape").map(({ name }) => name),
    );
    assert.equal(names.length, 2);
    assert.equal(names[0], names[1]);
    assert.match(names[0]!, /^!!morph-\d+$/u);
  });

  it("loads and builds without importing React at runtime", () => {
    const packageRoot = fileURLToPath(new URL("../", import.meta.url));
    const script = `
import { registerHooks } from "node:module";

registerHooks({
  resolve(specifier, context, nextResolve) {
    if (
      specifier === "react" ||
      specifier === "react-reconciler" ||
      specifier.startsWith("react/") ||
      specifier === "@pptxdk/react" ||
      specifier.startsWith("@pptxdk/react/")
    ) {
      throw new Error("pure document entry loaded " + specifier);
    }
    return nextResolve(specifier, context);
  },
});

const { build } = await import("pptxdk");
const entry = await import("pptxdk/vanilla");
const png = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);
const document = entry
  .deck({ width: 960, height: 540 })
  .append(
    entry
      .slide({ id: "pure", background: "#112233" })
      .append(
        entry
          .box()
          .style({ width: 640, height: 120 })
          .append(
            entry
              .text("No React")
              .set({ x: 8, y: 8, width: 320, height: 40 }),
          ),
        entry.shape({
          id: "pure-shape",
          geometry: "hexagon",
          x: 8,
          y: 140,
          width: 120,
          height: 60,
          text: "Shape",
          style: { backgroundColor: "#336699" },
        }),
        entry.image({
          id: "pure-image",
          src: png,
          x: 140,
          y: 140,
          width: 80,
          height: 60,
          fit: "cover",
        }),
        entry.group({ id: "pure-group" }).append(
          entry.shape({
            id: "pure-group-shape",
            x: 240,
            y: 140,
            width: 60,
            height: 40,
          }),
        ),
        entry
          .table({
            id: "pure-table",
            x: 8,
            y: 220,
            width: 240,
            columns: [80, "1fr"],
            style: { fontSize: 12 },
          })
          .append(
            entry
              .row({ style: { height: 24 } })
              .append(entry.cell("Pure table"), entry.cell("42")),
          ),
        entry.chart({
  id: "pure-chart",
  x: 260,
  y: 220,
  width: 240,
  height: 140,
  groups: [{
    kind: "column",
    data: { kind: "shared-categories", categories: ["Q1"] },
    series: [{ name: "Series", values: [1] }],
  }],
}),
      ),
  );
const result = await build(document);
if (result.report.slides[0]?.id !== "pure") {
  throw new Error("pure document build did not produce the expected slide");
}
if (result.report.slides[0]?.objects.length !== 6) {
  throw new Error("pure document build did not produce native objects");
}
const objects = result.report.slides[0]?.objects;
if (objects?.[4]?.kind !== "table" || objects?.[5]?.kind !== "chart") {
  throw new Error("pure document build did not produce table and chart objects");
}
process.stdout.write("ok");
`;

    const child = spawnSync(
      process.execPath,
      ["--conditions=source", "--input-type=module", "--eval", script],
      { cwd: packageRoot, encoding: "utf8" },
    );

    assert.equal(child.status, 0, child.stderr);
    assert.equal(child.stdout, "ok");
  });
});
