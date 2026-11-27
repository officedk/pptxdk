import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import {
  after,
  build,
  CommonContainer,
  CommonNode,
  createMorphScope,
  createMotionScopeHandle,
  defineTheme,
  emit,
  fade,
  font,
  linearGradient,
  motionSchedule,
  path,
  preset,
  radialGradient,
  resolve,
  transitions,
  type ChartGroup,
  type MotionSchedule,
} from "@officedk/pptxdk";
import { paint } from "@officedk/pptxdk/preview";
import {
  appendJSX,
  Box,
  Cell,
  Chart,
  Deck,
  expandDocument,
  Group,
  Image,
  MotionScope,
  Notes,
  P,
  Row,
  Section,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Step,
  Table,
  Text,
} from "@officedk/pptxdk/react";
import { type ChartProps } from "@officedk/pptxdk/react";
import { type ChartNode, type TextNode } from "@officedk/pptxdk/vanilla";
import { validatePptx, type FileObject } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { createContext, createElement, Fragment, useContext } from "react";
import { describe, it } from "vitest";

const jost = font("Jost*", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
      ),
    ),
  },
  {
    weight: 700,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url),
      ),
    ),
  },
]);

const panelPng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/n1sAAAAASUVORK5CYII=",
    "base64",
  ),
);

const theme = defineTheme({
  name: "Integrated acceptance",
  colors: {
    brand: "#123456",
    ink: "#162033",
    accent: "#1D4ED8",
    accent2: "#C2410C",
  },
  fonts: { major: jost, minor: jost },
});

function container(node: CommonNode): CommonContainer {
  assert.ok(node instanceof CommonContainer);
  return node;
}

function isText(node: CommonNode): node is TextNode {
  return node.kind === "text";
}

function isChart(node: CommonNode): node is ChartNode {
  return node.kind === "chart";
}

function sha256(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function slideXml(
  result: { readonly pptx: Uint8Array },
  index: number,
): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function readerObject(
  read: Awaited<ReturnType<typeof validatePptx>>,
  result: Awaited<ReturnType<typeof build>>,
  slideIndex: number,
  authorId: string,
): FileObject {
  const reportObject = result.report.slides[slideIndex]?.objects.find(
    (object) => object.authorId === authorId,
  );
  assert.ok(reportObject, `missing report object ${authorId}`);
  const object = read.slides[slideIndex]?.objects.find(
    (candidate) => candidate.id === reportObject.id,
  );
  assert.ok(object, `missing reader object ${authorId}`);
  return object;
}

function motion(
  result: Awaited<ReturnType<typeof build>>,
  slideIndex = 0,
): MotionSchedule {
  const slide = result.slides[slideIndex];
  assert.ok(slide);
  const value = motionSchedule(slide);
  assert.ok(value);
  return value;
}

function chartEffects(
  schedule: MotionSchedule,
  authorId: string,
): readonly MotionSchedule["steps"][number]["effects"][number][] {
  return schedule.steps
    .flatMap((step) => step.effects)
    .filter(({ target }) => target.authorId === authorId);
}

function chartXml(pptx: Uint8Array, marker: string): string {
  const parts = unzipSync(pptx);
  for (const [path, bytes] of Object.entries(parts)) {
    if (!/^ppt\/charts\/chart\d+\.xml$/u.test(path)) continue;
    const xml = strFromU8(bytes);
    if (xml.includes(marker)) return xml;
  }
  assert.fail(`missing chart XML containing ${marker}`);
}

describe("integrated acceptance", () => {
  it("edits, appends and moves React-authored nodes while earlier builds stay fixed", async () => {
    const heroKey = createMorphScope().handle("hero");
    const stageScope = createMotionScopeHandle();
    const heroStyle = { color: "brand", fontSize: 18, fontFamily: "minor" };
    const panelBytes = new Uint8Array(panelPng);
    const CaptionContext = createContext("default");

    function Caption({
      style,
    }: {
      readonly style: {
        readonly color: string;
        readonly fontSize: number;
        readonly fontFamily: string;
      };
    }) {
      return createElement(
        Text,
        {
          id: "appended-caption",
          style: { ...style, width: 300, height: 30 },
        },
        useContext(CaptionContext),
      );
    }

    const expanded = await expandDocument(
      <Deck width={960} height={540} theme={theme}>
        <Slide id="story">
          <Step n={1}>
            <Box
              id="source-panel"
              theme={{ colors: { brand: "#123456" } }}
              style={{
                position: "absolute",
                left: 40,
                top: 120,
                width: 260,
                height: 160,
                padding: 10,
                display: "flex",
                flexDirection: "column",
              }}
            >
              <Shape
                id="hero"
                morphKey={heroKey}
                enter={fade({ duration: 120 })}
                style={{
                  ...heroStyle,
                  width: 180,
                  height: 64,
                  background: "#336699",
                }}
              >
                Intro
              </Shape>
            </Box>
          </Step>
          <Step n={2}>
            <MotionScope scope={stageScope}>
              <Box
                id="stage"
                theme={{ colors: { brand: "#ABCDEF" } }}
                style={{
                  position: "absolute",
                  left: 360,
                  top: 80,
                  width: 420,
                  height: 260,
                  padding: 20,
                  display: "flex",
                  flexDirection: "column",
                  gap: 12,
                }}
              >
                <Text
                  id="stage-title"
                  style={{
                    width: 380,
                    height: 32,
                    color: "brand",
                    fontSize: 22,
                    fontFamily: "major",
                  }}
                >
                  Stage
                </Text>
              </Box>
            </MotionScope>
          </Step>
          <Text
            id="footer"
            style={{
              position: "absolute",
              left: 40,
              top: 480,
              width: 400,
              height: 24,
              color: "ink",
              fontSize: 12,
              fontFamily: "minor",
            }}
          >
            Footer
          </Text>
        </Slide>
        <Slide id="continuation">
          <Shape
            id="hero-next"
            morphKey={heroKey}
            style={{
              position: "absolute",
              left: 60,
              top: 60,
              width: 200,
              height: 70,
              background: "#336699",
            }}
          />
        </Slide>
      </Deck>,
    );

    const first = await build(expanded);
    const firstReportSnapshot = structuredClone(first.report);
    const firstBytes = new Uint8Array(first.pptx);
    const firstSha = sha256(first.pptx);
    const firstRead = await validatePptx(first.pptx);
    assert.deepEqual(firstRead.errors, []);
    const firstHero = readerObject(firstRead, first, 0, "hero");
    const firstContinuation = readerObject(firstRead, first, 1, "hero-next");
    const firstHeroReport = first.report.slides[0]?.objects.find(
      (object) => object.authorId === "hero",
    );
    assert.deepEqual(firstHeroReport?.box, {
      x: 50,
      y: 130,
      width: 180,
      height: 64,
    });
    assert.ok(firstHero.name.startsWith("!!"));
    assert.equal(firstHero.name, firstContinuation.name);
    assert.equal(firstHero.text?.paragraphs[0]?.runs[0]?.color, "#123456");
    assert.equal(firstHero.text?.text, "Intro");
    assert.deepEqual(
      motion(first).steps.map((step) => [
        step.number,
        step.trigger,
        step.effects.map(({ target }) => target.authorId),
      ]),
      [
        [1, "click", ["hero"]],
        [2, "click", []],
      ],
    );
    const firstHeroLocator = first
      .slide("story")
      .get("source-panel")
      .get("hero");
    const firstHeroBoxSnapshot = structuredClone(firstHeroLocator.box);
    assert.deepEqual(firstHeroBoxSnapshot, {
      x: 50,
      y: 130,
      width: 180,
      height: 64,
    });
    const firstQueryText = first
      .slide("story")
      .get("stage")
      .get("stage-title").text;
    assert.equal(firstQueryText, "Stage");

    const page = expanded.query.get("story");
    const stage = page.get("stage");
    const hero = page.get("source-panel").get("hero").node;
    hero.moveTo(stage.node);
    const stageTitle = page.get("stage").getAs("stage-title", isText);
    stageTitle.content("Stage updated");

    const draftStyle = { color: "brand", fontSize: 18, fontFamily: "minor" };
    const appended = await appendJSX(
      container(stage.node),
      createElement(
        Fragment,
        null,
        createElement(
          CaptionContext.Provider,
          { value: "Context caption" },
          createElement(Caption, { style: draftStyle }),
        ),
        createElement(Image, {
          id: "panel",
          src: panelBytes,
          style: {
            position: "absolute",
            left: 220,
            top: 20,
            width: 160,
            height: 80,
          },
        }),
      ),
    );
    assert.deepEqual(
      appended.map((node) => node.id),
      ["appended-caption", "panel"],
    );
    const caption = stage.getAs("appended-caption", isText);
    panelBytes[0] = 0;
    draftStyle.color = "#FF0000";
    draftStyle.fontSize = 42;
    caption.style({ fontSize: 24 });

    const second = await build(expanded);
    const secondRead = await validatePptx(second.pptx);
    assert.deepEqual(secondRead.errors, []);
    assert.deepEqual(first.report, firstReportSnapshot);
    assert.equal(sha256(first.pptx), firstSha);
    assert.deepEqual(new Uint8Array(first.pptx), firstBytes);
    const rereadFirst = await validatePptx(first.pptx);
    assert.deepEqual(
      rereadFirst.slides[0]?.objects.find(
        (object) => object.id === firstHero.id,
      )?.boxEmu,
      firstHero.boxEmu,
    );
    assert.deepEqual(
      first.slide("story").get("source-panel").get("hero").box,
      firstHeroBoxSnapshot,
    );
    assert.deepEqual(
      motion(first).steps.map((step) => [
        step.number,
        step.effects.map(({ target }) => target.authorId),
      ]),
      [
        [1, ["hero"]],
        [2, []],
      ],
    );
    assert.equal(
      first.slide("story").get("stage").get("stage-title").text,
      "Stage",
    );
    assert.match(slideXml(first, 1), /123456/u);
    assert.doesNotMatch(slideXml(first, 1), /Stage updated/u);

    const secondHero = readerObject(secondRead, second, 0, "hero");
    const secondContinuation = readerObject(secondRead, second, 1, "hero-next");
    const secondHeroReport = second.report.slides[0]?.objects.find(
      (object) => object.authorId === "hero",
    );
    assert.deepEqual(secondHeroReport?.box, {
      x: 380,
      y: 144,
      width: 180,
      height: 64,
    });
    assert.equal(secondHero.name, firstHero.name);
    assert.equal(secondHero.name, secondContinuation.name);
    assert.equal(secondHero.text?.paragraphs[0]?.runs[0]?.color, "#ABCDEF");
    assert.equal(secondHero.text?.text, "Intro");

    const secondCaption = readerObject(
      secondRead,
      second,
      0,
      "appended-caption",
    );
    assert.equal(secondCaption.text?.text, "Context caption");
    assert.equal(secondCaption.text?.paragraphs[0]?.runs[0]?.color, "#ABCDEF");
    assert.equal(secondCaption.text?.paragraphs[0]?.runs[0]?.fontSizePt, 18);
    assert.doesNotMatch(slideXml(second, 1), /FF0000/u);
    assert.equal(
      readerObject(secondRead, second, 0, "stage-title").text?.text,
      "Stage updated",
    );
    assert.match(slideXml(second, 1), /ABCDEF/u);

    const secondImage = readerObject(secondRead, second, 0, "panel");
    assert.ok(secondImage.image?.part);
    assert.equal(
      sha256(unzipSync(second.pptx)[secondImage.image.part]!),
      sha256(panelPng),
    );
    assert.deepEqual(secondImage.boxEmu, {
      x: 580 * 9525,
      y: 100 * 9525,
      width: 160 * 9525,
      height: 80 * 9525,
    });

    const secondMotion = motion(second);
    assert.deepEqual(
      secondMotion.steps.map((step) => [
        step.number,
        step.effects.map(({ target }) => target.authorId),
      ]),
      [[2, ["hero"]]],
    );
    const secondHeroId = secondHeroReport?.id;
    assert.ok(secondHeroId !== undefined);
    const scope = secondMotion.steps
      .flatMap((step) => step.scopes)
      .find(({ members }) =>
        members.some(
          (member) =>
            member.kind === "effect" && member.ref.targetId === secondHeroId,
        ),
      );
    assert.ok(scope);
    assert.equal(scope.members.length, 1);
    assert.equal(scope.members[0]?.kind, "effect");
    if (scope.members[0]?.kind === "effect") {
      assert.equal(scope.members[0].ref.targetId, secondHeroId);
    }
  });

  it("keeps dual-axis analysis, notes, reflow and independent rings together", async () => {
    const categories = ["North", "South", "East"];
    const actualValues = [12, 28, 19];
    const customPlus = [0.4, 0.5, 0.6];
    const customMinus = [0.1, 0.2, 0.3];
    const updatedCategories = ["North", "South", "East"];
    const panelBytes = new Uint8Array(panelPng);
    const badgeKey = createMorphScope().handle("badge");

    function comboGroups(
      values: readonly number[],
      labels: readonly string[] = categories,
      plus: readonly number[] = customPlus,
      minus: readonly number[] = customMinus,
    ): readonly ChartGroup[] {
      return [
        {
          kind: "column",
          axes: { category: "period", value: "amount" },
          data: { kind: "shared-categories", categories: labels },
          series: [
            {
              name: "Actual",
              values,
              errorBars: {
                type: "custom",
                plus,
                minus,
                line: { color: "#446688", width: 1.5, dash: "dash" },
              },
            },
            { name: "Plan", values: [10, 25, 22] },
          ],
          dataLabels: { value: true, position: "outsideEnd" },
        },
        {
          kind: "line",
          axes: { category: "period", value: "share" },
          data: { kind: "shared-categories", categories: labels },
          series: [
            {
              name: "Trend",
              values: [22, 31, 26],
              trendlines: [
                {
                  type: "linear",
                  displayEquation: true,
                  displayRSquared: true,
                },
              ],
              line: { color: "#C2410C", width: 2 },
            },
          ],
        },
      ];
    }

    const comboBase: Omit<ChartProps, "groups"> = {
      title: "Revenue and trend",
      legend: "bottom",
      style: {
        width: 552,
        height: 240,
        fontFamily: "minor",
        fontSize: "10pt",
        color: "ink",
      },
      axes: {
        period: { kind: "category", title: "Region" },
        amount: {
          kind: "value",
          position: "left",
          min: 0,
          max: 50,
          majorUnit: 10,
          title: "Revenue",
        },
        share: {
          kind: "value",
          position: "right",
          min: 0,
          max: 50,
          majorUnit: 10,
          title: "Trend index",
        },
      },
    };

    const firstBody =
      "A compact revenue narrative that wraps inside a flexible column and moves the chart below it. ";
    const secondBody =
      "A longer narrative that should wrap into additional lines inside the same flexible column so the chart below must move down without leaving the canvas. ".repeat(
        2,
      );

    const expanded = await expandDocument(
      <Deck width={960} height={540} theme={theme}>
        <Slide
          id="combo-page"
          theme={{ colors: { accent: "#C2410C" } }}
          style={{ fontSize: 40, lineHeight: "60pt" }}
        >
          <Text
            id="title"
            style={{
              fontFamily: "major",
              fontSize: 28,
              fontWeight: 700,
              color: "ink",
            }}
          >
            Integrated delivery
          </Text>
          <Text
            id="body"
            style={{
              fontFamily: "minor",
              fontSize: 16,
              lineHeight: "22pt",
              color: "ink",
              height: "auto",
              width: 552,
            }}
          >
            {firstBody}
          </Text>
          <Box
            id="badge-slot"
            style={{
              position: "absolute",
              left: 64,
              top: 470,
              width: 120,
              height: 40,
            }}
          >
            <Shape
              id="badge"
              morphKey={badgeKey}
              style={{
                width: 100,
                height: 32,
                background: "#1D4ED8",
                color: "#FFFFFF",
                fontFamily: "minor",
                fontSize: 14,
              }}
            >
              Move me
            </Shape>
          </Box>
          <Box
            id="badge-target"
            style={{
              position: "absolute",
              left: 200,
              top: 470,
              width: 120,
              height: 40,
            }}
          />
          <Notes id="speech" style={{ fontSize: "14pt", color: "accent" }}>
            <P id="opening">
              Opening <Span style={{ fontWeight: 700 }}>emphasis</Span>
            </P>
            <P id="second" style={{ fontSize: "16pt" }}>
              Second paragraph
            </P>
          </Notes>
          <Chart
            id="combo-chart"
            {...comboBase}
            groups={comboGroups(actualValues)}
          />
        </Slide>
        <Slide id="rings-page">
          <Chart
            id="rings-chart"
            title="Independent rings"
            legend="right"
            style={{
              position: "absolute",
              left: 0,
              top: 0,
              width: 520,
              height: 360,
              fontFamily: "minor",
              fontSize: "10pt",
              color: "ink",
            }}
            groups={[
              {
                kind: "doughnut",
                data: { kind: "independent-categories" },
                holeSize: 55,
                series: [
                  {
                    name: "Ring A",
                    categories: ["Direct", "Partner", "Online"],
                    values: [30, 20, null],
                    points: {
                      0: { color: "#C8102E", explosion: 6 },
                      2: { color: "#1D4E89" },
                    },
                  },
                  {
                    name: "Ring B",
                    categories: ["Partner", "Direct"],
                    values: [40, 10],
                    points: { 0: { color: "#1565C0" } },
                  },
                  {
                    name: "Ring C",
                    categories: ["Direct", "Partner", "Online", "Partner"],
                    values: [5, 6, 7, 8],
                    points: {
                      3: {
                        color: "#00838F",
                        dataLabels: { numberFormat: '0.0"K"' },
                      },
                    },
                  },
                ],
                dataLabels: {
                  value: true,
                  category: true,
                  leaderLines: {
                    color: "#445566",
                    width: 1.25,
                    dash: "dash",
                  },
                },
              },
            ]}
          />
          <Image
            id="panel"
            src={panelBytes}
            style={{
              position: "absolute",
              left: 560,
              top: 40,
              width: 320,
              height: 160,
            }}
          />
        </Slide>
      </Deck>,
    );

    categories[0] = "Mutated";
    actualValues[0] = 99;
    customPlus[0] = 99;
    customMinus[0] = 99;
    panelBytes[0] = 0;

    const first = await build(expanded);
    const firstReportSnapshot = structuredClone(first.report);
    const firstBytes = new Uint8Array(first.pptx);
    const firstSha = sha256(first.pptx);
    const firstRead = await validatePptx(first.pptx);
    assert.deepEqual(firstRead.errors, []);
    assert.ok(first.report.extensions.includes("c15"));

    const firstComboObject = readerObject(firstRead, first, 0, "combo-chart");
    const firstCombo = firstComboObject.chart;
    assert.ok(firstCombo);
    assert.deepEqual(firstCombo.types, ["barChart", "lineChart"]);
    assert.deepEqual(
      firstCombo.axes.map(({ kind, position }) => ({ kind, position })),
      [
        { kind: "category", position: "bottom" },
        { kind: "value", position: "left" },
        { kind: "value", position: "right" },
      ],
    );
    assert.deepEqual(
      firstCombo.series.map(
        ({ name, index, order, groupIndex, values, categories: domain }) => ({
          name,
          index,
          order,
          groupIndex,
          values,
          domain,
        }),
      ),
      [
        {
          name: "Actual",
          index: 0,
          order: 0,
          groupIndex: 0,
          values: [12, 28, 19],
          domain: ["North", "South", "East"],
        },
        {
          name: "Plan",
          index: 1,
          order: 1,
          groupIndex: 0,
          values: [10, 25, 22],
          domain: ["North", "South", "East"],
        },
        {
          name: "Trend",
          index: 2,
          order: 2,
          groupIndex: 1,
          values: [22, 31, 26],
          domain: ["North", "South", "East"],
        },
      ],
    );
    const firstActual = firstCombo.series[0];
    assert.deepEqual(
      firstActual?.errorBars?.map(({ type, plus, minus }) => ({
        type,
        plus: plus?.values,
        minus: minus?.values,
        plusFormula: plus?.formula,
        minusFormula: minus?.formula,
      })),
      [
        {
          type: "custom",
          plus: [0.4, 0.5, 0.6],
          minus: [0.1, 0.2, 0.3],
          plusFormula: "Sheet1!$E$2:$E$4",
          minusFormula: "Sheet1!$F$2:$F$4",
        },
      ],
    );
    assert.deepEqual(firstCombo.series[2]?.trendlines, [
      {
        type: "linear",
        displayEquation: true,
        displayRSquared: true,
      },
    ]);
    assert.deepEqual(
      firstCombo.embeddedWorkbook?.series[0]?.errorBars?.[0]?.plus?.values,
      [0.4, 0.5, 0.6],
    );
    const firstComboXml = chartXml(first.pptx, "Actual");
    assert.match(firstComboXml, /<c:errBars>/u);
    assert.match(firstComboXml, /<c:trendline>/u);
    assert.match(firstComboXml, /<c:crosses val="min"\/>/u);
    assert.match(firstComboXml, /<c:crosses val="max"\/>/u);

    const firstNotes = firstRead.slides[0]?.notes;
    assert.ok(firstNotes);
    assert.equal(firstNotes.text.text, "Opening emphasis\nSecond paragraph");
    assert.deepEqual(
      firstNotes.text.paragraphs[0]?.runs.map(
        ({ text, fontSizePt, color, bold }) => ({
          text,
          fontSizePt,
          color,
          bold,
        }),
      ),
      [
        {
          text: "Opening ",
          fontSizePt: 14,
          color: "#C2410C",
          bold: false,
        },
        {
          text: "emphasis",
          fontSizePt: 14,
          color: "#C2410C",
          bold: true,
        },
      ],
    );
    assert.equal(firstNotes.text.paragraphs[0]?.lineSpacingPoints, undefined);
    assert.equal(firstNotes.text.paragraphs[1]?.runs[0]?.fontSizePt, 16);
    assert.equal(firstRead.slides[1]?.notes, undefined);
    assert.ok(
      first.report.slides[0]?.objects.every(
        ({ authorId }) => authorId !== "speech",
      ),
    );
    assert.deepEqual(first.report.slides[0]?.capabilities?.["notes"], {
      text: "Opening emphasis\nSecond paragraph",
      paragraphCount: 2,
    });

    const firstBodyBox = first.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "body",
    )?.box;
    const firstComboBox = first.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "combo-chart",
    )?.box;
    assert.ok(firstBodyBox && firstComboBox);
    assert.ok(firstBodyBox.height > 40);
    assert.ok(
      Math.abs(firstComboBox.y - (firstBodyBox.y + firstBodyBox.height)) < 1e-6,
    );
    assert.ok(firstComboBox.y + firstComboBox.height <= 540);
    const firstBadge = readerObject(firstRead, first, 0, "badge");
    const firstBadgeReport = first.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "badge",
    );
    assert.deepEqual(firstBadgeReport?.box, {
      x: 64,
      y: 470,
      width: 100,
      height: 32,
    });
    assert.doesNotMatch(slideXml(first, 1), /Opening emphasis/u);

    const firstRings = readerObject(firstRead, first, 1, "rings-chart").chart;
    assert.ok(firstRings);
    assert.deepEqual(firstRings.extensions, ["c15"]);
    assert.deepEqual(
      firstRings.series.map(
        ({ name, index, order, categories: domain, values }) => ({
          name,
          index,
          order,
          domain,
          values,
        }),
      ),
      [
        {
          name: "Ring A",
          index: 0,
          order: 0,
          domain: ["Direct", "Partner", "Online"],
          values: [30, 20, null],
        },
        {
          name: "Ring B",
          index: 1,
          order: 1,
          domain: ["Partner", "Direct"],
          values: [40, 10],
        },
        {
          name: "Ring C",
          index: 2,
          order: 2,
          domain: ["Direct", "Partner", "Online", "Partner"],
          values: [5, 6, 7, 8],
        },
      ],
    );
    for (const [index, formula] of [
      [0, "Sheet1!$A$2:$A$4"],
      [1, "Sheet1!$C$2:$C$3"],
      [2, "Sheet1!$E$2:$E$5"],
    ] as const) {
      assert.equal(firstRings.series[index]?.datalabelsRange?.formula, formula);
      assert.equal(
        firstRings.series[index]?.datalabelsRange?.source,
        "series-category",
      );
      assert.equal(firstRings.series[index]?.showDataLabelsRange, true);
    }
    assert.equal(firstRings.series[0]?.pointExplosions?.[0], 6);
    assert.deepEqual(
      firstRings.seriesDataLabels?.map(({ dataLabels }) =>
        typeof dataLabels === "undefined" || dataLabels === false
          ? undefined
          : {
              value: dataLabels.value,
              leaderLines: dataLabels.leaderLines,
              leaderLine: dataLabels.leaderLine,
            },
      ),
      [
        {
          value: true,
          leaderLines: true,
          leaderLine: {
            color: "#445566",
            widthEmu: 11_906,
            dash: "dash",
          },
        },
        {
          value: true,
          leaderLines: true,
          leaderLine: {
            color: "#445566",
            widthEmu: 11_906,
            dash: "dash",
          },
        },
        {
          value: true,
          leaderLines: true,
          leaderLine: {
            color: "#445566",
            widthEmu: 11_906,
            dash: "dash",
          },
        },
      ],
    );
    const pointLabel = firstRings.seriesDataLabels?.[2]?.pointDataLabels?.[3];
    assert.ok(pointLabel !== undefined && pointLabel !== false);
    assert.equal(pointLabel.numberFormat, '0.0"K"');
    assert.equal(pointLabel.value, true);
    const ringXml = chartXml(first.pptx, "Ring A");
    assert.equal((ringXml.match(/<c15:datalabelsRange>/gu) ?? []).length, 3);
    assert.match(ringXml, /<c:leaderLines>/u);
    for (const color of ["C8102E", "1D4E89", "1565C0", "00838F"]) {
      assert.match(ringXml, new RegExp(`<a:srgbClr val="${color}"`, "u"));
    }
    assert.match(ringXml, /<c:explosion val="6"\/>/u);

    const firstImage = readerObject(firstRead, first, 1, "panel");
    assert.ok(firstImage.image?.part);
    assert.equal(
      sha256(unzipSync(first.pptx)[firstImage.image.part]!),
      sha256(panelPng),
    );
    const firstPanelBox = first.slide("rings-page").get("panel").box;
    assert.deepEqual(firstPanelBox, {
      x: 560,
      y: 40,
      width: 320,
      height: 160,
    });

    const p = expanded.query.get("combo-page");
    const body = p.getAs("body", isText);
    const combo = p.getAs("combo-chart", isChart);
    const badge = p.get("badge-slot").get("badge").node;
    badge.moveTo(container(p.get("badge-target").node));
    body.content(secondBody);
    combo.set({
      groups: comboGroups(
        [15, 30, 21],
        updatedCategories,
        [0.5, 0.6, 0.7],
        [0.2, 0.3, 0.4],
      ),
    });

    const second = await build(expanded);
    const secondRead = await validatePptx(second.pptx);
    assert.deepEqual(secondRead.errors, []);
    assert.deepEqual(first.report, firstReportSnapshot);
    assert.equal(sha256(first.pptx), firstSha);
    assert.deepEqual(new Uint8Array(first.pptx), firstBytes);
    assert.deepEqual(first.slide("combo-page").get("body").box, firstBodyBox);
    assert.deepEqual(
      first.slide("combo-page").get("combo-chart").box,
      firstComboBox,
    );
    const rereadFirst = await validatePptx(first.pptx);
    assert.deepEqual(
      rereadFirst.slides[1]?.objects.find(({ id }) => id === firstImage.id)
        ?.boxEmu,
      firstImage.boxEmu,
    );
    const rereadCombo = rereadFirst.slides[0]?.objects.find(
      ({ id }) => id === firstComboObject.id,
    )?.chart;
    assert.deepEqual(rereadCombo?.series[0]?.values, [12, 28, 19]);
    assert.deepEqual(
      rereadCombo?.series[0]?.errorBars?.[0]?.plus?.values,
      [0.4, 0.5, 0.6],
    );
    assert.equal(
      rereadFirst.slides[0]?.notes?.text.text,
      "Opening emphasis\nSecond paragraph",
    );

    const secondCombo = readerObject(
      secondRead,
      second,
      0,
      "combo-chart",
    ).chart;
    assert.ok(secondCombo);
    assert.deepEqual(
      secondCombo.series.map(
        ({ name, index, order, values, categories: domain }) => ({
          name,
          index,
          order,
          values,
          domain,
        }),
      ),
      [
        {
          name: "Actual",
          index: 0,
          order: 0,
          values: [15, 30, 21],
          domain: ["North", "South", "East"],
        },
        {
          name: "Plan",
          index: 1,
          order: 1,
          values: [10, 25, 22],
          domain: ["North", "South", "East"],
        },
        {
          name: "Trend",
          index: 2,
          order: 2,
          values: [22, 31, 26],
          domain: ["North", "South", "East"],
        },
      ],
    );
    assert.deepEqual(
      secondCombo.series[0]?.errorBars?.map(({ plus, minus }) => ({
        plus: plus?.values,
        minus: minus?.values,
      })),
      [{ plus: [0.5, 0.6, 0.7], minus: [0.2, 0.3, 0.4] }],
    );
    assert.deepEqual(secondCombo.series[2]?.trendlines, [
      {
        type: "linear",
        displayEquation: true,
        displayRSquared: true,
      },
    ]);
    assert.equal(secondCombo.axes.length, 3);
    const secondBadge = readerObject(secondRead, second, 0, "badge");
    const secondBadgeReport = second.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "badge",
    );
    assert.deepEqual(secondBadgeReport?.box, {
      x: 200,
      y: 470,
      width: 100,
      height: 32,
    });
    assert.equal(secondBadge.name, firstBadge.name);
    assert.deepEqual(
      first.slide("combo-page").get("badge-slot").get("badge").box,
      firstBadgeReport?.box,
    );

    const secondBodyBox = second.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "body",
    )?.box;
    const secondComboBox = second.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "combo-chart",
    )?.box;
    assert.ok(secondBodyBox && secondComboBox);
    assert.ok(
      secondBodyBox.height > firstBodyBox.height,
      `${firstBodyBox.height} -> ${secondBodyBox.height}`,
    );
    assert.ok(
      secondComboBox.y > firstComboBox.y,
      `${firstComboBox.y} -> ${secondComboBox.y}`,
    );
    assert.ok(
      Math.abs(secondComboBox.y - (secondBodyBox.y + secondBodyBox.height)) <
        1e-6,
    );
    assert.ok(secondComboBox.y + secondComboBox.height <= 540);

    assert.deepEqual(
      secondRead.slides[0]?.notes?.text.text,
      "Opening emphasis\nSecond paragraph",
    );
    assert.deepEqual(secondCombo.series[0]?.values, [15, 30, 21]);
    const secondRings = readerObject(
      secondRead,
      second,
      1,
      "rings-chart",
    ).chart;
    assert.deepEqual(
      secondRings?.series.map(({ values }) => values),
      [
        [30, 20, null],
        [40, 10],
        [5, 6, 7, 8],
      ],
    );
    const secondImage = readerObject(secondRead, second, 1, "panel");
    assert.ok(secondImage.image?.part);
    assert.equal(
      sha256(unzipSync(second.pptx)[secondImage.image.part]!),
      sha256(panelPng),
    );
  });

  it("fills empty category and XY charts with mixed data and new element orders", async () => {
    const sourceScope = createMotionScopeHandle();
    const followerScope = createMotionScopeHandle();
    const badgeKey = createMorphScope().handle("fill-badge");
    const expanded = await expandDocument(
      <Deck width={1280} height={720} theme={theme}>
        <Slide id="empty-fill-page">
          <Box
            id="fill-layout"
            style={{
              position: "absolute",
              left: 40,
              top: 100,
              width: 1200,
              height: 560,
            }}
          >
            <Text
              id="fill-heading"
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                width: 800,
                height: 36,
                color: "ink",
                fontFamily: "major",
                fontSize: 24,
                fontWeight: 700,
              }}
            >
              Empty to filled
            </Text>
            <MotionScope scope={sourceScope}>
              <Chart
                id="empty-category-chart"
                title="Empty category chart"
                legend="bottom"
                style={{
                  position: "absolute",
                  left: 0,
                  top: 48,
                  width: 560,
                  height: 300,
                  fontFamily: "minor",
                  fontSize: "10pt",
                  color: "ink",
                }}
                groups={[
                  {
                    kind: "column",
                    data: { kind: "shared-categories", categories: [] },
                    series: [
                      { name: "Zero and filled", values: [] },
                      { name: "Second", values: [] },
                    ],
                  },
                ]}
                enter={fade({ by: "seriesElements", duration: 100 })}
              />
              <Chart
                id="empty-xy-chart"
                title="Empty XY chart"
                legend="bottom"
                style={{
                  position: "absolute",
                  left: 580,
                  top: 48,
                  width: 560,
                  height: 300,
                  fontFamily: "minor",
                  fontSize: "10pt",
                  color: "ink",
                }}
                groups={[
                  {
                    kind: "scatter",
                    data: { kind: "xy" },
                    line: "straight",
                    markers: true,
                    series: [
                      { name: "Still empty", x: [], y: [] },
                      { name: "Mixed", x: [], y: [] },
                      { name: "Short", x: [], y: [] },
                    ],
                  },
                ]}
                enter={fade({ by: "categoryElements", duration: 100 })}
              />
            </MotionScope>
            <MotionScope
              scope={followerScope}
              startAt={after(sourceScope.end, 25)}
            >
              <Shape
                id="fill-follower"
                style={{
                  position: "absolute",
                  left: 0,
                  top: 370,
                  width: 100,
                  height: 40,
                  background: "#1D4ED8",
                }}
                enter={fade({ duration: 50 })}
              />
            </MotionScope>
            <Box
              id="fill-badge-slot"
              style={{
                position: "absolute",
                left: 0,
                top: 430,
                width: 120,
                height: 40,
              }}
            >
              <Shape
                id="fill-badge"
                morphKey={badgeKey}
                style={{
                  width: 100,
                  height: 32,
                  background: "#C2410C",
                  color: "#FFFFFF",
                  fontFamily: "minor",
                  fontSize: 14,
                }}
              >
                Badge
              </Shape>
            </Box>
            <Box
              id="fill-badge-target"
              style={{
                position: "absolute",
                left: 180,
                top: 430,
                width: 120,
                height: 40,
              }}
            />
          </Box>
          <Notes id="fill-notes">
            <P id="fill-note">Empty charts are filled from the same nodes.</P>
            <P id="fill-note-second">Notes stay off the canvas.</P>
          </Notes>
        </Slide>
      </Deck>,
    );

    const first = await build(expanded);
    const firstReportSnapshot = structuredClone(first.report);
    const firstBytes = new Uint8Array(first.pptx);
    const firstSha = sha256(first.pptx);
    const firstRead = await validatePptx(first.pptx);
    assert.deepEqual(firstRead.errors, []);
    const firstCategory = readerObject(
      firstRead,
      first,
      0,
      "empty-category-chart",
    ).chart;
    const firstXY = readerObject(firstRead, first, 0, "empty-xy-chart").chart;
    assert.ok(firstCategory && firstXY);
    assert.deepEqual(
      firstCategory.series.map(({ index, categories: domain, values }) => ({
        index,
        domain,
        values,
      })),
      [
        { index: 0, domain: [], values: [] },
        { index: 1, domain: [], values: [] },
      ],
    );
    assert.deepEqual(
      firstXY.series.map(({ index, x, y }) => ({ index, x, y })),
      [
        { index: 0, x: [], y: [] },
        { index: 1, x: [], y: [] },
        { index: 2, x: [], y: [] },
      ],
    );
    const emptyCategoryEffects = chartEffects(
      motion(first),
      "empty-category-chart",
    );
    const emptyXYEffects = chartEffects(motion(first), "empty-xy-chart");
    const backgroundRange = {
      type: "chart",
      step: "gridLegend",
      seriesIndex: -3,
      categoryIndex: -3,
    };
    assert.deepEqual(
      emptyCategoryEffects.map(({ target }) => target.range),
      [backgroundRange],
    );
    assert.deepEqual(
      emptyXYEffects.map(({ target }) => target.range),
      [backgroundRange],
    );
    const firstFollower = motion(first)
      .steps.flatMap((step) => step.effects)
      .find(({ target }) => target.authorId === "fill-follower");
    assert.deepEqual(
      [firstFollower?.startMs, firstFollower?.endMs],
      [125, 175],
    );
    const firstCategorySnapshot = structuredClone(
      first
        .slide("empty-fill-page")
        .get("fill-layout")
        .get("empty-category-chart").box,
    );
    const firstNotes = firstRead.slides[0]?.notes;
    assert.equal(
      firstNotes?.text.text,
      "Empty charts are filled from the same nodes.\nNotes stay off the canvas.",
    );
    assert.ok(
      first.report.slides[0]?.objects.every(
        ({ authorId }) => authorId !== "fill-notes",
      ),
    );

    const page = expanded.query.get("empty-fill-page");
    const layout = page.get("fill-layout");
    const category = layout.getAs("empty-category-chart", isChart);
    const xy = layout.getAs("empty-xy-chart", isChart);
    const badge = layout.get("fill-badge-slot").get("fill-badge").node;
    badge.moveTo(container(layout.get("fill-badge-target").node));
    await appendJSX(
      container(layout.node),
      createElement(Text, {
        id: "fill-appended",
        style: {
          position: "absolute",
          left: 0,
          top: 480,
          width: 400,
          height: 24,
          color: "ink",
          fontFamily: "minor",
          fontSize: 14,
        },
        children: "Appended after the empty build",
      }),
    );
    category.set({
      groups: [
        {
          kind: "column",
          data: {
            kind: "shared-categories",
            categories: ["A", "B", "C", "D"],
          },
          series: [
            { name: "Zero and filled", values: [0, 4, null, 7] },
            { name: "Second", values: [1, 0, 3, null] },
          ],
        },
      ],
    });
    xy.set({
      groups: [
        {
          kind: "scatter",
          data: { kind: "xy" },
          line: "straight",
          markers: true,
          series: [
            { name: "Still empty", x: [], y: [] },
            { name: "Mixed", x: [0, 2, null, 5], y: [1, null, 4, 0] },
            { name: "Short", x: [1, 3], y: [2, 6] },
          ],
        },
      ],
    });

    const second = await build(expanded);
    const secondRead = await validatePptx(second.pptx);
    assert.deepEqual(secondRead.errors, []);
    assert.deepEqual(first.report, firstReportSnapshot);
    assert.equal(sha256(first.pptx), firstSha);
    assert.deepEqual(new Uint8Array(first.pptx), firstBytes);
    const rereadFirst = await validatePptx(first.pptx);
    assert.deepEqual(
      rereadFirst.slides[0]?.objects
        .find(
          ({ id }) =>
            id ===
            first.report.slides[0]?.objects.find(
              ({ authorId }) => authorId === "empty-category-chart",
            )?.id,
        )
        ?.chart?.series.map(({ index, values }) => ({ index, values })),
      [
        { index: 0, values: [] },
        { index: 1, values: [] },
      ],
    );
    assert.deepEqual(
      first
        .slide("empty-fill-page")
        .get("fill-layout")
        .get("empty-category-chart").box,
      firstCategorySnapshot,
    );
    assert.deepEqual(
      motion(first).steps.map((step) =>
        step.effects.map(({ target }) => target.authorId),
      ),
      [["empty-category-chart", "empty-xy-chart", "fill-follower"]],
    );

    const secondCategory = readerObject(
      secondRead,
      second,
      0,
      "empty-category-chart",
    ).chart;
    const secondXY = readerObject(
      secondRead,
      second,
      0,
      "empty-xy-chart",
    ).chart;
    assert.ok(secondCategory && secondXY);
    assert.deepEqual(
      secondCategory.series.map(({ index, categories: domain, values }) => ({
        index,
        domain,
        values,
      })),
      [
        {
          index: 0,
          domain: ["A", "B", "C", "D"],
          values: [0, 4, null, 7],
        },
        {
          index: 1,
          domain: ["A", "B", "C", "D"],
          values: [1, 0, 3, null],
        },
      ],
    );
    assert.deepEqual(
      secondCategory.embeddedWorkbook?.series.map(({ index, values }) => ({
        index,
        values,
      })),
      [
        { index: 0, values: [0, 4, null, 7] },
        { index: 1, values: [1, 0, 3, null] },
      ],
    );
    assert.deepEqual(
      secondXY.series.map(({ index, x, y }) => ({ index, x, y })),
      [
        { index: 0, x: [], y: [] },
        { index: 1, x: [0, 2, null, 5], y: [1, null, 4, 0] },
        { index: 2, x: [1, 3], y: [2, 6] },
      ],
    );
    assert.deepEqual(
      secondXY.embeddedWorkbook?.series.map(({ index, x, y }) => ({
        index,
        x,
        y,
      })),
      [
        { index: 0, x: [], y: [] },
        { index: 1, x: [0, 2, null, 5], y: [1, null, 4, 0] },
        { index: 2, x: [1, 3], y: [2, 6] },
      ],
    );

    const seriesElements = [
      backgroundRange,
      ...[0, 1, 2, 3].map((categoryIndex) => ({
        type: "chart",
        step: "ptInSeries",
        seriesIndex: 0,
        categoryIndex,
      })),
      ...[0, 1, 2, 3].map((categoryIndex) => ({
        type: "chart",
        step: "ptInSeries",
        seriesIndex: 1,
        categoryIndex,
      })),
    ];
    const categoryElements = [
      backgroundRange,
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 1,
        categoryIndex: 0,
      },
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 2,
        categoryIndex: 0,
      },
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 1,
        categoryIndex: 1,
      },
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 2,
        categoryIndex: 1,
      },
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 1,
        categoryIndex: 2,
      },
      {
        type: "chart",
        step: "ptInCategory",
        seriesIndex: 1,
        categoryIndex: 3,
      },
    ];
    const secondMotion = motion(second);
    assert.deepEqual(
      chartEffects(secondMotion, "empty-category-chart").map(
        ({ target }) => target.range,
      ),
      seriesElements,
    );
    assert.deepEqual(
      chartEffects(secondMotion, "empty-xy-chart").map(
        ({ target }) => target.range,
      ),
      categoryElements,
    );
    assert.deepEqual(
      chartEffects(secondMotion, "empty-category-chart").map(
        ({ startMs, endMs }) => [startMs, endMs],
      ),
      Array.from({ length: 9 }, (_, index) => [index * 100, (index + 1) * 100]),
    );
    assert.deepEqual(
      chartEffects(secondMotion, "empty-xy-chart").map(({ startMs, endMs }) => [
        startMs,
        endMs,
      ]),
      Array.from({ length: 7 }, (_, index) => [index * 100, (index + 1) * 100]),
    );
    const follower = secondMotion.steps
      .flatMap((step) => step.effects)
      .find(({ target }) => target.authorId === "fill-follower");
    assert.deepEqual([follower?.startMs, follower?.endMs], [925, 975]);

    const query = motionSchedule(second.slide("empty-fill-page"));
    const queryEffects = query?.steps.flatMap((step) => step.effects) ?? [];
    const lastCategory = queryEffects
      .filter(({ target }) => target.authorId === "empty-category-chart")
      .at(-1);
    assert.deepEqual([lastCategory?.startMs, lastCategory?.endMs], [800, 900]);

    const fileCategoryEffects =
      secondRead.slides[0]?.animations?.[0]?.effects.filter(
        ({ targetId }) =>
          targetId ===
          second.report.slides[0]?.objects.find(
            ({ authorId }) => authorId === "empty-category-chart",
          )?.id,
      );
    const fileXYEffects = secondRead.slides[0]?.animations?.[0]?.effects.filter(
      ({ targetId }) =>
        targetId ===
        second.report.slides[0]?.objects.find(
          ({ authorId }) => authorId === "empty-xy-chart",
        )?.id,
    );
    assert.equal(fileCategoryEffects?.length, seriesElements.length);
    assert.equal(fileXYEffects?.length, categoryElements.length);
    assert.ok(
      fileCategoryEffects?.every(
        ({ nativeBuild }) => nativeBuild === "seriesEl",
      ),
    );
    assert.ok(
      fileXYEffects?.every(({ nativeBuild }) => nativeBuild === "categoryEl"),
    );
    assert.deepEqual(
      fileCategoryEffects?.map(({ targetRange }) => targetRange),
      seriesElements,
    );
    assert.deepEqual(
      fileXYEffects?.map(({ targetRange }) => targetRange),
      categoryElements,
    );

    const appended = second.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "fill-appended",
    );
    assert.equal(appended?.text, "Appended after the empty build");
    const secondBadge = second.report.slides[0]?.objects.find(
      ({ authorId }) => authorId === "fill-badge",
    );
    assert.deepEqual(secondBadge?.box, {
      x: 220,
      y: 530,
      width: 100,
      height: 32,
    });
    assert.deepEqual(
      secondRead.slides[0]?.notes?.text.text,
      firstNotes?.text.text,
    );
    assert.ok(
      second.report.slides[0]?.objects.every(
        ({ authorId }) => authorId !== "fill-notes",
      ),
    );
  });

  it("resolves once then emits and paints a 2-page deck mixing all eight canvas capabilities, notes, sections, transitions, and motion with zero preview errors", async () => {
    const notoSc = font("Noto Sans CJK SC", [
      {
        weight: 400,
        data: new Uint8Array(
          readFileSync(
            new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
          ),
        ),
      },
    ]);
    const validPng1x1 = new Uint8Array(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
        "base64",
      ),
    );
    const integrationTheme = defineTheme({
      name: "Full Canvas Parity Integration",
      colors: {
        brand: "#1E3A8A",
        ink: "#0F172A",
        accent: "#2563EB",
        accent2: "#059669",
      },
      fonts: { major: jost, minor: jost },
    });
    const morphScope = createMorphScope();
    const heroMorph = morphScope.handle("parity-hero");
    const openBezierArrow = path(
      "M 10 70 C 35 10 65 10 90 45 S 135 80 150 25",
      { viewBox: [160, 90] },
    );

    const expanded = await expandDocument(
      createElement(
        Deck,
        {
          title: "Full Canvas Parity Integration",
          width: 960,
          height: 540,
          theme: integrationTheme,
        },
        createElement(
          Section,
          { name: "Canvas Parity Showcase" },
          createElement(
            Slide,
            {
              id: "parity-canvas-1",
              background: linearGradient(135, "#F8FAFC", "#E0F2FE"),
              transition: transitions.fade({ duration: 400 }),
            },
            createElement(
              Notes,
              { id: "p1-notes" },
              "Speaker cues for page 1 cross-capability layout.",
            ),
            createElement(
              Box,
              {
                id: "styled-box",
                theme: { colors: { accent: "#0F766E" } },
                style: {
                  position: "absolute",
                  left: 24,
                  top: 20,
                  width: 240,
                  height: 110,
                  padding: 12,
                  background: radialGradient("#DBEAFE", "#93C5FD"),
                  border: { color: "brand", width: 2, style: "dashed" },
                  borderRadius: 14,
                  boxShadow: {
                    color: "#0F172A",
                    blur: 4,
                    distance: 3,
                    angle: 90,
                  },
                  softEdge: 3,
                },
              },
              createElement(
                Text,
                {
                  id: "box-caption",
                  style: { fontFamily: jost, fontSize: 15, color: "accent" },
                },
                "Radial Box + Theme Override",
              ),
            ),
            createElement(
              Step,
              { n: 1 },
              createElement(
                Shape,
                {
                  id: "morph-origin-card",
                  morphKey: heroMorph,
                  link: { id: "parity-canvas-2" },
                  enter: fade({ duration: 250 }),
                  geometry: preset("roundRect", { adj: 20000 }),
                  style: {
                    position: "absolute",
                    left: 24,
                    top: 144,
                    width: 240,
                    height: 60,
                    background: "accent",
                    fontFamily: jost,
                    fontSize: 14,
                    color: "#FFFFFF",
                    align: "center",
                    verticalAlign: "middle",
                  },
                },
                "Morph + Link Origin",
              ),
            ),
            createElement(
              Group,
              {
                id: "outer-rot-group",
                style: { rotate: 12 },
              },
              createElement(
                Group,
                {
                  id: "inner-nested-group",
                },
                createElement(
                  Shape,
                  {
                    id: "grouped-arrow-shape",
                    geometry: preset("rightArrow", { adj1: 30000 }),
                    style: {
                      position: "absolute",
                      left: 290,
                      top: 30,
                      width: 200,
                      height: 68,
                      background: "#7C3AED",
                      border: { color: "#4C1D95", width: 1.5 },
                      fontFamily: jost,
                      fontSize: 13,
                      color: "#FFFFFF",
                      align: "center",
                      verticalAlign: "middle",
                    },
                  },
                  "Nested Rich Arrow",
                ),
                createElement(Shape, {
                  id: "grouped-open-bezier",
                  geometry: openBezierArrow,
                  style: {
                    position: "absolute",
                    left: 300,
                    top: 110,
                    width: 190,
                    height: 68,
                    border: {
                      color: "#0F172A",
                      width: 2,
                      startArrow: "oval",
                      endArrow: "stealth",
                    },
                  },
                }),
              ),
            ),
            createElement(Image, {
              id: "clipped-hex-image",
              src: validPng1x1,
              geometry: preset("hexagon"),
              style: {
                position: "absolute",
                left: 560,
                top: 20,
                width: 110,
                height: 96,
                border: { color: "#1D4ED8", width: 2 },
                boxShadow: {
                  color: "#1E293B",
                  blur: 4,
                  distance: 3,
                  angle: 45,
                },
                softEdge: 2,
              },
            }),
            createElement(
              Text,
              {
                id: "multiscript-rich-text",
                style: {
                  position: "absolute",
                  left: 690,
                  top: 20,
                  width: 246,
                  height: 220,
                  fontFamily: jost,
                  fontSize: 13,
                  color: "ink",
                  padding: 10,
                  background: "#FFFFFF",
                  border: { color: "#94A3B8", width: 1 },
                },
              },
              createElement(
                P,
                { style: { align: "justify", spacingAfter: 6 } },
                "Slide ",
                createElement(SlideNumber, null),
                " ",
                createElement(
                  Span,
                  { style: { fontFamily: notoSc } },
                  "汉字 あア ∑≠",
                ),
                " justified multi-word copy for layout parity.",
              ),
              createElement(
                P,
                {
                  style: {
                    list: { type: "bullet" },
                    textDecoration: "underline",
                    letterSpacing: 1.5,
                  },
                },
                "Wide Bullet",
              ),
              createElement(
                P,
                {
                  style: {
                    list: { type: "number", startAt: 1 },
                    textDecoration: "line-through",
                    letterSpacing: -0.5,
                    textShadow: {
                      color: "#475569",
                      blur: 1,
                      distance: 1,
                      angle: 90,
                    },
                  },
                },
                "Tight Shadowed",
              ),
            ),
            createElement(
              Table,
              {
                id: "parity-table",
                columns: [180, "auto", "1fr"],
                style: {
                  position: "absolute",
                  left: 24,
                  top: 256,
                  width: 912,
                  fontFamily: jost,
                  fontSize: 13,
                  color: "ink",
                },
              },
              createElement(
                Row,
                { id: "tbl-header", style: { height: 32 } },
                createElement(
                  Cell,
                  {
                    id: "hdr-span",
                    colSpan: 2,
                    style: {
                      background: "#1E3A8A",
                      color: "#FFFFFF",
                      padding: 6,
                      border: { color: "#0F172A", width: 1, style: "solid" },
                    },
                  },
                  "Merged Header",
                ),
                createElement(
                  Cell,
                  {
                    id: "hdr-right",
                    style: {
                      background: "#EFF6FF",
                      padding: 6,
                      border: { color: "#2563EB", width: 1.5, style: "dashed" },
                    },
                  },
                  "Auto / Fr Tracks",
                ),
              ),
              createElement(
                Row,
                { id: "tbl-body-1", style: { height: 32 } },
                createElement(
                  Cell,
                  {
                    id: "row-span-cell",
                    rowSpan: 2,
                    style: {
                      padding: 6,
                      border: { color: "#059669", width: 1.5, style: "dotted" },
                    },
                  },
                  "Merged Side",
                ),
                createElement(
                  Cell,
                  {
                    id: "overset-cell",
                    style: {
                      padding: 2,
                      border: { color: "#DC2626", width: 1 },
                    },
                  },
                  "OVERSET_LINE_1\nOVERSET_LINE_2\nOVERSET_LINE_3\nOVERSET_LINE_4",
                ),
                createElement(
                  Cell,
                  {
                    id: "normal-cell-1",
                    style: {
                      padding: 4,
                      border: { color: "#CBD5E1", width: 1 },
                    },
                  },
                  "Dashed/Dotted Grid",
                ),
              ),
              createElement(
                Row,
                { id: "tbl-body-2", style: { height: 32 } },
                createElement(
                  Cell,
                  {
                    id: "normal-cell-2",
                    style: {
                      padding: 6,
                      border: { color: "#CBD5E1", width: 1 },
                    },
                  },
                  "Row 2 Col 2",
                ),
                createElement(
                  Cell,
                  {
                    id: "normal-cell-3",
                    style: {
                      padding: 6,
                      border: { color: "#CBD5E1", width: 1 },
                    },
                  },
                  "Row 2 Col 3",
                ),
              ),
            ),
          ),
          createElement(
            Slide,
            {
              id: "parity-canvas-2",
              background: { image: validPng1x1 },
              transition: transitions.morph({
                duration: 800,
                fallback: transitions.fade(),
              }),
            },
            createElement(
              Shape,
              {
                id: "morph-target-card",
                morphKey: heroMorph,
                geometry: preset("roundRect", { adj: 25000 }),
                style: {
                  position: "absolute",
                  left: 24,
                  top: 20,
                  width: 280,
                  height: 64,
                  background: "accent2",
                  fontFamily: jost,
                  fontSize: 15,
                  color: "#FFFFFF",
                  align: "center",
                  verticalAlign: "middle",
                },
              },
              "Morph Target Card",
            ),
            createElement(Chart, {
              id: "combo-dual-axis-chart",
              title: {
                text: "Combo Dual Axis + Trend",
                style: {
                  fontSize: 15,
                  color: "brand",
                  textDecoration: "underline",
                  textShadow: {
                    color: "#0F172A",
                    blur: 1,
                    distance: 1,
                    angle: 90,
                  },
                },
              },
              legend: "bottom",
              style: {
                position: "absolute",
                left: 24,
                top: 100,
                width: 520,
                height: 416,
                fontFamily: jost,
                fontSize: 11,
              },
              axes: {
                x: { kind: "category" },
                yLeft: { kind: "value", position: "left" },
                yRight: {
                  kind: "value",
                  position: "right",
                  numberFormat: "0%",
                },
              },
              groups: [
                {
                  kind: "column",
                  axes: { category: "x", value: "yLeft" },
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2", "Q3", "Q4"],
                  },
                  dataLabels: { value: true, position: "outsideEnd" },
                  series: [
                    {
                      name: "Revenue",
                      values: [40, 65, 55, 80],
                      errorBars: { type: "fixed", value: 5 },
                      trendlines: [
                        {
                          type: "linear",
                          displayEquation: true,
                          displayRSquared: true,
                          line: { color: "#DC2626", width: 1.5 },
                        },
                      ],
                    },
                  ],
                },
                {
                  kind: "line",
                  axes: { category: "x", value: "yRight" },
                  data: {
                    kind: "shared-categories",
                    categories: ["Q1", "Q2", "Q3", "Q4"],
                  },
                  series: [
                    {
                      name: "Margin",
                      values: [0.18, 0.24, 0.21, 0.29],
                    },
                  ],
                },
              ],
            }),
            createElement(Chart, {
              id: "multi-ring-doughnut",
              title: "Multi-Ring Independent Categories",
              legend: "top",
              style: {
                position: "absolute",
                left: 564,
                top: 100,
                width: 372,
                height: 416,
                fontFamily: jost,
                fontSize: 11,
              },
              groups: [
                {
                  kind: "doughnut",
                  holeSize: 45,
                  data: { kind: "independent-categories" },
                  dataLabels: { category: true, percent: true },
                  series: [
                    {
                      name: "Inner",
                      categories: ["Alpha", "Beta"],
                      values: [60, 40],
                    },
                    {
                      name: "Outer",
                      categories: ["North", "South", "East"],
                      values: [50, 30, 20],
                    },
                  ],
                },
              ],
            }),
          ),
        ),
      ),
    );

    const resolution = resolve(expanded);
    const parityTable = resolution.slide("parity-canvas-1").get("parity-table");
    const oversetMetrics = parityTable
      .get("tbl-body-1")
      .get("overset-cell").textMetrics;
    assert.equal(oversetMetrics?.status, "measured");
    if (oversetMetrics?.status === "measured") {
      assert.equal(oversetMetrics.fits, false);
    }
    const multiscriptMetrics = resolution
      .slide("parity-canvas-1")
      .get("multiscript-rich-text").textMetrics;
    assert.equal(multiscriptMetrics?.status, "measured");
    if (multiscriptMetrics?.status === "measured") {
      assert.equal(multiscriptMetrics.fits, true);
    }
    assert.ok(resolution.report.extensions.includes("p14"));
    assert.ok(resolution.report.extensions.includes("p159"));
    assert.ok(resolution.report.extensions.includes("c15"));

    const emitted = emit(resolution);
    const validation = await validatePptx(emitted.pptx, {
      allowedExtensions: ["p14", "p159", "c15"],
    });
    assert.deepEqual(validation.errors, []);
    assert.equal(validation.slides.length, 2);
    assert.deepEqual(validation.presentation.sections, [
      {
        name: "Canvas Parity Showcase",
        slideIds: ["parity-canvas-1", "parity-canvas-2"],
      },
    ]);
    assert.ok(
      validation.slides[0]?.notes?.text.text.includes(
        "Speaker cues for page 1",
      ),
    );
    assert.ok((validation.slides[0]?.animations?.length ?? 0) >= 1);
    const fileTable = validation.slides[0]?.objects.find(
      (o) => o.kind === "table",
    )?.table;
    assert.deepEqual(fileTable?.merges, [
      { row: 0, column: 0, rowSpan: 1, colSpan: 2 },
      { row: 1, column: 0, rowSpan: 2, colSpan: 1 },
    ]);
    const charts = validation.slides[1]?.objects.filter(
      (o) => o.chart !== undefined,
    );
    assert.equal(charts?.length, 2);
    assert.ok(charts?.every((c) => c.chart?.embeddedWorkbook !== undefined));

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 2);
    const pngSignature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
    for (const page of painted.pages) {
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.placeholders, []);
      assert.ok(page.png !== undefined);
      assert.ok((page.svg?.trim().length ?? 0) > 0);
      assert.deepEqual([...page.png.subarray(0, 8)], pngSignature);
      const view = new DataView(
        page.png.buffer,
        page.png.byteOffset,
        page.png.byteLength,
      );
      assert.equal(view.getUint32(16, false), 960);
      assert.equal(view.getUint32(20, false), 540);
    }

    const groupedArrow = painted.pages[0]?.nodes.find(
      (n) => n.authorId === "grouped-arrow-shape",
    );
    const groupedBezier = painted.pages[0]?.nodes.find(
      (n) => n.authorId === "grouped-open-bezier",
    );
    assert.ok(groupedArrow?.bounds !== undefined);
    assert.equal(groupedArrow?.corners?.length, 4);
    assert.ok(groupedBezier?.bounds !== undefined);
    assert.equal(groupedBezier?.corners?.length, 4);
  });
});
