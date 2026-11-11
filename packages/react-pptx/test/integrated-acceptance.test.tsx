import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import {
  appendJSX,
  Box,
  Chart,
  CommonContainer,
  CommonNode,
  createMorphScope,
  createMotionScopeHandle,
  Deck,
  defineTheme,
  expandDocument,
  fade,
  font,
  Image,
  MotionScope,
  Notes,
  P,
  Shape,
  Slide,
  Span,
  Step,
  Text,
  type MotionObservation,
} from "@react-pptx/react-pptx";
import {
  build,
  type ChartGroup,
  type ChartNode,
  type ChartProps,
  type TextNode,
} from "@react-pptx/react-pptx/document";
import { validatePptx, type FileObject } from "@react-pptx/validation";
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

function motion(result: Awaited<ReturnType<typeof build>>, slideIndex = 0) {
  const value = result.report.slides[slideIndex]?.capabilities?.["motion"] as
    | MotionObservation
    | undefined;
  assert.ok(value);
  return value;
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
        { id: "appended-caption", width: 300, height: 30, style },
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
                width={180}
                height={64}
                fill="#336699"
                text="Intro"
                enter={fade({ duration: 120 })}
                style={heroStyle}
              />
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
                  width={380}
                  height={32}
                  style={{
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
            x={40}
            y={480}
            width={400}
            height={24}
            style={{ color: "ink", fontSize: 12, fontFamily: "minor" }}
          >
            Footer
          </Text>
        </Slide>
        <Slide id="continuation">
          <Shape
            id="hero-next"
            morphKey={heroKey}
            x={60}
            y={60}
            width={200}
            height={70}
            fill="#336699"
          />
        </Slide>
      </Deck>,
    );

    const first = await build(expanded);
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
    const firstHeroObservation = first.query
      .getSlide("story")
      .get("source-panel")
      .get("hero")
      .observe();
    assert.deepEqual(firstHeroObservation.box, {
      x: 50,
      y: 130,
      width: 180,
      height: 64,
    });
    const firstQueryText = first.query
      .getSlide("story")
      .get("stage")
      .get("stage-title")
      .observe().text;
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
      first.query.getSlide("story").get("source-panel").get("hero").observe(),
      firstHeroObservation,
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
      first.query.getSlide("story").get("stage").get("stage-title").observe()
        .text,
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
    const scope = secondMotion.scopes?.find(({ members }) =>
      members.some(
        (member) => member.kind === "object" && member.id === secondHeroId,
      ),
    );
    assert.ok(scope);
    assert.deepEqual(scope.members, [{ kind: "object", id: secondHeroId }]);
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
      x: 0,
      y: 0,
      width: 552,
      height: 240,
      title: "Revenue and trend",
      legend: "bottom",
      style: { fontFamily: "minor", fontSize: "10pt", color: "ink" },
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
              width={100}
              height={32}
              fill="#1D4ED8"
              text="Move me"
              style={{
                color: "#FFFFFF",
                fontFamily: "minor",
                fontSize: 14,
              }}
            />
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
            x={0}
            y={0}
            width={520}
            height={360}
            title="Independent rings"
            legend="right"
            style={{ fontFamily: "minor", fontSize: "10pt", color: "ink" }}
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
            x={560}
            y={40}
            width={320}
            height={160}
          />
        </Slide>
      </Deck>,
    );

    const first = await build(expanded);
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
    const firstPanelObservation = first.query
      .getSlide("rings-page")
      .get("panel")
      .observe();
    assert.deepEqual(firstPanelObservation.box, {
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
    categories[0] = "Mutated";
    actualValues[0] = 99;
    customPlus[0] = 99;
    panelBytes[0] = 0;

    const second = await build(expanded);
    const secondRead = await validatePptx(second.pptx);
    assert.deepEqual(secondRead.errors, []);
    assert.equal(sha256(first.pptx), firstSha);
    assert.deepEqual(new Uint8Array(first.pptx), firstBytes);
    assert.deepEqual(
      first.query.getSlide("combo-page").get("body").observe().box,
      firstBodyBox,
    );
    assert.deepEqual(
      first.query.getSlide("combo-page").get("combo-chart").observe().box,
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
      first.query
        .getSlide("combo-page")
        .get("badge-slot")
        .get("badge")
        .observe().box,
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
});
