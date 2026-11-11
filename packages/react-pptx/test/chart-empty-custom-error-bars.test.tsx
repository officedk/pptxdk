import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Chart,
  Deck,
  Slide,
  type ChartErrorBars,
  type ChartGroup,
} from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  chart as documentChart,
  deck,
  slide as documentSlide,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync, type Unzipped } from "fflate";
import { describe, it } from "vitest";

function chartXml(parts: Unzipped, index = 1): string {
  return strFromU8(parts[`ppt/charts/chart${index}.xml`]!);
}

function sheetXml(parts: Unzipped, workbookPart: string): string {
  return strFromU8(
    unzipSync(parts[workbookPart]!)["xl/worksheets/sheet1.xml"]!,
  );
}

function numLitZero(side: "plus" | "minus"): string {
  return `<c:${side}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numLit></c:${side}>`;
}

const customBoth = {
  type: "custom" as const,
  plus: [],
  minus: [],
  endCap: false,
  line: { color: "#336699", width: 1 },
};

function lineZeroGroups(errorBars: ChartErrorBars): readonly ChartGroup[] {
  return [
    {
      kind: "line",
      data: { kind: "shared-categories", categories: [] },
      series: [{ name: "Empty", values: [], errorBars }],
    },
  ];
}

describe("zero-observation custom error bars", () => {
  for (const kind of ["line", "column"] as const) {
    it(`writes paired zero custom sources as literals for ${kind}`, async () => {
      const result = await build(
        <Deck width={800} height={600}>
          <Slide id="zero-custom">
            <Chart
              id="zero-custom"
              x={10}
              y={10}
              width={760}
              height={560}
              groups={[
                {
                  kind,
                  data: { kind: "shared-categories", categories: [] },
                  series: [
                    { name: "Empty", values: [], errorBars: customBoth },
                  ],
                } as ChartGroup,
              ]}
            />
          </Slide>
        </Deck>,
      );
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const chart = file.slides[0]?.objects[0]?.chart;
      assert.deepEqual(chart?.series[0]?.values, []);
      assert.deepEqual(
        chart?.series[0]?.errorBars?.map(
          ({ type, include, endCap, plus, minus }) => ({
            type,
            include,
            endCap,
            plus,
            minus,
          }),
        ),
        [
          {
            type: "custom",
            include: "both",
            endCap: false,
            plus: { source: "literal", values: [] },
            minus: { source: "literal", values: [] },
          },
        ],
      );
      const parts = unzipSync(result.pptx);
      const xml = chartXml(parts);
      assert.ok(xml.includes('<c:errValType val="cust"/>'));
      assert.ok(xml.includes('<c:errBarType val="both"/>'));
      assert.ok(xml.includes('<c:noEndCap val="1"/>'));
      assert.ok(xml.includes(numLitZero("plus")));
      assert.ok(xml.includes(numLitZero("minus")));
      assert.ok(!xml.includes("$C$2:$C$1"));
      assert.ok(!xml.includes("$D$2:$D$1"));
      assert.ok(!xml.includes("<c:numRef>"));
      const workbookPart = chart?.embeddedWorkbook?.part;
      assert.ok(workbookPart);
      const sheet = sheetXml(parts, workbookPart);
      assert.ok(sheet.includes("Empty +"));
      assert.ok(sheet.includes("Empty -"));
      assert.ok(!sheet.includes('r="C2"'));
      assert.ok(!sheet.includes('r="D2"'));
    });
  }

  it("writes a plus-only zero custom source with its include and style", async () => {
    const node = documentChart({
      id: "zero-custom-plus",
      x: 10,
      y: 10,
      width: 760,
      height: 560,
      groups: lineZeroGroups({
        type: "custom",
        include: "plus",
        plus: [],
        endCap: false,
        line: { color: "#AA5500", width: 2, dash: "dash" },
      }),
    });
    const document = deck({ width: 800, height: 600 }).append(
      documentSlide({ id: "zero-custom-plus" }).append(node),
    );
    const result = await buildDocument(document);
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const chart = file.slides[0]?.objects[0]?.chart;
    const bars = chart?.series[0]?.errorBars?.[0];
    assert.equal(bars?.type, "custom");
    assert.equal(bars?.include, "plus");
    assert.equal(bars?.endCap, false);
    assert.deepEqual(bars?.plus, { source: "literal", values: [] });
    assert.equal(bars?.minus, undefined);
    const xml = chartXml(unzipSync(result.pptx));
    assert.ok(xml.includes('<c:errBarType val="plus"/>'));
    assert.ok(xml.includes(numLitZero("plus")));
    assert.ok(!xml.includes("<c:minus>"));
    assert.ok(xml.includes('val="AA5500"'));
    assert.ok(xml.includes('w="19050"'));
  });

  it("writes a minus-only zero custom source", async () => {
    const result = await build(
      <Deck width={800} height={600}>
        <Slide id="zero-custom-minus">
          <Chart
            id="zero-custom-minus"
            x={10}
            y={10}
            width={760}
            height={560}
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: [] },
                series: [
                  {
                    name: "Empty",
                    values: [],
                    errorBars: {
                      type: "custom",
                      include: "minus",
                      minus: [],
                      endCap: true,
                    },
                  },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const bars = file.slides[0]?.objects[0]?.chart?.series[0]?.errorBars?.[0];
    assert.equal(bars?.include, "minus");
    assert.deepEqual(bars?.minus, { source: "literal", values: [] });
    assert.equal(bars?.plus, undefined);
    const xml = chartXml(unzipSync(result.pptx));
    assert.ok(xml.includes('<c:errBarType val="minus"/>'));
    assert.ok(xml.includes(numLitZero("minus")));
    assert.ok(!xml.includes("<c:plus>"));
    assert.ok(!xml.includes("<c:noEndCap"));
  });

  it("writes zero X and Y custom error sources without inverse ranges", async () => {
    const result = await build(
      <Deck width={800} height={600}>
        <Slide id="zero-custom-xy">
          <Chart
            id="zero-custom-xy"
            x={10}
            y={10}
            width={760}
            height={560}
            groups={[
              {
                kind: "scatter",
                data: { kind: "xy" },
                series: [
                  {
                    name: "Empty XY",
                    x: [],
                    y: [],
                    errorBars: {
                      x: { type: "custom", plus: [], minus: [] },
                      y: { type: "custom", include: "plus", plus: [] },
                    },
                  },
                ],
              },
            ]}
          />
        </Slide>
      </Deck>,
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const bars = file.slides[0]?.objects[0]?.chart?.series[0]?.errorBars;
    assert.deepEqual(
      bars?.map(({ direction, type, include, plus, minus }) => ({
        direction,
        type,
        include,
        plus,
        minus,
      })),
      [
        {
          direction: "x",
          type: "custom",
          include: "both",
          plus: { source: "literal", values: [] },
          minus: { source: "literal", values: [] },
        },
        {
          direction: "y",
          type: "custom",
          include: "plus",
          plus: { source: "literal", values: [] },
          minus: undefined,
        },
      ],
    );
    const parts = unzipSync(result.pptx);
    const xml = chartXml(parts);
    assert.ok(xml.includes('<c:errDir val="x"/>'));
    assert.ok(xml.includes('<c:errDir val="y"/>'));
    assert.ok(xml.includes(numLitZero("plus")));
    assert.ok(xml.includes(numLitZero("minus")));
    assert.ok(!xml.includes("$D$2:$D$1"));
    assert.ok(!xml.includes("$E$2:$E$1"));
    assert.ok(!xml.includes("$F$2:$F$1"));
    assert.ok(!xml.includes("$G$2:$G$1"));
  });

  it("fills a zero custom node and rebuilds the non-zero custom references", async () => {
    const node = documentChart({
      id: "zero-custom-fill",
      x: 10,
      y: 10,
      width: 760,
      height: 560,
      groups: lineZeroGroups(customBoth),
    });
    const document = deck({ width: 800, height: 600 }).append(
      documentSlide({ id: "zero-custom-fill" }).append(node),
    );
    const emptyResult = await buildDocument(document);
    const emptyFile = await validatePptx(emptyResult.pptx);
    assert.deepEqual(emptyFile.errors, []);
    const emptyPptx = emptyResult.pptx.slice();

    node.set({
      groups: [
        {
          kind: "line",
          data: { kind: "shared-categories", categories: ["A", "B"] },
          series: [
            {
              name: "Empty",
              values: [3, 7],
              errorBars: {
                type: "custom",
                plus: [1, 2],
                minus: [0.5, 1],
                endCap: false,
                line: { color: "#AA5500", width: 2 },
              },
            },
          ],
        },
      ],
    });
    const filledResult = await buildDocument(document);
    const filledFile = await validatePptx(filledResult.pptx);
    assert.deepEqual(filledFile.errors, []);
    const chart = filledFile.slides[0]?.objects[0]?.chart;
    assert.deepEqual(chart?.series[0]?.categories, ["A", "B"]);
    assert.deepEqual(chart?.series[0]?.values, [3, 7]);
    const bars = chart?.series[0]?.errorBars?.[0];
    assert.equal(bars?.type, "custom");
    assert.deepEqual(bars?.plus, {
      source: "reference",
      formula: "Sheet1!$C$2:$C$3",
      values: [1, 2],
    });
    assert.deepEqual(bars?.minus, {
      source: "reference",
      formula: "Sheet1!$D$2:$D$3",
      values: [0.5, 1],
    });
    const parts = unzipSync(filledResult.pptx);
    const xml = chartXml(parts);
    assert.ok(xml.includes("Sheet1!$C$2:$C$3"));
    assert.ok(xml.includes("Sheet1!$D$2:$D$3"));
    assert.ok(!xml.includes("$C$2:$C$1"));
    const workbookPart = chart?.embeddedWorkbook?.part;
    assert.ok(workbookPart);
    const sheet = sheetXml(parts, workbookPart);
    assert.ok(sheet.includes('<c r="C2"><v>1</v></c>'));
    assert.ok(sheet.includes('<c r="C3"><v>2</v></c>'));
    assert.ok(sheet.includes('<c r="D2"><v>0.5</v></c>'));
    assert.ok(sheet.includes('<c r="D3"><v>1</v></c>'));
    assert.deepEqual(emptyResult.pptx, emptyPptx);
  });

  for (const file of ["custom-both.pptx", "custom-plus.pptx"]) {
    it(`rejects the recorded inverse custom source in ${file}`, async () => {
      const bytes = new Uint8Array(
        readFileSync(
          new URL(`./assets/native-23-zero-custom/${file}`, import.meta.url),
        ),
      );
      const result = await validatePptx(bytes);
      assert.ok(
        result.errors.some(
          ({ code, part }) =>
            code === "invalid-chart-formula" &&
            part === "ppt/charts/chart1.xml",
        ),
        JSON.stringify(result.errors),
      );
    });
  }
});
