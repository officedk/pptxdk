import assert from "node:assert/strict";

import {
  Box,
  build,
  BuildError,
  Cell,
  Chart,
  Deck,
  Group,
  P,
  Row,
  Shape,
  Slide,
  Span,
  Table,
  Text,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, it } from "vitest";

const sampleChartData = {
  categories: ["Q1", "Q2"],
  series: [
    {
      name: "Revenue",
      values: [100, 200],
    },
  ],
};

describe("text language (lang) validation", () => {
  it("rejects non-string lang", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck
            // @ts-expect-error non-string lang
            style={{ lang: 123 }}
          >
            <Slide id="cover">
              <Text x={0} y={0} width={200} height={100}>
                Deck non-string lang
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );

    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide
              id="cover"
              // @ts-expect-error non-string lang
              style={{ lang: true }}
            >
              <Text x={0} y={0} width={200} height={100}>
                Slide non-string lang
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Deck", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck style={{ lang: "invalid_tag" }}>
            <Slide id="cover">
              <Text x={0} y={0} width={200} height={100}>
                Deck invalid lang
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects empty string lang on Slide", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover" style={{ lang: "" }}>
              <Text x={0} y={0} width={200} height={100}>
                Slide empty lang
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Box", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Box style={{ lang: "zh_CN" }}>
                <Text x={0} y={0} width={200} height={100}>
                  Box invalid lang
                </Text>
              </Box>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Group", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Group style={{ lang: "zh..CN" }}>
                <Text x={0} y={0} width={200} height={100}>
                  Group invalid lang
                </Text>
              </Group>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Text", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Text
                x={0}
                y={0}
                width={200}
                height={100}
                style={{ lang: "en_US" }}
              >
                Text invalid lang
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Shape", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Shape
                x={0}
                y={0}
                width={200}
                height={100}
                text="Shape text"
                style={{ lang: "bad_tag" }}
              />
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Table", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Table columns={[200]} style={{ lang: "invalid lang" }}>
                <Row style={{ height: 40 }}>
                  <Cell>
                    <P>
                      <Span>Cell content</Span>
                    </P>
                  </Cell>
                </Row>
              </Table>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Row and Cell", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Table columns={[200]}>
                <Row style={{ height: 40, lang: "invalid_row" }}>
                  <Cell>
                    <P>
                      <Span>Cell content</Span>
                    </P>
                  </Cell>
                </Row>
              </Table>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );

    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Table columns={[200]}>
                <Row style={{ height: 40 }}>
                  <Cell style={{ lang: "invalid_cell" }}>
                    <P>
                      <Span>Cell content</Span>
                    </P>
                  </Cell>
                </Row>
              </Table>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on Chart", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Chart
                x={0}
                y={0}
                width={300}
                height={200}
                kind="bar"
                data={sampleChartData}
                style={{ lang: "invalid_chart" }}
              />
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });

  it("rejects invalid lang on P and Span", async () => {
    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Text x={0} y={0} width={200} height={100}>
                <P style={{ lang: "bad_p" }}>
                  <Span>Paragraph content</Span>
                </P>
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );

    await assert.rejects(
      async () => {
        await build(
          <Deck>
            <Slide id="cover">
              <Text x={0} y={0} width={200} height={100}>
                <P>
                  <Span style={{ lang: "bad_span" }}>Span content</Span>
                </P>
              </Text>
            </Slide>
          </Deck>,
        );
      },
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-text-lang");
        return true;
      },
    );
  });
});

describe("omission of lang attribute when undeclared", () => {
  it("omits lang attribute on a:rPr, a:defRPr, and a:endParaRPr when lang is not declared", async () => {
    const result = await build(
      <Deck>
        <Slide id="slide-undeclared">
          <Text x={50} y={50} width={300} height={150}>
            <P>
              <Span>Run text</Span>
            </P>
            <P>{/* Empty paragraph to generate a:endParaRPr */}</P>
          </Text>
          <Shape
            x={50}
            y={220}
            width={300}
            height={100}
            text="Shape text content"
          />
          <Table x={50} y={340} columns={[150, 150]}>
            <Row style={{ height: 40 }}>
              <Cell colSpan={2}>
                <P>
                  <Span>Merged cell</Span>
                </P>
              </Cell>
            </Row>
            <Row style={{ height: 40 }}>
              <Cell>
                <P>
                  <Span>Cell 2.1</Span>
                </P>
              </Cell>
              <Cell>
                <P>
                  <Span>Cell 2.2</Span>
                </P>
              </Cell>
            </Row>
          </Table>
          <Chart
            x={400}
            y={50}
            width={300}
            height={200}
            kind="bar"
            data={sampleChartData}
          />
        </Slide>
      </Deck>,
    );

    const unzipped = unzipSync(result.pptx);
    const slideXml = strFromU8(unzipped["ppt/slides/slide1.xml"]!);
    assert.ok(slideXml, "slide1.xml should exist");

    // Must NOT have lang="en-US" or any lang attribute in slide text
    assert.ok(
      !slideXml.includes('lang="en-US"'),
      "slide1.xml should not contain hardcoded lang='en-US'",
    );
    assert.ok(
      !slideXml.includes("lang="),
      "slide1.xml should not contain any lang attribute when lang is undeclared",
    );

    // Verify empty paragraph and table merged continuation cells have <a:endParaRPr/> without lang
    assert.ok(
      slideXml.includes("<a:endParaRPr"),
      "slide1.xml should contain a:endParaRPr",
    );

    // Check chart XML
    const chartXml = strFromU8(unzipped["ppt/charts/chart1.xml"]!);
    assert.ok(chartXml, "chart1.xml should exist");
    // Chart <c:lang> falls back to en-US when undeclared
    assert.ok(
      chartXml.includes('<c:lang val="en-US"/>'),
      "chart1.xml should fallback <c:lang val='en-US'/> when undeclared",
    );
    // But chart text properties <a:defRPr> and <a:endParaRPr> must omit lang
    assert.ok(
      !chartXml.includes('<a:endParaRPr lang="en-US"/>'),
      "chart1.xml txPr should not have lang='en-US' on a:endParaRPr",
    );

    // Validation must pass
    const validated = await validatePptx(result.pptx);
    assert.deepEqual(validated.errors, []);
  });
});

describe("inheritance and local overrides of text language (lang)", () => {
  it("inherits lang from Deck and allows hierarchical overrides at Slide, Box, Group, Table, Row, Cell, Text, P, Span, and Chart", async () => {
    const result = await build(
      <Deck style={{ lang: "zh-CN" }}>
        {/* Slide 1: inherits zh-CN from Deck */}
        <Slide id="slide-inherit">
          <Text x={50} y={50} width={300} height={100}>
            <P>
              <Span>Inherited zh-CN run</Span>
            </P>
          </Text>
          <Shape
            x={50}
            y={160}
            width={300}
            height={80}
            text="Inherited zh-CN shape"
          />
          <Table x={50} y={260} columns={[150, 150]}>
            <Row style={{ height: 35 }}>
              {/* Cell overrides to ja with colSpan=2 */}
              <Cell colSpan={2} style={{ lang: "ja" }}>
                <P>
                  <Span>Japanese cell</Span>
                </P>
              </Cell>
            </Row>
            <Row style={{ height: 35 }}>
              {/* Cell inherits zh-CN */}
              <Cell>
                <P>
                  <Span>Chinese cell</Span>
                </P>
              </Cell>
              {/* Cell overrides to ko-KR */}
              <Cell style={{ lang: "ko-KR" }}>
                <P>
                  <Span>Korean cell</Span>
                </P>
              </Cell>
            </Row>
          </Table>
          <Chart
            x={400}
            y={50}
            width={300}
            height={200}
            kind="bar"
            data={sampleChartData}
          />
        </Slide>

        {/* Slide 2: overrides lang to en-US */}
        <Slide id="slide-override" style={{ lang: "en-US" }}>
          <Box style={{ width: 320, height: 180, lang: "de-DE" }}>
            <Text x={10} y={10} width={300} height={150}>
              <P>
                <Span>German paragraph (inherited from Box)</Span>
              </P>
              <P style={{ lang: "fr-FR" }}>
                <Span>French span (inherited from P)</Span>
                <Span style={{ lang: "pt-BR" }}>
                  Portuguese span (overridden)
                </Span>
              </P>
            </Text>
          </Box>
          <Group style={{ width: 320, height: 100, lang: "es-ES" }}>
            <Text x={10} y={10} width={300} height={80}>
              Spanish run (inherited from Group)
            </Text>
          </Group>
          <Chart
            x={400}
            y={50}
            width={300}
            height={200}
            kind="bar"
            data={sampleChartData}
            style={{ lang: "it-IT" }}
          />
        </Slide>
      </Deck>,
    );

    const unzipped = unzipSync(result.pptx);
    const slide1Xml = strFromU8(unzipped["ppt/slides/slide1.xml"]!);
    const slide2Xml = strFromU8(unzipped["ppt/slides/slide2.xml"]!);
    const chart1Xml = strFromU8(unzipped["ppt/charts/chart1.xml"]!);
    const chart2Xml = strFromU8(unzipped["ppt/charts/chart2.xml"]!);

    // Slide 1 checks:
    // Text and Shape inherit zh-CN
    assert.ok(
      slide1Xml.includes('lang="zh-CN"'),
      "slide1.xml should contain inherited lang='zh-CN'",
    );
    // Table cell overrides to ja
    assert.ok(
      slide1Xml.includes('lang="ja"'),
      "slide1.xml should contain overridden lang='ja' on cell",
    );
    // Continuation emptyCell of the ja cell has <a:endParaRPr lang="ja"/>
    assert.ok(
      slide1Xml.includes('<a:endParaRPr lang="ja"/>'),
      "slide1.xml should have lang='ja' on continuation emptyCell of merged cell",
    );
    // Table cell overrides to ko-KR
    assert.ok(
      slide1Xml.includes('lang="ko-KR"'),
      "slide1.xml should contain overridden lang='ko-KR'",
    );

    // Chart 1 checks:
    // Inherits zh-CN from Deck
    assert.ok(
      chart1Xml.includes('<c:lang val="zh-CN"/>'),
      "chart1.xml should have <c:lang val='zh-CN'/> inherited from Deck",
    );
    assert.ok(
      chart1Xml.includes('lang="zh-CN"'),
      "chart1.xml should have defRPr/endParaRPr with lang='zh-CN'",
    );

    // Slide 2 checks:
    // Box inherits de-DE
    assert.ok(
      slide2Xml.includes('lang="de-DE"'),
      "slide2.xml should contain lang='de-DE' inherited from Box",
    );
    // P overrides to fr-FR
    assert.ok(
      slide2Xml.includes('lang="fr-FR"'),
      "slide2.xml should contain lang='fr-FR' overridden at P",
    );
    // Span overrides to pt-BR
    assert.ok(
      slide2Xml.includes('lang="pt-BR"'),
      "slide2.xml should contain lang='pt-BR' overridden at Span",
    );
    // Group inherits es-ES
    assert.ok(
      slide2Xml.includes('lang="es-ES"'),
      "slide2.xml should contain lang='es-ES' inherited from Group",
    );

    // Chart 2 checks:
    // Overridden to it-IT
    assert.ok(
      chart2Xml.includes('<c:lang val="it-IT"/>'),
      "chart2.xml should have <c:lang val='it-IT'/>",
    );

    // Validate entire presentation
    const validated = await validatePptx(result.pptx);
    assert.deepEqual(validated.errors, []);

    // Check FileTextRun lang on reader output
    const slide1 = validated.slides[0];
    assert.ok(slide1);
    const textObject = slide1.objects.find((obj) =>
      obj.name.startsWith("Text"),
    );
    assert.ok(textObject && "text" in textObject && textObject.text);
    assert.equal(textObject.text.paragraphs[0]?.runs[0]?.lang, "zh-CN");
  });
});

describe("artifact validator (validatePptx) lang checks", () => {
  it("rejects invalid lang attribute in slide text XML", async () => {
    const result = await build(
      <Deck style={{ lang: "zh-CN" }}>
        <Slide id="valid-slide">
          <Text x={50} y={50} width={200} height={100}>
            Hello
          </Text>
        </Slide>
      </Deck>,
    );

    const unzipped = unzipSync(result.pptx);
    let slideXml = strFromU8(unzipped["ppt/slides/slide1.xml"]!);
    // Replace valid lang with invalid lang tag
    slideXml = slideXml.replace('lang="zh-CN"', 'lang="invalid_bcp_47!"');
    unzipped["ppt/slides/slide1.xml"] = strToU8(slideXml);

    const corruptedPptx = zipSync(unzipped);
    const file = await validatePptx(corruptedPptx);
    assert.ok(
      file.errors.some((err) => err.code === "invalid-text-lang"),
      "validator should report invalid-text-lang for malformed lang attribute in rPr",
    );
  });

  it("rejects invalid c:lang in chart XML", async () => {
    const result = await build(
      <Deck style={{ lang: "zh-CN" }}>
        <Slide id="valid-chart-slide">
          <Chart
            x={50}
            y={50}
            width={300}
            height={200}
            kind="bar"
            data={sampleChartData}
          />
        </Slide>
      </Deck>,
    );

    const unzipped = unzipSync(result.pptx);
    let chartXml = strFromU8(unzipped["ppt/charts/chart1.xml"]!);
    // Replace valid c:lang with invalid lang tag
    chartXml = chartXml.replace(
      '<c:lang val="zh-CN"/>',
      '<c:lang val="bad..lang"/>',
    );
    unzipped["ppt/charts/chart1.xml"] = strToU8(chartXml);

    const corruptedPptx = zipSync(unzipped);
    const file = await validatePptx(corruptedPptx);
    assert.ok(
      file.errors.some((err) => err.code === "invalid-text-lang"),
      "validator should report invalid-text-lang for malformed c:lang attribute",
    );
  });
});
