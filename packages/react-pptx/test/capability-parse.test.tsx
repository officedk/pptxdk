import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  build,
  Box,
  BuildError,
  Chart,
  Deck,
  font,
  Group,
  Image,
  P,
  Row,
  Cell,
  Shape,
  shapeCapability,
  Slide,
  Span,
  Table,
  Text,
  textCapability,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it, vi } from "vitest";

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
]);

const sampleImage = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);

describe("Ticket 08: Object capability parse phase, container TextStyle validation, and Text pilot", () => {
  describe("Container declaration site validation for inheritable TextStyle", () => {
    it("directly rejects invalid TextStyle on Deck and points error path to deck", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540} style={{ fontSize: -4 }}>
            <Slide slideId="slide-1">
              <Box testId="empty-box" />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "invalid-font-size",
          );
          assert(diag !== undefined, "Expected invalid-font-size diagnostic");
          assert.equal(diag.path, "deck");
          return true;
        },
      );
    });

    it("directly rejects invalid color on Deck and points error path to deck", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540} style={{ color: "invalid-color" }}>
            <Slide slideId="slide-1">
              <Box testId="empty-box" />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) =>
              d.code === "invalid-color" || d.code === "unknown-color-token",
          );
          assert(diag !== undefined, "Expected invalid color diagnostic");
          assert.equal(diag.path, "deck");
          return true;
        },
      );
    });

    it("directly rejects invalid TextStyle on Slide and points error path to that slide", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="slide-with-invalid-style" style={{ fontSize: -10 }}>
              <Box testId="empty-box" />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "invalid-font-size",
          );
          assert(diag !== undefined, "Expected invalid-font-size diagnostic");
          assert.equal(diag.path, "deck/slide[slide-with-invalid-style]");
          return true;
        },
      );
    });

    it("directly rejects invalid TextStyle on Box without text children and points path to that box", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="slide-box-test">
              <Box
                testId="invalid-box"
                style={{ fontSize: 0, width: 200, height: 100 }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "invalid-font-size",
          );
          assert(diag !== undefined, "Expected invalid-font-size diagnostic");
          assert(
            diag.path.includes("box[invalid-box]"),
            `Expected path to include box[invalid-box], got: ${diag.path}`,
          );
          return true;
        },
      );
    });

    it("directly rejects invalid TextStyle on Group without text children and points path to that group", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="slide-group-test">
              <Group
                testId="invalid-group"
                style={{ fontSize: -5, width: 200, height: 100 }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "invalid-font-size",
          );
          assert(diag !== undefined, "Expected invalid-font-size diagnostic");
          assert(
            diag.path.includes("group[invalid-group]"),
            `Expected path to include group[invalid-group], got: ${diag.path}`,
          );
          return true;
        },
      );
    });
  });

  describe("Text capability parse phase, parameter/style validation, and single parse invocation", () => {
    it("rejects unknown parameters on Text", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="test-unknown-param">
              <Text
                testId="bad-param-text"
                {...({ unsupportedProp: "foo" } as any)}
              >
                Hello
              </Text>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-parameter",
          );
          assert(diag !== undefined, "Expected unknown-parameter diagnostic");
          assert(
            diag.message.includes("unsupportedProp"),
            `Expected message to mention unsupportedProp, got: ${diag.message}`,
          );
          return true;
        },
      );
    });

    it("rejects unknown style keys on Text", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="test-unknown-style">
              <Text
                testId="bad-style-text"
                style={{ invalidStyleProp: 42 } as any}
              >
                Hello
              </Text>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-style",
          );
          assert(diag !== undefined, "Expected unknown-style diagnostic");
          assert(
            diag.message.includes("invalidStyleProp"),
            `Expected message to mention invalidStyleProp, got: ${diag.message}`,
          );
          return true;
        },
      );
    });

    it("rejects non-object style on Text", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="test-invalid-style">
              <Text testId="non-obj-style" style={"not-an-object" as any}>
                Hello
              </Text>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "invalid-style",
          );
          assert(diag !== undefined, "Expected invalid-style diagnostic");
          return true;
        },
      );
    });

    it("executes textCapability.parse exactly once per Text node during build and reuses parsed result", async () => {
      const parseSpy = vi.spyOn(textCapability, "parse");
      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide slideId="single-parse-slide">
              <Box style={{ display: "flex", flexDirection: "column" }}>
                <Text
                  testId="piloted-text"
                  width={300}
                  height={60}
                  style={{
                    fontSize: 24,
                    color: "#333333",
                    padding: 10,
                    backgroundColor: "#f0f0f0",
                  }}
                >
                  <P>
                    <Span>Single parse paragraph</Span>
                  </P>
                </Text>
              </Box>
            </Slide>
          </Deck>,
        );

        assert.equal(
          parseSpy.mock.calls.length,
          1,
          `Expected parse to be called exactly 1 time, but was called ${parseSpy.mock.calls.length} times`,
        );

        const validation = await validatePptx(result.pptx);
        assert.deepEqual(validation.errors, []);
      } finally {
        parseSpy.mockRestore();
      }
    });
  });

  describe("Text layout, baseline, and positive compilation", () => {
    it("successfully compiles Text with layoutStyle, textStyle, visualStyle, and paragraphs", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide slideId="positive-text-slide">
            <Text
              testId="hero-title"
              x={50}
              y={60}
              width={500}
              height={120}
              style={{
                fontFamily: notoFont,
                fontSize: 32,
                fontWeight: "bold",
                color: "#112233",
                lineHeight: 40,
                backgroundColor: "#e8f4fd",
                borderColor: "#3388cc",
                borderWidth: 2,
                borderRadius: 8,
                padding: 16,
              }}
            >
              <P>
                <Span style={{ color: "#3388cc" }}>Prefix </Span>
                <Span>Main headline text</Span>
              </P>
            </Text>
          </Slide>
        </Deck>,
      );

      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);

      // Verify observation in build report
      const slideReport = result.report.slides[0]!;
      const textObs = slideReport.objects.find((o) => o.name === "hero-title");
      assert(textObs !== undefined, "Expected text observation hero-title");
      assert.equal(textObs.kind, "text");
      assert.equal(textObs.text, "Prefix Main headline text");
      assert.deepEqual(textObs.box, { x: 50, y: 60, width: 500, height: 120 });
    });

    it("aligns baseline when font measurement is available and reports baseline-measurement-unavailable when font is missing", async () => {
      // 1. Successful baseline alignment when font resource is available
      const goodResult = await build(
        <Deck width={960} height={540}>
          <Slide slideId="baseline-slide">
            <Box
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
              }}
            >
              <Text testId="t1" style={{ fontFamily: notoFont, fontSize: 32 }}>
                Large
              </Text>
              <Text testId="t2" style={{ fontFamily: notoFont, fontSize: 16 }}>
                Small
              </Text>
            </Box>
          </Slide>
        </Deck>,
      );
      const validation = await validatePptx(goodResult.pptx);
      assert.deepEqual(validation.errors, []);

      // 2. Fails with baseline-measurement-unavailable when font resource is unavailable for baseline measurement
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="missing-font-baseline">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                }}
              >
                <Text testId="unmeasured-text" width={200} height={50}>
                  No font bytes
                </Text>
              </Box>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "baseline-measurement-unavailable",
          );
          assert(
            diag !== undefined,
            "Expected baseline-measurement-unavailable diagnostic",
          );
          return true;
        },
      );
    });
  });

  describe("Compatibility for unmigrated capabilities in transition period", () => {
    it("builds presentation containing Shape, Image, Table, Chart, Box, and Group without regressions", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide slideId="full-suite-slide">
            <Box
              testId="wrapper-box"
              style={{
                display: "flex",
                flexDirection: "column",
                gap: 10,
                width: 900,
                height: 500,
                backgroundColor: "#ffffff",
              }}
            >
              {/* Group */}
              <Group testId="sample-group" style={{ width: 400, height: 60 }}>
                <Shape
                  width={150}
                  height={50}
                  text="Grouped Shape"
                  style={{ backgroundColor: "#2b579a" }}
                />
              </Group>

              {/* Shape */}
              <Shape
                testId="sample-shape"
                width={150}
                height={50}
                text="Direct Shape"
                style={{ backgroundColor: "#1e395b" }}
              />

              {/* Image */}
              <Image
                testId="sample-image"
                src={sampleImage}
                width={60}
                height={60}
              />

              {/* Table */}
              <Table testId="sample-table" columns={[100, 100]}>
                <Row style={{ height: 40 }}>
                  <Cell>A1</Cell>
                  <Cell>B1</Cell>
                </Row>
              </Table>

              {/* Chart */}
              <Chart
                testId="sample-chart"
                kind="bar"
                width={300}
                height={150}
                data={{
                  categories: ["Q1", "Q2"],
                  series: [{ name: "Sales", values: [10, 20] }],
                }}
              />

              {/* Text */}
              <Text
                testId="sample-text"
                width={200}
                height={40}
                style={{ fontSize: 14 }}
              >
                Compatibility check text
              </Text>
            </Box>
          </Slide>
        </Deck>,
      );

      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);

      // Verify all objects are present in report
      const objects = result.report.slides[0]!.objects;
      const testIds = objects.map((o) => o.testId).filter(Boolean);
      assert(testIds.includes("wrapper-box"));
      assert(testIds.includes("sample-shape"));
      assert(testIds.includes("sample-image"));
      assert(testIds.includes("sample-table"));
      assert(testIds.includes("sample-chart"));
      assert(testIds.includes("sample-text"));
    });
  });

  describe("Ticket 09: Migrate Shape to capability parse, fix text style on children and baseline alignment", () => {
    it("exposes parse method on shapeCapability module", () => {
      assert.equal(typeof shapeCapability.parse, "function");
    });

    it("allows Shape with <P> and <Span> children to set fontSize, color, and padding in style", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide slideId="shape-children-style-slide">
            <Shape
              testId="child-text-shape"
              x={60}
              y={70}
              width={400}
              height={150}
              style={{
                fontFamily: notoFont,
                fontSize: 24,
                color: "#ff3300",
                backgroundColor: "#f5f5f5",
                borderColor: "#333333",
                borderWidth: 2,
                padding: 12,
              }}
            >
              <P>
                <Span>Child paragraph headline</Span>
              </P>
            </Shape>
          </Slide>
        </Deck>,
      );

      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);

      const obs = result.report.slides[0]?.objects.find(
        (o) => o.testId === "child-text-shape",
      );
      assert(obs !== undefined, "Expected shape observation");
      assert.equal(obs.text, "Child paragraph headline");

      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      // Verify bodyPr has padding in EMU (12 px = 114300 EMU)
      assert.match(slideXml, /<a:bodyPr[^>]*lIns="114300"/u);
      assert.match(slideXml, /<a:bodyPr[^>]*tIns="114300"/u);
      assert.match(slideXml, /<a:bodyPr[^>]*rIns="114300"/u);
      assert.match(slideXml, /<a:bodyPr[^>]*bIns="114300"/u);
      // Verify font size (24 px = 18 pt = 1800 centipoints in DrawingML sz) and color
      assert.match(slideXml, /<a:rPr[^>]*sz="1800"/u);
      assert.match(slideXml, /<a:srgbClr val="FF3300"/u);
    });

    it("aligns Shape with <P> and <Span> children in parent flex container with alignItems: 'baseline'", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide slideId="shape-child-baseline-slide">
            <Box
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
                width: 600,
              }}
            >
              <Text
                testId="baseline-reference-text"
                width={150}
                height={100}
                style={{
                  fontFamily: notoFont,
                  fontSize: 20,
                  verticalAlign: "top",
                }}
              >
                Text Ref
              </Text>
              <Shape
                testId="baseline-child-shape"
                width={200}
                height={100}
                style={{
                  fontFamily: notoFont,
                  fontSize: 40,
                  verticalAlign: "top",
                  paddingTop: 10,
                }}
              >
                <P>
                  <Span>Shape Ref</Span>
                </P>
              </Shape>
            </Box>
          </Slide>
        </Deck>,
      );

      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);

      const objects = result.report.slides[0]?.objects ?? [];
      const textObs = objects.find(
        (o) => o.testId === "baseline-reference-text",
      );
      const shapeObs = objects.find((o) => o.testId === "baseline-child-shape");
      assert(textObs?.box !== undefined && shapeObs?.box !== undefined);
      // Because shape has larger font (40) and paddingTop (10), text should shift down to align baselines
      assert(textObs.box.y > shapeObs.box.y);
    });

    it("fails with baseline-measurement-unavailable when Shape with child text defaults to middle anchor", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="shape-child-middle-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Text
                  testId="reference-text"
                  width={150}
                  height={100}
                  style={{
                    fontFamily: notoFont,
                    fontSize: 20,
                    verticalAlign: "top",
                  }}
                >
                  Ref
                </Text>
                <Shape
                  testId="middle-anchor-child-shape"
                  width={200}
                  height={100}
                  style={{
                    fontFamily: notoFont,
                    fontSize: 40,
                  }}
                >
                  <P>
                    <Span>Middle anchor by default</Span>
                  </P>
                </Shape>
              </Box>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "baseline-measurement-unavailable",
          );
          assert(
            diag !== undefined,
            "Expected baseline-measurement-unavailable diagnostic",
          );
          assert.match(diag.path ?? "", /shape\[middle-anchor-child-shape\]/u);
          return true;
        },
      );
    });

    it("fails with unsupported-baseline-context when Shape without text is placed in baseline propagation context", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="shape-no-text-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Text
                  testId="reference-text"
                  width={150}
                  height={100}
                  style={{
                    fontFamily: notoFont,
                    fontSize: 20,
                    verticalAlign: "top",
                  }}
                >
                  Ref
                </Text>
                {/* Single-child Box chain wrapping an empty Shape */}
                <Box
                  testId="wrapper-for-empty-shape"
                  style={{
                    display: "flex",
                    flexDirection: "row",
                    alignSelf: "baseline",
                  }}
                >
                  <Shape testId="textless-shape" width={100} height={100} />
                </Box>
              </Box>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unsupported-baseline-context",
          );
          assert(
            diag !== undefined,
            "Expected unsupported-baseline-context diagnostic for textless shape",
          );
          return true;
        },
      );
    });

    it("rejects unknown top-level parameters on Shape", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="unknown-param-slide">
              <Shape
                testId="invalid-shape"
                width={100}
                height={100}
                {...({ unsupportedParameter: 42 } as any)}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-parameter",
          );
          assert(diag !== undefined, "Expected unknown-parameter diagnostic");
          assert.match(
            diag.message,
            /Shape contains unsupported parameter unsupportedParameter/u,
          );
          return true;
        },
      );
    });

    it("rejects unsupported style properties on Shape", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="unknown-style-slide">
              <Shape
                testId="invalid-style-shape"
                width={100}
                height={100}
                style={{ unsupportedStyleProp: "red" } as any}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-style",
          );
          assert(diag !== undefined, "Expected unknown-style diagnostic");
          assert.match(
            diag.message,
            /style contains unsupported property unsupportedStyleProp/u,
          );
          return true;
        },
      );
    });

    it("rejects TextStyle and padding in style when Shape bears no text", async () => {
      // 1. fontSize on Shape without text
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="textless-shape-font-slide">
              <Shape
                testId="textless-font-shape"
                width={100}
                height={100}
                style={{ fontSize: 24 }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-style",
          );
          assert(
            diag !== undefined,
            "Expected unknown-style diagnostic for fontSize on textless shape",
          );
          return true;
        },
      );

      // 2. padding on Shape without text
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide slideId="textless-shape-padding-slide">
              <Shape
                testId="textless-padding-shape"
                width={100}
                height={100}
                style={{ padding: 16 }}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) =>
              d.code === "invalid-layout-context" || d.code === "unknown-style",
          );
          assert(
            diag !== undefined,
            "Expected invalid padding diagnostic on textless shape",
          );
          return true;
        },
      );
    });

    it("calls shapeCapability.parse once during layout and reuses parsed result in measureLayout and compile", async () => {
      const parseSpy = vi.spyOn(shapeCapability, "parse");
      const compileSpy = vi.spyOn(shapeCapability, "compile");

      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide slideId="parse-reuse-slide">
              <Shape
                testId="reused-shape"
                x={100}
                y={100}
                width={200}
                height={100}
                style={{
                  fontFamily: notoFont,
                  fontSize: 20,
                  verticalAlign: "top",
                  padding: 8,
                }}
              >
                <P>
                  <Span>Parsed once, reused twice</Span>
                </P>
              </Shape>
            </Slide>
          </Deck>,
        );

        assert.deepEqual(result.report.diagnostics, []);
        assert.equal(
          parseSpy.mock.calls.length,
          1,
          "parse should be called exactly once",
        );

        // Verify that compile received the parsed object produced by parse
        const parseReturn = parseSpy.mock.results[0]!.value as {
          parsed: unknown;
        };
        assert.equal(compileSpy.mock.calls[0]![0], parseReturn.parsed);
      } finally {
        parseSpy.mockRestore();
        compileSpy.mockRestore();
      }
    });
  });
});
