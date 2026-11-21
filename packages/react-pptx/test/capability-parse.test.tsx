import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { validatePptx } from "@pptxdk/validation";
import {
  boxCapability,
  build,
  Box,
  BuildError,
  createMorphScope,
  Deck,
  font,
  Group,
  groupCapability,
  P,
  Shape,
  shapeCapability,
  Slide,
  Span,
  Text,
  textCapability,
  transitions,
} from "@react-pptx/react-pptx";
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

describe("Ticket 08: Object capability parse phase, container TextStyle validation, and Text pilot", () => {
  describe("Container declaration site validation for inheritable TextStyle", () => {
    it("directly rejects invalid TextStyle on Deck and points error path to deck", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540} style={{ fontSize: -4 }}>
            <Slide id="slide-1">
              <Box id="empty-box" />
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
            <Slide id="slide-1">
              <Box id="empty-box" />
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
            <Slide id="slide-with-invalid-style" style={{ fontSize: -10 }}>
              <Box id="empty-box" />
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
            <Slide id="slide-box-test">
              <Box
                id="invalid-box"
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
            <Slide id="slide-group-test">
              <Group
                id="invalid-group"
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
            <Slide id="test-unknown-param">
              <Text
                id="bad-param-text"
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
            <Slide id="test-unknown-style">
              <Text id="bad-style-text" style={{ invalidStyleProp: 42 } as any}>
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
            <Slide id="test-invalid-style">
              <Text id="non-obj-style" style={"not-an-object" as any}>
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

    it("executes textCapability.bind exactly once per Text node during build and reuses bound handle", async () => {
      const bindSpy = vi.spyOn(textCapability, "bind");
      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide id="single-parse-slide">
              <Box style={{ display: "flex", flexDirection: "column" }}>
                <Text
                  id="piloted-text"
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
          bindSpy.mock.calls.length,
          1,
          `Expected bind to be called exactly 1 time, but was called ${bindSpy.mock.calls.length} times`,
        );

        const validation = await validatePptx(result.pptx);
        assert.deepEqual(validation.errors, []);
      } finally {
        bindSpy.mockRestore();
      }
    });
  });

  describe("Text layout, baseline, and positive compilation", () => {
    it("successfully compiles Text with layoutStyle, textStyle, visualStyle, and paragraphs", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="positive-text-slide">
            <Text
              id="hero-title"
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
      const textObs = slideReport.objects.find(
        (o) => o.authorId === "hero-title",
      );
      assert(textObs !== undefined, "Expected text observation hero-title");
      assert.equal(textObs.kind, "text");
      assert.equal(textObs.text, "Prefix Main headline text");
      assert.deepEqual(textObs.box, { x: 50, y: 60, width: 500, height: 120 });
    });

    it("aligns baseline when font measurement is available and reports baseline-measurement-unavailable when font is missing", async () => {
      // 1. Successful baseline alignment when font resource is available
      const goodResult = await build(
        <Deck width={960} height={540}>
          <Slide id="baseline-slide">
            <Box
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
              }}
            >
              <Text id="t1" style={{ fontFamily: notoFont, fontSize: 32 }}>
                Large
              </Text>
              <Text id="t2" style={{ fontFamily: notoFont, fontSize: 16 }}>
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
            <Slide id="missing-font-baseline">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                }}
              >
                <Text id="unmeasured-text" width={200} height={50}>
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

  describe("Ticket 09: Migrate Shape to capability parse, fix text style on children and baseline alignment", () => {
    it("allows Shape with <P> and <Span> children to set fontSize, color, and padding in style", async () => {
      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="shape-children-style-slide">
            <Shape
              id="child-text-shape"
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
        (o) => o.authorId === "child-text-shape",
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
          <Slide id="shape-child-baseline-slide">
            <Box
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "baseline",
                width: 600,
              }}
            >
              <Text
                id="baseline-reference-text"
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
                id="baseline-child-shape"
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
        (o) => o.authorId === "baseline-reference-text",
      );
      const shapeObs = objects.find(
        (o) => o.authorId === "baseline-child-shape",
      );
      assert(textObs?.box !== undefined && shapeObs?.box !== undefined);
      // Because shape has larger font (40) and paddingTop (10), text should shift down to align baselines
      assert(textObs.box.y > shapeObs.box.y);
    });

    it("fails with baseline-measurement-unavailable when Shape with child text defaults to middle anchor", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="shape-child-middle-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Text
                  id="reference-text"
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
                  id="middle-anchor-child-shape"
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
            <Slide id="shape-no-text-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  flexDirection: "row",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Text
                  id="reference-text"
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
                  id="wrapper-for-empty-shape"
                  style={{
                    display: "flex",
                    flexDirection: "row",
                    alignSelf: "baseline",
                  }}
                >
                  <Shape id="textless-shape" width={100} height={100} />
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

    it("fails with baseline-measurement-unavailable when Text with verticalAlign: 'middle' is placed in alignItems: 'baseline' container", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="text-middle-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Text
                  id="reference-text"
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
                <Text
                  id="middle-text"
                  width={150}
                  height={100}
                  style={{
                    fontFamily: notoFont,
                    fontSize: 20,
                    verticalAlign: "middle",
                  }}
                >
                  Middle
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
            "Expected baseline-measurement-unavailable diagnostic for middle-aligned text in baseline",
          );
          return true;
        },
      );
    });

    it("fails with unsupported-baseline-context when Shape without text is placed directly in alignItems: 'baseline' container", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="shape-direct-no-text-baseline-slide">
              <Box
                style={{
                  display: "flex",
                  alignItems: "baseline",
                  width: 500,
                }}
              >
                <Shape id="direct-textless-shape" width={100} height={100} />
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
            "Expected unsupported-baseline-context diagnostic for direct textless shape in baseline",
          );
          return true;
        },
      );
    });

    it("rejects unknown top-level parameters on Shape", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="unknown-param-slide">
              <Shape
                id="invalid-shape"
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
            <Slide id="unknown-style-slide">
              <Shape
                id="invalid-style-shape"
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
            <Slide id="textless-shape-font-slide">
              <Shape
                id="textless-font-shape"
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
            <Slide id="textless-shape-padding-slide">
              <Shape
                id="textless-padding-shape"
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

    it("calls shapeCapability.bind once during layout and reuses the bound handle for measureLayout and resolve", async () => {
      const bindSpy = vi.spyOn(shapeCapability, "bind");

      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide id="parse-reuse-slide">
              <Shape
                id="reused-shape"
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
          bindSpy.mock.calls.length,
          1,
          "bind should be called exactly once",
        );
      } finally {
        bindSpy.mockRestore();
      }
    });
  });

  describe("Ticket 11: Unify Box and Group, tighten parse contract, and remove leaf coupling", () => {
    it("allows Group in flex layout to use morphKey and link", async () => {
      const morph = createMorphScope();
      const groupKey = morph.handle("flex-group-key");

      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="slide-flex-group">
            <Box style={{ display: "flex", width: 500, height: 300, gap: 10 }}>
              <Group
                id="flex-group"
                morphKey={groupKey}
                link={{ url: "https://example.com/target" }}
                style={{ width: 200, height: 150 }}
              >
                <Shape id="inner-shape" width={100} height={80} />
              </Group>
            </Box>
          </Slide>
        </Deck>,
      );

      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);

      // Verify the slide XML contains <p:grpSp> with the morph name and hyperlink
      const unzipped = unzipSync(result.pptx);
      const slideXml = strFromU8(unzipped["ppt/slides/slide1.xml"]!);
      assert.match(slideXml, /<p:grpSp>/u);
      assert.match(slideXml, /name="!*react-pptx-morph-/u);
      assert.match(slideXml, /<a:hlinkClick/u);
    });

    it("supports cross-slide Morph pairing between Group objects based on nativeKind", async () => {
      const morph = createMorphScope();
      const groupKey = morph.handle("morphing-group");

      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="slide-1">
            <Box style={{ display: "flex", width: 400, height: 300 }}>
              <Group
                id="group-start"
                morphKey={groupKey}
                style={{ width: 150, height: 100 }}
              >
                <Shape width={50} height={50} />
              </Group>
            </Box>
          </Slide>
          <Slide
            id="slide-2"
            transition={transitions.morph({ fallback: transitions.fade() })}
          >
            <Box style={{ display: "flex", width: 600, height: 400 }}>
              <Group
                id="group-end"
                morphKey={groupKey}
                style={{ width: 250, height: 200 }}
              >
                <Shape width={80} height={80} />
              </Group>
            </Box>
          </Slide>
        </Deck>,
      );

      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
    });

    it("rejects Morph pairing between objects with different nativeKind (Group grpSp vs Shape sp)", async () => {
      const morph = createMorphScope();
      const mismatchedKey = morph.handle("mismatched-key");

      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="slide-group">
              <Group morphKey={mismatchedKey}>
                <Shape x={0} y={0} width={50} height={50} />
              </Group>
            </Slide>
            <Slide
              id="slide-shape"
              transition={transitions.morph({ fallback: transitions.fade() })}
            >
              <Shape
                x={0}
                y={0}
                morphKey={mismatchedKey}
                width={80}
                height={80}
              />
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "morph-object-kind-mismatch",
          );
          assert(
            diag !== undefined,
            "Expected morph-object-kind-mismatch error",
          );
          assert.match(diag.message, /grpSp/u);
          assert.match(diag.message, /sp/u);
          return true;
        },
      );
    });

    it("allows Morph pairing between Shape and Text because both provide nativeKind sp", async () => {
      const morph = createMorphScope();
      const sharedSpKey = morph.handle("shared-sp");

      const result = await build(
        <Deck width={960} height={540}>
          <Slide id="slide-shape-start">
            <Shape
              x={10}
              y={10}
              morphKey={sharedSpKey}
              width={100}
              height={50}
            />
          </Slide>
          <Slide
            id="slide-text-end"
            transition={transitions.morph({ fallback: transitions.fade() })}
          >
            <Text
              x={10}
              y={10}
              morphKey={sharedSpKey}
              width={400}
              height={100}
              style={{ fontFamily: notoFont }}
            >
              Morphing between Shape and Text
            </Text>
          </Slide>
        </Deck>,
      );

      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
    });

    it("rejects unknown parameters on Group", async () => {
      await assert.rejects(
        build(
          <Deck width={960} height={540}>
            <Slide id="slide-invalid-group">
              {/* @ts-expect-error -- testing unknown parameter */}
              <Group id="bad-group" unknownProp={123}>
                <Shape width={50} height={50} />
              </Group>
            </Slide>
          </Deck>,
        ),
        (error: unknown) => {
          assert(error instanceof BuildError);
          const diag = error.diagnostics.find(
            (d) => d.code === "unknown-parameter",
          );
          assert(
            diag !== undefined,
            "Expected unknown-parameter for unknownProp",
          );
          return true;
        },
      );
    });

    it("calls boxCapability.bind once during layout and reuses the bound handle in resolve", async () => {
      const bindSpy = vi.spyOn(boxCapability, "bind");

      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide id="parse-box-slide">
              <Box
                id="visible-box"
                style={{
                  width: 200,
                  height: 100,
                  backgroundColor: "#2B579A",
                }}
              >
                <Shape width={50} height={50} />
              </Box>
            </Slide>
          </Deck>,
        );

        assert.deepEqual(result.report.diagnostics, []);
        assert.equal(
          bindSpy.mock.calls.length,
          1,
          "boxCapability.bind should be called exactly once",
        );
      } finally {
        bindSpy.mockRestore();
      }
    });

    it("calls groupCapability.bind once during layout and reuses the bound handle in resolve", async () => {
      const bindSpy = vi.spyOn(groupCapability, "bind");

      try {
        const result = await build(
          <Deck width={960} height={540}>
            <Slide id="parse-group-slide">
              <Box style={{ display: "flex", width: 400, height: 200 }}>
                <Group id="parsed-group" style={{ width: 100, height: 80 }}>
                  <Shape width={40} height={40} />
                </Group>
              </Box>
            </Slide>
          </Deck>,
        );

        assert.deepEqual(result.report.diagnostics, []);
        assert.equal(
          bindSpy.mock.calls.length,
          1,
          "groupCapability.bind should be called exactly once",
        );
      } finally {
        bindSpy.mockRestore();
      }
    });
  });
});
