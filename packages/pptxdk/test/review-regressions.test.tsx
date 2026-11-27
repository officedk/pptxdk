import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  build,
  codegenDiagnosticSchema,
  defineSlideLayout,
  defineTheme,
  font,
  importPptx,
  previewErrorSchema,
  resolve,
} from "@officedk/pptxdk";
import { resolve as resolveTrimmed } from "@officedk/pptxdk/build";
import { preview } from "@officedk/pptxdk/preview";
import {
  Box,
  Cell,
  Deck,
  expandDocument,
  Image,
  Placeholder,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Table,
  Text,
} from "@officedk/pptxdk/react";
import { CommonContainer, CommonNode } from "@pptxdk/core/document";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { describe, expect, it } from "vitest";

const jostRegularBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

function makeCorruptPng(): Uint8Array {
  const valid = PNG.sync.write(new PNG({ width: 2, height: 2 }));
  const bytes = new Uint8Array(valid);
  // Corrupt the IDAT compressed payload while keeping the PNG header/IHDR valid for resolve().
  const idatIndex = Buffer.from(bytes).indexOf(Buffer.from("IDAT"));
  assert.ok(idatIndex > 0);
  bytes[idatIndex + 4] = 0xff;
  bytes[idatIndex + 5] = 0xff;
  return bytes;
}

describe("two-axis code review regressions", () => {
  it("preserves exact font family/weight/style and full >10 missing-glyph character lists on Table diagnostics", async () => {
    const doc1 = await expandDocument(
      <Deck width={960} height={540}>
        <Slide id="tbl-missing-bold">
          <Table id="t1" columns={["auto"]}>
            <Row>
              <Cell
                style={{
                  fontFamily: "UnbundledCorp",
                  fontWeight: 700,
                  fontStyle: "italic",
                  fontSize: 18,
                }}
              >
                Quarterly Revenue
              </Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
    );
    assert.throws(
      () => resolve(doc1),
      (error: unknown) => {
        assert(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-unavailable");
        if (diag?.code === "font-unavailable") {
          assert.equal(diag.reason, "missing-resource");
          assert.deepEqual(diag.font, {
            family: "UnbundledCorp",
            weight: 700,
            style: "italic",
          });
        }
        return true;
      },
    );

    const jostOnly = font("Jost", [{ weight: 400, data: jostRegularBytes }], {
      substitute: true,
    });
    const twelveMissingChars = [
      "一",
      "二",
      "三",
      "四",
      "五",
      "六",
      "七",
      "八",
      "九",
      "十",
      "百",
      "千",
    ];

    const doc2 = await expandDocument(
      <Deck width={960} height={540}>
        <Slide id="tbl-missing-glyphs">
          <Table id="t2" columns={[240]}>
            <Row>
              <Cell style={{ fontFamily: jostOnly, fontSize: 16 }}>
                {twelveMissingChars.join("")}
              </Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
    );
    assert.throws(
      () => resolve(doc2),
      (error: unknown) => {
        assert(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-missing-glyph");
        if (diag?.code === "font-missing-glyph") {
          assert.deepEqual(diag.font, {
            family: "Jost",
            weight: 400,
            style: "normal",
          });
          assert.deepEqual(diag.characters, twelveMissingChars);
        }
        return true;
      },
    );
  });

  it("reports font-unavailable instead of invalid-structure when baseline-aligned Text or Shape lacks font bytes", async () => {
    const docText = await expandDocument(
      <Deck width={480} height={270}>
        <Slide id="baseline-text">
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              width: 400,
              height: 100,
            }}
          >
            <Text
              id="b-text"
              style={{
                width: 150,
                height: 40,
                fontFamily: "MissingBaselineSans",
                fontWeight: 700,
                fontSize: 16,
              }}
            >
              Baseline Text
            </Text>
          </Box>
        </Slide>
      </Deck>,
    );
    assert.throws(
      () => resolve(docText),
      (error: unknown) => {
        assert(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-unavailable");
        if (diag?.code === "font-unavailable") {
          assert.equal(diag.reason, "missing-resource");
          assert.equal(diag.font.family, "MissingBaselineSans");
          assert.equal(diag.font.weight, 700);
        }
        return true;
      },
    );

    const docShape = await expandDocument(
      <Deck width={480} height={270}>
        <Slide id="baseline-shape">
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              width: 400,
              height: 100,
            }}
          >
            <Shape
              id="b-shape"
              style={{
                width: 150,
                height: 40,
                verticalAlign: "top",
                fontFamily: "MissingBaselineShapeSans",
                fontSize: 16,
              }}
            >
              Baseline Shape
            </Shape>
          </Box>
        </Slide>
      </Deck>,
    );
    assert.throws(
      () => resolve(docShape),
      (error: unknown) => {
        assert(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "font-unavailable");
        if (diag?.code === "font-unavailable") {
          assert.equal(diag.reason, "missing-resource");
          assert.equal(diag.font.family, "MissingBaselineShapeSans");
        }
        return true;
      },
    );
  });

  it("treats %-width fixed-height and absolute top+bottom/inset Text and Shape as fixed-region unmeasured text", async () => {
    const jostOnly = font("Jost", [{ weight: 400, data: jostRegularBytes }], {
      substitute: true,
    });

    const fixedDoc = await expandDocument(
      <Deck width={800} height={500}>
        <Slide id="fixed-regions">
          <Shape
            id="pct-width-shape"
            style={{
              width: "50%",
              height: 60,
              fontFamily: "UnbundledFixedSans",
              fontSize: 18,
            }}
          >
            Percent width fixed height shape
          </Shape>
          <Shape
            id="abs-tb-shape"
            style={{
              position: "absolute",
              left: 40,
              right: 40,
              top: 120,
              bottom: 280,
              fontFamily: "UnbundledFixedSans",
              fontSize: 18,
            }}
          >
            Pinned top and bottom shape
          </Shape>
          <Text
            id="pct-width-text"
            style={{
              width: "25%",
              height: 48,
              fontFamily: "UnbundledFixedSans",
              fontSize: 16,
            }}
          >
            Percent width fixed height text
          </Text>
          <Text
            id="abs-inset-text"
            style={{
              position: "absolute",
              inset: 50,
              fontFamily: jostOnly,
              fontSize: 16,
            }}
          >
            中文固定区域
          </Text>
        </Slide>
      </Deck>,
    );
    const resolution = resolve(fixedDoc);

    const slideLoc = resolution.slide("fixed-regions");
    assert.deepEqual(slideLoc.get("pct-width-shape").box, {
      x: 0,
      y: 0,
      width: 400,
      height: 60,
    });
    assert.equal(
      slideLoc.get("pct-width-shape").textMetrics?.status,
      "unmeasured",
    );

    assert.deepEqual(slideLoc.get("abs-tb-shape").box, {
      x: 40,
      y: 120,
      width: 720,
      height: 100,
    });
    assert.equal(
      slideLoc.get("abs-tb-shape").textMetrics?.status,
      "unmeasured",
    );

    assert.deepEqual(slideLoc.get("pct-width-text").box, {
      x: 0,
      y: 60,
      width: 200,
      height: 48,
    });
    assert.equal(
      slideLoc.get("pct-width-text").textMetrics?.status,
      "unmeasured",
    );

    assert.deepEqual(slideLoc.get("abs-inset-text").box, {
      x: 50,
      y: 50,
      width: 700,
      height: 400,
    });
    const insetMetrics = slideLoc.get("abs-inset-text").textMetrics;
    assert.equal(insetMetrics?.status, "unmeasured");
    if (insetMetrics?.status === "unmeasured") {
      assert.equal(insetMetrics.reason, "missing-glyphs");
    }
  });

  it("throws missing-capability on Slide.style.background when slide-background capability is omitted", async () => {
    await assert.rejects(
      () =>
        resolveTrimmed(
          <Deck width={640} height={360}>
            <Slide id="bg-trimmed" style={{ background: "#0F172A" }}>
              <Shape
                id="box"
                style={{
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 100,
                  height: 60,
                  background: "#38BDF8",
                }}
              />
            </Slide>
          </Deck>,
          { capabilities: ["shape"] },
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "missing-capability");
        if (diag?.code === "missing-capability") {
          assert.equal(diag.capability, "slide-background");
          assert.equal(diag.field, "style.background");
        }
        return true;
      },
    );
  });

  it("exposes optional field on PreviewError and CodegenDiagnostic schemas and populates image preview error fields", async () => {
    const corruptPng = makeCorruptPng();

    const previewed = await preview(
      <Deck width={400} height={240}>
        <Slide id="corrupt-img-src">
          <Image
            id="img1"
            src={corruptPng}
            style={{
              position: "absolute",
              left: 20,
              top: 20,
              width: 80,
              height: 60,
            }}
          />
        </Slide>
        <Slide id="corrupt-shape-bg">
          <Shape
            id="sh1"
            style={{
              position: "absolute",
              left: 20,
              top: 20,
              width: 80,
              height: 60,
              background: { image: corruptPng },
            }}
          />
        </Slide>
        <Slide
          id="corrupt-slide-bg"
          style={{ background: { image: corruptPng } }}
        />
      </Deck>,
    );

    const [p1, p2, p3] = previewed.pages;
    expect(previewErrorSchema.parse(p1?.errors[0])).toEqual(p1?.errors[0]);
    expect(p1?.errors[0]?.field).toBe("src");

    expect(previewErrorSchema.parse(p2?.errors[0])).toEqual(p2?.errors[0]);
    expect(p2?.errors[0]?.field).toBe("style.background");

    expect(previewErrorSchema.parse(p3?.errors[0])).toEqual(p3?.errors[0]);
    expect(p3?.errors[0]?.field).toBe("style.background");

    const parsedCodegenDiag = codegenDiagnosticSchema.parse({
      code: "unsupported-declaration",
      path: "deck/slide[s1]/widget[w1]",
      field: "children",
      declaration: "widget",
      message: "Unsupported widget declaration",
    });
    expect(parsedCodegenDiag.field).toBe("children");
  });

  it("emits and roundtrips placeholder textShadow/boxShadow/background 'none', flipH: false, and content Placeholder style overrides", async () => {
    const jost = font("Jost", [{ weight: 400, data: jostRegularBytes }], {
      substitute: true,
    });
    const theme = defineTheme({
      name: "PhClearTheme",
      fonts: { major: jost, minor: jost },
    });

    const clearedLayout = defineSlideLayout({
      id: "cleared-layout",
      name: "Cleared Layout",
      type: "obj",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 40,
            top: 30,
            width: 600,
            height: 60,
            background: "none",
            boxShadow: "none",
            textShadow: "none",
            flipH: false,
          },
        },
        content: {
          type: "obj",
          idx: 1,
          style: {
            position: "absolute",
            left: 40,
            top: 120,
            width: 600,
            height: 300,
          },
        },
      },
    });

    const built = await build(
      <Deck width={960} height={540} theme={theme}>
        <SlideMaster
          id="m1"
          placeholders={{
            title: {
              type: "title",
              style: {
                position: "absolute",
                left: 40,
                top: 30,
                width: 600,
                height: 60,
                background: "#1E293B",
                boxShadow: {
                  color: "#000000",
                  blur: 6,
                  distance: 4,
                  angle: 90,
                },
                textShadow: {
                  color: "#000000",
                  blur: 4,
                  distance: 2,
                  angle: 90,
                },
                flipH: true,
              },
            },
            body: {
              type: "body",
              style: {
                position: "absolute",
                left: 40,
                top: 120,
                width: 600,
                height: 300,
                background: "#E2E8F0",
                boxShadow: {
                  color: "#000000",
                  blur: 8,
                  distance: 4,
                  angle: 90,
                },
              },
            },
          }}
        >
          <SlideLayout layout={clearedLayout} />
        </SlideMaster>
        <Slide id="s1" layout={clearedLayout}>
          <Text
            id="slide-title"
            placeholder="title"
            style={{
              textShadow: "none",
              boxShadow: "none",
              background: "none",
              flipH: false,
            }}
          >
            Cleared Heading
          </Text>
          <Placeholder
            id="slide-content"
            placeholder="content"
            style={{
              background: "none",
              boxShadow: "none",
              border: { color: "#2563EB", width: 2 },
            }}
          />
          <Text
            id="free-text"
            style={{
              position: "absolute",
              left: 40,
              top: 450,
              width: 300,
              height: 40,
              textShadow: "none",
            }}
          >
            Free Text Without Empty EffectLst
          </Text>
        </Slide>
      </Deck>,
    );

    const validation = await validatePptx(built.pptx);
    assert.deepEqual(validation.errors, []);

    const zip = unzipSync(built.pptx);
    const masterXml = strFromU8(zip["ppt/slideMasters/slideMaster1.xml"]!);
    const layoutXml = strFromU8(zip["ppt/slideLayouts/slideLayout1.xml"]!);
    const slideXml = strFromU8(zip["ppt/slides/slide1.xml"]!);

    // Master placeholder defRPr emits outerShdw; layout placeholder defRPr clears it with <a:effectLst/>.
    expect(masterXml).toContain("<a:outerShdw");
    expect(layoutXml).toContain("<a:effectLst/>");
    // Slide title placeholder differential rPr clears textShadow with <a:effectLst/>, while free text omits empty <a:effectLst/>.
    expect(slideXml).toContain("Cleared Heading");
    const titlePhSp = slideXml
      .split("<p:sp>")
      .find((s) => s.includes("Cleared Heading"));
    assert.ok(titlePhSp !== undefined);
    expect(titlePhSp).toContain("<a:rPr><a:effectLst/></a:rPr>");

    const freeTextSp = slideXml
      .split("<p:sp>")
      .find((s) => s.includes("Free Text Without Empty EffectLst"));
    assert.ok(freeTextSp !== undefined);
    expect(freeTextSp).not.toContain("<a:effectLst/>");

    const imported = importPptx(built.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedMaster = imported.document.children.find(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide-master",
    )!;
    const importedLayout = importedMaster.children.find(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide-layout",
    )!;
    const layoutSpec = (importedLayout.props["layout"] ??
      importedLayout.props) as {
      readonly placeholders?: Record<
        string,
        {
          readonly style?: Record<string, unknown>;
          readonly listStyles?: Record<number, Record<string, unknown>>;
        }
      >;
    };
    expect(layoutSpec.placeholders?.["title"]?.style).toMatchObject({
      background: "none",
      boxShadow: "none",
      flipH: false,
    });
    expect(layoutSpec.placeholders?.["title"]?.listStyles?.[1]).toMatchObject({
      textShadow: "none",
    });

    const importedSlide = imported.document.children.find(
      (c): c is CommonContainer =>
        c instanceof CommonContainer && c.kind === "slide",
    )!;
    const importedTitle = importedSlide.children.find(
      (c): c is CommonContainer =>
        c instanceof CommonContainer &&
        c.kind === "text" &&
        c.props["placeholder"] === "title",
    );
    expect(
      importedTitle?.props["style"] as Record<string, unknown>,
    ).toMatchObject({
      background: "none",
      boxShadow: "none",
      textShadow: "none",
      flipH: false,
    });

    const importedContentPh = importedSlide.children.find(
      (c): c is CommonNode =>
        c instanceof CommonNode && c.kind === "placeholder",
    );
    expect(
      importedContentPh?.props["style"] as Record<string, unknown>,
    ).toMatchObject({
      background: "none",
      boxShadow: "none",
      border: { color: "#2563EB", width: 2 },
    });
  });
});
