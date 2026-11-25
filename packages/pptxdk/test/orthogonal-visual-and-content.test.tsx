import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

import {
  BuildError,
  build,
  codegen,
  defineSlideLayout,
  font,
  importPptx,
  linearGradient,
  radialGradient,
  spin,
  spinOptionsSchema,
} from "@officedk/pptxdk";
import { preview } from "@officedk/pptxdk/preview";
import {
  Box,
  Cell,
  Deck,
  Group,
  Image,
  P,
  Placeholder,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Span,
  Table,
  Text,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  deck,
  group,
  image,
  paragraph,
  placeholder,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  span,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [
  { weight: 400, data: jostBookBytes },
  { weight: 700, data: jostBoldBytes },
]);

const samplePng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

const tempAssetDir = mkdtempSync(join(tmpdir(), "pptxdk-t04-"));
const pixelPngPath = join(tempAssetDir, "pixel.png");
writeFileSync(pixelPngPath, samplePng);

afterAll(() => {
  rmSync(tempAssetDir, { recursive: true, force: true });
});

describe("Ticket 04: orthogonal visual and content props", () => {
  it("supports unified style.background across Slide, Master/Layout, canvas objects, placeholders, and Table/Row/Cell including Shape gradients and image fit/position/opacity", async () => {
    const grad = linearGradient(90, "#112233", "#445566");
    const radGrad = radialGradient("#FEF3C7", "#F59E0B");
    const fileUrl = pathToFileURL(pixelPngPath);

    const layoutDef = defineSlideLayout({
      id: "bg-layout",
      style: { background: "#F8FAFC" },
      placeholders: {
        hero: {
          type: "body",
          style: {
            position: "absolute",
            left: 40,
            top: 280,
            width: 240,
            height: 80,
            background: linearGradient(45, "#0EA5E9", "#2563EB"),
          },
        },
        slotPic: {
          type: "pic",
          style: {
            position: "absolute",
            left: 300,
            top: 280,
            width: 140,
            height: 80,
            background: "#E2E8F0",
          },
        },
      },
    });

    const jsxDoc = (
      <Deck width={960} height={540} theme={{ fonts: { minor: jostFont } }}>
        <SlideMaster id="master" style={{ background: "#0F172A" }}>
          <SlideLayout layout={layoutDef} />
        </SlideMaster>
        <Slide
          id="slide-bg-objects"
          layout={layoutDef}
          style={{ background: "#FFFFFF" }}
        >
          <Shape
            id="grad-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 30,
              width: 180,
              height: 80,
              background: grad,
            }}
          />
          <Box
            id="img-box"
            style={{
              position: "absolute",
              left: 240,
              top: 30,
              width: 160,
              height: 80,
              background: {
                image: samplePng,
                fit: "cover",
                position: "top left",
                opacity: 0.6,
              },
            }}
          />
          <Text
            id="radial-text"
            style={{
              position: "absolute",
              left: 420,
              top: 30,
              width: 180,
              height: 80,
              background: radGrad,
            }}
          >
            Radial
          </Text>
          <Image
            id="bg-image"
            src={samplePng}
            style={{
              position: "absolute",
              left: 620,
              top: 30,
              width: 100,
              height: 80,
              background: {
                image: fileUrl,
                fit: "contain",
                position: { x: "25%", y: "75%" },
                opacity: 0.8,
              },
            }}
          />
          <Table
            id="bg-table"
            columns={[120, 120]}
            style={{
              position: "absolute",
              left: 40,
              top: 130,
              width: 240,
              background: "#F1F5F9",
            }}
          >
            <Row
              style={{
                height: 40,
                background: linearGradient(0, "#E0F2FE", "#BAE6FD"),
              }}
            >
              <Cell id="c1">Inherited Row Grad</Cell>
              <Cell id="c2" style={{ background: "#FEF3C7" }}>
                Cell Solid Override
              </Cell>
            </Row>
            <Row style={{ height: 40 }}>
              <Cell
                id="c3"
                style={{ background: { image: samplePng, opacity: 0.5 } }}
              >
                Cell Image
              </Cell>
              <Cell id="c4" style={{ background: "none" }}>
                Cleared
              </Cell>
            </Row>
          </Table>
          <Text
            id="hero-ph"
            placeholder="hero"
            style={{ background: "#DCFCE7" }}
          >
            Solid overrides layout gradient
          </Text>
          <Placeholder
            id="pic-ph"
            placeholder="slotPic"
            style={{ background: "none" }}
          />
        </Slide>
        <Slide
          id="slide-cleared-bg"
          layout={layoutDef}
          style={{ background: "none" }}
        >
          <Shape
            id="ext-bg-shape"
            style={{
              position: "absolute",
              left: 220,
              top: 40,
              width: 160,
              height: 80,
              background: {
                image: { externalUrl: "https://example.com/bg.png" },
                fit: "cover",
                opacity: 0.75,
              },
            }}
          />
        </Slide>
      </Deck>
    );

    const vanMaster = slideMaster({
      id: "master",
      style: { background: "#0F172A" },
    }).append(slideLayout(layoutDef));

    const vanSlide1 = slide({
      id: "slide-bg-objects",
      layout: layoutDef,
      style: { background: "#FFFFFF" },
    }).append(
      shape({
        id: "grad-shape",
        geometry: "roundRect",
        style: {
          position: "absolute",
          left: 40,
          top: 30,
          width: 180,
          height: 80,
          background: grad,
        },
      }),
      box({
        id: "img-box",
        style: {
          position: "absolute",
          left: 240,
          top: 30,
          width: 160,
          height: 80,
          background: {
            image: samplePng,
            fit: "cover",
            position: "top left",
            opacity: 0.6,
          },
        },
      }),
      text("Radial", {
        id: "radial-text",
        style: {
          position: "absolute",
          left: 420,
          top: 30,
          width: 180,
          height: 80,
          background: radGrad,
        },
      }),
      image({
        id: "bg-image",
        src: samplePng,
        style: {
          position: "absolute",
          left: 620,
          top: 30,
          width: 100,
          height: 80,
          background: {
            image: fileUrl,
            fit: "contain",
            position: { x: "25%", y: "75%" },
            opacity: 0.8,
          },
        },
      }),
      table({
        id: "bg-table",
        columns: [120, 120],
        style: {
          position: "absolute",
          left: 40,
          top: 130,
          width: 240,
          background: "#F1F5F9",
        },
      }).append(
        row({
          style: {
            height: 40,
            background: linearGradient(0, "#E0F2FE", "#BAE6FD"),
          },
        }).append(
          cell("Inherited Row Grad", { id: "c1" }),
          cell("Cell Solid Override", {
            id: "c2",
            style: { background: "#FEF3C7" },
          }),
        ),
        row({ style: { height: 40 } }).append(
          cell("Cell Image", {
            id: "c3",
            style: { background: { image: samplePng, opacity: 0.5 } },
          }),
          cell("Cleared", { id: "c4", style: { background: "none" } }),
        ),
      ),
      text("Solid overrides layout gradient", {
        id: "hero-ph",
        placeholder: "hero",
        style: { background: "#DCFCE7" },
      }),
      placeholder({
        id: "pic-ph",
        placeholder: "slotPic",
        style: { background: "none" },
      }),
    );

    const vanSlide2 = slide({
      id: "slide-cleared-bg",
      layout: layoutDef,
      style: { background: "none" },
    }).append(
      shape({
        id: "ext-bg-shape",
        style: {
          position: "absolute",
          left: 220,
          top: 40,
          width: 160,
          height: 80,
          background: {
            image: { externalUrl: "https://example.com/bg.png" },
            fit: "cover",
            opacity: 0.75,
          },
        },
      }),
    );

    const vanDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: jostFont } },
    }).append(vanMaster, vanSlide1, vanSlide2);

    const jsxBuild = await build(jsxDoc);
    const vanBuild = await build(vanDoc);

    for (const built of [jsxBuild, vanBuild]) {
      const validated = await validatePptx(built.pptx);
      assert.deepEqual(validated.errors, []);

      const files = unzipSync(built.pptx);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);
      const slide2Xml = strFromU8(files["ppt/slides/slide2.xml"]!);
      const slide2Rels = strFromU8(files["ppt/slides/_rels/slide2.xml.rels"]!);

      assert.match(
        slide1Xml,
        /<a:gradFill/u,
        "Shape/Text/Row gradient emits <a:gradFill>",
      );
      assert.match(
        slide1Xml,
        /<a:alphaModFix amt="60000"\/>/u,
        "Box background image opacity emits alphaModFix",
      );
      assert.match(
        slide1Xml,
        /<a:noFill\/>/u,
        "Cell or placeholder background:'none' emits <a:noFill/>",
      );
      assert.match(
        slide2Xml,
        /<p:bg><p:bgPr><a:noFill\/><a:effectLst\/><\/p:bgPr><\/p:bg>/u,
        "Slide background:'none' emits explicit noFill background",
      );
      assert.match(
        slide2Rels,
        /Target="https:\/\/example\.com\/bg\.png" TargetMode="External"/u,
        "External image fill registers external relationship",
      );
    }

    const jsxPreview = await preview(jsxDoc);
    const vanPreview = await preview(vanDoc);
    for (const previewResult of [jsxPreview, vanPreview]) {
      assert.equal(previewResult.pages.length, 2);
      assert.deepEqual(previewResult.pages[0]?.diagnostics, []);
      assert.ok(
        previewResult.pages[0]?.svg?.includes("<linearGradient"),
        "Static preview paints Shape linearGradient",
      );
      assert.ok(
        previewResult.pages[0]?.svg?.includes("<radialGradient"),
        "Static preview paints Text radialGradient",
      );
      assert.ok(
        previewResult.pages[0]?.png !== undefined,
        "Page 1 PNG renders",
      );
      assert.ok(
        !previewResult.pages[1]?.svg?.includes("#F8FAFC") &&
          !previewResult.pages[1]?.svg?.includes("#0F172A"),
        "Slide background:'none' suppresses inherited layout and master backgrounds in preview",
      );
    }
  });

  it("supports style.border (including Shape arrows and Table/Row/Cell cascade + 4-side overrides), style.boxShadow:'none', and style.flipH/flipV", async () => {
    const layoutDef = defineSlideLayout({
      id: "visual-cascade-layout",
      placeholders: {
        clearedCard: {
          type: "body",
          style: {
            position: "absolute",
            left: 40,
            top: 320,
            width: 240,
            height: 80,
            border: { color: "#1E293B", width: 2, style: "dashed" },
            boxShadow: { color: "#000000", blur: 6, distance: 4, angle: 45 },
            flipH: true,
          },
        },
        overrideCard: {
          type: "body",
          style: {
            position: "absolute",
            left: 300,
            top: 320,
            width: 240,
            height: 80,
            border: { color: "#64748B", width: 1, style: "dotted" },
          },
        },
      },
    });

    const jsxDoc = (
      <Deck width={960} height={540} theme={{ fonts: { minor: jostFont } }}>
        <SlideMaster id="master">
          <SlideLayout layout={layoutDef} />
        </SlideMaster>
        <Slide id="visual-slide" layout={layoutDef}>
          <Shape
            id="arrow-line"
            geometry="line"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 160,
              height: 40,
              border: {
                color: "#0F172A",
                width: 2,
                style: "dashed",
                startArrow: "oval",
                endArrow: "triangle",
              },
              flipH: true,
              flipV: true,
            }}
          />
          <Box
            id="bordered-box"
            style={{
              position: "absolute",
              left: 220,
              top: 40,
              width: 120,
              height: 60,
              border: { color: "#6366F1", width: 2, style: "solid" },
              boxShadow: "none",
            }}
          />
          <Text
            id="flipped-text"
            style={{
              position: "absolute",
              left: 360,
              top: 40,
              width: 140,
              height: 60,
              border: { color: "#2563EB", width: 1, style: "dotted" },
              flipH: true,
              flipV: false,
            }}
          >
            Flipped
          </Text>
          <Image
            id="flipped-image"
            src={samplePng}
            style={{
              position: "absolute",
              left: 520,
              top: 40,
              width: 80,
              height: 60,
              border: { color: "#10B981", width: 2 },
              flipH: false,
              flipV: true,
            }}
          />
          <Group
            id="flipped-group"
            style={{
              position: "absolute",
              left: 620,
              top: 40,
              width: 120,
              height: 60,
              flipH: true,
              flipV: true,
            }}
          >
            <Shape
              id="group-child"
              geometry="rect"
              style={{
                position: "absolute",
                left: 10,
                top: 10,
                width: 80,
                height: 40,
                background: "#38BDF8",
              }}
            />
          </Group>
          <Table
            id="border-table"
            columns={[130, 130]}
            style={{
              position: "absolute",
              left: 40,
              top: 130,
              width: 260,
              border: { color: "#334155", width: 1, style: "solid" },
            }}
          >
            <Row
              style={{
                height: 44,
                border: { color: "#0EA5E9", width: 2, style: "dashed" },
              }}
            >
              <Cell id="tc1">Row Border</Cell>
              <Cell
                id="tc2"
                style={{
                  border: "none",
                  borderBottom: { color: "#DC2626", width: 3, style: "solid" },
                }}
              >
                Bottom Only
              </Cell>
            </Row>
            <Row style={{ height: 44 }}>
              <Cell
                id="tc3"
                style={{
                  borderLeft: "none",
                  borderTop: { color: "#16A34A", width: 2, style: "dotted" },
                }}
              >
                Mixed Sides
              </Cell>
              <Cell id="tc4" style={{ border: "none" }}>
                No Border
              </Cell>
            </Row>
          </Table>
          <Text
            id="ph-cleared"
            placeholder="clearedCard"
            style={{
              border: "none",
              boxShadow: "none",
              flipH: false,
              flipV: true,
            }}
          >
            Cleared Placeholder
          </Text>
          <Placeholder
            id="ph-overridden"
            placeholder="overrideCard"
            style={{
              border: { color: "#EF4444", width: 3, style: "solid" },
            }}
          />
        </Slide>
      </Deck>
    );

    const vanDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: jostFont } },
    }).append(
      slideMaster({ id: "master" }).append(slideLayout(layoutDef)),
      slide({ id: "visual-slide", layout: layoutDef }).append(
        shape({
          id: "arrow-line",
          geometry: "line",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 160,
            height: 40,
            border: {
              color: "#0F172A",
              width: 2,
              style: "dashed",
              startArrow: "oval",
              endArrow: "triangle",
            },
            flipH: true,
            flipV: true,
          },
        }),
        box({
          id: "bordered-box",
          style: {
            position: "absolute",
            left: 220,
            top: 40,
            width: 120,
            height: 60,
            border: { color: "#6366F1", width: 2, style: "solid" },
            boxShadow: "none",
          },
        }),
        text("Flipped", {
          id: "flipped-text",
          style: {
            position: "absolute",
            left: 360,
            top: 40,
            width: 140,
            height: 60,
            border: { color: "#2563EB", width: 1, style: "dotted" },
            flipH: true,
            flipV: false,
          },
        }),
        image({
          id: "flipped-image",
          src: samplePng,
          style: {
            position: "absolute",
            left: 520,
            top: 40,
            width: 80,
            height: 60,
            border: { color: "#10B981", width: 2 },
            flipH: false,
            flipV: true,
          },
        }),
        group({
          id: "flipped-group",
          style: {
            position: "absolute",
            left: 620,
            top: 40,
            width: 120,
            height: 60,
            flipH: true,
            flipV: true,
          },
        }).append(
          shape({
            id: "group-child",
            geometry: "rect",
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 80,
              height: 40,
              background: "#38BDF8",
            },
          }),
        ),
        table({
          id: "border-table",
          columns: [130, 130],
          style: {
            position: "absolute",
            left: 40,
            top: 130,
            width: 260,
            border: { color: "#334155", width: 1, style: "solid" },
          },
        }).append(
          row({
            style: {
              height: 44,
              border: { color: "#0EA5E9", width: 2, style: "dashed" },
            },
          }).append(
            cell("Row Border", { id: "tc1" }),
            cell("Bottom Only", {
              id: "tc2",
              style: {
                border: "none",
                borderBottom: { color: "#DC2626", width: 3, style: "solid" },
              },
            }),
          ),
          row({ style: { height: 44 } }).append(
            cell("Mixed Sides", {
              id: "tc3",
              style: {
                borderLeft: "none",
                borderTop: { color: "#16A34A", width: 2, style: "dotted" },
              },
            }),
            cell("No Border", { id: "tc4", style: { border: "none" } }),
          ),
        ),
        text("Cleared Placeholder", {
          id: "ph-cleared",
          placeholder: "clearedCard",
          style: {
            border: "none",
            boxShadow: "none",
            flipH: false,
            flipV: true,
          },
        }),
        placeholder({
          id: "ph-overridden",
          placeholder: "overrideCard",
          style: {
            border: { color: "#EF4444", width: 3, style: "solid" },
          },
        }),
      ),
    );

    for (const doc of [jsxDoc, vanDoc]) {
      const built = await build(doc);
      const validated = await validatePptx(built.pptx);
      assert.deepEqual(validated.errors, []);

      const files = unzipSync(built.pptx);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);

      assert.match(slide1Xml, /<a:headEnd type="oval"\/>/u);
      assert.match(slide1Xml, /<a:tailEnd type="triangle"\/>/u);
      assert.match(slide1Xml, /<a:ln><a:noFill\/><\/a:ln>/u);
      assert.match(slide1Xml, /<a:effectLst\/>/u);
      assert.match(slide1Xml, /<a:xfrm[^>]*flipH="1"[^>]*flipV="1"/u);
      assert.match(slide1Xml, /<a:lnL><a:noFill\/><\/a:lnL>/u);
      assert.match(
        slide1Xml,
        /<a:lnB w="28575"[^>]*><a:solidFill><a:srgbClr val="DC2626"\/>/u,
      );

      const previewResult = await preview(doc);
      assert.equal(previewResult.pages.length, 1);
      assert.deepEqual(previewResult.pages[0]?.diagnostics, []);
      const svg = previewResult.pages[0]?.svg ?? "";
      assert.ok(
        svg.includes("scale(-1 -1)"),
        "Group/Shape flipH+flipV reflected in SVG transform",
      );
      assert.ok(
        !svg.includes("feDropShadow"),
        "boxShadow:'none' clears inherited placeholder shadow in SVG",
      );
      assert.ok(
        svg.includes("#DC2626"),
        "Cell borderBottom color painted in SVG",
      );
    }

    await assert.rejects(
      () =>
        build(
          <Deck width={960} height={540}>
            <Slide id="invalid-box-border-slide">
              <Box
                style={{
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 100,
                  height: 50,
                  border: {
                    color: "#000000",
                    startArrow: "triangle",
                  } as never,
                }}
              />
            </Slide>
          </Deck>,
        ),
      (error) => error instanceof BuildError,
    );
  });

  it("supports Shape JSX children, shape(content?, props?) overloads, .content(...), and preserves text/padding styles on empty string ''", async () => {
    const jsxDoc = (
      <Deck width={960} height={540} theme={{ fonts: { minor: jostFont } }}>
        <Slide id="shape-content-slide">
          <Shape
            id="rich-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 240,
              height: 100,
              background: "#EFF6FF",
              border: { color: "#1D4ED8", width: 2 },
              fontSize: 20,
              color: "#0F172A",
              padding: [10, 14],
            }}
          >
            <P>
              Launch{" "}
              <Span style={{ color: "#10B981", fontWeight: "bold" }}>Now</Span>
            </P>
            <P>Ready</P>
          </Shape>
          <Shape
            id="empty-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 170,
              width: 220,
              height: 80,
              background: "#FEF2F2",
              fontSize: 32,
              color: "#FF0000",
              fontWeight: "bold",
              padding: 16,
            }}
          >
            {""}
          </Shape>
          <Shape
            id="styled-no-child-shape"
            geometry="rect"
            style={{
              position: "absolute",
              left: 300,
              top: 170,
              width: 180,
              height: 80,
              background: "#F8FAFC",
              fontSize: 18,
              padding: [8, 12],
            }}
          />
        </Slide>
      </Deck>
    );

    const vanDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: jostFont } },
    }).append(
      slide({ id: "shape-content-slide" }).append(
        shape(
          paragraph(
            "Launch ",
            span("Now", { style: { color: "#10B981", fontWeight: "bold" } }),
          ),
          paragraph("Ready"),
          {
            id: "rich-shape",
            geometry: "roundRect",
            style: {
              position: "absolute",
              left: 40,
              top: 40,
              width: 240,
              height: 100,
              background: "#EFF6FF",
              border: { color: "#1D4ED8", width: 2 },
              fontSize: 20,
              color: "#0F172A",
              padding: [10, 14],
            },
          },
        ),
        shape("", {
          id: "empty-shape",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 40,
            top: 170,
            width: 220,
            height: 80,
            background: "#FEF2F2",
            fontSize: 32,
            color: "#FF0000",
            fontWeight: "bold",
            padding: 16,
          },
        }),
        shape({
          id: "styled-no-child-shape",
          geometry: "rect",
          style: {
            position: "absolute",
            left: 300,
            top: 170,
            width: 180,
            height: 80,
            background: "#F8FAFC",
            fontSize: 18,
            padding: [8, 12],
          },
        }),
      ),
    );

    const contentOverloadDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: jostFont } },
    }).append(
      slide({ id: "overload-slide" }).append(
        shape("Label", {
          id: "str-overload",
          geometry: "roundRect",
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 180,
            height: 60,
            background: "#E2E8F0",
            fontSize: 16,
            padding: 8,
          },
        }),
        shape({
          id: "chain-overload",
          geometry: "rect",
          style: {
            position: "absolute",
            left: 240,
            top: 40,
            width: 180,
            height: 60,
            background: "#E2E8F0",
            fontSize: 16,
            padding: 8,
          },
        }).content("Chained"),
      ),
    );

    for (const doc of [jsxDoc, vanDoc]) {
      const built = await build(doc);
      const validated = await validatePptx(built.pptx);
      assert.deepEqual(validated.errors, []);

      const files = unzipSync(built.pptx);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);

      assert.match(
        slide1Xml,
        /lIns="152400" tIns="152400" rIns="152400" bIns="152400"/u,
        "Empty-string Shape retains padding=16 in <a:bodyPr>",
      );
      assert.match(
        slide1Xml,
        /<a:defRPr[^>]*sz="2400"[^>]*b="1"[^>]*><a:solidFill><a:srgbClr val="FF0000"\/>/u,
        "Empty-string Shape retains fontSize/color/fontWeight in <a:defRPr>",
      );
      assert.match(
        slide1Xml,
        /<a:endParaRPr[^>]*sz="2400"[^>]*b="1"/u,
        "Empty-string Shape retains run properties on <a:endParaRPr>",
      );

      const imported = importPptx(built.pptx);
      assert.equal(
        imported.report.diagnostics.some((d) => d.code === "unconsumed-xml"),
        false,
        "importPptx consumes <a:defRPr> and <a:endParaRPr> without unconsumed-xml",
      );

      const previewResult = await preview(doc);
      assert.equal(previewResult.pages.length, 1);
      assert.deepEqual(previewResult.pages[0]?.diagnostics, []);
    }

    const overloadBuilt = await build(contentOverloadDoc);
    const overloadValidated = await validatePptx(overloadBuilt.pptx);
    assert.deepEqual(overloadValidated.errors, []);
    assert.equal(overloadBuilt.slide("overload-slide").text, "Label\nChained");
  });

  it("enforces spin() mutually exclusive by vs from/to union in TypeScript, runtime build, importPptx, and codegen", async () => {
    // Compile-time mutual exclusion checks:
    // @ts-expect-error -- numeric by delta cannot be combined with from
    void spin({ from: 45, by: 90 });
    // @ts-expect-error -- numeric by delta cannot be combined with to
    void spin({ to: 225, by: 90 });
    // @ts-expect-error -- numeric by delta cannot be combined with from and to
    void spin({ from: 45, to: 225, by: 180 });

    assert.equal(spinOptionsSchema.safeParse({ by: 180 }).success, true);
    assert.equal(
      spinOptionsSchema.safeParse({ from: 45, to: 225 }).success,
      true,
    );
    assert.equal(
      spinOptionsSchema.safeParse({ from: 45, by: 90 }).success,
      false,
    );
    assert.equal(
      spinOptionsSchema.safeParse({ to: 225, by: 90 }).success,
      false,
    );

    const jsxDoc = (
      <Deck width={960} height={540} theme={{ fonts: { minor: jostFont } }}>
        <Slide id="spin-slide">
          <Shape
            id="spin-by-shape"
            geometry="rect"
            emphasis={spin({ by: 180, duration: 800 })}
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 120,
              height: 120,
              background: "#3B82F6",
            }}
          />
          <Shape
            id="spin-endpoints-shape"
            geometry="roundRect"
            emphasis={spin({ from: 45, to: 225, duration: 1000 })}
            style={{
              position: "absolute",
              left: 220,
              top: 40,
              width: 120,
              height: 120,
              background: "#10B981",
            }}
          />
          <Shape
            id="spin-to-only-shape"
            geometry="ellipse"
            emphasis={spin({ to: 180, duration: 600 })}
            style={{
              position: "absolute",
              left: 400,
              top: 40,
              width: 120,
              height: 120,
              background: "#F59E0B",
            }}
          />
        </Slide>
      </Deck>
    );

    const vanDoc = deck({
      width: 960,
      height: 540,
      theme: { fonts: { minor: jostFont } },
    }).append(
      slide({ id: "spin-slide" }).append(
        shape({
          id: "spin-by-shape",
          geometry: "rect",
          emphasis: spin({ by: 180, duration: 800 }),
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 120,
            height: 120,
            background: "#3B82F6",
          },
        }),
        shape({
          id: "spin-endpoints-shape",
          geometry: "roundRect",
          emphasis: spin({ from: 45, to: 225, duration: 1000 }),
          style: {
            position: "absolute",
            left: 220,
            top: 40,
            width: 120,
            height: 120,
            background: "#10B981",
          },
        }),
        shape({
          id: "spin-to-only-shape",
          geometry: "ellipse",
          emphasis: spin({ to: 180, duration: 600 }),
          style: {
            position: "absolute",
            left: 400,
            top: 40,
            width: 120,
            height: 120,
            background: "#F59E0B",
          },
        }),
      ),
    );

    for (const doc of [jsxDoc, vanDoc]) {
      const built = await build(doc);
      const validated = await validatePptx(built.pptx);
      assert.deepEqual(validated.errors, []);

      const files = unzipSync(built.pptx);
      const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]!);

      assert.match(slide1Xml, /<p:animRot by="10800000">/u);
      assert.match(slide1Xml, /<p:animRot from="2700000" to="13500000">/u);
      assert.match(slide1Xml, /<p:animRot from="0" to="10800000">/u);

      const imported = importPptx(built.pptx);
      assert.equal(
        imported.report.diagnostics.some((d) => d.code === "unconsumed-xml"),
        false,
      );
      const generated = codegen(imported.document, { target: "vanilla" });
      const rawEntry = generated.files[generated.entry];
      const entryFile = typeof rawEntry === "string" ? rawEntry : "";
      assert.match(entryFile, /spin\(\{\s*by:\s*180/u);
      assert.match(entryFile, /from:\s*45,\s*to:\s*225/u);
      assert.match(entryFile, /to:\s*180/u);
    }

    for (const invalidSpin of [
      spin({ from: 45, by: 90 } as never),
      spin({ to: 225, by: 90 } as never),
      spin({ from: 45, to: 225, by: 90 } as never),
    ]) {
      await assert.rejects(
        () =>
          build(
            <Deck width={960} height={540}>
              <Slide id="invalid-spin-slide">
                <Shape
                  geometry="rect"
                  emphasis={invalidSpin}
                  style={{
                    position: "absolute",
                    left: 20,
                    top: 20,
                    width: 100,
                    height: 100,
                  }}
                />
              </Slide>
            </Deck>,
          ),
        (error) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === "invalid-prop"),
      );
    }
  });
});
