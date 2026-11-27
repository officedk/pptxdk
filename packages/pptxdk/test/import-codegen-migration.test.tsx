import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  CodegenError,
  CommonContainer,
  CommonNode,
  DeckContainer,
  RichTextContainer,
  build,
  codegen,
  defineSlideLayout,
  font,
  importPptx,
  linearGradient,
} from "@officedk/pptxdk";
import { preview } from "@officedk/pptxdk/preview";
import {
  Box,
  Cell,
  Chart,
  Deck,
  Group,
  Image,
  P,
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
  deck as vanillaDeck,
  slide as vanillaSlide,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { describe, it } from "vitest";

import {
  previewGeneratedSourceRelocated,
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

const samplePng = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

const forbiddenNodeProps = new Set([
  "x",
  "y",
  "width",
  "height",
  "bounds",
  "fill",
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "startArrow",
  "endArrow",
  "flipH",
  "flipV",
  "text",
  "textStyle",
  "sz",
]);

const forbiddenStyleProps = new Set([
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "startArrow",
  "endArrow",
]);

const forbiddenPlaceholderSpecKeys = new Set([
  "x",
  "y",
  "width",
  "height",
  "flipH",
  "flipV",
  "sz",
  "backgroundColor",
  "backgroundImage",
  "borderColor",
  "borderWidth",
  "borderStyle",
]);

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function assertCleanStyleObject(style: unknown, path: string): void {
  if (style === undefined) return;
  assert.ok(isPlainRecord(style), `${path}.style must be an object`);
  for (const key of Object.keys(style)) {
    assert.equal(
      forbiddenStyleProps.has(key),
      false,
      `${path}.style must not contain legacy key "${key}"`,
    );
  }
}

function assertCleanPlaceholderSpec(spec: unknown, path: string): void {
  if (!isPlainRecord(spec)) return;
  for (const key of Object.keys(spec)) {
    assert.equal(
      forbiddenPlaceholderSpecKeys.has(key),
      false,
      `${path} must not contain legacy key "${key}"`,
    );
  }
  assertCleanStyleObject(spec["style"], path);
}

function collectAllNodes(
  nodes: readonly unknown[],
  out: Array<CommonNode<string, Record<string, unknown>>> = [],
): Array<CommonNode<string, Record<string, unknown>>> {
  for (const item of nodes) {
    if (item instanceof CommonNode) {
      out.push(item as CommonNode<string, Record<string, unknown>>);
      if (item instanceof CommonContainer) {
        collectAllNodes(item.children, out);
      }
    }
  }
  return out;
}

describe("Ticket 06 Part 1: importPptx single-track style, orthogonal visuals, Shape children, and PlaceholderSize", () => {
  it("imports representative deck without legacy props, preserves Group containing-block coordinates, and roundtrips cleanly through build, validatePptx, and preview", async () => {
    const layoutDef = defineSlideLayout({
      id: "hero-layout",
      name: "Hero Layout",
      placeholders: {
        hero: {
          type: "body",
          idx: 1,
          size: "half",
          style: {
            position: "absolute",
            left: 40,
            top: 400,
            width: 300,
            height: 70,
            background: "#E0F2FE",
            border: { color: "#0284C7", width: 1, style: "solid" },
            fontFamily: jostFont,
            fontSize: 14,
            color: "#0F172A",
          },
        },
      },
      footerPlaceholders: {
        footer: {
          text: "Master Footer",
          size: "quarter",
          style: {
            position: "absolute",
            left: 40,
            top: 496,
            width: 220,
            height: 24,
            fontFamily: jostFont,
            fontSize: 10,
            color: "#475569",
          },
        },
      },
    });

    const initial = await build(
      <Deck
        width={960}
        height={540}
        theme={{ fonts: { major: jostFont, minor: jostFont } }}
      >
        <SlideMaster
          id="main-master"
          headerFooter={{ footer: true }}
          placeholders={{
            title: {
              type: "title",
              size: "full",
              style: {
                position: "absolute",
                left: 40,
                top: 16,
                width: 880,
                height: 36,
              },
            },
          }}
        >
          <SlideLayout layout={layoutDef} />
        </SlideMaster>
        <Slide
          id="showcase"
          layout={layoutDef}
          style={{ background: "#F8FAFC" }}
          footer="Slide Footer"
        >
          <Shape
            id="hero-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 220,
              height: 80,
              background: linearGradient(90, "#0EA5E9", "#2563EB"),
              border: {
                color: "#0F172A",
                width: 2,
                style: "dashed",
                startArrow: "oval",
                endArrow: "triangle",
              },
              flipH: true,
              boxShadow: { color: "#000000", blur: 4, distance: 3, angle: 45 },
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
            }}
          >
            Shape Content
          </Shape>
          <Group
            id="pos-group"
            style={{
              position: "absolute",
              left: 300,
              top: 40,
              width: 260,
              height: 140,
              flipV: true,
            }}
          >
            <Shape
              id="pos-group-child"
              geometry="rect"
              style={{
                position: "absolute",
                left: 20,
                top: 30,
                width: 110,
                height: 50,
                background: "#38BDF8",
              }}
            />
            <Group
              id="nested-pos-group"
              style={{
                position: "absolute",
                left: 150,
                top: 20,
                width: 90,
                height: 80,
              }}
            >
              <Shape
                id="nested-pos-child"
                geometry="rect"
                style={{
                  position: "absolute",
                  left: 10,
                  top: 15,
                  width: 60,
                  height: 30,
                  background: "#0284C7",
                }}
              />
            </Group>
          </Group>
          <Group id="unpos-group" style={{ rotate: 15 }}>
            <Shape
              id="unpos-child-a"
              geometry="rect"
              style={{
                position: "absolute",
                left: 300,
                top: 200,
                width: 80,
                height: 40,
                background: "#A855F7",
              }}
            />
            <Shape
              id="unpos-child-b"
              geometry="rect"
              style={{
                position: "absolute",
                left: 400,
                top: 200,
                width: 80,
                height: 40,
                background: "#EC4899",
              }}
            />
          </Group>
          <Table
            id="summary-table"
            columns={[120, 120]}
            style={{
              position: "absolute",
              left: 40,
              top: 180,
              width: 240,
              fontFamily: jostFont,
              fontSize: 12,
            }}
          >
            <Row style={{ height: 40, background: "#E2E8F0" }}>
              <Cell
                id="cell-a"
                style={{
                  background: "#FFFFFF",
                  border: { color: "#334155", width: 1, style: "solid" },
                }}
              >
                Cell A
              </Cell>
              <Cell
                id="cell-b"
                style={{
                  borderBottom: { color: "#DC2626", width: 2, style: "solid" },
                }}
              >
                Cell B
              </Cell>
            </Row>
          </Table>
          <Chart
            id="sales-chart"
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "Sales", values: [10, 20] }],
              },
            ]}
            style={{
              position: "absolute",
              left: 600,
              top: 40,
              width: 300,
              height: 180,
              fontFamily: jostFont,
              fontSize: 11,
            }}
          />
          <Image
            id="brand-image"
            src={samplePng}
            style={{
              position: "absolute",
              left: 600,
              top: 240,
              width: 100,
              height: 80,
              border: { color: "#10B981", width: 2, style: "solid" },
              flipH: true,
            }}
          />
          <Text
            id="free-text"
            style={{
              position: "absolute",
              left: 40,
              top: 280,
              width: 220,
              height: 50,
              background: "#FEF3C7",
              border: { color: "#D97706", width: 1, style: "solid" },
              fontFamily: jostFont,
              fontSize: 14,
              color: "#0F172A",
            }}
          >
            Free Text
          </Text>
          <Text id="bound-hero" placeholder="hero">
            Placeholder Text
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );

    const imported = importPptx(initial.pptx);
    assert.ok(imported.document instanceof DeckContainer);
    const allNodes = collectAllNodes([
      ...imported.document.masters,
      ...imported.document.masters.flatMap((master) => master.layouts),
      ...imported.document.children,
    ]);
    assert.ok(allNodes.length > 0);

    for (const node of allNodes) {
      const props = node.props;
      for (const key of Object.keys(props)) {
        assert.equal(
          forbiddenNodeProps.has(key),
          false,
          `${node.kind}("${node.id ?? node.type}") must not emit legacy prop "${key}"`,
        );
      }
      assertCleanStyleObject(
        props["style"],
        `${node.kind}("${node.id ?? ""}")`,
      );

      if (node.kind === "slide-master" || node.kind === "slide-layout") {
        const placeholders = props["placeholders"];
        if (isPlainRecord(placeholders)) {
          for (const [phKey, phSpec] of Object.entries(placeholders)) {
            assertCleanPlaceholderSpec(
              phSpec,
              `${node.kind}.placeholders.${phKey}`,
            );
          }
        }
        const footerPlaceholders = props["footerPlaceholders"];
        if (isPlainRecord(footerPlaceholders)) {
          for (const [fKey, fSpec] of Object.entries(footerPlaceholders)) {
            assertCleanPlaceholderSpec(
              fSpec,
              `${node.kind}.footerPlaceholders.${fKey}`,
            );
          }
        }
      }
    }

    // Verify Shape uses child content + single-track style
    const shapeNode = allNodes.find(
      (node) =>
        node.kind === "shape" &&
        node instanceof RichTextContainer &&
        node.children.length > 0,
    );
    assert.ok(
      shapeNode !== undefined,
      "Expected imported Shape with child content",
    );
    const shapeStyle = shapeNode.props["style"] as Record<string, unknown>;
    assert.equal(shapeStyle["position"], "absolute");
    assert.equal(shapeStyle["left"], 40);
    assert.equal(shapeStyle["top"], 40);
    assert.equal(shapeStyle["width"], 220);
    assert.equal(shapeStyle["height"], 80);
    assert.equal(shapeStyle["flipH"], true);
    assert.ok(isPlainRecord(shapeStyle["background"]));
    assert.ok(isPlainRecord(shapeStyle["border"]));
    assert.equal(shapeStyle["border"]["startArrow"], "oval");
    assert.equal(shapeStyle["border"]["endArrow"], "triangle");
    assert.ok(isPlainRecord(shapeStyle["boxShadow"]));

    // Verify positioned Group vs unpositioned Group containing-block coordinates
    const groupNodes = allNodes.filter(
      (node): node is CommonContainer<string, Record<string, unknown>> =>
        node.kind === "group" && node instanceof CommonContainer,
    );
    const posGroup = groupNodes.find((groupNode) => {
      const style = groupNode.props["style"] as
        | Record<string, unknown>
        | undefined;
      return style?.["position"] === "absolute" && style?.["left"] === 300;
    });
    assert.ok(
      posGroup !== undefined,
      "Expected positioned outer Group at left=300",
    );
    const posGroupStyle = posGroup.props["style"] as Record<string, unknown>;
    assert.equal(posGroupStyle["top"], 40);
    assert.equal(posGroupStyle["width"], 260);
    assert.equal(posGroupStyle["height"], 140);
    assert.equal(posGroupStyle["flipV"], true);

    const posDirectShape = posGroup.children.find(
      (child): child is CommonNode<string, Record<string, unknown>> =>
        child instanceof CommonNode && child.kind === "shape",
    );
    assert.ok(posDirectShape !== undefined);
    const posDirectStyle = posDirectShape.props["style"] as Record<
      string,
      unknown
    >;
    assert.equal(posDirectStyle["position"], "absolute");
    assert.equal(posDirectStyle["left"], 20);
    assert.equal(posDirectStyle["top"], 30);

    const nestedGroup = posGroup.children.find(
      (child): child is CommonContainer<string, Record<string, unknown>> =>
        child instanceof CommonContainer && child.kind === "group",
    );
    assert.ok(nestedGroup !== undefined, "Expected nested positioned Group");
    const nestedGroupStyle = nestedGroup.props["style"] as Record<
      string,
      unknown
    >;
    assert.equal(nestedGroupStyle["position"], "absolute");
    assert.equal(nestedGroupStyle["left"], 150);
    assert.equal(nestedGroupStyle["top"], 20);
    const nestedChildShape = nestedGroup.children.find(
      (child): child is CommonNode<string, Record<string, unknown>> =>
        child instanceof CommonNode && child.kind === "shape",
    );
    assert.ok(nestedChildShape !== undefined);
    const nestedChildStyle = nestedChildShape.props["style"] as Record<
      string,
      unknown
    >;
    assert.equal(nestedChildStyle["left"], 10);
    assert.equal(nestedChildStyle["top"], 15);

    const unposGroup = groupNodes.find((groupNode) => {
      const style = groupNode.props["style"] as
        | Record<string, unknown>
        | undefined;
      return style?.["position"] === undefined && style?.["rotate"] === 15;
    });
    assert.ok(
      unposGroup !== undefined,
      "Expected unpositioned Group with rotate=15",
    );
    const unposChildA = unposGroup.children.find(
      (child): child is CommonNode<string, Record<string, unknown>> =>
        child instanceof CommonNode &&
        child.kind === "shape" &&
        isPlainRecord(child.props["style"]) &&
        child.props["style"]["left"] === 300,
    );
    assert.ok(
      unposChildA !== undefined,
      "Expected unpositioned Group child to retain slide-space left=300",
    );

    // Verify Table / Row / Cell, Chart, Image, SlideLayout placeholder size & footerPlaceholders
    const tableNode = allNodes.find((node) => node.kind === "table");
    assert.ok(tableNode !== undefined);
    const tableStyle = tableNode.props["style"] as Record<string, unknown>;
    assert.equal(tableStyle["position"], "absolute");
    assert.equal(tableStyle["left"], 40);
    assert.equal(tableStyle["top"], 180);

    const cellNodes = allNodes.filter((node) => node.kind === "cell");
    assert.ok(cellNodes.length >= 2);
    const cellAStyle = cellNodes[0]!.props["style"] as Record<string, unknown>;
    assert.equal(cellAStyle["background"], "#FFFFFF");
    assert.ok(isPlainRecord(cellAStyle["border"]));
    const cellBStyle = cellNodes[1]!.props["style"] as Record<string, unknown>;
    assert.ok(isPlainRecord(cellBStyle["borderBottom"]));

    const chartNode = allNodes.find((node) => node.kind === "chart");
    assert.ok(chartNode !== undefined);
    const chartStyle = chartNode.props["style"] as Record<string, unknown>;
    assert.equal(chartStyle["position"], "absolute");
    assert.equal(chartStyle["left"], 600);
    assert.equal(chartStyle["top"], 40);
    assert.equal(chartStyle["width"], 300);
    assert.equal(chartStyle["height"], 180);

    const imageNode = allNodes.find((node) => node.kind === "image");
    assert.ok(imageNode !== undefined);
    const imageStyle = imageNode.props["style"] as Record<string, unknown>;
    assert.equal(imageStyle["position"], "absolute");
    assert.equal(imageStyle["flipH"], true);
    assert.ok(isPlainRecord(imageStyle["border"]));

    const layoutNode = allNodes.find((node) => node.kind === "slide-layout");
    assert.ok(layoutNode !== undefined);
    const layoutPlaceholders = layoutNode.props["placeholders"] as Record<
      string,
      Record<string, unknown>
    >;
    const heroPh = Object.values(layoutPlaceholders).find(
      (ph) => ph["type"] === "body",
    );
    assert.ok(heroPh !== undefined);
    assert.equal(heroPh["size"], "half");
    assert.ok(isPlainRecord(heroPh["style"]));
    assert.equal(heroPh["style"]["position"], "absolute");
    assert.equal(heroPh["style"]["background"], "#E0F2FE");
    assert.ok(isPlainRecord(heroPh["style"]["border"]));

    const ftrPh = layoutPlaceholders["ftr"];
    assert.ok(ftrPh !== undefined);
    assert.equal(ftrPh["type"], "ftr");
    assert.equal(ftrPh["size"], "quarter");
    assert.ok(isPlainRecord(ftrPh["style"]));
    assert.equal(ftrPh["style"]["position"], "absolute");
    assert.equal(ftrPh["style"]["left"], 40);
    assert.equal(ftrPh["style"]["top"], 496);

    // Re-build, validatePptx, and preview the imported editable document
    const rebuilt = await build(imported.document);
    const validated = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validated.errors, []);

    const painted = await preview(imported.document);
    assert.equal(painted.pages.length, 1);
    assert.deepEqual(painted.pages[0]?.errors, []);

    // Verify nested Group coordinates roundtrip identically to slide space
    const rebuiltSlide = rebuilt.slides[0]!;
    const posGroupLoc = rebuiltSlide.objects.find(
      (obj) => obj.box?.x === 300 && obj.box?.y === 40 && obj.kind === "group",
    );
    assert.ok(posGroupLoc?.box !== undefined);
    assert.equal(posGroupLoc.box.width, 260);
    assert.equal(posGroupLoc.box.height, 140);
  });
});

describe("Ticket 06 Part 2: codegen vanilla + react new author syntax, strict consumer typecheck, buildFromSource, and previewFromSource", () => {
  it("generates clean vanilla and react sources for imported representative deck and verifies strict tsc, buildFromSource, validatePptx, and previewFromSource", async () => {
    const layoutDef = defineSlideLayout({
      id: "codegen-layout",
      name: "Codegen Layout",
      placeholders: {
        hero: {
          type: "body",
          idx: 1,
          size: "half",
          style: {
            position: "absolute",
            left: 40,
            top: 420,
            width: 300,
            height: 60,
            background: "#E0F2FE",
            border: { color: "#0284C7", width: 1, style: "solid" },
            fontFamily: jostFont,
            fontSize: 14,
            color: "#0F172A",
          },
        },
      },
      footerPlaceholders: {
        footer: {
          text: "Master Footer",
          size: "quarter",
          style: {
            position: "absolute",
            left: 40,
            top: 496,
            width: 220,
            height: 24,
            fontFamily: jostFont,
            fontSize: 10,
            color: "#475569",
          },
        },
      },
    });

    const initial = await build(
      <Deck
        width={960}
        height={540}
        theme={{ fonts: { major: jostFont, minor: jostFont } }}
      >
        <SlideMaster
          id="codegen-master"
          headerFooter={{ footer: true }}
          placeholders={{
            title: {
              type: "title",
              size: "full",
              style: {
                position: "absolute",
                left: 40,
                top: 16,
                width: 880,
                height: 36,
              },
            },
          }}
        >
          <SlideLayout layout={layoutDef} />
        </SlideMaster>
        <Slide
          id="codegen-slide"
          layout={layoutDef}
          background={{ image: samplePng, fit: "cover", opacity: 0.9 }}
          footer="Slide Footer"
        >
          <Shape
            id="plain-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 40,
              width: 220,
              height: 70,
              background: linearGradient(90, "#0EA5E9", "#2563EB"),
              border: {
                color: "#0F172A",
                width: 2,
                style: "dashed",
                startArrow: "oval",
                endArrow: "triangle",
              },
              flipH: true,
              boxShadow: { color: "#000000", blur: 4, distance: 3, angle: 45 },
              fontFamily: jostFont,
              fontSize: 16,
              color: "#FFFFFF",
            }}
          >
            Shape Content
          </Shape>
          <Shape
            id="empty-styled-shape"
            geometry="rect"
            style={{
              position: "absolute",
              left: 40,
              top: 120,
              width: 110,
              height: 44,
              background: "#E2E8F0",
              fontFamily: jostFont,
              fontSize: 13,
              color: "#1E293B",
            }}
          >
            {""}
          </Shape>
          <Shape
            id="rich-shape"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 160,
              top: 120,
              width: 120,
              height: 44,
              background: { image: samplePng, fit: "cover" },
            }}
          >
            <P style={{ align: "center" }}>
              <Span
                style={{
                  fontFamily: jostFont,
                  textDecoration: "underline",
                  color: "#0F172A",
                }}
              >
                Rich
              </Span>{" "}
              <Span style={{ fontFamily: jostFont, color: "#2563EB" }}>
                Shape
              </Span>
            </P>
          </Shape>
          <Group
            id="pos-group"
            style={{
              position: "absolute",
              left: 300,
              top: 40,
              width: 260,
              height: 130,
              flipV: true,
            }}
          >
            <Shape
              id="pos-group-child"
              geometry="rect"
              style={{
                position: "absolute",
                left: 20,
                top: 30,
                width: 110,
                height: 50,
                background: "#38BDF8",
              }}
            />
          </Group>
          <Group id="unpos-group" style={{ rotate: 15 }}>
            <Shape
              id="unpos-child-a"
              geometry="rect"
              style={{
                position: "absolute",
                left: 300,
                top: 200,
                width: 80,
                height: 40,
                background: "#A855F7",
              }}
            />
            <Shape
              id="unpos-child-b"
              geometry="rect"
              style={{
                position: "absolute",
                left: 400,
                top: 200,
                width: 80,
                height: 40,
                background: "#EC4899",
              }}
            />
          </Group>
          <Table
            id="summary-table"
            columns={[120, 120]}
            style={{
              position: "absolute",
              left: 40,
              top: 180,
              width: 240,
              fontFamily: jostFont,
              fontSize: 12,
            }}
          >
            <Row style={{ height: 40, background: "#E2E8F0" }}>
              <Cell
                id="cell-a"
                style={{
                  background: "#FFFFFF",
                  border: { color: "#334155", width: 1, style: "solid" },
                }}
              >
                Cell A
              </Cell>
              <Cell
                id="cell-b"
                style={{
                  borderBottom: { color: "#DC2626", width: 2, style: "solid" },
                }}
              >
                Cell B
              </Cell>
            </Row>
          </Table>
          <Chart
            id="sales-chart"
            groups={[
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["Q1", "Q2"] },
                series: [{ name: "Sales", values: [10, 20] }],
              },
            ]}
            style={{
              position: "absolute",
              left: 600,
              top: 40,
              width: 300,
              height: 180,
              fontFamily: jostFont,
              fontSize: 11,
            }}
          />
          <Image
            id="brand-image"
            src={samplePng}
            style={{
              position: "absolute",
              left: 600,
              top: 240,
              width: 100,
              height: 80,
              border: { color: "#10B981", width: 2, style: "solid" },
              flipH: true,
            }}
          />
          <Box
            id="author-box"
            style={{
              position: "absolute",
              left: 600,
              top: 340,
              width: 180,
              height: 60,
              background: "#FEF3C7",
              border: { color: "#D97706", width: 1, style: "solid" },
            }}
          >
            <Text
              id="free-text"
              style={{
                fontFamily: jostFont,
                fontSize: 14,
                color: "#0F172A",
              }}
            >
              Free Text
            </Text>
          </Box>
          <Text id="bound-hero" placeholder="hero">
            Placeholder Text
          </Text>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );

    const imported = importPptx(initial.pptx);
    const vanillaOut = codegen(imported.document, { target: "vanilla" });
    const reactOut = codegen(imported.document, { target: "react" });

    const vanillaSource = vanillaOut.files["deck.ts"];
    const reactSource = reactOut.files["deck.tsx"];
    assert.equal(typeof vanillaSource, "string");
    assert.equal(typeof reactSource, "string");

    for (const [label, source] of [
      ["vanilla", vanillaSource as string],
      ["react", reactSource as string],
    ] as const) {
      for (const forbidden of [
        /\bbackgroundColor\b/u,
        /\bbackgroundImage\b/u,
        /\bborderColor\b/u,
        /\bborderWidth\b/u,
        /\bborderStyle\b/u,
        /\btextStyle\b/u,
        /\bbounds:/u,
        /\bsz:/u,
      ]) {
        assert.doesNotMatch(
          source,
          forbidden,
          `${label} output must not contain legacy token ${String(forbidden)}`,
        );
      }
      assert.match(source, /position:\s*"absolute"/u);
      assert.match(source, /background:/u);
      assert.match(source, /border:/u);
      assert.match(source, /size:\s*"half"/u);
    }

    assert.match(vanillaSource as string, /shape\(\s*"Shape Content",/u);
    assert.match(vanillaSource as string, /shape\(\s*"",/u);
    assert.match(vanillaSource as string, /shape\(\s*\[/u);

    assert.match(
      reactSource as string,
      /<Shape[\s\S]*?>\s*Shape Content\s*<\/Shape>/u,
    );
    assert.match(reactSource as string, /\{\s*""\s*\}/u);
    assert.match(reactSource as string, /<P\b/u);

    await verifyGeneratedSourcesTypecheck([
      { label: "vanilla", result: vanillaOut },
      { label: "react", result: reactOut },
    ]);

    for (const out of [vanillaOut, reactOut] as const) {
      const rebuilt = await rebuildGeneratedSourceRelocated(out);
      assert.deepEqual(rebuilt.file.errors, []);

      const previewed = await previewGeneratedSourceRelocated(out);
      assert.equal(previewed.pages.length, 1);
      assert.deepEqual(previewed.pages[0]?.errors, []);
    }

    // Also verify external image URLs in Image.src and style.background image descriptors
    const extDoc = vanillaDeck({ width: 960, height: 540 }).append(
      vanillaSlide({ id: "ext-slide", background: "none" }).append(
        new CommonNode("image", "image", {
          id: "ext-img",
          src: { externalUrl: "https://example.com/logo.png" },
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 120,
            height: 60,
          },
        }) as never,
        new RichTextContainer("shape", "shape", {
          id: "ext-bg-shape",
          geometry: "rect",
          style: {
            position: "absolute",
            left: 160,
            top: 20,
            width: 140,
            height: 60,
            background: {
              image: { externalUrl: "https://example.com/bg.png" },
              fit: "contain",
              position: "center",
              opacity: 0.8,
            },
          },
        }) as never,
      ),
    );
    const extVanilla = codegen(extDoc, { target: "vanilla" });
    const extReact = codegen(extDoc, { target: "react" });
    assert.match(
      extVanilla.files["deck.ts"] as string,
      /externalUrl:\s*"https:\/\/example\.com\/bg\.png"/u,
    );
    assert.match(extVanilla.files["deck.ts"] as string, /fit:\s*"contain"/u);
    assert.match(
      extVanilla.files["deck.ts"] as string,
      /position:\s*"center"/u,
    );
    assert.match(extVanilla.files["deck.ts"] as string, /opacity:\s*0\.8/u);
    await verifyGeneratedSourcesTypecheck([
      { label: "ext-vanilla", result: extVanilla },
      { label: "ext-react", result: extReact },
    ]);
  });

  it("roundtrips a full-feature deck (master footer letterSpacing, multiple Chart kinds, merged Table, and Shape) with zero ImportDiagnostics and dual-target codegen rebuild", async () => {
    const masterFooterPlaceholders = {
      dateTime: {
        text: "2026-11-27",
        style: {
          position: "absolute" as const,
          left: 40,
          top: 496,
          width: 180,
          height: 24,
          fontFamily: jostFont,
          fontSize: 10,
          color: "#475569" as const,
          letterSpacing: 1.6,
        },
      },
      footer: {
        text: "Acme Full Feature",
        style: {
          position: "absolute" as const,
          left: 240,
          top: 496,
          width: 320,
          height: 24,
          fontFamily: jostFont,
          fontSize: 10,
          color: "#334155" as const,
          letterSpacing: 2,
          align: "center" as const,
        },
      },
      slideNumber: {
        style: {
          position: "absolute" as const,
          left: 780,
          top: 496,
          width: 140,
          height: 24,
          fontFamily: jostFont,
          fontSize: 10,
          color: "#0F172A" as const,
          letterSpacing: 0.8,
          align: "right" as const,
        },
      },
    };

    const fullFeatureLayout = defineSlideLayout({
      id: "full-feature-layout",
      name: "Full Feature Layout",
      type: "obj",
      preserve: true,
    });

    const initial = await build(
      <Deck
        width={960}
        height={540}
        theme={{ fonts: { major: jostFont, minor: jostFont } }}
      >
        <SlideMaster
          id="full-feature-master"
          preserve={true}
          footerPlaceholders={masterFooterPlaceholders}
        >
          <SlideLayout layout={fullFeatureLayout} />
        </SlideMaster>
        <Slide id="overview" layout={fullFeatureLayout}>
          <Shape
            id="kpi-banner"
            geometry="roundRect"
            style={{
              position: "absolute",
              left: 40,
              top: 28,
              width: 880,
              height: 52,
              background: linearGradient(90, "#0F172A", "#1D4ED8"),
              border: { color: "#1E40AF", width: 1, style: "solid" },
              fontFamily: jostFont,
              fontSize: 18,
              color: "#FFFFFF",
              letterSpacing: 1.2,
            }}
          >
            Quarterly Business Review
          </Shape>
          <Chart
            id="revenue-column-chart"
            groups={[
              {
                kind: "column",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [
                  { name: "Revenue", values: [120, 145, 178] },
                  { name: "Target", values: [110, 140, 170] },
                ],
              },
            ]}
            style={{
              position: "absolute",
              left: 40,
              top: 96,
              width: 280,
              height: 180,
              fontFamily: jostFont,
              fontSize: 10,
            }}
          />
          <Chart
            id="margin-line-chart"
            groups={[
              {
                kind: "line",
                data: {
                  kind: "shared-categories",
                  categories: ["Q1", "Q2", "Q3"],
                },
                series: [{ name: "Margin %", values: [22, 25, 28] }],
              },
            ]}
            style={{
              position: "absolute",
              left: 340,
              top: 96,
              width: 280,
              height: 180,
              fontFamily: jostFont,
              fontSize: 10,
            }}
          />
          <Chart
            id="mix-pie-chart"
            groups={[
              {
                kind: "pie",
                data: {
                  kind: "shared-categories",
                  categories: ["Enterprise", "SMB"],
                },
                series: [{ name: "Mix", values: [68, 32] }],
              },
            ]}
            style={{
              position: "absolute",
              left: 640,
              top: 96,
              width: 280,
              height: 180,
              fontFamily: jostFont,
              fontSize: 10,
            }}
          />
          <Table
            id="merged-metrics-table"
            columns={[180, 180, 180]}
            style={{
              position: "absolute",
              left: 40,
              top: 292,
              width: 540,
              fontFamily: jostFont,
              fontSize: 11,
            }}
          >
            <Row style={{ height: 34, background: "#1E293B" }}>
              <Cell
                colSpan={2}
                style={{
                  color: "#FFFFFF",
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                Merged Region Header
              </Cell>
              <Cell
                style={{
                  color: "#FFFFFF",
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                Status
              </Cell>
            </Row>
            <Row style={{ height: 34 }}>
              <Cell
                rowSpan={2}
                style={{
                  background: "#EFF6FF",
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                EMEA (2 rows)
              </Cell>
              <Cell
                style={{
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                Q2 Actual
              </Cell>
              <Cell
                style={{
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                On Track
              </Cell>
            </Row>
            <Row style={{ height: 34 }}>
              <Cell
                style={{
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                Q3 Forecast
              </Cell>
              <Cell
                style={{
                  border: { color: "#CBD5E1", width: 1, style: "solid" },
                }}
              >
                Ahead
              </Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
      { embedFonts: "full" },
    );

    const imported = importPptx(initial.pptx);
    assert.deepEqual(imported.report.diagnostics, []);

    const importedSlide = imported.document.query.get("overview").node;
    assert.equal(
      (
        importedSlide.props["dateTime"] as {
          style?: Record<string, unknown>;
        }
      )?.style?.["letterSpacing"],
      1.6,
    );
    assert.equal(
      (importedSlide.props["footer"] as { style?: Record<string, unknown> })
        ?.style?.["letterSpacing"],
      2,
    );
    assert.equal(
      (
        importedSlide.props["slideNumber"] as {
          style?: Record<string, unknown>;
        }
      )?.style?.["letterSpacing"],
      0.8,
    );

    const vanillaOut = codegen(imported.document, { target: "vanilla" });
    const reactOut = codegen(imported.document, { target: "react" });

    assert.match(
      vanillaOut.files["deck.ts"] as string,
      /letterSpacing:\s*2\b/u,
    );
    assert.match(reactOut.files["deck.tsx"] as string, /letterSpacing:\s*2\b/u);

    await verifyGeneratedSourcesTypecheck([
      { label: "full-feature-vanilla", result: vanillaOut },
      { label: "full-feature-react", result: reactOut },
    ]);

    for (const out of [vanillaOut, reactOut] as const) {
      const rebuilt = await rebuildGeneratedSourceRelocated(out);
      assert.deepEqual(rebuilt.file.errors, []);

      const reimported = importPptx(rebuilt.built.pptx);
      assert.deepEqual(reimported.report.diagnostics, []);
    }
  });

  it("rejects legacy top-level geometry and legacy visual style keys across canvas/media declarations", () => {
    const makeDocWith = (child: CommonNode): DeckContainer =>
      vanillaDeck({ width: 960, height: 540 }).append(
        vanillaSlide({ id: "s1" }).append(child as never),
      );

    const legacyCases: readonly CommonNode[] = [
      new RichTextContainer("shape", "shape", {
        x: 10,
        y: 20,
        width: 100,
        height: 50,
      }),
      new RichTextContainer("shape", "shape", { fill: "#FF0000" }),
      new RichTextContainer("shape", "shape", { text: "Legacy" }),
      new RichTextContainer("shape", "shape", {
        style: { backgroundColor: "#FF0000" },
      }),
      new CommonNode("image", "image", { src: samplePng, x: 10 }),
      new CommonNode("image", "image", {
        src: samplePng,
        nativePlaceholder: { type: "pic", sz: "half" },
      }),
      new CommonContainer("box", "box", {
        x: 10,
        y: 10,
        width: 100,
        height: 50,
      }),
      new CommonContainer("box", "box", { style: { borderColor: "#000000" } }),
      new CommonContainer("group", "group", {
        bounds: { x: 0, y: 0, width: 100, height: 100 },
      }),
      new RichTextContainer("text", "text", { x: 10, y: 10, width: 100 }),
      new RichTextContainer("text", "text", {
        style: { backgroundImage: { image: samplePng } },
      }),
    ];

    for (const badNode of legacyCases) {
      const doc = makeDocWith(badNode);
      assert.throws(
        () => codegen(doc, { target: "vanilla" }),
        (err: unknown) => err instanceof CodegenError,
        `Expected CodegenError for legacy ${badNode.type} node`,
      );
    }
  });
});
