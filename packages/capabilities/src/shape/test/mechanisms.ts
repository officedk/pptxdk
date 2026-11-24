import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { Box } from "@pptxdk/capabilities/box-elements";
import {
  box,
  deck,
  font,
  group,
  shape,
  slide,
  text,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import {
  linearGradient,
  preset,
  radialGradient,
  type CustomGeometry,
} from "@pptxdk/capabilities/shape";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import { CommonNode } from "@pptxdk/core";
import { Deck, Slide } from "@pptxdk/react";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";
import type { ExhaustiveMechanismVerificationMap } from "../../../test/mechanism-verification.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

const sampleFillPng = (() => {
  const img = new PNG({ width: 4, height: 4 });
  for (let i = 0; i < img.data.length; i += 4) {
    img.data.set([0x16, 0xa3, 0x4a, 255], i);
  }
  return new Uint8Array(PNG.sync.write(img));
})();

export const customBezierShape: CustomGeometry = {
  type: "custom",
  width: 100,
  height: 100,
  paths: [
    {
      commands: [
        { op: "M", x: 10, y: 80 },
        { op: "C", x1: 10, y1: 15, x2: 90, y2: 15, x: 90, y: 80 },
        { op: "Q", x1: 50, y1: 95, x: 10, y: 80 },
        { op: "Z" },
      ],
      fill: true,
      stroke: true,
    },
  ],
};

function findFirstSlideNode(root: CommonNode): CommonNode | undefined {
  if (root.type === "slide") return root;
  for (const child of root.children) {
    if (child instanceof CommonNode) {
      const found = findFirstSlideNode(child);
      if (found !== undefined) return found;
    }
  }
  return undefined;
}

function findShapeNodeByName(
  root: CommonNode,
  name: string,
): CommonNode | undefined {
  let matched: CommonNode | undefined;
  const visit = (node: CommonNode): void => {
    if (matched !== undefined) return;
    if (node.type === "shape" && node.props["name"] === name) {
      matched = node;
      return;
    }
    for (const child of node.children) {
      if (child instanceof CommonNode) {
        visit(child);
      }
    }
  };
  visit(root);
  return matched;
}

function extractShapeXmlByName(slideXml: string, name: string): string {
  const marker = `name="${name}"`;
  const markerIdx = slideXml.indexOf(marker);
  assert.ok(markerIdx >= 0, `Expected <p:sp> with ${marker} in slide XML`);
  const spStart = slideXml.lastIndexOf("<p:sp>", markerIdx);
  const spEnd = slideXml.indexOf("</p:sp>", markerIdx);
  assert.ok(
    spStart >= 0 && spEnd > spStart,
    `Expected enclosing <p:sp>...</p:sp> for ${marker}`,
  );
  return slideXml.slice(spStart, spEnd + "</p:sp>".length);
}

export const shapeAndGroupRoundtripScenario: CapabilityScenario = {
  id: "mechanism-shape-group-layout-roundtrip",
  capability: "shape",
  description:
    "Minimal dual-author Transitional scenario covering preset shapes, custom geometry, solid/gradient/image fills, lines/arrows, shadows, transforms, z-order, recursive groups, and expanded Box/Flex layout reconstruction.",
  visualCheckpoints: [
    "Page 1: rounded solid shape with shadow & dashed border, linear-gradient ellipse with softEdge, image-filled rect, arrow line, radial-gradient custom bezier shape, and z-ordered overlap.",
    "Page 2: rotated outer Group with Hexagon shape and nested rotated inner Group, plus Flex Box with visible rounded card containing a badge shape.",
    "Page 3: shape with embedded-font text and Group containing a labeled shape and Text member.",
  ],
  nativeCheckpoints: [
    "Initial and roundtrip PPTX retain preset/custom geometry, fills, arrow line ends, z-order, recursive <p:grpSp> members, shape/group text, and expanded layout shapes.",
  ],
  jsx() {
    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "shapes-and-fills" },
        createElement(Shape, {
          id: "preset-solid",
          name: "Preset Solid",
          geometry: "roundRect",
          x: 24,
          y: 24,
          width: 140,
          height: 72,
          fill: "#2563EB",
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "dashed",
            rotate: 15,
            boxShadow: {
              color: { color: "#000000", alpha: 0.35 },
              blur: 4,
              distance: 4,
              angle: 90,
            },
            zIndex: 1,
          },
        }),
        createElement(Shape, {
          id: "preset-linear-grad",
          name: "Preset Linear",
          geometry: "ellipse",
          x: 190,
          y: 24,
          width: 130,
          height: 76,
          style: {
            backgroundImage: linearGradient(135, "#1D4ED8", "#0EA5E9"),
            opacity: 0.8,
            softEdge: 4,
            zIndex: 2,
          },
        }),
        createElement(Shape, {
          id: "preset-image-fill",
          name: "Preset ImageFill",
          geometry: "rect",
          x: 344,
          y: 24,
          width: 120,
          height: 76,
          style: {
            backgroundImage: { image: sampleFillPng },
            borderColor: "#0F172A",
            borderWidth: 2,
          },
        }),
        createElement(Shape, {
          id: "arrow-connector",
          name: "Arrow Connector",
          geometry: preset("line"),
          x: 488,
          y: 56,
          width: 120,
          height: 2,
          flipH: true,
          style: {
            borderColor: "#DC2626",
            borderWidth: 2,
            startArrow: "oval",
            endArrow: "triangle",
          },
        }),
        createElement(Shape, {
          id: "custom-radial",
          name: "Custom Radial",
          geometry: customBezierShape,
          x: 24,
          y: 140,
          width: 140,
          height: 100,
          flipV: true,
          style: {
            backgroundImage: radialGradient("#FEF3C7", "#D97706"),
            borderColor: "#78350F",
            borderWidth: 2,
          },
        }),
        createElement(Shape, {
          id: "front-amber",
          name: "Front Amber",
          x: 220,
          y: 150,
          width: 100,
          height: 70,
          fill: "#F59E0B",
          style: { zIndex: 5 },
        }),
        createElement(Shape, {
          id: "back-slate",
          name: "Back Slate",
          x: 260,
          y: 180,
          width: 100,
          height: 70,
          fill: "#334155",
          style: { zIndex: 1 },
        }),
      ),
      createElement(
        Slide,
        { id: "groups-and-layout" },
        createElement(
          Group,
          { id: "outer-group", name: "Outer Cluster", style: { rotate: 12 } },
          createElement(Shape, {
            id: "member-shape",
            name: "Member Shape",
            geometry: "hexagon",
            x: 32,
            y: 36,
            width: 110,
            height: 64,
            fill: "#10B981",
          }),
          createElement(
            Group,
            {
              id: "inner-group",
              name: "Inner Cluster",
              style: { rotate: -18 },
            },
            createElement(Shape, {
              id: "nested-leaf",
              name: "Nested Leaf",
              geometry: "diamond",
              x: 80,
              y: 124,
              width: 90,
              height: 64,
              fill: "#EC4899",
              flipH: true,
            }),
          ),
        ),
        createElement(
          Box,
          {
            id: "transparent-flex",
            style: {
              position: "absolute",
              left: 340,
              top: 36,
              width: 260,
              display: "flex",
              flexDirection: "column",
              gap: 12,
            },
          },
          createElement(
            Box,
            {
              id: "visible-card",
              style: {
                padding: 12,
                display: "flex",
                flexDirection: "row",
                gap: 10,
                backgroundColor: "#EFF6FF",
                borderColor: "#2563EB",
                borderWidth: 2,
                borderRadius: 10,
              },
            },
            createElement(Shape, {
              id: "boxed-badge",
              name: "Boxed Badge",
              width: 48,
              height: 24,
              fill: "#1D4ED8",
            }),
          ),
        ),
      ),
      createElement(
        Slide,
        { id: "shape-and-group-text" },
        createElement(Shape, {
          id: "text-shape",
          name: "Preset Text",
          geometry: "roundRect",
          x: 24,
          y: 28,
          width: 160,
          height: 64,
          fill: "#1E40AF",
          text: "Solid",
          textStyle: {
            fontFamily: jostFont,
            fontSize: 14,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "middle",
          },
        }),
        createElement(
          Group,
          { id: "text-group", name: "Text Cluster" },
          createElement(Shape, {
            id: "labeled-hex",
            name: "Labeled Hex",
            geometry: "hexagon",
            x: 24,
            y: 120,
            width: 110,
            height: 64,
            fill: "#10B981",
            text: "Hex",
            textStyle: {
              fontFamily: jostFont,
              fontSize: 13,
              color: "#FFFFFF",
              align: "center",
              verticalAlign: "middle",
            },
          }),
          createElement(
            Text,
            {
              id: "member-text",
              name: "Member Text",
              x: 160,
              y: 120,
              width: 140,
              height: 40,
              style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
            },
            "Group Label",
          ),
        ),
      ),
    );
  },
  document() {
    return deck({ width: 640, height: 360 }).append(
      slide({ id: "shapes-and-fills" }).append(
        shape({
          id: "preset-solid",
          name: "Preset Solid",
          geometry: "roundRect",
          x: 24,
          y: 24,
          width: 140,
          height: 72,
          fill: "#2563EB",
          style: {
            borderColor: "#1E293B",
            borderWidth: 3,
            borderStyle: "dashed",
            rotate: 15,
            boxShadow: {
              color: { color: "#000000", alpha: 0.35 },
              blur: 4,
              distance: 4,
              angle: 90,
            },
            zIndex: 1,
          },
        }),
        shape({
          id: "preset-linear-grad",
          name: "Preset Linear",
          geometry: "ellipse",
          x: 190,
          y: 24,
          width: 130,
          height: 76,
          style: {
            backgroundImage: linearGradient(135, "#1D4ED8", "#0EA5E9"),
            opacity: 0.8,
            softEdge: 4,
            zIndex: 2,
          },
        }),
        shape({
          id: "preset-image-fill",
          name: "Preset ImageFill",
          geometry: "rect",
          x: 344,
          y: 24,
          width: 120,
          height: 76,
          style: {
            backgroundImage: { image: sampleFillPng },
            borderColor: "#0F172A",
            borderWidth: 2,
          },
        }),
        shape({
          id: "arrow-connector",
          name: "Arrow Connector",
          geometry: preset("line"),
          x: 488,
          y: 56,
          width: 120,
          height: 2,
          flipH: true,
          style: {
            borderColor: "#DC2626",
            borderWidth: 2,
            startArrow: "oval",
            endArrow: "triangle",
          },
        }),
        shape({
          id: "custom-radial",
          name: "Custom Radial",
          geometry: customBezierShape,
          x: 24,
          y: 140,
          width: 140,
          height: 100,
          flipV: true,
          style: {
            backgroundImage: radialGradient("#FEF3C7", "#D97706"),
            borderColor: "#78350F",
            borderWidth: 2,
          },
        }),
        shape({
          id: "front-amber",
          name: "Front Amber",
          x: 220,
          y: 150,
          width: 100,
          height: 70,
          fill: "#F59E0B",
          style: { zIndex: 5 },
        }),
        shape({
          id: "back-slate",
          name: "Back Slate",
          x: 260,
          y: 180,
          width: 100,
          height: 70,
          fill: "#334155",
          style: { zIndex: 1 },
        }),
      ),
      slide({ id: "groups-and-layout" }).append(
        group({
          id: "outer-group",
          name: "Outer Cluster",
          style: { rotate: 12 },
        }).append(
          shape({
            id: "member-shape",
            name: "Member Shape",
            geometry: "hexagon",
            x: 32,
            y: 36,
            width: 110,
            height: 64,
            fill: "#10B981",
          }),
          group({
            id: "inner-group",
            name: "Inner Cluster",
            style: { rotate: -18 },
          }).append(
            shape({
              id: "nested-leaf",
              name: "Nested Leaf",
              geometry: "diamond",
              x: 80,
              y: 124,
              width: 90,
              height: 64,
              fill: "#EC4899",
              flipH: true,
            }),
          ),
        ),
        box({
          id: "transparent-flex",
          style: {
            position: "absolute",
            left: 340,
            top: 36,
            width: 260,
            display: "flex",
            flexDirection: "column",
            gap: 12,
          },
        }).append(
          box({
            id: "visible-card",
            style: {
              padding: 12,
              display: "flex",
              flexDirection: "row",
              gap: 10,
              backgroundColor: "#EFF6FF",
              borderColor: "#2563EB",
              borderWidth: 2,
              borderRadius: 10,
            },
          }).append(
            shape({
              id: "boxed-badge",
              name: "Boxed Badge",
              width: 48,
              height: 24,
              fill: "#1D4ED8",
            }),
          ),
        ),
      ),
      slide({ id: "shape-and-group-text" }).append(
        shape({
          id: "text-shape",
          name: "Preset Text",
          geometry: "roundRect",
          x: 24,
          y: 28,
          width: 160,
          height: 64,
          fill: "#1E40AF",
          text: "Solid",
          textStyle: {
            fontFamily: jostFont,
            fontSize: 14,
            color: "#FFFFFF",
            align: "center",
            verticalAlign: "middle",
          },
        }),
        group({ id: "text-group", name: "Text Cluster" }).append(
          shape({
            id: "labeled-hex",
            name: "Labeled Hex",
            geometry: "hexagon",
            x: 24,
            y: 120,
            width: 110,
            height: 64,
            fill: "#10B981",
            text: "Hex",
            textStyle: {
              fontFamily: jostFont,
              fontSize: 13,
              color: "#FFFFFF",
              align: "center",
              verticalAlign: "middle",
            },
          }),
          text("Group Label", {
            id: "member-text",
            name: "Member Text",
            x: 160,
            y: 120,
            width: 140,
            height: 40,
            style: { fontFamily: jostFont, fontSize: 14, color: "#0F172A" },
          }),
        ),
      ),
    );
  },
  assert({ validation }) {
    assert.equal(validation.slides.length, 3);
    assert.equal(validation.slides[0]?.objects.length, 7);
    assert.equal(validation.slides[1]?.objects.length, 3);
    assert.equal(validation.slides[2]?.objects.length, 2);
  },
};

export const shapeMechanismVerifications: Pick<
  ExhaustiveMechanismVerificationMap,
  | "preset-shape"
  | "custom-geometry-shape"
  | "solid-fill"
  | "gradient-fill"
  | "image-fill"
> = {
  "preset-shape": {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const s1 = validation.slides[0]!;
      const presetSolid = s1.objects.find((o) => o.name === "Preset Solid");
      assert.equal(presetSolid?.geometry, "roundRect");
      assert.equal(presetSolid?.rotationDegrees, 15);
      const presetLinear = s1.objects.find((o) => o.name === "Preset Linear");
      assert.equal(presetLinear?.geometry, "ellipse");
      const presetImage = s1.objects.find((o) => o.name === "Preset ImageFill");
      assert.equal(presetImage?.geometry, "rect");
      const arrow = s1.objects.find((o) => o.name === "Arrow Connector");
      assert.equal(arrow?.geometry, "line");
      // Verify zIndex ordering in PPTX shape tree: Back Slate precedes Front Amber.
      const backIdx = s1.objects.findIndex((o) => o.name === "Back Slate");
      const frontIdx = s1.objects.findIndex((o) => o.name === "Front Amber");
      assert.ok(backIdx >= 0 && frontIdx > backIdx, `${stage} z-order`);

      const s3 = validation.slides[2]!;
      const textShape = s3.objects.find((o) => o.name === "Preset Text");
      assert.equal(textShape?.text?.text, "Solid");

      const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:headEnd type="oval"\/>/u);
      assert.match(slide1Xml, /<a:tailEnd type="triangle"\/>/u);
      assert.match(slide1Xml, /<a:outerShdw/u);
      assert.match(slide1Xml, /<a:softEdge/u);

      // Slide 2: transparent Box disappeared, visible Box became editable roundRect Shape before Boxed Badge.
      const s2 = validation.slides[1]!;
      assert.equal(s2.objects.length, 3);
      assert.equal(s2.objects[1]?.geometry, "roundRect");
      assert.deepEqual(s2.objects[1]?.boxEmu, {
        x: 340 * 9525,
        y: 36 * 9525,
        width: 260 * 9525,
        height: 48 * 9525,
      });
      assert.equal(s2.objects[2]?.name, "Boxed Badge");
      assert.deepEqual(s2.objects[2]?.boxEmu, {
        x: 352 * 9525,
        y: 48 * 9525,
        width: 48 * 9525,
        height: 24 * 9525,
      });
    },
    assertPaint({ initial, roundtrip }) {
      assert.deepEqual(initial.pages[0]?.diagnostics, []);
      assert.ok(initial.pages[0]?.png instanceof Uint8Array);
      assert.equal(roundtrip.pages.length, 3);
      assert.deepEqual(roundtrip.pages[0]?.diagnostics, []);
      assert.ok(roundtrip.pages[0]?.png instanceof Uint8Array);
      assert.ok((roundtrip.pages[0]?.svg ?? "").includes("<path "));
    },
  },
  "custom-geometry-shape": {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, pptx, validation, importedDocument }) {
      const prefix = `[${shapeAndGroupRoundtripScenario.id}:${stage}:Custom Radial]`;
      assert.equal(validation.errors.length, 0, `${prefix} validation errors`);
      const custom = validation.slides[0]!.objects.find(
        (o) => o.name === "Custom Radial",
      );
      assert.equal(custom?.geometry, "custom", `${prefix} geometry kind`);
      assert.deepEqual(
        custom?.boxEmu,
        {
          x: 24 * 9525,
          y: 140 * 9525,
          width: 140 * 9525,
          height: 100 * 9525,
        },
        `${prefix} boxEmu`,
      );
      const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      const customRadialXml = extractShapeXmlByName(slide1Xml, "Custom Radial");
      assert.match(customRadialXml, /flipV="1"/u, `${prefix} flipV="1"`);
      assert.match(
        customRadialXml,
        /<a:path w="100000" h="100000">/u,
        `${prefix} path viewBox`,
      );
      assert.match(
        customRadialXml,
        /<a:moveTo><a:pt x="10000" y="80000"\/><\/a:moveTo>/u,
        `${prefix} moveTo control point`,
      );
      assert.match(
        customRadialXml,
        /<a:cubicBezTo><a:pt x="10000" y="15000"\/><a:pt x="90000" y="15000"\/><a:pt x="90000" y="80000"\/><\/a:cubicBezTo>/u,
        `${prefix} cubicBezTo control points`,
      );
      assert.match(
        customRadialXml,
        /<a:quadBezTo><a:pt x="50000" y="95000"\/><a:pt x="10000" y="80000"\/><\/a:quadBezTo>/u,
        `${prefix} quadBezTo control points`,
      );
      assert.match(customRadialXml, /<a:close\/>/u, `${prefix} close command`);

      if (stage === "roundtrip" && importedDocument !== undefined) {
        const slide1 = findFirstSlideNode(importedDocument);
        assert.ok(slide1 !== undefined, `${prefix} imported slide1`);
        const importedCustom = findShapeNodeByName(slide1, "Custom Radial");
        assert.ok(
          importedCustom !== undefined,
          `${prefix} imported Custom Radial node`,
        );
        assert.deepEqual(
          JSON.parse(JSON.stringify(importedCustom.props["geometry"])),
          {
            type: "custom",
            width: 100,
            height: 100,
            paths: [{ commands: customBezierShape.paths[0]!.commands }],
          },
          `${prefix} imported geometry commands and control points`,
        );
      }
    },
    assertPaint({ initial, roundtrip }) {
      const expectedPath =
        'd="M38 220 C38 155 150 155 150 220 Q94 235 38 220 Z"';
      for (const [phase, preview] of [
        ["paint:initial", initial],
        ["paint:roundtrip", roundtrip],
      ] as const) {
        const prefix = `[${shapeAndGroupRoundtripScenario.id}:${phase}:Custom Radial]`;
        const page = preview.pages[0];
        assert.ok(page !== undefined, `${prefix} missing page 1`);
        assert.deepEqual(page.diagnostics, [], `${prefix} diagnostics`);
        assert.ok(
          page.svg?.includes(expectedPath),
          `${prefix} expected SVG path ${expectedPath}`,
        );
        assert.ok(page.png instanceof Uint8Array, `${prefix} expected PNG`);
        const decoded = PNG.sync.read(Buffer.from(page.png));
        const offset = (190 * decoded.width + 94) * 4;
        const r = decoded.data[offset] ?? 0;
        const g = decoded.data[offset + 1] ?? 0;
        const b = decoded.data[offset + 2] ?? 0;
        assert.ok(
          !(r === 255 && g === 255 && b === 255) && r > 200 && g > 120,
          `${prefix} expected warm radial-gradient pixel at (94, 190), got rgb(${r}, ${g}, ${b})`,
        );
      }
    },
  },
  "solid-fill": {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:solidFill><a:srgbClr val="2563EB"\/>/u);
    },
    assertPaint({ initial, roundtrip }) {
      assert.ok(initial.pages[0]?.svg?.includes('fill="#2563EB"'));
      assert.ok(roundtrip.pages[0]?.svg?.includes('fill="#2563EB"'));
    },
  },
  "gradient-fill": {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const slide1Xml = strFromU8(unzipSync(pptx)["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:gradFill/u);
      assert.match(slide1Xml, /<a:lin /u);
      assert.match(slide1Xml, /<a:path path="circle">/u);
    },
    assertPaint({ initial, roundtrip }) {
      assert.ok(initial.pages[0]?.svg?.includes("<linearGradient"));
      assert.ok(initial.pages[0]?.svg?.includes("<radialGradient"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("<linearGradient"));
      assert.ok(roundtrip.pages[0]?.svg?.includes("<radialGradient"));
    },
  },
  "image-fill": {
    scenario: shapeAndGroupRoundtripScenario,
    assertNative({ stage, pptx, validation }) {
      assert.equal(validation.errors.length, 0, `${stage} errors`);
      const parts = unzipSync(pptx);
      const slide1Xml = strFromU8(parts["ppt/slides/slide1.xml"]!);
      assert.match(slide1Xml, /<a:blipFill/u);
      assert.ok(
        Object.keys(parts).some((part) => part.startsWith("ppt/media/")),
        `${stage} expected embedded media part for image fill`,
      );
    },
    assertPaint({ initial, roundtrip }) {
      assert.ok(initial.pages[0]?.svg?.includes("<pattern "));
      assert.ok(roundtrip.pages[0]?.svg?.includes("<pattern "));
    },
  },
};
