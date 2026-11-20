import assert from "node:assert/strict";

import {
  deck,
  documentCapabilities,
  paragraph,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import {
  blockArc,
  linearGradient,
  preset,
} from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import { P } from "@react-pptx/capabilities/text-elements";
import type { CapabilityModule, CommonNode } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

function runDocPreview(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return {
      query: resolution.query,
      ...emit(resolution),
      ...paint(resolution),
    };
  };
}

describe("Slice 2: Shape static preview", () => {
  it("renders rect, ellipse, preset, custom path, transforms, zIndex, and screenGeometry", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const custom = {
      type: "custom" as const,
      width: 10,
      height: 10,
      paths: [
        {
          commands: [
            { op: "M" as const, x: 0, y: 0 },
            { op: "L" as const, x: 10, y: 0 },
            { op: "L" as const, x: 10, y: 10 },
            { op: "L" as const, x: 0, y: 10 },
            { op: "Z" as const },
          ],
          fill: true,
          stroke: false,
        },
        {
          commands: [
            { op: "M" as const, x: 0, y: 0 },
            { op: "Q" as const, x1: 5, y1: 10, x: 10, y: 0 },
            { op: "C" as const, x1: 8, y1: 4, x2: 2, y2: 4, x: 0, y: 0 },
          ],
          fill: false,
          stroke: true,
        },
      ],
    };
    const mutableShape = shape({
      id: "hero-rect",
      x: 20,
      y: 20,
      width: 100,
      height: 50,
      fill: "#FF0000",
      style: {
        borderColor: "#000000",
        borderWidth: 2,
        borderStyle: "solid",
        rotate: 90,
        zIndex: 2,
      },
    });
    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "shapes" }).append(
        mutableShape,
        shape({
          id: "underlay",
          geometry: preset("ellipse"),
          x: 20,
          y: 20,
          width: 100,
          height: 50,
          style: { backgroundColor: "#0000FF", zIndex: 1 },
        }),
        shape({
          id: "custom-path",
          geometry: custom,
          x: 160,
          y: 20,
          width: 100,
          height: 50,
          fill: "accent1",
          flipH: true,
          flipV: true,
          style: {
            borderColor: "#112233",
            borderWidth: 2,
            opacity: 0.5,
          },
        }),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // zIndex 1 underlay must appear before zIndex 2 hero-rect in SVG paint order.
    expect(page.svg.indexOf('fill="#0000FF"')).toBeLessThan(
      page.svg.indexOf('fill="#FF0000"'),
    );
    // Uniform path evaluation renders all shapes (including ellipse) as SVG <path> elements.
    expect(page.svg).toContain("<path ");
    expect(page.svg).toContain('fill="none"');
    expect(page.svg).toContain('stroke="none"');
    expect(page.svg).toContain('stroke-width="2"');

    const heroNode = page.nodes.find((node) => node.authorId === "hero-rect")!;
    expect(heroNode.screenBox?.x).toBeCloseTo(45);
    expect(heroNode.screenBox?.y).toBeCloseTo(-5);
    expect(heroNode.screenBox?.width).toBeCloseTo(50);
    expect(heroNode.screenBox?.height).toBeCloseTo(100);

    // Mutate after preview and rebuild: first result stays untouched.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    mutableShape.set({ fill: "#00AA00" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.notEqual(second.pages[0]?.svg, firstSvg);
    expect(second.pages[0]?.svg).toContain('fill="#00AA00"');
  });

  it("renders standard preset geometries (including adjust, shaded subpaths, local path w/h, and connectors) with zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({ width: 480, height: 240 }).append(
      slide({ id: "presets" }).append(
        shape({
          id: "star",
          geometry: "star5",
          x: 20,
          y: 20,
          width: 80,
          height: 80,
          fill: "#F59E0B",
        }),
        shape({
          id: "arc",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 120,
          y: 20,
          width: 80,
          height: 80,
          fill: "#10B981",
        }),
        shape({
          id: "bevel-box",
          geometry: preset("bevel", { adj: 20000 }),
          x: 220,
          y: 20,
          width: 100,
          height: 60,
          fill: "#3B82F6",
          style: { borderColor: "#1E3A8A", borderWidth: 2 },
        }),
        shape({
          id: "cloud-shape",
          geometry: "cloud",
          x: 340,
          y: 20,
          width: 110,
          height: 70,
          fill: "#E0F2FE",
          style: { borderColor: "#0284C7", borderWidth: 2 },
        }),
        shape({
          id: "donut-ring",
          geometry: "donut",
          x: 20,
          y: 120,
          width: 80,
          height: 80,
          fill: "#EC4899",
        }),
        shape({
          id: "connector",
          geometry: preset("bentConnector3", { adj1: 30000 }),
          x: 130,
          y: 130,
          width: 120,
          height: 60,
          style: { borderColor: "#0F172A", borderWidth: 2 },
        }),
      ),
    );
    const result = await previewDoc(doc);
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // Bevel produces shaded fill variants alongside the base #3B82F6 face.
    expect(page.svg).toContain('fill="#3B82F6"');
    expect(page.svg).not.toEqual(
      page.svg.replace(/fill="#(?!3B82F6)[0-9A-F]{6}"/u, ""),
    );
  });

  it("renders gradients, image fills, borderRadius, dashed/dotted borders, open-path arrows, boxShadow, and softEdge with zero diagnostics", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const { PNG } = await import("pngjs");
    const { arrowTypes } = await import("@react-pptx/capabilities");
    const { path } = await import("@react-pptx/capabilities/shape");

    const samplePng = new PNG({ width: 4, height: 4 });
    for (let i = 0; i < samplePng.data.length; i += 4) {
      samplePng.data.set([0x16, 0xa3, 0x4a, 255], i);
    }
    const pngBytes = new Uint8Array(PNG.sync.write(samplePng));

    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "visuals", background: "#FFFFFF" }).append(
        shape({
          id: "grad-shape",
          x: 10,
          y: 10,
          width: 90,
          height: 40,
          style: {
            backgroundImage: linearGradient(90, "#111111", "#888888"),
            borderRadius: 8,
            borderColor: "#0F172A",
            borderWidth: 2,
            borderStyle: "dashed",
          },
        }),
        shape({
          id: "img-shape",
          geometry: "ellipse",
          x: 120,
          y: 10,
          width: 80,
          height: 40,
          style: {
            backgroundImage: { image: pngBytes },
            borderColor: "#064E3B",
            borderWidth: 2,
            borderStyle: "dotted",
          },
        }),
        shape({
          id: "shadow-shape",
          x: 220,
          y: 15,
          width: 70,
          height: 35,
          fill: "#112233",
          style: {
            rotate: 45,
            boxShadow: { blur: 4, distance: 6, angle: 90 },
          },
        }),
        shape({
          id: "soft-shape",
          x: 320,
          y: 10,
          width: 80,
          height: 40,
          fill: "#1D4ED8",
          style: { softEdge: 6 },
        }),
        ...arrowTypes.map((arrowType, idx) =>
          shape({
            id: `arrow-${arrowType}`,
            geometry:
              idx % 2 === 0
                ? "line"
                : path("M 0 12 Q 50 -6 100 12", {
                    viewBox: [100, 18],
                    fill: false,
                  }),
            x: 15 + idx * 75,
            y: 100,
            width: 60,
            height: 16,
            style: {
              borderColor: "#0F172A",
              borderWidth: 2.5,
              startArrow: arrowType,
              endArrow: arrowType,
            },
          }),
        ),
      ),
    );
    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    expect(page.svg).toContain("<linearGradient");
    expect(page.svg).toContain("<pattern");
    expect(page.svg).toContain("<feDropShadow");
    expect(page.svg).toContain("<feGaussianBlur");
    expect(page.svg).toContain("stroke-dasharray");
  });

  it("rejects inline text inside shapes with located errors", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const cases: readonly {
      label: string;
      node: ReturnType<typeof shape>;
      expected: RegExp;
    }[] = [
      {
        label: "inline-text",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
          text: "Inline",
        }),
        expected: /text/i,
      },
      {
        label: "rich-children-text",
        node: shape({
          id: "bad",
          x: 10,
          y: 10,
          width: 60,
          height: 20,
          fill: "#112233",
        }).append(paragraph("Child")),
        expected: /text/i,
      },
    ];

    for (const item of cases) {
      const result = await previewDoc(
        deck({ width: 200, height: 120 }).append(
          slide({ id: "s1" }).append(item.node),
        ),
      );
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, [], item.label);
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /shape\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
    }

    // Verify JSX shape with children text also rejects cleanly.
    const jsxWithText = await expandDocument(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "jsx-text" },
          createElement(
            Shape,
            { id: "bad-jsx", x: 10, y: 10, width: 60, height: 20 },
            createElement(P, null, "Hello"),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxWithText);
    assert.equal(jsxResult.pages[0]?.png, undefined);
    assert.match(
      jsxResult.pages[0]?.diagnostics[0]?.path ?? "",
      /shape\[bad-jsx\]/,
    );
  });
});
