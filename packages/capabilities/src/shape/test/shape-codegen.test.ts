import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { createMorphScope } from "@pptxdk/core/capability";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  blockArc,
  deck,
  linearGradient,
  paragraph,
  path,
  preset,
  radialGradient,
  shape,
  slide,
  span,
} from "../../document.ts";

describe("shape capability codegen", () => {
  it("generates vanilla and react source for preset geometries, adjusts, custom paths, arrows, transforms, and text modes", () => {
    const morphScope = createMorphScope();
    const compositeGeometry = {
      type: "custom" as const,
      width: 100,
      height: 80,
      paths: [
        {
          commands: [
            { op: "M" as const, x: 0, y: 0 },
            { op: "L" as const, x: 100, y: 0 },
            { op: "L" as const, x: 100, y: 80 },
            { op: "Z" as const },
          ],
          fill: true,
          stroke: false,
        },
        {
          commands: [
            { op: "M" as const, x: 10, y: 70 },
            { op: "C" as const, x1: 20, y1: 10, x2: 80, y2: 10, x: 90, y: 70 },
            { op: "Q" as const, x1: 50, y1: 80, x: 10, y: 70 },
          ],
          fill: false,
          stroke: true,
        },
      ],
    };

    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "shapes" }).append(
        shape({
          id: "rounded",
          name: "Rounded Card",
          morphKey: morphScope.handle("card"),
          link: { url: "https://example.com" },
          geometry: preset("roundRect", { adj: 20000 }),
          x: 24,
          y: 24,
          width: 160,
          height: 80,
          fill: "#2563EB",
          text: "Preset Label",
          textStyle: { fontSize: 16, color: "#FFFFFF" },
          style: {
            borderColor: "#1E293B",
            borderWidth: 2,
            borderStyle: "dashed",
            rotate: 15,
          },
        }),
        shape({
          id: "arc-shape",
          geometry: blockArc({ startAngle: 0, endAngle: 180, thickness: 0.3 }),
          x: 200,
          y: 24,
          width: 100,
          height: 100,
          style: {
            backgroundImage: linearGradient(90, "#1D4ED8", "#0EA5E9"),
          },
        }),
        shape({
          id: "bezier-single",
          geometry: path("M 10 80 C 10 15 90 15 90 80 Q 50 95 10 80 Z", {
            viewBox: [100, 100],
            fill: true,
            stroke: true,
          }),
          x: 320,
          y: 24,
          width: 120,
          height: 100,
          flipH: true,
          flipV: true,
          style: {
            backgroundImage: radialGradient("#FDE047", "#D97706"),
            borderColor: "#78350F",
            borderWidth: 2,
          },
        }),
        shape({
          id: "composite-custom",
          geometry: compositeGeometry,
          x: 460,
          y: 24,
          width: 120,
          height: 96,
          fill: "#10B981",
          style: {
            borderColor: "#064E3B",
            borderWidth: 2,
          },
        }),
        shape({
          id: "arrow-line",
          geometry: "line",
          x: 24,
          y: 160,
          width: 180,
          height: 2,
          style: {
            borderColor: "#DC2626",
            borderWidth: 3,
            startArrow: "oval",
            endArrow: "triangle",
          },
        }),
        shape({
          id: "rich-shape",
          geometry: "diamond",
          x: 24,
          y: 200,
          width: 220,
          height: 100,
          fill: "#0F766E",
        }).append(
          paragraph([
            span("Bold ", { style: { fontWeight: "bold", color: "#FFFFFF" } }),
            span("body", { style: { color: "#E2E8F0" } }),
          ]),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry];
    assert.equal(typeof vanillaSource, "string");
    if (typeof vanillaSource !== "string") return;
    assert.match(vanillaSource, /preset\("roundRect", \{ adj: 20000 \}\)/u);
    assert.match(
      vanillaSource,
      /preset\("blockArc", \{ adj1: 0, adj2: 10800000, adj3: 30000 \}\)/u,
    );
    assert.match(
      vanillaSource,
      /path\("M 10 80 C 10 15 90 15 90 80 Q 50 95 10 80 Z", \{/u,
    );
    assert.match(vanillaSource, /type: "custom"/u);
    assert.match(vanillaSource, /startArrow: "oval"/u);
    assert.match(vanillaSource, /endArrow: "triangle"/u);
    assert.match(vanillaSource, /flipH: true/u);
    assert.match(vanillaSource, /flipV: true/u);
    assert.match(vanillaSource, /linearGradient\(90,/u);
    assert.match(vanillaSource, /radialGradient\(/u);
    assert.match(vanillaSource, /morphScope\.handle\("card"\)/u);
    assert.match(vanillaSource, /text: "Preset Label"/u);
    assert.match(vanillaSource, /id: "rich-shape"[\s\S]*?\.append\(/u);
    assert.doesNotMatch(vanillaSource, /id: "rich-shape"[\s\S]*?text:/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry];
    assert.equal(typeof reactSource, "string");
    if (typeof reactSource !== "string") return;
    assert.match(
      reactSource,
      /geometry=\{preset\("roundRect", \{ adj: 20000 \}\)\}/u,
    );
    assert.match(
      reactSource,
      /geometry=\{path\("M 10 80 C 10 15 90 15 90 80 Q 50 95 10 80 Z"/u,
    );
    assert.match(reactSource, /flipH/u);
    assert.match(reactSource, /flipV/u);
    assert.match(reactSource, /<Shape[\s\S]*?id="rich-shape"[\s\S]*?>\s*<P>/u);
  });

  it("throws CodegenError with precise path on non-finite custom geometry values or conflicting text and children", () => {
    const generate = createCodegen(documentCapabilities);

    const invalidCustomDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "bad-geom" }).append(
        shape({
          id: "bad-path",
          x: 10,
          y: 10,
          width: 100,
          height: 100,
          geometry: {
            type: "custom",
            width: 100,
            height: 100,
            paths: [
              {
                commands: [
                  { op: "M", x: 0, y: 0 },
                  { op: "L", x: Number.NaN, y: 50 },
                ],
              },
            ],
          },
        }),
      ),
    );

    assert.throws(
      () => generate(invalidCustomDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path.includes("shape[bad-path]@geometry"),
    );

    const conflictingShape = shape({
      id: "conflict",
      x: 10,
      y: 10,
      width: 100,
      height: 50,
      text: "Inline",
    }).append(paragraph("Child"));

    const conflictDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1" }).append(conflictingShape),
    );

    assert.throws(
      () => generate(conflictDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/shape[conflict]@text",
    );
  });
});
