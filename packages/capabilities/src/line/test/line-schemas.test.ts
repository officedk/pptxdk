import assert from "node:assert/strict";

import { BuildError, createResolver } from "@pptxdk/core/build";
import { createDeckNode, createSlideNode } from "@pptxdk/core/document";
import { describe, it } from "vitest";

import { boxCapability, createBoxNode } from "../../box/box.ts";
import { groupCapability, createGroupNode } from "../../group/group.ts";
import { shapeCapability, createShapeNode } from "../../shape/shape.ts";
import {
  createLineNode,
  lineAdjustSchema,
  lineCapability,
  lineDeclarationPropsSchema,
  lineEndpointSchema,
  linePortSchema,
  lineRoutingSchema,
  lineStrokeSchema,
  lineStrokeValueSchema,
  lineStyleSchema,
} from "../line.ts";

const resolveLineDeck = createResolver([
  lineCapability,
  shapeCapability,
  boxCapability,
  groupCapability,
]);

function expectBuildErrorCode(fn: () => unknown, code: string): void {
  assert.throws(
    fn,
    (error) =>
      error instanceof BuildError &&
      error.diagnostics.some((entry) => entry.code === code),
  );
}

describe("Line schemas and L1 contract", () => {
  it("validates Line flat schemas and rejects forbidden properties", () => {
    assert.equal(lineRoutingSchema.safeParse("straight").success, true);
    assert.equal(lineRoutingSchema.safeParse("elbow").success, true);
    assert.equal(lineRoutingSchema.safeParse("curved").success, true);
    assert.equal(lineRoutingSchema.safeParse("arc").success, false);

    assert.equal(linePortSchema.safeParse("top").success, true);
    assert.equal(linePortSchema.safeParse(2).success, true);
    assert.equal(linePortSchema.safeParse(-1).success, false);
    assert.equal(linePortSchema.safeParse(1.5).success, false);

    assert.equal(
      lineEndpointSchema.safeParse({ id: "node-a", port: "right" }).success,
      true,
    );
    assert.equal(lineEndpointSchema.safeParse({ x: 10, y: 20 }).success, true);
    assert.equal(
      lineEndpointSchema.safeParse({ id: "node-a", x: 10 }).success,
      false,
    );

    assert.equal(
      lineStrokeSchema.safeParse({
        color: "#2563EB",
        width: 2,
        style: "dashed",
        startArrow: "oval",
        endArrow: "triangle",
      }).success,
      true,
    );
    assert.equal(lineStrokeValueSchema.safeParse("none").success, true);
    assert.equal(lineAdjustSchema.safeParse({ adj1: 35000 }).success, true);
    assert.equal(lineAdjustSchema.safeParse({ adj: 35000 }).success, false);

    assert.equal(
      lineStyleSchema.safeParse({
        position: "absolute",
        left: 20,
        top: 40,
        width: 200,
        height: 0,
        opacity: 0.8,
      }).success,
      true,
    );
    assert.equal(lineStyleSchema.safeParse({ padding: 8 }).success, false);
    assert.equal(
      lineStyleSchema.safeParse({ background: "#FF0000" }).success,
      false,
    );
    assert.equal(
      lineStyleSchema.safeParse({ border: { color: "#000000" } }).success,
      false,
    );

    assert.equal(
      lineDeclarationPropsSchema.safeParse({
        routing: "elbow",
        adjust: { adj1: 50000 },
        stroke: { color: "accent1", width: 2, endArrow: "arrow" },
        from: { id: "a", port: "right" },
        to: { id: "b", port: "left" },
        style: { opacity: 0.9 },
      }).success,
      true,
    );
  });

  it("supports single-axis zero extent in box mode while rejecting 0x0 and children", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        createLineNode({
          id: "horiz",
          style: {
            position: "absolute",
            left: 40,
            top: 80,
            width: 300,
            height: 0,
          },
        }),
        createLineNode({
          id: "vert",
          stroke: {
            color: "#0F172A",
            width: 2,
            style: "dotted",
            endArrow: "stealth",
          },
          style: {
            position: "absolute",
            left: 100,
            top: 20,
            width: 0,
            height: 180,
          },
        }),
        createGroupNode({ id: "grp" }).append(
          createShapeNode({
            id: "grp-box",
            style: {
              position: "absolute",
              left: 400,
              top: 40,
              width: 120,
              height: 60,
            },
          }),
          createLineNode({
            id: "grp-line",
            style: {
              position: "absolute",
              left: 400,
              top: 120,
              width: 120,
              height: 0,
            },
          }),
        ),
      ),
    );

    const res = resolveLineDeck(doc);
    assert.deepEqual(res.slide("s1").get("horiz").box, {
      x: 40,
      y: 80,
      width: 300,
      height: 0,
    });
    assert.deepEqual(res.slide("s1").get("vert").box, {
      x: 100,
      y: 20,
      width: 0,
      height: 180,
    });
    assert.equal(res.slide("s1").get("horiz").objects[0]?.kind, "line");

    expectBuildErrorCode(
      () =>
        resolveLineDeck(
          createDeckNode({ width: 960, height: 540 }).append(
            createSlideNode({ id: "s1" }).append(
              createLineNode({
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 0,
                  height: 0,
                },
              }),
            ),
          ),
        ),
      "invalid-style",
    );
  });

  it("resolves endpoint mode geometry, port transforms, flip direction, and rejects invalid combinations", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        createShapeNode({
          id: "source",
          style: {
            position: "absolute",
            left: 300,
            top: 200,
            width: 100,
            height: 60,
          },
        }),
        // Forward ref + reverse quadrant (x2 < x1, y2 < y1 => flipH=true, flipV=true)
        createLineNode({
          id: "conn",
          routing: "curved",
          adjust: { adj1: 40000 },
          stroke: {
            color: "#DC2626",
            width: 2,
            startArrow: "oval",
            endArrow: "triangle",
          },
          from: { id: "source", port: "left" },
          to: { id: "target", port: "bottom" },
        }),
        createShapeNode({
          id: "target",
          style: {
            position: "absolute",
            left: 100,
            top: 50,
            width: 80,
            height: 40,
          },
        }),
        createLineNode({
          id: "rotated-conn",
          from: { id: "rotated-node", port: "top" },
          to: { x: 600, y: 200 },
        }),
        createShapeNode({
          id: "rotated-node",
          style: {
            position: "absolute",
            left: 500,
            top: 100,
            width: 80,
            height: 40,
            rotate: 90,
          },
        }),
      ),
    );

    const res = resolveLineDeck(doc);
    // source.left = (300, 230); target.bottom = (140, 90)
    assert.deepEqual(res.slide("s1").get("conn").box, {
      x: 140,
      y: 90,
      width: 160,
      height: 140,
    });
    // rotated-node center=(540, 120), local top=(0, -20) rotated +90 deg clockwise => (+20, 0) => (560, 120)
    assert.deepEqual(res.slide("s1").get("rotated-conn").box, {
      x: 560,
      y: 120,
      width: 40,
      height: 80,
    });

    // Single endpoint rejected
    expectBuildErrorCode(
      () =>
        resolveLineDeck(
          createDeckNode({ width: 960, height: 540 }).append(
            createSlideNode({ id: "s1" }).append(
              createLineNode({ from: { x: 0, y: 0 } }),
            ),
          ),
        ),
      "invalid-prop",
    );

    // Endpoint + manual width/left rejected
    expectBuildErrorCode(
      () =>
        resolveLineDeck(
          createDeckNode({ width: 960, height: 540 }).append(
            createSlideNode({ id: "s1" }).append(
              createLineNode({
                from: { x: 0, y: 0 },
                to: { x: 100, y: 50 },
                style: { width: 100 },
              }),
            ),
          ),
        ),
      "invalid-style",
    );

    // Out of bounds numeric port rejected
    expectBuildErrorCode(
      () =>
        resolveLineDeck(
          createDeckNode({ width: 960, height: 540 }).append(
            createSlideNode({ id: "s1" }).append(
              createShapeNode({
                id: "a",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 40,
                },
              }),
              createLineNode({
                from: { id: "a", port: 4 },
                to: { x: 200, y: 200 },
              }),
            ),
          ),
        ),
      "invalid-prop",
    );

    // Non-native Box rejected as broken-reference
    expectBuildErrorCode(
      () =>
        resolveLineDeck(
          createDeckNode({ width: 960, height: 540 }).append(
            createSlideNode({ id: "s1" }).append(
              createBoxNode({
                id: "logical",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 40,
                },
              }),
              createLineNode({
                from: { id: "logical", port: "right" },
                to: { x: 200, y: 200 },
              }),
            ),
          ),
        ),
      "broken-reference",
    );
  });
});
