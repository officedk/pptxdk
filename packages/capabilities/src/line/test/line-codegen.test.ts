import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import {
  CommonContainer,
  createDeckNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { describe, it } from "vitest";

import { fade } from "../../motion/motion-api.ts";
import { motionCapability } from "../../motion/motion.ts";
import { shapeCapability } from "../../shape/shape.ts";
import { createLineNode, lineCapability } from "../line.ts";

const generateLineCode = createCodegen([
  lineCapability,
  shapeCapability,
  motionCapability,
]);

describe("line codegen L1 unit tests", () => {
  it("emits vanilla and React source for box-mode and endpoint-mode lines", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        createLineNode({
          id: "divider",
          style: {
            position: "absolute",
            left: 40,
            top: 80,
            width: 880,
            height: 0,
            opacity: 0.75,
          },
          stroke: { color: "#94A3B8", width: 1.5, style: "dashed" },
        }),
        createLineNode({
          id: "arrow",
          routing: "elbow",
          adjust: { adj1: 45000 },
          stroke: { color: "accent1", width: 2, endArrow: "triangle" },
          from: { id: "a", port: "right" },
          to: { x: 400, y: 200 },
          enter: fade(),
        }),
      ),
    );

    const vanilla = generateLineCode(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /\bline\(\s*\{/u);
    assert.match(vanillaSource, /routing:\s*"elbow"/u);
    assert.match(
      vanillaSource,
      /from:\s*\{\s*id:\s*"a",\s*port:\s*"right"\s*\}/u,
    );
    assert.match(vanillaSource, /to:\s*\{\s*x:\s*400,\s*y:\s*200\s*\}/u);
    assert.match(vanillaSource, /endArrow:\s*"triangle"/u);
    assert.match(vanillaSource, /fade\(/u);

    const react = generateLineCode(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /<Line\b/u);
    assert.match(reactSource, /routing="elbow"/u);
    assert.match(
      reactSource,
      /from=\{\{\s*id:\s*"a",\s*port:\s*"right"\s*\}\}/u,
    );
    assert.match(reactSource, /endArrow:\s*"triangle"/u);
    assert.match(reactSource, /enter=\{fade\(/u);
  });

  it("omits empty style objects and rejects children or invalid styles", () => {
    const emptyStyleDoc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        createLineNode({
          from: { x: 10, y: 20 },
          to: { x: 110, y: 20 },
          style: {},
        }),
      ),
    );
    const vanilla = generateLineCode(emptyStyleDoc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.doesNotMatch(vanillaSource, /\bstyle:\s*\{\s*\}/u);

    const badStyleDoc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        createLineNode({
          style: { background: "#FF0000" } as Record<string, unknown>,
        }),
      ),
    );
    assert.throws(
      () => generateLineCode(badStyleDoc, { target: "vanilla" }),
      (error) =>
        error instanceof CodegenError &&
        error.diagnostics.some((entry) => entry.code === "unexpressible-value"),
    );

    const withChildren = new CommonContainer("line", "line", {
      style: { width: 100, height: 0 },
    });
    withChildren.append(createLineNode({ style: { width: 10, height: 0 } }));
    const childDoc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(withChildren),
    );
    assert.throws(
      () => generateLineCode(childDoc, { target: "react" }),
      (error) =>
        error instanceof CodegenError &&
        error.diagnostics.some((entry) => entry.code === "invalid-document"),
    );
  });
});
