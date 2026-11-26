import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { createMorphScope } from "@pptxdk/core/capability";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import { deck, group, shape, slide, text } from "../../document.ts";

describe("group capability codegen", () => {
  it("preserves nested Group hierarchy, member order, bounds, local transforms, and identity in vanilla and react source", () => {
    const morphScope = createMorphScope();
    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "groups" }).append(
        group({
          id: "outer-cluster",
          name: "Outer Cluster",
          morphKey: morphScope.handle("cluster"),
          style: {
            position: "absolute",
            left: 40,
            top: 40,
            width: 320,
            height: 200,
            rotate: 12,
          },
        }).append(
          shape({
            id: "first-member",
            name: "First Member",
            geometry: "hexagon",
            style: {
              position: "absolute",
              left: 10,
              top: 20,
              width: 100,
              height: 80,
              background: "#2563EB",
              flipH: true,
            },
          }),
          group({
            id: "inner-cluster",
            name: "Inner Cluster",
            style: { rotate: -18 },
          }).append(
            shape({
              id: "inner-leaf",
              name: "Inner Leaf",
              geometry: "diamond",
              style: {
                position: "absolute",
                left: 140,
                top: 40,
                width: 90,
                height: 60,
                background: "#10B981",
                flipV: true,
                rotate: 30,
              },
            }),
            text("Member Label", {
              id: "member-label",
              name: "Member Label",
              style: {
                position: "absolute",
                left: 140,
                top: 110,
                width: 120,
                height: 32,
              },
            }),
          ),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry];
    assert.equal(typeof vanillaSource, "string");
    if (typeof vanillaSource !== "string") return;
    assert.match(vanillaSource, /name: "Outer Cluster"/u);
    assert.match(
      vanillaSource,
      /style: \{\s*position: "absolute",\s*left: 40,\s*top: 40,\s*width: 320,\s*height: 200,\s*rotate: 12/u,
    );
    assert.match(vanillaSource, /flipH: true/u);
    assert.match(vanillaSource, /flipV: true/u);
    assert.match(vanillaSource, /morphScope\.handle\("cluster"\)/u);
    const firstIdx = vanillaSource.indexOf("First Member");
    const innerIdx = vanillaSource.indexOf("Inner Cluster");
    const leafIdx = vanillaSource.indexOf("Inner Leaf");
    const labelIdx = vanillaSource.indexOf("Member Label");
    assert.ok(firstIdx > 0 && innerIdx > firstIdx);
    assert.ok(leafIdx > innerIdx && labelIdx > leafIdx);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry];
    assert.equal(typeof reactSource, "string");
    if (typeof reactSource !== "string") return;
    assert.match(reactSource, /<Group[\s\S]*?id="outer-cluster"/u);
    assert.match(reactSource, /<Group[\s\S]*?id="inner-cluster"/u);
    assert.match(reactSource, /rotate: -18/u);
  });

  it("throws CodegenError with precise path when Group has non-finite bounds or rotation", () => {
    const generate = createCodegen(documentCapabilities);

    const badGroupDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1" }).append(
        group({
          id: "bad-group",
          style: { rotate: Number.POSITIVE_INFINITY },
        }).append(
          shape({
            style: {
              position: "absolute",
              left: 10,
              top: 10,
              width: 40,
              height: 40,
              background: "#000000",
            },
          }),
        ),
      ),
    );

    assert.throws(
      () => generate(badGroupDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide[s1]/group[bad-group]@style.rotate",
    );
  });
});
