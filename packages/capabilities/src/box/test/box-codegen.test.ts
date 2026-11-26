import assert from "node:assert/strict";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import { createMorphScope } from "@pptxdk/core/capability";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import { box, deck, shape, slide, text } from "../../document.ts";

describe("box capability codegen", () => {
  it("preserves Flex, Grid, and Block layout rules and nested Box children in vanilla and react source", () => {
    const morphScope = createMorphScope();
    const doc = deck({ width: 960, height: 540 }).append(
      slide({ id: "layout" }).append(
        box({
          id: "flex-wrap-container",
          morphKey: morphScope.handle("panel"),
          theme: { colors: { panelBg: "#F8FAFC" } },
          style: {
            position: "absolute",
            left: 32,
            top: 32,
            width: 300,
            display: "flex",
            flexDirection: "row",
            flexWrap: "wrap",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
            padding: 16,
            background: "panelBg",
            border: { color: "#CBD5E1", width: 1 },
            borderRadius: 8,
          },
        }).append(
          shape({
            id: "item-1",
            style: { width: 120, height: 48, background: "#2563EB" },
          }),
          shape({
            id: "item-2",
            style: { width: 120, height: 48, background: "#16A34A" },
          }),
          box({
            id: "nested-grid",
            style: {
              width: 268,
              display: "grid",
              gridTemplateColumns: "1fr 2fr",
              columnGap: 10,
              rowGap: 8,
            },
          }).append(
            text("Left", { id: "grid-left", style: { fontSize: 14 } }),
            text("Right", { id: "grid-right", style: { fontSize: 14 } }),
          ),
        ),
        box(),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry];
    assert.equal(typeof vanillaSource, "string");
    if (typeof vanillaSource !== "string") return;
    assert.match(vanillaSource, /display: "flex"/u);
    assert.match(vanillaSource, /flexDirection: "row"/u);
    assert.match(vanillaSource, /flexWrap: "wrap"/u);
    assert.match(vanillaSource, /justifyContent: "space-between"/u);
    assert.match(vanillaSource, /gap: 12/u);
    assert.match(vanillaSource, /display: "grid"/u);
    assert.match(vanillaSource, /gridTemplateColumns: "1fr 2fr"/u);
    assert.match(vanillaSource, /morphScope\.handle\("panel"\)/u);
    assert.match(vanillaSource, /box\(\)/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry];
    assert.equal(typeof reactSource, "string");
    if (typeof reactSource !== "string") return;
    assert.match(reactSource, /<Box[\s\S]*?id="flex-wrap-container"/u);
    assert.match(reactSource, /gridTemplateColumns: "1fr 2fr"/u);
    assert.match(reactSource, /<Box \/>/u);
  });

  it("throws CodegenError with precise path when Box contains invalid style numbers", () => {
    const generate = createCodegen(documentCapabilities);

    const badGapDoc = deck({ width: 960, height: 540 }).append(
      slide({ id: "s1" }).append(
        box({ id: "bad-box", style: { gap: Number.NaN } }),
      ),
    );

    assert.throws(
      () => generate(badGapDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/box[bad-box]@style.gap",
    );

    assert.throws(
      () => generate(badGapDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/box[bad-box]@style.gap",
    );
  });
});
