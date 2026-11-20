import assert from "node:assert/strict";

import { Box } from "@react-pptx/capabilities/box-elements";
import {
  box,
  deck,
  documentCapabilities,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import { linearGradient } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
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

describe("Slice 4 (Ticket 05): Box static preview positive and negative behaviors", () => {
  it("renders transparent and visible Box containers with solid fill/border, opacity, rotate, theme overrides, and post-preview isolation", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const mutableCard = box({
      id: "visible-card",
      theme: { colors: { cardTone: "#1D4ED8" } },
      style: {
        width: 140,
        height: 80,
        padding: 12,
        backgroundColor: "cardTone",
        borderColor: "#0F172A",
        borderWidth: 2,
        borderStyle: "solid",
        opacity: 0.85,
        rotate: 90,
      },
    }).append(
      shape({
        id: "card-dot",
        width: 24,
        height: 24,
        fill: "#FACC15",
      }),
    );

    const doc = deck({
      width: 360,
      height: 200,
      theme: { colors: { cardTone: "#991B1B" } },
    }).append(
      slide({ id: "boxes", background: "#FFFFFF" }).append(
        box({
          id: "transparent-row",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            display: "flex",
            flexDirection: "row",
            gap: 20,
            padding: 8,
          },
        }).append(
          mutableCard,
          box({
            id: "border-only",
            style: {
              width: 100,
              height: 60,
              padding: 8,
              borderColor: "#16A34A",
              borderWidth: 3,
            },
          }).append(
            shape({
              id: "inner-rect",
              width: 30,
              height: 20,
              fill: "#0EA5E9",
            }),
          ),
        ),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    // Transparent outer box does not emit a native shape object; only visible boxes and child shapes become objects.
    assert.equal(page.nodes.length, 4);
    assert.equal(
      page.nodes.some((node) => node.authorId === "transparent-row"),
      false,
    );
    assert.equal(page.layouts[0]?.authorId, "transparent-row");
    expect(page.svg).toContain('fill="#1D4ED8"');
    expect(page.svg).toContain('stroke="#16A34A"');
    // Visible box background appears before its child dot in SVG paint order.
    expect(page.svg.indexOf('fill="#1D4ED8"')).toBeLessThan(
      page.svg.indexOf('fill="#FACC15"'),
    );
    const rotatedCard = page.nodes.find((n) => n.authorId === "visible-card")!;
    expect(rotatedCard.screenBox?.width).toBeCloseTo(80);
    expect(rotatedCard.screenBox?.height).toBeCloseTo(140);

    // Mutate after preview and rebuild: prior PNG, SVG, and PPTX remain untouched.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    const firstPptx = first.pptx.slice();
    mutableCard.style({ backgroundColor: "#059669" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.pptx, firstPptx);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // JSX entry renders transparent and visible Box identically.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 240, height: 140 },
        createElement(
          Slide,
          { id: "jsx-box", background: "#FFFFFF" },
          createElement(
            Box,
            {
              id: "jsx-transparent",
              style: { display: "flex", padding: 10 },
            },
            createElement(
              Box,
              {
                id: "jsx-visible",
                style: {
                  width: 120,
                  height: 60,
                  padding: 8,
                  backgroundColor: "#E2E8F0",
                  borderColor: "#334155",
                  borderWidth: 2,
                },
              },
              createElement(Shape, {
                id: "jsx-child",
                width: 40,
                height: 20,
                fill: "#EF4444",
              }),
            ),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    expect(jsxResult.pages[0]?.svg).toContain('fill="#E2E8F0"');
  });

  it("rejects unsupported visible Box preview styles while retaining valid sibling pages and PPTX", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const rejectionCases: readonly {
      readonly label: string;
      readonly node: ReturnType<typeof box>;
      readonly expected: RegExp;
    }[] = [
      {
        label: "backgroundImage",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundImage: linearGradient(90, "#1D4ED8", "#9333EA"),
          },
        }),
        expected: /backgroundImage/i,
      },
      {
        label: "borderRadius",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            borderRadius: 12,
          },
        }),
        expected: /borderRadius/i,
      },
      {
        label: "boxShadow",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            boxShadow: { blur: 4, distance: 2 },
          },
        }),
        expected: /boxShadow/i,
      },
      {
        label: "softEdge",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            backgroundColor: "#E2E8F0",
            softEdge: 4,
          },
        }),
        expected: /softEdge/i,
      },
      {
        label: "dashed-border",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            borderColor: "#0F172A",
            borderStyle: "dashed",
          },
        }),
        expected: /non-solid/i,
      },
      {
        label: "dotted-border",
        node: box({
          id: "bad",
          style: {
            width: 100,
            height: 50,
            borderColor: "#0F172A",
            borderStyle: "dotted",
          },
        }),
        expected: /non-solid/i,
      },
    ];

    for (const item of rejectionCases) {
      const result = await previewDoc(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "bad-box-page" }).append(item.node),
          slide({ id: "good-box-page" }).append(
            box({
              id: "ok",
              style: { width: 100, height: 50, backgroundColor: "#10B981" },
            }),
          ),
        ),
      );
      assert.deepEqual(
        (await validatePptx(result.pptx)).errors,
        [],
        item.label,
      );
      assert.equal(result.pages[0]?.png, undefined, item.label);
      assert.equal(result.pages[0]?.svg, undefined, item.label);
      assert.equal(
        result.pages[0]?.diagnostics[0]?.severity,
        "error",
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /box\[bad\]/,
        item.label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        item.expected,
        item.label,
      );
      assert.ok(result.pages[1]?.png !== undefined, item.label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], item.label);
    }
  });
});
