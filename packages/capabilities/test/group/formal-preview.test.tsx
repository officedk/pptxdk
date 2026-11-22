import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  box,
  deck,
  font,
  group,
  image,
  paragraph,
  shape,
  slide,
  span,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import {
  BuildError,
  type CapabilityModule,
  type CommonNode,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, Slide, expandDocument } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

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

describe("Group static preview, screenGeometry, and member failure propagation", () => {
  it("renders nested rotated Group trees with flipped members, Box children, screenGeometry, and post-preview isolation", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const mutableShape = shape({
      id: "inner-leaf",
      x: 110,
      y: 70,
      width: 60,
      height: 20,
      fill: "#1D4ED8",
      flipH: true,
    });
    const doc = deck({
      width: 400,
      height: 240,
      theme: { colors: { groupAccent: "#111111" } },
    }).append(
      slide({ id: "groups-preview", background: "#FFFFFF" }).append(
        group({
          id: "outer",
          theme: { colors: { groupAccent: "#0F766E" } },
          style: { rotate: 90 },
        }).append(
          shape({
            id: "base-rect",
            x: 70,
            y: 40,
            width: 80,
            height: 40,
            fill: "groupAccent",
          }),
          box({
            id: "group-box",
            style: {
              position: "absolute",
              left: 150,
              top: 40,
              width: 60,
              height: 40,
              padding: 4,
              backgroundColor: "#E2E8F0",
            },
          }).append(
            shape({
              id: "box-child",
              width: 30,
              height: 20,
              fill: "#F59E0B",
            }),
          ),
          group({
            id: "inner",
            style: { rotate: -90 },
          }).append(mutableShape),
        ),
      ),
    );

    const first = await previewDoc(doc);
    const page = first.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain('transform="rotate(90 ');
    expect(page.svg).toContain('transform="rotate(-90 ');
    expect(page.svg).toContain('fill="#0F766E"');
    expect(page.svg).toContain('fill="#1D4ED8"');

    const outerNode = page.nodes.find((n) => n.authorId === "outer")!;
    // Outer bounds: x=[70..210] (width 140), y=[40..90] (height 50), center=(140, 65).
    // Rotated 90deg -> width=50, height=140, x=115, y=-5.
    assert.deepEqual(outerNode.box, { x: 70, y: 40, width: 140, height: 50 });
    expect(outerNode.screenBox?.x).toBeCloseTo(115);
    expect(outerNode.screenBox?.y).toBeCloseTo(-5);
    expect(outerNode.screenBox?.width).toBeCloseTo(50);
    expect(outerNode.screenBox?.height).toBeCloseTo(140);
    assert.equal(outerNode.corners?.length, 4);

    // Inner group is rotated -90 around (140, 80), then +90 around (140, 65) => net 0deg rotation!
    // Center (140, 80) rotated +90deg around (140, 65) goes to (125, 65).
    // Since net rotation on inner-leaf is 0deg, its screenBox has width=60, height=20 centered at (125, 65).
    const innerLeafNode = page.nodes.find((n) => n.authorId === "inner-leaf")!;
    expect(innerLeafNode.screenBox?.x).toBeCloseTo(95);
    expect(innerLeafNode.screenBox?.y).toBeCloseTo(55);
    expect(innerLeafNode.screenBox?.width).toBeCloseTo(60);
    expect(innerLeafNode.screenBox?.height).toBeCloseTo(20);

    // Mutate after preview: prior outputs remain unchanged.
    const firstSvg = page.svg;
    const firstPng = page.png.slice();
    const firstPptx = first.pptx.slice();
    mutableShape.set({ fill: "#DC2626" });
    const second = await previewDoc(doc);
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.pptx, firstPptx);
    expect(second.pages[0]?.svg).toContain('fill="#DC2626"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // JSX entry renders equivalent rotated Group.
    const jsxDoc = await expandDocument(
      createElement(
        Deck,
        { width: 240, height: 160 },
        createElement(
          Slide,
          { id: "jsx-group", background: "#FFFFFF" },
          createElement(
            Group,
            { id: "jg", style: { rotate: 90 } },
            createElement(Shape, {
              id: "js",
              x: 40,
              y: 30,
              width: 80,
              height: 40,
              fill: "#2563EB",
            }),
          ),
        ),
      ),
    );
    const jsxResult = await previewDoc(jsxDoc);
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    const jsxGroupNode = jsxResult.pages[0]?.nodes.find(
      (n) => n.authorId === "jg",
    );
    expect(jsxGroupNode?.screenBox?.width).toBeCloseTo(40);
    expect(jsxGroupNode?.screenBox?.height).toBeCloseTo(80);
  });

  it("propagates member preview failures out of Group without omitting failed members or breaking sibling pages", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const png = new PNG({ width: 4, height: 4 });
    png.data.fill(255);
    const corruptPng = new Uint8Array(PNG.sync.write(png));
    const idat = Buffer.from(corruptPng).indexOf(Buffer.from("IDAT"));
    corruptPng[idat + 6] = (corruptPng[idat + 6] ?? 0) ^ 0xff;
    corruptPng[idat + 7] = (corruptPng[idat + 7] ?? 0) ^ 0xff;

    const result = await previewDoc(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "group-with-bad-member" }).append(
          group({ id: "g", style: { rotate: 15 } }).append(
            shape({
              id: "ok-member",
              x: 20,
              y: 20,
              width: 50,
              height: 30,
              fill: "#10B981",
            }),
            image({
              id: "bad-member",
              src: corruptPng,
              x: 80,
              y: 20,
              width: 50,
              height: 30,
            }),
          ),
        ),
        slide({ id: "valid-sibling" }).append(
          group({ id: "ok-group" }).append(
            shape({
              id: "ok-shape",
              x: 20,
              y: 20,
              width: 60,
              height: 40,
              fill: "#2563EB",
            }),
          ),
        ),
      ),
    );

    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const badPage = result.pages[0]!;
    assert.equal(badPage.png, undefined);
    assert.equal(badPage.svg, undefined);
    assert.ok(
      badPage.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /group\[g\]\/image\[bad-member\]/u.test(d.path) &&
          /decode/i.test(d.message),
      ),
    );
    // Nodes and transformed screenBoxes are still reported for locator inspection.
    assert.ok(
      badPage.nodes.find((n) => n.authorId === "bad-member")?.screenBox !==
        undefined,
    );

    const goodPage = result.pages[1]!;
    assert.deepEqual(goodPage.diagnostics, []);
    assert.ok(goodPage.png !== undefined);

    // Non-finite rotation on Group fails formal build validation.
    await assert.rejects(
      () =>
        previewDoc(
          deck({ width: 320, height: 180 }).append(
            slide({ id: "bad-rot" }).append(
              group({ id: "g", style: { rotate: Number.NaN } }).append(
                shape({ id: "s", x: 10, y: 10, width: 20, height: 20 }),
              ),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "invalid-visual-style"),
        );
        return true;
      },
    );
  });

  it("composites 3-level nested rotated Group with inline-text flipped/opaque Shape, local ThemeScope, exact corners, and missing-glyph propagation", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const doc = deck({
      width: 400,
      height: 260,
      theme: { colors: { cardFill: "#991B1B", ink: "#000000" } },
    }).append(
      slide({ id: "deep-group", background: "#FFFFFF" }).append(
        group({
          id: "lvl1",
          theme: { colors: { cardFill: "#1D4ED8", ink: "#FFFFFF" } },
          style: { rotate: 90 },
        }).append(
          shape({
            id: "anchor",
            x: 100,
            y: 50,
            width: 40,
            height: 40,
            fill: "#E2E8F0",
          }),
          group({
            id: "lvl2",
            style: { rotate: -90 },
          }).append(
            group({
              id: "lvl3",
              style: { rotate: 45 },
            }).append(
              shape({
                id: "deep-rich-shape",
                geometry: "roundRect",
                x: 160,
                y: 90,
                width: 120,
                height: 60,
                fill: "cardFill",
                flipH: true,
                style: { opacity: 0.7 },
              }).append(
                paragraph([span("Deep")]).style({
                  fontFamily: jostFont,
                  fontSize: 16,
                  color: "ink",
                  align: "center",
                }),
              ),
            ),
          ),
        ),
      ),
      slide({ id: "deep-glyph-fail" }).append(
        group({ id: "outer-fail" }).append(
          group({ id: "inner-fail", style: { rotate: 15 } }).append(
            shape({
              id: "bad-shape-text",
              x: 40,
              y: 40,
              width: 100,
              height: 40,
              fill: "#111827",
              text: "汉",
              textStyle: { fontFamily: jostFont, color: "#FFFFFF" },
            }),
          ),
        ),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);

    const p1 = result.pages[0]!;
    assert.deepEqual(p1.diagnostics, []);
    assert.ok(p1.png !== undefined);
    assert.ok(p1.svg !== undefined);
    expect(p1.svg).toContain("scale(-1 1)");

    // Analytical corners check for deep-rich-shape:
    // lvl1 box: anchor [100..140, 50..90] + lvl3 [160..280, 90..150] => x=[100..280] (w=180, cx=190), y=[50..150] (h=100, cy=100).
    // lvl2 box: [160..280, 90..150] (cx=220, cy=120).
    // lvl3 box: [160..280, 90..150] (cx=220, cy=120).
    // Net rotation of deep-rich-shape: 45 - 90 + 90 = +45deg, around center:
    // (220, 120) rotated -90 around (220, 120) -> (220, 120); rotated +90 around (190, 100) -> (170, 130).
    // So deep-rich-shape screenBox is centered at (170, 130) with rotated 45deg bounding box side = (120 + 60) / sqrt(2).
    const deepNode = p1.nodes.find((n) => n.authorId === "deep-rich-shape")!;
    assert.equal(deepNode.corners?.length, 4);
    const expectedSpan = (120 + 60) / Math.SQRT2;
    expect(deepNode.screenBox?.width).toBeCloseTo(expectedSpan, 3);
    expect(deepNode.screenBox?.height).toBeCloseTo(expectedSpan, 3);
    expect(deepNode.screenBox!.x + deepNode.screenBox!.width / 2).toBeCloseTo(
      170,
      3,
    );
    expect(deepNode.screenBox!.y + deepNode.screenBox!.height / 2).toBeCloseTo(
      130,
      3,
    );

    const p2 = result.pages[1]!;
    assert.equal(p2.png, undefined);
    assert.equal(p2.svg, undefined);
    assert.ok(
      p2.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /group\[outer-fail\]\/group\[inner-fail\]\/shape\[bad-shape-text\]/u.test(
            d.path,
          ) &&
          /glyph/i.test(d.message),
      ),
    );
  });
});
