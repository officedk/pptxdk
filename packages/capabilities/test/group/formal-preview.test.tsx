import assert from "node:assert/strict";

import {
  box,
  deck,
  documentCapabilities,
  group,
  image,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import { Group } from "@react-pptx/capabilities/group-elements";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import {
  BuildError,
  type CapabilityModule,
  type CommonNode,
} from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
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

describe("Slice 5 (Ticket 08): Group static preview, screenGeometry, and member failure propagation", () => {
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
});
