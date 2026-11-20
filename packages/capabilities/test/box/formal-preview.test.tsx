import assert from "node:assert/strict";
import { createRequire } from "node:module";

import { Box } from "@react-pptx/capabilities/box-elements";
import {
  box,
  deck,
  documentCapabilities,
  shape,
  slide,
} from "@react-pptx/capabilities/document";
import { linearGradient, radialGradient } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import type { CapabilityModule, CommonNode } from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { Deck, Slide, expandDocument } from "@react-pptx/react";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

const require = createRequire(import.meta.url);
const { encode: encodeJpeg } = require("jpeg-js") as typeof import("jpeg-js");

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

  it("renders visible Box containers with linear/radial gradients, image fill, borderRadius, dashed/dotted borders, boxShadow, and softEdge", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const pngImg = new PNG({ width: 4, height: 4 });
    for (let i = 0; i < pngImg.data.length; i += 4) {
      pngImg.data.set([0x25, 0x63, 0xeb, 255], i);
    }
    const validPng = new Uint8Array(PNG.sync.write(pngImg));
    const validGif = new Uint8Array([
      0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0x80, 0, 0, 0x10, 0xb9,
      0x81, 0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0x00, 0x02, 0x02,
      0x44, 0x01, 0x00, 0x3b,
    ]);

    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "box-effects", background: "#FFFFFF" }).append(
        box({
          id: "grad-box",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 100,
            height: 50,
            backgroundImage: linearGradient(90, "#1D4ED8", "#9333EA"),
            borderRadius: 10,
          },
        }),
        box({
          id: "radial-box",
          style: {
            position: "absolute",
            left: 140,
            top: 20,
            width: 100,
            height: 50,
            backgroundImage: radialGradient("#FEF3C7", "#1E3A8A"),
          },
        }),
        box({
          id: "img-png-box",
          style: {
            position: "absolute",
            left: 260,
            top: 20,
            width: 80,
            height: 50,
            backgroundImage: { image: validPng },
            borderRadius: 8,
          },
        }),
        box({
          id: "img-gif-box",
          style: {
            position: "absolute",
            left: 360,
            top: 20,
            width: 80,
            height: 50,
            backgroundImage: { image: validGif },
          },
        }),
        box({
          id: "dashed-shadow-box",
          style: {
            position: "absolute",
            left: 20,
            top: 100,
            width: 140,
            height: 60,
            backgroundColor: "#E2E8F0",
            borderColor: "#0F172A",
            borderWidth: 2,
            borderStyle: "dashed",
            boxShadow: { blur: 4, distance: 6, angle: 90 },
          },
        }),
        box({
          id: "dotted-soft-box",
          style: {
            position: "absolute",
            left: 200,
            top: 100,
            width: 140,
            height: 60,
            backgroundColor: "#0EA5E9",
            borderColor: "#0369A1",
            borderWidth: 2,
            borderStyle: "dotted",
            softEdge: 6,
          },
        }),
      ),
    );

    const result = await previewDoc(doc);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const page = result.pages[0]!;
    assert.deepEqual(page.diagnostics, []);
    assert.ok(page.png !== undefined);
    assert.ok(page.svg !== undefined);
    expect(page.svg).toContain("<linearGradient");
    expect(page.svg).toContain("<radialGradient");
    expect(page.svg).toContain("<pattern");
    expect(page.svg).toContain("<feDropShadow");
    expect(page.svg).toContain("<feGaussianBlur");
    expect(page.svg).toContain("stroke-dasharray");
  });

  it("rejects corrupt PNG, JPEG, and GIF backgroundImage streams on Box with preview-corrupt-image while preserving valid sibling pages and PPTX", async () => {
    const previewDoc = runDocPreview(documentCapabilities);
    const source = new PNG({ width: 20, height: 10 });
    for (let y = 0; y < 10; y += 1) {
      for (let x = 0; x < 20; x += 1) {
        source.data.set(
          x < 10 ? [255, 0, 0, 255] : [0, 0, 255, 255],
          (y * 20 + x) * 4,
        );
      }
    }
    const corruptPng = new Uint8Array(PNG.sync.write(source));
    const idat = Buffer.from(corruptPng).indexOf(Buffer.from("IDAT"));
    corruptPng[idat + 6] = (corruptPng[idat + 6] ?? 0) ^ 0xff;
    corruptPng[idat + 7] = (corruptPng[idat + 7] ?? 0) ^ 0xff;

    const corruptJpeg = new Uint8Array(encodeJpeg(source, 90).data);
    const sos = Buffer.from(corruptJpeg).indexOf(Buffer.from([0xff, 0xda]));
    const sosLen =
      ((corruptJpeg[sos + 2] ?? 0) << 8) + (corruptJpeg[sos + 3] ?? 0);
    for (let i = sos + 2 + sosLen; i < corruptJpeg.length - 4; i++) {
      corruptJpeg[i] = 0x55;
    }

    const corruptGif = new Uint8Array([
      0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 2, 0, 0x80, 0, 0, 0xff, 0xff,
      0xff, 0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 2, 0, 2, 0, 0x00, 0x02, 0x02,
      0xbc, 0x00, 0x00, 0x3b,
    ]);

    for (const [label, bytes] of [
      ["corrupt-png", corruptPng],
      ["corrupt-jpeg", corruptJpeg],
      ["corrupt-gif", corruptGif],
    ] as const) {
      const result = await previewDoc(
        deck({ width: 240, height: 120 }).append(
          slide({ id: "bad-box-page" }).append(
            box({
              id: "bad",
              style: {
                width: 100,
                height: 50,
                backgroundImage: { image: bytes },
              },
            }),
          ),
          slide({ id: "good-box-page" }).append(
            box({
              id: "ok",
              style: { width: 100, height: 50, backgroundColor: "#10B981" },
            }),
          ),
        ),
      );
      assert.deepEqual((await validatePptx(result.pptx)).errors, [], label);
      assert.equal(result.pages[0]?.png, undefined, label);
      assert.equal(result.pages[0]?.svg, undefined, label);
      assert.equal(result.pages[0]?.diagnostics[0]?.severity, "error", label);
      assert.match(
        result.pages[0]?.diagnostics[0]?.path ?? "",
        /box\[bad\]/,
        label,
      );
      assert.match(
        result.pages[0]?.diagnostics[0]?.message ?? "",
        /decode/i,
        label,
      );
      assert.ok(result.pages[1]?.png !== undefined, label);
      assert.deepEqual(result.pages[1]?.diagnostics, [], label);
    }
  });
});
