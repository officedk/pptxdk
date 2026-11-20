import assert from "node:assert/strict";

import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { expandDocument } from "@react-pptx/react";
import {
  BuildError,
  Deck,
  Slide,
  linearGradient,
  slideBackgroundCapability,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const png1x1 = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
    "base64",
  ),
);

describe("slide capabilities resolved contract", () => {
  it("migrates slideBackgroundCapability to BoundSlideCapability with paint-only solid/default rendering, preview error on image/gradient, and emit-only resource registration", async () => {
    const resolve = createResolver([slideBackgroundCapability]);
    const doc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "default-bg" }),
        createElement(Slide, { id: "solid-bg", background: "#1E293B" }),
        createElement(Slide, {
          id: "gradient-bg",
          background: linearGradient(
            90,
            { color: "#0F172A", position: 0 },
            { color: "#38BDF8", position: 1 },
          ),
        }),
        createElement(Slide, {
          id: "image-bg",
          background: { image: png1x1 },
        }),
      ),
    );

    const resolution = resolve(doc);

    // Paint directly from Resolution without running emit().
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 4);

    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);
    assert.match(painted.pages[0]?.svg ?? "", /fill="white"/u);

    assert.deepEqual(painted.pages[1]?.diagnostics, []);
    assert.ok(painted.pages[1]?.png !== undefined);
    assert.match(painted.pages[1]?.svg ?? "", /fill="#1E293B"/u);

    // Gradient and image backgrounds report preview error diagnostics and omit PNG,
    // while emit() on the same Resolution succeeds.
    for (const pageIndex of [2, 3]) {
      const page = painted.pages[pageIndex]!;
      assert.equal(page.png, undefined);
      assert.ok(
        page.diagnostics.some(
          (diag) =>
            diag.severity === "error" &&
            diag.message.includes("image or gradient backgrounds"),
        ),
      );
    }

    const emitted1 = emit(resolution);
    const emitted2 = emit(resolution);
    assert.deepEqual((await validatePptx(emitted1.pptx)).errors, []);
    assert.deepEqual((await validatePptx(emitted2.pptx)).errors, []);

    const entries1 = unzipSync(emitted1.pptx);
    const entries2 = unzipSync(emitted2.pptx);
    const media1 = Object.keys(entries1).filter((name) =>
      name.startsWith("ppt/media/"),
    );
    const media2 = Object.keys(entries2).filter((name) =>
      name.startsWith("ppt/media/"),
    );
    assert.equal(media1.length, 1);
    assert.deepEqual(media2, media1);

    const slide4Xml = strFromU8(entries1["ppt/slides/slide4.xml"]!);
    assert.match(slide4Xml, /<p:bg><p:bgPr><a:blipFill>/u);
    const slide4Rels = strFromU8(entries1["ppt/slides/_rels/slide4.xml.rels"]!);
    assert.match(slide4Rels, /relationships\/image/u);

    // Invalid background fails during resolve.
    const invalidDoc = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, {
          id: "bad-bg",
          background: { image: png1x1, extra: true } as never,
        }),
      ),
    );
    assert.throws(
      () => resolve(invalidDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) =>
            diag.code === "invalid-slide-background" &&
            diag.path === "deck/slide[bad-bg]/background",
        ),
    );
  });
});
