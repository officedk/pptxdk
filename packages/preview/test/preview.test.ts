import assert from "node:assert/strict";
import { createRequire } from "node:module";

import { createResolver } from "@pptxdk/core/build";
import {
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  type ObjectCapabilityModule,
} from "@pptxdk/core/capability";
import {
  CommonNode,
  createDeckNode,
  createSlideNode,
} from "@pptxdk/core/document";
import {
  PreviewPageRangeError,
  deduplicateSvgRasterImages,
  paint,
  paintOptionsSchema,
  previewPageNumberSchema,
  sanitizeSvgForRasterization,
  screenGeometry,
  validatePaintOptions,
} from "@pptxdk/preview";
import { describe, it } from "vitest";

const unsupportedWidgetCapability: ObjectCapabilityModule = {
  declaration: "custom-widget",
  bind: (input, context) => ({
    contract: parseCapabilityLayoutContract({
      nativeKind: "sp",
      props: input.props,
      allowedStyleKeys: new Set(layoutStyleKeys),
      path: context.path,
    }),
    resolve: (ctx) => ({
      observation: {
        id: 1,
        kind: "custom-widget",
        name: "custom-widget",
        ...(ctx.layoutBox === undefined ? {} : { box: ctx.layoutBox }),
      },
      extensions: [],
      emit: () => ({ xml: "" }),
    }),
  }),
};

describe("@pptxdk/preview pure unit contract (without rasterizer)", () => {
  it("validates paint options and rejects non-positive page numbers, unknown keys, and bad signals", () => {
    assert.deepEqual(validatePaintOptions(), {});
    assert.deepEqual(validatePaintOptions({ pages: [1, 3] }), {
      pages: [1, 3],
    });
    assert.equal(paintOptionsSchema.safeParse({ pages: [] }).success, true);
    assert.equal(previewPageNumberSchema.safeParse(1).success, true);
    assert.equal(previewPageNumberSchema.safeParse(0).success, false);

    assert.throws(() => validatePaintOptions({ pages: [0] }), RangeError);
    assert.throws(() => validatePaintOptions({ pages: [1.5] }), RangeError);
    assert.throws(() => validatePaintOptions({ signal: "bad" }), TypeError);
    assert.throws(() => validatePaintOptions({ unknown: true }), TypeError);
    assert.throws(() => validatePaintOptions(null), TypeError);
  });

  it("computes rotated and flipped screen geometry deterministically", () => {
    const rect = { x: 10, y: 20, width: 100, height: 40 };
    const identity = screenGeometry(rect, []);
    assert.deepEqual(identity.screenBox, rect);
    assert.deepEqual(identity.corners, [
      { x: 10, y: 20 },
      { x: 110, y: 20 },
      { x: 110, y: 60 },
      { x: 10, y: 60 },
    ]);

    const rotated = screenGeometry(rect, [
      { box: rect, angle: 90, flipH: false, flipV: false },
    ]);
    assert.deepEqual(rotated.screenBox, {
      x: 40,
      y: -10,
      width: 40,
      height: 100,
    });
  });

  it("handles zero-page paint, out-of-range pages, aborted signals, and unsupported capability diagnostics without loading resvg", () => {
    const resolve = createResolver([unsupportedWidgetCapability]);
    const doc = createDeckNode({ width: 480, height: 270 }).append(
      createSlideNode({ id: "s1" }).append(
        new CommonNode("custom-widget", "custom-widget", {
          id: "w1",
          style: {
            position: "absolute",
            left: 10,
            top: 20,
            width: 100,
            height: 40,
          },
        }),
      ),
    );
    const resolution = resolve(doc);

    const zeroPages = paint(resolution, { pages: [] });
    assert.deepEqual(zeroPages.pages, []);

    assert.throws(
      () => paint(resolution, { pages: [2] }),
      (error: unknown) =>
        error instanceof PreviewPageRangeError &&
        error.totalSlides === 1 &&
        error.requestedPages.includes(2),
    );

    const controller = new AbortController();
    const reason = new Error("aborted paint");
    controller.abort(reason);
    assert.throws(
      () => paint(resolution, { signal: controller.signal }),
      (error: unknown) => error === reason,
    );

    // Unsupported capability short-circuits before calling Resvg rasterization.
    const errored = paint(resolution);
    assert.equal(errored.pages.length, 1);
    assert.equal(errored.pages[0]?.png, undefined);
    assert.equal(errored.pages[0]?.svg, undefined);
    assert.equal(
      errored.pages[0]?.errors[0]?.code,
      "preview-unsupported-capability",
    );
    assert.equal(errored.pages[0]?.errors[0]?.capability, "custom-widget");
  });

  it("sanitizes zero-area clips, degenerate gradients, non-finite geometry, and extreme off-canvas bounds before rasterization", () => {
    const dimensions = { width: 640, height: 360 };

    // 1. Zero-width, negative-width, sub-epsilon, and empty clipPaths strip referencing elements cleanly.
    const zeroClipSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360"><defs><clipPath id="c-zero"><rect x="10" y="10" width="0" height="50"/></clipPath><clipPath id="c-line"><path d="M10 10 L100 10"/></clipPath><clipPath id="c-empty"/></defs><rect x="0" y="0" width="640" height="360" fill="white"/><rect clip-path="url(#c-zero)" x="0" y="0" width="50" height="50" fill="#DC2626"/><path clip-path="url(#c-line)" d="M0 0 H100 V100 H0 Z" fill="#2563EB"/><g clip-path="url(#c-empty)"><rect x="0" y="0" width="20" height="20" fill="#16A34A"/></g></svg>`;
    const sanitizedZero = sanitizeSvgForRasterization(zeroClipSvg, dimensions);
    assert.equal(sanitizedZero.includes("c-zero"), false);
    assert.equal(sanitizedZero.includes("c-line"), false);
    assert.equal(sanitizedZero.includes("c-empty"), false);
    assert.equal(sanitizedZero.includes("#DC2626"), false);
    assert.equal(sanitizedZero.includes("#2563EB"), false);
    assert.equal(sanitizedZero.includes("#16A34A"), false);
    assert.equal(sanitizedZero.includes('fill="white"'), true);

    // 2. Degenerate linearGradient (zero-length vector) & radialGradient (r<=0) fall back to solid last-stop fill.
    const gradSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360"><defs><linearGradient id="g-zero" gradientUnits="objectBoundingBox" x1="0.5" y1="0.5" x2="0.5" y2="0.5"><stop offset="0%" stop-color="#111111" stop-opacity="1"/><stop offset="100%" stop-color="#047857" stop-opacity="0.75"/></linearGradient><radialGradient id="g-rad0" cx="0.5" cy="0.5" r="0"><stop offset="0%" stop-color="#000000"/><stop offset="100%" stop-color="#EA580C" stop-opacity="0.5"/></radialGradient></defs><rect x="10" y="10" width="100" height="50" fill="url(#g-zero)"/><rect x="120" y="10" width="100" height="50" fill="url(#g-rad0)"/></svg>`;
    const sanitizedGrad = sanitizeSvgForRasterization(gradSvg, dimensions);
    assert.equal(sanitizedGrad.includes("<linearGradient"), false);
    assert.equal(sanitizedGrad.includes("<radialGradient"), false);
    assert.equal(sanitizedGrad.includes('fill="#047857"'), true);
    assert.equal(sanitizedGrad.includes('fill="#EA580C"'), true);

    // 3. Off-canvas clipPaths, extreme translated clips, and non-finite geometry are dropped without touching valid content.
    const offscreenSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360"><defs><clipPath id="c-off"><rect x="10000" y="0" width="100" height="100"/></clipPath><clipPath id="c-local"><rect x="0" y="0" width="100" height="100"/></clipPath><filter id="f-bad" filterUnits="userSpaceOnUse" x="NaN" y="0" width="10" height="10"><feDropShadow dx="1" dy="1" stdDeviation="1"/></filter></defs><rect clip-path="url(#c-off)" x="10000" y="0" width="100" height="100" fill="#DC2626"/><g transform="translate(100000000 0)"><rect clip-path="url(#c-local)" x="0" y="0" width="100" height="100" fill="#B91C1C"/></g><rect x="NaN" y="10" width="50" height="50" fill="#991B1B"/><rect filter="url(#f-bad)" x="20" y="20" width="80" height="40" fill="#1D4ED8"/></svg>`;
    const sanitizedOffscreen = sanitizeSvgForRasterization(
      offscreenSvg,
      dimensions,
    );
    assert.equal(sanitizedOffscreen.includes("c-off"), false);
    assert.equal(sanitizedOffscreen.includes("100000000"), false);
    assert.equal(sanitizedOffscreen.includes("NaN"), false);
    assert.equal(sanitizedOffscreen.includes("f-bad"), false);
    assert.equal(sanitizedOffscreen.includes("#1D4ED8"), true);

    // 4. Viewport-spanning parallax clipPath/pattern with extreme horizontal extent is clamped to safe neighborhood.
    const parallaxSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360"><defs><clipPath id="c-para" clipPathUnits="userSpaceOnUse"><rect x="-500000" y="0" width="1000000" height="360"/></clipPath><pattern id="p-para" patternUnits="userSpaceOnUse" x="-500000" y="0" width="1000000" height="360"><rect x="0" y="0" width="1000000" height="360" fill="#0EA5E9"/></pattern></defs><rect clip-path="url(#c-para)" x="-500000" y="0" width="1000000" height="360" fill="url(#p-para)"/></svg>`;
    const sanitizedParallax = sanitizeSvgForRasterization(
      parallaxSvg,
      dimensions,
    );
    assert.equal(sanitizedParallax.includes("c-para"), true);
    assert.equal(sanitizedParallax.includes("p-para"), true);
    assert.equal(sanitizedParallax.includes("1000000"), false);
    assert.equal(sanitizedParallax.includes("-500000"), false);

    // 5. Valid rotated/flipped shapes and normal gradients/clips remain intact.
    const validSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360"><defs><linearGradient id="g-ok" gradientUnits="objectBoundingBox" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#1D4ED8"/><stop offset="100%" stop-color="#0EA5E9"/></linearGradient></defs><g transform="translate(100 100) scale(-1 1)"><path d="M0 0 L80 0 L80 40 L0 40 Z" fill="url(#g-ok)"/></g></svg>`;
    assert.equal(sanitizeSvgForRasterization(validSvg, dimensions), validSvg);
  });

  it("deduplicates repeated large raster data URIs into shared <defs> and leaves small or invalid images untouched without loading resvg", () => {
    const largeDataUri = `data:image/png;base64,${"A".repeat(5000)}`;
    const smallDataUri = "data:image/png;base64,AAAA";
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180" viewBox="0 0 320 180"><rect width="100%" height="100%" fill="white"/><image clip-path="url(#c1)" x="10" y="20" width="40" height="30" preserveAspectRatio="none" href="${largeDataUri}"/><image x="60" y="20" width="40" height="30" preserveAspectRatio="none" href="${largeDataUri}"/><image x="110" y="20" width="20" height="20" href="${smallDataUri}"/><image x="140" y="20" width="0" height="20" href="${largeDataUri}"/><image x="170" y="20" width="20" height="20" href="https://example.com/a.png"/></svg>`;

    const deduplicated = deduplicateSvgRasterImages(svg, "shared-img-s1");
    assert.equal(
      (deduplicated.match(/<image\b/gu) ?? []).length,
      4,
      "1 shared <defs><image> + 3 non-hoisted <image> elements",
    );
    assert.equal((deduplicated.match(/<use\b/gu) ?? []).length, 2);
    assert.ok(
      deduplicated.includes(
        `<defs><image id="shared-img-s1-1" x="0" y="0" width="1" height="1" preserveAspectRatio="none" href="${largeDataUri}"/></defs>`,
      ),
    );
    assert.ok(
      deduplicated.includes(
        '<g clip-path="url(#c1)"><use href="#shared-img-s1-1" transform="translate(10 20) scale(40 30)"/></g>',
      ),
    );
    assert.ok(
      deduplicated.includes(
        '<use href="#shared-img-s1-1" transform="translate(60 20) scale(40 30)"/>',
      ),
    );
    assert.ok(deduplicated.includes(`href="${smallDataUri}"`));
    assert.ok(deduplicated.includes('width="0" height="20"'));
    assert.ok(deduplicated.includes('href="https://example.com/a.png"'));

    // Negative / no-op cases: SVG without hoistable data URIs is returned unchanged.
    const unchangedSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><image x="0" y="0" width="10" height="10" href="${smallDataUri}"/></svg>`;
    assert.equal(deduplicateSvgRasterImages(unchangedSvg), unchangedSvg);

    // @pptxdk/preview#test must never load @resvg/resvg-js.
    const req = createRequire(import.meta.url);
    assert.equal(
      Object.keys(req.cache).some((id) => id.includes("resvg")),
      false,
      "@pptxdk/preview unit tests must not load @resvg/resvg-js",
    );
  });
});
