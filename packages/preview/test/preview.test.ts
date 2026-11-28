import assert from "node:assert/strict";

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
  paint,
  paintOptionsSchema,
  previewPageNumberSchema,
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
});
