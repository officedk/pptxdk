import assert from "node:assert/strict";

import { build, BuildError, CommonNode } from "@officedk/pptxdk";
import {
  Deck,
  expandFragment,
  Shape,
  Slide,
  Text,
  type ExpandFragmentOptions,
} from "@officedk/pptxdk/react";
import { deck, slide } from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { createElement, Fragment, Suspense, use, useEffect } from "react";
import { describe, it } from "vitest";

describe("public React expandFragment", () => {
  it("expands sync, empty, and multi-node fragments into frozen ownerless CommonNode arrays", async () => {
    const options: ExpandFragmentOptions = {};
    const empty = await expandFragment(null, options);
    assert.equal(Object.isFrozen(empty), true);
    assert.deepEqual(empty, []);

    const nodes = await expandFragment(
      createElement(
        Fragment,
        null,
        createElement(
          Text,
          {
            id: "frag-title",
            style: {
              position: "absolute",
              left: 24,
              top: 24,
              width: 300,
              height: 40,
            },
          },
          "Fragment heading",
        ),
        createElement(Shape, {
          id: "frag-box",
          style: {
            position: "absolute",
            left: 24,
            top: 80,
            width: 200,
            height: 60,
            background: "#E2E8F0",
          },
        }),
      ),
    );

    assert.equal(Object.isFrozen(nodes), true);
    assert.equal(nodes.length, 2);
    for (const node of nodes) {
      assert.ok(node instanceof CommonNode);
      assert.equal(node.parent, undefined);
    }

    const assembled = deck({ width: 640, height: 360 }).append(
      slide({ id: "main" }).append(...nodes),
    );
    const built = await build(assembled);
    assert.deepEqual((await validatePptx(built.pptx)).errors, []);
    assert.deepEqual(
      built.report.slides[0]?.objects.map((o) => o.authorId),
      ["frag-title", "frag-box"],
    );
  });

  it("waits for Suspense async content and supports AbortSignal cancellation", async () => {
    const gate = Promise.withResolvers<string>();
    function AsyncItem() {
      return createElement(
        Text,
        {
          id: "async-item",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 240,
            height: 36,
          },
        },
        use(gate.promise),
      );
    }

    const expanding = expandFragment(
      createElement(
        Suspense,
        {
          fallback: createElement(Text, { id: "loading" }, "Loading"),
        },
        createElement(AsyncItem),
      ),
    );
    gate.resolve("Resolved fragment");
    const resolved = await expanding;
    assert.equal(resolved.length, 1);
    assert.equal(resolved[0]?.id, "async-item");
    assert.equal(resolved[0]?.parent, undefined);

    const controller = new AbortController();
    const reason = new Error("abort fragment");
    const started = Promise.withResolvers<void>();
    const never = new Promise<string>(() => {});
    function Started() {
      useEffect(() => {
        started.resolve();
      }, []);
      return null;
    }
    function Pending() {
      return createElement(Text, { id: "pending" }, use(never));
    }

    const cancelled = expandFragment(
      createElement(
        Fragment,
        null,
        createElement(Started),
        createElement(Suspense, { fallback: null }, createElement(Pending)),
      ),
      { signal: controller.signal },
    );
    await started.promise;
    controller.abort(reason);
    await assert.rejects(cancelled, (error: unknown) => error === reason);
  });

  it("rejects top-level non-whitespace bare text and reports expand phase on render errors", async () => {
    await assert.rejects(
      expandFragment("bare top-level text"),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-structure");
        assert.equal(error.diagnostics[0]?.path, "fragment");
        return true;
      },
    );

    function Broken(): never {
      throw new Error("fragment boom");
    }
    await assert.rejects(
      expandFragment(createElement(Broken)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        const diag = error.diagnostics[0];
        assert.equal(diag?.code, "render-failed");
        if (diag?.code === "render-failed") {
          assert.equal(diag.phase, "expand");
          assert.equal(diag.path, "fragment");
        }
        return true;
      },
    );

    // Whitespace-only text around elements is allowed just like JSX indentation.
    const spaced = await expandFragment(
      createElement(
        Fragment,
        null,
        "\n  ",
        createElement(Deck, { width: 320, height: 180 }, createElement(Slide)),
        "\n",
      ),
    );
    assert.equal(spaced.length, 1);
  });
});
