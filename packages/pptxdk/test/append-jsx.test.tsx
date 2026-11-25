import assert from "node:assert/strict";
import { getEventListeners } from "node:events";

import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  after,
  build,
  BuildError,
  CommonContainer,
  CommonNode,
  createMorphScope,
  createMotionScopeHandle,
  defineTheme,
  fade,
  RichTextContainer,
  type MotionObservation,
} from "pptxdk";
import {
  appendJSX,
  Box,
  Deck,
  expandDocument,
  MotionScope,
  Shape,
  Slide,
  Step,
  Text,
  useMotionScope,
} from "pptxdk/react";
import { box, deck, shape, slide, text } from "pptxdk/vanilla";
import {
  Component,
  createContext,
  createElement,
  Fragment,
  Suspense,
  use,
  useContext,
  useEffect,
  useLayoutEffect,
  type ReactNode,
} from "react";
import { describe, it } from "vitest";

function slideXml(result: { readonly pptx: Uint8Array }, index: number) {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function container(node: CommonNode) {
  assert.ok(node instanceof CommonContainer);
  return node;
}

function findNode(parent: CommonNode, kind: string): CommonNode | undefined {
  for (const child of parent.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.kind === kind) return child;
    const found = findNode(child, kind);
    if (found !== undefined) return found;
  }
  return undefined;
}

describe("public React fragment append", () => {
  it("appends a complete JSX fragment after cleanup and keeps the nodes editable", async () => {
    const morph = createMorphScope();
    const hero = morph.handle("hero");
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "morph-source" },
          createElement(Shape, {
            id: "anchor",
            x: 0,
            y: 0,
            width: 20,
            height: 20,
            morphKey: hero,
          }),
        ),
        createElement(
          Slide,
          { id: "append-target" },
          createElement(
            Text,
            { id: "before", x: 0, y: 0, width: 160, height: 24 },
            "Before",
          ),
        ),
      ),
    );
    const target = container(expanded.query.get("append-target").node);

    const appended = await appendJSX(
      target,
      createElement(
        Fragment,
        null,
        createElement(
          Text,
          {
            id: "title",
            x: 0,
            y: 24,
            width: 200,
            height: 32,
            style: { color: "#123456" },
          },
          "Appended title",
        ),
        createElement(Shape, {
          id: "hero-copy",
          x: 0,
          y: 60,
          width: 40,
          height: 30,
          morphKey: hero,
        }),
      ),
    );

    assert.deepEqual(
      appended.map((node) => node.id),
      ["title", "hero-copy"],
    );
    const title = target.query.get("title").node;
    assert.ok(title instanceof RichTextContainer);
    assert.equal(title, appended[0]);
    title.content("Edited appended title");
    assert.equal(target.query.get("hero-copy").node.props["morphKey"], hero);

    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[1]?.objects.map(({ authorId, text, box }) => ({
        authorId,
        text,
        box,
      })),
      [
        {
          authorId: "before",
          text: "Before",
          box: { x: 0, y: 0, width: 160, height: 24 },
        },
        {
          authorId: "title",
          text: "Edited appended title",
          box: { x: 0, y: 24, width: 200, height: 32 },
        },
        {
          authorId: "hero-copy",
          text: undefined,
          box: { x: 0, y: 60, width: 40, height: 30 },
        },
      ],
    );
    assert.equal(
      result.report.slides[0]?.objects[0]?.name,
      result.report.slides[1]?.objects[2]?.name,
    );
    assert.match(slideXml(result, 2), /Edited appended title/u);
    assert.match(slideXml(result, 2), /123456/u);
  });

  it("waits for final suspended fragment content before attaching", async () => {
    const final = Promise.withResolvers<string>();
    const started = Promise.withResolvers<void>();
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "suspense-target" }),
      ),
    );
    const target = container(expanded.query.get("suspense-target").node);

    function Started() {
      useEffect(() => {
        started.resolve();
      }, []);
      return null;
    }
    function LateContent() {
      return createElement(
        Text,
        { id: "late", x: 0, y: 0, width: 160, height: 24 },
        use(final.promise),
      );
    }

    const appending = appendJSX(
      target,
      createElement(
        Fragment,
        null,
        createElement(Started),
        createElement(
          Suspense,
          {
            fallback: createElement(
              Text,
              { id: "fallback", x: 0, y: 0, width: 160, height: 24 },
              "Loading",
            ),
          },
          createElement(LateContent),
        ),
      ),
    );
    await started.promise;
    final.resolve("Final content");
    const appended = await appending;

    assert.deepEqual(
      appended.map((node) => node.id),
      ["late"],
    );
    assert.equal(target.query.getAll("fallback").length, 0);
    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Final content"],
    );
  });

  it("appends the final content of handled business exceptions", async () => {
    class Boundary extends Component<
      { readonly children: ReactNode },
      { readonly failed: boolean }
    > {
      override state = { failed: false };

      static getDerivedStateFromError() {
        return { failed: true };
      }

      override render() {
        return this.state.failed
          ? createElement(
              Text,
              { id: "business-fallback", x: 0, y: 0, width: 200, height: 24 },
              "Handled fallback",
            )
          : this.props.children;
      }
    }
    function ThrowingContent(): never {
      throw new Error("handled business failure");
    }
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "business-target" }),
      ),
    );
    const target = container(expanded.query.get("business-target").node);

    const appended = await appendJSX(
      target,
      createElement(Boundary, null, createElement(ThrowingContent)),
    );
    assert.deepEqual(
      appended.map((node) => node.id),
      ["business-fallback"],
    );
    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Handled fallback"],
    );
  });

  it("attaches nothing until unmount cleanup has completed", async () => {
    let cleanupRan = false;
    let fragmentVisibleDuringCleanup: boolean | undefined;
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "cleanup-target" }),
      ),
    );
    const target = container(expanded.query.get("cleanup-target").node);

    function CleanupProbe() {
      useLayoutEffect(
        () => () => {
          cleanupRan = true;
          fragmentVisibleDuringCleanup =
            target.query.getAll("cleanup-probe").length > 0;
        },
        [],
      );
      return createElement(
        Text,
        { id: "cleanup-probe", x: 0, y: 0, width: 160, height: 24 },
        "After cleanup",
      );
    }

    const appended = await appendJSX(target, createElement(CleanupProbe));
    assert.equal(cleanupRan, true);
    assert.equal(fragmentVisibleDuringCleanup, false);
    assert.equal(appended.length, 1);
    assert.equal(target.query.get("cleanup-probe").node, appended[0]);
  });

  it("provides fragment-local React context without reviving the finished root", async () => {
    const Label = createContext("default");
    function Labeled({ id }: { readonly id: string }) {
      return createElement(
        Text,
        { id, x: 0, y: 0, width: 160, height: 24 },
        useContext(Label),
      );
    }

    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "context-target" },
          createElement(
            Label.Provider,
            { value: "old root" },
            createElement(Labeled, { id: "old-root-context" }),
          ),
        ),
      ),
    );
    const target = container(expanded.query.get("context-target").node);

    await appendJSX(target, createElement(Labeled, { id: "default-context" }));
    await appendJSX(
      target,
      createElement(
        Label.Provider,
        { value: "fragment provider" },
        createElement(Labeled, { id: "provider-context" }),
      ),
    );

    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["old root", "default", "fragment provider"],
    );
  });

  it("creates new node instances for repeated appends with the same React key", async () => {
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "repeat-target" }),
      ),
    );
    const target = container(expanded.query.get("repeat-target").node);

    function Repeated({
      scope,
      label,
    }: {
      readonly scope: string;
      readonly label: string;
    }) {
      return createElement(
        Box,
        { id: scope },
        createElement(
          Text,
          { key: "same-key", id: "title", x: 0, y: 0, width: 160, height: 24 },
          label,
        ),
      );
    }

    const first = await appendJSX(
      target,
      createElement(Repeated, { scope: "scope-a", label: "First" }),
    );
    const second = await appendJSX(
      target,
      createElement(Repeated, { scope: "scope-b", label: "Second" }),
    );
    const firstTitle = target.query.get("scope-a").get("title").node;
    const secondTitle = target.query.get("scope-b").get("title").node;
    assert.notEqual(firstTitle, secondTitle);
    assert.equal(firstTitle, first[0]?.children[0]);
    assert.equal(secondTitle, second[0]?.children[0]);

    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ authorId, text: value }) => ({
        authorId,
        text: value,
      })),
      [
        { authorId: "title", text: "First" },
        { authorId: "title", text: "Second" },
      ],
    );
  });

  it("rejects a naming conflict against the current target structure without residue", async () => {
    const gate = Promise.withResolvers<void>();
    const started = Promise.withResolvers<void>();
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "conflict-target" }),
      ),
    );
    const target = container(expanded.query.get("conflict-target").node);

    function Started() {
      useEffect(() => {
        started.resolve();
      }, []);
      return null;
    }
    function LateFragment() {
      use(gate.promise);
      return createElement(
        Text,
        { id: "late", x: 0, y: 0, width: 160, height: 24 },
        "Appended",
      );
    }

    const appending = appendJSX(
      target,
      createElement(
        Fragment,
        null,
        createElement(Started),
        createElement(
          Suspense,
          { fallback: null },
          createElement(
            Box,
            null,
            createElement(LateFragment),
            createElement(
              Text,
              { id: "fragment-only", x: 0, y: 60, width: 160, height: 24 },
              "Fragment only",
            ),
          ),
        ),
      ),
    );
    await started.promise;
    const external = text("External edit", {
      id: "late",
      x: 0,
      y: 30,
      width: 160,
      height: 24,
    });
    target.append(external);

    gate.resolve();
    await assert.rejects(appending, (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "duplicate-author-id");
      return true;
    });
    assert.equal(target.query.get("late").node, external);
    assert.equal(target.query.getAll("late").length, 1);
    assert.equal(target.query.getAll("fragment-only").length, 0);
  });

  it("checks the full anonymous parent scope before appending", async () => {
    const gate = Promise.withResolvers<void>();
    const started = Promise.withResolvers<void>();
    const root = box();
    const wrapper = box();
    const target = box();
    root.append(wrapper);
    wrapper.append(
      shape({ id: "duplicate", x: 0, y: 0, width: 10, height: 10 }),
      target,
    );
    const before = [...root.children];

    function Started() {
      useEffect(() => {
        started.resolve();
      }, []);
      return null;
    }
    function LateDuplicate() {
      use(gate.promise);
      return createElement(Shape, {
        id: "duplicate",
        x: 10,
        y: 0,
        width: 10,
        height: 10,
      });
    }

    const appending = appendJSX(
      target,
      createElement(
        Fragment,
        null,
        createElement(Started),
        createElement(
          Suspense,
          { fallback: null },
          createElement(Box, null, createElement(LateDuplicate)),
        ),
      ),
    );
    await started.promise;
    const external = shape({
      id: "external",
      x: 0,
      y: 20,
      width: 10,
      height: 10,
    });
    root.append(external);

    gate.resolve();
    await assert.rejects(appending, (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "duplicate-author-id");
      return true;
    });
    assert.deepEqual([...root.children], [...before, external]);
    assert.equal(target.children.length, 0);
    assert.equal(root.query.getAll("duplicate").length, 1);
    assert.equal(root.query.get("external").node, external);
  });

  it("keeps ownerless named scopes reusable and later buildable", async () => {
    const root = box();
    const group = box({ id: "group" });
    const target = box();
    group.append(target);
    root.append(
      shape({ id: "shared", x: 0, y: 0, width: 10, height: 10 }),
      group,
    );

    const unique = await appendJSX(
      root,
      createElement(Shape, {
        id: "unique",
        x: 10,
        y: 0,
        width: 10,
        height: 10,
      }),
    );
    assert.equal(root.query.get("unique").node, unique[0]);

    const reused = await appendJSX(
      target,
      createElement(Shape, {
        id: "shared",
        x: 0,
        y: 0,
        width: 10,
        height: 10,
      }),
    );
    assert.equal(group.query.get("shared").node, reused[0]);
    assert.equal(root.query.getAll("shared").length, 1);

    const result = await build(
      deck({ width: 320, height: 180 }).append(
        slide({ id: "main" }).append(root),
      ),
    );
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const authorIds =
      result.report.slides[0]?.objects.map(({ authorId }) => authorId) ?? [];
    assert.equal(authorIds.filter((id) => id === "shared").length, 2);
    assert.ok(authorIds.includes("unique"));
  });

  it("rejects unhandled expansion and cleanup failures without touching the target", async () => {
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "failure-target" }),
      ),
    );
    const target = container(expanded.query.get("failure-target").node);

    function FailedExpansion(): never {
      throw new Error("fragment expansion failure");
    }
    await assert.rejects(
      appendJSX(
        target,
        createElement(
          Fragment,
          null,
          createElement(
            Text,
            { id: "failed-fragment", x: 0, y: 0, width: 160, height: 24 },
            "Failed",
          ),
          createElement(FailedExpansion),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
        assert.match(error.message, /fragment expansion failure/u);
        return true;
      },
    );

    function FailedCleanup() {
      useLayoutEffect(
        () => () => {
          throw new Error("fragment cleanup failure");
        },
        [],
      );
      return createElement(
        Text,
        { id: "cleanup-fragment", x: 0, y: 0, width: 160, height: 24 },
        "Cleanup",
      );
    }
    await assert.rejects(
      appendJSX(target, createElement(FailedCleanup)),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
        assert.match(error.message, /fragment cleanup failure/u);
        return true;
      },
    );

    assert.equal(target.children.length, 0);
  });

  it("cancels a pending append without attaching the fragment or leaving listeners", async () => {
    const controller = new AbortController();
    const reason = new Error("stop fragment expansion");
    const started = Promise.withResolvers<void>();
    const never = new Promise<string>(() => {});
    let cleaned = 0;
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "cancel-target" }),
      ),
    );
    const target = container(expanded.query.get("cancel-target").node);

    function Started() {
      useEffect(() => {
        started.resolve();
      }, []);
      return null;
    }
    function CleanupProbe() {
      useEffect(
        () => () => {
          cleaned += 1;
        },
        [],
      );
      return null;
    }
    function PendingFragment() {
      return createElement(
        Text,
        { id: "cancelled-fragment", x: 0, y: 0, width: 160, height: 24 },
        use(never),
      );
    }

    const appending = appendJSX(
      target,
      createElement(
        Fragment,
        null,
        createElement(Started),
        createElement(CleanupProbe),
        createElement(
          Suspense,
          { fallback: null },
          createElement(PendingFragment),
        ),
      ),
      { signal: controller.signal },
    );
    await started.promise;
    assert.equal(getEventListeners(controller.signal, "abort").length, 1);
    controller.abort(reason);
    await assert.rejects(appending, (error: unknown) => error === reason);
    assert.equal(cleaned, 1);
    assert.equal(getEventListeners(controller.signal, "abort").length, 0);
    assert.equal(target.children.length, 0);
  });

  it("applies the nearest named author scope when appending", async () => {
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(
          Slide,
          { id: "scope-target" },
          createElement(
            Text,
            {
              id: "title",
              x: 0,
              y: 0,
              width: 160,
              height: 24,
            },
            "Slide title",
          ),
          createElement(Box, { id: "card" }),
        ),
      ),
    );
    const slide = container(expanded.query.get("scope-target").node);
    const card = container(slide.query.get("card").node);

    await appendJSX(
      card,
      createElement(
        Text,
        {
          id: "title",
          x: 0,
          y: 0,
          width: 140,
          height: 24,
        },
        "Card title",
      ),
    );
    await assert.rejects(
      appendJSX(
        slide,
        createElement(
          Text,
          {
            id: "title",
            x: 0,
            y: 40,
            width: 160,
            height: 24,
          },
          "Duplicate slide title",
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "duplicate-author-id");
        return true;
      },
    );

    assert.equal(card.query.getAll("title").length, 1);
    assert.equal(slide.query.getAll("title").length, 1);
    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Slide title", "Card title"],
    );
  });

  it("appends slides to a deck and rejects duplicate deck-scoped slide ids", async () => {
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 320, height: 180 },
        createElement(Slide, { id: "first-slide" }),
      ),
    );
    const deck = container(expanded.query.node);

    const appended = await appendJSX(
      deck,
      createElement(
        Slide,
        { id: "second-slide" },
        createElement(
          Text,
          {
            id: "copy",
            x: 0,
            y: 0,
            width: 160,
            height: 24,
          },
          "Second slide",
        ),
      ),
    );
    assert.deepEqual(
      appended.map((node) => node.id),
      ["second-slide"],
    );
    await assert.rejects(
      appendJSX(deck, createElement(Slide, { id: "first-slide" })),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "duplicate-author-id");
        return true;
      },
    );

    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides.map(({ authorId }) => authorId),
      ["first-slide", "second-slide"],
    );
    assert.deepEqual(
      result.report.slides[1]?.objects.map(({ text: value }) => value),
      ["Second slide"],
    );
  });

  it("derives theme, layout, motion scope, step, default, disabled, and cue behavior from the target", async () => {
    const outer = createMotionScopeHandle();
    const expanded = await expandDocument(
      createElement(
        Deck,
        {
          width: 480,
          height: 320,
          theme: defineTheme({ colors: { accent: "#336699" } }),
        },
        createElement(
          Slide,
          { id: "inherit-target" },
          createElement(
            Box,
            {
              id: "layout-slot",
              style: {
                display: "block",
                position: "absolute",
                left: 240,
                top: 0,
                width: 220,
                height: 120,
              },
            },
            createElement(Shape, {
              id: "layout-existing",
              width: 40,
              height: 20,
            }),
          ),
          createElement(
            Box,
            {
              id: "default-slot",
              enter: fade({ duration: 120 }),
              style: {
                position: "absolute",
                left: 0,
                top: 0,
                width: 160,
                height: 80,
              },
            },
            createElement(Shape, {
              id: "default-existing",
              x: 0,
              y: 0,
              width: 10,
              height: 10,
            }),
          ),
          createElement(
            Box,
            {
              id: "disabled-slot",
              motion: false,
              style: {
                position: "absolute",
                left: 0,
                top: 100,
                width: 160,
                height: 80,
              },
            },
            createElement(Shape, {
              id: "disabled-existing",
              x: 0,
              y: 0,
              width: 10,
              height: 10,
            }),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(
              MotionScope,
              { scope: outer, publishStart: outer.cue("ready") },
              createElement(Shape, {
                id: "scope-existing",
                x: 0,
                y: 200,
                width: 10,
                height: 10,
                enter: fade({ duration: 100 }),
              }),
            ),
          ),
        ),
      ),
    );
    const slideQuery = expanded.query.get("inherit-target");
    const slide = container(slideQuery.node);
    const layoutSlot = container(slideQuery.get("layout-slot").node);
    const defaultSlot = container(slideQuery.get("default-slot").node);
    const disabledSlot = container(slideQuery.get("disabled-slot").node);
    const scopeNode = findNode(expanded, "motion-scope");
    assert.ok(scopeNode instanceof CommonContainer);

    await appendJSX(
      slide,
      createElement(
        Text,
        {
          id: "themed",
          x: 240,
          y: 200,
          width: 180,
          height: 30,
          style: { color: "accent" },
        },
        "Themed",
      ),
    );
    await appendJSX(
      layoutSlot,
      createElement(Text, { id: "flowed", width: 220, height: 24 }, "Flowed"),
    );
    await appendJSX(
      defaultSlot,
      createElement(Shape, {
        id: "default-inherited",
        x: 0,
        y: 20,
        width: 12,
        height: 12,
      }),
    );
    await appendJSX(
      disabledSlot,
      createElement(Shape, {
        id: "disabled-inherited",
        x: 0,
        y: 20,
        width: 12,
        height: 12,
        enter: fade({ duration: 40 }),
      }),
    );

    function ScopedFragment() {
      const inner = useMotionScope();
      return createElement(
        MotionScope,
        { scope: inner, startAt: after(outer.cue("ready"), 20) },
        createElement(Shape, {
          id: "scope-inherited",
          x: 20,
          y: 200,
          width: 12,
          height: 12,
          enter: fade({ duration: 50 }),
        }),
      );
    }
    await appendJSX(scopeNode, createElement(ScopedFragment));

    const result = await build(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const objects = result.report.slides[0]?.objects ?? [];
    const objectId = (authorId: string) => {
      const found = objects.find((object) => object.authorId === authorId);
      assert.ok(found, `missing built object for ${authorId}`);
      return found.id;
    };
    assert.equal(
      objects.find(({ authorId }) => authorId === "flowed")?.box?.y,
      20,
    );
    assert.equal(
      objects.find(({ authorId }) => authorId === "flowed")?.box?.width,
      220,
    );
    assert.match(slideXml(result, 1), /336699/u);

    const motion = result.report.slides[0]?.capabilities?.["motion"] as
      | MotionObservation
      | undefined;
    assert.ok(motion);
    const effects = motion.steps.flatMap((step) =>
      step.effects.map((effect) => ({ step, effect })),
    );
    const defaultEffect = effects.find(
      ({ effect }) => effect.target.id === objectId("default-inherited"),
    );
    assert.equal(defaultEffect?.effect.preset, "fade");
    assert.equal(defaultEffect?.effect.durationMs, 120);
    assert.equal(
      effects.some(
        ({ effect }) => effect.target.id === objectId("disabled-inherited"),
      ),
      false,
    );
    const scopedEffect = effects.find(
      ({ effect }) => effect.target.id === objectId("scope-inherited"),
    );
    assert.equal(scopedEffect?.step.trigger, "click");
    assert.equal(scopedEffect?.step.number, 2);

    const scopes = motion.scopes ?? [];
    const outerScope = scopes.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "object" && member.id === objectId("scope-existing"),
      ),
    );
    const innerScope = scopes.find(({ members }) =>
      members.some(
        (member) =>
          member.kind === "object" && member.id === objectId("scope-inherited"),
      ),
    );
    assert.ok(outerScope && innerScope);
    assert.equal(innerScope.parentId, outerScope.id);
    assert.equal(innerScope.startMs, 20);
  });
});
