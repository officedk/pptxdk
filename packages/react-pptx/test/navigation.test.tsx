import assert from "node:assert/strict";

import { BuildError } from "@react-pptx/core";
import {
  build,
  Deck,
  Box,
  Cell,
  Chart,
  createMorphScope,
  Group,
  Image,
  linearGradient,
  P,
  Row,
  Section,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Table,
  Text,
  transitions,
} from "@react-pptx/react-pptx";
import { validatePptx } from "@react-pptx/validation";
import { createElement } from "react";
import { describe, it } from "vitest";

const pixelPng = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49,
  0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06,
  0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0b, 0x49, 0x44,
  0x41, 0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x02, 0x00, 0x00, 0x05, 0x00, 0x01,
  0xa5, 0xf6, 0x45, 0x40, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

describe("slide navigation", () => {
  it("writes internal jump relationships by stable slideId even after slide reordering", async () => {
    const buildWithOrder = async (slideIds: readonly string[]) => {
      const pages = slideIds.map((slideId) =>
        createElement(
          Slide,
          { key: slideId, slideId },
          slideId === "home"
            ? createElement(Text, {
                testId: "open-details",
                link: { slideId: "details" },
                x: 0,
                y: 0,
                width: 120,
                height: 48,
                children: "查看详情",
              })
            : null,
        ),
      );
      const result = await build(createElement(Deck, {}, ...pages));
      const file = await validatePptx(result.pptx);
      assert.deepEqual(file.errors, []);
      const source = file.slides.find(({ id }) => id === "home");
      const target = file.slides.find(({ id }) => id === "details");
      const link = source?.objects[0]?.link;
      assert.ok(link !== undefined && "slideId" in link);
      assert.equal(link.slideId, target?.id);
      return result;
    };

    const firstOrder = await buildWithOrder(["home", "details", "other"]);
    const reordered = await buildWithOrder(["other", "home", "details"]);

    assert.equal(
      firstOrder.report.slides[0]?.objects[0]?.link,
      "slide:details",
    );
    assert.equal(reordered.report.slides[1]?.objects[0]?.link, "slide:details");
  });

  it("independently reads title, sections, hidden and auto-advance slides, background, and native slide number fields", async () => {
    const result = await build(
      createElement(
        Deck,
        { title: "导航示例" },
        createElement(
          Section,
          { name: "开场" },
          createElement(
            Slide,
            {
              slideId: "opening",
              hidden: true,
              advanceAfter: 3539,
              background: "#112233",
            },
            createElement(
              Text,
              {
                x: 20,
                y: 30,
                width: 420,
                height: 60,
                style: { fontWeight: "bold" },
              },
              createElement(
                P,
                { style: { color: "#336699" } },
                "第 ",
                createElement(
                  Span,
                  { style: { fontStyle: "italic" } },
                  "页码 ",
                ),
                createElement(SlideNumber, {}),
              ),
            ),
          ),
        ),
        createElement(
          Section,
          { name: "内容" },
          createElement(
            Slide,
            {
              slideId: "gradient",
              background: linearGradient(45, "#112233", "#abcdef"),
            },
            createElement(
              Text,
              { x: 20, y: 30, width: 420, height: 60 },
              "第 ",
              createElement(SlideNumber, {}),
            ),
          ),
          createElement(Slide, {
            slideId: "image",
            background: { image: pixelPng },
          }),
        ),
      ),
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.equal(file.presentation.title, "导航示例");
    assert.deepEqual(file.presentation.sections, [
      { name: "开场", slideIds: ["opening"] },
      { name: "内容", slideIds: ["gradient", "image"] },
    ]);
    assert.equal(file.slides[0]?.hidden, true);
    assert.equal(file.slides[0]?.advanceAfterMs, 3539);
    assert.deepEqual(file.slides[0]?.background, {
      type: "solid",
      color: "#112233",
    });
    assert.equal(file.slides[1]?.background?.type, "gradient");
    assert.equal(file.slides[2]?.background?.type, "image");
    const field = file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[2];
    assert.equal(field?.field, "slidenum");
    assert.equal(field?.text, "1");
    assert.equal(field?.bold, true);
    assert.equal(field?.color, "#336699");
    assert.equal(
      file.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[1]?.text,
      "2",
    );
  });

  it("generates slide number fields using current slide sequence via shared text service in tables", async () => {
    const makeSlide = (slideId: string) =>
      createElement(
        Slide,
        { key: slideId, slideId },
        createElement(
          Table,
          { testId: `number-${slideId}`, x: 10, y: 10, columns: [160] },
          createElement(
            Row,
            { style: { height: 32 } },
            createElement(
              Cell,
              null,
              "第 ",
              createElement(SlideNumber, { style: { fontWeight: "bold" } }),
            ),
          ),
        ),
      );
    const result = await build(
      createElement(Deck, {}, makeSlide("one"), makeSlide("two")),
    );
    const file = await validatePptx(result.pptx);

    assert.deepEqual(file.errors, []);
    assert.deepEqual(
      file.slides.map((slide) => {
        const cell = slide.objects[0]?.table?.cellTexts[0]?.text;
        return cell?.paragraphs[0]?.runs.map(({ text, field, bold }) => [
          text,
          field,
          bold,
        ]);
      }),
      [
        [
          ["第 ", undefined, false],
          ["1", "slidenum", true],
        ],
        [
          ["第 ", undefined, false],
          ["2", "slidenum", true],
        ],
      ],
    );
  });

  it("writes and independently reads URLs, internal slides, and first-slide actions for each native object", async () => {
    const result = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { slideId: "source" },
          createElement(Shape, {
            testId: "shape-url",
            link: { url: "https://example.org/path?q=1" },
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(Text, {
            testId: "text-link",
            link: { slideId: "target" },
            x: 50,
            y: 0,
            width: 80,
            height: 30,
            children: "go",
          }),
          createElement(Image, {
            testId: "image-link",
            link: { slideId: "target" },
            src: pixelPng,
            x: 140,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(
            Group,
            { testId: "group-home", link: { action: "home" } },
            createElement(Shape, {
              testId: "group-member",
              x: 200,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
          createElement(Box, {
            testId: "visible-box-link",
            link: { slideId: "target" },
            style: {
              position: "absolute",
              left: 0,
              top: 50,
              width: 80,
              height: 30,
              backgroundColor: "#ffffff",
            },
          }),
          createElement(
            Table,
            {
              testId: "table-link",
              link: { slideId: "target" },
              columns: [60],
            },
            createElement(
              Row,
              { style: { height: 30 } },
              createElement(Cell, {}, "table"),
            ),
          ),
          createElement(Chart, {
            testId: "chart-link",
            link: { slideId: "target" },
            kind: "line",
            x: 0,
            y: 100,
            width: 100,
            height: 80,
            data: {
              categories: ["Q1"],
              series: [{ name: "Sales", values: [1] }],
            },
          }),
        ),
        createElement(Slide, { slideId: "target" }),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const source = file.slides[0];
    assert.deepEqual(
      source?.objects.map(({ name, link }) => [name, link]),
      [
        ["shape-url", { url: "https://example.org/path?q=1" }],
        ["text-link", { slideId: "target" }],
        ["image-link", { slideId: "target" }],
        ["group-home", { action: "home" }],
        ["visible-box-link", { slideId: "target" }],
        ["table-link", { slideId: "target" }],
        ["chart-link", { slideId: "target" }],
      ],
    );
  });

  it("validates native object kinds with Morph across explicit internal jump sources and the preceding visible slide of the target", async () => {
    const scope = createMorphScope();
    const hero = scope.handle("hero");
    const declarations = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { slideId: "linked-source" },
        createElement(Shape, {
          testId: "source-shape",
          link: { slideId: "morph-target" },
          morphKey: hero,
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
      createElement(Slide, { slideId: "hidden-intermediate", hidden: true }),
      createElement(
        Slide,
        {
          slideId: "morph-target",
          transition: transitions.morph({ fallback: transitions.fade() }),
        },
        createElement(Shape, {
          testId: "target-shape",
          morphKey: hero,
          x: 80,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
    );

    const result = await build(declarations);
    assert.equal(result.report.slides[2]?.objects[0]?.kind, "shape");
    const mismatched = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { slideId: "linked-source" },
        createElement(Shape, {
          link: { slideId: "morph-target" },
          morphKey: hero,
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
      createElement(Slide, { slideId: "middle" }),
      createElement(
        Slide,
        {
          slideId: "morph-target",
          transition: transitions.morph({ fallback: transitions.fade() }),
        },
        createElement(Image, {
          morphKey: hero,
          src: pixelPng,
          x: 80,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
    );
    await assert.rejects(build(mismatched), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.ok(
        error.diagnostics.some(
          ({ code }) => code === "morph-object-kind-mismatch",
        ),
      );
      return true;
    });
  });

  it("identifies missing jump targets, duplicate slide IDs, and transparent Box links with exact locations", async () => {
    const missingTarget = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { slideId: "source" },
        createElement(Text, {
          link: { slideId: "missing" },
          x: 0,
          y: 0,
          width: 80,
          height: 30,
          children: "missing",
        }),
      ),
    );
    await assert.rejects(build(missingTarget), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "missing-link-target");
      assert.match(error.diagnostics[0]?.path ?? "", /slide\[source\].*text/u);
      return true;
    });

    const duplicateIds = createElement(
      Deck,
      {},
      createElement(Slide, { slideId: "same" }),
      createElement(Slide, { slideId: "same" }),
    );
    await assert.rejects(build(duplicateIds), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "duplicate-slide-id");
      return true;
    });

    const transparentBox = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { slideId: "source" },
        createElement(
          Box,
          { link: { action: "home" } },
          createElement(Shape, {
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
        ),
      ),
    );
    await assert.rejects(build(transparentBox), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "link-needs-object");
      return true;
    });
  });
});
