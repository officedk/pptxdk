import assert from "node:assert/strict";

import { BuildError } from "@pptxdk/core";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync, zipSync } from "fflate";
import {
  CommonContainer,
  build,
  createMorphScope,
  importPptx,
  linearGradient,
  transitions,
} from "pptxdk";
import {
  Deck,
  Box,
  Cell,
  Chart,
  Group,
  Image,
  P,
  Row,
  Section,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Table,
  Text,
} from "pptxdk/react";
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
  it("writes internal jump relationships by stable author id even after slide reordering", async () => {
    const buildWithOrder = async (slideIds: readonly string[]) => {
      const pages = slideIds.map((id) =>
        createElement(
          Slide,
          { key: id, id },
          id === "home"
            ? createElement(Text, {
                id: "open-details",
                link: { id: "details" },
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
              id: "opening",
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
              id: "gradient",
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
            id: "image",
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
    const makeSlide = (id: string) =>
      createElement(
        Slide,
        { key: id, id },
        createElement(
          Table,
          { id: `number-${id}`, x: 10, y: 10, columns: [160] },
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
          { id: "source" },
          createElement(Shape, {
            id: "shape-url",
            name: "shape-url",
            link: { url: "https://example.org/path?q=1" },
            x: 0,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(Text, {
            id: "text-link",
            name: "text-link",
            link: { id: "target" },
            x: 50,
            y: 0,
            width: 80,
            height: 30,
            children: "go",
          }),
          createElement(Image, {
            id: "image-link",
            name: "image-link",
            link: { id: "target" },
            src: pixelPng,
            x: 140,
            y: 0,
            width: 40,
            height: 30,
          }),
          createElement(
            Group,
            { id: "group-home", name: "group-home", link: { action: "home" } },
            createElement(Shape, {
              id: "group-member",
              x: 200,
              y: 0,
              width: 40,
              height: 30,
            }),
          ),
          createElement(Box, {
            id: "visible-box-link",
            link: { id: "target" },
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
              id: "table-link",
              name: "table-link",
              link: { id: "target" },
              columns: [60],
            },
            createElement(
              Row,
              { style: { height: 30 } },
              createElement(Cell, {}, "table"),
            ),
          ),
          createElement(Chart, {
            id: "chart-link",
            name: "chart-link",
            link: { id: "target" },
            x: 0,
            y: 100,
            width: 100,
            height: 80,
            groups: [
              {
                kind: "line",
                data: { kind: "shared-categories", categories: ["Q1"] },
                series: [{ name: "Sales", values: [1] }],
              },
            ],
          }),
        ),
        createElement(Slide, { id: "target" }),
      ),
    );
    const file = await validatePptx(result.pptx);
    assert.deepEqual(file.errors, []);
    const source = file.slides[0];
    assert.deepEqual(
      source?.objects.map(({ id, link }) => [
        result.report.slides[0]?.objects.find((object) => object.id === id)
          ?.authorId,
        link,
      ]),
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
        { id: "linked-source" },
        createElement(Shape, {
          id: "source-shape",
          link: { id: "morph-target" },
          morphKey: hero,
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
      createElement(Slide, { id: "hidden-intermediate", hidden: true }),
      createElement(
        Slide,
        {
          id: "morph-target",
          transition: transitions.morph({ fallback: transitions.fade() }),
        },
        createElement(Shape, {
          id: "target-shape",
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
        { id: "linked-source" },
        createElement(Shape, {
          link: { id: "morph-target" },
          morphKey: hero,
          x: 0,
          y: 0,
          width: 40,
          height: 30,
        }),
      ),
      createElement(Slide, { id: "middle" }),
      createElement(
        Slide,
        {
          id: "morph-target",
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

  it("identifies missing jump targets, duplicate author ids, and transparent Box links with exact locations", async () => {
    const missingTarget = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { id: "source" },
        createElement(Text, {
          link: { id: "missing" },
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
      createElement(Slide, { id: "same" }),
      createElement(Slide, { id: "same" }),
    );
    await assert.rejects(build(duplicateIds), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.equal(error.diagnostics[0]?.code, "duplicate-author-id");
      return true;
    });

    const transparentBox = createElement(
      Deck,
      {},
      createElement(
        Slide,
        { id: "source" },
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

  it("preserves duplicate slide display names and internal slide-jump targets to the disambiguated second slide", async () => {
    const initial = await build(
      createElement(
        Deck,
        {},
        createElement(
          Slide,
          { id: "first", name: "Section" },
          createElement(Shape, {
            id: "shape-jump",
            x: 20,
            y: 20,
            width: 100,
            height: 40,
            link: { id: "second" },
          }),
          createElement(Text, {
            id: "text-jump",
            x: 20,
            y: 70,
            width: 120,
            height: 40,
            link: { id: "second" },
            children: "Go to Section 2",
          }),
          createElement(
            Group,
            { id: "group-jump", link: { id: "second" } },
            createElement(Shape, {
              id: "group-child",
              x: 160,
              y: 20,
              width: 60,
              height: 40,
            }),
          ),
          createElement(Chart, {
            id: "chart-jump",
            x: 20,
            y: 130,
            width: 160,
            height: 100,
            link: { id: "second" },
            groups: [
              {
                kind: "line",
                data: { kind: "shared-categories", categories: ["Q1"] },
                series: [{ name: "S1", values: [1] }],
              },
            ],
          }),
        ),
        createElement(Slide, { id: "second", name: "Section" }),
      ),
    );

    const files = unzipSync(initial.pptx);
    const slide1XmlBefore = strFromU8(files["ppt/slides/slide1.xml"]!);
    const slide2XmlBefore = strFromU8(files["ppt/slides/slide2.xml"]!);
    assert.match(slide1XmlBefore, /<p:cSld name="Section">/u);
    assert.match(slide2XmlBefore, /<p:cSld name="Section">/u);

    const imported = importPptx(zipSync(files));
    assert.deepEqual(imported.report.diagnostics, []);

    const slides = imported.document.children.filter(
      (child): child is CommonContainer =>
        child instanceof CommonContainer && child.kind === "slide",
    );
    assert.equal(slides.length, 2);
    assert.equal(slides[0]?.id, "Section");
    assert.equal(slides[1]?.id, "Section-2");
    assert.equal(slides[1]?.props["name"], "Section");

    const rebuilt = await build(imported.document);
    const validation = await validatePptx(rebuilt.pptx);
    assert.deepEqual(validation.errors, []);

    const rebuiltFiles = unzipSync(rebuilt.pptx);
    const rebuiltSlide1 = strFromU8(rebuiltFiles["ppt/slides/slide1.xml"]!);
    const rebuiltSlide2 = strFromU8(rebuiltFiles["ppt/slides/slide2.xml"]!);
    const rebuiltSlide1Rels = strFromU8(
      rebuiltFiles["ppt/slides/_rels/slide1.xml.rels"]!,
    );
    assert.match(rebuiltSlide1, /<p:cSld name="Section">/u);
    assert.match(rebuiltSlide2, /<p:cSld name="Section">/u);

    const jumpRelIds = [
      ...rebuiltSlide1.matchAll(
        /<a:hlinkClick r:id="([^"]+)" action="ppaction:\/\/hlinksldjump"/gu,
      ),
    ].map((match) => match[1]!);
    assert.equal(jumpRelIds.length, 4);
    for (const relId of jumpRelIds) {
      assert.match(
        rebuiltSlide1Rels,
        new RegExp(`Id="${relId}"[^>]*Target="slide2\\.xml"`, "u"),
      );
    }
  });

  it("disambiguates duplicate slide <p:cSld name> in validation while mapping each slide to its own section", async () => {
    const built = await build(
      <Deck>
        <Section name="S1">
          <Slide id="a" name="Section" />
        </Section>
        <Section name="S2">
          <Slide id="b" name="Section" />
        </Section>
      </Deck>,
    );
    const validation = await validatePptx(built.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(validation.slides[0]?.section, "S1");
    assert.equal(validation.slides[1]?.section, "S2");
    assert.equal(validation.slides[0]?.name, "Section");
    assert.equal(validation.slides[1]?.name, "Section");

    const imported = importPptx(built.pptx);
    assert.deepEqual(imported.report.diagnostics, []);
    const importedSections = imported.document.children.filter(
      (child): child is CommonContainer =>
        child instanceof CommonContainer && child.kind === "section",
    );
    assert.equal(importedSections.length, 2);
    assert.equal(importedSections[0]?.props["name"], "S1");
    assert.equal(importedSections[1]?.props["name"], "S2");
    const s1Slide = importedSections[0]?.children[0] as CommonContainer;
    const s2Slide = importedSections[1]?.children[0] as CommonContainer;
    assert.equal(s1Slide.id, "Section");
    assert.equal(s2Slide.id, "Section-2");
    assert.equal(s2Slide.props["name"], "Section");
  });
});
