import assert from "node:assert/strict";

import {
  Box,
  BuildError,
  Cell,
  Chart,
  Deck,
  P,
  Row,
  Slide,
  Shape,
  SlideNumber,
  Span,
  Table,
  Text,
  build,
  expandDocument,
} from "@react-pptx/react-pptx";
import {
  box,
  build as buildDocument,
  deck,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
  text,
  type CommonNode,
  type TextNode,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { describe, it } from "vitest";

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function textContent(node: CommonNode): string {
  return node.children
    .map((child) => ("text" in child ? child.text : ""))
    .join("");
}

function isText(node: CommonNode): node is TextNode {
  return node.kind === "text";
}

function authorId(document: CommonNode, id: string) {
  return document.query.get(id);
}

describe("scoped author ids", () => {
  it("looks up editable nodes through anonymous layers and edits them before the formal build", async () => {
    const titleNode = text("Draft heading", {
      id: "title",
      x: 0,
      y: 0,
      width: 600,
      height: 40,
    });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "cover" }).append(
        box({ style: { display: "block", width: 800 } }).append(titleNode),
      ),
    );

    const slideQuery = authorId(document, "cover");
    const titleQuery = slideQuery.get("title");
    assert.equal(titleQuery.node, titleNode);
    assert.equal(slideQuery.getNode("title"), titleNode);
    assert.match(
      titleQuery.path,
      /^deck\/slide\[cover\]\/box\/text\[title\]$/u,
    );

    const title = slideQuery.getAs("title", isText);
    assert.equal(title, titleNode);
    title.content("Edited through the scoped query");

    const result = await buildDocument(document);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.equal(result.report.slides[0]?.id, "cover");
    assert.equal(result.report.slides[0]?.authorId, "cover");
    const builtTitle = result.query.getSlide("cover").get("title");
    assert.equal(builtTitle.observe().authorId, "title");
    assert.equal(builtTitle.observe().text, "Edited through the scoped query");
    assert.match(slideXml(result), /Edited through the scoped query/u);
  });

  it("keeps anonymous layers transparent and stops at the next named scope", () => {
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "scope" }).append(
        box().append(
          text("Anonymous title", {
            id: "title",
            x: 0,
            y: 0,
            width: 300,
            height: 30,
          }),
        ),
        box({ id: "sidebar" }).append(
          text("Sidebar title", {
            id: "title",
            x: 0,
            y: 40,
            width: 300,
            height: 30,
          }),
        ),
      ),
    );
    const slideQuery = authorId(document, "scope");
    assert.equal(
      textContent(slideQuery.getAs("title", isText)),
      "Anonymous title",
    );
    assert.equal(
      textContent(slideQuery.get("sidebar").getAs("title", isText)),
      "Sidebar title",
    );
    assert.throws(
      () => slideQuery.get("sidebar").get("missing"),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "missing-id");
        return true;
      },
    );
    assert.throws(
      () =>
        slideQuery
          .get("sidebar")
          .getAs(
            "title",
            (node): node is CommonNode<"image"> => node.kind === "image",
          ),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "type-mismatch");
        return true;
      },
    );
  });

  it("builds P, Span, and SlideNumber author ids through the shared rich text parser", async () => {
    const document = deck({ width: 480, height: 270 }).append(
      slide({ id: "rich" }).append(
        text(
          paragraph("Body", {
            id: "body",
            style: { align: "center", fontSize: 18 },
          }),
          paragraph(
            [
              span("Bold", {
                id: "emphasis",
                style: { fontWeight: "bold", color: "#C00000" },
              }),
              slideNumber({ id: "page-number" }),
            ],
            { id: "footer", style: { spacingAfter: 6 } },
          ),
        ).set({ id: "copy", x: 10, y: 10, width: 300, height: 80 }),
      ),
    );

    const copy = document.query.get("rich").get("copy");
    assert.equal(copy.get("body").node.kind, "paragraph");
    assert.equal(copy.get("footer").get("emphasis").node.kind, "span");
    assert.equal(
      copy.get("footer").get("page-number").node.kind,
      "slide-number",
    );

    const result = await buildDocument(document);
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, []);
    const richText = validation.slides[0]?.objects.find(
      ({ text: value }) => value !== undefined,
    )?.text;
    assert.equal(richText?.text, "Body\nBold1");
    assert.equal(richText?.paragraphs[0]?.alignment, "ctr");
    assert.equal(richText?.paragraphs[0]?.runs[0]?.fontSizePt, 13.5);
    assert.equal(richText?.paragraphs[1]?.runs[0]?.bold, true);
    assert.equal(richText?.paragraphs[1]?.runs[0]?.color, "#C00000");
    assert.equal(richText?.paragraphs[1]?.spacingAfterPoints, 4.5);
    assert.equal(richText?.paragraphs[1]?.runs[1]?.field, "slidenum");
    assert.equal(richText?.paragraphs[1]?.runs[1]?.text, "1");
    assert.equal(
      result.query.getSlide("rich").get("copy").observe().text,
      "Body\nBold1",
    );
  });

  it("expands JSX rich text author ids into the same formal build", async () => {
    const expanded = await expandDocument(
      <Deck width={480} height={270}>
        <Slide id="jsx-rich">
          <Text id="copy" x={10} y={10} width={300} height={80}>
            <P id="body" style={{ align: "center", fontSize: 18 }}>
              Body
            </P>
            <P id="footer">
              <Span id="emphasis" style={{ fontWeight: "bold" }}>
                Bold
              </Span>
              <SlideNumber id="page-number" />
            </P>
          </Text>
        </Slide>
      </Deck>,
    );
    const copy = expanded.query.get("jsx-rich").get("copy");
    assert.equal(copy.get("body").node.kind, "paragraph");
    assert.equal(copy.get("footer").get("emphasis").node.kind, "span");
    assert.equal(
      copy.get("footer").get("page-number").node.kind,
      "slide-number",
    );

    const result = await buildDocument(expanded);
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, []);
    const richText = validation.slides[0]?.objects.find(
      ({ text: value }) => value !== undefined,
    )?.text;
    assert.equal(richText?.text, "Body\nBold1");
    assert.equal(richText?.paragraphs[0]?.alignment, "ctr");
    assert.equal(richText?.paragraphs[1]?.runs[0]?.bold, true);
    assert.equal(richText?.paragraphs[1]?.runs[1]?.field, "slidenum");
    assert.equal(
      result.query.getSlide("jsx-rich").get("copy").observe().text,
      "Body\nBold1",
    );
    assert.match(slideXml(result), /type="slidenum"/u);
  });

  it("validates rich text author ids in their named scopes", async () => {
    const duplicatedParagraph = deck({ width: 480, height: 270 }).append(
      slide({ id: "duplicate-paragraph" }).append(
        text(
          paragraph("First", { id: "same" }),
          paragraph("Second", { id: "same" }),
        ).set({ x: 0, y: 0, width: 240, height: 60 }),
      ),
    );
    await assert.rejects(
      buildDocument(duplicatedParagraph),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(({ code }) => code === "duplicate-author-id"),
        );
        return true;
      },
    );

    const duplicatedSpan = deck({ width: 480, height: 270 }).append(
      slide({ id: "duplicate-span" }).append(
        text(
          paragraph([
            span("First", { id: "same" }),
            span("Second", { id: "same" }),
          ]),
        ).set({ x: 0, y: 0, width: 240, height: 30 }),
      ),
    );
    await assert.rejects(buildDocument(duplicatedSpan), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.ok(
        error.diagnostics.some(({ code }) => code === "duplicate-author-id"),
      );
      return true;
    });

    const scoped = deck({ width: 480, height: 270 }).append(
      slide({ id: "scoped-rich" }).append(
        text(
          paragraph(["Left", span("L", { id: "part" })], { id: "left" }),
          paragraph(["Right", span("R", { id: "part" })], { id: "right" }),
          paragraph("Anonymous"),
        ).set({ id: "copy", x: 0, y: 0, width: 240, height: 90 }),
      ),
    );
    const result = await buildDocument(scoped);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.equal(
      result.query.getSlide("scoped-rich").get("copy").observe().text,
      "LeftL\nRightR\nAnonymous",
    );
  });

  it("keeps unknown rich text parameter diagnostics", async () => {
    for (const [label, node] of [
      ["P", paragraph("Body").set({ mystery: true } as never)],
      ["Span", span("Body").set({ mystery: true } as never)],
    ] as const) {
      const document = deck({ width: 320, height: 180 }).append(
        slide({ id: "unknown-rich" }).append(
          text(node).set({ x: 0, y: 0, width: 240, height: 40 }),
        ),
      );
      await assert.rejects(buildDocument(document), (error: unknown) => {
        assert.ok(error instanceof BuildError);
        const message =
          error.diagnostics.find(({ code }) => code === "unknown-parameter")
            ?.message ?? "";
        assert.match(
          message,
          new RegExp(`${label} contains unimplemented parameter mystery`, "u"),
        );
        return true;
      });
    }
  });

  it("rejects duplicate ids in one scope while allowing reuse across named scopes", async () => {
    const duplicated = deck({ width: 960, height: 540 }).append(
      slide({ id: "duplicates" }).append(
        text("First", { id: "same", x: 0, y: 0, width: 100, height: 20 }),
        text("Second", { id: "same", x: 0, y: 30, width: 100, height: 20 }),
      ),
    );
    assert.throws(
      () => authorId(duplicated, "duplicates").get("same"),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "ambiguous-id");
        return true;
      },
    );
    await assert.rejects(buildDocument(duplicated), (error: unknown) => {
      assert.ok(error instanceof BuildError);
      assert.ok(
        error.diagnostics.some(({ code }) => code === "duplicate-author-id"),
      );
      return true;
    });

    const independent = deck({ width: 960, height: 540 }).append(
      slide({ id: "independent" }).append(
        box({ id: "left" }).append(
          text("Left", { id: "title", x: 0, y: 0, width: 100, height: 20 }),
        ),
        box({ id: "right" }).append(
          text("Right", { id: "title", x: 0, y: 30, width: 100, height: 20 }),
        ),
      ),
    );
    const result = await buildDocument(independent);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.equal(
      result.query.getSlide("independent").get("left").get("title").observe()
        .text,
      "Left",
    );
    assert.equal(
      result.query.getSlide("independent").get("right").get("title").observe()
        .text,
      "Right",
    );
  });

  it("expands real React components and reuses internal names across instances", async () => {
    function Card({ caption }: { readonly caption: string }) {
      return (
        <Text
          id="title"
          x={0}
          y={0}
          width={180}
          height={30}
          style={{ fontSize: 16 }}
        >
          {caption}
        </Text>
      );
    }

    const expanded = await expandDocument(
      <Deck width={960} height={540}>
        <Slide id="instances">
          <Box id="left">
            <Card caption="Left card" />
          </Box>
          <Box id="right">
            <Card caption="Right card" />
          </Box>
        </Slide>
      </Deck>,
    );
    const slideQuery = expanded.query.get("instances");
    const leftTitle = slideQuery.get("left").getAs("title", isText);
    const rightTitle = slideQuery.get("right").getAs("title", isText);
    assert.equal(textContent(leftTitle), "Left card");
    assert.equal(textContent(rightTitle), "Right card");
    assert.notEqual(leftTitle, rightTitle);
    leftTitle.content("Edited left card");

    const result = await buildDocument(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    assert.deepEqual(
      result.report.slides[0]?.objects.map(({ text: value }) => value),
      ["Edited left card", "Right card"],
    );
    assert.equal(result.report.slides[0]?.objects[0]?.authorId, "title");
    assert.match(slideXml(result), /Edited left card/u);
  });

  it("keeps node identity while a move changes the query path, and freezes build observations", async () => {
    const moved = text("Before move", {
      id: "moved",
      x: 0,
      y: 0,
      width: 160,
      height: 30,
    });
    const source = box({ id: "source" }).append(moved);
    const target = box({ id: "target" });
    const document = deck({ width: 960, height: 540 }).append(
      slide({ id: "first" }).append(source),
      slide({ id: "second" }).append(target),
    );
    const first = await buildDocument(document);
    const firstObservation = first.query
      .getSlide("first")
      .get("source")
      .get("moved")
      .observe();
    const beforePath = authorId(document, "first")
      .get("source")
      .get("moved").path;

    moved.moveTo(target);
    assert.equal(
      authorId(document, "second").get("target").get("moved").node,
      moved,
    );
    const afterPath = authorId(document, "second")
      .get("target")
      .get("moved").path;
    assert.notEqual(afterPath, beforePath);
    assert.throws(
      () => authorId(document, "first").get("source").get("moved"),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "missing-id");
        return true;
      },
    );

    moved.content("After move");
    const second = await buildDocument(document);
    assert.equal(
      second.query.getSlide("second").get("target").get("moved").observe().text,
      "After move",
    );
    assert.equal(firstObservation.text, "Before move");
    assert.equal(firstObservation.authorId, "moved");
    assert.equal(
      first.query.getSlide("first").get("source").get("moved").observe().text,
      "Before move",
    );
    assert.ok(
      first.report.slides[0]?.objects.some(
        ({ authorId: value }) => value === "moved",
      ),
    );
    assert.ok(
      first.report.slides[1]?.objects.every(
        ({ authorId: value }) => value !== "moved",
      ),
    );

    moved.remove();
    assert.throws(
      () => authorId(document, "second").get("target").get("moved"),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "missing-id");
        return true;
      },
    );
    const third = await buildDocument(document);
    assert.ok(
      third.report.slides.every(({ objects }) =>
        objects.every(({ authorId: value }) => value !== "moved"),
      ),
    );
  });

  it("does not derive Morph or React key behavior from equal author ids", async () => {
    function RepeatedLabels() {
      return (
        <>
          {["first", "second"].map((label) => (
            <Box id={`box-${label}`} key={label}>
              <Shape id="shared" x={0} y={0} width={80} height={30}>
                {label}
              </Shape>
            </Box>
          ))}
        </>
      );
    }

    const expanded = await expandDocument(
      <Deck width={960} height={540}>
        <Slide id="keys">
          <RepeatedLabels />
        </Slide>
      </Deck>,
    );
    const slideQuery = expanded.query.get("keys");
    assert.equal(
      textContent(slideQuery.get("box-first").get("shared").node),
      "first",
    );
    assert.equal(
      textContent(slideQuery.get("box-second").get("shared").node),
      "second",
    );

    const result = await buildDocument(expanded);
    assert.deepEqual((await validatePptx(result.pptx)).errors, []);
    const shared = result.report.slides[0]?.objects.filter(
      ({ authorId }) => authorId === "shared",
    );
    assert.equal(shared?.length, 2);
    assert.equal(new Set(shared?.map(({ id }) => id)).size, 2);
    assert.ok(
      shared?.every(({ name }) => !name.startsWith("!!react-pptx-morph-")),
    );
    assert.equal(result.report.slides[0]?.capabilities?.transition, undefined);
    assert.throws(
      () => result.query.getSlide("missing"),
      (error: unknown) => {
        assert.equal((error as { code?: string }).code, "missing-slide");
        return true;
      },
    );
  });

  it("carries author ids through chart objects and table parts", async () => {
    const result = await build(
      <Deck width={480} height={270}>
        <Slide id="parts">
          <Chart
            id="sales"
            kind="column"
            x={10}
            y={10}
            width={200}
            height={120}
            data={{
              categories: ["Q1"],
              series: [{ name: "Bookings", values: [4] }],
            }}
          />
          <Table
            id="grid"
            x={10}
            y={140}
            columns={[120]}
            style={{ fontSize: 12 }}
          >
            <Row id="row" style={{ height: 24 }}>
              <Cell id="cell">Booked</Cell>
            </Row>
          </Table>
        </Slide>
      </Deck>,
    );
    const slide = result.query.getSlide("parts");
    assert.equal(slide.get("sales").observe().authorId, "sales");
    const row = slide.get("grid").get("row");
    assert.equal(row.observe().authorId, "row");
    assert.equal(row.get("cell").observe().authorId, "cell");
    assert.equal(row.get("cell").observe().text, "Booked");
    assert.ok(
      result.report.slides[0]?.objects.some(
        ({ authorId }) => authorId === "sales",
      ),
    );
  });

  it("matches React expansion and command authoring through the same formal build", async () => {
    const commandDocument = deck({ width: 320, height: 180 }).append(
      slide({ id: "same-intent" }).append(
        shape({ id: "badge", x: 10, y: 10, width: 120, height: 40 }).content(
          "Same intent",
        ),
      ),
    );
    const commandResult = await buildDocument(commandDocument);

    const reactResult = await build(
      <Deck width={320} height={180}>
        <Slide id="same-intent">
          <Shape id="badge" x={10} y={10} width={120} height={40}>
            Same intent
          </Shape>
        </Slide>
      </Deck>,
    );
    const reactObject = reactResult.report.slides[0]?.objects[0];
    const commandObject = commandResult.report.slides[0]?.objects[0];
    assert.equal(reactObject?.authorId, "badge");
    assert.equal(commandObject?.authorId, "badge");
    assert.equal(reactObject?.text, commandObject?.text);
    assert.deepEqual(reactObject?.box, commandObject?.box);
    assert.deepEqual(
      commandResult.query.getSlide("same-intent").get("badge").observe().box,
      reactResult.query.getSlide("same-intent").get("badge").observe().box,
    );
    assert.deepEqual((await validatePptx(reactResult.pptx)).errors, []);
    assert.deepEqual((await validatePptx(commandResult.pptx)).errors, []);
  });
});
