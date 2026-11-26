import assert from "node:assert/strict";

import {
  build,
  build as buildDocument,
  BuildError,
  CommonContainer,
  CommonNode,
  createMotionScopeHandle,
  resolve as resolveDocument,
  RichTextContainer,
} from "@officedk/pptxdk";
import {
  Box,
  Cell,
  Chart,
  Deck,
  expandDocument,
  Group,
  Image,
  MotionScope,
  Notes,
  P,
  Row,
  Section,
  Shape,
  Slide,
  SlideNumber,
  Span,
  Step,
  Table,
  Text,
} from "@officedk/pptxdk/react";
import {
  box,
  cell,
  chart,
  deck,
  group,
  image,
  motionScope,
  notes,
  paragraph,
  row,
  section,
  shape,
  slide,
  slideNumber,
  span,
  step,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  createElement,
  useEffect,
  useLayoutEffect,
  useState,
  type ReactNode,
} from "react";
import { describe, it } from "vitest";

function reactDeck(children: ReactNode) {
  return createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(Slide, { id: "handoff" }, children),
  );
}

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function findTextNode(parent: CommonNode, id: string): CommonNode | undefined {
  for (const child of parent.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.kind === "text" && child.props["id"] === id) {
      return child;
    }
    const found = findTextNode(child, id);
    if (found !== undefined) return found;
  }
  return undefined;
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

describe("public React handoff", () => {
  it("expands real components into editable common nodes and builds the edited document", async () => {
    const committed: string[] = [];

    function AuthoredContent() {
      const [heading] = useState("Handoff title");
      useEffect(() => {
        committed.push("committed");
      }, []);
      useLayoutEffect(
        () => () => {
          committed.push("layout-cleanup");
        },
        [],
      );
      return createElement(
        Text,
        {
          id: "handoff-title",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
            fontSize: 24,
            color: "#123456",
          },
        },
        heading,
      );
    }

    const expanded = await expandDocument(
      reactDeck(createElement(AuthoredContent)),
    );
    assert.deepEqual(committed, ["committed", "layout-cleanup"]);
    assert.ok(expanded instanceof CommonContainer);
    assert.equal(expanded.type, "deck");

    const title = findTextNode(expanded, "handoff-title");
    assert.ok(title instanceof RichTextContainer);
    title.content("Edited after handoff");
    title.style({ color: "#00AA00" });

    const first = await buildDocument(expanded);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    assert.equal(
      first.report.slides[0]?.objects[0]?.text,
      "Edited after handoff",
    );
    const firstReport = first.report;
    const firstPptx = first.pptx.slice();
    const firstXml = slideXml(first);
    assert.match(firstXml, /00AA00/u);

    title.content("Second edit");
    const second = await buildDocument(expanded);
    assert.equal(second.report.slides[0]?.objects[0]?.text, "Second edit");
    assert.match(slideXml(second), /Second edit/u);
    assert.doesNotMatch(firstXml, /Second edit/u);
    assert.deepEqual(first.report, firstReport);
    assert.deepEqual(first.pptx, firstPptx);
  });

  it("edits JSX shape rich text through the same node contract as the command factory", async () => {
    const props = {
      id: "handoff-shape",
      style: {
        position: "absolute" as const,
        left: 20,
        top: 24,
        width: 320,
        height: 96,
      },
    };
    const expanded = await expandDocument(
      reactDeck(
        createElement(
          Shape,
          props,
          createElement(P, null, createElement(Span, null, "Original")),
        ),
      ),
    );
    const shapeNode = findNode(expanded, "shape");
    assert.ok(shapeNode instanceof RichTextContainer);

    const first = await buildDocument(expanded);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    assert.equal(first.report.slides[0]?.objects[0]?.text, "Original");

    shapeNode.content(paragraph().content(span("Edited")));
    const second = await buildDocument(expanded);
    assert.deepEqual((await validatePptx(second.pptx)).errors, []);
    assert.equal(second.report.slides[0]?.objects[0]?.text, "Edited");
    assert.match(slideXml(second), /Edited/u);
    assert.equal(first.report.slides[0]?.objects[0]?.text, "Original");
    assert.doesNotMatch(slideXml(first), /Edited/u);

    const commandShape = shape(props).content(
      paragraph().content(span("Original")),
    );
    const commandDocument = deck({ width: 960, height: 540 }).append(
      slide({ id: "handoff" }).append(commandShape),
    );
    commandShape.content(paragraph().content(span("Edited")));
    const commandResult = await buildDocument(commandDocument);
    assert.deepEqual((await validatePptx(commandResult.pptx)).errors, []);
    assert.equal(commandResult.report.slides[0]?.objects[0]?.text, "Edited");
  });

  it("edits expanded table and chart nodes through the shared common document", async () => {
    const expanded = await expandDocument(
      reactDeck([
        createElement(
          Table,
          {
            id: "handoff-table",
            columns: [120, 120],
            style: { width: 240 },
          },
          createElement(
            Row,
            { style: { height: 30 } },
            createElement(Cell, { id: "handoff-cell" }, "Original"),
            createElement(Cell, null, "Value"),
          ),
        ),
        createElement(Chart, {
          id: "handoff-chart",
          style: {
            position: "absolute",
            left: 0,
            top: 40,
            width: 240,
            height: 160,
          },
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["Q1"] },
              series: [{ name: "Original", values: [1] }],
            },
          ],
        }),
      ]),
    );

    const tableNode = findNode(expanded, "table");
    assert.ok(tableNode instanceof CommonContainer);
    const cellNode = findNode(tableNode, "cell");
    assert.ok(cellNode instanceof RichTextContainer);
    const chartNode = findNode(expanded, "chart");
    assert.ok(chartNode instanceof CommonNode);

    const first = await buildDocument(expanded);
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(
      firstFile.slides[0]?.objects.find(({ kind }) => kind === "table")?.table
        ?.cells[0]?.[0],
      "Original",
    );
    assert.deepEqual(
      firstFile.slides[0]?.objects.find(({ kind }) => kind === "chart")?.chart
        ?.series,
      [
        {
          name: "Original",
          categories: ["Q1"],
          values: [1],
          index: 0,
          order: 0,
          groupIndex: 0,
        },
      ],
    );

    cellNode.content("Edited cell");
    chartNode.set({
      groups: [
        {
          kind: "column",
          data: { kind: "shared-categories", categories: ["R1"] },
          series: [{ name: "Revised", values: [7] }],
        },
      ],
    });
    const second = await buildDocument(expanded);
    const secondFile = await validatePptx(second.pptx);
    assert.deepEqual(secondFile.errors, []);
    assert.equal(
      secondFile.slides[0]?.objects.find(({ kind }) => kind === "table")?.table
        ?.cells[0]?.[0],
      "Edited cell",
    );
    assert.deepEqual(
      secondFile.slides[0]?.objects.find(({ kind }) => kind === "chart")?.chart
        ?.series,
      [
        {
          name: "Revised",
          categories: ["R1"],
          values: [7],
          index: 0,
          order: 0,
          groupIndex: 0,
        },
      ],
    );
    assert.equal(
      firstFile.slides[0]?.objects.find(({ kind }) => kind === "table")?.table
        ?.cells[0]?.[0],
      "Original",
    );
  });

  it("fails public expansion when React unmount cleanup throws", async () => {
    function CleanupFailure() {
      useLayoutEffect(
        () => () => {
          throw new Error("layout cleanup failure");
        },
        [],
      );
      return createElement(
        Text,
        {
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
          },
        },
        "Cleanup failure",
      );
    }

    await assert.rejects(
      expandDocument(reactDeck(createElement(CleanupFailure))),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "render-failed");
        assert.match(error.message, /layout cleanup failure/u);
        return true;
      },
    );
  });

  it("keeps received style values when unmount cleanup mutates the original style object", async () => {
    const style = { color: "#FF0000" };

    function MutatingCleanup() {
      useLayoutEffect(
        () => () => {
          style.color = "#0000FF";
        },
        [],
      );
      return createElement(
        Text,
        {
          id: "cleanup-style",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
            ...style,
          },
        },
        "Owned input",
      );
    }

    const expanded = await expandDocument(
      reactDeck(createElement(MutatingCleanup)),
    );
    const title = findTextNode(expanded, "cleanup-style");
    assert.ok(title instanceof RichTextContainer);

    const receivedXml = slideXml(await buildDocument(expanded));
    assert.match(receivedXml, /FF0000/u);
    assert.doesNotMatch(receivedXml, /0000FF/u);

    title.style({ color: "#00AA00" });
    const editedXml = slideXml(await buildDocument(expanded));
    assert.match(editedXml, /00AA00/u);
    assert.doesNotMatch(editedXml, /FF0000/u);
  });

  it("replaces complete JSX props on commit instead of accumulating earlier declarations", async () => {
    const renderedSteps: number[] = [];

    function ChangingContent() {
      const [step, setStep] = useState(0);
      renderedSteps.push(step);
      useEffect(() => {
        setStep(1);
      }, []);
      return createElement(
        Text,
        {
          ...(step === 0 ? { id: "changing-style" } : {}),
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
            ...(step === 0
              ? { color: "#FF0000", fontSize: 32 }
              : { color: "#0000FF" }),
          },
        },
        "Complete props",
      );
    }

    const result = await build(reactDeck(createElement(ChangingContent)));
    const file = await validatePptx(result.pptx);
    const run = file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0];

    assert.deepEqual(renderedSteps, [0, 1]);
    assert.deepEqual(file.errors, []);
    assert.equal(run?.color, "#0000FF");
    assert.equal(run?.fontSizePt, 18);
    assert.equal(result.report.slides[0]?.objects[0]?.authorId, undefined);
  });

  it("creates an independent document on repeated expansion instead of merging manual edits", async () => {
    function StaticContent() {
      return createElement(
        Text,
        {
          id: "repeat",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: 640,
            height: 48,
          },
        },
        "Original",
      );
    }

    const first = await expandDocument(reactDeck(createElement(StaticContent)));
    const firstTitle = findTextNode(first, "repeat");
    assert.ok(firstTitle instanceof RichTextContainer);
    firstTitle.content("Manual edit");

    const second = await expandDocument(
      reactDeck(createElement(StaticContent)),
    );
    assert.notEqual(first, second);
    const secondTitle = findTextNode(second, "repeat");
    assert.ok(secondTitle instanceof RichTextContainer);
    assert.notEqual(firstTitle, secondTitle);

    assert.equal(
      (await buildDocument(first)).report.slides[0]?.objects[0]?.text,
      "Manual edit",
    );
    assert.equal(
      (await buildDocument(second)).report.slides[0]?.objects[0]?.text,
      "Original",
    );
  });

  it("uses identical plain host names across JSX and imperative trees, diagnostics, and native PPTX output", async () => {
    const png1x1 = new Uint8Array([
      137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1,
      0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84,
      120, 156, 99, 248, 207, 192, 240, 31, 0, 5, 0, 1, 255, 137, 153, 61, 29,
      0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ]);
    const scopeHandle = createMotionScopeHandle();
    const chartData = [
      {
        kind: "column" as const,
        data: { kind: "shared-categories" as const, categories: ["Q1"] },
        series: [{ name: "Sales", values: [12] }],
      },
    ];

    const jsxExpanded = await expandDocument(
      createElement(
        Deck,
        { width: 960, height: 540 },
        createElement(
          Section,
          { name: "Main" },
          createElement(
            Slide,
            { id: "overview" },
            createElement(
              MotionScope,
              { scope: scopeHandle },
              createElement(
                Step,
                { n: 1 },
                createElement(
                  Box,
                  {
                    id: "container",
                    style: { width: 200, height: 60 },
                  },
                  createElement(
                    Shape,
                    {
                      id: "hero-shape",
                      style: { width: 120, height: 48 },
                    },
                    "Boxed",
                  ),
                ),
              ),
            ),
            createElement(
              Text,
              {
                id: "title",
                style: {
                  position: "absolute",
                  left: 20,
                  top: 100,
                  width: 240,
                  height: 40,
                },
              },
              createElement(
                P,
                null,
                createElement(Span, null, "Page "),
                createElement(SlideNumber, null),
              ),
            ),
            createElement(Image, {
              id: "logo",
              src: png1x1,
              style: {
                position: "absolute",
                left: 280,
                top: 20,
                width: 32,
                height: 32,
              },
            }),
            createElement(
              Group,
              { id: "pair" },
              createElement(Shape, {
                id: "in-group",
                style: {
                  position: "absolute",
                  left: 280,
                  top: 80,
                  width: 80,
                  height: 40,
                },
              }),
            ),
            createElement(
              Table,
              {
                id: "grid",
                columns: [100],
                style: { position: "absolute", left: 20, top: 160 },
              },
              createElement(
                Row,
                { style: { height: 32 } },
                createElement(Cell, null, "Cell"),
              ),
            ),
            createElement(Chart, {
              id: "sales-chart",
              style: {
                position: "absolute",
                left: 400,
                top: 20,
                width: 240,
                height: 160,
              },
              groups: chartData,
            }),
            createElement(Notes, null, createElement(P, null, "Speaker note")),
          ),
        ),
      ),
    );

    const imperativeDoc = deck({ width: 960, height: 540 }).append(
      section({ name: "Main" }).append(
        slide({ id: "overview" }).append(
          motionScope(scopeHandle).append(
            step({ n: 1 }).append(
              box({
                id: "container",
                style: { width: 200, height: 60 },
              }).append(
                shape("Boxed", {
                  id: "hero-shape",
                  style: { width: 120, height: 48 },
                }),
              ),
            ),
          ),
          text([paragraph([span("Page "), slideNumber()])], {
            id: "title",
            style: {
              position: "absolute",
              left: 20,
              top: 100,
              width: 240,
              height: 40,
            },
          }),
          image({
            id: "logo",
            src: png1x1,
            style: {
              position: "absolute",
              left: 280,
              top: 20,
              width: 32,
              height: 32,
            },
          }),
          group({ id: "pair" }).append(
            shape({
              id: "in-group",
              style: {
                position: "absolute",
                left: 280,
                top: 80,
                width: 80,
                height: 40,
              },
            }),
          ),
          table({
            id: "grid",
            columns: [100],
            style: { position: "absolute", left: 20, top: 160 },
          }).append(row({ style: { height: 32 } }).append(cell("Cell"))),
          chart({
            id: "sales-chart",
            style: {
              position: "absolute",
              left: 400,
              top: 20,
              width: 240,
              height: 160,
            },
            groups: chartData,
          }),
          notes(paragraph("Speaker note")),
        ),
      ),
    );

    const collectHostTypes = (root: CommonNode): string[] => {
      const types: string[] = [];
      const visit = (node: unknown) => {
        if (!(node instanceof CommonNode)) return;
        assert.equal(node.type, node.kind);
        assert.equal(node.type.includes(":"), false);
        types.push(node.type);
        for (const child of node.children) visit(child);
      };
      visit(root);
      return types;
    };

    const expectedTypes = [
      "deck",
      "section",
      "slide",
      "motion-scope",
      "motion-step",
      "box",
      "shape",
      "text",
      "paragraph",
      "span",
      "slide-number",
      "image",
      "group",
      "shape",
      "table",
      "row",
      "cell",
      "chart",
      "notes",
      "paragraph",
    ];
    assert.deepEqual(collectHostTypes(jsxExpanded), expectedTypes);
    assert.deepEqual(collectHostTypes(imperativeDoc), expectedTypes);

    const jsxBuilt = await build(jsxExpanded);
    const impBuilt = await buildDocument(imperativeDoc);

    const jsxValidated = await validatePptx(jsxBuilt.pptx);
    const impValidated = await validatePptx(impBuilt.pptx);
    assert.deepEqual(jsxValidated.errors, []);
    assert.deepEqual(impValidated.errors, []);
    assert.deepEqual(
      jsxBuilt.report.slides[0]?.objects.map((o) => o.kind),
      ["shape", "text", "image", "group", "table", "chart"],
    );
    assert.deepEqual(
      impBuilt.report.slides[0]?.objects.map((o) => o.kind),
      ["shape", "text", "image", "group", "table", "chart"],
    );
    assert.equal(jsxValidated.slides[0]?.objects.length, 6);
    assert.equal(impValidated.slides[0]?.objects.length, 6);

    const xml = slideXml(impBuilt);
    assert.match(xml, /<p:sp>/u);
    assert.match(xml, /<p:pic>/u);
    assert.match(xml, /<p:grpSp>/u);
    assert.match(xml, /<p:graphicFrame>/u);

    const badDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "diag" }).append(
        shape({
          id: "invalid-shape",
          style: {
            position: "absolute",
            left: 0,
            top: 0,
            width: -10,
            height: 20,
          },
        }),
      ),
    );
    assert.throws(
      () => resolveDocument(badDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.path === "deck/slide[diag]/shape[invalid-shape]",
    );
  });
});
