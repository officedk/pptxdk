import assert from "node:assert/strict";

import { validatePptx } from "@pptxdk/validation";
import {
  build,
  BuildError,
  Cell,
  Chart,
  CommonContainer,
  CommonNode,
  Deck,
  expandDocument,
  P,
  RichTextContainer,
  Row,
  Shape,
  Slide,
  Span,
  Table,
  Text,
} from "@react-pptx/react-pptx";
import {
  build as buildDocument,
  deck,
  paragraph,
  shape,
  slide,
  span,
} from "@react-pptx/react-pptx/document";
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
          x: 0,
          y: 0,
          width: 640,
          height: 48,
          style: { fontSize: 24, color: "#123456" },
        },
        heading,
      );
    }

    const expanded = await expandDocument(
      reactDeck(createElement(AuthoredContent)),
    );
    assert.deepEqual(committed, ["committed", "layout-cleanup"]);
    assert.ok(expanded instanceof CommonContainer);
    assert.equal(expanded.type, "rp:deck");

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
      x: 20,
      y: 24,
      width: 320,
      height: 96,
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
          { id: "handoff-table", width: 240, columns: [120, 120] },
          createElement(
            Row,
            { style: { height: 30 } },
            createElement(Cell, { id: "handoff-cell" }, "Original"),
            createElement(Cell, null, "Value"),
          ),
        ),
        createElement(Chart, {
          id: "handoff-chart",
          x: 0,
          y: 40,
          width: 240,
          height: 160,
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
        { x: 0, y: 0, width: 640, height: 48 },
        "Cleanup failure",
      );
    }

    await assert.rejects(
      expandDocument(reactDeck(createElement(CleanupFailure))),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "react-render");
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
          x: 0,
          y: 0,
          width: 640,
          height: 48,
          style,
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
          x: 0,
          y: 0,
          width: 640,
          height: 48,
          style:
            step === 0
              ? { color: "#FF0000", fontSize: 32 }
              : { color: "#0000FF" },
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
        { id: "repeat", x: 0, y: 0, width: 640, height: 48 },
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
});
