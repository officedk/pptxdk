import assert from "node:assert/strict";

import {
  build,
  BuildError,
  CommonContainer,
  CommonNode,
  Deck,
  expandDocument,
  RichTextContainer,
  Slide,
  Text,
} from "@react-pptx/react-pptx";
import { build as buildDocument } from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
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
    createElement(Slide, { slideId: "handoff" }, children),
  );
}

function slideXml(result: { readonly pptx: Uint8Array }, index = 1): string {
  return strFromU8(unzipSync(result.pptx)[`ppt/slides/slide${index}.xml`]!);
}

function findTextNode(
  parent: CommonNode,
  testId: string,
): CommonNode | undefined {
  for (const child of parent.children) {
    if (!(child instanceof CommonNode)) continue;
    if (child.kind === "text" && child.props["testId"] === testId) {
      return child;
    }
    const found = findTextNode(child, testId);
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
          testId: "handoff-title",
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
          testId: "cleanup-style",
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
          ...(step === 0 ? { testId: "changing-style" } : {}),
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
    assert.equal(result.report.slides[0]?.objects[0]?.testId, undefined);
  });

  it("creates an independent document on repeated expansion instead of merging manual edits", async () => {
    function StaticContent() {
      return createElement(
        Text,
        { testId: "repeat", x: 0, y: 0, width: 640, height: 48 },
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
