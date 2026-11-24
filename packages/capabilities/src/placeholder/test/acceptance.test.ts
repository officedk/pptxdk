import assert from "node:assert/strict";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  deck,
  defineSlideLayout,
  slide,
  slideLayout,
  slideMaster,
} from "@pptxdk/capabilities/document";
import { createPlaceholderNode } from "@pptxdk/capabilities/placeholder";
import { BuildError } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import { paint } from "@pptxdk/preview";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { createElement } from "react";
import { describe, expect, it } from "vitest";

import { runCapabilityScenario } from "../../../test/acceptance-runner.ts";
import { contentPlaceholderRoundtripScenario } from "./mechanisms.ts";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);

describe("Placeholder capability acceptance & isolation", () => {
  it("runs unpopulated content placeholder dual-author scenario through shared runner", async () => {
    const execution = await runCapabilityScenario(
      contentPlaceholderRoundtripScenario,
    );
    assert.equal(execution.runs.length, 2);
    assert.deepEqual(
      execution.runs.map((run) => run.author),
      ["jsx", "document"],
    );
    assert.deepEqual(execution.roundtrip.importReport.diagnostics, []);
  });

  it("rejects invalid placeholder declarations at resolve time", async () => {
    const validLayout = defineSlideLayout({
      id: "reject-layout",
      name: "Reject Layout",
      type: "obj",
      placeholders: {
        content: { type: "obj", idx: 1, x: 40, y: 80, width: 240, height: 180 },
      },
    });

    assert.throws(
      () =>
        resolveDocument(
          deck().append(
            slide({ id: "missing-binding" }).append(createPlaceholderNode({})),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some(
          (diag) => diag.code === "missing-placeholder-binding",
        ),
    );

    assert.throws(
      () =>
        resolveDocument(
          deck().append(
            slideMaster({ id: "m-rej" }).append(slideLayout(validLayout)),
            slide({ id: "unknown-prop", layout: validLayout }).append(
              createPlaceholderNode({
                placeholder: "content",
                bogus: 1,
              } as never),
            ),
          ),
        ),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "unknown-parameter"),
    );

    const badChildDoc = await expandDocument(
      createElement(
        Deck,
        null,
        createElement(
          SlideMaster,
          { id: "m-rej" },
          createElement(SlideLayout, { layout: validLayout }),
        ),
        createElement(
          Slide,
          { id: "with-child", layout: validLayout },
          createElement(
            "placeholder" as never,
            { placeholder: "content" },
            createElement("shape" as never, {
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            }),
          ),
        ),
      ),
    );
    assert.throws(
      () => resolveDocument(badChildDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics.some((diag) => diag.code === "unsupported-children"),
    );
  });

  it("isolates a corrupt unpopulated placeholder missing <p:spPr> with exact location and disposition while preserving healthy siblings", async () => {
    const twoPhLayout = defineSlideLayout({
      id: "two-ph-layout",
      name: "Two Placeholders",
      type: "twoObj",
      preserve: true,
      placeholders: {
        left: { type: "obj", idx: 1, x: 40, y: 80, width: 240, height: 200 },
        right: { type: "obj", idx: 2, x: 340, y: 80, width: 240, height: 200 },
      },
    });
    const sourceDoc = deck({ width: 640, height: 360 }).append(
      slideMaster({ id: "ph-m", name: "PH Master", preserve: true }).append(
        slideLayout(twoPhLayout),
      ),
      slide({ id: "s1", layout: twoPhLayout }).append(
        createPlaceholderNode({
          id: "ph-corrupt",
          name: "Corrupt PH",
          placeholder: "left",
        }),
        createPlaceholderNode({
          id: "ph-healthy-sibling",
          name: "Healthy Sibling PH",
          placeholder: "right",
        }),
      ),
      slide({ id: "s2", layout: twoPhLayout }).append(
        createPlaceholderNode({
          id: "ph-s2-healthy",
          name: "Slide 2 Healthy PH",
          placeholder: "left",
        }),
      ),
    );

    const basePptx = emit(resolveDocument(sourceDoc)).pptx;
    const entries = unzipSync(basePptx);
    const slide1Xml = strFromU8(entries["ppt/slides/slide1.xml"]!);
    const corruptedSlide1Xml = slide1Xml.replace(/<p:spPr\/>/u, "");
    assert.notEqual(corruptedSlide1Xml, slide1Xml);
    entries["ppt/slides/slide1.xml"] = strToU8(corruptedSlide1Xml);
    const corruptedBytes = new Uint8Array(zipSync(entries));

    const imported = importDocument(corruptedBytes);
    corruptedBytes.fill(0);

    expect(imported.report.diagnostics).toHaveLength(1);
    const diag = imported.report.diagnostics[0]!;
    expect(diag.code).toBe("corrupt-part");
    expect(diag.category).toBe("corrupt-input");
    expect(diag.disposition).toBe("omitted");
    expect(diag.location.part).toBe("ppt/slides/slide1.xml");
    expect(diag.location.slideIndex).toBe(0);
    expect(diag.location.xpath).toMatch(/\/p:sp\[1\]$/u);

    const recoveredResolution = resolveDocument(imported.document);
    const recoveredEmit = emit(recoveredResolution);
    const recoveredValidation = await validatePptx(recoveredEmit.pptx);
    assert.deepEqual(recoveredValidation.errors, []);
    expect(recoveredValidation.slides).toHaveLength(2);
    expect(recoveredValidation.slides[0]?.objects).toHaveLength(1);
    expect(recoveredValidation.slides[0]?.objects[0]?.name).toBe(
      "Healthy Sibling PH",
    );
    expect(recoveredValidation.slides[0]?.objects[0]?.placeholder).toEqual({
      type: "obj",
      idx: 2,
    });
    expect(recoveredValidation.slides[1]?.objects).toHaveLength(1);
    expect(recoveredValidation.slides[1]?.objects[0]?.name).toBe(
      "Slide 2 Healthy PH",
    );

    const recoveredPaint = paint(recoveredResolution);
    expect(recoveredPaint.pages).toHaveLength(2);
    expect(recoveredPaint.pages[0]?.diagnostics).toEqual([]);
    expect(recoveredPaint.pages[1]?.diagnostics).toEqual([]);
  });
});
