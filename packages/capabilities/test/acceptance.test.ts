import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  BuildError,
  CommonNode,
  Deck,
  Slide,
  createPreview,
  expandDocument,
  pxToEmu,
  xmlEscape,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
  type Diagnostic,
} from "@react-pptx/core";
import { parseCapabilityLayoutContract } from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";
import { createDocumentPreview } from "@react-pptx/core/preview";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement, type ComponentType } from "react";
import { describe, expect, it } from "vitest";

import { boxCapability } from "../src/box.ts";
import {
  box,
  deck,
  documentCapabilities,
  font,
  notes,
  paragraph,
  shape,
  slide,
  slideNumber,
  span,
  text,
} from "../src/document.ts";
import { Notes } from "../src/notes-elements.ts";
import { notesCapability } from "../src/notes.ts";
import { Shape } from "../src/shape-elements.ts";
import { shapeCapability } from "../src/shape.ts";
import { slideBackgroundCapability } from "../src/slide-background.ts";
import { P, SlideNumber, Span, Text } from "../src/text-elements.ts";
import { runCapabilityScenario } from "./acceptance-runner.ts";
import { boxPositiveScenario, boxRejectionScenario } from "./scenarios/box.ts";
import {
  groupPositiveScenario,
  groupRejectionScenario,
} from "./scenarios/group.ts";
import {
  imagePositiveScenario,
  imageRejectionScenario,
} from "./scenarios/image.ts";
import {
  layoutBlockFlexPositiveScenario,
  layoutBlockFlexRejectionCases,
} from "./scenarios/layout-block-flex.ts";
import {
  layoutGridPositiveScenario,
  layoutGridRejectionCases,
} from "./scenarios/layout-grid.ts";
import {
  notesPositiveScenario,
  notesPreviewIsolationScenario,
} from "./scenarios/notes.ts";
import {
  shapePositiveScenario,
  shapeRejectionScenario,
} from "./scenarios/shape.ts";
import {
  slideBackgroundPositiveScenario,
  slideBackgroundRejectionScenario,
} from "./scenarios/slide-background.ts";
import {
  textPositiveScenario,
  textRejectionScenario,
} from "./scenarios/text.ts";
import {
  runTransitionRejectionScenarios,
  transitionPositiveScenario,
} from "./scenarios/transition.ts";

type CustomMarkerProps = {
  readonly id?: string;
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly color: string;
  readonly warnMessage?: string;
};

function createMarkerNode(
  props: CustomMarkerProps,
): CommonNode<"marker", CustomMarkerProps> {
  return new CommonNode("marker", "rp:custom-marker", props);
}

describe("Capability-owned acceptance scenarios & custom bound preview extensions", () => {
  it("runs Box, Shape, Image, slide-background, Text, and Notes dual-author scenarios through shared runner and collects artifacts", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-"));
    try {
      for (const scenario of [
        boxPositiveScenario,
        boxRejectionScenario,
        groupPositiveScenario,
        groupRejectionScenario,
        imagePositiveScenario,
        imageRejectionScenario,
        layoutBlockFlexPositiveScenario,
        layoutGridPositiveScenario,
        shapePositiveScenario,
        shapeRejectionScenario,
        slideBackgroundPositiveScenario,
        slideBackgroundRejectionScenario,
        textPositiveScenario,
        textRejectionScenario,
        notesPositiveScenario,
        notesPreviewIsolationScenario,
      ]) {
        const runs = await runCapabilityScenario(scenario, {
          outputDir: outDir,
        });
        assert.equal(runs.length, 2);
        assert.deepEqual(
          runs.map((r) => r.author),
          ["jsx", "document"],
        );
      }
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 30_000);

  it("enforces unsupported Block/Flex and Grid layout rejections consistently across JSX and document entries", async () => {
    const jsxPreview = createPreview(documentCapabilities);
    const docPreview = createDocumentPreview(documentCapabilities);
    for (const entry of [
      ...layoutBlockFlexRejectionCases,
      ...layoutGridRejectionCases,
    ]) {
      for (const author of ["jsx", "document"] as const) {
        await assert.rejects(
          () =>
            author === "jsx"
              ? jsxPreview(entry.jsx())
              : docPreview(entry.document()),
          (error: unknown) => {
            assert(
              error instanceof BuildError,
              `${entry.id} (${author}) should throw BuildError`,
            );
            assert.ok(
              error.diagnostics.some(
                (d) =>
                  d.code === entry.expectedCode &&
                  d.path.includes(entry.expectedPathSubstring),
              ),
              `${entry.id} (${author}) expected ${entry.expectedCode} at ${entry.expectedPathSubstring}, got ${JSON.stringify(error.diagnostics)}`,
            );
            return true;
          },
        );
      }
    }
  });

  it("runs Transition and Morph dual-author scenarios through shared runner and verifies rejections", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-trans-"));
    try {
      const runs = await runCapabilityScenario(transitionPositiveScenario, {
        outputDir: outDir,
      });
      assert.equal(runs.length, 2);
      assert.deepEqual(
        runs.map((r) => r.author),
        ["jsx", "document"],
      );
      await runTransitionRejectionScenarios();
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 15_000);

  it("supports JSX expansion followed by imperative Text edit, rebuild, and preview without mutating prior outputs", async () => {
    const jostFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL(
            "../../react-pptx/test/assets/fonts/Jost-400-Book.ttf",
            import.meta.url,
          ),
        ),
      },
    ]);
    const expanded = await expandDocument(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "handoff" },
          createElement(
            Text,
            {
              id: "headline",
              x: 24,
              y: 24,
              width: 200,
              style: {
                fontFamily: jostFont,
                fontSize: 20,
                color: "#1E293B",
                padding: 8,
                backgroundColor: "#F1F5F9",
              },
            },
            "Initial title",
          ),
        ),
      ),
    );

    const previewDocument = createDocumentPreview(documentCapabilities);
    const first = await previewDocument(expanded);
    assert.deepEqual(first.pages[0]?.diagnostics, []);
    assert.deepEqual((await validatePptx(first.build.pptx)).errors, []);
    const firstPng = first.pages[0].png!.slice();
    const firstSvg = first.pages[0].svg!;
    const firstPptx = first.build.pptx.slice();
    const firstHeight = first.build.query
      .getSlide("handoff")
      .get("headline")
      .observe().box!.height;

    const headline = expanded.query
      .get("handoff")
      .getAs(
        "headline",
        (node): node is RichTextContainer => node instanceof RichTextContainer,
      );
    headline.content("Updated longer headline wrapping across multiple lines");
    headline.style({ color: "#059669" });

    const second = await previewDocument(expanded);
    assert.deepEqual(second.pages[0]?.diagnostics, []);
    assert.deepEqual((await validatePptx(second.build.pptx)).errors, []);
    assert.equal(
      second.build.query.getSlide("handoff").get("headline").observe().text,
      "Updated longer headline wrapping across multiple lines",
    );
    const secondHeight = second.build.query
      .getSlide("handoff")
      .get("headline")
      .observe().box!.height;
    assert.ok(secondHeight > firstHeight);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // Prior BuildResult and PreviewResult remain unchanged.
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.build.pptx, firstPptx);
    assert.equal(
      first.build.query.getSlide("handoff").get("headline").observe().text,
      "Initial title",
    );
  });

  it("supports registering custom bound extension modules, warning retention, and missing preview rejection", async () => {
    const calls: string[] = [];

    const customMarkerWithPreview: BoundObjectCapabilityModule = {
      declaration: "rp:custom-marker",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const x = Number(input.props["x"]);
        const y = Number(input.props["y"]);
        const size = Number(input.props["size"]);
        const color = String(input.props["color"]);
        const warnMessage =
          typeof input.props["warnMessage"] === "string"
            ? input.props["warnMessage"]
            : undefined;
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: { x, y, width: size, height: size },
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        calls.push(`bind:${parseContext.path}`);

        return {
          contract,
          compile(compileContext) {
            calls.push(`compile:${compileContext.path}`);
            const box = compileContext.layoutBox ?? {
              x,
              y,
              width: size,
              height: size,
            };
            const name = xmlEscape(`Marker ${compileContext.id}`);
            const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`;
            return {
              xml,
              observation: {
                id: compileContext.id,
                kind: "shape",
                name,
                geometry: "ellipse",
                box,
              },
              extensions: [],
            };
          },
          preview(previewContext) {
            calls.push(`preview:${previewContext.path}`);
            const box = previewContext.box!;
            const resolved = previewContext.theme.resolveColor(color);
            return {
              svg: `<circle cx="${box.x + box.width / 2}" cy="${box.y + box.height / 2}" r="${box.width / 2}" fill="#${resolved.hex}"/>`,
              ...(warnMessage === undefined
                ? {}
                : { warnings: [{ message: warnMessage }] }),
            };
          },
        };
      },
    };

    const customMarkerWithoutPreview: BoundObjectCapabilityModule = {
      declaration: "rp:custom-no-preview",
      bind(input: CapabilityInput, parseContext: CapabilityParseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(),
          allowPadding: false,
          path: parseContext.path,
        });
        return {
          contract,
          compile(compileContext) {
            const box = compileContext.layoutBox ?? {
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            };
            return {
              xml: `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="NoPreview"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
              observation: {
                id: compileContext.id,
                kind: "shape",
                name: "NoPreview",
                geometry: "rect",
                box,
              },
              extensions: [],
            };
          },
        };
      },
    };

    const previewCustom = createDocumentPreview([
      shapeCapability,
      customMarkerWithPreview,
      customMarkerWithoutPreview,
    ]);

    const doc = deck({ width: 200, height: 120 }).append(
      slide({ id: "warn-keeps-png" }).append(
        createMarkerNode({
          id: "dot",
          x: 60,
          y: 20,
          size: 80,
          color: "#DC2626",
          warnMessage: "Custom approximation notice",
        }),
      ),
      slide({ id: "warn-cannot-clear-error" }).append(
        new CommonNode("noprev", "rp:custom-no-preview", {
          id: "opaque",
          x: 10,
          y: 10,
          width: 20,
          height: 20,
        }),
        createMarkerNode({
          id: "later-warn",
          x: 100,
          y: 20,
          size: 40,
          color: "#2563EB",
          warnMessage: "Later warning must not clear earlier error",
        }),
      ),
    );

    const result = await previewCustom(doc);
    // Page 1: warning preserves PNG and draws the red circle.
    assert.equal(result.pages[0]?.diagnostics.length, 1);
    assert.equal(result.pages[0]?.diagnostics[0]?.severity, "warning");
    assert.ok(result.pages[0]?.png !== undefined);
    const pixels = PNG.sync.read(Buffer.from(result.pages[0].png));
    const center = (60 * pixels.width + 100) * 4;
    assert.deepEqual(
      [...pixels.data.subarray(center, center + 4)],
      [0xdc, 0x26, 0x26, 255],
    );
    assert.ok(
      result.build.query.getSlide("warn-keeps-png").get("dot").observe().box
        ?.width === 80,
    );

    // Page 2: missing preview handler causes error; subsequent warning does NOT clear error or fabricate PNG.
    assert.equal(result.pages[1]?.png, undefined);
    expect(result.pages[1]?.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          severity: "error",
          path: "deck/slide[warn-cannot-clear-error]/custom-no-preview[opaque]",
        }),
        expect.objectContaining({
          severity: "warning",
          path: "deck/slide[warn-cannot-clear-error]/custom-marker[later-warn]",
        }),
      ]),
    );

    // Same custom bound extension works via JSX createPreview and shares trimmed constraints.
    const jsxCustomPreview = createPreview([shapeCapability]);
    const trimmedOk = await jsxCustomPreview(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(Slide, { id: "trimmed" }),
      ),
    );
    assert.ok(trimmedOk.pages[0]?.png !== undefined);

    await assert.rejects(
      () =>
        jsxCustomPreview(
          createElement(
            Deck,
            { width: 200, height: 120 },
            createElement(Slide, { id: "missing-bg", background: "#123456" }),
          ),
        ),
      BuildError,
    );
    void shape;
    assert.ok(calls.some((entry) => entry.startsWith("bind:")));
    assert.ok(calls.some((entry) => entry.startsWith("compile:")));
    assert.ok(calls.some((entry) => entry.startsWith("preview:")));
  });

  it("binds notesCapability via BoundSlideCapabilityModule without fake previewSlide, trims textCapability cleanly, and enforces rejections", async () => {
    assert.equal("bindSlide" in notesCapability, true);
    assert.equal("compileSlide" in notesCapability, false);
    assert.equal("measureLayout" in notesCapability, false);

    // Notes-only slide previews cleanly as an empty canvas without fake previewSlide errors or leaked notes text.
    const notesOnlyPreview = createDocumentPreview([
      notesCapability,
      slideBackgroundCapability,
    ]);
    const notesOnlyResult = await notesOnlyPreview(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "notes-only" }).append(notes("Speaker only")),
      ),
    );
    assert.deepEqual(notesOnlyResult.pages[0]?.diagnostics, []);
    assert.ok(notesOnlyResult.pages[0]?.png !== undefined);
    const notesOnlyPng = PNG.sync.read(
      Buffer.from(notesOnlyResult.pages[0].png),
    );
    for (let i = 0; i < notesOnlyPng.data.length; i += 4) {
      assert.deepEqual(
        [...notesOnlyPng.data.subarray(i, i + 4)],
        [255, 255, 255, 255],
      );
    }

    const trimmedModules = [
      shapeCapability,
      notesCapability,
      slideBackgroundCapability,
    ] as const;
    const previewDocTrimmed = createDocumentPreview(trimmedModules);
    const previewJsxTrimmed = createPreview(trimmedModules);

    const doc = deck({ width: 480, height: 270 }).append(
      slide({ id: "trimmed-text" }).append(
        shape({
          id: "card",
          x: 40,
          y: 40,
          width: 200,
          height: 100,
          fill: "#1D4ED8",
        }),
        notes(
          [
            paragraph([
              span("Trimmed text check ", { style: { fontWeight: "bold" } }),
              slideNumber(),
            ]),
          ],
          { id: "speaker" },
        ),
      ),
    );
    const docResult = await previewDocTrimmed(doc);
    assert.deepEqual(docResult.pages[0]?.diagnostics, []);
    assert.ok(docResult.pages[0]?.png !== undefined);
    assert.deepEqual(docResult.build.report.slides[0]?.capabilities, {
      notes: { text: "Trimmed text check 1", paragraphCount: 1 },
    });

    const jsxResult = await previewJsxTrimmed(
      createElement(
        Deck,
        { width: 480, height: 270 },
        createElement(
          Slide,
          { id: "trimmed-text" },
          createElement(Shape, {
            id: "card",
            x: 40,
            y: 40,
            width: 200,
            height: 100,
            fill: "#1D4ED8",
          }),
          createElement(
            Notes,
            { id: "speaker" },
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontWeight: "bold" } },
                "Trimmed text check ",
              ),
              createElement(SlideNumber, null),
            ),
          ),
        ),
      ),
    );
    assert.deepEqual(jsxResult.pages[0]?.diagnostics, []);
    assert.ok(jsxResult.pages[0]?.png !== undefined);
    assert.deepEqual(jsxResult.build.report.slides[0]?.capabilities, {
      notes: { text: "Trimmed text check 1", paragraphCount: 1 },
    });

    async function expectBuildErrorCode(
      run: () => Promise<unknown>,
    ): Promise<readonly Diagnostic[]> {
      try {
        await run();
      } catch (error) {
        assert.ok(error instanceof BuildError);
        return error.diagnostics;
      }
      assert.fail("Expected BuildError");
    }

    const withBoxModules = [...trimmedModules, boxCapability] as const;
    const previewDocWithBox = createDocumentPreview(withBoxModules);
    const previewJsxWithBox = createPreview(withBoxModules);
    const previewWithoutNotes = createDocumentPreview([
      shapeCapability,
      boxCapability,
    ]);
    const previewJsxWithoutNotes = createPreview([
      shapeCapability,
      boxCapability,
    ]);

    const dupHiddenDiags = await expectBuildErrorCode(() =>
      previewDocWithBox(
        deck().append(
          slide({ id: "dup" }).append(
            box({ style: { display: "none" } }).append(notes("hidden")),
            notes("visible"),
          ),
        ),
      ),
    );
    assert.deepEqual(
      dupHiddenDiags.map(({ code }) => code),
      ["duplicate-notes"],
    );

    const dupJsxDiags = await expectBuildErrorCode(() =>
      previewJsxWithBox(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "dup-jsx" },
            createElement(Notes, null, createElement(P, null, "one")),
            createElement(Notes, null, createElement(P, null, "two")),
          ),
        ),
      ),
    );
    assert.deepEqual(
      dupJsxDiags.map(({ code }) => code),
      ["duplicate-notes"],
    );

    const badChildDiags = await expectBuildErrorCode(() =>
      previewDocTrimmed(
        deck().append(
          slide({ id: "bad-child" }).append(notes(text("illegal"))),
        ),
      ),
    );
    assert.equal(badChildDiags[0]?.code, "unsupported-text-child");

    const unknownPropDiags = await expectBuildErrorCode(() =>
      previewJsxTrimmed(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "unknown-prop" },
            createElement(
              Notes as unknown as ComponentType<Record<string, unknown>>,
              { layout: { x: 1 } },
              createElement(P, null, "bad"),
            ),
          ),
        ),
      ),
    );
    assert.equal(unknownPropDiags[0]?.code, "unknown-parameter");

    const missingHiddenDocDiags = await expectBuildErrorCode(() =>
      previewWithoutNotes(
        deck().append(
          slide({ id: "missing-hidden" }).append(
            box({ style: { display: "none" } }).append(notes("hidden")),
          ),
        ),
      ),
    );
    assert.deepEqual(
      missingHiddenDocDiags.map(({ code, path }) => ({ code, path })),
      [
        {
          code: "missing-capability",
          path: "deck/slide[missing-hidden]/box[1]/notes[1]",
        },
      ],
    );

    const missingJsxDiags = await expectBuildErrorCode(() =>
      previewJsxWithoutNotes(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "missing-jsx" },
            createElement(Notes, null, createElement(P, null, "trimmed")),
          ),
        ),
      ),
    );
    assert.equal(missingJsxDiags[0]?.code, "missing-capability");

    const dupIdDiags = await expectBuildErrorCode(() =>
      previewDocTrimmed(
        deck().append(
          slide({ id: "dup-id" }).append(
            shape({ id: "same", x: 10, y: 10, width: 40, height: 20 }),
            notes("body", { id: "same" }),
          ),
        ),
      ),
    );
    assert.equal(dupIdDiags[0]?.code, "duplicate-author-id");
  });
});
