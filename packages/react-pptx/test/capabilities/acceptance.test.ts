import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { slideBackgroundCapability } from "@react-pptx/capabilities";
import { boxCapability } from "@react-pptx/capabilities/box";
import { chartCapability } from "@react-pptx/capabilities/chart";
import { Chart } from "@react-pptx/capabilities/chart-elements";
import {
  box,
  chart,
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
} from "@react-pptx/capabilities/document";
import {
  fade,
  motionCapability,
  observeMotion,
} from "@react-pptx/capabilities/motion";
import { Step } from "@react-pptx/capabilities/motion-elements";
import { notesCapability } from "@react-pptx/capabilities/notes";
import { Notes } from "@react-pptx/capabilities/notes-elements";
import { shapeCapability } from "@react-pptx/capabilities/shape";
import { Shape } from "@react-pptx/capabilities/shape-elements";
import {
  P,
  SlideNumber,
  Span,
  Text,
} from "@react-pptx/capabilities/text-elements";
import {
  BuildError,
  BuildQueryError,
  CapabilityError,
  CommonNode,
  Deck,
  Slide,
  createMorphScope,
  expandDocument,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityModule,
  type CapabilityParseContext,
  type Diagnostic,
  type ObjectCapabilityModule,
  type SlideCapabilityModule,
} from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { parseCapabilityLayoutContract } from "@react-pptx/core/capability";
import { RichTextContainer } from "@react-pptx/core/document";
import { emit, pxToEmu, xmlEscape } from "@react-pptx/emit";
import { paint } from "@react-pptx/preview";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement, type ComponentType, type ReactNode } from "react";
import { describe, expect, it } from "vitest";

function previewDocWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode) => {
    const resolution = resolveWith(document);
    return { resolution, ...emit(resolution), ...paint(resolution) };
  };
}

function previewJsxWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (element: ReactNode) => {
    const resolution = resolveWith(await expandDocument(element));
    return { resolution, ...emit(resolution), ...paint(resolution) };
  };
}

import { runCapabilityScenario } from "./acceptance-runner.ts";
import { boxPositiveScenario, boxRejectionScenario } from "./scenarios/box.ts";
import {
  chartPositiveScenario,
  chartRejectionScenario,
} from "./scenarios/chart.ts";
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
  motionPositiveScenario,
  verifyMotionCollaborationAndRejections,
} from "./scenarios/motion.ts";
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
  tablePositiveScenario,
  tableRejectionScenario,
} from "./scenarios/table.ts";
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
  it("runs Box, Group, Shape, Image, slide-background, Table, Text, Notes, and layout dual-author scenarios through shared runner and collects artifacts", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-cap-accept-"));
    try {
      for (const scenario of [
        boxPositiveScenario,
        boxRejectionScenario,
        chartPositiveScenario,
        chartRejectionScenario,
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
        tablePositiveScenario,
        tableRejectionScenario,
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
    const jsxPreview = previewJsxWith(documentCapabilities);
    const docPreview = previewDocWith(documentCapabilities);
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

  it("runs Motion dual-author choreography scenario, static canvas preview, Group/Chart collaboration, and rejections", async () => {
    const outDir = await mkdtemp(join(tmpdir(), "rp-motion-accept-"));
    try {
      const runs = await runCapabilityScenario(motionPositiveScenario, {
        outputDir: outDir,
      });
      assert.equal(runs.length, 2);
      assert.deepEqual(
        runs.map((r) => r.author),
        ["jsx", "document"],
      );
      await verifyMotionCollaborationAndRejections();
    } finally {
      await rm(outDir, { recursive: true, force: true });
    }
  }, 15_000);

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
          new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
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

    const previewDocument = previewDocWith(documentCapabilities);
    const first = await previewDocument(expanded);
    assert.deepEqual(first.pages[0]?.diagnostics, []);
    assert.deepEqual((await validatePptx(first.pptx)).errors, []);
    const firstPng = first.pages[0].png!.slice();
    const firstSvg = first.pages[0].svg!;
    const firstPptx = first.pptx.slice();
    const firstReportSnapshot = JSON.stringify(first.report);
    const firstHeight = first.query
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
    assert.deepEqual((await validatePptx(second.pptx)).errors, []);
    assert.equal(
      second.query.getSlide("handoff").get("headline").observe().text,
      "Updated longer headline wrapping across multiple lines",
    );
    const secondHeight = second.query
      .getSlide("handoff")
      .get("headline")
      .observe().box!.height;
    assert.ok(secondHeight > firstHeight);
    expect(second.pages[0]?.svg).toContain('fill="#059669"');
    assert.notDeepEqual(second.pages[0]?.png, firstPng);

    // Prior BuildResult and PreviewResult remain unchanged.
    assert.equal(first.pages[0]?.svg, firstSvg);
    assert.deepEqual(first.pages[0]?.png, firstPng);
    assert.deepEqual(first.pptx, firstPptx);
    assert.equal(JSON.stringify(first.report), firstReportSnapshot);
    assert.equal(
      first.query.getSlide("handoff").get("headline").observe().text,
      "Initial title",
    );

    // Cross-build motion observation references are rejected cleanly.
    const motionDoc = deck({ width: 480, height: 270 }).append(
      slide({ id: "m" }).append(
        shape({
          id: "hero",
          x: 20,
          y: 20,
          width: 80,
          height: 40,
          enter: fade({ duration: 200 }),
        }),
      ),
    );
    const motionBuildA = await previewDocument(motionDoc);
    const motionBuildB = await previewDocument(motionDoc);
    const effectA = observeMotion(motionBuildA.report.slides[0]!)!.steps[0]!
      .effects[0]!;
    const stepB = observeMotion(motionBuildB.report.slides[0]!)!.steps[0]!;
    assert.throws(
      () => stepB.before(effectA.start.ref, effectA.end.ref),
      (error: unknown) =>
        error instanceof BuildQueryError && error.code === "foreign-reference",
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
          resolve(resolveContext) {
            calls.push(`resolve:${resolveContext.path}`);
            const box = resolveContext.layoutBox ?? {
              x,
              y,
              width: size,
              height: size,
            };
            const name = xmlEscape(`Marker ${resolveContext.id}`);
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name,
                geometry: "ellipse",
                box,
              },
              extensions: [],
              emit(emitContext) {
                const xml = `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`;
                return { xml };
              },
              paint(paintContext) {
                calls.push(`paint:${paintContext.path}`);
                const resolved = paintContext.theme.resolveColor(color);
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
          resolve(resolveContext) {
            const box = resolveContext.layoutBox ?? {
              x: 0,
              y: 0,
              width: 20,
              height: 20,
            };
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: "NoPreview",
                geometry: "rect",
                box,
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="NoPreview"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const previewCustom = previewDocWith([
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
      result.query.getSlide("warn-keeps-png").get("dot").observe().box
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

    // Same custom bound extension works via JSX expansion and shares trimmed constraints.
    const jsxCustomPreview = previewJsxWith([shapeCapability]);
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
    assert.ok(calls.some((entry) => entry.startsWith("resolve:")));
    assert.ok(calls.some((entry) => entry.startsWith("paint:")));
  });

  it("binds notesCapability via BoundSlideCapabilityModule, trims textCapability cleanly, and enforces rejections", async () => {
    assert.equal("bindSlide" in notesCapability, true);
    assert.equal("measureLayout" in notesCapability, false);

    // Notes-only slide previews cleanly as an empty canvas without preview errors or leaked notes text.
    const notesOnlyPreview = previewDocWith([
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
    const previewDocTrimmed = previewDocWith(trimmedModules);
    const previewJsxTrimmed = previewJsxWith(trimmedModules);

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
    assert.deepEqual(docResult.report.slides[0]?.capabilities, {
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
    assert.deepEqual(jsxResult.report.slides[0]?.capabilities, {
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
    const previewDocWithBox = previewDocWith(withBoxModules);
    const previewJsxWithBox = previewJsxWith(withBoxModules);
    const previewWithoutNotes = previewDocWith([
      shapeCapability,
      boxCapability,
    ]);
    const previewJsxWithoutNotes = previewJsxWith([
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

  it("rejects non-bound capability modules at registration and invalid bound handles at bind time", async () => {
    const unboundObjectModule = {
      declaration: "rp:unbound-object",
      parse() {
        return {
          contract: { nativeKind: "sp", layoutStyle: {} },
          parsed: {},
        };
      },
      resolve() {
        return {
          observation: { id: 2, kind: "shape", name: "Unbound" },
          extensions: [],
          emit: () => ({ xml: "" }),
        };
      },
    } as unknown as CapabilityModule;

    const unboundSlideModule = {
      declaration: "unbound-slide",
      slideProp: "unboundSlide",
      slot: "timing",
      resolveSlide() {
        return { extensions: [], emit: () => ({ xml: "" }) };
      },
    } as unknown as CapabilityModule;

    for (const invalidModule of [unboundObjectModule, unboundSlideModule]) {
      assert.throws(
        () => createResolver([invalidModule]),
        (error: unknown) => {
          assert(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-capability-module");
          return true;
        },
      );
    }

    const brokenObjectBinding: ObjectCapabilityModule = {
      declaration: "rp:broken-object",
      bind() {
        return {
          contract: { nativeKind: "sp" },
        } as unknown as ReturnType<ObjectCapabilityModule["bind"]>;
      },
    };
    const brokenObjectResolve = createResolver([brokenObjectBinding]);
    assert.throws(
      () =>
        brokenObjectResolve(
          deck().append(
            slide({ id: "s1" }).append(
              new CommonNode("broken", "rp:broken-object", { id: "bad-obj" }),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-capability-binding");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[s1]/broken-object[bad-obj]",
        );
        return true;
      },
    );

    const brokenSlideBinding: SlideCapabilityModule = {
      declaration: "broken-slide",
      slideProp: "background",
      slot: "background",
      bindSlide() {
        return {} as unknown as ReturnType<SlideCapabilityModule["bindSlide"]>;
      },
    };
    const brokenSlideResolve = createResolver([brokenSlideBinding]);
    assert.throws(
      () =>
        brokenSlideResolve(
          deck().append(slide({ id: "s1", background: "#123456" })),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-capability-binding");
        assert.equal(error.diagnostics[0]?.path, "deck/slide[s1]");
        return true;
      },
    );
  });

  it("delegates transparent declaration prop validation (including morphKey) to compileDeclaration without core prop readbacks", async () => {
    const morphKey = createMorphScope().handle("step-key");
    const stepBuild = previewJsxWith([motionCapability, shapeCapability]);
    await assert.rejects(
      stepBuild(
        createElement(
          Deck,
          null,
          createElement(
            Slide,
            { id: "transparent-own-validation" },
            createElement(
              Step as unknown as ComponentType<Record<string, unknown>>,
              { n: 1, morphKey },
              createElement(Shape, { x: 0, y: 0, width: 40, height: 20 }),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.deepEqual(
          error.diagnostics.map(({ code }) => code),
          ["unknown-parameter"],
        );
        return true;
      },
    );
  });

  it("supports custom bound SlideCapabilityModule with paint, rejects missing background paint, and enforces [slideBackgroundCapability, chartCapability] trimming", async () => {
    const customSlideCalls: string[] = [];
    const customBackdropModule: SlideCapabilityModule = {
      declaration: "custom-backdrop",
      slideProp: "background",
      slot: "background",
      bindSlide(input, parseContext) {
        const raw = input.props["background"];
        if (typeof raw !== "string" || !/^#[0-9A-Fa-f]{6}$/u.test(raw)) {
          throw new CapabilityError(
            "invalid-background",
            "Custom backdrop expects a 6-digit hex color",
            parseContext.path,
          );
        }
        const hex = raw.slice(1).toUpperCase();
        customSlideCalls.push(`bindSlide:${parseContext.path}`);
        return {
          resolveSlide(resolveContext) {
            customSlideCalls.push(`resolveSlide:${resolveContext.path}`);
            return {
              extensions: [],
              observation: { backdropColor: `#${hex}` },
              emit() {
                return {
                  xml: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
                };
              },
              paint(paintContext) {
                customSlideCalls.push(`paint:${paintContext.path}`);
                return {
                  svg: `<rect width="${paintContext.dimensions.width}" height="${paintContext.dimensions.height}" fill="#${hex}"/>`,
                };
              },
            };
          },
        };
      },
    };

    const backdropDocPreview = previewDocWith([
      customBackdropModule,
      shapeCapability,
    ]);
    const backdropJsxPreview = previewJsxWith([
      customBackdropModule,
      shapeCapability,
    ]);

    const docResult = await backdropDocPreview(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "bg-doc", background: "#0F766E" }).append(
          shape({
            id: "card",
            x: 40,
            y: 20,
            width: 120,
            height: 80,
            fill: "#F59E0B",
          }),
        ),
      ),
    );
    const jsxResult = await backdropJsxPreview(
      createElement(
        Deck,
        { width: 200, height: 120 },
        createElement(
          Slide,
          { id: "bg-jsx", background: "#0F766E" },
          createElement(Shape, {
            id: "card",
            x: 40,
            y: 20,
            width: 120,
            height: 80,
            fill: "#F59E0B",
          }),
        ),
      ),
    );

    for (const result of [docResult, jsxResult]) {
      assert.deepEqual(result.pages[0]?.diagnostics, []);
      assert.deepEqual((await validatePptx(result.pptx)).errors, []);
      assert.deepEqual(result.report.slides[0]?.capabilities, {
        "custom-backdrop": { backdropColor: "#0F766E" },
      });
      const slideXml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      expect(slideXml).toContain('<a:srgbClr val="0F766E"/>');
      const png = PNG.sync.read(Buffer.from(result.pages[0].png!));
      // Top-left corner (5, 5) shows custom backdrop #0F766E; center (100, 60) shows shape #F59E0B.
      const cornerOffset = (5 * png.width + 5) * 4;
      assert.deepEqual(
        [...png.data.subarray(cornerOffset, cornerOffset + 4)],
        [0x0f, 0x76, 0x6e, 255],
      );
      const centerOffset = (60 * png.width + 100) * 4;
      assert.deepEqual(
        [...png.data.subarray(centerOffset, centerOffset + 4)],
        [0xf5, 0x9e, 0x0b, 255],
      );
    }
    assert.deepEqual(customSlideCalls, [
      "bindSlide:deck/slide[bg-doc]",
      "resolveSlide:deck/slide[bg-doc]",
      "paint:deck/slide[bg-doc]",
      "bindSlide:deck/slide[bg-jsx]",
      "resolveSlide:deck/slide[bg-jsx]",
      "paint:deck/slide[bg-jsx]",
    ]);

    // A background-slot SlideCapabilityModule without paint produces png: undefined + located error.
    const backdropWithoutPreview: SlideCapabilityModule = {
      declaration: "custom-backdrop-no-preview",
      slideProp: "background",
      slot: "background",
      bindSlide() {
        return {
          resolveSlide() {
            return {
              extensions: [],
              emit() {
                return {
                  xml: `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,
                };
              },
            };
          },
        };
      },
    };
    const noPreviewResult = await previewDocWith([backdropWithoutPreview])(
      deck({ width: 200, height: 120 }).append(
        slide({ id: "no-bg-prev", background: "#1E293B" }),
      ),
    );
    assert.equal(noPreviewResult.pages[0]?.png, undefined);
    assert.deepEqual(noPreviewResult.pages[0]?.diagnostics, [
      {
        severity: "error",
        page: 1,
        path: "deck/slide[no-bg-prev]",
        message: "Preview does not support custom-backdrop-no-preview",
      },
    ]);

    // Trimmed [slideBackgroundCapability, chartCapability] builds & previews chart without shape/text, and rejects Shape.
    const chartFont = font("Jost*", [
      {
        weight: 400,
        data: readFileSync(
          new URL("../assets/fonts/Jost-400-Book.ttf", import.meta.url),
        ),
      },
    ]);
    const chartOnlyModules = [
      slideBackgroundCapability,
      chartCapability,
    ] as const;
    const chartTrimmedDoc = previewDocWith(chartOnlyModules);
    const chartTrimmedJsx = previewJsxWith(chartOnlyModules);
    const chartGroup = {
      kind: "column" as const,
      data: { kind: "shared-categories" as const, categories: ["Q1", "Q2"] },
      series: [{ name: "Sales", values: [12, 24] }],
    };
    const chartDocOut = await chartTrimmedDoc(
      deck({ width: 320, height: 200 }).append(
        slide({ id: "chart-trim", background: "#FFFFFF" }).append(
          chart({
            id: "c1",
            x: 20,
            y: 20,
            width: 280,
            height: 160,
            style: { fontFamily: chartFont },
            groups: [chartGroup],
          }),
        ),
      ),
    );
    const chartJsxOut = await chartTrimmedJsx(
      createElement(
        Deck,
        { width: 320, height: 200 },
        createElement(
          Slide,
          { id: "chart-trim", background: "#FFFFFF" },
          createElement(Chart, {
            id: "c1",
            x: 20,
            y: 20,
            width: 280,
            height: 160,
            style: { fontFamily: chartFont },
            groups: [chartGroup],
          }),
        ),
      ),
    );
    assert.ok(chartDocOut.pages[0]?.png !== undefined);
    assert.ok(chartJsxOut.pages[0]?.png !== undefined);
    assert.deepEqual((await validatePptx(chartDocOut.pptx)).errors, []);
    assert.deepEqual((await validatePptx(chartJsxOut.pptx)).errors, []);
    await assert.rejects(
      () =>
        chartTrimmedDoc(
          deck({ width: 320, height: 200 }).append(
            slide({ id: "missing-shape" }).append(
              shape({ id: "s1", x: 10, y: 10, width: 50, height: 50 }),
            ),
          ),
        ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        return true;
      },
    );
  });
});
