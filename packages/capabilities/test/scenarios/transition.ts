import assert from "node:assert/strict";

import {
  BuildError,
  createMorphScope,
  Deck,
  expandDocument,
  Slide,
  targets,
  type BuildOptions,
  type CommonNode,
} from "@react-pptx/core";
import { createResolver } from "@react-pptx/core/build";
import { emit } from "@react-pptx/emit";
import { validatePptx } from "@react-pptx/validation";
import { PNG } from "pngjs";
import { createElement, type ReactNode } from "react";

import { Box } from "../../src/box-elements.ts";
import { Chart } from "../../src/chart-elements.ts";
import {
  box,
  chart,
  deck,
  documentCapabilities,
  group,
  shape,
  slide,
} from "../../src/document.ts";
import { Group } from "../../src/group-elements.ts";
import { Shape } from "../../src/shape-elements.ts";
import { transitions } from "../../src/transition.ts";
import type { CapabilityScenario } from "../acceptance-runner.ts";

function readPixel(
  pngBytes: Uint8Array,
  x: number,
  y: number,
): readonly [number, number, number, number] {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  return [
    decoded.data[offset]!,
    decoded.data[offset + 1]!,
    decoded.data[offset + 2]!,
    decoded.data[offset + 3]!,
  ];
}

export const transitionPositiveScenario: CapabilityScenario = {
  id: "transition-and-morph-adjacent",
  capability: "transition",
  description:
    "Dual-author adjacent-slide transition and Morph scenario covering fade with p14 duration, advanceAfter, p159 morph pairing (moved/recolored/resized + entering), and p14 flip with fallback, while keeping static edit canvases unchanged.",
  visualCheckpoints: [
    "Page 1 (morph-origin): clean off-white canvas (#F8FAFC) with blue rectangle (#2563EB, dark navy border) on the left and dark slate circle (#0F172A) at center-left; fade transition and advanceAfter do not alter static pixels.",
    "Page 2 (morph-target): light slate canvas (#F1F5F9) with emerald rectangle (#059669) on the right, amber circle (#D97706) at bottom-left, and purple entering card (#7C3AED) at top-left; Morph declaration does not alter static edit canvas.",
    "Page 3 (flip-outro): dark navy canvas (#0F172A) with centered amber rectangle (#F59E0B); Flip transition does not alter static edit canvas.",
  ],
  nativeCheckpoints: [
    "Open candidate PPTX in PowerPoint/WPS and run Slide Show: Slide 1 fades in (500ms) and auto-advances after 2000ms.",
    "Verify Slide 1 -> Slide 2 performs Morph: the blue left rectangle smoothly moves/resizes/recolors into the emerald right rectangle, the dark circle morphs into the bottom-left amber circle, and the purple card fades/enters.",
    "Verify Slide 2 -> Slide 3 performs Flip (left) after 2500ms or click, and remains editable/saveable on reopen.",
  ],
  jsx() {
    const scope = createMorphScope();
    const hero = scope.handle("hero-panel");
    const badge = scope.handle("accent-dot");
    const entering = scope.handle("entering-card");

    return createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        {
          id: "morph-origin",
          background: "#F8FAFC",
          transition: transitions.fade({ duration: 500 }),
          advanceAfter: 2000,
        },
        createElement(Shape, {
          id: "hero-before",
          morphKey: hero,
          x: 60,
          y: 110,
          width: 140,
          height: 140,
          fill: "#2563EB",
          style: { borderColor: "#1E3A8A", borderWidth: 2 },
        }),
        createElement(Shape, {
          id: "badge-before",
          morphKey: badge,
          geometry: "ellipse",
          x: 240,
          y: 150,
          width: 60,
          height: 60,
          fill: "#0F172A",
        }),
      ),
      createElement(
        Slide,
        {
          id: "morph-target",
          background: "#F1F5F9",
          transition: transitions.morph({
            duration: 1200,
            by: "object",
            fallback: transitions.fade(),
          }),
          advanceAfter: 2500,
        },
        createElement(Shape, {
          id: "hero-after",
          morphKey: hero,
          x: 360,
          y: 70,
          width: 220,
          height: 160,
          fill: "#059669",
          style: { borderColor: "#064E3B", borderWidth: 2 },
        }),
        createElement(Shape, {
          id: "badge-after",
          morphKey: badge,
          geometry: "ellipse",
          x: 80,
          y: 210,
          width: 90,
          height: 90,
          fill: "#D97706",
        }),
        createElement(Shape, {
          id: "entering-card",
          morphKey: entering,
          x: 80,
          y: 50,
          width: 180,
          height: 80,
          fill: "#7C3AED",
        }),
      ),
      createElement(
        Slide,
        {
          id: "flip-outro",
          background: "#0F172A",
          transition: transitions.flip({
            direction: "left",
            fallback: transitions.push({ from: "right" }),
          }),
        },
        createElement(Shape, {
          id: "outro-card",
          x: 180,
          y: 110,
          width: 280,
          height: 140,
          fill: "#F59E0B",
        }),
      ),
    );
  },
  document() {
    const scope = createMorphScope();
    const hero = scope.handle("hero-panel");
    const badge = scope.handle("accent-dot");
    const entering = scope.handle("entering-card");

    return deck({ width: 640, height: 360 }).append(
      slide({
        id: "morph-origin",
        background: "#F8FAFC",
        transition: transitions.fade({ duration: 500 }),
        advanceAfter: 2000,
      }).append(
        shape({
          id: "hero-before",
          morphKey: hero,
          x: 60,
          y: 110,
          width: 140,
          height: 140,
          fill: "#2563EB",
          style: { borderColor: "#1E3A8A", borderWidth: 2 },
        }),
        shape({
          id: "badge-before",
          morphKey: badge,
          geometry: "ellipse",
          x: 240,
          y: 150,
          width: 60,
          height: 60,
          fill: "#0F172A",
        }),
      ),
      slide({
        id: "morph-target",
        background: "#F1F5F9",
        transition: transitions.morph({
          duration: 1200,
          by: "object",
          fallback: transitions.fade(),
        }),
        advanceAfter: 2500,
      }).append(
        shape({
          id: "hero-after",
          morphKey: hero,
          x: 360,
          y: 70,
          width: 220,
          height: 160,
          fill: "#059669",
          style: { borderColor: "#064E3B", borderWidth: 2 },
        }),
        shape({
          id: "badge-after",
          morphKey: badge,
          geometry: "ellipse",
          x: 80,
          y: 210,
          width: 90,
          height: 90,
          fill: "#D97706",
        }),
        shape({
          id: "entering-card",
          morphKey: entering,
          x: 80,
          y: 50,
          width: 180,
          height: 80,
          fill: "#7C3AED",
        }),
      ),
      slide({
        id: "flip-outro",
        background: "#0F172A",
        transition: transitions.flip({
          direction: "left",
          fallback: transitions.push({ from: "right" }),
        }),
      }).append(
        shape({
          id: "outro-card",
          x: 180,
          y: 110,
          width: 280,
          height: 140,
          fill: "#F59E0B",
        }),
      ),
    );
  },
  async assert({ author, pptx, preview }) {
    assert.equal(preview.pages.length, 3, author);
    for (const page of preview.pages) {
      assert.deepEqual(page.diagnostics, [], `${author} p${page.page}`);
      assert.ok(page.png !== undefined, `${author} p${page.page} png`);
    }

    // Verify static canvas pixel colors match declared positions regardless of transition/advanceAfter.
    assert.deepEqual(
      readPixel(preview.pages[0]!.png!, 130, 180),
      [0x25, 0x63, 0xeb, 255],
      author,
    );
    assert.deepEqual(
      readPixel(preview.pages[0]!.png!, 270, 180),
      [0x0f, 0x17, 0x2a, 255],
      author,
    );
    assert.deepEqual(
      readPixel(preview.pages[1]!.png!, 470, 150),
      [0x05, 0x96, 0x69, 255],
      author,
    );
    assert.deepEqual(
      readPixel(preview.pages[1]!.png!, 125, 255),
      [0xd9, 0x77, 0x06, 255],
      author,
    );
    assert.deepEqual(
      readPixel(preview.pages[1]!.png!, 170, 90),
      [0x7c, 0x3a, 0xed, 255],
      author,
    );
    assert.deepEqual(
      readPixel(preview.pages[2]!.png!, 320, 180),
      [0xf5, 0x9e, 0x0b, 255],
      author,
    );

    const file = await validatePptx(pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(file.errors, [], author);
    assert.deepEqual(file.extensions, ["p14", "p159"], author);

    assert.equal(file.slides[0]?.advanceAfterMs, 2000, author);
    assert.deepEqual(
      file.slides[0]?.transition,
      {
        kind: "fade",
        speed: "fast",
        durationMs: 500,
        fallback: { kind: "fade", speed: "fast" },
      },
      author,
    );

    assert.equal(file.slides[1]?.advanceAfterMs, 2500, author);
    assert.deepEqual(
      file.slides[1]?.transition,
      {
        kind: "morph",
        speed: "slow",
        durationMs: 1200,
        by: "object",
        fallback: { kind: "fade", speed: "slow" },
      },
      author,
    );

    assert.deepEqual(
      file.slides[2]?.transition,
      {
        kind: "flip",
        speed: "slow",
        direction: "l",
        fallback: { kind: "push", speed: "slow", direction: "r" },
      },
      author,
    );

    // Verify paired !!morph names match across Slide 1 & Slide 2 and ordinary objects omit !!.
    assert.equal(
      file.slides[0]?.objects[0]?.name,
      file.slides[1]?.objects[0]?.name,
      author,
    );
    assert.match(file.slides[0]?.objects[0]?.name ?? "", /^!!/u, author);
    assert.equal(
      file.slides[0]?.objects[1]?.name,
      file.slides[1]?.objects[1]?.name,
      author,
    );
    assert.match(file.slides[0]?.objects[1]?.name ?? "", /^!!/u, author);
    assert.match(file.slides[1]?.objects[2]?.name ?? "", /^!!/u, author);
    assert.ok(
      !(file.slides[2]?.objects[0]?.name ?? "").startsWith("!!"),
      author,
    );

    // Baseline-only validation must reject the exported file because p14/p159 extensions are used.
    const baselineCheck = await validatePptx(pptx, {
      allowedExtensions: [],
    });
    assert.ok(
      baselineCheck.errors.some(({ code }) => code === "forbidden-extension"),
      author,
    );
  },
};

type DualAuthorRejectionCase = {
  readonly id: string;
  readonly expectedCode: string;
  readonly options?: BuildOptions;
  jsx(): ReactNode;
  document(): CommonNode;
};

const resolveDefault = createResolver(documentCapabilities);
const resolveOptionsFrom = (options?: BuildOptions) =>
  options?.signal === undefined ? {} : { signal: options.signal };
const runJsxBuild = async (element: ReactNode, options?: BuildOptions) =>
  emit(
    resolveDefault(
      await expandDocument(
        element,
        options?.signal === undefined ? {} : { signal: options.signal },
      ),
      resolveOptionsFrom(options),
    ),
    options,
  );
const runDocBuild = async (document: CommonNode, options?: BuildOptions) =>
  emit(resolveDefault(document, resolveOptionsFrom(options)), options);

export const transitionRejectionCases: readonly DualAuthorRejectionCase[] = [
  {
    id: "morph-with-fallback-rejected-on-baseline-only",
    expectedCode: "forbidden-mechanism",
    options: { target: targets.baselineOnly },
    jsx() {
      return createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "morph-fallback",
          transition: transitions.morph({ fallback: transitions.fade() }),
        }),
      );
    },
    document() {
      return deck().append(
        slide({
          id: "morph-fallback",
          transition: transitions.morph({ fallback: transitions.fade() }),
        }),
      );
    },
  },
  {
    id: "morph-duration-rejected-when-p14-omitted",
    expectedCode: "forbidden-mechanism",
    options: {
      target: {
        ...targets.transitional,
        name: "p159-only",
        extensions: ["p159"],
      },
    },
    jsx() {
      return createElement(
        Deck,
        {},
        createElement(Slide, {
          id: "morph-p14",
          transition: transitions.morph({
            duration: 900,
            fallback: transitions.fade(),
          }),
        }),
      );
    },
    document() {
      return deck().append(
        slide({
          id: "morph-p14",
          transition: transitions.morph({
            duration: 900,
            fallback: transitions.fade(),
          }),
        }),
      );
    },
  },
  {
    id: "cross-object-morph-kind-mismatch-box-vs-chart",
    expectedCode: "morph-object-kind-mismatch",
    jsx() {
      const key = createMorphScope().handle("panel");
      return createElement(
        Deck,
        {},
        createElement(
          Slide,
          { id: "src" },
          createElement(Box, {
            morphKey: key,
            style: { width: 120, height: 80, backgroundColor: "#1D4ED8" },
          }),
        ),
        createElement(
          Slide,
          {
            id: "dst",
            transition: transitions.morph({ fallback: transitions.fade() }),
          },
          createElement(Chart, {
            morphKey: key,
            x: 20,
            y: 20,
            width: 240,
            height: 140,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "S", values: [1] }],
              },
            ],
          }),
        ),
      );
    },
    document() {
      const key = createMorphScope().handle("panel");
      return deck().append(
        slide({ id: "src" }).append(
          box({
            morphKey: key,
            style: { width: 120, height: 80, backgroundColor: "#1D4ED8" },
          }),
        ),
        slide({
          id: "dst",
          transition: transitions.morph({ fallback: transitions.fade() }),
        }).append(
          chart({
            morphKey: key,
            x: 20,
            y: 20,
            width: 240,
            height: 140,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "S", values: [1] }],
              },
            ],
          }),
        ),
      );
    },
  },
  {
    id: "morph-inside-group-rejected",
    expectedCode: "morph-inside-group",
    jsx() {
      const key = createMorphScope().handle("member");
      return createElement(
        Deck,
        {},
        createElement(
          Slide,
          { id: "grouped" },
          createElement(
            Group,
            null,
            createElement(Shape, {
              morphKey: key,
              x: 20,
              y: 20,
              width: 60,
              height: 40,
            }),
          ),
        ),
      );
    },
    document() {
      const key = createMorphScope().handle("member");
      return deck().append(
        slide({ id: "grouped" }).append(
          group().append(
            shape({
              morphKey: key,
              x: 20,
              y: 20,
              width: 60,
              height: 40,
            }),
          ),
        ),
      );
    },
  },
];

/** Executes dual-author negative transition and Morph rejection checks across JSX and document entries. */
export async function runTransitionRejectionScenarios(): Promise<void> {
  for (const testCase of transitionRejectionCases) {
    for (const author of ["jsx", "document"] as const) {
      await assert.rejects(
        author === "jsx"
          ? runJsxBuild(testCase.jsx(), testCase.options)
          : runDocBuild(testCase.document(), testCase.options),
        (error: unknown) => {
          assert.ok(error instanceof BuildError, `${testCase.id} (${author})`);
          assert.ok(
            error.diagnostics.some(
              (item) => item.code === testCase.expectedCode,
            ),
            `${testCase.id} (${author}): expected ${testCase.expectedCode}, got ${JSON.stringify(error.diagnostics)}`,
          );
          return true;
        },
      );
    }
  }
}
