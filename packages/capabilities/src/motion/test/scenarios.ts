import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { documentCapabilities } from "@pptxdk/capabilities/all";
import { Box } from "@pptxdk/capabilities/box-elements";
import { Chart } from "@pptxdk/capabilities/chart-elements";
import {
  box,
  chart,
  createMotionScopeHandle,
  deck,
  font,
  group,
  motionScope,
  section,
  shape,
  slide,
  step,
  text,
} from "@pptxdk/capabilities/document";
import { Group } from "@pptxdk/capabilities/group-elements";
import {
  after,
  breathe,
  fade,
  fillColor,
  floatUp,
  fly,
  motionSchedule,
  repeat,
  spin,
  wipe,
  zoom,
  type MotionObservation,
} from "@pptxdk/capabilities/motion";
import {
  MotionScope,
  Step,
  useMotionScope,
} from "@pptxdk/capabilities/motion-elements";
import { Shape } from "@pptxdk/capabilities/shape-elements";
import { Text } from "@pptxdk/capabilities/text-elements";
import {
  BuildError,
  targets,
  type BuildOptions,
  type CommonNode,
} from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { Deck, expandDocument, Section, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import { PNG } from "pngjs";
import { createElement, type ReactNode } from "react";

import type { CapabilityScenario } from "../../../test/acceptance-runner.ts";

const jostFont = font("Jost*", [
  {
    weight: 400,
    data: readFileSync(
      new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
    ),
  },
]);

function pixelRgba(
  pngBytes: Uint8Array,
  x: number,
  y: number,
): readonly [number, number, number, number] {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const px = Math.round(x);
  const py = Math.round(y);
  const offset = (py * decoded.width + px) * 4;
  return [
    decoded.data[offset]!,
    decoded.data[offset + 1]!,
    decoded.data[offset + 2]!,
    decoded.data[offset + 3]!,
  ];
}

function slideXml(pptx: Uint8Array, slideNumber: number): string {
  const entry = unzipSync(pptx)[`ppt/slides/slide${slideNumber}.xml`];
  assert.ok(entry !== undefined, `missing ppt/slides/slide${slideNumber}.xml`);
  return strFromU8(entry);
}

export const motionPositiveScenario: CapabilityScenario = {
  id: "motion-choreography-and-static-canvas",
  capability: "motion",
  description:
    "Dual-author Motion scenario covering enter/emphasis effects, Step click groups, MotionScope sequence/stagger/after(cue) timing, container default entrance & motion={false} overrides, continuous breathe, static canvas visibility in PNGs, and real PPTX <p:timing> output.",
  visualCheckpoints: [
    "Page 1 (motion-choreography): light slate #F8FAFC canvas; all objects with enter animations remain statically visible on the edit canvas — blue #2563EB hero-badge at top-left, #EFF6FF bordered Text 'Step 1: Sequence Source' at top-center, emerald #10B981 follower-card at top-right, and in the second row amber #F59E0B, purple #8B5CF6, white bordered 'Step 2: Own Zoom', and muted #E2E8F0 box framing slate #64748B shape.",
    "Page 2 (motion-auto-and-continuous): white canvas; sky #0EA5E9 and teal #14B8A6 auto-staggered rectangles at top-left, #F1F5F9 blue-bordered 'Auto after stagger' text at top-right, and dark green #164E3D continuous breathe lamp at bottom-left all visible statically.",
  ],
  nativeCheckpoints: [
    "Open exported PPTX in PowerPoint/WPS slideshow: Page 1 starts blank for enter-animated objects, Click 1 sequences hero-badge (fade + 2x breathe) then hero-title (fly from left) followed 80ms later by follower-card (wipe + fillColor); Click 2 floats up amber and purple shapes together with own-entry zoom text while muted-box and muted-shape remain static.",
    "On Page 2, auto-a and auto-b stagger automatically on slide entry (100ms apart) followed 50ms later by auto-summary, while pulse-lamp breathes continuously until slide exit.",
  ],
  jsx() {
    function MotionDeck() {
      const source = useMotionScope();
      const follower = useMotionScope();
      const body = useMotionScope();
      const afterBody = useMotionScope();
      const lamp = useMotionScope();
      const ready = source.cue("ready");
      const bodyDone = body.cue("done");

      return createElement(
        Deck,
        { width: 640, height: 360 },
        createElement(
          Slide,
          { id: "motion-choreography", background: "#F8FAFC" },
          createElement(
            Step,
            { n: 1 },
            createElement(
              MotionScope,
              { scope: source, mode: "sequence", publishEnd: ready },
              createElement(Shape, {
                id: "hero-badge",
                enter: fade({ duration: 240 }),
                emphasis: repeat(
                  breathe({ period: 800, from: "#2563EB", to: "#60A5FA" }),
                  2,
                ),
                style: {
                  position: "absolute",
                  left: 32,
                  top: 36,
                  width: 120,
                  height: 68,
                  background: "#2563EB",
                },
              }),
              createElement(
                Text,
                {
                  id: "hero-title",
                  enter: fly({ from: "left", duration: 220 }),
                  style: {
                    position: "absolute",
                    left: 172,
                    top: 36,
                    width: 220,
                    height: 68,
                    fontFamily: jostFont,
                    fontSize: 16,
                    color: "#0F172A",
                    verticalAlign: "middle",
                    padding: 10,
                    background: "#EFF6FF",
                    border: { color: "#3B82F6", width: 2 },
                  },
                },
                "Step 1: Sequence Source",
              ),
            ),
            createElement(
              MotionScope,
              { scope: follower, startAt: after(ready, 80) },
              createElement(Shape, {
                id: "follower-card",
                enter: wipe({ from: "top", duration: 200 }),
                emphasis: fillColor({
                  from: "#10B981",
                  to: "#059669",
                  duration: 300,
                }),
                style: {
                  position: "absolute",
                  left: 412,
                  top: 36,
                  width: 196,
                  height: 68,
                  background: "#10B981",
                },
              }),
            ),
          ),
          createElement(
            Step,
            { n: 2 },
            createElement(
              Box,
              {
                id: "stagger-box",
                enter: floatUp({ duration: 180 }),
                style: {
                  position: "absolute",
                  left: 32,
                  top: 136,
                  width: 576,
                  display: "flex",
                  flexDirection: "row",
                  gap: 16,
                },
              },
              createElement(Shape, {
                id: "inherited-a",
                style: { width: 130, height: 76, background: "#F59E0B" },
              }),
              createElement(Shape, {
                id: "inherited-b",
                style: { width: 130, height: 76, background: "#8B5CF6" },
              }),
              createElement(
                Text,
                {
                  id: "own-entry-note",
                  enter: zoom({ duration: 160 }),
                  style: {
                    width: 170,
                    height: 76,
                    fontFamily: jostFont,
                    fontSize: 15,
                    color: "#0F172A",
                    verticalAlign: "middle",
                    padding: 10,
                    background: "#FFFFFF",
                    border: { color: "#CBD5E1", width: 1 },
                  },
                },
                "Step 2: Own Zoom",
              ),
              createElement(
                Box,
                {
                  id: "muted-box",
                  motion: false,
                  enter: fade(),
                  style: {
                    width: 98,
                    height: 76,
                    padding: 8,
                    background: "#E2E8F0",
                  },
                },
                createElement(Shape, {
                  id: "muted-shape",
                  enter: fade(),
                  style: { width: 82, height: 60, background: "#64748B" },
                }),
              ),
            ),
          ),
        ),
        createElement(
          Slide,
          { id: "motion-auto-and-continuous", background: "#FFFFFF" },
          createElement(
            MotionScope,
            {
              scope: body,
              mode: "stagger",
              each: 100,
              publishEnd: bodyDone,
            },
            createElement(Shape, {
              id: "auto-a",
              enter: fade({ duration: 200 }),
              style: {
                position: "absolute",
                left: 36,
                top: 44,
                width: 150,
                height: 80,
                background: "#0EA5E9",
              },
            }),
            createElement(Shape, {
              id: "auto-b",
              enter: fade({ duration: 200 }),
              style: {
                position: "absolute",
                left: 206,
                top: 44,
                width: 150,
                height: 80,
                background: "#14B8A6",
              },
            }),
          ),
          createElement(
            MotionScope,
            { scope: afterBody, startAt: after(bodyDone, 50) },
            createElement(
              Text,
              {
                id: "auto-summary",
                enter: wipe({ from: "left", duration: 180 }),
                style: {
                  position: "absolute",
                  left: 376,
                  top: 44,
                  width: 228,
                  height: 80,
                  fontFamily: jostFont,
                  fontSize: 16,
                  color: "#0F172A",
                  verticalAlign: "middle",
                  padding: 10,
                  background: "#F1F5F9",
                  border: { color: "#0284C7", width: 2 },
                },
              },
              "Auto after stagger",
            ),
          ),
          createElement(
            MotionScope,
            { scope: lamp },
            createElement(Shape, {
              id: "pulse-lamp",
              emphasis: repeat(
                breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
                "forever",
              ),
              style: {
                position: "absolute",
                left: 36,
                top: 160,
                width: 240,
                height: 80,
                background: "#164E3D",
              },
            }),
          ),
        ),
      );
    }

    return createElement(MotionDeck);
  },
  document() {
    const source = createMotionScopeHandle();
    const follower = createMotionScopeHandle();
    const body = createMotionScopeHandle();
    const afterBody = createMotionScopeHandle();
    const lamp = createMotionScopeHandle();
    const ready = source.cue("ready");
    const bodyDone = body.cue("done");

    return deck({ width: 640, height: 360 }).append(
      slide({ id: "motion-choreography", background: "#F8FAFC" }).append(
        step({ n: 1 }).append(
          motionScope(source, { mode: "sequence", publishEnd: ready }).append(
            shape({
              id: "hero-badge",
              enter: fade({ duration: 240 }),
              emphasis: repeat(
                breathe({ period: 800, from: "#2563EB", to: "#60A5FA" }),
                2,
              ),
              style: {
                position: "absolute",
                left: 32,
                top: 36,
                width: 120,
                height: 68,
                background: "#2563EB",
              },
            }),
            text("Step 1: Sequence Source", {
              id: "hero-title",
              enter: fly({ from: "left", duration: 220 }),
              style: {
                position: "absolute",
                left: 172,
                top: 36,
                width: 220,
                height: 68,
                fontFamily: jostFont,
                fontSize: 16,
                color: "#0F172A",
                verticalAlign: "middle",
                padding: 10,
                background: "#EFF6FF",
                border: { color: "#3B82F6", width: 2 },
              },
            }),
          ),
          motionScope(follower, { startAt: after(ready, 80) }).append(
            shape({
              id: "follower-card",
              enter: wipe({ from: "top", duration: 200 }),
              emphasis: fillColor({
                from: "#10B981",
                to: "#059669",
                duration: 300,
              }),
              style: {
                position: "absolute",
                left: 412,
                top: 36,
                width: 196,
                height: 68,
                background: "#10B981",
              },
            }),
          ),
        ),
        step({ n: 2 }).append(
          box({
            id: "stagger-box",
            enter: floatUp({ duration: 180 }),
            style: {
              position: "absolute",
              left: 32,
              top: 136,
              width: 576,
              display: "flex",
              flexDirection: "row",
              gap: 16,
            },
          }).append(
            shape({
              id: "inherited-a",
              style: { width: 130, height: 76, background: "#F59E0B" },
            }),
            shape({
              id: "inherited-b",
              style: { width: 130, height: 76, background: "#8B5CF6" },
            }),
            text("Step 2: Own Zoom", {
              id: "own-entry-note",
              enter: zoom({ duration: 160 }),
              style: {
                width: 170,
                height: 76,
                fontFamily: jostFont,
                fontSize: 15,
                color: "#0F172A",
                verticalAlign: "middle",
                padding: 10,
                background: "#FFFFFF",
                border: { color: "#CBD5E1", width: 1 },
              },
            }),
            box({
              id: "muted-box",
              motion: false,
              enter: fade(),
              style: {
                width: 98,
                height: 76,
                padding: 8,
                background: "#E2E8F0",
              },
            }).append(
              shape({
                id: "muted-shape",
                enter: fade(),
                style: { width: 82, height: 60, background: "#64748B" },
              }),
            ),
          ),
        ),
      ),
      slide({ id: "motion-auto-and-continuous", background: "#FFFFFF" }).append(
        motionScope(body, {
          mode: "stagger",
          each: 100,
          publishEnd: bodyDone,
        }).append(
          shape({
            id: "auto-a",
            enter: fade({ duration: 200 }),
            style: {
              position: "absolute",
              left: 36,
              top: 44,
              width: 150,
              height: 80,
              background: "#0EA5E9",
            },
          }),
          shape({
            id: "auto-b",
            enter: fade({ duration: 200 }),
            style: {
              position: "absolute",
              left: 206,
              top: 44,
              width: 150,
              height: 80,
              background: "#14B8A6",
            },
          }),
        ),
        motionScope(afterBody, { startAt: after(bodyDone, 50) }).append(
          text("Auto after stagger", {
            id: "auto-summary",
            enter: wipe({ from: "left", duration: 180 }),
            style: {
              position: "absolute",
              left: 376,
              top: 44,
              width: 228,
              height: 80,
              fontFamily: jostFont,
              fontSize: 16,
              color: "#0F172A",
              verticalAlign: "middle",
              padding: 10,
              background: "#F1F5F9",
              border: { color: "#0284C7", width: 2 },
            },
          }),
        ),
        motionScope(lamp).append(
          shape({
            id: "pulse-lamp",
            emphasis: repeat(
              breathe({ period: 1600, from: "#164E3D", to: "#4ADE80" }),
              "forever",
            ),
            style: {
              position: "absolute",
              left: 36,
              top: 160,
              width: 240,
              height: 80,
              background: "#164E3D",
            },
          }),
        ),
      ),
    );
  },
  assert({ author, resolution, pptx, preview, validation }) {
    assert.deepEqual(validation.extensions, [], author);
    const p1 = preview.pages[0]!;
    const p2 = preview.pages[1]!;
    assert.deepEqual(p1.diagnostics, [], author);
    assert.deepEqual(p2.diagnostics, [], author);
    assert.ok(p1.png !== undefined, author);
    assert.ok(p2.png !== undefined, author);

    // Static edit canvas keeps enter-animated objects visibly rendered.
    assert.deepEqual(
      pixelRgba(p1.png, 92, 70),
      [0x25, 0x63, 0xeb, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p1.png, 510, 70),
      [0x10, 0xb9, 0x81, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p1.png, 97, 174),
      [0xf5, 0x9e, 0x0b, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p1.png, 243, 174),
      [0x8b, 0x5c, 0xf6, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p1.png, 559, 174),
      [0x64, 0x74, 0x8b, 255],
      author,
    );

    assert.deepEqual(
      pixelRgba(p2.png, 111, 84),
      [0x0e, 0xa5, 0xe9, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p2.png, 281, 84),
      [0x14, 0xb8, 0xa6, 255],
      author,
    );
    assert.deepEqual(
      pixelRgba(p2.png, 156, 200),
      [0x16, 0x4e, 0x3d, 255],
      author,
    );

    const slide1 = resolution.report.slides[0]!;
    const slide2 = resolution.report.slides[1]!;
    const motion1 = slide1.capabilities?.["motion"] as MotionObservation;
    const motion2 = slide2.capabilities?.["motion"] as MotionObservation;
    assert.ok(motion1 !== undefined && motion2 !== undefined, author);

    assert.deepEqual(
      motion1.steps.map((s) => [
        s.trigger,
        s.number,
        s.effects.map((e) => [e.target.authorId, e.preset, e.startMs, e.endMs]),
      ]),
      [
        [
          "click",
          1,
          [
            ["hero-badge", "fade", 0, 240],
            ["hero-badge", "breathe", 0, 1600],
            ["hero-title", "fly", 1600, 1820],
            ["follower-card", "wipe", 1900, 2100],
            ["follower-card", "fillColor", 1900, 2200],
          ],
        ],
        [
          "click",
          2,
          [
            ["inherited-a", "floatUp", 0, 180],
            ["inherited-b", "floatUp", 0, 180],
            ["own-entry-note", "zoom", 0, 160],
          ],
        ],
      ],
      author,
    );

    const motionQuery1 = motionSchedule(slide1);
    assert.ok(motionQuery1 !== undefined, author);
    assert.equal(motionQuery1.steps.length, 2, author);
    assert.equal(motionQuery1.steps[0]?.cues[0]?.name, "ready", author);

    assert.equal(motion2.steps.length, 1, author);
    assert.equal(motion2.steps[0]?.trigger, "auto", author);
    assert.deepEqual(
      motion2.steps[0]?.effects.map((e) => [
        e.target.authorId,
        e.preset,
        e.startMs,
        e.endMs,
        e.completion,
      ]),
      [
        ["auto-a", "fade", 0, 200, undefined],
        ["pulse-lamp", "breathe", 0, undefined, "no-natural-completion"],
        ["auto-b", "fade", 100, 300, undefined],
        ["auto-summary", "wipe", 350, 530, undefined],
      ],
      author,
    );

    // Independent PPTX structure & target inspection.
    const fileSlide1 = validation.slides[0]!;
    const fileSlide2 = validation.slides[1]!;
    assert.equal(fileSlide1.animations?.length, 2, author);
    assert.deepEqual(
      fileSlide1.animations?.map((stepItem) => stepItem.trigger),
      ["click", "click"],
      author,
    );
    assert.equal(fileSlide2.animations?.[0]?.trigger, "auto", author);

    const xml1 = slideXml(pptx, 1);
    const xml2 = slideXml(pptx, 2);
    assert.match(xml1, /<p:timing>/u, author);
    assert.match(xml1, /<p:bldLst>/u, author);
    assert.match(xml2, /<p:timing>/u, author);
    assert.match(xml2, /repeatCount="indefinite"/u, author);
  },
};

export async function verifyMotionCollaborationAndRejections(): Promise<void> {
  const resolveDefault = createResolver(documentCapabilities);
  const resolveOptionsFrom = (options?: BuildOptions) =>
    options?.signal === undefined ? {} : { signal: options.signal };
  const buildJsx = async (element: ReactNode, options?: BuildOptions) =>
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
  const buildDocument = async (document: CommonNode, options?: BuildOptions) =>
    emit(resolveDefault(document, resolveOptionsFrom(options)), options);

  // 1. Preserve 100% Group and Chart native motion target collaboration via formal build + validatePptx.
  const chartGroups = [
    {
      kind: "column" as const,
      data: {
        kind: "shared-categories" as const,
        categories: ["Q1", "Q2", "Q3"],
      },
      series: [
        { name: "Baseline", values: [12, 18, 24] },
        { name: "Actual", values: [15, 21, 29] },
      ],
    },
  ];

  const jsxCollab = await buildJsx(
    createElement(
      Deck,
      { width: 640, height: 360 },
      createElement(
        Slide,
        { id: "group-and-chart-motion" },
        createElement(
          Step,
          { n: 1 },
          createElement(
            Group,
            { id: "spin-group", emphasis: spin({ by: 90, duration: 400 }) },
            createElement(Shape, {
              id: "group-chip",
              enter: fade({ duration: 150 }),
              style: {
                position: "absolute",
                left: 32,
                top: 40,
                width: 90,
                height: 48,
                background: "#2563EB",
              },
            }),
            createElement(
              Text,
              {
                id: "group-label",
                enter: fade({ duration: 150 }),
                style: {
                  position: "absolute",
                  left: 32,
                  top: 96,
                  width: 90,
                  height: 32,
                  fontFamily: jostFont,
                  fontSize: 13,
                },
              },
              "Rotated",
            ),
          ),
          createElement(Chart, {
            id: "animated-chart",
            groups: chartGroups,
            enter: fade({ duration: 180, by: "series" }),
            style: {
              position: "absolute",
              left: 180,
              top: 36,
              width: 420,
              height: 260,
            },
          }),
        ),
      ),
    ),
  );

  const docCollab = await buildDocument(
    deck({ width: 640, height: 360 }).append(
      slide({ id: "group-and-chart-motion" }).append(
        step({ n: 1 }).append(
          group({
            id: "spin-group",
            emphasis: spin({ by: 90, duration: 400 }),
          }).append(
            shape({
              id: "group-chip",
              enter: fade({ duration: 150 }),
              style: {
                position: "absolute",
                left: 32,
                top: 40,
                width: 90,
                height: 48,
                background: "#2563EB",
              },
            }),
            text("Rotated", {
              id: "group-label",
              enter: fade({ duration: 150 }),
              style: {
                position: "absolute",
                left: 32,
                top: 96,
                width: 90,
                height: 32,
                fontFamily: jostFont,
                fontSize: 13,
              },
            }),
          ),
          chart({
            id: "animated-chart",
            groups: chartGroups,
            enter: fade({ duration: 180, by: "series" }),
            style: {
              position: "absolute",
              left: 180,
              top: 36,
              width: 420,
              height: 260,
            },
          }),
        ),
      ),
    ),
  );

  for (const [author, result] of [
    ["jsx", jsxCollab],
    ["document", docCollab],
  ] as const) {
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, [], author);
    const motion = result.report.slides[0]?.capabilities?.[
      "motion"
    ] as MotionObservation;
    assert.equal(motion.nativeProjections?.length, 1, author);
    assert.equal(
      motion.nativeProjections?.[0]?.nativeEffects.length,
      2,
      author,
    );
    assert.equal(
      motion.steps[0]?.effects.filter(
        (e) => e.target.authorId === "animated-chart",
      ).length,
      3,
      author,
    );
    const xml = slideXml(result.pptx, 1);
    assert.match(xml, /<p:bldGraphic/u, author);
    assert.match(xml, /<a:bldChart bld="series"/u, author);
    assert.match(xml, /<p:animRot/u, author);
  }

  // 2. Rejections across missing target/cue, invalid combinations, granular chart constraints, and output target limits.
  const expectBuildError = async (
    label: string,
    run: () => Promise<unknown>,
    expectedCodes: readonly string[],
  ): Promise<void> => {
    await assert.rejects(run, (error: unknown) => {
      assert.ok(error instanceof BuildError, `${label}: expected BuildError`);
      assert.ok(
        error.diagnostics.some((d) => expectedCodes.includes(d.code)),
        `${label}: expected one of ${expectedCodes.join(", ")}, got ${JSON.stringify(error.diagnostics)}`,
      );
      return true;
    });
  };

  // Missing cue / missing scope
  const missingSource = createMotionScopeHandle();
  const missingFollower = createMotionScopeHandle();
  const missingCue = missingSource.cue("never-mounted");
  await expectBuildError(
    "jsx-missing-cue",
    () =>
      buildJsx(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "missing" },
            createElement(
              MotionScope,
              { scope: missingFollower, startAt: after(missingCue, 20) },
              createElement(Shape, {
                id: "orphan",
                enter: fade(),
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                },
              }),
            ),
          ),
        ),
      ),
    ["broken-reference"],
  );
  await expectBuildError(
    "doc-missing-cue",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          slide({ id: "missing" }).append(
            motionScope(missingFollower, {
              startAt: after(missingCue, 20),
            }).append(
              shape({
                id: "orphan",
                enter: fade(),
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                },
              }),
            ),
          ),
        ),
      ),
    ["broken-reference"],
  );

  // Cross-step cue dependency
  const crossA = createMotionScopeHandle();
  const crossB = createMotionScopeHandle();
  const crossCue = crossA.cue("done");
  await expectBuildError(
    "cross-step-cue",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          slide({ id: "cross-step" }).append(
            step({ n: 1 }).append(
              motionScope(crossA, { publishEnd: crossCue }).append(
                shape({
                  id: "s1",
                  enter: fade(),
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 10,
                    width: 40,
                    height: 20,
                  },
                }),
              ),
            ),
            step({ n: 2 }).append(
              motionScope(crossB, { startAt: after(crossCue, 10) }).append(
                shape({
                  id: "s2",
                  enter: fade(),
                  style: {
                    position: "absolute",
                    left: 10,
                    top: 40,
                    width: 40,
                    height: 20,
                  },
                }),
              ),
            ),
          ),
        ),
      ),
    ["invalid-structure"],
  );

  // Invalid breathe target on Group, fillColor on unfilled Shape, chart granularity on Shape, and repeated word iteration
  await expectBuildError(
    "breathe-on-group",
    () =>
      buildJsx(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Slide,
            { id: "bad-breathe-group" },
            createElement(
              Group,
              {
                id: "g",
                emphasis: breathe({
                  period: 800,
                  from: "#164E3D",
                  to: "#4ADE80",
                }),
              },
              createElement(Shape, {
                id: "m",
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                  background: "#164E3D",
                },
              }),
            ),
          ),
        ),
      ),
    ["invalid-prop"],
  );
  await expectBuildError(
    "fillColor-unfilled-shape",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          slide({ id: "bad-fill-color" }).append(
            shape({
              id: "hollow",
              emphasis: fillColor({
                from: "#10B981",
                to: "#059669",
                duration: 200,
              }),
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 40,
                height: 20,
              },
            }),
          ),
        ),
      ),
    ["invalid-prop"],
  );
  await expectBuildError(
    "chart-granularity-on-shape",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          slide({ id: "bad-granularity" }).append(
            shape({
              id: "not-a-chart",
              enter: fade({ by: "series" }),
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 40,
                height: 20,
              },
            }),
          ),
        ),
      ),
    ["invalid-prop"],
  );
  await expectBuildError(
    "repeat-word-combination",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          slide({ id: "bad-repeat" }).append(
            text("Hello world", {
              id: "t",
              enter: repeat(fade({ by: "word" }) as never, 2),
              style: {
                position: "absolute",
                left: 10,
                top: 10,
                width: 120,
                fontFamily: jostFont,
              },
            }),
          ),
        ),
      ),
    ["invalid-prop"],
  );

  // Output target & vendor extension rejections (p14 sections and c15 chart labels under strict extensions: [])
  const noExtensionsTarget = {
    name: "strict-transitional",
    format: targets.transitional.format,
    extensions: [],
  } as const;
  await expectBuildError(
    "p14-section-rejected-with-motion-jsx",
    () =>
      buildJsx(
        createElement(
          Deck,
          { width: 320, height: 180 },
          createElement(
            Section,
            { name: "Part 1" },
            createElement(
              Slide,
              { id: "s1" },
              createElement(Shape, {
                id: "animated",
                enter: fade(),
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                  background: "#2563EB",
                },
              }),
            ),
          ),
        ),
        { target: noExtensionsTarget },
      ),
    ["target-violation"],
  );
  await expectBuildError(
    "p14-section-rejected-with-motion-doc",
    () =>
      buildDocument(
        deck({ width: 320, height: 180 }).append(
          section({ name: "Part 1" }).append(
            slide({ id: "s1" }).append(
              shape({
                id: "animated",
                enter: fade(),
                style: {
                  position: "absolute",
                  left: 10,
                  top: 10,
                  width: 40,
                  height: 20,
                  background: "#2563EB",
                },
              }),
            ),
          ),
        ),
        { target: noExtensionsTarget },
      ),
    ["target-violation"],
  );
  await expectBuildError(
    "c15-chart-rejected-with-motion",
    () =>
      buildDocument(
        deck({ width: 480, height: 270 }).append(
          slide({ id: "s1" }).append(
            chart({
              id: "c15-chart",
              groups: [
                {
                  kind: "doughnut",
                  data: { kind: "independent-categories" },
                  series: [
                    { name: "R1", categories: ["A", "B"], values: [10, 20] },
                    { name: "R2", categories: ["B", "A"], values: [15, 25] },
                  ],
                  dataLabels: { value: true, category: true },
                },
              ],
              enter: fade({ by: "series" }),
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 320,
                height: 200,
              },
            }),
          ),
        ),
        { target: noExtensionsTarget },
      ),
    ["target-violation"],
  );

  // Pure Motion without vendor extensions succeeds under strict extensions: [].
  const strictMotion = await buildDocument(
    deck({ width: 320, height: 180 }).append(
      slide({ id: "pure-motion" }).append(
        shape({
          id: "ok",
          enter: fade(),
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 40,
            height: 20,
            background: "#2563EB",
          },
        }),
      ),
    ),
    { target: noExtensionsTarget },
  );
  assert.deepEqual(
    (await validatePptx(strictMotion.pptx, { allowedExtensions: [] })).errors,
    [],
  );
}
