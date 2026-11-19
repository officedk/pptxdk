import assert from "node:assert/strict";

import {
  BuildError,
  CommonContainer,
  CommonNode,
  createMorphScope,
  Deck,
  Slide,
  targets,
  type BoundObjectCapabilityModule,
  type BoundSlideCapabilityModule,
  type BuildOptions,
  type CapabilityModule,
  type MorphKey,
} from "@react-pptx/core";
import { createResolver, emit } from "@react-pptx/core/build";
import { paint } from "@react-pptx/core/preview";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, expect, it, vi } from "vitest";

function docBuildWith(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode, options?: BuildOptions) =>
    emit(
      resolveWith(document, options),
      options?.signal === undefined ? {} : { signal: options.signal },
    );
}

import {
  box,
  chart,
  deck,
  documentCapabilities,
  group,
  shape,
  slide,
} from "../src/document.ts";
import { shapeCapability } from "../src/shape.ts";
import { slideBackgroundCapability } from "../src/slide-background.ts";
import { transitionCapability, transitions } from "../src/transition.ts";

describe("transitionCapability bound slide migration & generic continuity collaboration", () => {
  it("implements BoundSlideCapabilityModule, parses once per slide, and leaves static preview canvas unchanged", async () => {
    assert.equal("bindSlide" in transitionCapability, true);
    assert.equal("compileSlide" in transitionCapability, false);

    const bindSpy = vi.spyOn(transitionCapability, "bindSlide");
    try {
      const morph = createMorphScope();
      const hero = morph.handle("hero");

      const previewResolve = createResolver([
        shapeCapability,
        slideBackgroundCapability,
        transitionCapability,
      ]);

      const staticDoc = deck({ width: 320, height: 180 }).append(
        slide({ id: "first", background: "#F8FAFC" }).append(
          shape({
            id: "hero-a",
            morphKey: hero,
            x: 30,
            y: 40,
            width: 80,
            height: 60,
            fill: "#2563EB",
          }),
        ),
        slide({ id: "second", background: "#F8FAFC" }).append(
          shape({
            id: "hero-b",
            morphKey: hero,
            x: 180,
            y: 50,
            width: 100,
            height: 80,
            fill: "#1D4ED8",
          }),
        ),
      );

      const transitionDoc = deck({ width: 320, height: 180 }).append(
        slide({
          id: "first",
          background: "#F8FAFC",
          advanceAfter: 1800,
          transition: transitions.fade({ duration: 650 }),
        }).append(
          shape({
            id: "hero-a",
            morphKey: hero,
            x: 30,
            y: 40,
            width: 80,
            height: 60,
            fill: "#2563EB",
          }),
        ),
        slide({
          id: "second",
          background: "#F8FAFC",
          advanceAfter: 2500,
          transition: transitions.morph({
            duration: 1100,
            fallback: transitions.fade(),
          }),
        }).append(
          shape({
            id: "hero-b",
            morphKey: hero,
            x: 180,
            y: 50,
            width: 100,
            height: 80,
            fill: "#1D4ED8",
          }),
        ),
      );

      const staticPreview = paint(previewResolve(staticDoc));
      assert.equal(bindSpy.mock.calls.length, 0);

      const transitionResolution = previewResolve(transitionDoc);
      const transitionPreview = paint(transitionResolution);
      const transitionEmit = emit(transitionResolution);
      assert.equal(bindSpy.mock.calls.length, 2);

      // Static edit canvas pixels are identical whether transition/advanceAfter are declared or not.
      assert.deepEqual(transitionPreview.pages[0]?.diagnostics, []);
      assert.deepEqual(transitionPreview.pages[1]?.diagnostics, []);
      assert.deepEqual(
        Buffer.from(transitionPreview.pages[0].png!),
        Buffer.from(staticPreview.pages[0]!.png!),
      );
      assert.deepEqual(
        Buffer.from(transitionPreview.pages[1].png!),
        Buffer.from(staticPreview.pages[1]!.png!),
      );

      // Transition and advanceAfter observations live on transition capability observation.
      assert.deepEqual(
        transitionResolution.report.slides[0]?.capabilities?.["transition"],
        {
          kind: "fade",
          speed: "med",
          durationMs: 650,
          fallback: { kind: "fade", speed: "med" },
          advanceAfterMs: 1800,
        },
      );
      assert.deepEqual(
        transitionResolution.report.slides[1]?.capabilities?.["transition"],
        {
          kind: "morph",
          by: "object",
          speed: "slow",
          durationMs: 1100,
          fallback: { kind: "fade", speed: "slow" },
          advanceAfterMs: 2500,
        },
      );

      const file = await validatePptx(transitionEmit.pptx, {
        allowedExtensions: ["p14", "p159"],
      });
      assert.deepEqual(file.errors, []);
      assert.equal(file.slides[0]?.advanceAfterMs, 1800);
      assert.equal(file.slides[1]?.advanceAfterMs, 2500);
      assert.equal(file.slides[1]?.transition?.kind, "morph");
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("validates Slide.advanceAfter in transitionCapability and rejects when trimmed", async () => {
    const buildWithTransition = docBuildWith([
      shapeCapability,
      transitionCapability,
    ]);
    const valid = await buildWithTransition(
      deck().append(slide({ id: "timed", advanceAfter: 1500.4 })),
    );
    assert.deepEqual(valid.report.slides[0]?.capabilities?.["transition"], {
      advanceAfterMs: 1500,
    });
    const file = await validatePptx(valid.pptx);
    assert.deepEqual(file.errors, []);
    assert.equal(file.slides[0]?.advanceAfterMs, 1500);

    await assert.rejects(
      buildWithTransition(
        deck().append(slide({ id: "bad-timing", advanceAfter: -10 })),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "invalid-advance-after");
        assert.equal(error.diagnostics[0]?.path, "deck/slide[bad-timing]");
        return true;
      },
    );

    const buildWithoutTransition = docBuildWith([shapeCapability]);
    await assert.rejects(
      buildWithoutTransition(
        deck().append(slide({ id: "trimmed", advanceAfter: 1000 })),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "missing-capability");
        assert.equal(error.diagnostics[0]?.path, "deck/slide[trimmed]");
        return true;
      },
    );
  });

  it("coordinates cross-slide Morph kind matching via generic BoundSlideCapability continuity without inspecting transition props", async () => {
    const customContinuitySlideCapability: BoundSlideCapabilityModule = {
      declaration: "custom-continuity",
      slideProp: "background",
      slot: "background",
      bindSlide(input) {
        const enableCheck = input.props["background"] === "#112233";
        return {
          ...(enableCheck ? { continuity: { matchSourceKinds: true } } : {}),
          compileSlide() {
            return { xml: "", extensions: [] };
          },
        };
      },
    };

    const scope = createMorphScope();
    const shared = scope.handle("shared");

    const customGraphicModule = {
      declaration: "rp:custom-graphic",
      bind(input: { props: Record<string, unknown> }) {
        return {
          contract: {
            nativeKind: "graphicFrame" as const,
            layoutStyle: {},
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            isContainer: false,
            hasExplicitSize: true,
          },
          native: {
            morphKey: input.props["morphKey"] as ReturnType<
              typeof createMorphScope
            >["handle"] extends (n: string) => infer R
              ? R
              : never,
          },
          compile(compileContext: { id: number; morphName?: string }) {
            const name = compileContext.morphName ?? "Custom";
            return {
              xml: `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
              observation: {
                id: compileContext.id,
                kind: "custom",
                name,
                box: { x: 0, y: 0, width: 20, height: 20 },
              },
              extensions: [],
            };
          },
        };
      },
    };

    const continuityMismatchBuild = docBuildWith([
      shapeCapability,
      customGraphicModule,
      customContinuitySlideCapability,
    ]);

    const { CommonNode } = await import("@react-pptx/core");
    const mismatchDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "src" }).append(
        shape({
          id: "s1",
          morphKey: shared,
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
      ),
      slide({ id: "dst", background: "#112233" }).append(
        new CommonNode("custom", "rp:custom-graphic", {
          id: "g1",
          morphKey: shared,
        }),
      ),
    );

    await assert.rejects(
      continuityMismatchBuild(mismatchDoc),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "morph-object-kind-mismatch");
        return true;
      },
    );

    // When continuity.matchSourceKinds is not set, cross-slide kind difference is allowed.
    const noMatchDoc = deck({ width: 320, height: 180 }).append(
      slide({ id: "src" }).append(
        shape({
          id: "s1",
          morphKey: shared,
          x: 10,
          y: 10,
          width: 40,
          height: 40,
        }),
      ),
      slide({ id: "dst", background: "#445566" }).append(
        new CommonNode("custom", "rp:custom-graphic", {
          id: "g1",
          morphKey: shared,
        }),
      ),
    );
    const okResult = await continuityMismatchBuild(noMatchDoc);
    assert.equal(okResult.report.slides.length, 2);

    void Deck;
    void Slide;
    void createElement;
    void expect;
  });

  it("enforces generic BoundObjectNativeCollaboration emitsObject, grpSp membership, and explicit name conflicts", async () => {
    const scope = createMorphScope();
    const key = scope.handle("obj");

    const nonEmittingWrapper: BoundObjectCapabilityModule = {
      declaration: "rp:non-emitting",
      bind(input) {
        return {
          contract: {
            nativeKind: "sp",
            layoutStyle: {},
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            isContainer: false,
            hasExplicitSize: true,
          },
          native: {
            emitsObject: false,
            ...(input.props["morphKey"] !== undefined
              ? { morphKey: input.props["morphKey"] as MorphKey }
              : {}),
            ...(input.props["link"] !== undefined
              ? { link: input.props["link"] as { readonly id: string } }
              : {}),
          },
          compile(compileContext) {
            return {
              xml: "",
              observation: {
                id: compileContext.id,
                kind: "shape",
                name: "NonEmitting",
                box: { x: 0, y: 0, width: 10, height: 10 },
              },
              extensions: [],
            };
          },
        };
      },
    };

    const realCustomGroup: BoundObjectCapabilityModule = {
      declaration: "rp:custom-group",
      bind(input) {
        return {
          contract: {
            nativeKind: "grpSp",
            layoutStyle: {},
            padding: { top: 0, right: 0, bottom: 0, left: 0 },
            isContainer: false,
            hasExplicitSize: true,
          },
          compile(compileContext) {
            const compiled = compileContext.compileChildren(input.children);
            return {
              xml: `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${compileContext.id}" name="CustomGroup"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${compiled.xml}</p:grpSp>`,
              observation: {
                id: compileContext.id,
                kind: "group",
                name: "CustomGroup",
                box: { x: 0, y: 0, width: 100, height: 100 },
              },
              extensions: [],
            };
          },
        };
      },
    };

    const customBuild = docBuildWith([
      shapeCapability,
      nonEmittingWrapper,
      realCustomGroup,
    ]);

    await assert.rejects(
      customBuild(
        deck().append(
          slide({ id: "s1" }).append(
            new CommonNode("ne", "rp:non-emitting", { morphKey: key }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "morph-needs-object");
        return true;
      },
    );

    await assert.rejects(
      customBuild(
        deck().append(
          slide({ id: "s1" }).append(
            new CommonNode("ne", "rp:non-emitting", {
              link: { id: "s2" },
            }),
          ),
          slide({ id: "s2" }),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "link-needs-object");
        return true;
      },
    );

    await assert.rejects(
      customBuild(
        deck().append(
          slide({ id: "s1" }).append(
            new CommonContainer("cg", "rp:custom-group", {}).append(
              shape({
                id: "inner",
                morphKey: key,
                x: 0,
                y: 0,
                width: 40,
                height: 30,
              }),
            ),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "morph-inside-group");
        return true;
      },
    );
  });

  it("collaborates across Shape, Box (visible vs transparent), Group, and Chart via generic native contributions", async () => {
    const buildAll = docBuildWith(documentCapabilities);
    const scope = createMorphScope();
    const spKey = scope.handle("sp-card");
    const grpKey = scope.handle("grp");
    const chartKey = scope.handle("chart");

    // Same-kind pairings across visible Box <-> Shape, Group <-> Group, and Chart <-> Chart succeed.
    const valid = await buildAll(
      deck({ width: 960, height: 540 }).append(
        slide({ id: "s1" }).append(
          box({
            id: "box-sp",
            morphKey: spKey,
            style: {
              position: "absolute",
              left: 20,
              top: 20,
              width: 140,
              height: 80,
              backgroundColor: "#2563EB",
            },
          }),
          group({ id: "grp-1", morphKey: grpKey }).append(
            shape({ id: "g1-a", x: 200, y: 20, width: 60, height: 40 }),
          ),
          chart({
            id: "chart-1",
            morphKey: chartKey,
            x: 320,
            y: 20,
            width: 240,
            height: 160,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "S", values: [2] }],
              },
            ],
          }),
        ),
        slide({
          id: "s2",
          transition: transitions.morph({ fallback: transitions.fade() }),
        }).append(
          shape({
            id: "shape-sp",
            morphKey: spKey,
            x: 60,
            y: 220,
            width: 160,
            height: 90,
            fill: "#1D4ED8",
          }),
          group({ id: "grp-2", morphKey: grpKey }).append(
            shape({ id: "g2-a", x: 260, y: 220, width: 70, height: 50 }),
          ),
          chart({
            id: "chart-2",
            morphKey: chartKey,
            x: 400,
            y: 200,
            width: 280,
            height: 180,
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A"] },
                series: [{ name: "S", values: [5] }],
              },
            ],
          }),
        ),
      ),
    );
    const validFile = await validatePptx(valid.pptx, {
      allowedExtensions: ["p14", "p159"],
    });
    assert.deepEqual(validFile.errors, []);
    assert.equal(
      validFile.slides[0]?.objects[0]?.name,
      validFile.slides[1]?.objects[0]?.name,
    );
    assert.equal(
      validFile.slides[0]?.objects[1]?.name,
      validFile.slides[1]?.objects[1]?.name,
    );
    assert.equal(
      validFile.slides[0]?.objects[2]?.name,
      validFile.slides[1]?.objects[2]?.name,
    );

    // Transparent Box rejects morphKey with morph-needs-object.
    await assert.rejects(
      buildAll(
        deck().append(
          slide({ id: "transparent-box" }).append(
            box({ morphKey: spKey, style: { width: 100, height: 60 } }),
          ),
        ),
      ),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "morph-needs-object");
        return true;
      },
    );

    // Cross-kind mismatch across visible Box (sp) <-> Chart (graphicFrame) and Group (grpSp) <-> Shape (sp),
    // including across hyperlink-reachable non-adjacent slides!
    const mismatchPairs = [
      {
        label: "visible Box vs Chart",
        srcObj: box({
          morphKey: spKey,
          style: { width: 100, height: 60, backgroundColor: "#112233" },
        }),
        dstObj: chart({
          morphKey: spKey,
          x: 20,
          y: 20,
          width: 200,
          height: 120,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", values: [1] }],
            },
          ],
        }),
      },
      {
        label: "Group vs Shape",
        srcObj: group({ morphKey: grpKey }).append(
          shape({ x: 10, y: 10, width: 40, height: 30 }),
        ),
        dstObj: shape({
          morphKey: grpKey,
          x: 40,
          y: 40,
          width: 80,
          height: 60,
        }),
      },
      {
        label: "Group vs Chart",
        srcObj: group({ morphKey: grpKey }).append(
          shape({ x: 10, y: 10, width: 40, height: 30 }),
        ),
        dstObj: chart({
          morphKey: grpKey,
          x: 20,
          y: 20,
          width: 200,
          height: 120,
          groups: [
            {
              kind: "column",
              data: { kind: "shared-categories", categories: ["A"] },
              series: [{ name: "S", values: [1] }],
            },
          ],
        }),
      },
    ];

    for (const pair of mismatchPairs) {
      await assert.rejects(
        buildAll(
          deck().append(
            slide({ id: "src" }).append(
              pair.srcObj,
              shape({
                id: "jump",
                link: { id: "dst" },
                x: 400,
                y: 20,
                width: 40,
                height: 20,
              }),
            ),
            slide({ id: "middle" }),
            slide({
              id: "dst",
              transition: transitions.morph({ fallback: transitions.fade() }),
            }).append(pair.dstObj),
          ),
        ),
        (error: unknown) => {
          assert.ok(error instanceof BuildError, pair.label);
          assert.equal(
            error.diagnostics[0]?.code,
            "morph-object-kind-mismatch",
            pair.label,
          );
          return true;
        },
      );
    }
  });

  it("enforces vendor extensions, parameter-dependent p14/p159 requirements, and rejects forbidden mechanisms even with fallback", async () => {
    const buildAll = docBuildWith(documentCapabilities);

    // Baseline-only allows standard transitions without duration and advanceAfter alone.
    const baselineDoc = deck().append(
      slide({
        id: "base-fade",
        transition: transitions.fade(),
        advanceAfter: 1200,
      }),
    );
    const baseResult = await buildAll(baselineDoc, {
      target: targets.baselineOnly,
    });
    assert.deepEqual(baseResult.report.extensions, []);
    const baseFile = await validatePptx(baseResult.pptx, {
      allowedExtensions: [],
    });
    assert.deepEqual(baseFile.errors, []);
    assert.equal(baseFile.slides[0]?.advanceAfterMs, 1200);

    // Duration on fade requires p14 and emits mc:AlternateContent with fallback; baselineOnly must reject!
    const timedFadeDoc = deck().append(
      slide({
        id: "timed-fade",
        transition: transitions.fade({ duration: 750 }),
      }),
    );
    const timedFade = await buildAll(timedFadeDoc);
    assert.deepEqual(timedFade.report.extensions, ["p14"]);
    const timedXml = strFromU8(
      unzipSync(timedFade.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      timedXml,
      /<mc:AlternateContent[^>]*>.*<mc:Choice Requires="p14">.*p14:dur="750".*<mc:Fallback><p:transition spd="med"><p:fade\/><\/p:transition><\/mc:Fallback>/u,
    );

    await assert.rejects(
      buildAll(timedFadeDoc, { target: targets.baselineOnly }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        return true;
      },
    );

    // Morph without duration requires p159 only; Morph with duration requires BOTH p14 and p159.
    const morphWithDurationDoc = deck().append(
      slide({
        id: "morph-dur",
        transition: transitions.morph({
          duration: 950,
          by: "word",
          fallback: transitions.push({ from: "left" }),
        }),
      }),
    );
    const morphDurResult = await buildAll(morphWithDurationDoc);
    assert.deepEqual(morphDurResult.report.extensions, ["p14", "p159"]);
    const morphXml = strFromU8(
      unzipSync(morphDurResult.pptx)["ppt/slides/slide1.xml"]!,
    );
    assert.match(
      morphXml,
      /<mc:Choice Requires="p14 p159">.*p14:dur="950".*<p159:morph option="byWord"\/>.*<mc:Fallback><p:transition spd="slow"><p:push dir="l"\/><\/p:transition><\/mc:Fallback>/u,
    );

    for (const target of [
      targets.baselineOnly,
      {
        ...targets.transitional,
        name: "p14-only",
        extensions: ["p14"] as const,
      },
      {
        ...targets.transitional,
        name: "p159-only",
        extensions: ["p159"] as const,
      },
    ]) {
      await assert.rejects(
        buildAll(morphWithDurationDoc, { target }),
        (error: unknown) => {
          assert.ok(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
          return true;
        },
      );
    }

    // Flip with fallback requires p14 and rejects on p159-only even with fallback declared.
    const flipDoc = deck().append(
      slide({
        id: "flip",
        transition: transitions.flip({
          direction: "left",
          fallback: transitions.strips({ direction: "rightUp" }),
        }),
      }),
    );
    await assert.rejects(
      buildAll(flipDoc, {
        target: {
          ...targets.transitional,
          name: "p159-only",
          extensions: ["p159"],
        },
      }),
      (error: unknown) => {
        assert.ok(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "forbidden-mechanism");
        return true;
      },
    );
  });
});
