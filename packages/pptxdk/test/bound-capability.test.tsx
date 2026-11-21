import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  CommonContainer,
  CommonNode,
  createMorphScope,
  type BoundContainerCollaboration,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
  type LinkTarget,
} from "@pptxdk/core";
import {
  createResolver,
  type BuildOptions,
  type CapabilityModule,
} from "@pptxdk/core/build";
import {
  CapabilityError,
  isMorphKey,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  textStyleFrom,
  textStyleKeys,
  type EmitContext,
  type MorphKey,
  type ObjectResolveContext,
  type PaintContext,
  type ResolvedObject,
  type ResolvedSlide,
  type SlideCapabilityModule,
  type SlideEmitContext,
  type SlideResolveContext,
  type ThemeScope,
} from "@pptxdk/core/capability";
import { createDeckNode, createSlideNode } from "@pptxdk/core/document";
import { hostNodeCreator } from "@pptxdk/core/host";
import { emit, emittedBuildQuery, pxToEmu, xmlEscape } from "@pptxdk/emit";
import { paint } from "@pptxdk/preview";
import { Deck, expandDocument, Slide } from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import {
  boxCapability,
  fade,
  font,
  groupCapability,
  imageCapability,
  motionCapability,
  notesCapability,
  observeMotion,
  preset,
  shapeCapability,
  targets,
  transitionCapability,
  transitions,
} from "pptxdk";
import {
  createMorphScope as createDocMorphScope,
  font as impFont,
} from "pptxdk";
import { Box, Group, P, Shape, Span } from "pptxdk/react";
import {
  box as impBox,
  deck as impDeck,
  group as impGroup,
  image as impImage,
  paragraph as impParagraph,
  shape as impShape,
  slide as impSlide,
  span as impSpan,
} from "pptxdk/vanilla";
import { createElement, type ReactNode } from "react";
import { describe, it } from "vitest";

function resolveOptionsFrom(options?: BuildOptions) {
  return options?.signal === undefined ? {} : { signal: options.signal };
}

function runJsxModules(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (element: ReactNode, options?: BuildOptions) => {
    const resolution = resolveWith(
      await expandDocument(
        element,
        options?.signal === undefined ? {} : { signal: options.signal },
      ),
      resolveOptionsFrom(options),
    );
    const emitted = emit(resolution, options);
    return {
      ...emitted,
      query: emittedBuildQuery(emitted.report, resolution.query),
    };
  };
}

function runDocModules(modules: readonly CapabilityModule[]) {
  const resolveWith = createResolver(modules);
  return async (document: CommonNode, options?: BuildOptions) => {
    const resolution = resolveWith(document, resolveOptionsFrom(options));
    const emitted = emit(resolution, options);
    return {
      ...emitted,
      query: emittedBuildQuery(emitted.report, resolution.query),
    };
  };
}

const notoFont = font("Noto Sans CJK SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
]);
const impNotoFont = impFont("Noto Sans CJK SC", [
  {
    weight: 400,
    data: new Uint8Array(
      readFileSync(
        new URL("./assets/fonts/NotoSansSC-Regular.woff2", import.meta.url),
      ),
    ),
  },
]);

type BadgeProps = {
  readonly id?: string;
  readonly label: string;
  readonly name?: string;
  readonly morphKey?: MorphKey;
  readonly link?: LinkTarget;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly style?: {
    readonly width?: number;
    readonly height?: number;
  };
};

function createBadgeNode(props: BadgeProps): CommonNode<"badge", BadgeProps> {
  return new CommonNode("badge", "custom-badge", props);
}

function CustomBadge(props: BadgeProps) {
  return createElement("custom-badge", {
    ...props,
    [hostNodeCreator]: createBadgeNode,
  });
}

function createTrackedBadgeCapability() {
  const bindCalls: string[] = [];
  const measureCalls: string[] = [];
  const resolveCalls: string[] = [];

  const module: BoundObjectCapabilityModule = {
    declaration: "custom-badge",
    bind(
      input: CapabilityInput,
      context: CapabilityParseContext,
    ): BoundObjectCapability {
      bindCalls.push(context.path);
      const label = input.props["label"];
      if (typeof label !== "string" || label.length === 0) {
        throw new CapabilityError(
          "invalid-parameter",
          "CustomBadge.label must be a non-empty string",
          context.path,
        );
      }
      const rawName = input.props["name"];
      if (
        rawName !== undefined &&
        (typeof rawName !== "string" || rawName.trim() === "")
      ) {
        throw new CapabilityError(
          "invalid-parameter",
          "CustomBadge.name must be a non-empty string",
          context.path,
        );
      }
      const rawMorphKey = input.props["morphKey"];
      let morphKey: MorphKey | undefined;
      if (rawMorphKey !== undefined) {
        if (!isMorphKey(rawMorphKey)) {
          throw new CapabilityError(
            "invalid-morph-key",
            "morphKey must be created by createMorphScope().handle(name)",
            context.path,
          );
        }
        morphKey = rawMorphKey;
      }
      const rawLink = input.props["link"] as LinkTarget | undefined;

      const contract = parseCapabilityLayoutContract({
        nativeKind: "sp",
        props: input.props,
        allowedStyleKeys: new Set<string>(layoutStyleKeys),
        path: context.path,
      });

      const declaredX =
        typeof input.props["x"] === "number" ? input.props["x"] : 0;
      const declaredY =
        typeof input.props["y"] === "number" ? input.props["y"] : 0;
      const declaredWidth =
        typeof contract.width === "number" ? contract.width : label.length * 10;
      const declaredHeight =
        typeof contract.height === "number" ? contract.height : 36;

      // Encapsulated private state: never returned on contract or exposed to core.
      const privateLabel = label;
      const privateExplicitName = rawName;

      return {
        contract,
        native: {
          ...(morphKey === undefined ? {} : { morphKey }),
          ...(privateExplicitName === undefined
            ? {}
            : { hasExplicitName: true }),
          ...(rawLink === undefined ? {} : { link: rawLink }),
        },
        measureLayout(request) {
          measureCalls.push(context.path);
          const width =
            request.availableWidth !== undefined
              ? Math.min(request.availableWidth, declaredWidth)
              : declaredWidth;
          return {
            width,
            height: declaredHeight,
            minContentWidth: Math.min(declaredWidth, 40),
            minContentHeight: declaredHeight,
          };
        },
        resolve(resolveContext) {
          resolveCalls.push(resolveContext.path);
          const box = resolveContext.layoutBox ?? {
            x: declaredX,
            y: declaredY,
            width: declaredWidth,
            height: declaredHeight,
          };
          const name =
            resolveContext.morphName ??
            privateExplicitName ??
            `Badge ${resolveContext.id - 1}`;
          return {
            observation: {
              id: resolveContext.id,
              kind: "badge",
              name,
              box,
              text: privateLabel,
            },
            extensions: [],
            emit(emitContext) {
              const link = emitContext.nativeLink;
              const linkAttrs = [
                ...(link?.relationshipId !== undefined &&
                link.relationshipId !== ""
                  ? [`r:id="${xmlEscape(link.relationshipId)}"`]
                  : []),
                ...(link?.action !== undefined
                  ? [`action="${xmlEscape(link.action)}"`]
                  : []),
              ].join(" ");
              const linkXml =
                linkAttrs.length > 0 ? `<a:hlinkClick ${linkAttrs}/>` : "";
              const xml =
                `<p:sp><p:nvSpPr>` +
                `<p:cNvPr id="${emitContext.id}" name="${xmlEscape(name)}">${linkXml}</p:cNvPr>` +
                `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
                `<p:spPr>` +
                `<a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm>` +
                `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                `<a:solidFill><a:srgbClr val="2B579A"/></a:solidFill>` +
                `</p:spPr>` +
                `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(privateLabel)}</a:t></a:r></a:p></p:txBody>` +
                `</p:sp>`;
              return { xml };
            },
          };
        },
      };
    },
  };

  return { module, bindCalls, measureCalls, resolveCalls };
}

describe("bound object capability contract", () => {
  it("supports custom bound object capability via both JSX and imperative builds without exposing parsed state", async () => {
    const tracker = createTrackedBadgeCapability();
    const morph = createMorphScope();
    const badgeMorph = morph.handle("badge-morph");

    // 1. Imperative build via createResolver + emit
    const deckNode = createDeckNode({ width: 960, height: 540 });
    const slide1 = createSlideNode({ id: "first" });
    const badge1 = createBadgeNode({
      id: "imp-badge",
      label: "Imperative",
      morphKey: badgeMorph,
      link: { id: "second" },
      x: 40,
      y: 50,
      width: 160,
      height: 44,
    });
    const slide2 = createSlideNode({ id: "second" });
    slide1.append(badge1);
    deckNode.append(slide1, slide2);

    const imperativeBuild = runDocModules([tracker.module]);
    const imperativeResult = await imperativeBuild(deckNode);
    const impValidation = await validatePptx(imperativeResult.pptx);
    assert.deepEqual(impValidation.errors, []);

    assert.equal(tracker.bindCalls.length, 1);
    assert.equal(tracker.resolveCalls.length, 1);
    assert.equal(
      imperativeResult.report.slides[0]?.objects[0]?.text,
      "Imperative",
    );
    assert.equal(
      imperativeResult.report.slides[0]?.objects[0]?.link,
      "slide:second",
    );
    assert.match(
      imperativeResult.report.slides[0]?.objects[0]?.name ?? "",
      /^!!morph-\d+$/u,
    );

    // 2. JSX build inside Box Flex container (exercises measureLayout + container collaboration)
    tracker.bindCalls.length = 0;
    tracker.measureCalls.length = 0;
    tracker.resolveCalls.length = 0;

    const jsxBuild = runJsxModules([boxCapability, tracker.module]);
    const jsxResult = await jsxBuild(
      <Deck width={960} height={540}>
        <Slide id="jsx-first">
          <Box
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "flex-start",
              width: 500,
              height: 120,
              padding: 10,
            }}
          >
            <CustomBadge
              id="jsx-badge"
              label="JSX-Flex"
              link={{ url: "https://example.com/badge" }}
            />
          </Box>
        </Slide>
      </Deck>,
    );
    const jsxValidation = await validatePptx(jsxResult.pptx);
    assert.deepEqual(jsxValidation.errors, []);
    assert.equal(tracker.bindCalls.length, 1);
    assert.ok(tracker.measureCalls.length >= 1);
    assert.equal(tracker.resolveCalls.length, 1);

    const observed = jsxResult.query
      .getSlide("jsx-first")
      .get("jsx-badge")
      .observe();
    assert.equal(observed.authorId, "jsx-badge");
    assert.equal(observed.text, "JSX-Flex");
    assert.deepEqual(observed.box, {
      x: 10,
      y: 10,
      width: 80,
      height: 36,
    });
    assert.equal(observed.objects[0]?.link, "url:https://example.com/badge");
  });

  it("enforces morph-object-name-conflict via native.hasExplicitName and missing-capability when trimmed", async () => {
    const tracker = createTrackedBadgeCapability();
    const morph = createMorphScope();
    const conflictKey = morph.handle("conflict");

    const buildWithBadge = runJsxModules([tracker.module]);
    await assert.rejects(
      buildWithBadge(
        <Deck width={960} height={540}>
          <Slide id="conflict-slide">
            <CustomBadge
              id="conflicting-badge"
              label="Conflict"
              name="ExplicitName"
              morphKey={conflictKey}
              x={10}
              y={10}
              width={100}
              height={40}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some(
            (d) => d.code === "morph-object-name-conflict",
          ),
        );
        return true;
      },
    );

    const trimmedBuild = runJsxModules([]);
    await assert.rejects(
      trimmedBuild(
        <Deck width={960} height={540}>
          <Slide id="trimmed-slide">
            <CustomBadge
              label="Missing"
              x={10}
              y={10}
              width={100}
              height={40}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "missing-capability"),
        );
        return true;
      },
    );
  });

  it("migrates shapeCapability to BoundObjectCapabilityModule with single bind per node per build", async () => {
    const trimmedShapeBuild = runJsxModules([shapeCapability]);
    const result = await trimmedShapeBuild(
      <Deck width={960} height={540}>
        <Slide id="standalone-shape">
          <Shape
            id="hero-shape"
            geometry={preset("roundRect")}
            x={80}
            y={90}
            width={260}
            height={100}
            fill="#2B579A"
            style={{ fontFamily: notoFont, fontSize: 18, color: "#FFFFFF" }}
          >
            <P>
              <Span>Bound Shape</Span>
            </P>
          </Shape>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.diagnostics, []);
    const validation = await validatePptx(result.pptx);
    assert.deepEqual(validation.errors, []);
    assert.equal(validation.slides[0]?.objects[0]?.text?.text, "Bound Shape");
  });

  it("produces consistent native PPTX and observations for Shape across JSX and imperative entries inside unmigrated containers", async () => {
    const jsxMorph = createMorphScope().handle("shape-card");
    const impMorph = createDocMorphScope().handle("shape-card");

    const jsxBuild = runJsxModules([
      shapeCapability,
      boxCapability,
      groupCapability,
      transitionCapability,
      motionCapability,
    ]);
    const docBuild = runDocModules([
      shapeCapability,
      boxCapability,
      groupCapability,
      transitionCapability,
      motionCapability,
    ]);

    const jsxResult = await jsxBuild(
      <Deck width={960} height={540}>
        <Slide id="main">
          <Shape
            id="standalone"
            geometry="hexagon"
            morphKey={jsxMorph}
            link={{ id: "next" }}
            enter={fade()}
            x={40}
            y={40}
            width={180}
            height={70}
            fill="#1D4ED8"
            style={{ fontFamily: notoFont, fontSize: 16, color: "#FFFFFF" }}
            text="Standalone"
          />
          <Box
            id="flex-row"
            style={{
              position: "absolute",
              left: 40,
              top: 140,
              width: 500,
              height: 120,
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              gap: 16,
            }}
          >
            <Shape
              id="flex-shape-a"
              width={180}
              fill="#0F766E"
              style={{
                fontFamily: notoFont,
                fontSize: 24,
                color: "#FFFFFF",
                verticalAlign: "top",
                padding: 8,
              }}
            >
              <P>
                <Span>Baseline A</Span>
              </P>
            </Shape>
            <Shape
              id="flex-shape-b"
              width={160}
              fill="#B45309"
              style={{
                fontFamily: notoFont,
                fontSize: 14,
                color: "#FFFFFF",
                verticalAlign: "top",
                padding: 8,
              }}
            >
              <P>
                <Span>Baseline B</Span>
              </P>
            </Shape>
          </Box>
          <Box
            id="grid-box"
            style={{
              position: "absolute",
              left: 40,
              top: 280,
              width: 320,
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: 12,
            }}
          >
            <Shape
              id="grid-shape-1"
              fill="#334155"
              style={{
                fontFamily: notoFont,
                fontSize: 14,
                color: "#FFFFFF",
                padding: 8,
              }}
              text="Grid Cell 1"
            />
            <Shape
              id="grid-shape-2"
              fill="#475569"
              style={{
                fontFamily: notoFont,
                fontSize: 14,
                color: "#FFFFFF",
                padding: 8,
              }}
              text="Grid Cell 2"
            />
          </Box>
          <Group id="grouped-shapes">
            <Shape
              id="grouped-rect"
              x={600}
              y={140}
              width={140}
              height={80}
              fill="#7C3AED"
              style={{ fontFamily: notoFont }}
              text="In Group"
            />
          </Group>
        </Slide>
        <Slide
          id="next"
          transition={transitions.morph({ fallback: transitions.fade() })}
        >
          <Shape
            id="standalone-next"
            geometry="hexagon"
            morphKey={jsxMorph}
            x={300}
            y={80}
            width={220}
            height={90}
            fill="#1D4ED8"
            style={{ fontFamily: notoFont }}
            text="Standalone"
          />
        </Slide>
      </Deck>,
    );

    const impDeckNode = impDeck({ width: 960, height: 540 });
    const impSlideMain = impSlide({ id: "main" });
    const impStandalone = impShape({
      id: "standalone",
      geometry: "hexagon",
      morphKey: impMorph,
      link: { id: "next" },
      enter: fade(),
      x: 40,
      y: 40,
      width: 180,
      height: 70,
      fill: "#1D4ED8",
      style: { fontFamily: impNotoFont, fontSize: 16, color: "#FFFFFF" },
      text: "Standalone",
    });
    const impFlexRow = impBox({
      id: "flex-row",
      style: {
        position: "absolute",
        left: 40,
        top: 140,
        width: 500,
        height: 120,
        display: "flex",
        flexDirection: "row",
        alignItems: "baseline",
        gap: 16,
      },
    }).append(
      impShape({
        id: "flex-shape-a",
        width: 180,
        fill: "#0F766E",
        style: {
          fontFamily: impNotoFont,
          fontSize: 24,
          color: "#FFFFFF",
          verticalAlign: "top",
          padding: 8,
        },
      }).append(impParagraph().append(impSpan("Baseline A"))),
      impShape({
        id: "flex-shape-b",
        width: 160,
        fill: "#B45309",
        style: {
          fontFamily: impNotoFont,
          fontSize: 14,
          color: "#FFFFFF",
          verticalAlign: "top",
          padding: 8,
        },
      }).append(impParagraph().append(impSpan("Baseline B"))),
    );
    const impGridBox = impBox({
      id: "grid-box",
      style: {
        position: "absolute",
        left: 40,
        top: 280,
        width: 320,
        display: "grid",
        gridTemplateColumns: "1fr 1fr",
        gap: 12,
      },
    }).append(
      impShape({
        id: "grid-shape-1",
        fill: "#334155",
        style: {
          fontFamily: impNotoFont,
          fontSize: 14,
          color: "#FFFFFF",
          padding: 8,
        },
        text: "Grid Cell 1",
      }),
      impShape({
        id: "grid-shape-2",
        fill: "#475569",
        style: {
          fontFamily: impNotoFont,
          fontSize: 14,
          color: "#FFFFFF",
          padding: 8,
        },
        text: "Grid Cell 2",
      }),
    );
    const impGrouped = impGroup({ id: "grouped-shapes" }).append(
      impShape({
        id: "grouped-rect",
        x: 600,
        y: 140,
        width: 140,
        height: 80,
        fill: "#7C3AED",
        style: { fontFamily: impNotoFont },
        text: "In Group",
      }),
    );
    impSlideMain.append(impStandalone, impFlexRow, impGridBox, impGrouped);
    const impSlideNext = impSlide({
      id: "next",
      transition: transitions.morph({ fallback: transitions.fade() }),
    }).append(
      impShape({
        id: "standalone-next",
        geometry: "hexagon",
        morphKey: impMorph,
        x: 300,
        y: 80,
        width: 220,
        height: 90,
        fill: "#1D4ED8",
        style: { fontFamily: impNotoFont },
        text: "Standalone",
      }),
    );
    impDeckNode.append(impSlideMain, impSlideNext);
    const impResult = await docBuild(impDeckNode);

    for (const result of [jsxResult, impResult]) {
      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
      assert.equal(validation.slides[0]?.objects.length, 6);
      assert.equal(
        result.query.getSlide("main").get("standalone").observe().objects[0]
          ?.link,
        "slide:next",
      );
      assert.equal(
        result.query
          .getSlide("main")
          .get("grouped-shapes")
          .get("grouped-rect")
          .observe().text,
        "In Group",
      );
      const slide1Xml = strFromU8(
        unzipSync(result.pptx)["ppt/slides/slide1.xml"]!,
      );
      assert.match(slide1Xml, /<p:timing>/u);
      assert.match(slide1Xml, /<a:prstGeom prst="hexagon">/u);
    }

    const jsxFlexA = jsxResult.query
      .getSlide("main")
      .get("flex-row")
      .get("flex-shape-a")
      .observe().box!;
    const jsxFlexB = jsxResult.query
      .getSlide("main")
      .get("flex-row")
      .get("flex-shape-b")
      .observe().box!;
    const impFlexA = impResult.query
      .getSlide("main")
      .get("flex-row")
      .get("flex-shape-a")
      .observe().box!;
    const impFlexB = impResult.query
      .getSlide("main")
      .get("flex-row")
      .get("flex-shape-b")
      .observe().box!;
    assert.deepEqual(jsxFlexA, impFlexA);
    assert.deepEqual(jsxFlexB, impFlexB);
    assert.ok(impFlexB.y > impFlexA.y);
  });

  it("isolates per-build bound state across document edits and rebuilds", async () => {
    const node = impShape({
      id: "mutable-shape",
      geometry: "rect",
      x: 20,
      y: 30,
      width: 140,
      height: 60,
      fill: "#112233",
      text: "Version 1",
    });
    const deckNode = impDeck({ width: 960, height: 540 }).append(
      impSlide({ id: "s1" }).append(node),
    );
    const docBuild = runDocModules([shapeCapability]);

    const first = await docBuild(deckNode);

    node.set({
      geometry: "ellipse",
      width: 220,
      fill: "#445566",
      text: "Version 2",
    });

    const second = await docBuild(deckNode);

    // First build outputs remain bound to Version 1
    const firstFile = await validatePptx(first.pptx);
    assert.deepEqual(firstFile.errors, []);
    assert.equal(firstFile.slides[0]?.objects[0]?.text?.text, "Version 1");
    assert.equal(
      first.query.getSlide("s1").get("mutable-shape").observe().box?.width,
      140,
    );

    // Second build outputs reflect Version 2
    const secondFile = await validatePptx(second.pptx);
    assert.deepEqual(secondFile.errors, []);
    assert.equal(secondFile.slides[0]?.objects[0]?.text?.text, "Version 2");
    assert.equal(
      second.query.getSlide("s1").get("mutable-shape").observe().box?.width,
      220,
    );
  });

  it("validates Shape link declarations in bind and resolves valid links in core", async () => {
    const shapeBuild = runJsxModules([shapeCapability]);
    await assert.rejects(
      shapeBuild(
        <Deck width={960} height={540}>
          <Slide id="bad-link-shape">
            <Shape
              id="malformed-link"
              x={10}
              y={10}
              width={100}
              height={40}
              link={"not-an-object" as unknown as LinkTarget}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(error.diagnostics.some((d) => d.code === "invalid-link"));
        return true;
      },
    );

    await assert.rejects(
      shapeBuild(
        <Deck width={960} height={540}>
          <Slide id="bad-url-shape">
            <Shape
              id="malformed-url"
              x={10}
              y={10}
              width={100}
              height={40}
              link={{ url: "not a valid url" }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(error.diagnostics.some((d) => d.code === "invalid-link"));
        return true;
      },
    );
  });

  it("migrates tableCapability to BoundObjectCapabilityModule and resolves generic hierarchical sub-part locators", async () => {
    const compositeWithParts: BoundObjectCapabilityModule = {
      declaration: "composite-grid",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "graphicFrame",
          props: input.props,
          allowedStyleKeys: new Set(layoutStyleKeys),
          path: context.path,
        });
        return {
          contract,
          resolve(resolveContext) {
            const box = resolveContext.layoutBox ?? {
              x: 10,
              y: 20,
              width: 120,
              height: 40,
            };
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: "Composite",
                geometry: "rect",
                box,
              },
              parts: [
                {
                  authorId: "band-1",
                  box: { x: 10, y: 20, width: 120, height: 20 },
                  text: "Alpha\tBeta",
                  parts: [
                    {
                      authorId: "slot-a",
                      box: { x: 10, y: 20, width: 60, height: 20 },
                      text: "Alpha",
                    },
                    {
                      authorId: "slot-b",
                      box: { x: 70, y: 20, width: 60, height: 20 },
                      text: "Beta",
                    },
                  ],
                },
              ],
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="Composite"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const customBuild = runDocModules([compositeWithParts]);
    const doc = impDeck({ width: 320, height: 180 }).append(
      impSlide({ id: "generic-parts" }).append(
        new CommonNode("composite", "composite-grid", { id: "grid" }),
      ),
    );
    const result = await customBuild(doc);
    const grid = result.query.getSlide("generic-parts").get("grid");
    assert.equal(grid.observe().text, "Alpha\tBeta");
    const band = grid.get("band-1");
    assert.equal(band.observe().text, "Alpha\tBeta");
    assert.deepEqual(band.observe().objects, []);
    assert.equal(band.children().length, 2);
    assert.deepEqual(band.get("slot-a").observe().box, {
      x: 10,
      y: 20,
      width: 60,
      height: 20,
    });
    assert.equal(band.get("slot-b").observe().text, "Beta");
  });

  it("migrates boxCapability to BoundObjectCapabilityModule with container role", async () => {
    const trimmedBuild = runJsxModules([boxCapability, shapeCapability]);
    const result = await trimmedBuild(
      <Deck width={960} height={540}>
        <Slide id="box-container-slide">
          <Box
            id="transparent-card"
            style={{ display: "flex", width: 320, padding: 12 }}
          >
            <Shape id="child-shape" width={120} height={60} fill="#1D4ED8" />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.diagnostics, []);
    assert.equal(result.report.slides[0]?.objects.length, 1);
    assert.equal(result.report.slides[0]?.layouts?.length, 1);
    const observed = result.query
      .getSlide("box-container-slide")
      .get("transparent-card")
      .observe();
    assert.deepEqual(observed.box, { x: 0, y: 0, width: 320, height: 84 });
    assert.equal(observed.objects.length, 1);
    assert.equal(observed.objects[0]?.authorId, "child-shape");

    const withoutBoxBuild = runJsxModules([shapeCapability]);
    await assert.rejects(
      withoutBoxBuild(
        <Deck width={960} height={540}>
          <Slide id="missing-box">
            <Box id="cropped-box" style={{ width: 100, height: 40 }}>
              <Shape width={40} height={20} />
            </Box>
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "missing-capability"),
        );
        return true;
      },
    );
  });

  it("coordinates generic non-Box BoundContainerCollaboration modules for layout, scoped theme, baseline, overflow, and native objects", async () => {
    type PanelProps = {
      readonly id?: string;
      readonly filled?: boolean;
      readonly link?: LinkTarget;
      readonly morphKey?: MorphKey;
      readonly theme?: ThemeScope;
      readonly style?: Record<string, unknown>;
    };
    const createPanelNode = (props: PanelProps) =>
      new CommonContainer("panel", "custom-panel", props);
    const CustomPanel = (
      props: PanelProps & { readonly children?: React.ReactNode },
    ) =>
      createElement(
        "custom-panel",
        { ...props, [hostNodeCreator]: createPanelNode },
        props.children,
      );

    const panelAllowedStyle = new Set([
      ...layoutStyleKeys,
      ...textStyleKeys,
      "padding",
    ]);
    const customPanelModule: BoundObjectCapabilityModule = {
      declaration: "custom-panel",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: panelAllowedStyle,
          allowPadding: true,
          isContainer: true,
          path: context.path,
        });
        const emitsObject = input.props["filled"] === true;
        const derivedTheme = context.theme.with(
          input.props["theme"] as ThemeScope | undefined,
          textStyleFrom(input.props["style"]),
        );
        const container: BoundContainerCollaboration = {
          children: input.children,
          theme: derivedTheme,
          emitsObject,
          propagatesBaseline: true,
          checksCanvasOverflow: true,
        };
        const morphKey = input.props["morphKey"] as MorphKey | undefined;
        const link = input.props["link"] as LinkTarget | undefined;
        return {
          contract,
          container,
          ...(morphKey !== undefined || link !== undefined
            ? {
                native: {
                  ...(morphKey === undefined ? {} : { morphKey }),
                  ...(link === undefined ? {} : { link }),
                },
              }
            : {}),
          resolve(resolveContext) {
            const box = resolveContext.layoutBox!;
            const name =
              resolveContext.morphName ?? `Panel ${resolveContext.id}`;
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name,
                geometry: "rect",
                box,
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: emitsObject
                    ? `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${xmlEscape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></p:spPr></p:sp>`
                    : "",
                };
              },
            };
          },
        };
      },
    };

    const panelMorph = createMorphScope().handle("panel-morph");
    const panelBuild = runJsxModules([customPanelModule, shapeCapability]);
    const valid = await panelBuild(
      <Deck
        width={400}
        height={240}
        theme={{ colors: { panelAccent: "#111111" } }}
      >
        <Slide id="panel-slide">
          <CustomPanel
            id="outer-panel"
            filled
            morphKey={panelMorph}
            theme={{ colors: { panelAccent: "#2563EB" } }}
            style={{
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              width: 300,
              padding: 10,
              gap: 12,
              fontFamily: notoFont,
              fontSize: 18,
            }}
          >
            <CustomPanel
              id="inner-transparent-panel"
              style={{
                display: "flex",
                flexDirection: "row",
                alignItems: "flex-start",
              }}
            >
              <Shape
                id="panel-shape-a"
                width={100}
                fill="panelAccent"
                style={{ verticalAlign: "top", padding: 4 }}
              >
                <P>
                  <Span>Row A</Span>
                </P>
              </Shape>
            </CustomPanel>
            <Shape
              id="panel-shape-b"
              width={100}
              fill="#0F766E"
              style={{ fontSize: 12, verticalAlign: "top", padding: 4 }}
            >
              <P>
                <Span>Row B</Span>
              </P>
            </Shape>
          </CustomPanel>
          <CustomPanel
            id="overflow-panel"
            style={{
              position: "absolute",
              left: 350,
              top: 10,
              width: 100,
              height: 40,
            }}
          />
        </Slide>
      </Deck>,
    );

    assert.equal(
      valid.report.diagnostics.filter((d) => d.severity === "error").length,
      0,
    );
    assert.ok(
      valid.report.diagnostics.some(
        (w) =>
          w.severity === "warning" &&
          w.code === "overflow" &&
          w.path.includes("custom-panel[overflow-panel]"),
      ),
    );
    const outerObserved = valid.query
      .getSlide("panel-slide")
      .get("outer-panel")
      .observe();
    assert.equal(outerObserved.objects[0]?.kind, "shape");
    const innerObserved = valid.query
      .getSlide("panel-slide")
      .get("outer-panel")
      .get("inner-transparent-panel")
      .observe();
    // Transparent container has layout observation box, and its native objects are its descendants.
    assert.ok(innerObserved.box !== undefined);
    assert.equal(innerObserved.objects.length, 1);
    assert.equal(innerObserved.objects[0]?.authorId, "panel-shape-a");
    // Baseline shifted smaller font shape downward to align with larger font shape.
    assert.ok(
      valid.query
        .getSlide("panel-slide")
        .get("outer-panel")
        .get("panel-shape-b")
        .observe().box!.y > innerObserved.objects[0].box!.y,
    );
    const unzipped = unzipSync(valid.pptx);
    assert.match(strFromU8(unzipped["ppt/slides/slide1.xml"]!), /2563EB/u);

    // Transparent custom container rejects link and morphKey.
    await assert.rejects(
      panelBuild(
        <Deck width={400} height={240}>
          <Slide id="bad-transparent-link">
            <CustomPanel
              link={{ action: "home" }}
              style={{ width: 100, height: 40 }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "link-needs-object"),
        );
        return true;
      },
    );
    await assert.rejects(
      panelBuild(
        <Deck width={400} height={240}>
          <Slide id="bad-transparent-morph">
            <CustomPanel
              morphKey={panelMorph}
              style={{ width: 100, height: 40 }}
            />
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "morph-needs-object"),
        );
        return true;
      },
    );
  });

  it("collects annotations from transparent bound containers and binds motionCapability onto BoundSlideCapabilityModule", async () => {
    const observedMarkers: unknown[] = [];
    const customTimingSlide: import("@pptxdk/core").BoundSlideCapabilityModule =
      {
        declaration: "custom-timing",
        slideProp: "customTiming",
        slot: "timing",
        always: true,
        bindSlide() {
          return {
            resolveSlide(resolveContext) {
              for (const contribution of resolveContext.contributions) {
                if (contribution.capability === "custom-timing") {
                  observedMarkers.push(contribution.value);
                }
              }
              return {
                extensions: [],
                observation: { count: observedMarkers.length },
                emit() {
                  return { xml: "" };
                },
              };
            },
          };
        },
      };

    const transparentAnnotator: BoundObjectCapabilityModule = {
      declaration: "transparent-annotator",
      bind(input, parseContext) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(layoutStyleKeys),
          isContainer: true,
          path: parseContext.path,
        });
        return {
          contract,
          container: {
            children: input.children,
            emitsObject: false,
          },
          resolve() {
            return {
              observation: {
                id: 0,
                kind: "shape",
                name: "transparent",
                box: { x: 0, y: 0, width: 100, height: 40 },
              },
              extensions: [],
              annotations: [
                {
                  capability: "custom-timing",
                  value: { marker: "from-transparent-container" },
                },
              ],
              emit() {
                return { xml: "" };
              },
            };
          },
        };
      },
    };

    const customBuild = runDocModules([
      transparentAnnotator,
      shapeCapability,
      customTimingSlide,
    ]);
    const doc = impDeck({ width: 320, height: 180 }).append(
      impSlide({ id: "annotator-slide" }).append(
        new CommonContainer("annotator", "transparent-annotator", {
          style: { width: 100, height: 40 },
        }).append(impShape({ id: "child", x: 0, y: 0, width: 20, height: 20 })),
      ),
    );
    const result = await customBuild(doc);
    assert.deepEqual(observedMarkers, [
      { marker: "from-transparent-container" },
    ]);
    assert.deepEqual(result.report.slides[0]?.capabilities?.["custom-timing"], {
      count: 1,
    });
  });

  it("solves Block/Flex automatic minimums, baseline alignment, and absolute coordinates strictly from bound contracts without reading raw leaf props", async () => {
    type MetricTileProps = {
      readonly id?: string;
      readonly spanW?: number;
      readonly spanH?: number;
      readonly posX?: number;
      readonly posY?: number;
      readonly baselineAt?: number;
      readonly minW?: number;
      readonly style?: Record<string, unknown>;
    };
    const createTileNode = (props: MetricTileProps) =>
      new CommonNode("tile", "metric-tile", props);
    const MetricTile = (props: MetricTileProps) =>
      createElement("metric-tile", {
        ...props,
        [hostNodeCreator]: createTileNode,
      });

    const tileModule: BoundObjectCapabilityModule = {
      declaration: "metric-tile",
      bind(input, context) {
        // Deliberately omit props.width/height/x/y; supply them strictly on contract.
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: {
            ...(input.props["posX"] === undefined
              ? {}
              : { x: input.props["posX"] }),
            ...(input.props["posY"] === undefined
              ? {}
              : { y: input.props["posY"] }),
            ...(input.props["spanW"] === undefined
              ? {}
              : { width: input.props["spanW"] }),
            ...(input.props["spanH"] === undefined
              ? {}
              : { height: input.props["spanH"] }),
            ...(input.props["style"] === undefined
              ? {}
              : { style: input.props["style"] }),
          },
          allowedStyleKeys: new Set<string>(layoutStyleKeys),
          path: context.path,
        });
        const baselineAt =
          typeof input.props["baselineAt"] === "number"
            ? input.props["baselineAt"]
            : undefined;
        const minW =
          typeof input.props["minW"] === "number" ? input.props["minW"] : 40;
        const naturalW =
          typeof contract.width === "number" ? contract.width : 80;
        const naturalH =
          typeof contract.height === "number" ? contract.height : 30;
        return {
          contract,
          measureLayout() {
            return {
              width: naturalW,
              height: naturalH,
              minContentWidth: minW,
              minContentHeight: naturalH,
              ...(baselineAt === undefined ? {} : { baseline: baselineAt }),
            };
          },
          resolve(resolveContext) {
            const box = resolveContext.layoutBox!;
            const name = `Tile ${resolveContext.id}`;
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name,
                geometry: "rect",
                box,
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const customBuild = runJsxModules([tileModule, boxCapability]);
    const result = await customBuild(
      <Deck width={400} height={240}>
        <Slide id="contract-only-layout">
          <Box
            id="narrow-row"
            style={{
              position: "absolute",
              left: 20,
              top: 20,
              width: 140,
              display: "flex",
              flexDirection: "row",
              alignItems: "baseline",
              gap: 10,
            }}
          >
            {/* spanW=60 clamps minContentWidth=100 down to 60 via contract.width */}
            <MetricTile
              id="clamped-tile"
              spanW={60}
              spanH={40}
              minW={100}
              baselineAt={28}
            />
            <MetricTile
              id="shifted-tile"
              spanW={120}
              spanH={24}
              minW={80}
              baselineAt={14}
            />
          </Box>
          <MetricTile
            id="abs-tile"
            posX={220}
            posY={40}
            spanW={90}
            spanH={50}
            style={{ position: "absolute" }}
          />
        </Slide>
      </Deck>,
    );

    const clamped = result.query
      .getSlide("contract-only-layout")
      .get("narrow-row")
      .get("clamped-tile")
      .observe().box!;
    const shifted = result.query
      .getSlide("contract-only-layout")
      .get("narrow-row")
      .get("shifted-tile")
      .observe().box!;
    const abs = result.query
      .getSlide("contract-only-layout")
      .get("abs-tile")
      .observe().box!;

    // Total width 140, gap 10 -> 130 available; clamped-tile min is min(100, 60)=60, shifted-tile min is 80 -> 60 + 80 > 130 so flexShrink keeps at least 60/80 or proportional shrink above contract clamp.
    assert.ok(clamped.width >= 60);
    assert.ok(shifted.width >= 70);
    // Baseline 28 vs 14 shifts shifted-tile down by 14px.
    assert.equal(clamped.y, 20);
    assert.equal(shifted.y, 34);
    // Absolute position comes from contract.x / contract.y (posX / posY).
    assert.deepEqual(abs, { x: 220, y: 40, width: 90, height: 50 });
  });

  it("migrates groupCapability to BoundObjectCapabilityModule", async () => {
    const trimmedBuild = runJsxModules([groupCapability, shapeCapability]);
    const result = await trimmedBuild(
      <Deck width={960} height={540} theme={{ colors: { ink: "#111111" } }}>
        <Slide id="group-bound-slide">
          <Group id="loose-group" theme={{ colors: { ink: "#2563EB" } }}>
            <Shape id="m1" x={10} y={20} width={50} height={30} fill="ink" />
            <Shape
              id="m2"
              x={70}
              y={30}
              width={40}
              height={40}
              fill="#0F766E"
            />
          </Group>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(result.report.diagnostics, []);
    assert.equal(result.report.slides[0]?.layouts, undefined);
    const observed = result.query
      .getSlide("group-bound-slide")
      .get("loose-group")
      .observe();
    assert.deepEqual(observed.box, { x: 10, y: 20, width: 100, height: 50 });
    assert.deepEqual(
      observed.objects.map((object) => object.authorId),
      ["loose-group", "m1", "m2"],
    );

    const withoutGroupBuild = runJsxModules([shapeCapability]);
    await assert.rejects(
      withoutGroupBuild(
        <Deck width={960} height={540}>
          <Slide id="missing-group">
            <Group id="cropped-group">
              <Shape x={0} y={0} width={40} height={20} />
            </Group>
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "missing-capability"),
        );
        return true;
      },
    );
  });

  it("coordinates generic non-Group BoundContainerCollaboration modules with when-styled layout, member composition, and continuity restrictions", async () => {
    type ClusterProps = {
      readonly id?: string;
      readonly name?: string;
      readonly morphKey?: MorphKey;
      readonly link?: LinkTarget;
      readonly theme?: ThemeScope;
      readonly style?: Record<string, unknown>;
    };
    const createClusterNode = (props: ClusterProps) =>
      new CommonContainer("cluster", "custom-cluster", props);
    const CustomCluster = (
      props: ClusterProps & { readonly children?: React.ReactNode },
    ) =>
      createElement(
        "custom-cluster",
        { ...props, [hostNodeCreator]: createClusterNode },
        props.children,
      );

    const clusterAllowedStyle = new Set([
      ...layoutStyleKeys,
      ...textStyleKeys,
      "padding",
    ]);
    const customClusterModule: BoundObjectCapabilityModule = {
      declaration: "custom-cluster",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "grpSp",
          props: input.props,
          allowedStyleKeys: clusterAllowedStyle,
          allowPadding: true,
          isContainer: true,
          path: context.path,
        });
        const derivedTheme = context.theme.with(
          input.props["theme"] as ThemeScope | undefined,
          textStyleFrom(input.props["style"]),
        );
        const morphKey = input.props["morphKey"] as MorphKey | undefined;
        const link = input.props["link"] as LinkTarget | undefined;
        const requestedName =
          typeof input.props["name"] === "string"
            ? input.props["name"]
            : undefined;
        const container: BoundContainerCollaboration = {
          children: input.children,
          theme: derivedTheme,
          emitsObject: true,
          layoutMode: "when-styled",
          childComposition: "members",
          restrictsMemberContinuity: true,
          observesLayout: false,
          propagatesBaseline: false,
          checksCanvasOverflow: false,
        };
        return {
          contract,
          container,
          ...(morphKey !== undefined ||
          requestedName !== undefined ||
          link !== undefined
            ? {
                native: {
                  ...(morphKey === undefined ? {} : { morphKey }),
                  ...(requestedName === undefined
                    ? {}
                    : { hasExplicitName: true }),
                  ...(link === undefined ? {} : { link }),
                },
              }
            : {}),
          resolve(resolveContext) {
            const resolved = resolveContext.resolveChildren(input.children);
            const boxes = resolved.objects
              .map((member) => member.box)
              .filter((b): b is NonNullable<typeof b> => b !== undefined);
            const x = Math.min(...boxes.map((b) => b.x));
            const y = Math.min(...boxes.map((b) => b.y));
            const right = Math.max(...boxes.map((b) => b.x + b.width));
            const bottom = Math.max(...boxes.map((b) => b.y + b.height));
            const box = resolveContext.layoutBox ?? {
              x,
              y,
              width: right - x,
              height: bottom - y,
            };
            const name =
              resolveContext.morphName ??
              requestedName ??
              `Cluster ${resolveContext.id}`;
            return {
              observation: {
                id: resolveContext.id,
                kind: "group",
                name,
                box,
                members: resolved.objects,
              },
              extensions: [],
              emit(emitContext) {
                return {
                  xml: `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${emitContext.id}" name="${xmlEscape(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/><a:chOff x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:chExt cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm></p:grpSpPr>${emitContext.childrenXml}</p:grpSp>`,
                };
              },
            };
          },
        };
      },
    };

    const clusterMorph = createMorphScope().handle("cluster-morph");
    const clusterBuild = runJsxModules([
      customClusterModule,
      boxCapability,
      shapeCapability,
    ]);
    const valid = await clusterBuild(
      <Deck width={400} height={240} theme={{ colors: { tone: "#111111" } }}>
        <Slide id="cluster-slide">
          <CustomCluster
            id="loose-cluster"
            theme={{ colors: { tone: "#DC2626" } }}
          >
            <Shape id="c1" x={15} y={25} width={40} height={30} fill="tone" />
            <Shape
              id="c2"
              x={75}
              y={40}
              width={30}
              height={25}
              fill="#0F766E"
            />
          </CustomCluster>
          <Box
            id="cluster-host"
            style={{
              position: "absolute",
              left: 20,
              top: 120,
              display: "flex",
              flexDirection: "row",
              gap: 12,
            }}
          >
            <CustomCluster
              id="layout-cluster"
              morphKey={clusterMorph}
              style={{ display: "flex", width: 120, height: 60 }}
            >
              <Shape id="lc1" width={50} height={40} fill="#1D4ED8" />
            </CustomCluster>
            <Shape id="after-cluster" width={40} height={40} fill="#F59E0B" />
          </Box>
        </Slide>
      </Deck>,
    );

    assert.deepEqual(valid.report.diagnostics, []);
    assert.deepEqual((await validatePptx(valid.pptx)).errors, []);
    const looseObs = valid.query
      .getSlide("cluster-slide")
      .get("loose-cluster")
      .observe();
    assert.deepEqual(looseObs.box, { x: 15, y: 25, width: 90, height: 40 });
    assert.equal(looseObs.objects[0]?.kind, "group");
    assert.equal(looseObs.objects[0]?.members?.length, 2);

    const afterObs = valid.query
      .getSlide("cluster-slide")
      .get("cluster-host")
      .get("after-cluster")
      .observe();
    assert.deepEqual(afterObs.box, { x: 152, y: 120, width: 40, height: 40 });
    const unzipped = unzipSync(valid.pptx);
    assert.match(strFromU8(unzipped["ppt/slides/slide1.xml"]!), /DC2626/u);

    // Members inside a childComposition="members" / restrictsMemberContinuity container cannot declare morphKey.
    await assert.rejects(
      clusterBuild(
        <Deck width={400} height={240}>
          <Slide id="bad-cluster-member-morph">
            <CustomCluster id="bad-cluster">
              <Shape
                id="bad-member"
                morphKey={clusterMorph}
                x={0}
                y={0}
                width={30}
                height={30}
              />
            </CustomCluster>
          </Slide>
        </Deck>,
      ),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "morph-inside-group"),
        );
        return true;
      },
    );
  });

  it("separates resolve, emit, and paint stages without resource/rId conflicts, and paints without running emit", async () => {
    const png1x1 = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
      "base64",
    );
    let emitCalls = 0;
    let paintCalls = 0;
    const allocatedFieldIds: string[] = [];
    const resolveKeysSeen: string[][] = [];

    const newTileModule: BoundObjectCapabilityModule = {
      declaration: "new-tile",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(layoutStyleKeys),
          path: context.path,
        });
        const label =
          typeof input.props["label"] === "string"
            ? input.props["label"]
            : "tile";
        const box = {
          x: typeof input.props["x"] === "number" ? input.props["x"] : 20,
          y: typeof input.props["y"] === "number" ? input.props["y"] : 20,
          width:
            typeof input.props["width"] === "number"
              ? input.props["width"]
              : 100,
          height:
            typeof input.props["height"] === "number"
              ? input.props["height"]
              : 40,
        };
        const rotate =
          typeof input.props["rotate"] === "number" ? input.props["rotate"] : 0;
        const link = input.props["link"] as LinkTarget | undefined;
        const attachImage = input.props["attachImage"] === true;

        return {
          contract,
          ...(link === undefined ? {} : { native: { link } }),
          resolve(resolveContext: ObjectResolveContext): ResolvedObject {
            resolveKeysSeen.push(Object.keys(resolveContext));
            const resolvedBox = resolveContext.layoutBox ?? box;
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: `NewTile ${resolveContext.id}`,
                geometry: "rect",
                box: resolvedBox,
                text: label,
              },
              ...(rotate === 0
                ? {}
                : {
                    transform: {
                      box: resolvedBox,
                      angle: rotate,
                      flipH: false,
                      flipV: false,
                    },
                  }),
              extensions: [],
              emit(emitContext: EmitContext) {
                emitCalls += 1;
                const fieldId = emitContext.createTextFieldId();
                allocatedFieldIds.push(fieldId);
                if (attachImage) {
                  emitContext.registerResource({
                    bytes: new Uint8Array(png1x1),
                    extension: "png",
                    contentType: "image/png",
                    relationshipType:
                      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                  });
                }
                const linkXml =
                  emitContext.nativeLink?.relationshipId !== undefined &&
                  emitContext.nativeLink.relationshipId !== ""
                    ? `<a:hlinkClick r:id="${xmlEscape(emitContext.nativeLink.relationshipId)}"${
                        emitContext.nativeLink.action === undefined
                          ? ""
                          : ` action="${xmlEscape(emitContext.nativeLink.action)}"`
                      }/>`
                    : "";
                return {
                  xml:
                    `<p:sp><p:nvSpPr>` +
                    `<p:cNvPr id="${emitContext.id}" name="${xmlEscape(`NewTile ${emitContext.id}`)}">${linkXml}</p:cNvPr>` +
                    `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
                    `<p:spPr>` +
                    `<a:xfrm><a:off x="${pxToEmu(resolvedBox.x)}" y="${pxToEmu(resolvedBox.y)}"/><a:ext cx="${pxToEmu(resolvedBox.width)}" cy="${pxToEmu(resolvedBox.height)}"/></a:xfrm>` +
                    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
                    `<a:solidFill><a:srgbClr val="1D4ED8"/></a:solidFill>` +
                    `</p:spPr>` +
                    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="${xmlEscape(fieldId)}" type="slidenum"><a:t>${emitContext.slideNumber}</a:t></a:fld><a:r><a:t>${xmlEscape(label)}</a:t></a:r></a:p></p:txBody>` +
                    `</p:sp>`,
                };
              },
              paint(paintContext: PaintContext) {
                paintCalls += 1;
                const cx = resolvedBox.x + resolvedBox.width / 2;
                const cy = resolvedBox.y + resolvedBox.height / 2;
                return {
                  svg: `<g transform="rotate(${rotate} ${cx} ${cy})" data-id="${paintContext.objectId}"><rect x="${resolvedBox.x}" y="${resolvedBox.y}" width="${resolvedBox.width}" height="${resolvedBox.height}" fill="#1D4ED8"/></g>`,
                };
              },
            };
          },
        };
      },
    };

    const newClusterModule: BoundObjectCapabilityModule = {
      declaration: "new-cluster",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "grpSp",
          props: input.props,
          allowedStyleKeys: new Set<string>(layoutStyleKeys),
          isContainer: true,
          path: context.path,
        });
        return {
          contract,
          container: {
            children: input.children,
            emitsObject: true,
            layoutMode: "when-styled",
            childComposition: "members",
            restrictsMemberContinuity: true,
            observesLayout: false,
            propagatesBaseline: false,
            checksCanvasOverflow: false,
          },
          resolve(resolveContext: ObjectResolveContext): ResolvedObject {
            const members = resolveContext.resolveChildren(input.children);
            const boxes = members.objects
              .map((m) => m.box)
              .filter((b): b is NonNullable<typeof b> => b !== undefined);
            const x = Math.min(...boxes.map((b) => b.x));
            const y = Math.min(...boxes.map((b) => b.y));
            const right = Math.max(...boxes.map((b) => b.x + b.width));
            const bottom = Math.max(...boxes.map((b) => b.y + b.height));
            const box = { x, y, width: right - x, height: bottom - y };
            const name = `NewCluster ${resolveContext.id}`;
            return {
              observation: {
                id: resolveContext.id,
                kind: "group",
                name,
                box,
                members: members.objects,
              },
              extensions: [],
              emit(emitContext: EmitContext) {
                return {
                  xml:
                    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${emitContext.id}" name="${xmlEscape(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
                    `<p:grpSpPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/><a:chOff x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:chExt cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm></p:grpSpPr>` +
                    `${emitContext.childrenXml}</p:grpSp>`,
                };
              },
              paint(paintContext: PaintContext) {
                return {
                  svg: `<g>${paintContext.childrenSvg ?? ""}</g>`,
                };
              },
            };
          },
        };
      },
    };

    const newSlideMetaModule: SlideCapabilityModule = {
      declaration: "new-slide-meta",
      slideProp: "slideMeta",
      slot: "transition",
      bindSlide(input) {
        const enabled = input.props["slideMeta"] === true;
        return {
          resolveSlide(resolveContext: SlideResolveContext): ResolvedSlide {
            assert.equal("registerSharedPart" in resolveContext, false);
            assert.equal("registerPackagePart" in resolveContext, false);
            assert.equal("registerResource" in resolveContext, false);
            assert.equal("createTextFieldId" in resolveContext, false);
            return {
              observation: { enabled },
              extensions: enabled ? ["p14"] : [],
              emit(emitContext: SlideEmitContext) {
                emitContext.registerSharedPart("new-meta-shared", {
                  part: {
                    pathTemplate: "ppt/tags/tag{n}.xml",
                    contentType:
                      "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
                  },
                  content: () =>
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
                });
                return { xml: "" };
              },
            };
          },
        };
      },
    };

    const noPaintModule: BoundObjectCapabilityModule = {
      declaration: "no-paint-tile",
      bind(input, context) {
        const contract = parseCapabilityLayoutContract({
          nativeKind: "sp",
          props: input.props,
          allowedStyleKeys: new Set<string>(layoutStyleKeys),
          path: context.path,
        });
        return {
          contract,
          resolve(resolveContext: ObjectResolveContext): ResolvedObject {
            const box = { x: 10, y: 10, width: 50, height: 25 };
            return {
              observation: {
                id: resolveContext.id,
                kind: "shape",
                name: `NoPaint ${resolveContext.id}`,
                geometry: "rect",
                box,
              },
              extensions: [],
              emit(emitContext: EmitContext) {
                return {
                  xml:
                    `<p:sp><p:nvSpPr><p:cNvPr id="${emitContext.id}" name="NoPaint"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
                    `<p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`,
                };
              },
            };
          },
        };
      },
    };

    const modules = [
      shapeCapability,
      boxCapability,
      groupCapability,
      imageCapability,
      motionCapability,
      notesCapability,
      newTileModule,
      newClusterModule,
      newSlideMetaModule,
      noPaintModule,
    ] as const;

    const deckNode = createDeckNode({ width: 960, height: 540 });
    const slide1 = createSlideNode({
      id: "mixed-1",
      slideMeta: true,
    } as never);
    const slide2 = createSlideNode({ id: "mixed-2" });

    const newTile1 = new CommonNode("new-tile", "new-tile", {
      id: "nt1",
      label: "Hello",
      x: 20,
      y: 30,
      width: 120,
      height: 40,
      rotate: 30,
      attachImage: true,
      link: { id: "mixed-2" },
    });
    const oldShape = impShape({
      id: "old-s1",
      x: 200,
      y: 30,
      width: 100,
      height: 50,
      fill: "#16A34A",
      link: { url: "https://example.com" },
      enter: fade(),
    });
    const oldImage = impImage({
      id: "old-img",
      src: new Uint8Array(png1x1),
      x: 320,
      y: 30,
      width: 40,
      height: 40,
    });
    const mixedCluster = new CommonContainer("new-cluster", "new-cluster", {
      id: "mixed-cluster",
    });
    mixedCluster.append(
      impShape({
        id: "cluster-old",
        x: 40,
        y: 120,
        width: 60,
        height: 40,
        fill: "#DC2626",
      }),
      new CommonNode("new-tile", "new-tile", {
        id: "cluster-new",
        label: "InsideCluster",
        x: 120,
        y: 120,
        width: 80,
        height: 40,
      }),
    );
    const oldGroupWithNewChild = impGroup({ id: "old-group" });
    oldGroupWithNewChild.append(
      new CommonNode("new-tile", "new-tile", {
        id: "inside-old-group",
        label: "InOldGroup",
        x: 360,
        y: 120,
        width: 80,
        height: 40,
      }),
    );
    slide1.append(
      newTile1,
      oldShape,
      oldImage,
      mixedCluster,
      oldGroupWithNewChild,
    );
    slide2.append(
      impShape({
        id: "s2-shape",
        x: 20,
        y: 20,
        width: 80,
        height: 40,
        fill: "#0EA5E9",
      }),
    );
    deckNode.append(slide1, slide2);

    const resolve = createResolver(modules);
    const resolution = resolve({ children: [deckNode] });

    assert.equal(Object.isFrozen(resolution), true);
    assert.deepEqual(Object.keys(resolution), ["report", "query"]);
    assert.equal(Object.getOwnPropertySymbols(resolution).length, 0);

    // Neither emit nor paint ran during resolve.
    assert.equal(emitCalls, 0);
    assert.equal(paintCalls, 0);
    assert.ok(
      resolveKeysSeen.every(
        (keys) =>
          !keys.includes("registerResource") &&
          !keys.includes("registerPackagePart") &&
          !keys.includes("createTextFieldId"),
      ),
    );

    // Mutating the editable CommonNode after resolve must not affect the Resolution snapshot.
    slide2.append(
      impShape({
        id: "mutated-after-resolve",
        x: 100,
        y: 100,
        width: 50,
        height: 50,
      }),
    );

    // Paint BEFORE emit: must succeed without emit having run.
    const painted = paint(resolution);
    assert.equal(emitCalls, 0);
    assert.equal(paintCalls, 3);
    assert.equal(painted.pages.length, 2);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);
    const rotatedNode = painted.pages[0]?.nodes.find(
      (n) => n.authorId === "nt1",
    );
    assert.ok(rotatedNode?.screenBox !== undefined);
    assert.ok(rotatedNode.screenBox.width > 120);
    assert.ok(rotatedNode.screenBox.height > 40);
    assert.equal(painted.pages[1]?.nodes.length, 1);

    // Emit twice on the same Resolution: idempotent, no rId/field-id/resource conflicts.
    const emitted1 = emit(resolution);
    const emitted2 = emit(resolution);
    assert.deepEqual((await validatePptx(emitted1.pptx)).errors, []);
    assert.deepEqual((await validatePptx(emitted2.pptx)).errors, []);
    assert.equal(emitted1.report.slides.length, 2);
    assert.equal(
      resolution.query.getSlide("mixed-2").observe().objects.length,
      1,
    );
    const mediaParts = Object.keys(unzipSync(emitted1.pptx)).filter((part) =>
      /^ppt\/media\/resource\d+\.png$/u.test(part),
    );
    assert.equal(mediaParts.length, 1);
    assert.equal(
      new Set(allocatedFieldIds.slice(0, 3)).size,
      allocatedFieldIds.slice(0, 3).length,
    );
    assert.ok(observeMotion(emitted1.report.slides[0]!) !== undefined);
    assert.deepEqual(
      emitted1.report.slides[0]?.capabilities?.["new-slide-meta"],
      {
        enabled: true,
      },
    );

    // Target check happens in emit, while resolve + paint still succeed for baselineOnly.
    const baselineResolution = resolve({ children: [deckNode] });
    assert.equal(paint(baselineResolution).pages.length, 2);
    assert.throws(
      () => emit(baselineResolution, { target: targets.baselineOnly }),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(
          error.diagnostics.some((d) => d.code === "forbidden-mechanism"),
        );
        return true;
      },
    );

    // Missing paint() produces a page error diagnostic and omits svg/png without blocking emit().
    const noPaintDeck = createDeckNode({ width: 400, height: 300 });
    const noPaintSlide = createSlideNode({ id: "no-paint" });
    noPaintSlide.append(
      new CommonNode("no-paint-tile", "no-paint-tile", { id: "np1" }),
    );
    noPaintDeck.append(noPaintSlide);
    const noPaintResolution = resolve({ children: [noPaintDeck] });
    const noPaintPainted = paint(noPaintResolution);
    assert.equal(noPaintPainted.pages[0]?.svg, undefined);
    assert.equal(noPaintPainted.pages[0]?.png, undefined);
    assert.ok(
      noPaintPainted.pages[0]?.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          d.message.includes("Preview does not support no-paint-tile"),
      ),
    );
    const noPaintEmitted = emit(noPaintResolution);
    assert.deepEqual((await validatePptx(noPaintEmitted.pptx)).errors, []);

    // New-contract background SlideCapabilityModule: paints when paint() is present and reports unsupported when omitted.
    const newBgSupported: SlideCapabilityModule = {
      declaration: "new-bg",
      slideProp: "newBg",
      slot: "background",
      bindSlide(input) {
        const color =
          typeof input.props["newBg"] === "string"
            ? input.props["newBg"]
            : "#ffffff";
        return {
          resolveSlide(): ResolvedSlide {
            return {
              extensions: [],
              emit: () => ({ xml: "" }),
              paint: () => ({
                svg: `<rect width="100%" height="100%" fill="${color}"/>`,
              }),
            };
          },
        };
      },
    };
    const newBgUnsupported: SlideCapabilityModule = {
      declaration: "new-bg-nopaint",
      slideProp: "newBgNoPaint",
      slot: "background",
      bindSlide() {
        return {
          resolveSlide(): ResolvedSlide {
            return {
              extensions: [],
              emit: () => ({ xml: "" }),
            };
          },
        };
      },
    };
    const bgOkDeck = createDeckNode({ width: 400, height: 300 });
    bgOkDeck.append(
      createSlideNode({ id: "bg-ok", newBg: "#123456" } as never),
    );
    const bgOkPainted = paint(
      createResolver([newBgSupported])({ children: [bgOkDeck] }),
    );
    assert.match(bgOkPainted.pages[0]?.svg ?? "", /#123456/u);
    assert.ok(bgOkPainted.pages[0]?.png !== undefined);

    const bgBadDeck = createDeckNode({ width: 400, height: 300 });
    bgBadDeck.append(
      createSlideNode({ id: "bg-bad", newBgNoPaint: true } as never),
    );
    const bgBadPainted = paint(
      createResolver([newBgUnsupported])({ children: [bgBadDeck] }),
    );
    assert.equal(bgBadPainted.pages[0]?.svg, undefined);
    assert.equal(bgBadPainted.pages[0]?.png, undefined);
    assert.ok(
      bgBadPainted.pages[0]?.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          d.message.includes("Preview does not support new-bg-nopaint"),
      ),
    );
  });

  it("migrates box, shape, and group to resolve/emit/paint without emit-coupled preview or paint-returned transforms", async () => {
    const resolve = createResolver([
      boxCapability,
      shapeCapability,
      groupCapability,
    ]);

    // 1. XML 1.0 control characters in Shape.name and Group.name fail in resolve() with non-doubled paths.
    for (const [badDoc, expectedPath] of [
      [
        impDeck({ width: 320, height: 180 }).append(
          impSlide({ id: "bad-shape-name" }).append(
            impShape({
              id: "s1",
              name: "Bad\u0007Shape",
              x: 10,
              y: 10,
              width: 40,
              height: 20,
            }),
          ),
        ),
        "deck/slide[bad-shape-name]/shape[s1]",
      ],
      [
        impDeck({ width: 320, height: 180 }).append(
          impSlide({ id: "bad-group-name" }).append(
            impGroup({ id: "g1", name: "Bad\u0008Group" }).append(
              impShape({ id: "s1", x: 10, y: 10, width: 40, height: 20 }),
            ),
          ),
        ),
        "deck/slide[bad-group-name]/group[g1]",
      ],
    ] as const) {
      assert.throws(
        () => resolve({ children: [badDoc] }),
        (error: unknown) => {
          assert(error instanceof BuildError);
          assert.equal(error.diagnostics[0]?.code, "invalid-xml-character");
          assert.equal(error.diagnostics[0]?.path, expectedPath);
          return true;
        },
      );
    }

    // Box errors in resolve() or non-layout bind() also report exact single node paths (never doubled).
    const wrapperCapability: CapabilityModule = {
      declaration: "custom-wrapper",
      bind(input) {
        return {
          contract: {
            nativeKind: "grpSp",
            layoutStyle: {},
          },
          resolve(resolveContext) {
            resolveContext.resolveChildren(input.children);
            return {
              observation: {
                id: resolveContext.id,
                kind: "group",
                name: "Wrap",
                box: { x: 0, y: 0, width: 10, height: 10 },
              },
              extensions: [],
              emit() {
                return { xml: "" };
              },
            };
          },
        };
      },
    };
    const resolveWithWrapper = createResolver([
      wrapperCapability,
      boxCapability,
    ]);
    const unlaidOutBoxDoc = impDeck({ width: 320, height: 180 }).append(
      impSlide({ id: "bad-box-slide" }).append(
        new CommonContainer("wrap", "custom-wrapper", { id: "wrap" }).append(
          impBox({ id: "b1", style: { backgroundColor: "#112233" } }),
        ),
      ),
    );
    assert.throws(
      () => resolveWithWrapper({ children: [unlaidOutBoxDoc] }),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.equal(error.diagnostics[0]?.code, "layout-unavailable");
        assert.equal(
          error.diagnostics[0]?.path,
          "deck/slide[bad-box-slide]/custom-wrapper[wrap]/box[b1]",
        );
        return true;
      },
    );

    // 2. Empty Group fails in resolve().
    assert.throws(
      () =>
        resolve({
          children: [
            impDeck({ width: 320, height: 180 }).append(
              impSlide({ id: "empty-group-slide" }).append(
                impGroup({ id: "empty" }),
              ),
            ),
          ],
        }),
      (error: unknown) => {
        assert(error instanceof BuildError);
        assert.ok(error.diagnostics.some((d) => d.code === "empty-group"));
        return true;
      },
    );

    // 3. Screen geometry depends only on resolve facts: even when a rotated member's paint() fails
    //    due to a corrupt image stream, its own rotation + enclosing group rotation still produce exact corners and screenBox.
    const corruptPng = new Uint8Array(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
        "base64",
      ),
    );
    const idatOffset = Buffer.from(corruptPng).indexOf(Buffer.from("IDAT"));
    corruptPng[idatOffset + 6] = (corruptPng[idatOffset + 6] ?? 0) ^ 0xff;
    corruptPng[idatOffset + 7] = (corruptPng[idatOffset + 7] ?? 0) ^ 0xff;

    const resolveWithImage = createResolver([
      boxCapability,
      shapeCapability,
      groupCapability,
      imageCapability,
    ]);
    const doc = impDeck({ width: 480, height: 320 }).append(
      impSlide({ id: "failed-paint-transforms" }).append(
        impGroup({ id: "rg", style: { rotate: 90 } }).append(
          impImage({
            id: "corrupt-member",
            src: corruptPng,
            x: 100,
            y: 100,
            width: 40,
            height: 40,
            style: { rotate: 45 },
          }),
          impShape({
            id: "rect-member",
            x: 160,
            y: 120,
            width: 40,
            height: 20,
            fill: "#16A34A",
          }),
        ),
      ),
    );
    const resolution = resolveWithImage({ children: [doc] });
    const painted = paint(resolution);
    const page = painted.pages[0]!;
    assert.equal(page.png, undefined);
    assert.ok(
      page.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /decode PNG/i.test(d.message) &&
          d.path.includes("image[corrupt-member]"),
      ),
    );
    // Group [100..200, 100..140] center (150, 120); corrupt-member center (120, 120) -> rotates 90° CW to (150, 90).
    const corruptNode = page.nodes.find(
      (n) => n.authorId === "corrupt-member",
    )!;
    const halfDiag = 20 * Math.SQRT2;
    assert.ok(corruptNode.screenBox !== undefined);
    assert.ok(Math.abs(corruptNode.screenBox.x - (150 - halfDiag)) < 1e-6);
    assert.ok(Math.abs(corruptNode.screenBox.y - (90 - halfDiag)) < 1e-6);
    assert.ok(Math.abs(corruptNode.screenBox.width - 2 * halfDiag) < 1e-6);
    assert.ok(Math.abs(corruptNode.screenBox.height - 2 * halfDiag) < 1e-6);

    // 4. resolve() without emit options + independent emit() with subset font embedding & Shape inline text,
    //    even when a sibling corrupt image fails paint().
    const jostFont = impFont("Jost*", [
      {
        weight: 400,
        data: new Uint8Array(
          readFileSync(
            new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url),
          ),
        ),
      },
    ]);
    const textShapeDoc = impDeck({ width: 480, height: 270 }).append(
      impSlide({ id: "s1" }).append(
        impBox({
          id: "card",
          style: {
            width: 240,
            height: 80,
            padding: 8,
            backgroundColor: "#F8FAFC",
          },
        }).append(
          impShape({
            id: "shaped-text",
            width: 200,
            height: 48,
            fill: "#1E293B",
            style: { fontFamily: jostFont, fontSize: 16, color: "#FFFFFF" },
            text: "Resolved once",
          }),
        ),
        impImage({
          id: "corrupt-blocker",
          src: corruptPng,
          x: 10,
          y: 120,
          width: 20,
          height: 20,
        }),
      ),
    );
    const textResolution = resolveWithImage({ children: [textShapeDoc] });
    const textPainted = paint(textResolution);
    assert.equal(textPainted.pages[0]?.png, undefined);
    assert.ok(
      textPainted.pages[0]?.diagnostics.some(
        (d) =>
          d.severity === "error" &&
          /decode PNG/i.test(d.message) &&
          d.path.includes("image[corrupt-blocker]"),
      ),
    );
    const textEmitted = emit(textResolution, { embedFonts: "subset" });
    assert.deepEqual((await validatePptx(textEmitted.pptx)).errors, []);
    const files = unzipSync(textEmitted.pptx);
    const slideXml = strFromU8(files["ppt/slides/slide1.xml"]!);
    assert.match(slideXml, /Resolved once/u);
  });

  it("migrates imageCapability to BoundObjectCapability with resolve-only placement/XML1.0 checks, paint without emit, and emit-only resource registration", async () => {
    const png1x1 = new Uint8Array(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
        "base64",
      ),
    );
    const jpeg1x1 = new Uint8Array(
      Buffer.from(
        "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAFR/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIRAxEAPwD/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oACAEBAAE/IV//2Q==",
        "base64",
      ),
    );

    const resolveImages = createResolver([imageCapability]);
    const deckNode = createDeckNode({ width: 640, height: 360 });
    const slide1 = createSlideNode({ id: "img-s1" });
    slide1.append(
      impImage({
        id: "hero-png",
        name: "Hero PNG",
        alt: "Square sample",
        src: png1x1,
        fit: "contain",
        x: 20,
        y: 40,
        width: 120,
        height: 60,
        style: { rotate: 30 },
        flipH: true,
        link: { id: "img-s2" },
      }),
    );
    const slide2 = createSlideNode({ id: "img-s2" });
    slide2.append(
      impImage({
        id: "thumb-jpeg",
        src: jpeg1x1,
        crop: { left: 0.25 },
        x: 50,
        y: 50,
        width: 80,
        height: 80,
      }),
    );
    deckNode.append(slide1, slide2);

    const resolution = resolveImages({ children: [deckNode] });

    // Placement (contain of 1x1 into 120x60 -> 60x60 centered at x=50) happens in resolve;
    // resource registration does NOT happen during resolve.
    const heroObs = resolution.query
      .getSlide("img-s1")
      .get("hero-png")
      .observe();
    assert.deepEqual(heroObs.box, { x: 50, y: 40, width: 60, height: 60 });
    assert.equal(heroObs.resources, undefined);

    // Paint succeeds directly from Resolution before emit() runs, and computes screenBox from transform.
    const painted = paint(resolution);
    assert.equal(painted.pages.length, 2);
    assert.deepEqual(painted.pages[0]?.diagnostics, []);
    assert.ok(painted.pages[0]?.png !== undefined);
    assert.match(painted.pages[0]?.svg ?? "", /<image/u);
    const paintedHero = painted.pages[0]?.nodes.find(
      (n) => n.authorId === "hero-png",
    );
    assert.ok(paintedHero?.screenBox !== undefined);
    assert.ok(paintedHero.screenBox.width > 60);
    assert.ok(paintedHero.screenBox.height > 60);

    // Emit registers media resources and produces valid PPTX; repeat emit is idempotent.
    const emitted1 = emit(resolution);
    const emitted2 = emit(resolution);
    const val1 = await validatePptx(emitted1.pptx);
    assert.deepEqual(val1.errors, []);
    assert.deepEqual((await validatePptx(emitted2.pptx)).errors, []);
    const heroImagePart = val1.slides[0]?.objects[0]?.image?.part ?? "";
    assert.match(heroImagePart, /^ppt\/media\/resource\d+\.png$/u);
    const thumbImagePart = val1.slides[1]?.objects[0]?.image?.part ?? "";
    assert.match(thumbImagePart, /^ppt\/media\/resource\d+\.jpeg$/u);

    // Invalid XML 1.0 characters in name or alt fail immediately during resolve.
    for (const badProps of [
      { name: "bad\u0000name", src: png1x1 },
      { alt: "bad\u000balt", src: png1x1 },
    ]) {
      const badDeck = createDeckNode({ width: 320, height: 180 });
      const badSlide = createSlideNode({ id: "bad-xml" });
      badSlide.append(impImage({ id: "bad-img", ...badProps }));
      badDeck.append(badSlide);
      assert.throws(
        () => resolveImages({ children: [badDeck] }),
        (error: unknown) =>
          error instanceof BuildError &&
          error.diagnostics.some((d) => d.code === "invalid-xml-character"),
      );
    }
  });
});
