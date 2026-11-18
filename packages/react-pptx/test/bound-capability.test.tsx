import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  CommonContainer,
  CommonNode,
  createBuilder,
  createMorphScope,
  Deck,
  Slide,
  type BoundContainerCollaboration,
  type BoundObjectCapability,
  type BoundObjectCapabilityModule,
  type CapabilityInput,
  type CapabilityParseContext,
  type LinkTarget,
} from "@react-pptx/core";
import {
  CapabilityError,
  isMorphKey,
  layoutStyleKeys,
  parseCapabilityLayoutContract,
  pxToEmu,
  textStyleFrom,
  textStyleKeys,
  xmlEscape,
  type MorphKey,
  type ThemeScope,
} from "@react-pptx/core/capability";
import {
  createDeckNode,
  createDocumentBuild,
  createSlideNode,
  hostNodeCreator,
} from "@react-pptx/core/document";
import {
  Box,
  boxCapability,
  fade,
  font,
  Group,
  groupCapability,
  motionCapability,
  P,
  preset,
  Shape,
  shapeCapability,
  Span,
  transitionCapability,
  transitions,
} from "@react-pptx/react-pptx";
import {
  box as impBox,
  createDocumentBuild as createDocBuild,
  createMorphScope as createDocMorphScope,
  deck as impDeck,
  font as impFont,
  group as impGroup,
  paragraph as impParagraph,
  shape as impShape,
  slide as impSlide,
  span as impSpan,
} from "@react-pptx/react-pptx/document";
import { validatePptx } from "@react-pptx/validation";
import { strFromU8, unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it, vi } from "vitest";

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
  return new CommonNode("badge", "rp:custom-badge", props);
}

function CustomBadge(props: BadgeProps) {
  return createElement("rp:custom-badge", {
    ...props,
    [hostNodeCreator]: createBadgeNode,
  });
}

function createTrackedBadgeCapability() {
  const bindCalls: string[] = [];
  const measureCalls: string[] = [];
  const compileCalls: string[] = [];

  const module: BoundObjectCapabilityModule = {
    declaration: "rp:custom-badge",
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
        compile(compileContext) {
          compileCalls.push(compileContext.path);
          const box = compileContext.layoutBox ?? {
            x: declaredX,
            y: declaredY,
            width: declaredWidth,
            height: declaredHeight,
          };
          const name =
            compileContext.morphName ??
            privateExplicitName ??
            `Badge ${compileContext.id - 1}`;
          const link = compileContext.nativeLink;
          const linkAttrs = [
            ...(link?.relationshipId !== undefined && link.relationshipId !== ""
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
            `<p:cNvPr id="${compileContext.id}" name="${xmlEscape(name)}">${linkXml}</p:cNvPr>` +
            `<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
            `<p:spPr>` +
            `<a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm>` +
            `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
            `<a:solidFill><a:srgbClr val="2B579A"/></a:solidFill>` +
            `</p:spPr>` +
            `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(privateLabel)}</a:t></a:r></a:p></p:txBody>` +
            `</p:sp>`;
          return {
            xml,
            observation: {
              id: compileContext.id,
              kind: "badge",
              name,
              box,
              text: privateLabel,
            },
            extensions: [],
          };
        },
      };
    },
  };

  return { module, bindCalls, measureCalls, compileCalls };
}

describe("bound object capability contract", () => {
  it("supports custom bound object capability via both JSX and imperative builds without exposing parsed state", async () => {
    const tracker = createTrackedBadgeCapability();
    const morph = createMorphScope();
    const badgeMorph = morph.handle("badge-morph");

    // 1. Imperative build via createDocumentBuild
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

    const imperativeBuild = createDocumentBuild([tracker.module]);
    const imperativeResult = await imperativeBuild(deckNode);
    const impValidation = await validatePptx(imperativeResult.pptx);
    assert.deepEqual(impValidation.errors, []);

    assert.equal(tracker.bindCalls.length, 1);
    assert.equal(tracker.compileCalls.length, 1);
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
      /^!!react-pptx-morph-/u,
    );

    // 2. JSX build inside legacy Box Flex container (exercises measureLayout + container collaboration)
    tracker.bindCalls.length = 0;
    tracker.measureCalls.length = 0;
    tracker.compileCalls.length = 0;

    const jsxBuild = createBuilder([boxCapability, tracker.module]);
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
    assert.equal(tracker.compileCalls.length, 1);

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

    const buildWithBadge = createBuilder([tracker.module]);
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

    const trimmedBuild = createBuilder([]);
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
    assert.equal("bind" in shapeCapability, true);
    assert.equal("parse" in shapeCapability, false);
    assert.equal("compile" in shapeCapability, false);

    const bindSpy = vi.spyOn(shapeCapability, "bind");
    try {
      const trimmedShapeBuild = createBuilder([shapeCapability]);
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

      assert.equal(bindSpy.mock.calls.length, 1);
      assert.deepEqual(result.report.diagnostics, []);
      const validation = await validatePptx(result.pptx);
      assert.deepEqual(validation.errors, []);
      assert.equal(validation.slides[0]?.objects[0]?.text?.text, "Bound Shape");
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("produces consistent native PPTX and observations for Shape across JSX and imperative entries inside unmigrated containers", async () => {
    const jsxMorph = createMorphScope().handle("shape-card");
    const impMorph = createDocMorphScope().handle("shape-card");

    const jsxBuild = createBuilder([
      shapeCapability,
      boxCapability,
      groupCapability,
      transitionCapability,
      motionCapability,
    ]);
    const docBuild = createDocBuild([
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
    const bindSpy = vi.spyOn(shapeCapability, "bind");
    try {
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
      const docBuild = createDocBuild([shapeCapability]);

      const first = await docBuild(deckNode);
      assert.equal(bindSpy.mock.calls.length, 1);

      node.set({
        geometry: "ellipse",
        width: 220,
        fill: "#445566",
        text: "Version 2",
      });

      const second = await docBuild(deckNode);
      assert.equal(bindSpy.mock.calls.length, 2);

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
    } finally {
      bindSpy.mockRestore();
    }
  });

  it("validates Shape link declarations in bind and resolves valid links in core", async () => {
    const shapeBuild = createBuilder([shapeCapability]);
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

  it("migrates boxCapability to BoundObjectCapabilityModule with container role and single bind per build", async () => {
    assert.equal("bind" in boxCapability, true);
    assert.equal("parse" in boxCapability, false);
    assert.equal("compile" in boxCapability, false);

    const bindSpy = vi.spyOn(boxCapability, "bind");
    try {
      const trimmedBuild = createBuilder([boxCapability, shapeCapability]);
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

      assert.equal(bindSpy.mock.calls.length, 1);
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
    } finally {
      bindSpy.mockRestore();
    }

    const withoutBoxBuild = createBuilder([shapeCapability]);
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
      new CommonContainer("panel", "rp:custom-panel", props);
    const CustomPanel = (
      props: PanelProps & { readonly children?: React.ReactNode },
    ) =>
      createElement(
        "rp:custom-panel",
        { ...props, [hostNodeCreator]: createPanelNode },
        props.children,
      );

    const panelAllowedStyle = new Set([
      ...layoutStyleKeys,
      ...textStyleKeys,
      "padding",
    ]);
    const customPanelModule: BoundObjectCapabilityModule = {
      declaration: "rp:custom-panel",
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
          compile(compileContext) {
            const box = compileContext.layoutBox!;
            const name =
              compileContext.morphName ?? `Panel ${compileContext.id}`;
            return {
              xml: emitsObject
                ? `<p:sp><p:nvSpPr><p:cNvPr id="${compileContext.id}" name="${xmlEscape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></p:spPr></p:sp>`
                : "",
              observation: {
                id: compileContext.id,
                kind: "shape",
                name,
                geometry: "rect",
                box,
              },
              extensions: [],
            };
          },
        };
      },
    };

    const panelMorph = createMorphScope().handle("panel-morph");
    const panelBuild = createBuilder([customPanelModule, shapeCapability]);
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

  it("migrates groupCapability to BoundObjectCapabilityModule with single bind per node per build", async () => {
    assert.equal("bind" in groupCapability, true);
    assert.equal("parse" in groupCapability, false);
    assert.equal("compile" in groupCapability, false);

    const bindSpy = vi.spyOn(groupCapability, "bind");
    try {
      const trimmedBuild = createBuilder([groupCapability, shapeCapability]);
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

      assert.equal(bindSpy.mock.calls.length, 1);
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
    } finally {
      bindSpy.mockRestore();
    }

    const withoutGroupBuild = createBuilder([shapeCapability]);
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
      new CommonContainer("cluster", "rp:custom-cluster", props);
    const CustomCluster = (
      props: ClusterProps & { readonly children?: React.ReactNode },
    ) =>
      createElement(
        "rp:custom-cluster",
        { ...props, [hostNodeCreator]: createClusterNode },
        props.children,
      );

    const clusterAllowedStyle = new Set([
      ...layoutStyleKeys,
      ...textStyleKeys,
      "padding",
    ]);
    const customClusterModule: BoundObjectCapabilityModule = {
      declaration: "rp:custom-cluster",
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
          compile(compileContext) {
            const compiled = compileContext.compileChildren(input.children);
            const boxes = compiled.objects
              .map((member) => member.box)
              .filter((b): b is NonNullable<typeof b> => b !== undefined);
            const x = Math.min(...boxes.map((b) => b.x));
            const y = Math.min(...boxes.map((b) => b.y));
            const right = Math.max(...boxes.map((b) => b.x + b.width));
            const bottom = Math.max(...boxes.map((b) => b.y + b.height));
            const box = compileContext.layoutBox ?? {
              x,
              y,
              width: right - x,
              height: bottom - y,
            };
            const name =
              compileContext.morphName ??
              requestedName ??
              `Cluster ${compileContext.id}`;
            return {
              xml: `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${compileContext.id}" name="${xmlEscape(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:ext cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/><a:chOff x="${pxToEmu(box.x)}" y="${pxToEmu(box.y)}"/><a:chExt cx="${pxToEmu(box.width)}" cy="${pxToEmu(box.height)}"/></a:xfrm></p:grpSpPr>${compiled.xml}</p:grpSp>`,
              observation: {
                id: compileContext.id,
                kind: "group",
                name,
                box,
                members: compiled.objects,
              },
              extensions: [],
            };
          },
        };
      },
    };

    const clusterMorph = createMorphScope().handle("cluster-morph");
    const clusterBuild = createBuilder([
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
});
