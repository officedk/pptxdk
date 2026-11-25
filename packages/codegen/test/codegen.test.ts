import assert from "node:assert/strict";

import {
  createMorphScope,
  defaultTheme,
  defineSlideLayout,
  defineTheme,
  lighter,
} from "@pptxdk/core";
import type { RegisteredCapabilityModule } from "@pptxdk/core/capability";
import {
  CommonNode,
  RichTextContainer,
  createDeckNode,
  createSectionNode,
  createSlideLayoutNode,
  createSlideMasterNode,
  createSlideNode,
} from "@pptxdk/core/document";
import { describe, it } from "vitest";

import {
  CodegenError,
  codegenDiagnosticCodeSchema,
  codegenDiagnosticSchema,
  codegenOptionsSchema,
  codegenTargetSchema,
  createCodegen,
  defineObjectCodegen,
  defineSlideCodegen,
  expr,
} from "../src/index.ts";

describe("@pptxdk/codegen core contract", () => {
  it("parses supported targets, options, and diagnostics via exported Zod schemas", () => {
    assert.equal(codegenTargetSchema.parse("vanilla"), "vanilla");
    assert.equal(codegenTargetSchema.parse("react"), "react");
    assert.equal(codegenTargetSchema.safeParse("other").success, false);
    assert.deepEqual(codegenOptionsSchema.parse({ target: "react" }), {
      target: "react",
    });
    assert.equal(
      codegenOptionsSchema.safeParse({ target: "other" }).success,
      false,
    );
    assert.equal(
      codegenDiagnosticCodeSchema.parse("invalid-options"),
      "invalid-options",
    );
  });

  it("generates deterministic vanilla and react entry files for sectioned empty decks", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSectionNode({ name: "Overview" }).append(
        createSlideNode({ id: "intro" }),
      ),
      createSectionNode({ name: "Empty Section" }),
    );

    const generate = createCodegen([]);
    const vanilla = generate(doc, { target: "vanilla" });
    assert.equal(vanilla.target, "vanilla");
    assert.equal(vanilla.entry, "deck.ts");
    assert.deepEqual(Object.keys(vanilla.files), ["deck.ts"]);
    assert.equal(typeof vanilla.files["deck.ts"], "string");
    assert.match(vanilla.files["deck.ts"] as string, /export default deck\(/u);
    assert.match(
      vanilla.files["deck.ts"] as string,
      /section\(\{\s*name:\s*"Overview"/u,
    );
    assert.match(
      vanilla.files["deck.ts"] as string,
      /section\(\{\s*name:\s*"Empty Section"/u,
    );

    const react = generate(doc, { target: "react" });
    assert.equal(react.target, "react");
    assert.equal(react.entry, "deck.tsx");
    assert.deepEqual(Object.keys(react.files), ["deck.tsx"]);
    assert.match(
      react.files["deck.tsx"] as string,
      /export default function GeneratedDeck\(\)/u,
    );
    assert.match(
      react.files["deck.tsx"] as string,
      /<Section name="Overview">/u,
    );
    assert.match(
      react.files["deck.tsx"] as string,
      /<Section name="Empty Section" \/>/u,
    );
  });

  it("throws CodegenError with precise paths on invalid input, unsupported declarations, unsupported slide props, and unexpressible values", () => {
    const generate = createCodegen([]);

    assert.throws(
      () => generate(createSlideNode({ id: "s1" }), { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck" &&
        !("severity" in error.diagnostics[0]),
    );

    const invalidDeckProp = createDeckNode({
      width: 960,
      height: 540,
      unknownDeckProp: true,
    } as Record<string, unknown>).append(createSlideNode({ id: "s1" }));
    assert.throws(
      () => generate(invalidDeckProp, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path === "deck@unknownDeckProp" &&
        error.diagnostics[0]?.field === "unknownDeckProp" &&
        !("severity" in error.diagnostics[0]),
    );

    const validEmptyDeck = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }),
    );
    assert.throws(
      () =>
        generate(validEmptyDeck, {
          target: "invalid" as "vanilla",
        }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-options" &&
        error.diagnostics[0]?.path === "deck" &&
        error.diagnostics[0]?.field === "target" &&
        !("severity" in error.diagnostics[0]),
    );

    const mixedDeck = createDeckNode({ width: 960, height: 540 }).append(
      createSectionNode({ name: "Sec" }).append(createSlideNode({ id: "s1" })),
      createSlideNode({ id: "s2" }),
    );
    assert.throws(
      () => generate(mixedDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "invalid-document" &&
        error.diagnostics[0]?.path.startsWith("deck/") &&
        !("severity" in error.diagnostics[0]),
    );

    const unsupportedNodeDeck = createDeckNode({
      width: 960,
      height: 540,
    }).append(
      createSlideNode({ id: "s1" }).append(
        new CommonNode("text", "text", {
          id: "t1",
          style: {
            position: "absolute",
            left: 10,
            top: 10,
            width: 100,
            height: 40,
          },
        }),
      ),
    );
    assert.throws(
      () => generate(unsupportedNodeDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-declaration" &&
        error.diagnostics[0]?.declaration === "text" &&
        error.diagnostics[0]?.path === "deck/slide[s1]/text[t1]" &&
        !("severity" in error.diagnostics[0]),
    );

    const unsupportedSlidePropDeck = createDeckNode({
      width: 960,
      height: 540,
    }).append(
      createSlideNode({
        id: "s1",
        transition: { kind: "fade" },
      }),
    );
    assert.throws(
      () => generate(unsupportedSlidePropDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unsupported-slide-property" &&
        error.diagnostics[0]?.field === "transition" &&
        error.diagnostics[0]?.path === "deck/slide[s1]@transition" &&
        !("severity" in error.diagnostics[0]),
    );

    const unexpressibleValueDeck = createDeckNode({
      width: Number.NaN,
      height: 540,
    }).append(createSlideNode({ id: "s1" }));
    assert.throws(
      () => generate(unexpressibleValueDeck, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.field === "width" &&
        error.diagnostics[0]?.path === "deck@width" &&
        !("severity" in error.diagnostics[0]) &&
        codegenDiagnosticSchema.safeParse(error.diagnostics[0]).success,
    );

    assert.equal(
      codegenDiagnosticSchema.safeParse({
        code: "unsupported-declaration",
        path: "deck/slide[s1]/widget[w1]",
        message: "Unsupported",
      }).success,
      false,
    );
    assert.equal(
      codegenDiagnosticSchema.safeParse({
        code: "unsupported-slide-property",
        path: "deck/slide[s1]@customProp",
        message: "Unsupported prop",
      }).success,
      false,
    );
    assert.equal(
      codegenDiagnosticSchema.safeParse({
        severity: "error",
        code: "invalid-document",
        path: "deck",
        message: "Legacy severity field must be rejected",
      }).success,
      false,
    );

    try {
      generate(createSlideNode({ id: "s1" }), { target: "vanilla" });
      assert.fail("Expected CodegenError");
    } catch (error) {
      assert.ok(error instanceof CodegenError);
      for (const diag of error.diagnostics) {
        assert.deepEqual(codegenDiagnosticSchema.parse(diag), diag);
      }
    }
  });

  it("disambiguates reserved, SDK-imported, non-Latin, and colliding identifiers deterministically", () => {
    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "slide" }),
      createSlideNode({ id: "封面" }),
      createSlideNode({ id: "hero-1" }),
      createSlideNode({ id: "hero_1" }),
    );

    const generate = createCodegen([]);
    const v1 = generate(doc, { target: "vanilla" });
    const v2 = generate(doc, { target: "vanilla" });
    assert.equal(v1.files["deck.ts"], v2.files["deck.ts"]);

    const vanillaSource = v1.files["deck.ts"] as string;
    assert.doesNotMatch(vanillaSource, /\bconst slide = slide\(/u);
    assert.match(vanillaSource, /\bconst slide2 = slide\(/u);
    assert.match(vanillaSource, /id: "封面"/u);
    assert.match(vanillaSource, /\bconst hero1Slide = slide\(/u);
    assert.match(vanillaSource, /\bconst hero1Slide2 = slide\(/u);

    const r1 = generate(doc, { target: "react" });
    const reactSource = r1.files["deck.tsx"] as string;
    assert.doesNotMatch(reactSource, /\bfunction Slide\(/u);
    assert.match(reactSource, /\bfunction Slide2\(/u);
    assert.match(reactSource, /\bfunction Hero1Slide\(/u);
    assert.match(reactSource, /\bfunction Hero1Slide2\(/u);
  });

  it("orders 4-section output, deduplicates assets/themes, packs MorphScopes greedily, and protects JSX whitespace", () => {
    const fakePng = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
    const scopeA = createMorphScope();
    const scopeB = createMorphScope();
    const sharedTheme = defineTheme({
      name: "Brand",
      colors: { brandAccent: lighter("accent2", 0.25) },
    });

    const dummyObjectCap = {
      declaration: "text",
      codegen: defineObjectCodegen({
        declarations: {
          text: {
            emitVanilla: (node, ctx) => {
              const fn = ctx.useImport("text");
              const children = ctx.emitRichTextChildren(node.children);
              const extra = ctx.emitObjectExtraProps(node.props);
              const propsEntries = [
                ...Object.entries(node.props).filter(
                  ([key]) => !(key in extra),
                ),
                ...Object.entries(extra),
              ];
              return ctx.formatVanillaCall(
                fn,
                [ctx.formatProps(propsEntries)],
                children,
              );
            },
            emitReact: (node, ctx) => {
              const comp = ctx.useImport("Text");
              const children = ctx.emitRichTextChildren(node.children);
              const extra = ctx.emitObjectExtraProps(node.props);
              const propsEntries = [
                ...Object.entries(node.props).filter(
                  ([key]) => !(key in extra),
                ),
                ...Object.entries(extra),
              ];
              return ctx.formatJsxElement(comp, propsEntries, children);
            },
          },
          image: {
            emitVanilla: (node, ctx) => {
              const fn = ctx.useImport("image");
              const assetVar = ctx.registerImageAsset(
                node.props["bytes"] as Uint8Array,
                "png",
              );
              return ctx.formatVanillaCall(fn, [
                ctx.formatProps([
                  ["id", node.props["id"]],
                  ["src", expr(assetVar)],
                ]),
              ]);
            },
            emitReact: (node, ctx) => {
              const comp = ctx.useImport("Image");
              const assetVar = ctx.registerImageAsset(
                node.props["bytes"] as Uint8Array,
                "png",
              );
              return ctx.formatJsxElement(comp, [
                ["id", node.props["id"]],
                ["src", expr(assetVar)],
              ]);
            },
          },
        },
        objectProps: {
          enter: (value, ctx) =>
            expr(`${ctx.useImport("fade")}(${ctx.emitValueExpr(value)})`),
        },
      }),
      bind() {
        throw new Error("not used in codegen unit test");
      },
    } as unknown as RegisteredCapabilityModule;

    const dummySlideCap = {
      declaration: "transition",
      slideProp: "transition",
      slot: "transition",
      codegen: defineSlideCodegen({
        slideProps: {
          transition: (value, ctx) =>
            expr(
              `${ctx.useImport("transitions")}.fade(${ctx.emitValueExpr(value)})`,
            ),
        },
      }),
      bindSlide() {
        throw new Error("not used in codegen unit test");
      },
    } as unknown as RegisteredCapabilityModule;

    const titleDef = defineSlideLayout({ id: "title", name: "Title" });
    const blankDef = defineSlideLayout({ id: "blank", name: "Blank" });
    const master1 = createSlideMasterNode({
      id: "m1",
      theme: sharedTheme,
    }).append(createSlideLayoutNode(titleDef), createSlideLayoutNode(blankDef));
    const master2 = createSlideMasterNode({
      id: "m2",
      theme: sharedTheme,
    }).append(createSlideLayoutNode({ id: "unused", name: "Unused" }));

    const t1 = new RichTextContainer("text", "text", {
      id: "t1",
      morphKey: scopeA.handle("hero"),
      enter: { duration: 300 },
    }).append(" world", "line1\nline2", "{brace}", "a & b");

    const t2 = new RichTextContainer("text", "text", {
      id: "t2",
      morphKey: scopeB.handle("hero"),
    }).append("Safe text");

    const img1 = new CommonNode("image", "image", { id: "i1", bytes: fakePng });
    const img2 = new CommonNode("image", "image", { id: "i2", bytes: fakePng });

    const doc = createDeckNode({
      width: 960,
      height: 540,
      theme: defaultTheme,
    }).append(
      master1,
      master2,
      createSectionNode({ name: "Main" }).append(
        createSlideNode({
          id: "s1",
          layout: titleDef,
          transition: { duration: "fast" },
        }).append(t1, img1),
        createSlideNode({ id: "s2" }).append(t2, img2),
      ),
    );

    const generate = createCodegen([dummyObjectCap, dummySlideCap]);
    const react = generate(doc, { target: "react" });
    const assetKeys = Object.keys(react.files).filter((key) =>
      key.startsWith("assets/"),
    );
    assert.equal(assetKeys.length, 1);

    const source = react.files["deck.tsx"] as string;
    // Section ordering: shared theme + morphScopes before masters/layouts, before slides, before GeneratedDeck
    const themeIdx = source.indexOf("defineTheme(");
    const morphIdx = source.indexOf("createMorphScope()");
    const layoutIdx = source.indexOf("defineSlideLayout(");
    const slideIdx = source.indexOf("function S1Slide()");
    const deckIdx = source.indexOf("export default function GeneratedDeck()");
    assert.ok(
      themeIdx >= 0 &&
        morphIdx > themeIdx &&
        layoutIdx > morphIdx &&
        slideIdx > layoutIdx &&
        deckIdx > slideIdx,
    );
    // Two distinct MorphScopes allocated for same-label distinct MorphKey handles
    assert.match(source, /const morphScope = createMorphScope\(\);/u);
    assert.match(source, /const morphScope2 = createMorphScope\(\);/u);
    // Unused layout from master2 is preserved
    assert.match(source, /name: "Unused"/u);
    // JSX whitespace/newline/entity protection
    assert.match(source, /\{" world"\}/u);
    assert.match(source, /\{"line1\\nline2"\}/u);
    assert.match(source, /\{"\{brace\}"\}/u);
    assert.match(source, /\{"a & b"\}/u);
    assert.match(source, />Safe text</u);
    // lighter() helper preserved
    assert.match(source, /lighter\("accent2", 0\.25\)/u);
  });

  it("escapes JSX string attributes containing backslashes or ampersands as expression containers", () => {
    const attrCap = {
      declaration: "shape",
      codegen: defineObjectCodegen({
        declarations: {
          shape: {
            emitVanilla: (node, ctx) =>
              ctx.formatVanillaCall(ctx.useImport("shape"), [
                ctx.formatProps(Object.entries(node.props)),
              ]),
            emitReact: (node, ctx) =>
              ctx.formatJsxElement(
                ctx.useImport("Shape"),
                Object.entries(node.props),
              ),
          },
        },
      }),
      bind() {
        throw new Error("not used in codegen unit test");
      },
    } as unknown as RegisteredCapabilityModule;

    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSectionNode({ name: "Path C:\\Temp &copy; 2026" }).append(
        createSlideNode({ id: "s-esc" }).append(
          new CommonNode("shape", "shape", {
            id: "sh-esc",
            x: 10,
            y: 10,
            width: 160,
            height: 50,
            text: "C:\\Users\\Alice &lt;tag&gt; &amp; Co.",
          }),
        ),
      ),
    );

    const generate = createCodegen([attrCap]);
    const react = generate(doc, { target: "react" });
    const source = react.files["deck.tsx"] as string;
    assert.ok(
      source.includes('name={"Path C:\\\\Temp &copy; 2026"}'),
      `Expected JSX expression attr for section name, got:\n${source}`,
    );
    assert.ok(
      source.includes('text={"C:\\\\Users\\\\Alice &lt;tag&gt; &amp; Co."}'),
      `Expected JSX expression attr for shape text, got:\n${source}`,
    );
  });

  it("wraps adjacent #text children in JSX expression containers while preserving clean single-child text", () => {
    const richCap = {
      declaration: "text",
      codegen: defineObjectCodegen({
        declarations: {
          text: {
            emitVanilla: (node, ctx) =>
              ctx.formatVanillaCall(
                ctx.useImport("text"),
                [ctx.formatProps(Object.entries(node.props))],
                ctx.emitRichTextChildren(node.children),
              ),
            emitReact: (node, ctx) =>
              ctx.formatJsxElement(
                ctx.useImport("Text"),
                Object.entries(node.props),
                ctx.emitRichTextChildren(node.children),
              ),
          },
          paragraph: {
            emitVanilla: (node, ctx) =>
              ctx.formatVanillaCall(
                ctx.useImport("paragraph"),
                [],
                ctx.emitRichTextChildren(node.children),
              ),
            emitReact: (node, ctx) =>
              ctx.formatJsxElement(
                ctx.useImport("P"),
                [],
                ctx.emitRichTextChildren(node.children),
              ),
          },
          span: {
            emitVanilla: (node, ctx) =>
              ctx.formatVanillaCall(
                ctx.useImport("span"),
                [],
                ctx.emitRichTextChildren(node.children),
              ),
            emitReact: (node, ctx) =>
              ctx.formatJsxElement(
                ctx.useImport("Span"),
                [],
                ctx.emitRichTextChildren(node.children),
              ),
          },
          box: {
            emitVanilla: (_node, ctx) =>
              ctx.formatVanillaCall(ctx.useImport("box"), []),
            emitReact: (_node, ctx) =>
              ctx.formatJsxElement(
                ctx.useImport("Box"),
                [],
                ["Hello", "World"],
              ),
          },
        },
      }),
      bind() {
        throw new Error("not used in codegen unit test");
      },
    } as unknown as RegisteredCapabilityModule;

    const adjText = new RichTextContainer("text", "text", {
      id: "adj-text",
    }).append("Hello", "World");
    const adjPara = new RichTextContainer("text", "text", {
      id: "adj-para",
    }).append(
      new RichTextContainer("paragraph", "paragraph", {}).append(
        "A",
        "B",
        new RichTextContainer("span", "span", {}).append("C"),
        "D",
        "E",
      ),
    );
    const singleText = new RichTextContainer("text", "text", {
      id: "single-text",
    }).append("Simple heading");
    const multiBareBox = new CommonNode("box", "box", { id: "raw-box" });

    const doc = createDeckNode({ width: 960, height: 540 }).append(
      createSlideNode({ id: "s1" }).append(
        adjText,
        adjPara,
        singleText,
        multiBareBox,
      ),
    );

    const generate = createCodegen([richCap]);
    const react = generate(doc, { target: "react" });
    const source = react.files["deck.tsx"] as string;
    assert.match(
      source,
      /<Text id="adj-text">\s*\{"Hello"\}\s*\{"World"\}\s*<\/Text>/u,
    );
    assert.match(
      source,
      /<P>\s*\{"A"\}\s*\{"B"\}\s*<Span>C<\/Span>\s*\{"D"\}\s*\{"E"\}\s*<\/P>/u,
    );
    assert.match(source, /<Text id="single-text">Simple heading<\/Text>/u);
    assert.match(source, /<Box>\s*\{"Hello"\}\s*\{"World"\}\s*<\/Box>/u);
  });
});
