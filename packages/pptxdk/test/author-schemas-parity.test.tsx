import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import * as coreMain from "@pptxdk/core";
import * as coreCapability from "@pptxdk/core/capability";
import * as coreDocument from "@pptxdk/core/document";
import * as coreTheme from "@pptxdk/core/theme";
import { validatePptx } from "@pptxdk/validation";
import {
  BuildError,
  authorIdSchema,
  build,
  createMorphScope,
  deckPropsSchema,
  defineTheme,
  font,
  layoutLengthSchema,
  layoutSpacingSchema,
  layoutStyleSchema,
  linkTargetSchema,
  paddingStyleSchema,
  resolve,
  sectionPropsSchema,
  slideBackgroundSchema,
  slidePropsSchema,
  textStyleSchema,
  themeInputSchema,
  themeSchema,
  themeScopeSchema,
} from "pptxdk";
import {
  Box,
  Deck,
  MotionScope,
  P,
  Section,
  Slide,
  Span,
  Text,
  expandDocument,
  useMotionScope,
} from "pptxdk/react";
import {
  box,
  createMotionScopeHandle,
  deck,
  motionScope,
  paragraph,
  section,
  slide,
  span,
  text,
} from "pptxdk/vanilla";
import type { ReactNode } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);

const jostFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);

function ScopedMotion({
  children,
  onHandle,
}: {
  readonly children?: ReactNode;
  readonly onHandle?: (handle: ReturnType<typeof useMotionScope>) => void;
}) {
  const scope = useMotionScope();
  onHandle?.(scope);
  return <MotionScope scope={scope}>{children}</MotionScope>;
}

describe("Ticket 16: core author Zod schemas & dual-entry parity", () => {
  it("exposes Zod schemas across @pptxdk/core public subpaths (., ./capability, ./document, ./theme)", () => {
    for (const key of [
      "authorIdSchema",
      "deckPropsSchema",
      "slidePropsSchema",
      "sectionPropsSchema",
      "layoutLengthSchema",
      "layoutSpacingSchema",
      "layoutStyleSchema",
      "paddingStyleSchema",
      "textStyleSchema",
      "themeInputSchema",
      "themeSchema",
      "themeScopeSchema",
      "linkTargetSchema",
    ] as const) {
      assert.ok(key in coreMain, `@pptxdk/core must export ${key}`);
      assert.ok(
        key in coreCapability,
        `@pptxdk/core/capability must export ${key}`,
      );
      assert.ok(
        key in coreDocument,
        `@pptxdk/core/document must export ${key}`,
      );
    }
    for (const key of [
      "textStyleSchema",
      "themeInputSchema",
      "themeSchema",
      "themeScopeSchema",
      "colorValueSchema",
      "fontSchema",
    ] as const) {
      assert.ok(key in coreTheme, `@pptxdk/core/theme must export ${key}`);
    }

    assert.equal(authorIdSchema.parse("hero"), "hero");
    assert.equal(layoutLengthSchema.parse("50%"), "50%");
    assert.deepEqual(layoutSpacingSchema.parse([8, 16]), [8, 16]);
    assert.deepEqual(paddingStyleSchema.parse({ paddingTop: 12 }), {
      paddingTop: 12,
    });
    assert.equal(layoutStyleSchema.parse({ display: "grid" }).display, "grid");
    assert.equal(textStyleSchema.parse({ fontSize: "32pt" }).fontSize, "32pt");
    assert.equal(deckPropsSchema.parse({ width: 960, height: 540 }).width, 960);
    assert.equal(slidePropsSchema.parse({ id: "s1" }).id, "s1");
    assert.equal(sectionPropsSchema.parse({ name: "Sec" }).name, "Sec");
    assert.deepEqual(linkTargetSchema.parse({ action: "home" }), {
      action: "home",
    });
    assert.equal(slideBackgroundSchema.parse("#112233"), "#112233");
    assert.equal(
      themeInputSchema.parse({ colors: { ink: "dk1" } }).colors?.["ink"],
      "dk1",
    );
    assert.equal(themeSchema.parse(defineTheme()).name, "Office Theme");
    assert.equal(
      themeScopeSchema.parse({ colors: { ink: "#112233" } }).colors?.["ink"],
      "#112233",
    );
  });

  it("produces identical valid documents and PPTX outputs from pptxdk/react and pptxdk/vanilla", async () => {
    const theme = defineTheme({
      name: "Shared Brand",
      colors: { ink: "#1E293B", accent: "accent1" },
      fonts: { major: jostFont, minor: jostFont },
    });

    const jsxDoc = await expandDocument(
      <Deck
        width={960}
        height={540}
        theme={theme}
        title="Dual Entry Parity"
        style={{ fontFamily: "minor", color: "ink", fontSize: 18 }}
      >
        <Section name="Overview">
          <Slide
            id="cover"
            theme={{ colors: { ink: "#0F172A" } }}
            style={{
              display: "flex",
              flexDirection: "column",
              padding: [24, 32],
              gap: 16,
            }}
            background="#F8FAFC"
          >
            <Box
              id="grid-panel"
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 12,
                padding: 12,
              }}
            >
              <Text
                id="pt-title"
                link={{ id: "details" }}
                style={{ fontSize: "32pt", fontWeight: 400, lang: "en-US" }}
              >
                <P style={{ align: "left", spacingAfter: 6 }}>
                  <Span style={{ textDecoration: "underline" }}>
                    32pt Title
                  </Span>
                </P>
              </Text>
              <Text id="px-title" style={{ fontSize: "32px", color: "accent" }}>
                32px Title
              </Text>
            </Box>
          </Slide>
          <Slide id="details">
            <Text
              id="back"
              link={{ action: "home" }}
              x={40}
              y={40}
              width={200}
              height={40}
            >
              Home
            </Text>
          </Slide>
        </Section>
      </Deck>,
    );

    const vanillaDoc = deck({
      width: 960,
      height: 540,
      theme,
      title: "Dual Entry Parity",
      style: { fontFamily: "minor", color: "ink", fontSize: 18 },
    }).append(
      section({ name: "Overview" }).append(
        slide({
          id: "cover",
          theme: { colors: { ink: "#0F172A" } },
          style: {
            display: "flex",
            flexDirection: "column",
            padding: [24, 32],
            gap: 16,
          },
          background: "#F8FAFC",
        }).append(
          box({
            id: "grid-panel",
            style: {
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gap: 12,
              padding: 12,
            },
          }).append(
            text(
              paragraph(
                [
                  span("32pt Title", {
                    style: { textDecoration: "underline" },
                  }),
                ],
                {
                  style: { align: "left", spacingAfter: 6 },
                },
              ),
              {
                id: "pt-title",
                link: { id: "details" },
                style: { fontSize: "32pt", fontWeight: 400, lang: "en-US" },
              },
            ),
            text("32px Title", {
              id: "px-title",
              style: { fontSize: "32px", color: "accent" },
            }),
          ),
        ),
        slide({ id: "details" }).append(
          text("Home", {
            id: "back",
            link: { action: "home" },
            x: 40,
            y: 40,
            width: 200,
            height: 40,
          }),
        ),
      ),
    );

    const jsxBuilt = await build(jsxDoc);
    const vanillaBuilt = await build(vanillaDoc);

    assert.deepEqual(jsxBuilt.report, vanillaBuilt.report);
    const jsxValidation = await validatePptx(jsxBuilt.pptx);
    const vanillaValidation = await validatePptx(vanillaBuilt.pptx);
    assert.deepEqual(jsxValidation.errors, []);
    assert.deepEqual(vanillaValidation.errors, []);
    assert.deepEqual(jsxValidation.slides, vanillaValidation.slides);
  });

  it("rejects invalid inputs identically across pptxdk/react and pptxdk/vanilla with localized diagnostics", async () => {
    const cases: readonly {
      readonly label: string;
      readonly expectedCode: string;
      readonly expectedPath: string;
      readonly jsx: () => ReactNode;
      readonly vanilla: () => ReturnType<typeof deck>;
    }[] = [
      {
        label: "invalid Deck dimension",
        expectedCode: "invalid-dimension",
        expectedPath: "deck",
        jsx: () => (
          <Deck width={10}>
            <Slide id="s1" />
          </Deck>
        ),
        vanilla: () => deck({ width: 10 }).append(slide({ id: "s1" })),
      },
      {
        label: "unknown Deck parameter",
        expectedCode: "unknown-parameter",
        expectedPath: "deck",
        jsx: () => (
          <Deck {...({ unknownProp: 1 } as Record<string, unknown>)}>
            <Slide id="s1" />
          </Deck>
        ),
        vanilla: () =>
          deck({ unknownProp: 1 } as never).append(slide({ id: "s1" })),
      },
      {
        label: "invalid Section name",
        expectedCode: "invalid-section-name",
        expectedPath: "deck/section[1]",
        jsx: () => (
          <Deck>
            <Section name="  ">
              <Slide id="s1" />
            </Section>
          </Deck>
        ),
        vanilla: () =>
          deck().append(section({ name: "  " }).append(slide({ id: "s1" }))),
      },
      {
        label: "invalid Slide author id",
        expectedCode: "invalid-author-id",
        expectedPath: "deck/slide[1]",
        jsx: () => (
          <Deck>
            <Slide id="   " />
          </Deck>
        ),
        vanilla: () => deck().append(slide({ id: "   " })),
      },
      {
        label: "Slide zIndex rejected",
        expectedCode: "unsupported-layout-style",
        expectedPath: "deck/slide[s1]",
        jsx: () => (
          <Deck>
            <Slide id="s1" style={{ zIndex: 1 } as never} />
          </Deck>
        ),
        vanilla: () =>
          deck().append(slide({ id: "s1", style: { zIndex: 1 } as never })),
      },
      {
        label: "invalid Text fontSize",
        expectedCode: "invalid-font-size",
        expectedPath: "deck/slide[s1]/text[t1]",
        jsx: () => (
          <Deck>
            <Slide id="s1">
              <Text id="t1" style={{ fontSize: -8 }}>
                Bad
              </Text>
            </Slide>
          </Deck>
        ),
        vanilla: () =>
          deck().append(
            slide({ id: "s1" }).append(
              text("Bad", { id: "t1", style: { fontSize: -8 } }),
            ),
          ),
      },
      {
        label: "invalid layout enum value",
        expectedCode: "unsupported-layout-value",
        expectedPath: "deck/slide[s1]/box[b1]",
        jsx: () => (
          <Deck>
            <Slide id="s1">
              <Box id="b1" style={{ display: "inline" as never }} />
            </Slide>
          </Deck>
        ),
        vanilla: () =>
          deck().append(
            slide({ id: "s1" }).append(
              box({ id: "b1", style: { display: "inline" as never } }),
            ),
          ),
      },
      {
        label: "invalid local theme scope",
        expectedCode: "invalid-theme-scope",
        expectedPath: "deck/slide[s1]/box[b1]",
        jsx: () => (
          <Deck>
            <Slide id="s1">
              <Box id="b1" theme={{ colors: { dk1: "#112233" } }} />
            </Slide>
          </Deck>
        ),
        vanilla: () =>
          deck().append(
            slide({ id: "s1" }).append(
              box({ id: "b1", theme: { colors: { dk1: "#112233" } } }),
            ),
          ),
      },
      {
        label: "invalid LinkTarget",
        expectedCode: "invalid-link",
        expectedPath: "deck/slide[s1]/text[t1]",
        jsx: () => (
          <Deck>
            <Slide id="s1">
              <Text id="t1" link={{ id: "s1", url: "https://example.test" }}>
                Bad link
              </Text>
            </Slide>
          </Deck>
        ),
        vanilla: () =>
          deck().append(
            slide({ id: "s1" }).append(
              text("Bad link", {
                id: "t1",
                link: { id: "s1", url: "https://example.test" } as never,
              }),
            ),
          ),
      },
    ];

    for (const item of cases) {
      const jsxDoc = await expandDocument(item.jsx());
      const vanillaDoc = item.vanilla();

      let jsxError: unknown;
      try {
        resolve(jsxDoc);
      } catch (error) {
        jsxError = error;
      }

      let vanillaError: unknown;
      try {
        resolve(vanillaDoc);
      } catch (error) {
        vanillaError = error;
      }

      assert(
        jsxError instanceof BuildError,
        `${item.label} JSX should throw BuildError`,
      );
      assert(
        vanillaError instanceof BuildError,
        `${item.label} vanilla should throw BuildError`,
      );
      assert.equal(
        jsxError.diagnostics[0]?.code,
        item.expectedCode,
        `${item.label} JSX code`,
      );
      assert.equal(
        vanillaError.diagnostics[0]?.code,
        item.expectedCode,
        `${item.label} vanilla code`,
      );
      assert.equal(
        jsxError.diagnostics[0]?.path,
        item.expectedPath,
        `${item.label} JSX path`,
      );
      assert.equal(
        vanillaError.diagnostics[0]?.path,
        item.expectedPath,
        `${item.label} vanilla path`,
      );
    }
  });

  it("preserves 32pt vs 32 vs 32px, omitted vs explicit equal defaults, patch undefined semantics, snapshotting, and identity", async () => {
    const morphScope = createMorphScope();
    const heroKey = morphScope.handle("hero");
    const motionHandle = createMotionScopeHandle();
    const bgUrl = new URL("https://example.test/bg.png");
    const mutableStyle = { fontSize: "32pt" as const, color: "#112233" };

    const vanillaDoc = deck({}).append(
      slide({ id: "s1", background: { image: bgUrl } }).append(
        motionScope(motionHandle).append(
          text("Pt", { id: "t-pt", morphKey: heroKey, style: mutableStyle }),
          text("Num", { id: "t-num", style: { fontSize: 32 } }),
          text("Px", { id: "t-px", style: { fontSize: "32px" } }),
          text("Default", { id: "t-default" }),
          text("ExplicitEqual", { id: "t-explicit", style: { fontSize: 24 } }),
        ),
      ),
    );
    mutableStyle.color = "#999999";

    let jsxMotionHandle: ReturnType<typeof useMotionScope> | undefined;
    const jsxDoc = await expandDocument(
      <Deck>
        <Slide id="s1" background={{ image: bgUrl }}>
          <ScopedMotion
            onHandle={(handle) => {
              jsxMotionHandle = handle;
            }}
          >
            <Text
              id="t-pt"
              morphKey={heroKey}
              style={{ fontSize: "32pt", color: "#112233" }}
            >
              Pt
            </Text>
            <Text id="t-num" style={{ fontSize: 32 }}>
              Num
            </Text>
            <Text id="t-px" style={{ fontSize: "32px" }}>
              Px
            </Text>
            <Text id="t-default">Default</Text>
            <Text id="t-explicit" style={{ fontSize: 24 }}>
              ExplicitEqual
            </Text>
          </ScopedMotion>
        </Slide>
      </Deck>,
    );

    for (const [label, root] of [
      ["vanilla", vanillaDoc],
      ["jsx", jsxDoc],
    ] as const) {
      assert.equal(
        Object.hasOwn(root.props, "width"),
        false,
        `${label} omitted deck width`,
      );
      const s1 = root.query.get("s1");
      assert.equal(
        ((s1.node.props["background"] as { image?: unknown }).image as URL)
          .href,
        "https://example.test/bg.png",
      );

      const tPt = s1.getNode("t-pt");
      const tNum = s1.getNode("t-num");
      const tPx = s1.getNode("t-px");
      const tDefault = s1.getNode("t-default");
      const tExplicit = s1.getNode("t-explicit");

      assert.equal(
        (tPt.props["style"] as { fontSize?: unknown }).fontSize,
        "32pt",
      );
      assert.equal(
        (tPt.props["style"] as { color?: unknown }).color,
        "#112233",
      );
      assert.equal(
        (tNum.props["style"] as { fontSize?: unknown }).fontSize,
        32,
      );
      assert.equal(
        (tPx.props["style"] as { fontSize?: unknown }).fontSize,
        "32px",
      );
      assert.equal(Object.hasOwn(tDefault.props, "style"), false);
      assert.equal(
        (tExplicit.props["style"] as { fontSize?: unknown }).fontSize,
        24,
      );
      assert.equal(tPt.props["morphKey"], heroKey);

      // .style(patch) omission preserves fontSize, explicit undefined deletes only fontSize
      tPt.style({ fontWeight: "bold" });
      assert.equal(
        (tPt.props["style"] as { fontSize?: unknown }).fontSize,
        "32pt",
      );
      assert.equal(
        (tPt.props["style"] as { color?: unknown }).color,
        "#112233",
      );
      assert.equal(
        (tPt.props["style"] as { fontWeight?: unknown }).fontWeight,
        "bold",
      );

      tPt.style({ fontSize: undefined });
      const afterClear = tPt.props["style"] as Record<string, unknown>;
      assert.equal(Object.hasOwn(afterClear, "fontSize"), false);
      assert.equal(afterClear["color"], "#112233");
      assert.equal(afterClear["fontWeight"], "bold");
    }

    const vanillaSlide = vanillaDoc.children[0];
    assert(vanillaSlide !== undefined && "children" in vanillaSlide);
    const vanillaMotion = vanillaSlide.children[0];
    assert(vanillaMotion !== undefined && "props" in vanillaMotion);
    assert.equal(vanillaMotion.props["scope"], motionHandle);

    const jsxSlide = jsxDoc.children[0];
    assert(jsxSlide !== undefined && "children" in jsxSlide);
    const jsxMotion = jsxSlide.children[0];
    assert(jsxMotion !== undefined && "props" in jsxMotion);
    assert.equal(jsxMotion.props["scope"], jsxMotionHandle);
  });
});
