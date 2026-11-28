import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  BuildError,
  CapabilityError,
  codegen,
  defineSlideLayout,
  defineTheme,
  emit,
  font,
  resolve,
} from "@officedk/pptxdk";
import {
  createResolver,
  emit as buildEmit,
  selectCapabilities,
  type DeclarationNode,
} from "@officedk/pptxdk/build";
import { paint } from "@officedk/pptxdk/preview";
import {
  Cell,
  Chart,
  Deck,
  P,
  Row,
  Shape,
  Slide,
  SlideLayout,
  SlideMaster,
  Span,
  Table,
  Text,
  expandDocument,
} from "@officedk/pptxdk/react";
import {
  cell,
  chart,
  deck,
  image,
  paragraph,
  row,
  shape,
  slide,
  slideLayout,
  slideMaster,
  span,
  table,
  text,
} from "@officedk/pptxdk/vanilla";
import { validatePptx } from "@pptxdk/validation";
import { unzipSync } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

const jostBookBytes = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const widePngBytes = new Uint8Array(
  readFileSync(new URL("./assets/wide-20000x1.png", import.meta.url)),
);
// Two valid 1x1 PNG buffers of identical byte length (67 bytes) for in-place mutation tests.
const png1x1BlackBytes = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgYGAAAAAEAAH2FzhVAAAAAElFTkSuQmCC",
    "base64",
  ),
);
const png1x1WhiteBytes = new Uint8Array(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4//8/AAX+Av4N70a4AAAAAElFTkSuQmCC",
    "base64",
  ),
);

describe("sealed font zero-copy sharing and digest memoization", () => {
  it("isolates sealed font() bytes from both input mutation and defensive .data reads while freezing structure", async () => {
    const inputBytes = new Uint8Array(jostBookBytes);
    const sealed = font("Jost*", [{ weight: 400, data: inputBytes }]);
    const noDataFont = font("Aptos", [{ weight: 400, style: "normal" }]);

    assert.equal(Object.isFrozen(sealed), true);
    assert.equal(Object.isFrozen(sealed.faces), true);
    assert.equal(Object.isFrozen(sealed.faces![0]), true);
    assert.equal("data" in noDataFont.faces![0]!, false);

    // Mutate caller input buffer after font() creation.
    inputBytes.fill(0xff, 0, 64);

    // Reading .data returns a fresh defensive copy; mutating it must not alter sealed font bytes.
    const readCopy = sealed.faces![0]!.data!;
    assert.notEqual(readCopy, sealed.faces![0]!.data!);
    readCopy.fill(0x00, 0, 64);
    assert.deepEqual(sealed.faces![0]!.data!, jostBookBytes);

    const buildDoc = () =>
      deck({ width: 640, height: 360 }).append(
        slide({ id: "s1" }).append(
          text(
            [
              paragraph([
                span("Sealed Jost Isolation", {
                  style: { fontFamily: sealed, fontSize: 20 },
                }),
              ]),
            ],
            {
              id: "t1",
              style: {
                position: "absolute",
                left: 32,
                top: 32,
                width: 320,
                height: 60,
              },
            },
          ),
        ),
      );

    const resolution = resolve(buildDoc());
    assert.equal(resolution.report.fonts[0]?.hasResource, true);
    assert.equal(resolution.report.fonts[0]?.embeddable, true);

    const unembedded = emit(resolution);
    const full = emit(resolution, { embedFonts: "full" });
    const subset = emit(resolution, { embedFonts: "subset" });
    assert.deepEqual((await validatePptx(unembedded.pptx)).errors, []);
    assert.deepEqual((await validatePptx(full.pptx)).errors, []);
    assert.deepEqual((await validatePptx(subset.pptx)).errors, []);

    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);
    assert.ok(painted.pages[0]?.png instanceof Uint8Array);

    const vanilla = codegen(buildDoc(), { target: "vanilla" });
    const react = codegen(buildDoc(), { target: "react" });
    const vanillaFontAsset = Object.keys(vanilla.files).find((k) =>
      k.startsWith("assets/font-"),
    );
    const reactFontAsset = Object.keys(react.files).find((k) =>
      k.startsWith("assets/font-"),
    );
    assert.ok(vanillaFontAsset !== undefined);
    assert.ok(reactFontAsset !== undefined);
    assert.deepEqual(vanilla.files[vanillaFontAsset], jostBookBytes);
    assert.deepEqual(react.files[reactFontAsset], jostBookBytes);
  });

  it("preserves exact sealed Font reference identity across theme, masters, slides, and canvas nodes while keeping Zod font-face and reserved-slot diagnostics", async () => {
    const sealed = font("Jost*", [{ weight: 400, data: jostBookBytes }]);
    const theme = defineTheme({
      name: "ZeroCopyTheme",
      fonts: { major: sealed, minor: sealed },
    });
    assert.equal(theme.fonts.major, sealed);
    assert.equal(theme.fonts.minor, sealed);

    const chartData = {
      kind: "shared-categories" as const,
      categories: ["Q1", "Q2"],
    };
    const chartSeries = [{ name: "Series", values: [10, 20] }];

    const masterNode = slideMaster({
      id: "m1",
      theme,
      style: { fontFamily: sealed },
    }).append(slideLayout({ id: "l1" }));
    const textNode = text(
      [paragraph([span("Hello", { style: { fontFamily: sealed } })])],
      {
        id: "txt",
        style: {
          position: "absolute",
          left: 20,
          top: 20,
          width: 200,
          height: 40,
          fontFamily: sealed,
          fontSize: 16,
        },
      },
    );
    const shapeNode = shape("Rect", {
      id: "shp",
      style: {
        position: "absolute",
        left: 20,
        top: 80,
        width: 200,
        height: 40,
        fontFamily: sealed,
        fontSize: 14,
      },
    });
    const tableNode = table({
      id: "tbl",
      columns: [200],
      style: {
        position: "absolute",
        left: 20,
        top: 140,
        width: 200,
        height: 40,
        fontFamily: sealed,
      },
    }).append(
      row().append(
        cell("Cell", {
          style: { fontFamily: sealed, fontSize: 12 },
        }),
      ),
    );
    const chartNode = chart({
      id: "cht",
      style: {
        position: "absolute",
        left: 260,
        top: 20,
        width: 320,
        height: 200,
        fontFamily: sealed,
      },
      groups: [{ kind: "column", data: chartData, series: chartSeries }],
    });
    const slideNode = slide({
      id: "s1",
      style: { fontFamily: sealed },
    }).append(textNode, shapeNode, tableNode, chartNode);
    const vanillaDoc = deck({ width: 640, height: 360, theme }).append(
      masterNode,
      slideNode,
    );

    assert.equal(
      (masterNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    assert.equal(
      (slideNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    assert.equal(
      (textNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    assert.equal(
      (shapeNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    assert.equal(
      (tableNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    assert.equal(
      (chartNode.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );

    const reactDoc = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360, theme },
        createElement(
          SlideMaster,
          { id: "m1", style: { fontFamily: sealed } },
          createElement(SlideLayout, { id: "l1" }),
        ),
        createElement(
          Slide,
          { id: "s1", style: { fontFamily: sealed } },
          createElement(
            Text,
            {
              id: "txt",
              style: {
                position: "absolute",
                left: 20,
                top: 20,
                width: 200,
                height: 40,
                fontFamily: sealed,
                fontSize: 16,
              },
            },
            createElement(
              P,
              null,
              createElement(Span, { style: { fontFamily: sealed } }, "Hello"),
            ),
          ),
          createElement(
            Shape,
            {
              id: "shp",
              style: {
                position: "absolute",
                left: 20,
                top: 80,
                width: 200,
                height: 40,
                fontFamily: sealed,
                fontSize: 14,
              },
            },
            "Rect",
          ),
          createElement(
            Table,
            {
              id: "tbl",
              columns: [200],
              style: {
                position: "absolute",
                left: 20,
                top: 140,
                width: 200,
                height: 40,
                fontFamily: sealed,
              },
            },
            createElement(
              Row,
              null,
              createElement(
                Cell,
                { style: { fontFamily: sealed, fontSize: 12 } },
                "Cell",
              ),
            ),
          ),
          createElement(Chart, {
            id: "cht",
            style: {
              position: "absolute",
              left: 260,
              top: 20,
              width: 320,
              height: 200,
              fontFamily: sealed,
            },
            groups: [{ kind: "column", data: chartData, series: chartSeries }],
          }),
        ),
      ),
    );
    const reactSlide = reactDoc.children[1]!;
    assert.ok("props" in reactSlide);
    assert.equal(
      (reactSlide.props["style"] as { fontFamily?: unknown }).fontFamily,
      sealed,
    );
    for (const child of reactSlide.children) {
      assert.ok("props" in child);
      assert.equal(
        (child.props["style"] as { fontFamily?: unknown }).fontFamily,
        sealed,
      );
    }

    assert.deepEqual(
      (await validatePptx(emit(resolve(vanillaDoc)).pptx)).errors,
      [],
    );
    assert.deepEqual(
      (await validatePptx(emit(resolve(reactDoc)).pptx)).errors,
      [],
    );

    // Invalid face weight in unsealed literal must still yield "invalid-font-face".
    const badFaceDoc = deck({ width: 640, height: 360 }).append(
      slide({ id: "bad" }).append(
        text("Bad", {
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 200,
            height: 40,
            fontFamily: { family: "Bad", faces: [{ weight: 0 }] } as never,
          },
        }),
      ),
    );
    assert.throws(
      () => resolve(badFaceDoc),
      (error: unknown) =>
        error instanceof BuildError &&
        error.diagnostics[0]?.code === "invalid-prop" &&
        error.diagnostics[0]?.message ===
          "Font.faces contains invalid weight, style, or resource",
    );

    // Sealed font with reserved slot name must still be rejected by defineTheme.
    for (const reserved of ["major", "minor", "+mj-lt", "+mn-ea"] as const) {
      const reservedFont = font(reserved);
      assert.throws(
        () => defineTheme({ fonts: { major: reservedFont } }),
        (error: unknown) =>
          error instanceof CapabilityError && error.code === "invalid-theme",
      );
    }
  });

  it("isolates unsealed { family, faces } literals and Image bytes from caller and node.props mutations while sharing assets across stages", async () => {
    const mutableFontBytes = new Uint8Array(jostBookBytes);
    const mutableImageBytes = new Uint8Array(widePngBytes);
    const literalFont = {
      family: "Jost*",
      faces: [
        { weight: 400, style: "normal" as const, data: mutableFontBytes },
      ],
    };

    const textNode = text("Unsealed literal font", {
      id: "literal-text",
      style: {
        position: "absolute",
        left: 20,
        top: 20,
        width: 300,
        height: 40,
        fontFamily: literalFont,
        fontSize: 16,
      },
    });
    const imageNode = image({
      id: "img1",
      src: mutableImageBytes,
      style: {
        position: "absolute",
        left: 20,
        top: 80,
        width: 200,
        height: 20,
      },
    });
    const imageNodeDuplicate = image({
      id: "img2",
      src: mutableImageBytes,
      style: {
        position: "absolute",
        left: 240,
        top: 80,
        width: 200,
        height: 20,
      },
    });

    // Mutate caller buffers immediately after node creation.
    mutableFontBytes[0]! ^= 0xff;
    mutableImageBytes[0]! ^= 0xff;

    // Public node.props must still return defensive copies for unsealed literals and images.
    const readFont1 = (
      textNode.props["style"] as {
        fontFamily: { faces: readonly { data: Uint8Array }[] };
      }
    ).fontFamily.faces[0]!.data;
    const readFont2 = (
      textNode.props["style"] as {
        fontFamily: { faces: readonly { data: Uint8Array }[] };
      }
    ).fontFamily.faces[0]!.data;
    assert.notEqual(readFont1, readFont2);
    readFont1[0]! ^= 0xff;
    assert.deepEqual(readFont2, jostBookBytes);

    const readImg1 = imageNode.props["src"] as Uint8Array;
    const readImg2 = imageNode.props["src"] as Uint8Array;
    assert.notEqual(readImg1, readImg2);
    readImg1[0]! ^= 0xff;
    assert.deepEqual(readImg2, widePngBytes);

    const doc = deck({ width: 640, height: 360 }).append(
      slide({ id: "s1" }).append(textNode, imageNode, imageNodeDuplicate),
    );
    const resolution = resolve(doc);
    assert.equal(resolution.report.fonts[0]?.hasResource, true);

    const emitted = emit(resolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    const mediaParts = Object.keys(unzipSync(emitted.pptx)).filter((name) =>
      name.startsWith("ppt/media/"),
    );
    assert.equal(mediaParts.length, 1);

    const painted = paint(resolution);
    assert.deepEqual(painted.pages[0]?.errors, []);

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(doc, { target });
      const imageAssets = Object.keys(generated.files).filter((key) =>
        key.startsWith("assets/image-"),
      );
      const fontAssets = Object.keys(generated.files).filter((key) =>
        key.startsWith("assets/font-"),
      );
      assert.equal(imageAssets.length, 1);
      assert.equal(fontAssets.length, 1);
      assert.deepEqual(generated.files[imageAssets[0]!], widePngBytes);
      assert.deepEqual(generated.files[fontAssets[0]!], jostBookBytes);
    }
  });

  it("reuses a single sealed Font across multiple masters, slides, text, tables, and charts cleanly through resolve, emit, paint, codegen, and validatePptx", async () => {
    const sharedFont = font("Jost*", [{ weight: 400, data: jostBookBytes }]);
    const sharedTheme = defineTheme({
      name: "HighFreqTheme",
      fonts: { major: sharedFont, minor: sharedFont },
    });

    const doc = deck({ width: 960, height: 540, theme: sharedTheme }).append(
      slideMaster({ id: "m1", theme: sharedTheme }).append(
        slideLayout({ id: "l1" }),
      ),
      ...Array.from({ length: 5 }, (_, slideIndex) =>
        slide({
          id: `s-${slideIndex + 1}`,
          style: { fontFamily: sharedFont },
        }).append(
          text(
            [
              paragraph([
                span(`Slide ${slideIndex + 1} title`, {
                  style: { fontFamily: sharedFont, fontSize: 18 },
                }),
              ]),
            ],
            {
              id: `t-${slideIndex + 1}`,
              style: {
                position: "absolute",
                left: 32,
                top: 24,
                width: 400,
                height: 40,
                fontFamily: sharedFont,
              },
            },
          ),
          table({
            id: `tbl-${slideIndex + 1}`,
            columns: [180, 180],
            style: {
              position: "absolute",
              left: 32,
              top: 80,
              width: 360,
              height: 80,
              fontFamily: sharedFont,
            },
          }).append(
            row().append(
              cell("Metric", {
                style: { fontFamily: sharedFont, fontSize: 12 },
              }),
              cell(`${slideIndex + 1}x`, {
                style: { fontFamily: sharedFont, fontSize: 12 },
              }),
            ),
          ),
          chart({
            id: `cht-${slideIndex + 1}`,
            style: {
              position: "absolute",
              left: 420,
              top: 80,
              width: 480,
              height: 260,
              fontFamily: sharedFont,
            },
            groups: [
              {
                kind: "column",
                data: { kind: "shared-categories", categories: ["A", "B"] },
                series: [{ name: "Series", values: [10 + slideIndex, 20] }],
              },
            ],
          }),
        ),
      ),
    );

    const resolution = resolve(doc);
    assert.equal(resolution.slides.length, 5);
    assert.equal(resolution.report.fonts.length, 1);
    assert.equal(resolution.report.fonts[0]?.hasResource, true);

    const unembedded = emit(resolution);
    const full = emit(resolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(unembedded.pptx)).errors, []);
    assert.deepEqual((await validatePptx(full.pptx)).errors, []);

    const painted = paint(resolution);
    assert.equal(painted.pages.length, 5);
    for (const page of painted.pages) {
      assert.deepEqual(page.errors, []);
    }

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(doc, { target });
      const fontAssets = Object.keys(generated.files).filter((key) =>
        key.startsWith("assets/font-"),
      );
      assert.equal(fontAssets.length, 1);
      assert.deepEqual(generated.files[fontAssets[0]!], jostBookBytes);
    }
  });

  it("snapshots defineSlideLayout inputs so mutating themeOverride font bytes and background image bytes cannot mutate the layout", async () => {
    const layoutFontBytes = new Uint8Array(jostBookBytes);
    const layoutBgBytes = new Uint8Array(widePngBytes);

    const layoutDef = defineSlideLayout({
      id: "isolated-layout",
      name: "Isolated Layout",
      themeOverride: {
        fonts: {
          major: {
            family: "Jost*",
            faces: [{ weight: 400, style: "normal", data: layoutFontBytes }],
          },
        },
      },
      background: { image: layoutBgBytes },
    });

    // Mutate caller buffers immediately after defineSlideLayout().
    layoutFontBytes.fill(0xff, 0, 64);
    layoutBgBytes.fill(0x00, 0, 16);

    const buildDoc = () =>
      deck({ width: 640, height: 360 }).append(
        slideMaster({ id: "m1" }).append(slideLayout(layoutDef)),
        slide({ id: "s1", layout: layoutDef }).append(
          text(
            [
              paragraph([
                span("Layout Override Heading", {
                  style: { fontFamily: "major", fontSize: 20 },
                }),
              ]),
            ],
            {
              id: "heading",
              style: {
                position: "absolute",
                left: 32,
                top: 32,
                width: 320,
                height: 48,
              },
            },
          ),
        ),
      );

    const firstDoc = buildDoc();
    const firstResolution = resolve(firstDoc);
    assert.equal(firstResolution.report.fonts[0]?.hasResource, true);
    const firstEmitted = emit(firstResolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(firstEmitted.pptx)).errors, []);
    const firstPaint = paint(firstResolution);
    assert.deepEqual(firstPaint.pages[0]?.errors, []);

    // Mutate caller buffers again after the first resolve/emit/paint run.
    layoutFontBytes.fill(0x11, 0, 128);
    layoutBgBytes.fill(0x22, 0, 32);

    const secondDoc = buildDoc();
    const secondResolution = resolve(secondDoc);
    const secondEmitted = emit(secondResolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(secondEmitted.pptx)).errors, []);
    const secondPaint = paint(secondResolution);
    assert.deepEqual(secondPaint.pages[0]?.errors, []);
    assert.deepEqual(secondPaint.pages[0]?.png, firstPaint.pages[0]?.png);

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(secondDoc, { target });
      const imageAssets = Object.keys(generated.files).filter((key) =>
        key.startsWith("assets/image-"),
      );
      const fontAssets = Object.keys(generated.files).filter((key) =>
        key.startsWith("assets/font-"),
      );
      assert.equal(imageAssets.length, 1);
      assert.equal(fontAssets.length, 1);
      assert.deepEqual(generated.files[imageAssets[0]!], widePngBytes);
      assert.deepEqual(generated.files[fontAssets[0]!], jostBookBytes);
    }
  });

  it("does not poison bytesSha256 when createResolver consumes plain DeclarationNodes with mutated Uint8Array buffers", async () => {
    assert.equal(png1x1BlackBytes.byteLength, png1x1WhiteBytes.byteLength);

    const mutableFontBuf = new Uint8Array(jostBookBytes.byteLength + 4);
    mutableFontBuf.set(jostBookBytes);
    mutableFontBuf[jostBookBytes.byteLength] = 0x01;
    const mutableImageBuf = new Uint8Array(png1x1BlackBytes);

    const plainDeck: DeclarationNode = {
      type: "deck",
      props: { width: 640, height: 360 },
      children: [
        {
          type: "slide",
          props: { id: "s1" },
          children: [
            {
              type: "text",
              props: {
                id: "t1",
                style: {
                  position: "absolute",
                  left: 20,
                  top: 20,
                  width: 320,
                  height: 60,
                  fontFamily: {
                    family: "Jost*",
                    faces: [{ data: mutableFontBuf }],
                  },
                  fontSize: 18,
                },
              },
              children: [{ type: "#text", text: "Plain Declaration Node" }],
            },
            {
              type: "image",
              props: {
                id: "img1",
                src: mutableImageBuf,
                style: {
                  position: "absolute",
                  left: 20,
                  top: 100,
                  width: 40,
                  height: 40,
                },
              },
              children: [],
            },
          ],
        },
      ],
    };

    const modules = await selectCapabilities(["text", "image"]);
    const customResolve = createResolver(modules);

    const firstResolution = customResolve(plainDeck);
    assert.equal(firstResolution.report.fonts[0]?.weight, 400);
    assert.equal(firstResolution.report.fonts[0]?.hasResource, true);
    const firstEmitted = buildEmit(firstResolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(firstEmitted.pptx)).errors, []);
    const firstZip = unzipSync(firstEmitted.pptx);
    assert.deepEqual(firstZip["ppt/media/resource1.png"], png1x1BlackBytes);
    const firstFontPartName = Object.keys(firstZip).find((k) =>
      k.startsWith("ppt/fonts/"),
    );
    assert.ok(firstFontPartName !== undefined);

    // Mutate both buffers in-place on the same plain DeclarationNode and re-run resolve + emit.
    mutableFontBuf[jostBookBytes.byteLength] = 0x02;
    mutableImageBuf.set(png1x1WhiteBytes);

    const secondResolution = customResolve(plainDeck);
    assert.equal(secondResolution.report.fonts[0]?.weight, 400);
    assert.equal(secondResolution.report.fonts[0]?.hasResource, true);
    const secondEmitted = buildEmit(secondResolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(secondEmitted.pptx)).errors, []);
    const secondZip = unzipSync(secondEmitted.pptx);
    assert.deepEqual(secondZip["ppt/media/resource1.png"], png1x1WhiteBytes);
    const secondFontPartName = Object.keys(secondZip).find((k) =>
      k.startsWith("ppt/fonts/"),
    );
    assert.ok(secondFontPartName !== undefined);
    assert.notDeepEqual(
      secondZip[secondFontPartName],
      firstZip[firstFontPartName],
    );
  });

  it("isolates inline deck theme unsealed font literals and image bytes across resolve, emit, codegen, and node.props mutations", async () => {
    const themeFontBytes = new Uint8Array(jostBookBytes);
    const imgBytes = new Uint8Array(png1x1BlackBytes);
    const bgBytes = new Uint8Array(png1x1WhiteBytes);

    const imageNode = image({
      id: "hero-img",
      src: imgBytes,
      style: {
        position: "absolute",
        left: 20,
        top: 100,
        width: 80,
        height: 80,
      },
    });
    const slideNode = slide({
      id: "s1",
      background: { image: bgBytes },
    }).append(
      text(
        [
          paragraph([
            span("Inline Theme Font", {
              style: { fontFamily: "major", fontSize: 20 },
            }),
          ]),
        ],
        {
          id: "title",
          style: {
            position: "absolute",
            left: 20,
            top: 20,
            width: 320,
            height: 48,
          },
        },
      ),
      imageNode,
    );
    const doc = deck({
      width: 640,
      height: 360,
      theme: {
        fonts: {
          major: {
            family: "Jost*",
            faces: [{ weight: 400, style: "normal", data: themeFontBytes }],
          },
        },
      },
    }).append(slideNode);

    // Mutate caller buffers after node creation.
    themeFontBytes.fill(0xff, 0, 64);
    imgBytes.fill(0x00, 0, 16);
    bgBytes.fill(0x00, 0, 16);

    // Mutate defensive copies returned by public CommonNode.props getters.
    const propsTheme = doc.props["theme"] as {
      fonts: { major: { faces: readonly { data: Uint8Array }[] } };
    };
    const readThemeFont = propsTheme.fonts.major.faces[0]!.data;
    readThemeFont.fill(0xee, 0, 64);
    const readImg = imageNode.props["src"] as Uint8Array;
    readImg.fill(0xee, 0, 16);
    const readBg = (slideNode.props["background"] as { image: Uint8Array })
      .image;
    readBg.fill(0xee, 0, 16);

    const resolution = resolve(doc);
    assert.equal(resolution.report.fonts[0]?.hasResource, true);
    const emitted = emit(resolution, { embedFonts: "full" });
    assert.deepEqual((await validatePptx(emitted.pptx)).errors, []);
    const zip = unzipSync(emitted.pptx);
    const mediaEntries = Object.keys(zip)
      .filter((k) => k.startsWith("ppt/media/"))
      .map((k) => zip[k]);
    assert.equal(mediaEntries.length, 2);

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(doc, { target });
      const fontAssets = Object.keys(generated.files).filter((k) =>
        k.startsWith("assets/font-"),
      );
      const imageAssets = Object.keys(generated.files).filter((k) =>
        k.startsWith("assets/image-"),
      );
      assert.equal(fontAssets.length, 1);
      assert.equal(imageAssets.length, 2);
      assert.deepEqual(generated.files[fontAssets[0]!], jostBookBytes);
      assert.ok(
        imageAssets.some((k) =>
          Buffer.from(generated.files[k] as Uint8Array).equals(
            png1x1BlackBytes,
          ),
        ),
      );
      assert.ok(
        imageAssets.some((k) =>
          Buffer.from(generated.files[k] as Uint8Array).equals(
            png1x1WhiteBytes,
          ),
        ),
      );
    }
  });
});
