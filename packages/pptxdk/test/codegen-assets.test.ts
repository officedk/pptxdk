import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  mkdir,
  mkdtemp,
  readFile,
  rename,
  rm,
  symlink,
  writeFile,
} from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { unpackEotFontData } from "@pptxdk/import";
import { validatePptx } from "@pptxdk/validation";
import { strFromU8, unzipSync } from "fflate";
import * as fontkit from "fontkit";
import { PNG } from "pngjs";
import {
  build,
  codegen,
  defineTheme,
  importPptx,
  type BuildOptions,
  type CodegenResult,
} from "pptxdk";
import { buildFromSource, previewFromSource } from "pptxdk/node";
import {
  Deck,
  Image,
  P,
  Slide,
  Span,
  Text,
  darker,
  defineSlideLayout,
  expandDocument,
  font,
  lighter,
  linearGradient,
  radialGradient,
  SlideLayout,
  SlideMaster,
} from "pptxdk/react";
import {
  deck,
  defineSlideLayout as vanillaDefineSlideLayout,
  font as vanillaFont,
  image,
  linearGradient as vanillaLinearGradient,
  paragraph,
  preset,
  radialGradient as vanillaRadialGradient,
  slide,
  slideLayout,
  slideMaster,
  span,
  text,
} from "pptxdk/vanilla";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
  writeCodegenFiles,
} from "./helpers/codegen-acceptance.ts";

const require = createRequire(import.meta.url);
const { encode: encodeJpeg } = require("jpeg-js") as typeof import("jpeg-js");
const packageRoot = fileURLToPath(new URL("../", import.meta.url));

const jostRegularSource = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Book.ttf", import.meta.url)),
);
const jostBoldSource = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-700-Bold.ttf", import.meta.url)),
);
const jostItalicSource = new Uint8Array(
  readFileSync(new URL("./assets/fonts/Jost-400-Italic.ttf", import.meta.url)),
);

function createTwoTonePng(): Uint8Array {
  // Left half (x < 30) red, right half (x >= 30) blue.
  const img = new PNG({ width: 60, height: 30 });
  for (let y = 0; y < img.height; y += 1) {
    for (let x = 0; x < img.width; x += 1) {
      img.data.set(
        x < 30 ? [220, 40, 40, 255] : [30, 100, 220, 255],
        (y * img.width + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(img));
}

function createSolidJpeg(r: number, g: number, b: number): Uint8Array {
  const img = new PNG({ width: 16, height: 16 });
  for (let i = 0; i < img.data.length; i += 4) {
    img.data[i] = r;
    img.data[i + 1] = g;
    img.data[i + 2] = b;
    img.data[i + 3] = 255;
  }
  return new Uint8Array(encodeJpeg(img, 100).data);
}

// 1x1 teal (#0D9488) GIF89a.
const sampleGifBytes = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0x80, 0, 0, 0x0d, 0x94, 0x88,
  0x00, 0x00, 0x00, 0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0x00, 0x02, 0x02, 0x44, 0x01,
  0x00, 0x3b,
]);

function readPngPixel(
  pngBytes: Uint8Array,
  x: number,
  y: number,
): readonly [number, number, number, number] {
  const decoded = PNG.sync.read(Buffer.from(pngBytes));
  const offset = (y * decoded.width + x) * 4;
  return [
    decoded.data[offset] ?? 0,
    decoded.data[offset + 1] ?? 0,
    decoded.data[offset + 2] ?? 0,
    decoded.data[offset + 3] ?? 0,
  ];
}

function corruptPngIdatStream(validPng: Uint8Array): Uint8Array {
  const copy = new Uint8Array(validPng);
  const idatOffset = Buffer.from(copy).indexOf(Buffer.from("IDAT"));
  assert.ok(idatOffset > 0);
  copy[idatOffset + 6] = (copy[idatOffset + 6] ?? 0) ^ 0xff;
  copy[idatOffset + 7] = (copy[idatOffset + 7] ?? 0) ^ 0xff;
  return copy;
}

async function withRelocatedConsumerProject<T>(
  result: CodegenResult,
  fn: (ctx: {
    readonly relocatedDir: string;
    readonly entryPath: string;
  }) => Promise<T>,
): Promise<T> {
  const tempRoot = await mkdtemp(join(tmpdir(), "pptxdk-assets-relocated-"));
  try {
    const consumerRoot = join(tempRoot, "consumer");
    const nodeModulesDir = join(consumerRoot, "node_modules");
    await mkdir(nodeModulesDir, { recursive: true });
    await symlink(packageRoot, join(nodeModulesDir, "pptxdk"), "dir");
    await symlink(
      join(packageRoot, "node_modules/react"),
      join(nodeModulesDir, "react"),
      "dir",
    );

    const originDir = join(tempRoot, "origin");
    await writeCodegenFiles(originDir, result.files);

    const relocatedParent = join(consumerRoot, "deep", "nested", "workspace");
    await mkdir(relocatedParent, { recursive: true });
    const relocatedDir = join(relocatedParent, "project");
    await rename(originDir, relocatedDir);

    const entryPath = join(relocatedDir, result.entry);
    return await fn({ relocatedDir, entryPath });
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

function inspectEmbeddedFontsInPptx(pptx: Uint8Array): {
  readonly typefaces: readonly string[];
  readonly unpackedFaces: readonly {
    readonly familyName: string | null;
    readonly preferredFamily: string | null;
    readonly weightClass: number | undefined;
    readonly italic: boolean;
  }[];
} {
  const parts = unzipSync(pptx);
  const presXml = strFromU8(parts["ppt/presentation.xml"]!);
  const typefaces = [
    ...presXml.matchAll(/<p:font\b[^>]*typeface="([^"]+)"/gu),
  ].map((m) => m[1]!);

  const fontParts = Object.keys(parts)
    .filter(
      (name) => name.startsWith("ppt/fonts/") && name.endsWith(".fntdata"),
    )
    .toSorted();

  const unpackedFaces = fontParts.map((partPath) => {
    const unpacked = unpackEotFontData(parts[partPath]!);
    assert.equal(unpacked.status, "ok", `${partPath} must unpack cleanly`);
    const opened = fontkit.create(Buffer.from(unpacked.ttfBytes));
    assert.equal(opened.type, "TTF");
    return {
      familyName: opened.familyName,
      preferredFamily: opened.getName("preferredFamily", "en"),
      weightClass: opened["OS/2"]?.usWeightClass,
      italic:
        Boolean(opened["OS/2"]?.fsSelection?.italic) ||
        Boolean(opened["OS/2"]?.fsSelection?.oblique) ||
        opened.italicAngle !== 0,
    };
  });

  return { typefaces, unpackedFaces };
}

async function assertAssetDeckRebuild(
  pptx: Uint8Array,
  label: string,
  expectedHeroPng: Uint8Array,
  expectedBgJpeg: Uint8Array,
  expectedWatermarkGif: Uint8Array,
  expectedExternalUrl: string,
): Promise<void> {
  const validation = await validatePptx(pptx);
  assert.deepEqual(
    validation.errors,
    [],
    `${label} validation errors: ${JSON.stringify(validation.errors)}`,
  );
  assert.equal(validation.slides.length, 4, `${label} slide count`);

  const parts = unzipSync(pptx);

  // Slide 1: cropped PNG + external image URL + rotated/styled GIF + embedded font text
  const s1Objects = validation.slides[0]!.objects;
  const heroObj = s1Objects.find((obj) => obj.image !== undefined && obj.crop);
  assert.ok(heroObj !== undefined, `${label} cropped hero image`);
  assert.ok(heroObj.image?.part !== undefined, `${label} hero media part`);
  assert.deepEqual(
    heroObj.crop,
    { left: 50_000 },
    `${label} preserved hero crop (50000 = 50% in DrawingML)`,
  );
  assert.deepEqual(
    parts[heroObj.image.part],
    expectedHeroPng,
    `${label} hero PNG media bytes match`,
  );

  const extObj = s1Objects.find(
    (obj) => obj.image?.externalTarget === expectedExternalUrl,
  );
  assert.ok(extObj !== undefined, `${label} external image object`);
  assert.equal(extObj.image?.part, undefined, `${label} external has no part`);
  const s1RelsXml = strFromU8(parts["ppt/slides/_rels/slide1.xml.rels"]!);
  assert.ok(
    s1RelsXml.includes(`Target="${expectedExternalUrl}"`) &&
      s1RelsXml.includes('TargetMode="External"'),
    `${label} slide1 rels preserves TargetMode="External" and target URL`,
  );

  const gifObj = s1Objects.find(
    (obj) => obj.image?.contentType === "image/gif",
  );
  assert.ok(gifObj?.image?.part !== undefined, `${label} GIF image part`);
  assert.equal(gifObj.rotationDegrees, 15, `${label} GIF rotation`);
  assert.deepEqual(
    parts[gifObj.image.part],
    expectedWatermarkGif,
    `${label} GIF media bytes match`,
  );

  // Slide 2: JPEG image background + shared PNG image object
  const s2Bg = validation.slides[1]!.background;
  assert.equal(s2Bg?.type, "image", `${label} slide 2 image background`);
  if (s2Bg?.type === "image") {
    assert.equal(s2Bg.contentType, "image/jpeg");
    assert.deepEqual(
      parts[s2Bg.part],
      expectedBgJpeg,
      `${label} JPEG background bytes match`,
    );
  }
  const s2SharedPng = validation.slides[1]!.objects.find(
    (obj) => obj.image?.contentType === "image/png",
  );
  assert.ok(
    s2SharedPng?.image?.part !== undefined,
    `${label} shared PNG on slide 2`,
  );
  assert.deepEqual(
    parts[s2SharedPng.image.part],
    expectedHeroPng,
    `${label} slide 2 shared PNG bytes match`,
  );

  // Slide 3: linear gradient background
  const s3Bg = validation.slides[2]!.background;
  assert.equal(s3Bg?.type, "gradient", `${label} slide 3 gradient`);
  if (s3Bg?.type === "gradient") {
    assert.equal(s3Bg.kind, "linear");
    assert.equal(s3Bg.angle, 90);
    assert.equal(s3Bg.stops.length, 2);
  }

  // Slide 4: radial gradient background
  const s4Bg = validation.slides[3]!.background;
  assert.equal(s4Bg?.type, "gradient", `${label} slide 4 radial gradient`);
  if (s4Bg?.type === "gradient") {
    assert.equal(s4Bg.kind, "radial");
    assert.equal(s4Bg.stops.length, 2);
  }

  const headline = s1Objects.find(
    (obj) => obj.text?.text === "Regular Bold Italic",
  );
  assert.ok(headline !== undefined, `${label} headline text object`);
  assert.equal(headline.text?.paragraphs[0]?.runs[0]?.fontFamily, "Jost*");
  assert.equal(headline.text?.paragraphs[0]?.runs[1]?.bold, true);
  assert.equal(headline.text?.paragraphs[0]?.runs[2]?.fontFamily, "Jost");
  assert.equal(headline.text?.paragraphs[0]?.runs[2]?.italic, true);

  assert.equal(
    validation.masters?.[0]?.layouts?.[0]?.background?.type,
    "solid",
    `${label} layout background preserved`,
  );

  // Embedded fonts: Jost* regular (400) & bold (700), plus Jost italic (400 italic)
  const fontInfo = inspectEmbeddedFontsInPptx(pptx);
  assert.ok(
    fontInfo.typefaces.includes("Jost*"),
    `${label} embedded typeface includes Jost*: ${JSON.stringify(fontInfo.typefaces)}`,
  );
  assert.ok(
    fontInfo.typefaces.includes("Jost"),
    `${label} embedded typeface includes Jost: ${JSON.stringify(fontInfo.typefaces)}`,
  );
  assert.ok(
    fontInfo.unpackedFaces.some(
      (face) => face.weightClass === 400 && !face.italic,
    ),
    `${label} has 400 normal face`,
  );
  assert.ok(
    fontInfo.unpackedFaces.some(
      (face) => face.weightClass === 700 && !face.italic,
    ),
    `${label} has 700 bold face`,
  );
  assert.ok(
    fontInfo.unpackedFaces.some(
      (face) => face.weightClass === 400 && face.italic,
    ),
    `${label} has 400 italic face`,
  );
}

describe("pptxdk asset codegen acceptance (images, backgrounds, embedded and unembedded fonts)", () => {
  it("generates typechecked, relocated-rebuildable vanilla and react sources with full/subset embedded fonts, images, backgrounds, and external URLs after releasing input bytes", async () => {
    const heroPng = createTwoTonePng();
    const bgJpeg = createSolidJpeg(0x1e, 0x3a, 0x8a);
    const watermarkGif = new Uint8Array(sampleGifBytes);
    const externalUrl = "https://assets.example.com/external-hero.png";

    const localInputDir = await mkdtemp(join(tmpdir(), "pptxdk-assets-input-"));
    const localGifPath = join(localInputDir, "watermark.gif");
    const localJpegPath = join(localInputDir, "bg.jpg");
    await writeFile(localGifPath, watermarkGif);
    await writeFile(localJpegPath, bgJpeg);

    const jostRegularCopy = new Uint8Array(jostRegularSource);
    const jostBoldCopy = new Uint8Array(jostBoldSource);
    const jostItalicCopy = new Uint8Array(jostItalicSource);

    const jostVanilla = vanillaFont("Jost*", [
      { weight: 400, style: "normal", data: jostRegularCopy },
      { weight: 700, style: "normal", data: jostBoldCopy },
    ]);
    const jostItalicVanilla = vanillaFont("Jost", [
      { weight: 400, style: "italic", data: jostItalicCopy },
    ]);
    const brandThemeVanilla = defineTheme({
      name: "Asset Brand",
      colors: { accent1: "#2563EB", accent2: "#7C3AED" },
      fonts: { major: jostVanilla, minor: jostVanilla },
    });
    const blankLayoutVanilla = vanillaDefineSlideLayout({
      id: "asset-blank",
      name: "Asset Blank",
      type: "blank",
      background: lighter("accent1", 0.9),
    });

    const vanillaDeck = deck({
      width: 640,
      height: 360,
      theme: brandThemeVanilla,
    }).append(
      slideMaster({
        id: "master-main",
        theme: brandThemeVanilla,
      }).append(slideLayout(blankLayoutVanilla)),
      slide({
        id: "media-and-fonts",
        layout: blankLayoutVanilla,
        background: "#FFFFFF",
      }).append(
        text(
          [
            paragraph([
              span("Regular ", { style: { fontFamily: jostVanilla } }),
              span("Bold ", {
                style: { fontFamily: jostVanilla, fontWeight: "bold" },
              }),
              span("Italic", {
                style: { fontFamily: jostItalicVanilla, fontStyle: "italic" },
              }),
            ]),
          ],
          {
            id: "font-headline",
            x: 24,
            y: 20,
            width: 360,
            height: 60,
            style: { fontSize: 20, color: "#0F172A" },
          },
        ),
        image({
          id: "hero-cropped",
          src: new Uint8Array(heroPng),
          x: 24,
          y: 100,
          width: 120,
          height: 120,
          crop: { left: 0.5 },
          geometry: preset("roundRect", { adj: 15000 }),
          link: { url: "https://pptxdk.example.com/docs" },
        }),
        image({
          id: "watermark-gif",
          src: localGifPath,
          x: 170,
          y: 100,
          width: 60,
          height: 60,
          flipH: true,
          style: { rotate: 15, opacity: 0.9 },
        }),
        image({
          id: "external-pic",
          src: { externalUrl },
          x: 260,
          y: 100,
          width: 140,
          height: 90,
        }),
      ),
      slide({
        id: "bg-jpeg-and-shared-png",
        layout: blankLayoutVanilla,
        background: { image: pathToFileURL(localJpegPath) },
      }).append(
        image({
          id: "shared-png-reuse",
          src: new Uint8Array(heroPng),
          x: 40,
          y: 40,
          width: 120,
          height: 60,
        }),
        text("Card Fill", {
          id: "card-gradient-fill",
          x: 200,
          y: 40,
          width: 240,
          height: 80,
          style: {
            color: "#FFFFFF",
            backgroundImage: vanillaLinearGradient(
              90,
              "#0F172A",
              darker("accent1", 0.2),
            ),
          },
        }),
      ),
      slide({
        id: "bg-linear-grad",
        layout: blankLayoutVanilla,
        background: vanillaLinearGradient(
          90,
          { color: lighter("accent1", 0.35), position: 0 },
          { color: "#38BDF8", position: 1 },
        ),
      }),
      slide({
        id: "bg-radial-grad",
        layout: blankLayoutVanilla,
        background: vanillaRadialGradient(
          { color: "#FEF3C7", position: 0 },
          { color: darker("accent2", 0.25), position: 1 },
        ),
      }),
    );

    const jostReact = font("Jost*", [
      { weight: 400, style: "normal", data: jostRegularCopy },
      { weight: 700, style: "normal", data: jostBoldCopy },
    ]);
    const jostItalicReact = font("Jost", [
      { weight: 400, style: "italic", data: jostItalicCopy },
    ]);
    const brandThemeReact = defineTheme({
      name: "Asset Brand",
      colors: { accent1: "#2563EB", accent2: "#7C3AED" },
      fonts: { major: jostReact, minor: jostReact },
    });
    const blankLayoutReact = defineSlideLayout({
      id: "asset-blank",
      name: "Asset Blank",
      type: "blank",
      background: lighter("accent1", 0.9),
    });
    const reactDeck = await expandDocument(
      createElement(
        Deck,
        { width: 640, height: 360, theme: brandThemeReact },
        createElement(
          SlideMaster,
          { id: "master-main", theme: brandThemeReact },
          createElement(SlideLayout, { layout: blankLayoutReact }),
        ),
        createElement(
          Slide,
          {
            id: "media-and-fonts",
            layout: blankLayoutReact,
            background: "#FFFFFF",
          },
          createElement(
            Text,
            {
              id: "font-headline",
              x: 24,
              y: 20,
              width: 360,
              height: 60,
              style: { fontSize: 20, color: "#0F172A" },
            },
            createElement(
              P,
              null,
              createElement(
                Span,
                { style: { fontFamily: jostReact } },
                "Regular ",
              ),
              createElement(
                Span,
                { style: { fontFamily: jostReact, fontWeight: "bold" } },
                "Bold ",
              ),
              createElement(
                Span,
                { style: { fontFamily: jostItalicReact, fontStyle: "italic" } },
                "Italic",
              ),
            ),
          ),
          createElement(Image, {
            id: "hero-cropped",
            src: new Uint8Array(heroPng),
            x: 24,
            y: 100,
            width: 120,
            height: 120,
            crop: { left: 0.5 },
            geometry: preset("roundRect", { adj: 15000 }),
            link: { url: "https://pptxdk.example.com/docs" },
          }),
          createElement(Image, {
            id: "watermark-gif",
            src: new Uint8Array(watermarkGif),
            x: 170,
            y: 100,
            width: 60,
            height: 60,
            flipH: true,
            style: { rotate: 15, opacity: 0.9 },
          }),
          createElement(Image, {
            id: "external-pic",
            src: { externalUrl },
            x: 260,
            y: 100,
            width: 140,
            height: 90,
          }),
        ),
        createElement(
          Slide,
          {
            id: "bg-jpeg-and-shared-png",
            layout: blankLayoutReact,
            background: { image: new Uint8Array(bgJpeg) },
          },
          createElement(Image, {
            id: "shared-png-reuse",
            src: new Uint8Array(heroPng),
            x: 40,
            y: 40,
            width: 120,
            height: 60,
          }),
          createElement(
            Text,
            {
              id: "card-gradient-fill",
              x: 200,
              y: 40,
              width: 240,
              height: 80,
              style: {
                color: "#FFFFFF",
                backgroundImage: linearGradient(
                  90,
                  "#0F172A",
                  darker("accent1", 0.2),
                ),
              },
            },
            "Card Fill",
          ),
        ),
        createElement(Slide, {
          id: "bg-linear-grad",
          layout: blankLayoutReact,
          background: linearGradient(
            90,
            { color: lighter("accent1", 0.35), position: 0 },
            { color: "#38BDF8", position: 1 },
          ),
        }),
        createElement(Slide, {
          id: "bg-radial-grad",
          layout: blankLayoutReact,
          background: radialGradient(
            { color: "#FEF3C7", position: 0 },
            { color: darker("accent2", 0.25), position: 1 },
          ),
        }),
      ),
    );

    // Build real PPTX samples with BOTH embedFonts: "full" and embedFonts: "subset",
    // import them via importPptx, and immediately zero the input PPTX buffers.
    const fullBuilt = await build(vanillaDeck, { embedFonts: "full" });
    const fullCopy = new Uint8Array(fullBuilt.pptx);
    const fullImported = importPptx(fullCopy);
    fullCopy.fill(0);
    assert.deepEqual(fullImported.report.diagnostics, []);

    const subsetBuilt = await build(vanillaDeck, { embedFonts: "subset" });
    const subsetCopy = new Uint8Array(subsetBuilt.pptx);
    const subsetImported = importPptx(subsetCopy);
    subsetCopy.fill(0);
    assert.deepEqual(subsetImported.report.diagnostics, []);

    // Also zero the caller font copies used to construct the direct decks after generating sources.
    const sourceDecks = [
      {
        label: "direct-vanilla",
        doc: vanillaDeck,
        embedMode: "full" as BuildOptions["embedFonts"],
      },
      {
        label: "direct-react",
        doc: reactDeck,
        embedMode: "full" as BuildOptions["embedFonts"],
      },
      {
        label: "imported-full",
        doc: fullImported.document,
        embedMode: "full" as BuildOptions["embedFonts"],
      },
      {
        label: "imported-subset",
        doc: subsetImported.document,
        embedMode: "subset" as BuildOptions["embedFonts"],
      },
    ] as const;

    const generatedCases = sourceDecks.flatMap(({ label, doc, embedMode }) => {
      const vanillaResult = codegen(doc, { target: "vanilla" });
      const reactResult = codegen(doc, { target: "react" });
      return [
        { label: `${label}:vanilla`, result: vanillaResult, embedMode },
        { label: `${label}:react`, result: reactResult, embedMode },
      ];
    });

    // Zero caller input byte copies and remove original local image files before running consumer typecheck & relocated rebuild.
    jostRegularCopy.fill(0);
    jostBoldCopy.fill(0);
    jostItalicCopy.fill(0);
    await rm(localInputDir, { recursive: true, force: true });

    await verifyGeneratedSourcesTypecheck(generatedCases);

    for (const item of generatedCases) {
      // Standard relocated rebuild (vanilla isolation check + validatePptx)
      const standardOutcome = await rebuildGeneratedSourceRelocated(
        item.result,
      );
      assert.deepEqual(standardOutcome.file.errors, [], item.label);

      // Relocated rebuild with explicit embedFonts ("full" / "subset") + full asset/font verification
      await withRelocatedConsumerProject(item.result, async ({ entryPath }) => {
        const rebuilt = await buildFromSource(entryPath, {
          embedFonts: item.embedMode,
        });
        await assertAssetDeckRebuild(
          rebuilt.pptx,
          `${item.label}:${item.embedMode}`,
          heroPng,
          bgJpeg,
          watermarkGif,
          externalUrl,
        );
      });
    }
  }, 90_000);

  it("verifies static preview pixels on relocated sources and pinpoints missing or corrupt embedded image assets", async () => {
    const heroPng = createTwoTonePng();
    const bgJpeg = createSolidJpeg(0x0f, 0x76, 0x6e);

    const doc = deck({ width: 320, height: 180 }).append(
      slide({ id: "img-slide", background: "#FFFFFF" }).append(
        image({
          id: "cropped-hero",
          src: heroPng,
          x: 20,
          y: 20,
          width: 60,
          height: 60,
          crop: { left: 0.5 },
        }),
      ),
      slide({
        id: "bg-slide",
        background: { image: bgJpeg },
      }),
    );

    for (const target of ["vanilla", "react"] as const) {
      const generated = codegen(doc, { target });
      const pngAssetKey = Object.keys(generated.files).find(
        (key) => key.startsWith("assets/image-") && key.endsWith(".png"),
      )!;
      assert.ok(pngAssetKey !== undefined);

      await withRelocatedConsumerProject(
        generated,
        async ({ relocatedDir, entryPath }) => {
          // 1. Healthy preview renders expected cropped blue pixels on slide 1 & teal bg on slide 2.
          const healthyPreview = await previewFromSource(entryPath);
          assert.equal(healthyPreview.pages.length, 2);
          assert.deepEqual(healthyPreview.pages[0]?.diagnostics, []);
          assert.deepEqual(healthyPreview.pages[1]?.diagnostics, []);
          assert.ok(healthyPreview.pages[0]?.png !== undefined);
          assert.ok(healthyPreview.pages[1]?.png !== undefined);

          const [r1, , b1] = readPngPixel(healthyPreview.pages[0].png, 50, 50);
          assert.ok(r1 < 35 && b1 > 190, "cropped hero right half is blue");
          const [r2, g2, b2] = readPngPixel(
            healthyPreview.pages[1].png,
            100,
            90,
          );
          assert.ok(r2 < 40 && g2 > 90 && b2 > 80, "slide 2 JPEG bg is teal");

          // 2. Corrupt compressed IDAT stream of the PNG asset -> preview pinpoints damaged image while slide 2 still renders.
          const assetDiskPath = join(relocatedDir, pngAssetKey);
          const originalBytes = new Uint8Array(await readFile(assetDiskPath));
          await writeFile(assetDiskPath, corruptPngIdatStream(originalBytes));

          const corruptPreview = await previewFromSource(entryPath);
          assert.equal(corruptPreview.pages[0]?.png, undefined);
          assert.equal(
            corruptPreview.pages[0]?.diagnostics[0]?.severity,
            "error",
          );
          assert.match(
            corruptPreview.pages[0]?.diagnostics[0]?.path ?? "",
            /image\[cropped-hero\]/u,
          );
          assert.match(
            corruptPreview.pages[0]?.diagnostics[0]?.message ?? "",
            /decode PNG/iu,
          );
          assert.ok(corruptPreview.pages[1]?.png !== undefined);
          assert.deepEqual(corruptPreview.pages[1]?.diagnostics, []);

          // 3. Missing image asset file -> buildFromSource fails locating the missing resource path.
          await rm(assetDiskPath);
          await assert.rejects(
            () => buildFromSource(entryPath),
            (error: unknown) =>
              error instanceof Error &&
              /image-read-failed|Failed to read image resource/iu.test(
                error.message,
              ),
          );
        },
      );
    }
  }, 60_000);

  it("produces 100% deterministic entries, paths, source, and asset bytes across repeated codegen runs and handles unembedded fonts with/without extra font resources", async () => {
    const sharedPng = createTwoTonePng();
    const sharedFont = font("Jost*", [
      { weight: 400, style: "normal", data: jostRegularSource },
    ]);
    const sharedTheme = defineTheme({
      name: "Shared Assets",
      fonts: { major: sharedFont, minor: sharedFont },
    });

    const deterministicDoc = deck({
      width: 640,
      height: 360,
      theme: sharedTheme,
    }).append(
      slide({
        id: "shared-1",
        background: { image: sharedPng },
      }).append(
        image({
          id: "shared-img",
          src: new Uint8Array(sharedPng),
          x: 40,
          y: 40,
          width: 120,
          height: 60,
        }),
        text("Shared Asset Text", {
          id: "shared-txt",
          x: 40,
          y: 120,
          width: 300,
          height: 60,
          style: {
            fontFamily: sharedFont,
            backgroundImage: { image: new Uint8Array(sharedPng) },
          },
        }),
      ),
    );

    for (const target of ["vanilla", "react"] as const) {
      const first = codegen(deterministicDoc, { target });
      const second = codegen(deterministicDoc, { target });
      assert.equal(first.entry, second.entry);
      const keys = Object.keys(first.files);
      assert.deepEqual(keys, Object.keys(second.files));
      // 1 entry + 1 deduplicated image asset + 1 deduplicated font asset
      assert.equal(keys.length, 3, `${target} deduplicated file count`);
      for (const key of keys) {
        assert.deepEqual(first.files[key], second.files[key]);
      }
    }

    // Unembedded font handling:
    // 1) Without extra font environment: PPTX preserves declared fontFamily, preview reports missing outline font diagnostic.
    const unembeddedDeck = deck({ width: 480, height: 270 }).append(
      slide({ id: "unembedded" }).append(
        text("Unembedded Family", {
          id: "plain-text",
          x: 24,
          y: 24,
          width: 280,
          height: 60,
          style: { fontFamily: "MissingPreviewFamily", fontSize: 18 },
        }),
      ),
    );
    const unembeddedPptx = new Uint8Array((await build(unembeddedDeck)).pptx);
    const unembeddedImported = importPptx(unembeddedPptx);
    unembeddedPptx.fill(0);
    assert.deepEqual(unembeddedImported.report.diagnostics, []);

    const unembeddedCases = (["vanilla", "react"] as const).map((target) => ({
      label: `unembedded:${target}`,
      result: codegen(unembeddedImported.document, { target }),
    }));
    await verifyGeneratedSourcesTypecheck(unembeddedCases);

    for (const item of unembeddedCases) {
      assert.equal(Object.keys(item.result.files).length, 1);
      await withRelocatedConsumerProject(item.result, async ({ entryPath }) => {
        const rebuilt = await buildFromSource(entryPath);
        const validation = await validatePptx(rebuilt.pptx);
        assert.deepEqual(validation.errors, []);
        assert.equal(
          validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
            ?.fontFamily,
          "MissingPreviewFamily",
        );

        const noExtraFontPreview = await previewFromSource(entryPath);
        assert.equal(noExtraFontPreview.pages[0]?.png, undefined);
        assert.ok(
          (noExtraFontPreview.pages[0]?.diagnostics.length ?? 0) > 0,
          "preview reports missing outline font when no extra font environment is provided",
        );
      });
    }

    // 2) Observation condition with explicit extra font resource (`substitute: true` supplies outlines for preview/measurement
    //    while keeping the declared output family name unembedded in PPTX).
    const withSubstituteFont = deck({ width: 480, height: 270 }).append(
      slide({ id: "with-extra-font" }).append(
        text("Observed With Extra Font", {
          id: "sub-text",
          x: 24,
          y: 24,
          width: 280,
          height: 60,
          style: {
            fontFamily: font(
              "MissingPreviewFamily",
              [{ weight: 400, style: "normal", data: jostRegularSource }],
              { substitute: true },
            ),
            fontSize: 18,
          },
        }),
      ),
    );
    const subGenerated = codegen(withSubstituteFont, { target: "vanilla" });
    await verifyGeneratedSourcesTypecheck([
      { label: "extra-font-env", result: subGenerated },
    ]);
    await withRelocatedConsumerProject(subGenerated, async ({ entryPath }) => {
      const previewWithExtraFont = await previewFromSource(entryPath);
      assert.deepEqual(previewWithExtraFont.pages[0]?.diagnostics, []);
      assert.ok(previewWithExtraFont.pages[0]?.png instanceof Uint8Array);

      const builtUnembedded = await buildFromSource(entryPath);
      const validation = await validatePptx(builtUnembedded.pptx);
      assert.deepEqual(validation.errors, []);
      assert.equal(
        validation.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]
          ?.fontFamily,
        "MissingPreviewFamily",
      );
      assert.deepEqual(
        inspectEmbeddedFontsInPptx(builtUnembedded.pptx).typefaces,
        [],
      );
    });
  }, 60_000);
});
