import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";

import { createCodegen } from "@pptxdk/codegen";
import { DeckContainer, defineTheme, lighter } from "@pptxdk/core";
import { createResolver } from "@pptxdk/core/build";
import { emit } from "@pptxdk/emit";
import { createImporter } from "@pptxdk/import";
import {
  Deck,
  Slide,
  SlideLayout,
  SlideMaster,
  expandDocument,
} from "@pptxdk/react";
import { validatePptx } from "@pptxdk/validation";
import { PNG } from "pngjs";
import { createElement } from "react";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  deck,
  defineSlideLayout,
  font,
  image,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "../../document.ts";
import { Image } from "../../image/image-elements.ts";
import { Text } from "../../text/text-elements.ts";

const BRAND_FIXTURE_SHA256 =
  "de3376969353a2322ccce48182393f63eea9a44f5a231565985bcb5d3d336536";

const importDocument = createImporter(documentCapabilities);
const resolveDocument = createResolver(documentCapabilities);
const generateSource = createCodegen(documentCapabilities);

function readEntrySource(
  files: Readonly<Record<string, string | Uint8Array>>,
  entry: string,
): string {
  const value = files[entry];
  if (typeof value !== "string") {
    throw new Error(`Expected string entry file ${entry}`);
  }
  return value;
}

const jostRegularBytes = new Uint8Array(
  readFileSync(
    new URL("../../../test/assets/fonts/Jost-400-Book.ttf", import.meta.url),
  ),
);

function createBrandBadgePng(): Uint8Array {
  const png = new PNG({ width: 32, height: 16 });
  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      png.data.set(
        x < 16 ? [29, 78, 216, 255] : [15, 23, 42, 255],
        (y * png.width + x) * 4,
      );
    }
  }
  return new Uint8Array(PNG.sync.write(png));
}

describe("brand master capability codegen", () => {
  it("imports PowerPoint brand-master-inheritance-override.pptx and generates deterministic vanilla & react sources preserving background, shared badge, footer, idle layout, and equal explicit override", async () => {
    const raw = new Uint8Array(
      readFileSync(
        new URL(
          "./fixtures/brand-master-inheritance-override.pptx",
          import.meta.url,
        ),
      ),
    );
    assert.equal(
      createHash("sha256").update(raw).digest("hex"),
      BRAND_FIXTURE_SHA256,
    );
    const imported = importDocument(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 2);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 1);
    assert.equal(imported.document.masters[0]?.layouts.length, 2);

    const vanilla1 = generateSource(imported.document, { target: "vanilla" });
    const vanilla2 = generateSource(imported.document, { target: "vanilla" });
    assert.equal(
      vanilla1.files[vanilla1.entry],
      vanilla2.files[vanilla2.entry],
    );
    const vanSrc = readEntrySource(vanilla1.files, vanilla1.entry);

    // 4-section ordering: theme -> layouts/master -> slides -> deck export
    const themeIdx = vanSrc.indexOf("defineTheme(");
    const heroLayoutIdx = vanSrc.indexOf(
      "const brandMasterBrandHeroLayout = defineSlideLayout(",
    );
    const idleLayoutIdx = vanSrc.indexOf(
      "const brandMasterBrandIdleContentLayout = defineSlideLayout(",
    );
    const masterIdx = vanSrc.indexOf("const brandMaster = slideMaster(");
    const slide1Idx = vanSrc.indexOf("const inheritedBrandSlide = slide(");
    const deckIdx = vanSrc.indexOf("export default deck(");
    assert.ok(
      themeIdx >= 0 &&
        heroLayoutIdx > themeIdx &&
        idleLayoutIdx > heroLayoutIdx &&
        masterIdx > idleLayoutIdx &&
        slide1Idx > masterIdx &&
        deckIdx > slide1Idx,
      "Expected 4-section top-to-bottom ordering in generated vanilla source",
    );

    assert.match(vanSrc, /background:\s*"#F8FAFC"/u);
    assert.match(vanSrc, /name:\s*"BrandBadge"/u);
    assert.match(vanSrc, /"Confidential Brand Deck"/u);
    assert.match(vanSrc, /name:\s*"Brand Idle Content"/u);

    const s1Section = vanSrc.slice(
      slide1Idx,
      vanSrc.indexOf("const explicitOverrideSlide"),
    );
    const s2Section = vanSrc.slice(
      vanSrc.indexOf("const explicitOverrideSlide"),
      deckIdx,
    );
    assert.doesNotMatch(s1Section, /fontSize:\s*"32pt"/u);
    assert.match(s2Section, /fontSize:\s*"32pt"/u);

    const react1 = generateSource(imported.document, { target: "react" });
    const react2 = generateSource(imported.document, { target: "react" });
    assert.equal(react1.files[react1.entry], react2.files[react2.entry]);
    const reactSrc = readEntrySource(react1.files, react1.entry);
    assert.match(reactSrc, /function BrandMaster\(/u);
    assert.match(
      reactSrc,
      /<SlideLayout layout=\{brandMasterBrandIdleContentLayout\} \/>/u,
    );
    assert.match(reactSrc, /Locked Brand Title/u);
    assert.match(reactSrc, /fontSize:\s*"32pt"/u);

    // Pure-file roundtrip check without reattaching external assets.
    const pureValidated = await validatePptx(
      emit(resolveDocument(imported.document)).pptx,
    );
    assert.deepEqual(pureValidated.errors, []);
    assert.equal(pureValidated.masters?.length, 1);
    assert.equal(pureValidated.masters?.[0]?.layouts.length, 2);
    assert.equal(pureValidated.masters?.[0]?.objects?.[0]?.name, "BrandBadge");
    assert.equal(
      pureValidated.slides[0]?.objects.some((o) => o.name === "BrandBadge"),
      false,
    );
  });

  it("generates vanilla & react sources for direct dual-author brand masters with embedded image & font assets, background, footer, shared objects, and idle layout", async () => {
    const logoPng = createBrandBadgePng();
    const jostFont = font("Jost*", [{ weight: 400, data: jostRegularBytes }]);
    const brandTheme = defineTheme({
      name: "Apex Brand",
      colors: {
        accent1: "#1D4ED8",
        accentSoft: lighter("accent1", 0.35),
      },
      fonts: { major: jostFont, minor: jostFont },
    });

    const heroLayout = defineSlideLayout({
      id: "hero",
      name: "Hero",
      type: "title",
      preserve: true,
      headerFooter: { dateTime: false, footer: true, slideNumber: true },
      footerPlaceholders: {
        footer: { text: "Apex Internal", style: { fontSize: "10pt" } },
      },
      placeholders: {
        title: {
          type: "ctrTitle",
          style: {
            position: "absolute",
            left: 64,
            top: 88,
            width: 832,
            height: 80,
            fontSize: "32pt",
            color: "#0F172A",
          },
        },
        subtitle: {
          type: "subTitle",
          idx: 1,
          style: {
            position: "absolute",
            left: 64,
            top: 192,
            width: 832,
            height: 120,
          },
        },
      },
    });

    const idleAgendaLayout = defineSlideLayout({
      id: "agenda-idle",
      name: "Agenda Idle",
      type: "obj",
      preserve: true,
      placeholders: {
        title: {
          type: "title",
          style: {
            position: "absolute",
            left: 64,
            top: 56,
            width: 832,
            height: 64,
            fontSize: "24pt",
            color: "#1D4ED8",
          },
        },
        body: {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 64,
            top: 144,
            width: 832,
            height: 280,
          },
        },
      },
    });

    const vanillaDoc = deck({
      width: 960,
      height: 540,
      theme: brandTheme,
    }).append(
      slideMaster({
        id: "apex-master",
        name: "Apex Master",
        theme: brandTheme,
        style: { background: "#F8FAFC" },
        headerFooter: { dateTime: false, footer: true, slideNumber: true },
        footerPlaceholders: {
          footer: { text: "Apex Internal" },
          slideNumber: { idx: 4 },
        },
      }).append(
        image({
          src: logoPng,
          id: "brand-logo",
          name: "Brand Logo",
          style: {
            position: "absolute",
            left: 24,
            top: 16,
            width: 64,
            height: 32,
          },
        }),
        text("APEX", {
          id: "brand-wordmark",
          name: "Brand Wordmark",
          style: {
            position: "absolute",
            left: 96,
            top: 16,
            width: 140,
            height: 32,
            fontSize: "14pt",
            color: "#1D4ED8",
          },
        }),
        slideLayout(heroLayout),
        slideLayout(idleAgendaLayout),
      ),
      slide({
        id: "s-inherited",
        layout: heroLayout,
        footer: "Apex Internal",
        slideNumber: true,
      }).append(
        text("Inherited Headline", { placeholder: "title" }),
        text("Layout-driven size", { placeholder: "subtitle" }),
      ),
      slide({
        id: "s-explicit",
        layout: heroLayout,
        footer: "Apex Internal",
        slideNumber: true,
      }).append(
        text("Locked Headline", {
          placeholder: "title",
          style: { fontSize: "32pt" },
        }),
        text("Equal explicit size", { placeholder: "subtitle" }),
      ),
    );

    const reactDoc = await expandDocument(
      createElement(
        Deck,
        { width: 960, height: 540, theme: brandTheme },
        createElement(
          SlideMaster,
          {
            id: "apex-master",
            name: "Apex Master",
            theme: brandTheme,
            style: { background: "#F8FAFC" },
            headerFooter: { dateTime: false, footer: true, slideNumber: true },
            footerPlaceholders: {
              footer: { text: "Apex Internal" },
              slideNumber: { idx: 4 },
            },
          },
          createElement(Image, {
            id: "brand-logo",
            name: "Brand Logo",
            src: logoPng,
            style: {
              position: "absolute",
              left: 24,
              top: 16,
              width: 64,
              height: 32,
            },
          }),
          createElement(
            Text,
            {
              id: "brand-wordmark",
              name: "Brand Wordmark",
              style: {
                position: "absolute",
                left: 96,
                top: 16,
                width: 140,
                height: 32,
                fontSize: "14pt",
                color: "#1D4ED8",
              },
            },
            "APEX",
          ),
          createElement(SlideLayout, { layout: heroLayout }),
          createElement(SlideLayout, { layout: idleAgendaLayout }),
        ),
        createElement(
          Slide,
          {
            id: "s-inherited",
            layout: heroLayout,
            footer: "Apex Internal",
            slideNumber: true,
          },
          createElement(Text, { placeholder: "title" }, "Inherited Headline"),
          createElement(
            Text,
            { placeholder: "subtitle" },
            "Layout-driven size",
          ),
        ),
        createElement(
          Slide,
          {
            id: "s-explicit",
            layout: heroLayout,
            footer: "Apex Internal",
            slideNumber: true,
          },
          createElement(
            Text,
            { placeholder: "title", style: { fontSize: "32pt" } },
            "Locked Headline",
          ),
          createElement(
            Text,
            { placeholder: "subtitle" },
            "Equal explicit size",
          ),
        ),
      ),
    );

    for (const [label, doc] of [
      ["vanilla-author", vanillaDoc],
      ["react-author", reactDoc],
    ] as const) {
      const van = generateSource(doc, { target: "vanilla" });
      const react = generateSource(doc, { target: "react" });
      for (const res of [van, react]) {
        const assetPaths = Object.keys(res.files).filter((k) =>
          k.startsWith("assets/"),
        );
        assert.equal(
          assetPaths.some((p) => p.endsWith(".png")),
          true,
          `${label} -> ${res.target} must emit embedded PNG asset`,
        );
        assert.equal(
          assetPaths.some((p) => p.endsWith(".ttf")),
          true,
          `${label} -> ${res.target} must emit embedded TTF font asset`,
        );
        const src = readEntrySource(res.files, res.entry);
        assert.match(src, /background[:=]/u);
        assert.match(src, /Brand Logo/u);
        assert.match(src, /Brand Wordmark/u);
        assert.match(src, /Agenda Idle/u);
        assert.match(src, /Locked Headline/u);
        assert.match(src, /fontSize:\s*"32pt"/u);
      }
    }
  });
});
