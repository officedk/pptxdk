import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { CodegenError, createCodegen } from "@pptxdk/codegen";
import {
  DeckContainer,
  defaultTheme,
  defineTheme,
  lighter,
} from "@pptxdk/core";
import { createImporter } from "@pptxdk/import";
import { describe, it } from "vitest";

import { documentCapabilities } from "../../all.ts";
import {
  deck,
  defineSlideLayout,
  section,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "../../document.ts";

const importDocument = createImporter(documentCapabilities);

describe("masters capability codegen", () => {
  it("generates 4-section source for dual masters, unused layouts, headerFooter, footerPlaceholders, and equal explicit overrides", () => {
    const brandTheme = defineTheme({
      name: "Brand Theme",
      colors: {
        accent1: "#1D4ED8",
        brandSoft: lighter("accent1", 0.35),
      },
    });

    const m1TitleLayout = defineSlideLayout({
      id: "m1-title",
      name: "Title Slide",
      type: "title",
      preserve: true,
      headerFooter: { footer: true, slideNumber: true },
      footerPlaceholders: {
        footer: { text: "Layout Footer", style: { fontSize: "10pt" } },
      },
      placeholders: {
        title: {
          type: "ctrTitle",
          style: {
            position: "absolute",
            left: 60,
            top: 120,
            width: 840,
            height: 80,
            fontSize: "28pt",
            color: "#0F172A",
          },
        },
        "body-1": {
          type: "body",
          idx: 1,
          style: {
            position: "absolute",
            left: 60,
            top: 220,
            width: 400,
            height: 200,
          },
        },
        "body-2": {
          type: "body",
          idx: 2,
          style: {
            position: "absolute",
            left: 500,
            top: 220,
            width: 400,
            height: 200,
          },
        },
      },
    });

    const m1UnusedLayout = defineSlideLayout({
      id: "m1-blank",
      name: "Blank",
      type: "blank",
      preserve: true,
    });

    const m2UnusedLayout = defineSlideLayout({
      id: "m2-blank",
      name: "Blank",
      type: "blank",
      preserve: true,
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "m1",
        name: "Primary Master",
        theme: brandTheme,
        headerFooter: { dateTime: false, footer: true, slideNumber: true },
        footerPlaceholders: {
          dateTime: { text: "2026-01-01", visible: false },
          footer: { text: "Confidential" },
          slideNumber: { idx: 4 },
        },
      }).append(
        text("Master Watermark", {
          id: "m1-watermark",
          style: {
            position: "absolute",
            left: 24,
            top: 16,
            width: 240,
            height: 28,
          },
        }),
        slideLayout(m1TitleLayout),
        slideLayout(m1UnusedLayout),
      ),
      slideMaster({
        id: "m2",
        name: "Secondary Master",
        preserve: true,
        theme: defaultTheme,
      }).append(slideLayout(m2UnusedLayout)),
      section({ name: "Deck Slides" }).append(
        slide({
          id: "s-inherited",
          layout: m1TitleLayout,
          footer: "Custom Slide Footer",
          dateTime: "2026-11-24",
          slideNumber: true,
        }).append(
          text("Inherited Title", { placeholder: "title" }),
          text("Left Column", { placeholder: "body-1" }),
          text("Right Column", { placeholder: "body-2" }),
        ),
        slide({
          id: "s-explicit-equal",
          layout: m1TitleLayout,
        }).append(
          text("Locked Title", {
            placeholder: "title",
            style: { fontSize: "28pt" },
          }),
        ),
      ),
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(doc, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;

    const themeIdx = vanillaSource.indexOf("defineTheme(");
    const layoutIdx = vanillaSource.indexOf("defineSlideLayout(");
    const masterIdx = vanillaSource.indexOf("slideMaster(");
    const slideIdx = vanillaSource.indexOf("const sInheritedSlide = slide(");
    const deckIdx = vanillaSource.indexOf("export default deck(");
    assert.ok(
      themeIdx >= 0 &&
        layoutIdx > themeIdx &&
        masterIdx > layoutIdx &&
        slideIdx > masterIdx &&
        deckIdx > slideIdx,
    );
    assert.match(vanillaSource, /lighter\("accent1", 0\.35\)/u);
    assert.match(vanillaSource, /theme:\s*defaultTheme/u);
    assert.match(vanillaSource, /name:\s*"Blank"/u);
    assert.match(vanillaSource, /headerFooter:\s*\{/u);
    assert.match(vanillaSource, /footerPlaceholders:\s*\{/u);
    assert.match(vanillaSource, /footer:\s*"Custom Slide Footer"/u);
    assert.match(vanillaSource, /dateTime:\s*"2026-11-24"/u);
    assert.match(vanillaSource, /slideNumber:\s*true/u);
    assert.match(vanillaSource, /placeholder:\s*"body-1"/u);
    assert.match(vanillaSource, /placeholder:\s*"body-2"/u);
    // Equal explicit override on s-explicit-equal title must not be stripped.
    assert.match(vanillaSource, /"Locked Title"[\s\S]*?fontSize:\s*"28pt"/u);

    const react = generate(doc, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /function M1\(/u);
    assert.match(reactSource, /function M2\(/u);
    assert.match(reactSource, /<SlideLayout layout=\{/u);
    assert.match(reactSource, /headerFooter=\{\{/u);
    assert.match(reactSource, /footerPlaceholders=\{\{/u);
    assert.match(reactSource, /footer=\{"Custom Slide Footer"\}/u);
    assert.match(reactSource, /dateTime=\{"2026-11-24"\}/u);
    assert.match(reactSource, /slideNumber=\{true\}/u);
    assert.match(reactSource, /placeholder="body-1"/u);
    assert.match(reactSource, /placeholder="body-2"/u);
    assert.match(reactSource, /fontSize:\s*"28pt"[\s\S]*?>\s*Locked Title/u);
  });

  it("inspects prs-slide-masters.pptx AST and generates dual masters, backgrounds, themes, and idle layouts", () => {
    const raw = new Uint8Array(
      readFileSync(
        new URL("./fixtures/prs-slide-masters.pptx", import.meta.url),
      ),
    );
    const imported = importDocument(raw);
    raw.fill(0);

    assert.equal(imported.report.slideCount, 0);
    assert.ok(imported.document instanceof DeckContainer);
    assert.equal(imported.document.masters.length, 2);
    const [m1, m2] = imported.document.masters;
    assert.equal(m1?.id, "master-1");
    assert.equal(m2?.id, "master-2");
    assert.equal(m1?.props["preserve"], undefined);
    assert.equal(m2?.props["preserve"], true);
    assert.equal(m1?.layouts.length, 1);
    assert.equal(m2?.layouts.length, 1);
    assert.equal(m1?.layouts[0]?.props["name"], "Blank");
    assert.equal(m2?.layouts[0]?.props["name"], "Blank");
    assert.deepEqual(
      { ...(m1?.props["background"] as Record<string, unknown>) },
      { ref: 1001, color: "bg1" },
    );
    assert.deepEqual(
      { ...(m2?.props["background"] as Record<string, unknown>) },
      { ref: 1001, color: "bg1" },
    );

    const generate = createCodegen(documentCapabilities);

    const vanilla = generate(imported.document, { target: "vanilla" });
    const vanillaSource = vanilla.files[vanilla.entry] as string;
    assert.match(vanillaSource, /name:\s*"Slide Master 1"/u);
    assert.match(vanillaSource, /name:\s*"Slide Master 2"/u);
    assert.match(
      vanillaSource,
      /const master1BlankLayout = defineSlideLayout\(/u,
    );
    assert.match(
      vanillaSource,
      /const master2BlankLayout = defineSlideLayout\(/u,
    );
    assert.match(vanillaSource, /preserve:\s*true/u);
    assert.match(
      vanillaSource,
      /background:\s*\{\s*ref:\s*1001,\s*color:\s*"bg1"/u,
    );

    const react = generate(imported.document, { target: "react" });
    const reactSource = react.files[react.entry] as string;
    assert.match(reactSource, /function Master1\(/u);
    assert.match(reactSource, /function Master2\(/u);
    assert.match(
      reactSource,
      /<SlideLayout layout=\{master1BlankLayout\} \/>/u,
    );
    assert.match(
      reactSource,
      /<SlideLayout layout=\{master2BlankLayout\} \/>/u,
    );
    assert.match(
      reactSource,
      /background=\{\{\s*ref:\s*1001,\s*color:\s*"bg1"/u,
    );
  });

  it("deterministically disambiguates colliding theme, master, and layout names and reports invalid headerFooter/footerPlaceholders", () => {
    const themeA = defineTheme({
      name: "Shared Name",
      colors: { accent1: "#111111" },
    });
    const themeB = defineTheme({
      name: "Shared Name",
      colors: { accent1: "#222222" },
    });

    const doc = deck({ width: 960, height: 540 }).append(
      slideMaster({ name: "Corporate", theme: themeA }).append(
        slideLayout({ id: "blank-a", name: "Blank", type: "blank" }),
      ),
      slideMaster({ name: "Corporate", theme: themeB }).append(
        slideLayout({ id: "blank-b", name: "Blank", type: "blank" }),
      ),
    );

    const generate = createCodegen(documentCapabilities);
    for (const target of ["vanilla", "react"] as const) {
      const first = generate(doc, { target });
      const second = generate(doc, { target });
      assert.equal(first.files[first.entry], second.files[second.entry]);
      const source = first.files[first.entry] as string;
      assert.match(source, /const sharedNameTheme = defineTheme\(/u);
      assert.match(source, /const sharedNameTheme2 = defineTheme\(/u);
    }

    const badHeaderFooterDoc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "m1",
        headerFooter: { footer: "yes" as never },
      }),
    );
    assert.throws(
      () => generate(badHeaderFooterDoc, { target: "vanilla" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path === "deck/slide-master[m1]@headerFooter",
    );

    const badFooterPlaceholdersDoc = deck({ width: 960, height: 540 }).append(
      slideMaster({
        id: "m1",
        footerPlaceholders: { unknownSlot: {} } as never,
      }),
    );
    assert.throws(
      () => generate(badFooterPlaceholdersDoc, { target: "react" }),
      (error: unknown) =>
        error instanceof CodegenError &&
        error.diagnostics[0]?.code === "unexpressible-value" &&
        error.diagnostics[0]?.path ===
          "deck/slide-master[m1]@footerPlaceholders",
    );
  });
});
