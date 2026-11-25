import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import {
  DeckContainer,
  build,
  codegen,
  defaultTheme,
  defineTheme,
  importPptx,
  lighter,
} from "@officedk/pptxdk";
import {
  Deck,
  Placeholder,
  Section,
  Slide,
  SlideLayout,
  SlideMaster,
  Text,
  expandDocument,
} from "@officedk/pptxdk/react";
import {
  deck,
  defineSlideLayout,
  placeholder,
  section,
  slide,
  slideLayout,
  slideMaster,
  text,
} from "@officedk/pptxdk/vanilla";
import { unzipSync, strFromU8 } from "fflate";
import { createElement } from "react";
import { describe, it } from "vitest";

import {
  rebuildGeneratedSourceRelocated,
  verifyGeneratedSourcesTypecheck,
} from "./helpers/codegen-acceptance.ts";

function buildDualMasterFixtures() {
  const m1Theme = defineTheme({
    name: "Studio Light",
    colors: {
      accent1: "#1D4ED8",
      brandTint: lighter("accent1", 0.4),
    },
  });

  const m2Theme = defineTheme({
    name: "Executive Slate",
    colors: {
      accent1: "#0F766E",
      accent2: "#134E4A",
    },
    fonts: { major: "Georgia", minor: "Calibri" },
  });

  const m1TitleLayout = defineSlideLayout({
    id: "m1-title",
    name: "Title Slide",
    type: "title",
    preserve: true,
    headerFooter: { dateTime: false, footer: true, slideNumber: true },
    footerPlaceholders: {
      footer: { text: "Master 1 Footer", style: { fontSize: "10pt" } },
    },
    placeholders: {
      title: {
        type: "ctrTitle",
        x: 60,
        y: 80,
        width: 840,
        height: 80,
        style: { fontSize: "28pt", color: "#0F172A" },
      },
      "body-1": {
        type: "body",
        idx: 1,
        x: 60,
        y: 190,
        width: 390,
        height: 220,
      },
      "body-2": {
        type: "body",
        idx: 2,
        x: 490,
        y: 190,
        width: 390,
        height: 220,
      },
    },
  });

  const m1IdleLayout = defineSlideLayout({
    id: "m1-blank",
    name: "Blank",
    type: "blank",
    preserve: true,
  });

  const m2UnusedReportLayout = defineSlideLayout({
    id: "m2-report",
    name: "Blank",
    type: "cust",
    preserve: true,
    placeholders: {
      title: {
        type: "title",
        x: 72,
        y: 56,
        width: 816,
        height: 64,
        style: { fontSize: "26pt", color: "#134E4A" },
      },
      content: {
        type: "obj",
        explicitType: false,
        idx: 1,
        x: 72,
        y: 144,
        width: 816,
        height: 320,
      },
    },
  });

  const vanillaDoc = deck({ width: 960, height: 540 }).append(
    slideMaster({
      id: "master-1",
      name: "Primary",
      theme: m1Theme,
      headerFooter: { dateTime: true, footer: true, slideNumber: true },
      footerPlaceholders: {
        dateTime: { text: "2026-11-24", visible: false },
        footer: { text: "Internal Draft" },
        slideNumber: { idx: 4 },
      },
    }).append(
      text("Primary Header Badge", {
        id: "m1-badge",
        x: 24,
        y: 16,
        width: 240,
        height: 24,
      }),
      slideLayout(m1TitleLayout),
      slideLayout(m1IdleLayout),
    ),
    slideMaster({
      id: "master-2",
      name: "Secondary",
      preserve: true,
      theme: m2Theme,
    }).append(
      text("Secondary Shared Watermark", {
        id: "m2-watermark",
        x: 24,
        y: 492,
        width: 280,
        height: 28,
      }),
      slideLayout(m2UnusedReportLayout),
    ),
    section({ name: "Briefing" }).append(
      slide({
        id: "s-inherited",
        layout: m1TitleLayout,
        footer: "Slide Footer Text",
        slideNumber: true,
      }).append(
        text("Inherited Headline", { placeholder: "title" }),
        text("Left Body", { placeholder: "body-1" }),
        text("Right Body", { placeholder: "body-2" }),
      ),
      slide({
        id: "s-explicit-equal",
        layout: m1TitleLayout,
      }).append(
        text("Locked Headline", {
          placeholder: "title",
          style: { fontSize: "28pt" },
        }),
        placeholder({
          id: "unpopulated-slot",
          placeholder: "body-1",
        }),
      ),
    ),
  );

  const reactElement = createElement(
    Deck,
    { width: 960, height: 540 },
    createElement(
      SlideMaster,
      {
        id: "master-1",
        name: "Primary",
        theme: m1Theme,
        headerFooter: { dateTime: true, footer: true, slideNumber: true },
        footerPlaceholders: {
          dateTime: { text: "2026-11-24", visible: false },
          footer: { text: "Internal Draft" },
          slideNumber: { idx: 4 },
        },
      },
      createElement(
        Text,
        { id: "m1-badge", x: 24, y: 16, width: 240, height: 24 },
        "Primary Header Badge",
      ),
      createElement(SlideLayout, { layout: m1TitleLayout }),
      createElement(SlideLayout, { layout: m1IdleLayout }),
    ),
    createElement(
      SlideMaster,
      {
        id: "master-2",
        name: "Secondary",
        preserve: true,
        theme: m2Theme,
      },
      createElement(
        Text,
        { id: "m2-watermark", x: 24, y: 492, width: 280, height: 28 },
        "Secondary Shared Watermark",
      ),
      createElement(SlideLayout, { layout: m2UnusedReportLayout }),
    ),
    createElement(
      Section,
      { name: "Briefing" },
      createElement(
        Slide,
        {
          id: "s-inherited",
          layout: m1TitleLayout,
          footer: "Slide Footer Text",
          slideNumber: true,
        },
        createElement(Text, { placeholder: "title" }, "Inherited Headline"),
        createElement(Text, { placeholder: "body-1" }, "Left Body"),
        createElement(Text, { placeholder: "body-2" }, "Right Body"),
      ),
      createElement(
        Slide,
        { id: "s-explicit-equal", layout: m1TitleLayout },
        createElement(
          Text,
          { placeholder: "title", style: { fontSize: "28pt" } },
          "Locked Headline",
        ),
        createElement(Placeholder, {
          id: "unpopulated-slot",
          placeholder: "body-1",
        }),
      ),
    ),
  );

  return { vanillaDoc, reactElement };
}

function extractMaster2VanillaFragment(generatedSource: string): string {
  const importMatch = generatedSource.match(
    /^import\s+\{[\s\S]*?\}\s+from\s+"@officedk\/pptxdk\/vanilla";/u,
  );
  const themeMatch = generatedSource.match(
    /const executiveSlateTheme = defineTheme\([\s\S]*?\n\}\);/u,
  );
  const layoutMatch = generatedSource.match(
    /const (master2\w*Layout) = defineSlideLayout\([\s\S]*?\n\}\);/u,
  );
  const masterMatch = generatedSource.match(
    /const master2 = slideMaster\([\s\S]*?\n\);/u,
  );
  assert.ok(importMatch, "Expected vanilla SDK import");
  assert.ok(themeMatch, "Expected copied Executive Slate theme");
  assert.ok(layoutMatch, "Expected copied Master 2 idle layout");
  assert.ok(masterMatch, "Expected copied Master 2 declaration");
  const layoutVar = layoutMatch[1]!;

  return `${importMatch[0]}

${themeMatch[0]}

${layoutMatch[0]}

${masterMatch[0]}

const reusedSlide = slide({
  id: "reused-page",
  layout: ${layoutVar},
}).append(
  text("Quarterly Summary", { placeholder: "title" }),
  text("Copied from Master 2", { placeholder: "content" }),
);

export default deck({ width: 960, height: 540 }).append(master2, reusedSlide);
`;
}

function extractMaster2ReactFragment(generatedSource: string): string {
  const importMatch = generatedSource.match(
    /^import\s+\{[\s\S]*?\}\s+from\s+"@officedk\/pptxdk\/react";/u,
  );
  const themeMatch = generatedSource.match(
    /const executiveSlateTheme = defineTheme\([\s\S]*?\n\}\);/u,
  );
  const layoutMatch = generatedSource.match(
    /const (master2\w*Layout) = defineSlideLayout\([\s\S]*?\n\}\);/u,
  );
  const masterMatch = generatedSource.match(
    /function Master2\(\)\s*\{[\s\S]*?\n\}/u,
  );
  assert.ok(importMatch, "Expected react SDK import");
  assert.ok(themeMatch, "Expected copied Executive Slate theme");
  assert.ok(layoutMatch, "Expected copied Master 2 idle layout");
  assert.ok(masterMatch, "Expected copied Master 2 component");
  const layoutVar = layoutMatch[1]!;

  return `${importMatch[0]}

${themeMatch[0]}

${layoutMatch[0]}

${masterMatch[0]}

function ReusedSlide() {
  return (
    <Slide id="reused-page" layout={${layoutVar}}>
      <Text placeholder="title">Quarterly Summary</Text>
      <Text placeholder="content">Copied from Master 2</Text>
    </Slide>
  );
}

export default function GeneratedDeck() {
  return (
    <Deck width={960} height={540}>
      <Master2 />
      <ReusedSlide />
    </Deck>
  );
}
`;
}

describe("pptxdk masters & placeholders codegen acceptance", () => {
  it("rebuilds vanilla and react sources for direct and imported dual-master decks including prs-slide-masters.pptx", async () => {
    const { vanillaDoc, reactElement } = buildDualMasterFixtures();
    const reactDoc = await expandDocument(reactElement);

    const builtOriginal = await build(vanillaDoc);
    assert.deepEqual(
      builtOriginal.report.diagnostics.filter((d) => d.severity === "error"),
      [],
    );
    const importedDual = importPptx(builtOriginal.pptx);
    assert.deepEqual(importedDual.report.diagnostics, []);

    const prsRaw = new Uint8Array(
      readFileSync(
        new URL("./fixtures/prs-slide-masters.pptx", import.meta.url),
      ),
    );
    const importedPrs = importPptx(prsRaw);
    prsRaw.fill(0);
    assert.equal(importedPrs.report.slideCount, 0);
    assert.ok(importedPrs.document instanceof DeckContainer);
    assert.equal(importedPrs.document.masters.length, 2);

    const dualCases = [
      { label: "direct-vanilla", doc: vanillaDoc },
      { label: "direct-react", doc: reactDoc },
      { label: "imported-dual", doc: importedDual.document },
    ].flatMap(({ label, doc }) => {
      const vanilla = codegen(doc, { target: "vanilla" });
      const react = codegen(doc, { target: "react" });
      return [
        { label: `${label}:vanilla`, result: vanilla },
        { label: `${label}:react`, result: react },
      ];
    });

    const prsVanilla = codegen(importedPrs.document, { target: "vanilla" });
    const prsReact = codegen(importedPrs.document, { target: "react" });
    const allCases = [
      ...dualCases,
      { label: "prs-slide-masters:vanilla", result: prsVanilla },
      { label: "prs-slide-masters:react", result: prsReact },
    ];

    await verifyGeneratedSourcesTypecheck(allCases);

    for (const item of dualCases) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], `${item.label} errors`);
      assert.equal(file.masters?.length, 2, `${item.label} master count`);
      assert.equal(file.masters?.[0]?.layouts.length, 2);
      assert.equal(file.masters?.[1]?.layouts.length, 1);
      assert.equal(file.masters?.[1]?.preserve, true);
      assert.equal(file.slides.length, 2);
      assert.equal(
        file.slides[0]?.layoutPart,
        file.masters?.[0]?.layouts[0]?.part,
      );
      assert.equal(file.slides[0]?.masterPart, file.masters?.[0]?.part);
      assert.equal(
        file.slides[1]?.objects[1]?.placeholder?.type,
        "body",
        `${item.label} unpopulated Placeholder preserved`,
      );
      assert.equal(
        file.slides[1]?.objects[1]?.text,
        undefined,
        `${item.label} unpopulated Placeholder has no synthesized txBody`,
      );
    }

    for (const item of [
      { label: "prs:vanilla", result: prsVanilla },
      { label: "prs:react", result: prsReact },
    ]) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], `${item.label} errors`);
      assert.equal(file.presentation.slideCount, 0);
      assert.equal(file.masters?.length, 2);
      assert.equal(file.masters?.[0]?.preserve, undefined);
      assert.equal(file.masters?.[1]?.preserve, true);
      assert.equal(file.masters?.[0]?.layouts[0]?.preserve, true);
      assert.equal(file.masters?.[1]?.layouts[0]?.preserve, true);
      assert.equal(file.masters?.[0]?.theme?.name, "Slide Master 1");
      assert.equal(file.masters?.[1]?.theme?.name, "Slide Master 2");
    }
  }, 60_000);

  it("supports copying Master 2 with its idle layout, shared object, and theme into a new deck and filling named placeholders", async () => {
    const { vanillaDoc } = buildDualMasterFixtures();
    const generatedVanilla = codegen(vanillaDoc, { target: "vanilla" });
    const generatedReact = codegen(vanillaDoc, { target: "react" });

    const copiedVanillaResult = {
      target: "vanilla" as const,
      entry: "deck.ts" as const,
      files: {
        "deck.ts": extractMaster2VanillaFragment(
          generatedVanilla.files["deck.ts"] as string,
        ),
      },
    };
    const copiedReactResult = {
      target: "react" as const,
      entry: "deck.tsx" as const,
      files: {
        "deck.tsx": extractMaster2ReactFragment(
          generatedReact.files["deck.tsx"] as string,
        ),
      },
    };

    // Also verify copying Master 2 from prs-slide-masters.pptx (zero-slide external fixture).
    const prsRaw = new Uint8Array(
      readFileSync(
        new URL("./fixtures/prs-slide-masters.pptx", import.meta.url),
      ),
    );
    const importedPrs = importPptx(prsRaw);
    prsRaw.fill(0);
    const prsVanillaSrc = codegen(importedPrs.document, { target: "vanilla" })
      .files["deck.ts"] as string;
    const prsImport = prsVanillaSrc.match(
      /^import\s+\{[\s\S]*?\}\s+from\s+"@officedk\/pptxdk\/vanilla";/u,
    )?.[0];
    const prsTheme2 = prsVanillaSrc.match(
      /const slideMaster2Theme = defineTheme\([\s\S]*?\n\}\);/u,
    )?.[0];
    const prsLayout2 = prsVanillaSrc.match(
      /const master2BlankLayout = defineSlideLayout\([\s\S]*?\n\}\);/u,
    )?.[0];
    const prsMaster2 = prsVanillaSrc.match(
      /const master2 = slideMaster\([\s\S]*?\n\);/u,
    )?.[0];
    assert.ok(prsImport && prsTheme2 && prsLayout2 && prsMaster2);
    const copiedPrsMaster2Result = {
      target: "vanilla" as const,
      entry: "deck.ts" as const,
      files: {
        "deck.ts": `${prsImport}
import { slide, text } from "@officedk/pptxdk/vanilla";

${prsTheme2}

${prsLayout2}

${prsMaster2}

const cover = slide({ id: "prs-cover", layout: master2BlankLayout }).append(
  text("From PRS Master 2", { id: "t1", x: 64, y: 64, width: 400, height: 48 }),
);

export default deck({ width: 960, height: 540 }).append(master2, cover);
`,
      },
    };

    await verifyGeneratedSourcesTypecheck([
      { label: "copied-master2:vanilla", result: copiedVanillaResult },
      { label: "copied-master2:react", result: copiedReactResult },
      { label: "copied-prs-master2:vanilla", result: copiedPrsMaster2Result },
    ]);

    for (const item of [
      { label: "copied-master2:vanilla", result: copiedVanillaResult },
      { label: "copied-master2:react", result: copiedReactResult },
    ]) {
      const { file } = await rebuildGeneratedSourceRelocated(item.result);
      assert.deepEqual(file.errors, [], `${item.label} errors`);
      assert.equal(file.masters?.length, 1);
      assert.equal(file.slides.length, 1);
      // Slide binds to the copied Master 2 and its previously idle layout.
      assert.equal(
        file.slides[0]?.layoutPart,
        file.masters?.[0]?.layouts[0]?.part,
      );
      assert.equal(file.slides[0]?.masterPart, file.masters?.[0]?.part);
      // Named placeholders resolve to the layout's coordinates & <p:ph> metadata.
      assert.equal(file.slides[0]?.objects[0]?.text?.text, "Quarterly Summary");
      assert.deepEqual(file.slides[0]?.objects[0]?.placeholder, {
        type: "title",
      });
      assert.equal(file.slides[0]?.objects[0]?.boxEmu.x, 72 * 9525);
      assert.equal(
        file.slides[0]?.objects[1]?.text?.text,
        "Copied from Master 2",
      );
      assert.deepEqual(file.slides[0]?.objects[1]?.placeholder, {
        type: "obj",
        idx: 1,
      });
      // Shared object stays on the master rather than duplicated onto slides.
      assert.equal(
        file.masters?.[0]?.objects?.[0]?.text?.text,
        "Secondary Shared Watermark",
      );
      assert.equal(file.slides[0]?.objects.length, 2);
      // Theme reference preserved.
      assert.equal(file.masters?.[0]?.theme?.name, "Executive Slate");
      assert.equal(file.masters?.[0]?.theme?.fonts.major, "Georgia");
    }

    const prsRebuilt = await rebuildGeneratedSourceRelocated(
      copiedPrsMaster2Result,
    );
    assert.deepEqual(prsRebuilt.file.errors, []);
    assert.equal(prsRebuilt.file.masters?.length, 1);
    assert.equal(prsRebuilt.file.masters?.[0]?.preserve, true);
    assert.equal(prsRebuilt.file.masters?.[0]?.theme?.name, "Slide Master 2");
    assert.equal(
      prsRebuilt.file.slides[0]?.layoutPart,
      prsRebuilt.file.masters?.[0]?.layouts[0]?.part,
    );
  }, 60_000);

  it("propagates layout fontSize edits to inheriting slides while preserving equal explicit overrides, and emits deterministic output for colliding names", async () => {
    const sharedA = defineTheme({
      name: "Shared Theme",
      colors: { accent1: "#2563EB" },
    });
    const sharedB = defineTheme({
      name: "Shared Theme",
      colors: { accent1: "#DC2626" },
    });

    const layoutA = defineSlideLayout({
      id: "layout-a",
      name: "Blank",
      type: "title",
      placeholders: {
        title: {
          type: "ctrTitle",
          x: 60,
          y: 80,
          width: 720,
          height: 60,
          style: { fontSize: "28pt", color: "#111827" },
        },
        "body-1": {
          type: "body",
          idx: 1,
          x: 60,
          y: 180,
          width: 340,
          height: 180,
        },
        "body-2": {
          type: "body",
          idx: 2,
          x: 440,
          y: 180,
          width: 340,
          height: 180,
        },
      },
    });
    const layoutB = defineSlideLayout({
      id: "layout-b",
      name: "Blank",
      type: "blank",
      preserve: true,
    });

    const doc = deck({ width: 960, height: 540, theme: defaultTheme }).append(
      slideMaster({ name: "Brand Master", theme: sharedA }).append(
        slideLayout(layoutA),
      ),
      slideMaster({ name: "Brand Master", theme: sharedB }).append(
        slideLayout(layoutB),
      ),
      slide({ id: "s-inherit", layout: layoutA }).append(
        text("Inherits Layout", { placeholder: "title" }),
        text("First Body", { placeholder: "body-1" }),
        text("Second Body", { placeholder: "body-2" }),
      ),
      slide({ id: "s-locked", layout: layoutA }).append(
        text("Equal Override", {
          placeholder: "title",
          style: { fontSize: "28pt" },
        }),
      ),
    );

    for (const target of ["vanilla", "react"] as const) {
      const first = codegen(doc, { target });
      const second = codegen(doc, { target });
      assert.equal(first.files[first.entry], second.files[second.entry]);

      const originalSource = first.files[first.entry] as string;
      const editedSource = originalSource.replace(
        'fontSize: "28pt", color: "#111827"',
        'fontSize: "40pt", color: "#111827"',
      );
      assert.notEqual(
        editedSource,
        originalSource,
        `${target} layout fontSize edit should apply`,
      );

      const editedResult = {
        target,
        entry: first.entry,
        files: { [first.entry]: editedSource },
      };
      await verifyGeneratedSourcesTypecheck([
        { label: `deterministic:${target}`, result: first },
        { label: `edited-layout:${target}`, result: editedResult },
      ]);

      const { built, file } =
        await rebuildGeneratedSourceRelocated(editedResult);
      assert.deepEqual(file.errors, [], `${target} edited errors`);
      assert.equal(file.masters?.length, 2);
      assert.equal(file.slides.length, 2);
      // Slide 1 omits local fontSize and inherits the edited 40pt from slideLayout1.xml.
      assert.equal(
        file.slides[0]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        undefined,
      );
      const layoutXml = strFromU8(
        unzipSync(built.pptx)["ppt/slideLayouts/slideLayout1.xml"]!,
      );
      assert.match(layoutXml, /sz="4000"/u);
      // Slide 2 retains its explicit 28pt override even though it originally equaled the layout's 28pt.
      assert.equal(
        file.slides[1]?.objects[0]?.text?.paragraphs[0]?.runs[0]?.fontSizePt,
        28,
      );
    }
  }, 60_000);
});
